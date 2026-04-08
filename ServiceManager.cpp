#include "pch.h"
#include "ServiceManager.h"

#include "ServiceRpc.h"

#include <rpc.h>
#include <tlhelp32.h>
#include <userenv.h>
#include <wtsapi32.h>

#if defined(ZIVPO_SERVICE_HOST)
#include <mutex>
#include <unordered_map>
#include <vector>
#endif

namespace
{
    constexpr wchar_t kServiceName[] = L"ZIVPO.SessionLauncher";
    constexpr wchar_t kServiceDisplayName[] = L"ZIVPO Session Launcher";
    constexpr wchar_t kServiceDescription[] = L"Starts ZIVPO in user sessions in hidden mode.";
    constexpr wchar_t kServiceExecutableName[] = L"ZIVPO.Service.exe";
    constexpr wchar_t kGuiExecutableName[] = L"ZIVPO.exe";
    constexpr wchar_t kRpcProtocol[] = L"ncalrpc";
    constexpr wchar_t kRpcEndpoint[] = L"ZIVPO.Service.Control";
    constexpr DWORD kServicePollIntervalMs = 500;
    constexpr DWORD kServiceStartTimeoutMs = 60000;

    void TraceService(std::wstring_view message) noexcept
    {
        OutputDebugStringW(L"[ZIVPO.Service] ");
        OutputDebugStringW(message.data());
        OutputDebugStringW(L"\r\n");
    }

    std::wstring GetModulePath()
    {
        std::wstring path(MAX_PATH, L'\0');
        DWORD length = 0;

        while (true)
        {
            length = GetModuleFileNameW(nullptr, path.data(), static_cast<DWORD>(path.size()));
            if (length == 0)
            {
                return {};
            }

            if (length < path.size() - 1)
            {
                path.resize(length);
                return path;
            }

            path.resize(path.size() * 2);
        }
    }

    std::wstring GetDirectoryPath(std::wstring const& path)
    {
        size_t const separator = path.find_last_of(L"\\/");
        if (separator == std::wstring::npos)
        {
            return {};
        }

        return path.substr(0, separator + 1);
    }

    std::wstring BuildSiblingPath(std::wstring const& executablePath, std::wstring_view fileName)
    {
        std::wstring directory = GetDirectoryPath(executablePath);
        if (directory.empty())
        {
            return {};
        }

        directory.append(fileName);
        return directory;
    }

    std::wstring BuildServiceBinaryPath(std::wstring const& guiExecutablePath)
    {
        std::wstring servicePath = BuildSiblingPath(guiExecutablePath, kServiceExecutableName);
        if (servicePath.empty())
        {
            return {};
        }

        std::wstring quoted = L"\"";
        quoted.append(servicePath);
        quoted.push_back(L'"');
        return quoted;
    }

    SC_HANDLE OpenServiceWithAccess(SC_HANDLE scm, DWORD accessMask)
    {
        return OpenServiceW(scm, kServiceName, accessMask);
    }

    void EnsureServiceConfiguration(SC_HANDLE serviceHandle, std::wstring const& binaryPath)
    {
        ChangeServiceConfigW(
            serviceHandle,
            SERVICE_NO_CHANGE,
            SERVICE_AUTO_START,
            SERVICE_NO_CHANGE,
            binaryPath.empty() ? nullptr : binaryPath.c_str(),
            nullptr,
            nullptr,
            nullptr,
            nullptr,
            nullptr,
            nullptr);
    }

    SC_HANDLE CreateServiceIfMissing(SC_HANDLE scm)
    {
        std::wstring executablePath = GetModulePath();
        if (executablePath.empty())
        {
            return nullptr;
        }

        std::wstring binaryPath = BuildServiceBinaryPath(executablePath);
        if (binaryPath.empty())
        {
            return nullptr;
        }

        SC_HANDLE serviceHandle = OpenServiceWithAccess(
            scm,
            SERVICE_QUERY_STATUS | SERVICE_START | SERVICE_CHANGE_CONFIG);

        if (serviceHandle != nullptr)
        {
            EnsureServiceConfiguration(serviceHandle, binaryPath);
            return serviceHandle;
        }

        if (GetLastError() != ERROR_SERVICE_DOES_NOT_EXIST)
        {
            return nullptr;
        }

        serviceHandle = CreateServiceW(
            scm,
            kServiceName,
            kServiceDisplayName,
            SERVICE_QUERY_STATUS | SERVICE_START | SERVICE_CHANGE_CONFIG,
            SERVICE_WIN32_OWN_PROCESS,
            SERVICE_AUTO_START,
            SERVICE_ERROR_NORMAL,
            binaryPath.c_str(),
            nullptr,
            nullptr,
            nullptr,
            nullptr,
            nullptr);

        if (serviceHandle == nullptr)
        {
            return nullptr;
        }

        SERVICE_DESCRIPTIONW description{};
        description.lpDescription = const_cast<LPWSTR>(kServiceDescription);
        ChangeServiceConfig2W(serviceHandle, SERVICE_CONFIG_DESCRIPTION, &description);
        return serviceHandle;
    }

    bool QueryServiceStatusProcess(SC_HANDLE serviceHandle, SERVICE_STATUS_PROCESS& status)
    {
        DWORD bytesNeeded = 0;
        return QueryServiceStatusEx(
            serviceHandle,
            SC_STATUS_PROCESS_INFO,
            reinterpret_cast<LPBYTE>(&status),
            sizeof(status),
            &bytesNeeded) != FALSE;
    }

    bool WaitForServiceRunning(SC_HANDLE serviceHandle, DWORD timeoutMs)
    {
        ULONGLONG const deadline = GetTickCount64() + timeoutMs;
        SERVICE_STATUS_PROCESS status{};

        while (GetTickCount64() <= deadline)
        {
            if (!QueryServiceStatusProcess(serviceHandle, status))
            {
                return false;
            }

            if (status.dwCurrentState == SERVICE_RUNNING)
            {
                return true;
            }

            if (status.dwCurrentState == SERVICE_STOPPED)
            {
                return false;
            }

            Sleep(kServicePollIntervalMs);
        }

        return false;
    }

    bool EnsureServiceRunning(SC_HANDLE serviceHandle, bool& startedByCaller)
    {
        startedByCaller = false;

        SERVICE_STATUS_PROCESS status{};
        if (!QueryServiceStatusProcess(serviceHandle, status))
        {
            return false;
        }

        if (status.dwCurrentState == SERVICE_RUNNING)
        {
            return true;
        }

        if (status.dwCurrentState == SERVICE_START_PENDING)
        {
            return WaitForServiceRunning(serviceHandle, kServiceStartTimeoutMs);
        }

        if (!StartServiceW(serviceHandle, 0, nullptr))
        {
            if (GetLastError() != ERROR_SERVICE_ALREADY_RUNNING)
            {
                return false;
            }
        }
        else
        {
            startedByCaller = true;
        }

        return WaitForServiceRunning(serviceHandle, kServiceStartTimeoutMs);
    }

    DWORD GetParentProcessId(DWORD processId)
    {
        HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
        if (snapshot == INVALID_HANDLE_VALUE)
        {
            return 0;
        }

        PROCESSENTRY32W processEntry{};
        processEntry.dwSize = sizeof(processEntry);

        DWORD parentProcessId = 0;
        if (Process32FirstW(snapshot, &processEntry))
        {
            do
            {
                if (processEntry.th32ProcessID == processId)
                {
                    parentProcessId = processEntry.th32ParentProcessID;
                    break;
                }
            } while (Process32NextW(snapshot, &processEntry));
        }

        CloseHandle(snapshot);
        return parentProcessId;
    }

    std::wstring GetProcessExeName(DWORD processId)
    {
        HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
        if (snapshot == INVALID_HANDLE_VALUE)
        {
            return {};
        }

        PROCESSENTRY32W processEntry{};
        processEntry.dwSize = sizeof(processEntry);
        std::wstring result;

        if (Process32FirstW(snapshot, &processEntry))
        {
            do
            {
                if (processEntry.th32ProcessID == processId)
                {
                    result = processEntry.szExeFile;
                    break;
                }
            } while (Process32NextW(snapshot, &processEntry));
        }

        CloseHandle(snapshot);
        return result;
    }

    bool IsCurrentProcessParentService()
    {
        DWORD const parentProcessId = GetParentProcessId(GetCurrentProcessId());
        if (parentProcessId == 0)
        {
            return false;
        }

        std::wstring parentExeName = GetProcessExeName(parentProcessId);
        if (parentExeName.empty())
        {
            return false;
        }

        return _wcsicmp(parentExeName.c_str(), kServiceExecutableName) == 0;
    }
}

namespace ZIVPO::Service
{
    GuiStartupDecision PrepareGuiStartup()
    {
        SC_HANDLE scm = OpenSCManagerW(nullptr, nullptr, SC_MANAGER_CONNECT | SC_MANAGER_CREATE_SERVICE);
        if (scm == nullptr)
        {
            scm = OpenSCManagerW(nullptr, nullptr, SC_MANAGER_CONNECT);
            if (scm == nullptr)
            {
                return GuiStartupDecision::Error;
            }
        }

        SC_HANDLE serviceHandle = CreateServiceIfMissing(scm);
        if (serviceHandle == nullptr)
        {
            serviceHandle = OpenServiceWithAccess(scm, SERVICE_QUERY_STATUS | SERVICE_START);
        }

        if (serviceHandle == nullptr)
        {
            CloseServiceHandle(scm);
            return GuiStartupDecision::Error;
        }

        bool startedByCaller = false;
        bool running = EnsureServiceRunning(serviceHandle, startedByCaller);

        CloseServiceHandle(serviceHandle);
        CloseServiceHandle(scm);

        if (!running)
        {
            return GuiStartupDecision::Error;
        }

        if (startedByCaller)
        {
            return GuiStartupDecision::Exit;
        }

        return IsCurrentProcessParentService()
            ? GuiStartupDecision::Continue
            : GuiStartupDecision::Exit;
    }
}

#if defined(ZIVPO_SERVICE_HOST)
namespace
{
    constexpr wchar_t kHiddenSwitch[] = L"--hidden";

    SERVICE_STATUS_HANDLE g_statusHandle = nullptr;
    SERVICE_STATUS g_serviceStatus{};
    DWORD g_checkPoint = 1;
    DWORD g_serviceProcessId = 0;
    bool g_rpcStopRequested = false;
    std::mutex g_launchedProcessesMutex;
    std::unordered_map<DWORD, DWORD> g_launchedGuiProcessBySession;

    bool IsProcessRunning(DWORD processId)
    {
        HANDLE process = OpenProcess(SYNCHRONIZE, FALSE, processId);
        if (process == nullptr)
        {
            return false;
        }

        DWORD waitResult = WaitForSingleObject(process, 0);
        CloseHandle(process);
        return waitResult == WAIT_TIMEOUT;
    }

    bool IsSessionGuiAlreadyRunning(DWORD sessionId)
    {
        std::lock_guard lock(g_launchedProcessesMutex);
        auto it = g_launchedGuiProcessBySession.find(sessionId);
        if (it == g_launchedGuiProcessBySession.end())
        {
            return false;
        }

        if (IsProcessRunning(it->second))
        {
            return true;
        }

        g_launchedGuiProcessBySession.erase(it);
        return false;
    }

    void TrackLaunchedGuiProcess(DWORD sessionId, DWORD processId)
    {
        std::lock_guard lock(g_launchedProcessesMutex);
        g_launchedGuiProcessBySession[sessionId] = processId;
    }

    std::vector<DWORD> CollectTrackedGuiProcesses()
    {
        std::vector<DWORD> result;
        std::lock_guard lock(g_launchedProcessesMutex);
        result.reserve(g_launchedGuiProcessBySession.size());
        for (auto const& entry : g_launchedGuiProcessBySession)
        {
            result.push_back(entry.second);
        }
        g_launchedGuiProcessBySession.clear();
        return result;
    }

    std::vector<DWORD> CollectServiceChildGuiProcesses()
    {
        std::vector<DWORD> result;
        HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
        if (snapshot == INVALID_HANDLE_VALUE)
        {
            return result;
        }

        PROCESSENTRY32W processEntry{};
        processEntry.dwSize = sizeof(processEntry);

        if (Process32FirstW(snapshot, &processEntry))
        {
            do
            {
                if (_wcsicmp(processEntry.szExeFile, kGuiExecutableName) != 0)
                {
                    continue;
                }

                if (processEntry.th32ParentProcessID == g_serviceProcessId)
                {
                    result.push_back(processEntry.th32ProcessID);
                }
            } while (Process32NextW(snapshot, &processEntry));
        }

        CloseHandle(snapshot);
        return result;
    }

    void TerminateProcessById(DWORD processId)
    {
        HANDLE process = OpenProcess(PROCESS_TERMINATE | SYNCHRONIZE, FALSE, processId);
        if (process == nullptr)
        {
            return;
        }

        TerminateProcess(process, 0);
        WaitForSingleObject(process, 5000);
        CloseHandle(process);
    }

    void TerminateLaunchedGuiProcesses()
    {
        std::vector<DWORD> processIds = CollectTrackedGuiProcesses();
        std::vector<DWORD> childIds = CollectServiceChildGuiProcesses();
        processIds.insert(processIds.end(), childIds.begin(), childIds.end());

        for (DWORD processId : processIds)
        {
            TerminateProcessById(processId);
        }
    }

    void UpdateServiceStatus(DWORD state, DWORD win32ExitCode = NO_ERROR, DWORD waitHint = 0)
    {
        if (g_statusHandle == nullptr)
        {
            return;
        }

        g_serviceStatus.dwCurrentState = state;
        g_serviceStatus.dwWin32ExitCode = win32ExitCode;
        g_serviceStatus.dwWaitHint = waitHint;
        g_serviceStatus.dwControlsAccepted = 0;

        if (state == SERVICE_START_PENDING || state == SERVICE_STOP_PENDING)
        {
            g_serviceStatus.dwCheckPoint = g_checkPoint++;
        }
        else
        {
            g_serviceStatus.dwCheckPoint = 0;
        }

        if (state == SERVICE_RUNNING)
        {
            g_serviceStatus.dwControlsAccepted = SERVICE_ACCEPT_SESSIONCHANGE;
        }

        SetServiceStatus(g_statusHandle, &g_serviceStatus);
    }

    bool IsExplorerRunningInSession(DWORD sessionId)
    {
        HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
        if (snapshot == INVALID_HANDLE_VALUE)
        {
            return true;
        }

        PROCESSENTRY32W processEntry{};
        processEntry.dwSize = sizeof(processEntry);
        bool found = false;

        if (Process32FirstW(snapshot, &processEntry))
        {
            do
            {
                if (_wcsicmp(processEntry.szExeFile, L"explorer.exe") != 0)
                {
                    continue;
                }

                DWORD processSessionId = 0;
                if (!ProcessIdToSessionId(processEntry.th32ProcessID, &processSessionId))
                {
                    continue;
                }

                if (processSessionId == sessionId)
                {
                    found = true;
                    break;
                }
            } while (Process32NextW(snapshot, &processEntry));
        }

        CloseHandle(snapshot);
        return found;
    }

    void WaitForExplorerInSession(DWORD sessionId)
    {
        constexpr int kMaxAttempts = 60;
        constexpr DWORD kDelayMs = 500;

        for (int attempt = 0; attempt < kMaxAttempts; ++attempt)
        {
            if (IsExplorerRunningInSession(sessionId))
            {
                return;
            }

            Sleep(kDelayMs);
        }
    }

    bool LaunchGuiHiddenInSession(DWORD sessionId)
    {
        if (sessionId == 0 || IsSessionGuiAlreadyRunning(sessionId))
        {
            return true;
        }

        WaitForExplorerInSession(sessionId);

        HANDLE impersonationToken = nullptr;
        if (!WTSQueryUserToken(sessionId, &impersonationToken))
        {
            return false;
        }

        HANDLE primaryToken = nullptr;
        bool duplicated = DuplicateTokenEx(
            impersonationToken,
            TOKEN_ASSIGN_PRIMARY | TOKEN_DUPLICATE | TOKEN_QUERY | TOKEN_ADJUST_DEFAULT | TOKEN_ADJUST_SESSIONID,
            nullptr,
            SecurityImpersonation,
            TokenPrimary,
            &primaryToken) != FALSE;

        CloseHandle(impersonationToken);
        if (!duplicated)
        {
            return false;
        }

        LPVOID environment = nullptr;
        if (!CreateEnvironmentBlock(&environment, primaryToken, FALSE))
        {
            environment = nullptr;
        }

        std::wstring servicePath = GetModulePath();
        std::wstring appPath = BuildSiblingPath(servicePath, kGuiExecutableName);
        if (appPath.empty())
        {
            if (environment != nullptr)
            {
                DestroyEnvironmentBlock(environment);
            }
            CloseHandle(primaryToken);
            return false;
        }

        std::wstring commandLine = L"\"";
        commandLine.append(appPath);
        commandLine.append(L"\" ");
        commandLine.append(kHiddenSwitch);

        STARTUPINFOW startupInfo{};
        startupInfo.cb = sizeof(startupInfo);
        startupInfo.lpDesktop = const_cast<LPWSTR>(L"winsta0\\default");
        startupInfo.dwFlags = STARTF_USESHOWWINDOW;
        startupInfo.wShowWindow = SW_HIDE;

        PROCESS_INFORMATION processInfo{};
        DWORD creationFlags = CREATE_UNICODE_ENVIRONMENT | CREATE_NEW_PROCESS_GROUP;

        bool started = CreateProcessAsUserW(
            primaryToken,
            appPath.c_str(),
            commandLine.data(),
            nullptr,
            nullptr,
            FALSE,
            creationFlags,
            environment,
            nullptr,
            &startupInfo,
            &processInfo) != FALSE;

        if (environment != nullptr)
        {
            DestroyEnvironmentBlock(environment);
        }

        CloseHandle(primaryToken);
        if (!started)
        {
            return false;
        }

        TrackLaunchedGuiProcess(sessionId, processInfo.dwProcessId);
        CloseHandle(processInfo.hThread);
        CloseHandle(processInfo.hProcess);
        return true;
    }

    void LaunchGuiInAllActiveSessions()
    {
        PWTS_SESSION_INFOW sessions = nullptr;
        DWORD sessionCount = 0;

        if (!WTSEnumerateSessionsW(WTS_CURRENT_SERVER_HANDLE, 0, 1, &sessions, &sessionCount))
        {
            return;
        }

        for (DWORD index = 0; index < sessionCount; ++index)
        {
            auto const& session = sessions[index];
            if (session.SessionId == 0)
            {
                continue;
            }

            if (session.State == WTSActive || session.State == WTSConnected)
            {
                LaunchGuiHiddenInSession(session.SessionId);
            }
        }

        WTSFreeMemory(sessions);
    }

    RPC_STATUS StartRpcServer()
    {
        RPC_STATUS status = RpcServerUseProtseqEpW(
            reinterpret_cast<RPC_WSTR>(const_cast<wchar_t*>(kRpcProtocol)),
            RPC_C_PROTSEQ_MAX_REQS_DEFAULT,
            reinterpret_cast<RPC_WSTR>(const_cast<wchar_t*>(kRpcEndpoint)),
            nullptr);

        if (status != RPC_S_OK && status != RPC_S_DUPLICATE_ENDPOINT)
        {
            return status;
        }

        status = RpcServerRegisterIf2(
            ZivpoServiceControl_v1_0_s_ifspec,
            nullptr,
            nullptr,
            RPC_IF_ALLOW_LOCAL_ONLY,
            RPC_C_LISTEN_MAX_CALLS_DEFAULT,
            static_cast<unsigned int>(-1),
            nullptr);

        if (status != RPC_S_OK && status != RPC_S_TYPE_ALREADY_REGISTERED)
        {
            return status;
        }

        return RPC_S_OK;
    }

    RPC_STATUS RunRpcServerLoop()
    {
        RPC_STATUS status = RpcServerListen(1, RPC_C_LISTEN_MAX_CALLS_DEFAULT, FALSE);
        if (status != RPC_S_OK && status != RPC_S_ALREADY_LISTENING)
        {
            return status;
        }

        status = RpcMgmtWaitServerListen();
        if (status == RPC_S_OK || status == RPC_S_NOT_LISTENING)
        {
            return RPC_S_OK;
        }

        return status;
    }

    void UnregisterRpcInterface()
    {
        RpcServerUnregisterIf(ZivpoServiceControl_v1_0_s_ifspec, nullptr, FALSE);
    }

    DWORD WINAPI ServiceControlHandler(
        DWORD controlCode,
        DWORD eventType,
        LPVOID eventData,
        [[maybe_unused]] LPVOID context)
    {
        switch (controlCode)
        {
        case SERVICE_CONTROL_STOP:
        case SERVICE_CONTROL_SHUTDOWN:
            // Stop/Shutdown handling must stay disabled by assignment.
            return NO_ERROR;

        case SERVICE_CONTROL_SESSIONCHANGE:
            if (eventData != nullptr)
            {
                auto const* notification = static_cast<WTSSESSION_NOTIFICATION*>(eventData);
                if (notification->dwSessionId != 0 &&
                    (eventType == WTS_SESSION_LOGON ||
                     eventType == WTS_CONSOLE_CONNECT ||
                     eventType == WTS_REMOTE_CONNECT))
                {
                    LaunchGuiHiddenInSession(notification->dwSessionId);
                }
            }
            return NO_ERROR;

        default:
            return NO_ERROR;
        }
    }

    void WINAPI ServiceMain([[maybe_unused]] DWORD argc, [[maybe_unused]] LPWSTR* argv)
    {
        g_statusHandle = RegisterServiceCtrlHandlerExW(kServiceName, ServiceControlHandler, nullptr);
        if (g_statusHandle == nullptr)
        {
            return;
        }

        g_serviceStatus = {};
        g_serviceStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS;
        g_serviceStatus.dwServiceSpecificExitCode = 0;
        g_checkPoint = 1;
        g_serviceProcessId = GetCurrentProcessId();
        g_rpcStopRequested = false;

        UpdateServiceStatus(SERVICE_START_PENDING, NO_ERROR, 5000);

        RPC_STATUS rpcStatus = StartRpcServer();
        if (rpcStatus != RPC_S_OK)
        {
            UpdateServiceStatus(SERVICE_STOPPED, rpcStatus, 0);
            return;
        }

        LaunchGuiInAllActiveSessions();
        UpdateServiceStatus(SERVICE_RUNNING, NO_ERROR, 0);

        rpcStatus = RunRpcServerLoop();

        UpdateServiceStatus(SERVICE_STOP_PENDING, NO_ERROR, 5000);
        UnregisterRpcInterface();
        TerminateLaunchedGuiProcesses();

        DWORD exitCode = (rpcStatus == RPC_S_OK || g_rpcStopRequested) ? NO_ERROR : rpcStatus;
        UpdateServiceStatus(SERVICE_STOPPED, exitCode, 0);
    }
}

extern "C" void RpcStopService([[maybe_unused]] handle_t hBinding)
{
    g_rpcStopRequested = true;
    RpcMgmtStopServerListening(nullptr);
}

namespace ZIVPO::Service
{
    int RunServiceMode()
    {
        SERVICE_TABLE_ENTRYW serviceTable[] =
        {
            { const_cast<LPWSTR>(kServiceName), &ServiceMain },
            { nullptr, nullptr }
        };

        if (!StartServiceCtrlDispatcherW(serviceTable))
        {
            return static_cast<int>(GetLastError());
        }

        return 0;
    }

    bool RequestServiceStop()
    {
        return false;
    }
}
#else
namespace ZIVPO::Service
{
    int RunServiceMode()
    {
        return static_cast<int>(ERROR_CALL_NOT_IMPLEMENTED);
    }

    bool RequestServiceStop()
    {
        RPC_WSTR bindingString = nullptr;
        RPC_BINDING_HANDLE bindingHandle = nullptr;

        RPC_STATUS status = RpcStringBindingComposeW(
            nullptr,
            reinterpret_cast<RPC_WSTR>(const_cast<wchar_t*>(kRpcProtocol)),
            nullptr,
            reinterpret_cast<RPC_WSTR>(const_cast<wchar_t*>(kRpcEndpoint)),
            nullptr,
            &bindingString);

        if (status != RPC_S_OK)
        {
            return false;
        }

        status = RpcBindingFromStringBindingW(bindingString, &bindingHandle);
        RpcStringFreeW(&bindingString);
        if (status != RPC_S_OK)
        {
            return false;
        }

        bool success = true;

        RpcTryExcept
        {
            RpcStopService(bindingHandle);
        }
        RpcExcept(EXCEPTION_EXECUTE_HANDLER)
        {
            success = false;
        }
        RpcEndExcept;

        RpcBindingFree(&bindingHandle);
        return success;
    }
}
#endif
