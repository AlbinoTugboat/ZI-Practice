#include "pch.h"
#include "ServiceManager.h"

#include <userenv.h>
#include <wtsapi32.h>

#pragma comment(lib, "Userenv.lib")
#pragma comment(lib, "Wtsapi32.lib")

namespace ZIVPO::Service
{
    namespace
    {
        constexpr wchar_t kServiceName[] = L"ZIVPO.SessionLauncher";
        constexpr wchar_t kServiceDisplayName[] = L"ZIVPO Session Launcher";
        constexpr wchar_t kServiceDescription[] = L"Starts ZIVPO in user sessions in hidden mode.";
        constexpr wchar_t kHiddenSwitch[] = L"--hidden";
        constexpr wchar_t kGuiExecutableName[] = L"ZIVPO.exe";
        constexpr wchar_t kServiceExecutableName[] = L"ZIVPO.Service.exe";

        SERVICE_STATUS_HANDLE g_statusHandle = nullptr;
        SERVICE_STATUS g_serviceStatus{};
        HANDLE g_stopEvent = nullptr;
        DWORD g_checkPoint = 1;

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

            std::wstring value = L"\"";
            value.append(servicePath);
            value.push_back(L'"');
            return value;
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
                g_serviceStatus.dwControlsAccepted =
                    SERVICE_ACCEPT_STOP |
                    SERVICE_ACCEPT_SHUTDOWN |
                    SERVICE_ACCEPT_SESSIONCHANGE;
            }

            SetServiceStatus(g_statusHandle, &g_serviceStatus);
        }

        bool LaunchGuiHiddenInSession(DWORD sessionId)
        {
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
                if (session.State == WTSActive || session.State == WTSConnected)
                {
                    LaunchGuiHiddenInSession(session.SessionId);
                }
            }

            WTSFreeMemory(sessions);
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
                UpdateServiceStatus(SERVICE_STOP_PENDING, NO_ERROR, 5000);
                if (g_stopEvent != nullptr)
                {
                    SetEvent(g_stopEvent);
                }
                return NO_ERROR;

            case SERVICE_CONTROL_SESSIONCHANGE:
                if (eventData != nullptr)
                {
                    auto const* notification = static_cast<WTSSESSION_NOTIFICATION*>(eventData);
                    if (eventType == WTS_SESSION_LOGON ||
                        eventType == WTS_CONSOLE_CONNECT ||
                        eventType == WTS_REMOTE_CONNECT)
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

            UpdateServiceStatus(SERVICE_START_PENDING, NO_ERROR, 5000);

            g_stopEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
            if (g_stopEvent == nullptr)
            {
                UpdateServiceStatus(SERVICE_STOPPED, GetLastError(), 0);
                return;
            }

            LaunchGuiInAllActiveSessions();

            UpdateServiceStatus(SERVICE_RUNNING, NO_ERROR, 0);
            WaitForSingleObject(g_stopEvent, INFINITE);

            CloseHandle(g_stopEvent);
            g_stopEvent = nullptr;

            UpdateServiceStatus(SERVICE_STOPPED, NO_ERROR, 0);
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

        void StartServiceIfStopped(SC_HANDLE serviceHandle)
        {
            SERVICE_STATUS_PROCESS status{};
            DWORD bytesNeeded = 0;
            if (!QueryServiceStatusEx(
                serviceHandle,
                SC_STATUS_PROCESS_INFO,
                reinterpret_cast<LPBYTE>(&status),
                sizeof(status),
                &bytesNeeded))
            {
                return;
            }

            if (status.dwCurrentState == SERVICE_RUNNING ||
                status.dwCurrentState == SERVICE_START_PENDING)
            {
                return;
            }

            StartServiceW(serviceHandle, 0, nullptr);
        }
    }

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

    void EnsureServiceRunningForGui()
    {
        SC_HANDLE scm = OpenSCManagerW(nullptr, nullptr, SC_MANAGER_CONNECT | SC_MANAGER_CREATE_SERVICE);
        if (scm == nullptr)
        {
            scm = OpenSCManagerW(nullptr, nullptr, SC_MANAGER_CONNECT);
            if (scm == nullptr)
            {
                return;
            }
        }

        SC_HANDLE serviceHandle = CreateServiceIfMissing(scm);
        if (serviceHandle == nullptr)
        {
            serviceHandle = OpenServiceWithAccess(scm, SERVICE_QUERY_STATUS | SERVICE_START);
        }

        if (serviceHandle != nullptr)
        {
            StartServiceIfStopped(serviceHandle);
            CloseServiceHandle(serviceHandle);
        }
        else
        {
            TraceService(L"Unable to open or create service.");
        }

        CloseServiceHandle(scm);
    }
}
