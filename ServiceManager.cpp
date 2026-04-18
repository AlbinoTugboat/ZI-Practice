#include "pch.h"
#include "ServiceManager.h"

#include "ServiceRpc.h"

#include <rpc.h>
#include <tlhelp32.h>
#include <userenv.h>
#include <winhttp.h>
#include <wtsapi32.h>

#include <algorithm>
#include <array>
#include <chrono>
#include <cctype>
#include <ctime>
#include <cwctype>
#include <iomanip>
#include <optional>
#include <regex>
#include <sstream>
#include <string>

#if defined(ZIVPO_SERVICE_HOST)
#include <atomic>
#include <mutex>
#include <thread>
#include <unordered_map>
#include <unordered_set>
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
    constexpr wchar_t kDefaultApiBaseUrl[] = L"https://localhost:8444";
    constexpr wchar_t kApiAuthLoginPath[] = L"/api/auth/login";
    constexpr wchar_t kApiAuthRefreshPath[] = L"/api/auth/refresh";
    constexpr wchar_t kApiUserMePath[] = L"/api/user/me";
    constexpr wchar_t kApiLicenseCheckPath[] = L"/api/user/licenses/check";
    constexpr wchar_t kApiLicenseActivatePath[] = L"/api/user/licenses/activate";
    constexpr long long kTokenRefreshSafetyWindowSeconds = 30;
    constexpr long long kLicenseRefreshSafetyWindowSeconds = 30;
    constexpr long long kDefaultProductId = 1;
    constexpr int kHttpTimeoutMs = 15000;

    void TraceService(std::wstring_view message) noexcept
    {
        OutputDebugStringW(L"[ZIVPO.Service] ");
        OutputDebugStringW(message.data());
        OutputDebugStringW(L"\r\n");
    }

    std::wstring GetEnvValue(std::wstring_view name)
    {
        DWORD required = GetEnvironmentVariableW(name.data(), nullptr, 0);
        if (required == 0)
        {
            return {};
        }

        std::wstring value(required, L'\0');
        DWORD written = GetEnvironmentVariableW(name.data(), value.data(), static_cast<DWORD>(value.size()));
        if (written == 0)
        {
            return {};
        }

        value.resize(written);
        return value;
    }

    bool IsTrueFlag(std::wstring_view value)
    {
        if (value.empty())
        {
            return false;
        }

        std::wstring normalized(value);
        std::transform(normalized.begin(), normalized.end(), normalized.begin(), [](wchar_t ch)
        {
            return static_cast<wchar_t>(std::towlower(ch));
        });

        return normalized == L"1" || normalized == L"true" || normalized == L"yes" || normalized == L"on";
    }

    std::wstring GetApiBaseUrl()
    {
        std::wstring configured = GetEnvValue(L"ZIVPO_API_BASE_URL");
        if (configured.empty())
        {
            return kDefaultApiBaseUrl;
        }

        return configured;
    }

    bool IsInsecureTlsAllowed()
    {
        std::wstring flag = GetEnvValue(L"ZIVPO_ALLOW_INSECURE_TLS");
        if (flag.empty())
        {
            return true;
        }

        std::transform(flag.begin(), flag.end(), flag.begin(), [](wchar_t ch)
        {
            return static_cast<wchar_t>(std::towlower(ch));
        });

        return flag == L"1" || flag == L"true" || flag == L"yes" || flag == L"on";
    }

    std::wstring Utf8ToWide(std::string const& value)
    {
        if (value.empty())
        {
            return {};
        }

        int required = MultiByteToWideChar(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), nullptr, 0);
        if (required <= 0)
        {
            return {};
        }

        std::wstring result(required, L'\0');
        MultiByteToWideChar(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), result.data(), required);
        return result;
    }

    std::string WideToUtf8(std::wstring_view value)
    {
        if (value.empty())
        {
            return {};
        }

        int required = WideCharToMultiByte(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), nullptr, 0, nullptr, nullptr);
        if (required <= 0)
        {
            return {};
        }

        std::string result(required, '\0');
        WideCharToMultiByte(CP_UTF8, 0, value.data(), static_cast<int>(value.size()), result.data(), required, nullptr, nullptr);
        return result;
    }

    std::optional<std::string> ExtractJsonString(std::string const& json, std::string const& fieldName)
    {
        std::regex pattern("\"" + fieldName + "\"\\s*:\\s*\"([^\"]*)\"");
        std::smatch match;
        if (std::regex_search(json, match, pattern) && match.size() > 1)
        {
            return match[1].str();
        }

        return std::nullopt;
    }

    std::optional<long long> ExtractJsonInt64(std::string const& json, std::string const& fieldName)
    {
        std::regex pattern("\"" + fieldName + "\"\\s*:\\s*(-?\\d+)");
        std::smatch match;
        if (std::regex_search(json, match, pattern) && match.size() > 1)
        {
            try
            {
                return std::stoll(match[1].str());
            }
            catch (...)
            {
                return std::nullopt;
            }
        }

        return std::nullopt;
    }

    std::optional<bool> ExtractJsonBool(std::string const& json, std::string const& fieldName)
    {
        std::regex pattern("\"" + fieldName + "\"\\s*:\\s*(true|false)");
        std::smatch match;
        if (std::regex_search(json, match, pattern) && match.size() > 1)
        {
            return match[1].str() == "true";
        }

        return std::nullopt;
    }

    std::optional<std::chrono::system_clock::time_point> ParseIsoDateTimeLocal(std::string const& value)
    {
        if (value.size() < 19)
        {
            return std::nullopt;
        }

        std::tm tm{};
        std::istringstream stream(value.substr(0, 19));
        stream >> std::get_time(&tm, "%Y-%m-%dT%H:%M:%S");
        if (stream.fail())
        {
            return std::nullopt;
        }

        time_t localTime = mktime(&tm);
        if (localTime == static_cast<time_t>(-1))
        {
            return std::nullopt;
        }

        return std::chrono::system_clock::from_time_t(localTime);
    }

    ZIVPO::Service::RpcStatusCode StatusCodeFromRpcLong(long status)
    {
        switch (status)
        {
        case ZIVPO_RPC_STATUS_OK:
            return ZIVPO::Service::RpcStatusCode::Ok;
        case ZIVPO_RPC_STATUS_NOT_AUTHENTICATED:
            return ZIVPO::Service::RpcStatusCode::NotAuthenticated;
        case ZIVPO_RPC_STATUS_INVALID_CREDENTIALS:
            return ZIVPO::Service::RpcStatusCode::InvalidCredentials;
        case ZIVPO_RPC_STATUS_NO_LICENSE:
            return ZIVPO::Service::RpcStatusCode::NoLicense;
        case ZIVPO_RPC_STATUS_ACTIVATION_FAILED:
            return ZIVPO::Service::RpcStatusCode::ActivationFailed;
        case ZIVPO_RPC_STATUS_NETWORK_ERROR:
            return ZIVPO::Service::RpcStatusCode::NetworkError;
        default:
            return ZIVPO::Service::RpcStatusCode::InternalError;
        }
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
        if (IsTrueFlag(GetEnvValue(L"ZIVPO_ALLOW_DIRECT_GUI")))
        {
            return GuiStartupDecision::Continue;
        }

        SC_HANDLE scm = OpenSCManagerW(nullptr, nullptr, SC_MANAGER_CONNECT | SC_MANAGER_CREATE_SERVICE);
        if (scm == nullptr)
        {
            scm = OpenSCManagerW(nullptr, nullptr, SC_MANAGER_CONNECT);
            if (scm == nullptr)
            {
                return GuiStartupDecision::Error;
            }
        }

        // Query service state using minimal access first. A GUI process launched in a
        // user session often has no SERVICE_START/SERVICE_CHANGE_CONFIG rights.
        SC_HANDLE serviceHandle = OpenServiceWithAccess(scm, SERVICE_QUERY_STATUS);
        DWORD openError = (serviceHandle == nullptr) ? GetLastError() : ERROR_SUCCESS;

        if (serviceHandle == nullptr && openError == ERROR_SERVICE_DOES_NOT_EXIST)
        {
            serviceHandle = CreateServiceIfMissing(scm);
        }

        if (serviceHandle == nullptr)
        {
            serviceHandle = OpenServiceWithAccess(scm, SERVICE_QUERY_STATUS | SERVICE_START);
        }

        if (serviceHandle == nullptr)
        {
            CloseServiceHandle(scm);
            return GuiStartupDecision::Error;
        }

        bool observedStopped = false;
        bool startedByCaller = false;

        SERVICE_STATUS_PROCESS status{};
        bool running = QueryServiceStatusProcess(serviceHandle, status);
        if (running)
        {
            if (status.dwCurrentState == SERVICE_RUNNING)
            {
                running = true;
            }
            else if (status.dwCurrentState == SERVICE_START_PENDING)
            {
                running = WaitForServiceRunning(serviceHandle, kServiceStartTimeoutMs);
            }
            else
            {
                observedStopped = true;
                CloseServiceHandle(serviceHandle);
                serviceHandle = OpenServiceWithAccess(scm, SERVICE_QUERY_STATUS | SERVICE_START);
                if (serviceHandle == nullptr)
                {
                    CloseServiceHandle(scm);
                    return GuiStartupDecision::Error;
                }

                running = EnsureServiceRunning(serviceHandle, startedByCaller);
            }
        }

        CloseServiceHandle(serviceHandle);
        CloseServiceHandle(scm);

        if (!running)
        {
            return GuiStartupDecision::Error;
        }

        // Assignment requirement: if service was stopped at startup, GUI must start it and exit.
        if (startedByCaller || observedStopped)
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
    constexpr int kExplorerWaitMaxAttempts = 240;
    constexpr int kUserTokenWaitMaxAttempts = 240;
    constexpr DWORD kRetryDelayMs = 500;
    constexpr int kAsyncLaunchRetryAttempts = 120;
    constexpr DWORD kAsyncLaunchRetryDelayMs = 2000;
    constexpr DWORD kSessionMonitorIntervalMs = 5000;

    SERVICE_STATUS_HANDLE g_statusHandle = nullptr;
    SERVICE_STATUS g_serviceStatus{};
    DWORD g_checkPoint = 1;
    DWORD g_serviceProcessId = 0;
    bool g_rpcStopRequested = false;
    std::atomic<bool> g_serviceStopping{ false };
    std::mutex g_launchedProcessesMutex;
    std::unordered_map<DWORD, DWORD> g_launchedGuiProcessBySession;
    std::mutex g_launchSlotsMutex;
    std::unordered_set<DWORD> g_launchSlotsBySession;
    std::mutex g_retrySessionsMutex;
    std::unordered_set<DWORD> g_retrySessions;
    HANDLE g_sessionMonitorStopEvent = nullptr;
    HANDLE g_sessionMonitorThread = nullptr;
    HANDLE g_authWorkerStopEvent = nullptr;
    HANDLE g_authWorkerThread = nullptr;

    struct ApiUrlParts
    {
        std::wstring host;
        INTERNET_PORT port{ INTERNET_DEFAULT_HTTPS_PORT };
        std::wstring basePath;
        bool secure{ true };
    };

    struct ServiceAuthState
    {
        bool authenticated{ false };
        std::wstring username;
        std::wstring accessToken;
        std::wstring refreshToken;
        std::chrono::system_clock::time_point accessTokenExpiresAt{};
        std::chrono::system_clock::time_point refreshTokenExpiresAt{};
        bool hasLicenseTicket{ false };
        bool licenseBlocked{ false };
        std::wstring licenseExpirationDate;
        std::chrono::system_clock::time_point licenseRefreshAt{};
    };

    std::mutex g_authStateMutex;
    ServiceAuthState g_authState{};

    void ClearAuthStateLocked(ServiceAuthState& state)
    {
        state = ServiceAuthState{};
    }

    void SetRpcUserInfo(ZIVPO_RPC_USER_INFO* userInfo, ServiceAuthState const& state)
    {
        if (userInfo == nullptr)
        {
            return;
        }

        userInfo->authenticated = state.authenticated ? 1 : 0;
        userInfo->username[0] = L'\0';
        if (!state.username.empty())
        {
            StringCchCopyW(userInfo->username, ARRAYSIZE(userInfo->username), state.username.c_str());
        }
    }

    void SetRpcLicenseInfo(ZIVPO_RPC_LICENSE_INFO* licenseInfo, ServiceAuthState const& state)
    {
        if (licenseInfo == nullptr)
        {
            return;
        }

        licenseInfo->hasLicense = state.hasLicenseTicket ? 1 : 0;
        licenseInfo->licenseBlocked = state.licenseBlocked ? 1 : 0;
        licenseInfo->expirationDate[0] = L'\0';
        if (!state.licenseExpirationDate.empty())
        {
            StringCchCopyW(licenseInfo->expirationDate, ARRAYSIZE(licenseInfo->expirationDate), state.licenseExpirationDate.c_str());
        }
    }

    std::chrono::system_clock::time_point EpochSecondsToTimePoint(long long secondsSinceEpoch)
    {
        return std::chrono::system_clock::time_point(std::chrono::seconds(secondsSinceEpoch));
    }

    std::optional<long long> GetJwtExpClaim(std::wstring const& token)
    {
        std::string utf8Token = WideToUtf8(token);
        size_t firstDot = utf8Token.find('.');
        if (firstDot == std::string::npos)
        {
            return std::nullopt;
        }

        size_t secondDot = utf8Token.find('.', firstDot + 1);
        if (secondDot == std::string::npos || secondDot <= firstDot + 1)
        {
            return std::nullopt;
        }

        std::string payload = utf8Token.substr(firstDot + 1, secondDot - firstDot - 1);
        std::replace(payload.begin(), payload.end(), '-', '+');
        std::replace(payload.begin(), payload.end(), '_', '/');
        while ((payload.size() % 4) != 0)
        {
            payload.push_back('=');
        }

        static char const decodeTable[] =
            "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
        std::array<int, 256> lookup{};
        lookup.fill(-1);
        for (int i = 0; i < 64; ++i)
        {
            lookup[static_cast<unsigned char>(decodeTable[i])] = i;
        }

        std::string decoded;
        decoded.reserve((payload.size() * 3) / 4);

        int val = 0;
        int bits = -8;
        for (unsigned char c : payload)
        {
            if (c == '=')
            {
                break;
            }

            int const index = lookup[c];
            if (index < 0)
            {
                return std::nullopt;
            }

            val = (val << 6) + index;
            bits += 6;
            if (bits >= 0)
            {
                decoded.push_back(static_cast<char>((val >> bits) & 0xFF));
                bits -= 8;
            }
        }

        return ExtractJsonInt64(decoded, "exp");
    }

    bool ParseApiUrlParts(std::wstring const& baseUrl, ApiUrlParts& parts)
    {
        URL_COMPONENTS components{};
        components.dwStructSize = sizeof(components);

        std::wstring mutableUrl = baseUrl;
        wchar_t hostBuffer[256]{};
        wchar_t pathBuffer[1024]{};
        components.lpszHostName = hostBuffer;
        components.dwHostNameLength = ARRAYSIZE(hostBuffer);
        components.lpszUrlPath = pathBuffer;
        components.dwUrlPathLength = ARRAYSIZE(pathBuffer);

        if (!WinHttpCrackUrl(mutableUrl.c_str(), static_cast<DWORD>(mutableUrl.size()), 0, &components))
        {
            return false;
        }

        parts.host.assign(components.lpszHostName, components.dwHostNameLength);
        parts.port = components.nPort;
        parts.secure = components.nScheme == INTERNET_SCHEME_HTTPS;
        parts.basePath.assign(components.lpszUrlPath, components.dwUrlPathLength);
        if (parts.basePath.empty())
        {
            parts.basePath = L"";
        }
        if (!parts.basePath.empty() && parts.basePath.back() == L'/')
        {
            parts.basePath.pop_back();
        }
        return !parts.host.empty();
    }

    struct HttpResult
    {
        bool ok{ false };
        DWORD statusCode{ 0 };
        std::string body;
    };

    HttpResult SendJsonRequest(
        ApiUrlParts const& api,
        std::wstring const& method,
        std::wstring const& path,
        std::string const& requestBodyUtf8,
        std::wstring const& bearerToken)
    {
        HttpResult result{};

        HINTERNET session = WinHttpOpen(
            L"ZIVPO.Service/1.0",
            WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY,
            WINHTTP_NO_PROXY_NAME,
            WINHTTP_NO_PROXY_BYPASS,
            0);
        if (session == nullptr)
        {
            return result;
        }

        WinHttpSetTimeouts(session, kHttpTimeoutMs, kHttpTimeoutMs, kHttpTimeoutMs, kHttpTimeoutMs);

        HINTERNET connect = WinHttpConnect(session, api.host.c_str(), api.port, 0);
        if (connect == nullptr)
        {
            WinHttpCloseHandle(session);
            return result;
        }

        std::wstring fullPath = api.basePath + path;
        DWORD requestFlags = api.secure ? WINHTTP_FLAG_SECURE : 0;
        HINTERNET request = WinHttpOpenRequest(
            connect,
            method.c_str(),
            fullPath.c_str(),
            nullptr,
            WINHTTP_NO_REFERER,
            WINHTTP_DEFAULT_ACCEPT_TYPES,
            requestFlags);
        if (request == nullptr)
        {
            WinHttpCloseHandle(connect);
            WinHttpCloseHandle(session);
            return result;
        }

        if (api.secure && IsInsecureTlsAllowed())
        {
            DWORD securityFlags =
                SECURITY_FLAG_IGNORE_UNKNOWN_CA |
                SECURITY_FLAG_IGNORE_CERT_DATE_INVALID |
                SECURITY_FLAG_IGNORE_CERT_CN_INVALID |
                SECURITY_FLAG_IGNORE_CERT_WRONG_USAGE;
            WinHttpSetOption(request, WINHTTP_OPTION_SECURITY_FLAGS, &securityFlags, sizeof(securityFlags));
        }

        std::wstring headers = L"Content-Type: application/json\r\nAccept: application/json\r\n";
        if (!bearerToken.empty())
        {
            headers.append(L"Authorization: Bearer ");
            headers.append(bearerToken);
            headers.append(L"\r\n");
        }

        BOOL sent = WinHttpSendRequest(
            request,
            headers.c_str(),
            static_cast<DWORD>(headers.size()),
            requestBodyUtf8.empty() ? WINHTTP_NO_REQUEST_DATA : const_cast<char*>(requestBodyUtf8.data()),
            static_cast<DWORD>(requestBodyUtf8.size()),
            static_cast<DWORD>(requestBodyUtf8.size()),
            0);
        if (!sent || !WinHttpReceiveResponse(request, nullptr))
        {
            WinHttpCloseHandle(request);
            WinHttpCloseHandle(connect);
            WinHttpCloseHandle(session);
            return result;
        }

        DWORD statusCode = 0;
        DWORD statusCodeSize = sizeof(statusCode);
        WinHttpQueryHeaders(
            request,
            WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
            WINHTTP_HEADER_NAME_BY_INDEX,
            &statusCode,
            &statusCodeSize,
            WINHTTP_NO_HEADER_INDEX);
        result.statusCode = statusCode;

        std::string body;
        DWORD availableBytes = 0;
        while (WinHttpQueryDataAvailable(request, &availableBytes) && availableBytes > 0)
        {
            std::string chunk(availableBytes, '\0');
            DWORD downloaded = 0;
            if (!WinHttpReadData(request, chunk.data(), availableBytes, &downloaded))
            {
                break;
            }
            chunk.resize(downloaded);
            body.append(chunk);
            if (downloaded == 0)
            {
                break;
            }
        }

        result.ok = true;
        result.body = std::move(body);

        WinHttpCloseHandle(request);
        WinHttpCloseHandle(connect);
        WinHttpCloseHandle(session);
        return result;
    }

    std::wstring GetDeviceName();

    std::wstring GetDeviceMacAddress()
    {
        std::wstring deviceName = GetDeviceName();
        size_t hash = std::hash<std::wstring>{}(deviceName);
        unsigned char bytes[6]{};
        for (int i = 0; i < 6; ++i)
        {
            bytes[i] = static_cast<unsigned char>((hash >> (i * 8)) & 0xFF);
        }

        wchar_t pseudoMac[32]{};
        StringCchPrintfW(
            pseudoMac,
            ARRAYSIZE(pseudoMac),
            L"%02X-%02X-%02X-%02X-%02X-%02X",
            bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5]);
        return pseudoMac;
    }

    std::wstring GetDeviceName()
    {
        wchar_t computerName[MAX_COMPUTERNAME_LENGTH + 1]{};
        DWORD size = ARRAYSIZE(computerName);
        if (GetComputerNameW(computerName, &size))
        {
            return computerName;
        }

        return L"UnknownDevice";
    }

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

    bool AcquireRetrySlot(DWORD sessionId)
    {
        std::lock_guard lock(g_retrySessionsMutex);
        return g_retrySessions.insert(sessionId).second;
    }

    void ReleaseRetrySlot(DWORD sessionId)
    {
        std::lock_guard lock(g_retrySessionsMutex);
        g_retrySessions.erase(sessionId);
    }

    bool TryAcquireLaunchSlot(DWORD sessionId)
    {
        std::lock_guard lock(g_launchSlotsMutex);
        return g_launchSlotsBySession.insert(sessionId).second;
    }

    void ReleaseLaunchSlot(DWORD sessionId)
    {
        std::lock_guard lock(g_launchSlotsMutex);
        g_launchSlotsBySession.erase(sessionId);
    }

    struct LaunchSlotGuard final
    {
        explicit LaunchSlotGuard(DWORD session) noexcept
            : sessionId(session), acquired(TryAcquireLaunchSlot(session))
        {
        }

        ~LaunchSlotGuard()
        {
            if (acquired)
            {
                ReleaseLaunchSlot(sessionId);
            }
        }

        DWORD sessionId{ 0 };
        bool acquired{ false };
    };

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
        for (int attempt = 0; attempt < kExplorerWaitMaxAttempts; ++attempt)
        {
            if (IsExplorerRunningInSession(sessionId))
            {
                return;
            }

            if (g_serviceStopping.load())
            {
                return;
            }

            Sleep(kRetryDelayMs);
        }
    }

    HANDLE WaitForUserToken(DWORD sessionId)
    {
        for (int attempt = 0; attempt < kUserTokenWaitMaxAttempts; ++attempt)
        {
            HANDLE token = nullptr;
            if (WTSQueryUserToken(sessionId, &token))
            {
                return token;
            }

            if (g_serviceStopping.load())
            {
                return nullptr;
            }

            Sleep(kRetryDelayMs);
        }

        return nullptr;
    }

    bool LaunchGuiHiddenInSession(DWORD sessionId)
    {
        if (g_serviceStopping.load())
        {
            return false;
        }

        if (sessionId == 0 || IsSessionGuiAlreadyRunning(sessionId))
        {
            return true;
        }

        // Multiple service paths (startup scan, monitor loop, session notifications)
        // can race and launch duplicate GUI processes for the same session.
        // Serialize launch attempts per session.
        LaunchSlotGuard launchGuard(sessionId);
        if (!launchGuard.acquired)
        {
            return true;
        }

        WaitForExplorerInSession(sessionId);

        HANDLE impersonationToken = WaitForUserToken(sessionId);
        if (impersonationToken == nullptr)
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

    DWORD WINAPI RetryLaunchThreadProc(LPVOID context)
    {
        DWORD sessionId = static_cast<DWORD>(reinterpret_cast<ULONG_PTR>(context));

        for (int attempt = 0; attempt < kAsyncLaunchRetryAttempts; ++attempt)
        {
            if (g_serviceStopping.load())
            {
                break;
            }

            if (IsSessionGuiAlreadyRunning(sessionId) || LaunchGuiHiddenInSession(sessionId))
            {
                break;
            }

            Sleep(kAsyncLaunchRetryDelayMs);
        }

        ReleaseRetrySlot(sessionId);
        return 0;
    }

    void ScheduleLaunchRetry(DWORD sessionId)
    {
        if (sessionId == 0 || g_serviceStopping.load())
        {
            return;
        }

        if (!AcquireRetrySlot(sessionId))
        {
            return;
        }

        HANDLE retryThread = CreateThread(
            nullptr,
            0,
            &RetryLaunchThreadProc,
            reinterpret_cast<LPVOID>(static_cast<ULONG_PTR>(sessionId)),
            0,
            nullptr);

        if (retryThread == nullptr)
        {
            ReleaseRetrySlot(sessionId);
            return;
        }

        CloseHandle(retryThread);
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
                if (!LaunchGuiHiddenInSession(session.SessionId))
                {
                    ScheduleLaunchRetry(session.SessionId);
                }
            }
        }

        WTSFreeMemory(sessions);
    }

    DWORD WINAPI SessionMonitorThreadProc([[maybe_unused]] LPVOID context)
    {
        while (!g_serviceStopping.load())
        {
            LaunchGuiInAllActiveSessions();

            if (g_sessionMonitorStopEvent == nullptr)
            {
                Sleep(kSessionMonitorIntervalMs);
                continue;
            }

            DWORD waitResult = WaitForSingleObject(g_sessionMonitorStopEvent, kSessionMonitorIntervalMs);
            if (waitResult == WAIT_OBJECT_0)
            {
                break;
            }
        }

        return 0;
    }

    void StartSessionMonitorThread()
    {
        if (g_sessionMonitorStopEvent == nullptr || g_sessionMonitorThread != nullptr)
        {
            return;
        }

        g_sessionMonitorThread = CreateThread(
            nullptr,
            0,
            &SessionMonitorThreadProc,
            nullptr,
            0,
            nullptr);
    }

    void StopSessionMonitorThread()
    {
        if (g_sessionMonitorStopEvent != nullptr)
        {
            SetEvent(g_sessionMonitorStopEvent);
        }

        if (g_sessionMonitorThread != nullptr)
        {
            WaitForSingleObject(g_sessionMonitorThread, 10000);
            CloseHandle(g_sessionMonitorThread);
            g_sessionMonitorThread = nullptr;
        }

        if (g_sessionMonitorStopEvent != nullptr)
        {
            CloseHandle(g_sessionMonitorStopEvent);
            g_sessionMonitorStopEvent = nullptr;
        }
    }

    std::string EscapeJsonString(std::wstring_view value)
    {
        std::string utf8 = WideToUtf8(value);
        std::string escaped;
        escaped.reserve(utf8.size() + 8);

        for (char ch : utf8)
        {
            switch (ch)
            {
            case '\\':
                escaped.append("\\\\");
                break;
            case '"':
                escaped.append("\\\"");
                break;
            case '\n':
                escaped.append("\\n");
                break;
            case '\r':
                escaped.append("\\r");
                break;
            case '\t':
                escaped.append("\\t");
                break;
            default:
                escaped.push_back(ch);
                break;
            }
        }

        return escaped;
    }

    bool ParseTokenPairResponse(
        std::string const& responseJson,
        std::wstring& accessToken,
        std::wstring& refreshToken)
    {
        std::optional<std::string> access = ExtractJsonString(responseJson, "accessToken");
        std::optional<std::string> refresh = ExtractJsonString(responseJson, "refreshToken");
        if (!access.has_value() || !refresh.has_value())
        {
            return false;
        }

        accessToken = Utf8ToWide(*access);
        refreshToken = Utf8ToWide(*refresh);
        return !accessToken.empty() && !refreshToken.empty();
    }

    bool ParseLicenseTicketResponse(
        std::string const& responseJson,
        bool& blocked,
        std::wstring& expirationDate,
        std::chrono::system_clock::time_point& licenseRefreshAt)
    {
        std::optional<std::string> expiration = ExtractJsonString(responseJson, "licenseExpirationDate");
        std::optional<bool> blockedValue = ExtractJsonBool(responseJson, "licenseBlocked");
        std::optional<std::string> serverDate = ExtractJsonString(responseJson, "serverDate");
        std::optional<long long> ticketLifetimeSeconds = ExtractJsonInt64(responseJson, "ticketLifetimeSeconds");

        if (!expiration.has_value() || !blockedValue.has_value() || !ticketLifetimeSeconds.has_value())
        {
            return false;
        }

        blocked = *blockedValue;
        expirationDate = Utf8ToWide(*expiration);

        auto now = std::chrono::system_clock::now();
        auto baseTime = now;
        if (serverDate.has_value())
        {
            std::optional<std::chrono::system_clock::time_point> parsedServer = ParseIsoDateTimeLocal(*serverDate);
            if (parsedServer.has_value())
            {
                baseTime = *parsedServer;
            }
        }

        long long lifetime = std::max<long long>(10, *ticketLifetimeSeconds);
        long long refreshDelta = std::max<long long>(5, lifetime - kLicenseRefreshSafetyWindowSeconds);
        licenseRefreshAt = baseTime + std::chrono::seconds(refreshDelta);
        if (licenseRefreshAt <= now)
        {
            licenseRefreshAt = now + std::chrono::seconds(5);
        }

        return true;
    }

    long CheckLicenseWithAccessToken(
        ApiUrlParts const& api,
        std::wstring const& accessToken,
        std::wstring const& deviceMac,
        bool& blocked,
        std::wstring& expirationDate,
        std::chrono::system_clock::time_point& licenseRefreshAt)
    {
        std::string requestBody =
            "{\"deviceMac\":\"" + EscapeJsonString(deviceMac) +
            "\",\"productId\":" + std::to_string(kDefaultProductId) + "}";

        HttpResult response = SendJsonRequest(api, L"POST", kApiLicenseCheckPath, requestBody, accessToken);
        if (!response.ok)
        {
            return ZIVPO_RPC_STATUS_NETWORK_ERROR;
        }

        if (response.statusCode == 401 || response.statusCode == 403)
        {
            return ZIVPO_RPC_STATUS_NOT_AUTHENTICATED;
        }

        if (response.statusCode == 404 || response.statusCode == 400)
        {
            return ZIVPO_RPC_STATUS_NO_LICENSE;
        }

        if (response.statusCode != 200)
        {
            return ZIVPO_RPC_STATUS_INTERNAL_ERROR;
        }

        if (!ParseLicenseTicketResponse(response.body, blocked, expirationDate, licenseRefreshAt))
        {
            return ZIVPO_RPC_STATUS_INTERNAL_ERROR;
        }

        return ZIVPO_RPC_STATUS_OK;
    }

    long GetUsernameFromAccessToken(ApiUrlParts const& api, std::wstring const& accessToken, std::wstring& username)
    {
        HttpResult response = SendJsonRequest(api, L"GET", kApiUserMePath, "", accessToken);
        if (!response.ok)
        {
            return ZIVPO_RPC_STATUS_NETWORK_ERROR;
        }

        if (response.statusCode == 401 || response.statusCode == 403)
        {
            return ZIVPO_RPC_STATUS_NOT_AUTHENTICATED;
        }

        if (response.statusCode != 200)
        {
            return ZIVPO_RPC_STATUS_INTERNAL_ERROR;
        }

        std::optional<std::string> extractedUsername = ExtractJsonString(response.body, "username");
        if (!extractedUsername.has_value())
        {
            return ZIVPO_RPC_STATUS_INTERNAL_ERROR;
        }

        username = Utf8ToWide(*extractedUsername);
        return username.empty() ? ZIVPO_RPC_STATUS_INTERNAL_ERROR : ZIVPO_RPC_STATUS_OK;
    }

    long RefreshTokens(ApiUrlParts const& api, std::wstring const& refreshToken, std::wstring& accessTokenOut, std::wstring& refreshTokenOut)
    {
        std::string requestBody =
            "{\"refreshToken\":\"" + EscapeJsonString(refreshToken) + "\"}";

        HttpResult response = SendJsonRequest(api, L"POST", kApiAuthRefreshPath, requestBody, L"");
        if (!response.ok)
        {
            return ZIVPO_RPC_STATUS_NETWORK_ERROR;
        }

        if (response.statusCode == 401)
        {
            return ZIVPO_RPC_STATUS_NOT_AUTHENTICATED;
        }

        if (response.statusCode != 200)
        {
            return ZIVPO_RPC_STATUS_INTERNAL_ERROR;
        }

        if (!ParseTokenPairResponse(response.body, accessTokenOut, refreshTokenOut))
        {
            return ZIVPO_RPC_STATUS_INTERNAL_ERROR;
        }

        return ZIVPO_RPC_STATUS_OK;
    }

    void UpdateTokenExpirations(ServiceAuthState& state)
    {
        auto now = std::chrono::system_clock::now();
        std::optional<long long> accessExp = GetJwtExpClaim(state.accessToken);
        std::optional<long long> refreshExp = GetJwtExpClaim(state.refreshToken);

        state.accessTokenExpiresAt = accessExp.has_value()
            ? EpochSecondsToTimePoint(*accessExp)
            : (now + std::chrono::minutes(10));
        state.refreshTokenExpiresAt = refreshExp.has_value()
            ? EpochSecondsToTimePoint(*refreshExp)
            : (now + std::chrono::hours(12));
    }

    void EnsureAuthWorkerStarted()
    {
        if (g_authWorkerStopEvent == nullptr || g_authWorkerThread != nullptr)
        {
            return;
        }

        auto workerProc = [](LPVOID) -> DWORD
        {
            ApiUrlParts api{};
            if (!ParseApiUrlParts(GetApiBaseUrl(), api))
            {
                api.host = L"localhost";
                api.port = 8444;
                api.secure = true;
            }

            std::wstring deviceMac = GetDeviceMacAddress();

            while (!g_serviceStopping.load())
            {
                ServiceAuthState snapshot{};
                {
                    std::lock_guard lock(g_authStateMutex);
                    snapshot = g_authState;
                }

                DWORD waitMs = 30000;
                auto now = std::chrono::system_clock::now();

                if (snapshot.authenticated)
                {
                    auto nextTokenRefresh = snapshot.accessTokenExpiresAt - std::chrono::seconds(kTokenRefreshSafetyWindowSeconds);
                    if (nextTokenRefresh <= now)
                    {
                        std::wstring newAccessToken;
                        std::wstring newRefreshToken;
                        long refreshStatus = RefreshTokens(api, snapshot.refreshToken, newAccessToken, newRefreshToken);
                        if (refreshStatus == ZIVPO_RPC_STATUS_OK)
                        {
                            std::lock_guard lock(g_authStateMutex);
                            g_authState.accessToken = std::move(newAccessToken);
                            g_authState.refreshToken = std::move(newRefreshToken);
                            UpdateTokenExpirations(g_authState);
                        }
                        else if (refreshStatus == ZIVPO_RPC_STATUS_NOT_AUTHENTICATED)
                        {
                            std::lock_guard lock(g_authStateMutex);
                            ClearAuthStateLocked(g_authState);
                        }
                    }

                    if (snapshot.hasLicenseTicket)
                    {
                        if (snapshot.licenseRefreshAt <= now)
                        {
                            bool blocked = false;
                            std::wstring expirationDate;
                            std::chrono::system_clock::time_point refreshAt{};
                            long checkStatus = CheckLicenseWithAccessToken(
                                api,
                                snapshot.accessToken,
                                deviceMac,
                                blocked,
                                expirationDate,
                                refreshAt);

                            std::lock_guard lock(g_authStateMutex);
                            if (checkStatus == ZIVPO_RPC_STATUS_OK)
                            {
                                g_authState.hasLicenseTicket = true;
                                g_authState.licenseBlocked = blocked;
                                g_authState.licenseExpirationDate = std::move(expirationDate);
                                g_authState.licenseRefreshAt = refreshAt;
                            }
                            else if (checkStatus == ZIVPO_RPC_STATUS_NO_LICENSE)
                            {
                                g_authState.hasLicenseTicket = false;
                                g_authState.licenseBlocked = false;
                                g_authState.licenseExpirationDate.clear();
                                g_authState.licenseRefreshAt = {};
                            }
                        }

                        auto delay = std::chrono::duration_cast<std::chrono::milliseconds>(snapshot.licenseRefreshAt - now).count();
                        if (delay > 0)
                        {
                            waitMs = static_cast<DWORD>(std::min<long long>(waitMs, delay));
                        }
                    }

                    auto tokenDelay = std::chrono::duration_cast<std::chrono::milliseconds>(
                        (snapshot.accessTokenExpiresAt - std::chrono::seconds(kTokenRefreshSafetyWindowSeconds)) - now).count();
                    if (tokenDelay > 0)
                    {
                        waitMs = static_cast<DWORD>(std::min<long long>(waitMs, tokenDelay));
                    }
                }

                HANDLE waitHandles[2] = { g_authWorkerStopEvent, g_sessionMonitorStopEvent };
                DWORD waitResult = WaitForMultipleObjects(2, waitHandles, FALSE, waitMs);
                if (waitResult == WAIT_OBJECT_0 || waitResult == WAIT_OBJECT_0 + 1)
                {
                    break;
                }
            }

            return 0;
        };

        g_authWorkerThread = CreateThread(nullptr, 0, workerProc, nullptr, 0, nullptr);
    }

    void StopAuthWorkerThread()
    {
        if (g_authWorkerStopEvent != nullptr)
        {
            SetEvent(g_authWorkerStopEvent);
        }

        if (g_authWorkerThread != nullptr)
        {
            WaitForSingleObject(g_authWorkerThread, 10000);
            CloseHandle(g_authWorkerThread);
            g_authWorkerThread = nullptr;
        }

        if (g_authWorkerStopEvent != nullptr)
        {
            CloseHandle(g_authWorkerStopEvent);
            g_authWorkerStopEvent = nullptr;
        }
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
                     eventType == WTS_SESSION_UNLOCK ||
                     eventType == WTS_CONSOLE_CONNECT ||
                     eventType == WTS_REMOTE_CONNECT))
                {
                    if (!LaunchGuiHiddenInSession(notification->dwSessionId))
                    {
                        ScheduleLaunchRetry(notification->dwSessionId);
                    }
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
        g_serviceStopping = false;
        g_sessionMonitorStopEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
        g_sessionMonitorThread = nullptr;
        g_authWorkerStopEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
        g_authWorkerThread = nullptr;
        {
            std::lock_guard lock(g_authStateMutex);
            ClearAuthStateLocked(g_authState);
        }

        UpdateServiceStatus(SERVICE_START_PENDING, NO_ERROR, 5000);

        RPC_STATUS rpcStatus = StartRpcServer();
        if (rpcStatus != RPC_S_OK)
        {
            g_serviceStopping = true;
            StopAuthWorkerThread();
            StopSessionMonitorThread();
            UpdateServiceStatus(SERVICE_STOPPED, rpcStatus, 0);
            return;
        }

        LaunchGuiInAllActiveSessions();
        UpdateServiceStatus(SERVICE_RUNNING, NO_ERROR, 0);
        StartSessionMonitorThread();

        rpcStatus = RunRpcServerLoop();

        g_serviceStopping = true;
        UpdateServiceStatus(SERVICE_STOP_PENDING, NO_ERROR, 5000);
        StopAuthWorkerThread();
        StopSessionMonitorThread();
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

extern "C" long RpcGetCurrentUser([[maybe_unused]] handle_t hBinding, ZIVPO_RPC_USER_INFO* userInfo)
{
    std::lock_guard lock(g_authStateMutex);
    SetRpcUserInfo(userInfo, g_authState);
    return g_authState.authenticated ? ZIVPO_RPC_STATUS_OK : ZIVPO_RPC_STATUS_NOT_AUTHENTICATED;
}

extern "C" long RpcLogin(
    [[maybe_unused]] handle_t hBinding,
    const wchar_t* username,
    const wchar_t* password,
    ZIVPO_RPC_USER_INFO* userInfo,
    ZIVPO_RPC_LICENSE_INFO* licenseInfo)
{
    if (username == nullptr || password == nullptr)
    {
        return ZIVPO_RPC_STATUS_INVALID_CREDENTIALS;
    }

    ApiUrlParts api{};
    if (!ParseApiUrlParts(GetApiBaseUrl(), api))
    {
        return ZIVPO_RPC_STATUS_NETWORK_ERROR;
    }

    std::string requestBody =
        "{\"username\":\"" + EscapeJsonString(username) +
        "\",\"password\":\"" + EscapeJsonString(password) + "\"}";

    HttpResult loginResponse = SendJsonRequest(api, L"POST", kApiAuthLoginPath, requestBody, L"");
    if (!loginResponse.ok)
    {
        return ZIVPO_RPC_STATUS_NETWORK_ERROR;
    }

    if (loginResponse.statusCode == 401)
    {
        return ZIVPO_RPC_STATUS_INVALID_CREDENTIALS;
    }

    if (loginResponse.statusCode != 200)
    {
        return ZIVPO_RPC_STATUS_INTERNAL_ERROR;
    }

    std::wstring accessToken;
    std::wstring refreshToken;
    if (!ParseTokenPairResponse(loginResponse.body, accessToken, refreshToken))
    {
        return ZIVPO_RPC_STATUS_INTERNAL_ERROR;
    }

    std::wstring resolvedUsername;
    if (GetUsernameFromAccessToken(api, accessToken, resolvedUsername) != ZIVPO_RPC_STATUS_OK)
    {
        resolvedUsername = username;
    }

    ServiceAuthState newState{};
    newState.authenticated = true;
    newState.username = std::move(resolvedUsername);
    newState.accessToken = std::move(accessToken);
    newState.refreshToken = std::move(refreshToken);
    UpdateTokenExpirations(newState);

    bool blocked = false;
    std::wstring expirationDate;
    std::chrono::system_clock::time_point refreshAt{};
    std::wstring deviceMac = GetDeviceMacAddress();
    long licenseStatus = CheckLicenseWithAccessToken(
        api,
        newState.accessToken,
        deviceMac,
        blocked,
        expirationDate,
        refreshAt);

    if (licenseStatus == ZIVPO_RPC_STATUS_OK)
    {
        newState.hasLicenseTicket = true;
        newState.licenseBlocked = blocked;
        newState.licenseExpirationDate = std::move(expirationDate);
        newState.licenseRefreshAt = refreshAt;
    }

    {
        std::lock_guard lock(g_authStateMutex);
        g_authState = std::move(newState);
        SetRpcUserInfo(userInfo, g_authState);
        SetRpcLicenseInfo(licenseInfo, g_authState);
    }

    EnsureAuthWorkerStarted();
    return ZIVPO_RPC_STATUS_OK;
}

extern "C" long RpcLogout([[maybe_unused]] handle_t hBinding)
{
    std::lock_guard lock(g_authStateMutex);
    ClearAuthStateLocked(g_authState);
    return ZIVPO_RPC_STATUS_OK;
}

extern "C" long RpcGetLicenseInfo([[maybe_unused]] handle_t hBinding, ZIVPO_RPC_LICENSE_INFO* licenseInfo)
{
    std::lock_guard lock(g_authStateMutex);
    SetRpcLicenseInfo(licenseInfo, g_authState);

    if (!g_authState.authenticated)
    {
        return ZIVPO_RPC_STATUS_NOT_AUTHENTICATED;
    }

    if (!g_authState.hasLicenseTicket)
    {
        return ZIVPO_RPC_STATUS_NO_LICENSE;
    }

    return ZIVPO_RPC_STATUS_OK;
}

extern "C" long RpcActivateProduct(
    [[maybe_unused]] handle_t hBinding,
    const wchar_t* activationKey,
    ZIVPO_RPC_LICENSE_INFO* licenseInfo)
{
    if (activationKey == nullptr || activationKey[0] == L'\0')
    {
        return ZIVPO_RPC_STATUS_ACTIVATION_FAILED;
    }

    ServiceAuthState snapshot{};
    {
        std::lock_guard lock(g_authStateMutex);
        snapshot = g_authState;
    }

    if (!snapshot.authenticated || snapshot.accessToken.empty())
    {
        return ZIVPO_RPC_STATUS_NOT_AUTHENTICATED;
    }

    ApiUrlParts api{};
    if (!ParseApiUrlParts(GetApiBaseUrl(), api))
    {
        return ZIVPO_RPC_STATUS_NETWORK_ERROR;
    }

    std::wstring deviceMac = GetDeviceMacAddress();
    std::wstring deviceName = GetDeviceName();
    std::string requestBody =
        "{\"activationKey\":\"" + EscapeJsonString(activationKey) +
        "\",\"deviceMac\":\"" + EscapeJsonString(deviceMac) +
        "\",\"deviceName\":\"" + EscapeJsonString(deviceName) + "\"}";

    HttpResult activateResponse = SendJsonRequest(
        api,
        L"POST",
        kApiLicenseActivatePath,
        requestBody,
        snapshot.accessToken);
    if (!activateResponse.ok)
    {
        return ZIVPO_RPC_STATUS_NETWORK_ERROR;
    }

    if (activateResponse.statusCode == 401 || activateResponse.statusCode == 403)
    {
        return ZIVPO_RPC_STATUS_NOT_AUTHENTICATED;
    }

    bool blocked = false;
    std::wstring expirationDate;
    std::chrono::system_clock::time_point refreshAt{};
    bool parsedTicket = false;

    if (activateResponse.statusCode == 200)
    {
        parsedTicket = ParseLicenseTicketResponse(
            activateResponse.body,
            blocked,
            expirationDate,
            refreshAt);
    }

    if (!parsedTicket)
    {
        long checkStatus = CheckLicenseWithAccessToken(
            api,
            snapshot.accessToken,
            deviceMac,
            blocked,
            expirationDate,
            refreshAt);
        if (checkStatus != ZIVPO_RPC_STATUS_OK)
        {
            return (checkStatus == ZIVPO_RPC_STATUS_NO_LICENSE)
                ? ZIVPO_RPC_STATUS_ACTIVATION_FAILED
                : checkStatus;
        }
    }

    {
        std::lock_guard lock(g_authStateMutex);
        g_authState.hasLicenseTicket = true;
        g_authState.licenseBlocked = blocked;
        g_authState.licenseExpirationDate = std::move(expirationDate);
        g_authState.licenseRefreshAt = refreshAt;
        SetRpcLicenseInfo(licenseInfo, g_authState);
    }

    return ZIVPO_RPC_STATUS_OK;
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
    namespace
    {
        bool CreateRpcBinding(RPC_BINDING_HANDLE& bindingHandle)
        {
            bindingHandle = nullptr;
            RPC_WSTR bindingString = nullptr;

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
            return status == RPC_S_OK && bindingHandle != nullptr;
        }

        RpcCallResult BuildResultFromRpcStatus(long rpcStatus)
        {
            RpcCallResult result{};
            result.status = StatusCodeFromRpcLong(rpcStatus);
            result.ok = rpcStatus == ZIVPO_RPC_STATUS_OK;
            return result;
        }

        void FillUserInfo(ZIVPO_RPC_USER_INFO const& rpcUser, AuthUserInfo& userInfo)
        {
            userInfo.authenticated = rpcUser.authenticated != 0;
            userInfo.username = rpcUser.username;
        }

        void FillLicenseInfo(ZIVPO_RPC_LICENSE_INFO const& rpcLicense, LicenseInfo& licenseInfo)
        {
            licenseInfo.hasLicense = rpcLicense.hasLicense != 0;
            licenseInfo.blocked = rpcLicense.licenseBlocked != 0;
            licenseInfo.expirationDate = rpcLicense.expirationDate;
        }

        long RpcStopServiceSafe(RPC_BINDING_HANDLE bindingHandle) noexcept
        {
            long rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            __try
            {
                RpcStopService(bindingHandle);
                rpcStatus = ZIVPO_RPC_STATUS_OK;
            }
            __except (EXCEPTION_EXECUTE_HANDLER)
            {
                rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            }

            return rpcStatus;
        }

        long RpcGetCurrentUserSafe(RPC_BINDING_HANDLE bindingHandle, ZIVPO_RPC_USER_INFO* userInfo) noexcept
        {
            long rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            __try
            {
                rpcStatus = RpcGetCurrentUser(bindingHandle, userInfo);
            }
            __except (EXCEPTION_EXECUTE_HANDLER)
            {
                rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            }

            return rpcStatus;
        }

        long RpcLoginSafe(
            RPC_BINDING_HANDLE bindingHandle,
            wchar_t const* username,
            wchar_t const* password,
            ZIVPO_RPC_USER_INFO* userInfo,
            ZIVPO_RPC_LICENSE_INFO* licenseInfo) noexcept
        {
            long rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            __try
            {
                rpcStatus = RpcLogin(bindingHandle, username, password, userInfo, licenseInfo);
            }
            __except (EXCEPTION_EXECUTE_HANDLER)
            {
                rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            }

            return rpcStatus;
        }

        long RpcLogoutSafe(RPC_BINDING_HANDLE bindingHandle) noexcept
        {
            long rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            __try
            {
                rpcStatus = RpcLogout(bindingHandle);
            }
            __except (EXCEPTION_EXECUTE_HANDLER)
            {
                rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            }

            return rpcStatus;
        }

        long RpcGetLicenseInfoSafe(RPC_BINDING_HANDLE bindingHandle, ZIVPO_RPC_LICENSE_INFO* licenseInfo) noexcept
        {
            long rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            __try
            {
                rpcStatus = RpcGetLicenseInfo(bindingHandle, licenseInfo);
            }
            __except (EXCEPTION_EXECUTE_HANDLER)
            {
                rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            }

            return rpcStatus;
        }

        long RpcActivateProductSafe(
            RPC_BINDING_HANDLE bindingHandle,
            wchar_t const* activationKey,
            ZIVPO_RPC_LICENSE_INFO* licenseInfo) noexcept
        {
            long rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            __try
            {
                rpcStatus = RpcActivateProduct(bindingHandle, activationKey, licenseInfo);
            }
            __except (EXCEPTION_EXECUTE_HANDLER)
            {
                rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            }

            return rpcStatus;
        }
    }

    int RunServiceMode()
    {
        return static_cast<int>(ERROR_CALL_NOT_IMPLEMENTED);
    }

    bool RequestServiceStop()
    {
        RPC_BINDING_HANDLE bindingHandle = nullptr;
        if (!CreateRpcBinding(bindingHandle))
        {
            return false;
        }

        long rpcStatus = RpcStopServiceSafe(bindingHandle);
        RpcBindingFree(&bindingHandle);
        return rpcStatus == ZIVPO_RPC_STATUS_OK;
    }

    RpcCallResult GetCurrentUser(AuthUserInfo& userInfo)
    {
        userInfo = {};
        RPC_BINDING_HANDLE bindingHandle = nullptr;
        if (!CreateRpcBinding(bindingHandle))
        {
            return { RpcStatusCode::NetworkError, false };
        }

        ZIVPO_RPC_USER_INFO rpcUser{};
        long rpcStatus = RpcGetCurrentUserSafe(bindingHandle, &rpcUser);
        RpcBindingFree(&bindingHandle);
        FillUserInfo(rpcUser, userInfo);
        return BuildResultFromRpcStatus(rpcStatus);
    }

    RpcCallResult Login(std::wstring_view username, std::wstring_view password, AuthUserInfo& userInfo, LicenseInfo& licenseInfo)
    {
        userInfo = {};
        licenseInfo = {};
        RPC_BINDING_HANDLE bindingHandle = nullptr;
        if (!CreateRpcBinding(bindingHandle))
        {
            return { RpcStatusCode::NetworkError, false };
        }

        ZIVPO_RPC_USER_INFO rpcUser{};
        ZIVPO_RPC_LICENSE_INFO rpcLicense{};
        std::wstring usernameValue(username);
        std::wstring passwordValue(password);

        long rpcStatus = RpcLoginSafe(
            bindingHandle,
            usernameValue.c_str(),
            passwordValue.c_str(),
            &rpcUser,
            &rpcLicense);
        RpcBindingFree(&bindingHandle);
        FillUserInfo(rpcUser, userInfo);
        FillLicenseInfo(rpcLicense, licenseInfo);
        return BuildResultFromRpcStatus(rpcStatus);
    }

    RpcCallResult Logout()
    {
        RPC_BINDING_HANDLE bindingHandle = nullptr;
        if (!CreateRpcBinding(bindingHandle))
        {
            return { RpcStatusCode::NetworkError, false };
        }

        long rpcStatus = RpcLogoutSafe(bindingHandle);
        RpcBindingFree(&bindingHandle);
        return BuildResultFromRpcStatus(rpcStatus);
    }

    RpcCallResult GetLicenseInfo(LicenseInfo& licenseInfo)
    {
        licenseInfo = {};
        RPC_BINDING_HANDLE bindingHandle = nullptr;
        if (!CreateRpcBinding(bindingHandle))
        {
            return { RpcStatusCode::NetworkError, false };
        }

        ZIVPO_RPC_LICENSE_INFO rpcLicense{};
        long rpcStatus = RpcGetLicenseInfoSafe(bindingHandle, &rpcLicense);
        RpcBindingFree(&bindingHandle);
        FillLicenseInfo(rpcLicense, licenseInfo);
        return BuildResultFromRpcStatus(rpcStatus);
    }

    RpcCallResult ActivateProduct(std::wstring_view activationKey, LicenseInfo& licenseInfo)
    {
        licenseInfo = {};
        RPC_BINDING_HANDLE bindingHandle = nullptr;
        if (!CreateRpcBinding(bindingHandle))
        {
            return { RpcStatusCode::NetworkError, false };
        }

        ZIVPO_RPC_LICENSE_INFO rpcLicense{};
        std::wstring activationKeyValue(activationKey);

        long rpcStatus = RpcActivateProductSafe(bindingHandle, activationKeyValue.c_str(), &rpcLicense);
        RpcBindingFree(&bindingHandle);
        FillLicenseInfo(rpcLicense, licenseInfo);
        return BuildResultFromRpcStatus(rpcStatus);
    }
}
#endif
