#include "pch.h"
#include "ServiceManager.h"

#include "ServiceRpc.h"

#include <rpc.h>
#include <bcrypt.h>
#include <tlhelp32.h>
#include <userenv.h>
#include <wincrypt.h>
#include <winhttp.h>
#include <wtsapi32.h>

#include <algorithm>
#include <array>
#include <chrono>
#include <cctype>
#include <cstdint>
#include <ctime>
#include <cwctype>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <limits>
#include <map>
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
    constexpr wchar_t kApiUserSignaturesPath[] = L"/api/user/signatures";
    constexpr wchar_t kApiBinarySignaturesFullPath[] = L"/api/binary/signatures/full";
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
        case ZIVPO_RPC_STATUS_INVALID_ARGUMENT:
            return ZIVPO::Service::RpcStatusCode::InvalidArgument;
        case ZIVPO_RPC_STATUS_SCAN_FAILED:
            return ZIVPO::Service::RpcStatusCode::ScanFailed;
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
    constexpr DWORD kAvBaseRefreshIntervalMs = 300000;

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
    HANDLE g_avWorkerStopEvent = nullptr;
    HANDLE g_avWorkerThread = nullptr;

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

    enum class AvObjectType : long
    {
        Unknown = 0,
        PeFile = 1,
        PowerShellScript = 2
    };

    struct AvSignatureRecord
    {
        unsigned long long objectSignaturePrefix{ 0 };
        std::vector<unsigned char> objectSignaturePrefixBytes;
        unsigned long objectSignatureLength{ 0 };
        unsigned long remainderLength{ 0 };
        std::vector<unsigned char> objectSignature;
        long long offsetBegin{ 0 };
        long long offsetEnd{ 0 };
        AvObjectType objectType{ AvObjectType::Unknown };
        std::vector<unsigned char> avRecordSignature;
        std::wstring threatName;
    };

    struct AvBaseState
    {
        bool loaded{ false };
        std::map<unsigned long long, std::vector<AvSignatureRecord>> recordsByPrefix;
        std::map<unsigned int, std::vector<AvSignatureRecord>> shortPrefixRecordsByFirstByte;
        std::wstring releaseDate;
        unsigned long long recordsCount{ 0 };
        std::chrono::system_clock::time_point loadedAt{};
    };

    struct ScanAggregateResult
    {
        bool completed{ false };
        bool malicious{ false };
        unsigned long long scannedObjects{ 0 };
        unsigned long long infectedObjects{ 0 };
        std::wstring targetPath;
        std::wstring detectedThreat;
        std::wstring details;
    };

    std::mutex g_authStateMutex;
    ServiceAuthState g_authState{};
    std::mutex g_avBaseMutex;
    AvBaseState g_avBaseState{};

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

    void SetRpcAvBaseInfo(ZIVPO_RPC_AV_BASE_INFO* avBaseInfo, AvBaseState const& state)
    {
        if (avBaseInfo == nullptr)
        {
            return;
        }

        avBaseInfo->loaded = state.loaded ? 1 : 0;
        avBaseInfo->recordsCount = static_cast<decltype(avBaseInfo->recordsCount)>(state.recordsCount);
        avBaseInfo->releaseDate[0] = L'\0';
        if (!state.releaseDate.empty())
        {
            StringCchCopyW(avBaseInfo->releaseDate, ARRAYSIZE(avBaseInfo->releaseDate), state.releaseDate.c_str());
        }
    }

    void SetRpcScanResult(ZIVPO_RPC_SCAN_RESULT* scanResult, ScanAggregateResult const& value)
    {
        if (scanResult == nullptr)
        {
            return;
        }

        scanResult->completed = value.completed ? 1 : 0;
        scanResult->malicious = value.malicious ? 1 : 0;
        scanResult->scannedObjects = static_cast<decltype(scanResult->scannedObjects)>(value.scannedObjects);
        scanResult->infectedObjects = static_cast<decltype(scanResult->infectedObjects)>(value.infectedObjects);
        scanResult->targetPath[0] = L'\0';
        scanResult->detectedThreat[0] = L'\0';
        scanResult->details[0] = L'\0';

        if (!value.targetPath.empty())
        {
            StringCchCopyW(scanResult->targetPath, ARRAYSIZE(scanResult->targetPath), value.targetPath.c_str());
        }

        if (!value.detectedThreat.empty())
        {
            StringCchCopyW(scanResult->detectedThreat, ARRAYSIZE(scanResult->detectedThreat), value.detectedThreat.c_str());
        }

        if (!value.details.empty())
        {
            StringCchCopyW(scanResult->details, ARRAYSIZE(scanResult->details), value.details.c_str());
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
        std::wstring contentType;
        std::string body;
    };

    HttpResult SendHttpRequest(
        ApiUrlParts const& api,
        std::wstring const& method,
        std::wstring const& path,
        std::string const& requestBodyUtf8,
        std::wstring const& bearerToken,
        std::wstring const& contentTypeHeader,
        std::wstring const& acceptHeader)
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

        std::wstring headers;
        if (!contentTypeHeader.empty())
        {
            headers.append(L"Content-Type: ");
            headers.append(contentTypeHeader);
            headers.append(L"\r\n");
        }
        if (!acceptHeader.empty())
        {
            headers.append(L"Accept: ");
            headers.append(acceptHeader);
            headers.append(L"\r\n");
        }
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

        DWORD contentTypeSizeBytes = 0;
        WinHttpQueryHeaders(
            request,
            WINHTTP_QUERY_CONTENT_TYPE,
            WINHTTP_HEADER_NAME_BY_INDEX,
            WINHTTP_NO_OUTPUT_BUFFER,
            &contentTypeSizeBytes,
            WINHTTP_NO_HEADER_INDEX);
        if (GetLastError() == ERROR_INSUFFICIENT_BUFFER && contentTypeSizeBytes >= sizeof(wchar_t))
        {
            std::vector<wchar_t> contentTypeBuffer(contentTypeSizeBytes / sizeof(wchar_t), L'\0');
            if (WinHttpQueryHeaders(
                request,
                WINHTTP_QUERY_CONTENT_TYPE,
                WINHTTP_HEADER_NAME_BY_INDEX,
                contentTypeBuffer.data(),
                &contentTypeSizeBytes,
                WINHTTP_NO_HEADER_INDEX))
            {
                result.contentType.assign(contentTypeBuffer.data());
            }
        }

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

    HttpResult SendJsonRequest(
        ApiUrlParts const& api,
        std::wstring const& method,
        std::wstring const& path,
        std::string const& requestBodyUtf8,
        std::wstring const& bearerToken)
    {
        return SendHttpRequest(
            api,
            method,
            path,
            requestBodyUtf8,
            bearerToken,
            L"application/json",
            L"application/json");
    }

    HttpResult SendBinaryRequest(
        ApiUrlParts const& api,
        std::wstring const& method,
        std::wstring const& path,
        std::wstring const& bearerToken)
    {
        return SendHttpRequest(
            api,
            method,
            path,
            "",
            bearerToken,
            L"",
            L"multipart/mixed");
    }

    void ClearAvBaseStateLocked(AvBaseState& state)
    {
        state = AvBaseState{};
    }

    unsigned long long ReadPrefixAsU64(unsigned char const* bytes)
    {
        unsigned long long value = 0;
        for (int i = 0; i < 8; ++i)
        {
            value |= (static_cast<unsigned long long>(bytes[i]) << (i * 8));
        }
        return value;
    }

    std::optional<std::vector<unsigned char>> DecodeHex(std::string const& hex)
    {
        if (hex.empty() || (hex.size() % 2) != 0)
        {
            return std::nullopt;
        }

        std::vector<unsigned char> bytes;
        bytes.reserve(hex.size() / 2);
        auto hexValue = [](char ch) -> int
        {
            if (ch >= '0' && ch <= '9')
            {
                return ch - '0';
            }
            if (ch >= 'a' && ch <= 'f')
            {
                return 10 + (ch - 'a');
            }
            if (ch >= 'A' && ch <= 'F')
            {
                return 10 + (ch - 'A');
            }
            return -1;
        };

        for (size_t i = 0; i < hex.size(); i += 2)
        {
            int const hi = hexValue(hex[i]);
            int const lo = hexValue(hex[i + 1]);
            if (hi < 0 || lo < 0)
            {
                return std::nullopt;
            }
            bytes.push_back(static_cast<unsigned char>((hi << 4) | lo));
        }

        return bytes;
    }

    std::optional<std::vector<unsigned char>> DecodeBase64(std::string const& base64)
    {
        if (base64.empty())
        {
            return std::vector<unsigned char>{};
        }

        DWORD binarySize = 0;
        if (!CryptStringToBinaryA(
            base64.c_str(),
            static_cast<DWORD>(base64.size()),
            CRYPT_STRING_BASE64,
            nullptr,
            &binarySize,
            nullptr,
            nullptr))
        {
            return std::nullopt;
        }

        std::vector<unsigned char> bytes(binarySize);
        if (!CryptStringToBinaryA(
            base64.c_str(),
            static_cast<DWORD>(base64.size()),
            CRYPT_STRING_BASE64,
            bytes.data(),
            &binarySize,
            nullptr,
            nullptr))
        {
            return std::nullopt;
        }

        bytes.resize(binarySize);
        return bytes;
    }

    bool ComputeHash(
        wchar_t const* algorithmName,
        std::vector<unsigned char> const& data,
        std::vector<unsigned char>& hashOut)
    {
        hashOut.clear();
        if (algorithmName == nullptr)
        {
            return false;
        }

        BCRYPT_ALG_HANDLE algorithm = nullptr;
        NTSTATUS status = BCryptOpenAlgorithmProvider(&algorithm, algorithmName, nullptr, 0);
        if (status != 0)
        {
            return false;
        }

        DWORD hashObjectSize = 0;
        DWORD bytesWritten = 0;
        status = BCryptGetProperty(
            algorithm,
            BCRYPT_OBJECT_LENGTH,
            reinterpret_cast<PUCHAR>(&hashObjectSize),
            sizeof(hashObjectSize),
            &bytesWritten,
            0);
        if (status != 0 || hashObjectSize == 0)
        {
            BCryptCloseAlgorithmProvider(algorithm, 0);
            return false;
        }

        DWORD hashLength = 0;
        status = BCryptGetProperty(
            algorithm,
            BCRYPT_HASH_LENGTH,
            reinterpret_cast<PUCHAR>(&hashLength),
            sizeof(hashLength),
            &bytesWritten,
            0);
        if (status != 0 || hashLength == 0)
        {
            BCryptCloseAlgorithmProvider(algorithm, 0);
            return false;
        }

        std::vector<unsigned char> hashObject(hashObjectSize);
        BCRYPT_HASH_HANDLE hashHandle = nullptr;
        status = BCryptCreateHash(
            algorithm,
            &hashHandle,
            hashObject.data(),
            hashObjectSize,
            nullptr,
            0,
            0);
        if (status != 0)
        {
            BCryptCloseAlgorithmProvider(algorithm, 0);
            return false;
        }

        status = BCryptHashData(
            hashHandle,
            const_cast<PUCHAR>(data.data()),
            static_cast<ULONG>(data.size()),
            0);
        if (status != 0)
        {
            BCryptDestroyHash(hashHandle);
            BCryptCloseAlgorithmProvider(algorithm, 0);
            return false;
        }

        hashOut.resize(hashLength);
        status = BCryptFinishHash(hashHandle, hashOut.data(), hashLength, 0);
        BCryptDestroyHash(hashHandle);
        BCryptCloseAlgorithmProvider(algorithm, 0);
        return status == 0;
    }

    bool ComputeSha256(std::vector<unsigned char> const& data, std::vector<unsigned char>& hashOut)
    {
        return ComputeHash(BCRYPT_SHA256_ALGORITHM, data, hashOut);
    }

    std::vector<std::string> SplitJsonArrayObjects(std::string const& json)
    {
        std::vector<std::string> objects;
        int depth = 0;
        size_t objectStart = std::string::npos;
        bool inString = false;
        bool escaped = false;

        for (size_t i = 0; i < json.size(); ++i)
        {
            char const ch = json[i];
            if (inString)
            {
                if (escaped)
                {
                    escaped = false;
                    continue;
                }
                if (ch == '\\')
                {
                    escaped = true;
                    continue;
                }
                if (ch == '"')
                {
                    inString = false;
                }
                continue;
            }

            if (ch == '"')
            {
                inString = true;
                continue;
            }

            if (ch == '{')
            {
                if (depth == 0)
                {
                    objectStart = i;
                }
                ++depth;
                continue;
            }

            if (ch == '}')
            {
                --depth;
                if (depth == 0 && objectStart != std::string::npos)
                {
                    objects.emplace_back(json.substr(objectStart, i - objectStart + 1));
                    objectStart = std::string::npos;
                }
            }
        }

        return objects;
    }

    AvObjectType ParseObjectType(std::string const& fileType)
    {
        std::string normalized = fileType;
        std::transform(normalized.begin(), normalized.end(), normalized.begin(), [](unsigned char ch)
        {
            return static_cast<char>(std::tolower(ch));
        });

        if (normalized.find("pe") != std::string::npos ||
            normalized.find("exe") != std::string::npos ||
            normalized.find("dll") != std::string::npos)
        {
            return AvObjectType::PeFile;
        }

        if (normalized.find("powershell") != std::string::npos ||
            normalized.find("ps1") != std::string::npos)
        {
            return AvObjectType::PowerShellScript;
        }

        return AvObjectType::Unknown;
    }

    bool ParseSignatureRecord(std::string const& jsonObject, AvSignatureRecord& record, std::string& updatedAtIso)
    {
        auto firstBytesHex = ExtractJsonString(jsonObject, "firstBytesHex");
        auto objectSignatureHex = ExtractJsonString(jsonObject, "remainderHashHex");
        auto remainderLength = ExtractJsonInt64(jsonObject, "remainderLength");
        auto fileType = ExtractJsonString(jsonObject, "fileType");
        auto offsetStart = ExtractJsonInt64(jsonObject, "offsetStart");
        auto offsetEnd = ExtractJsonInt64(jsonObject, "offsetEnd");
        auto digitalSignatureBase64 = ExtractJsonString(jsonObject, "digitalSignatureBase64");
        auto status = ExtractJsonString(jsonObject, "status");
        auto threatName = ExtractJsonString(jsonObject, "threatName");
        auto updatedAt = ExtractJsonString(jsonObject, "updatedAt");

        if (!firstBytesHex.has_value() ||
            !objectSignatureHex.has_value() ||
            !remainderLength.has_value() ||
            !fileType.has_value() ||
            !offsetStart.has_value() ||
            !offsetEnd.has_value() ||
            !status.has_value())
        {
            return false;
        }

        if (*status == "DELETED")
        {
            return false;
        }

        std::optional<std::vector<unsigned char>> prefixBytes = DecodeHex(*firstBytesHex);
        std::optional<std::vector<unsigned char>> objectSignature = DecodeHex(*objectSignatureHex);
        if (!prefixBytes.has_value() || !objectSignature.has_value() || prefixBytes->empty() || *remainderLength < 0)
        {
            return false;
        }

        record.objectSignaturePrefixBytes = std::move(*prefixBytes);
        if (record.objectSignaturePrefixBytes.size() >= 8)
        {
            record.objectSignaturePrefix = ReadPrefixAsU64(record.objectSignaturePrefixBytes.data());
        }
        record.remainderLength = static_cast<unsigned long>(*remainderLength);
        record.objectSignatureLength = static_cast<unsigned long>(record.objectSignaturePrefixBytes.size() + record.remainderLength);
        record.objectSignature = std::move(*objectSignature);
        record.offsetBegin = *offsetStart;
        record.offsetEnd = *offsetEnd;
        record.objectType = ParseObjectType(*fileType);
        record.threatName = threatName.has_value() ? Utf8ToWide(*threatName) : L"UnknownThreat";
        if (updatedAt.has_value())
        {
            updatedAtIso = *updatedAt;
        }

        if (digitalSignatureBase64.has_value())
        {
            auto signatureBytes = DecodeBase64(*digitalSignatureBase64);
            if (signatureBytes.has_value())
            {
                record.avRecordSignature = std::move(*signatureBytes);
            }
        }

        return true;
    }

    std::string ToLowerAscii(std::string value)
    {
        std::transform(value.begin(), value.end(), value.begin(), [](unsigned char ch)
        {
            return static_cast<char>(std::tolower(ch));
        });
        return value;
    }

    std::string TrimAscii(std::string value)
    {
        auto isSpace = [](unsigned char ch)
        {
            return std::isspace(ch) != 0;
        };

        while (!value.empty() && isSpace(static_cast<unsigned char>(value.front())))
        {
            value.erase(value.begin());
        }
        while (!value.empty() && isSpace(static_cast<unsigned char>(value.back())))
        {
            value.pop_back();
        }
        return value;
    }

    std::optional<std::string> ExtractMultipartBoundary(std::wstring const& contentType)
    {
        if (contentType.empty())
        {
            return std::nullopt;
        }

        std::string raw = WideToUtf8(contentType);
        std::string lower = ToLowerAscii(raw);
        size_t boundaryPos = lower.find("boundary=");
        if (boundaryPos == std::string::npos)
        {
            return std::nullopt;
        }

        boundaryPos += 9;
        if (boundaryPos >= raw.size())
        {
            return std::nullopt;
        }

        if (raw[boundaryPos] == '"' || raw[boundaryPos] == '\'')
        {
            char const quote = raw[boundaryPos];
            size_t const endQuote = raw.find(quote, boundaryPos + 1);
            if (endQuote == std::string::npos || endQuote <= boundaryPos + 1)
            {
                return std::nullopt;
            }
            return raw.substr(boundaryPos + 1, endQuote - (boundaryPos + 1));
        }

        size_t boundaryEnd = raw.find(';', boundaryPos);
        if (boundaryEnd == std::string::npos)
        {
            boundaryEnd = raw.size();
        }
        std::string boundary = TrimAscii(raw.substr(boundaryPos, boundaryEnd - boundaryPos));
        if (boundary.empty())
        {
            return std::nullopt;
        }
        return boundary;
    }

    struct MultipartParts
    {
        std::vector<unsigned char> manifest;
        std::vector<unsigned char> data;
    };

    bool ParseMultipartMixedParts(std::string const& payload, std::string const& boundary, MultipartParts& parts)
    {
        parts = {};
        if (payload.empty() || boundary.empty())
        {
            return false;
        }

        std::string const marker = "--" + boundary;
        size_t markerPos = payload.find(marker);
        while (markerPos != std::string::npos)
        {
            size_t cursor = markerPos + marker.size();
            if (cursor + 2 <= payload.size() && payload.compare(cursor, 2, "--") == 0)
            {
                break;
            }

            if (cursor + 2 <= payload.size() && payload.compare(cursor, 2, "\r\n") == 0)
            {
                cursor += 2;
            }
            else if (cursor < payload.size() && payload[cursor] == '\n')
            {
                cursor += 1;
            }

            size_t const headersEnd = payload.find("\r\n\r\n", cursor);
            if (headersEnd == std::string::npos)
            {
                return false;
            }

            std::string headersBlock = payload.substr(cursor, headersEnd - cursor);
            cursor = headersEnd + 4;

            std::optional<size_t> contentLength;
            std::string partName;
            std::string fileName;

            size_t lineStart = 0;
            while (lineStart < headersBlock.size())
            {
                size_t lineEnd = headersBlock.find("\r\n", lineStart);
                if (lineEnd == std::string::npos)
                {
                    lineEnd = headersBlock.size();
                }

                std::string line = headersBlock.substr(lineStart, lineEnd - lineStart);
                std::string lowerLine = ToLowerAscii(line);

                if (lowerLine.rfind("content-length:", 0) == 0)
                {
                    std::string value = TrimAscii(line.substr(15));
                    try
                    {
                        contentLength = static_cast<size_t>(std::stoull(value));
                    }
                    catch (...)
                    {
                        contentLength = std::nullopt;
                    }
                }
                else if (lowerLine.rfind("content-disposition:", 0) == 0)
                {
                    auto extractToken = [&](std::string const& tokenName) -> std::string
                    {
                        std::string lowerTokenName = ToLowerAscii(tokenName);
                        size_t tokenPos = lowerLine.find(lowerTokenName + "=");
                        if (tokenPos == std::string::npos)
                        {
                            return {};
                        }

                        size_t valuePos = tokenPos + lowerTokenName.size() + 1;
                        if (valuePos >= line.size())
                        {
                            return {};
                        }

                        if (line[valuePos] == '"' || line[valuePos] == '\'')
                        {
                            char const quote = line[valuePos];
                            size_t valueEnd = line.find(quote, valuePos + 1);
                            if (valueEnd == std::string::npos)
                            {
                                return {};
                            }
                            return line.substr(valuePos + 1, valueEnd - (valuePos + 1));
                        }

                        size_t valueEnd = line.find(';', valuePos);
                        if (valueEnd == std::string::npos)
                        {
                            valueEnd = line.size();
                        }
                        return TrimAscii(line.substr(valuePos, valueEnd - valuePos));
                    };

                    partName = extractToken("name");
                    fileName = extractToken("filename");
                }

                if (lineEnd == headersBlock.size())
                {
                    break;
                }
                lineStart = lineEnd + 2;
            }

            size_t dataEnd = std::string::npos;
            if (contentLength.has_value())
            {
                if (cursor + *contentLength > payload.size())
                {
                    return false;
                }
                dataEnd = cursor + *contentLength;
            }
            else
            {
                std::string const markerWithCrlf = "\r\n" + marker;
                size_t markerWithCrlfPos = payload.find(markerWithCrlf, cursor);
                if (markerWithCrlfPos == std::string::npos)
                {
                    return false;
                }
                dataEnd = markerWithCrlfPos;
            }

            std::vector<unsigned char> dataPart(
                payload.begin() + static_cast<std::ptrdiff_t>(cursor),
                payload.begin() + static_cast<std::ptrdiff_t>(dataEnd));

            std::string lowerName = ToLowerAscii(partName);
            std::string lowerFile = ToLowerAscii(fileName);
            if (lowerName == "manifest" || lowerFile == "manifest.bin")
            {
                parts.manifest = std::move(dataPart);
            }
            else if (lowerName == "data" || lowerFile == "data.bin")
            {
                parts.data = std::move(dataPart);
            }

            cursor = dataEnd;
            if (cursor + 2 <= payload.size() && payload.compare(cursor, 2, "\r\n") == 0)
            {
                cursor += 2;
            }
            markerPos = payload.find(marker, cursor);
        }

        return !parts.data.empty();
    }

    class BinaryReader
    {
    public:
        explicit BinaryReader(std::vector<unsigned char> const& bytes) : m_bytes(bytes) {}

        bool ReadU8(unsigned char& value)
        {
            if (m_offset + 1 > m_bytes.size())
            {
                return false;
            }
            value = m_bytes[m_offset++];
            return true;
        }

        bool ReadU16(unsigned short& value)
        {
            if (m_offset + 2 > m_bytes.size())
            {
                return false;
            }
            value = static_cast<unsigned short>(
                (static_cast<unsigned short>(m_bytes[m_offset]) << 8) |
                static_cast<unsigned short>(m_bytes[m_offset + 1]));
            m_offset += 2;
            return true;
        }

        bool ReadU32(unsigned long& value)
        {
            if (m_offset + 4 > m_bytes.size())
            {
                return false;
            }
            value = (static_cast<unsigned long>(m_bytes[m_offset]) << 24) |
                (static_cast<unsigned long>(m_bytes[m_offset + 1]) << 16) |
                (static_cast<unsigned long>(m_bytes[m_offset + 2]) << 8) |
                static_cast<unsigned long>(m_bytes[m_offset + 3]);
            m_offset += 4;
            return true;
        }

        bool ReadI64(long long& value)
        {
            if (m_offset + 8 > m_bytes.size())
            {
                return false;
            }

            unsigned long long raw = 0;
            for (int i = 0; i < 8; ++i)
            {
                raw = (raw << 8) | static_cast<unsigned long long>(m_bytes[m_offset + i]);
            }
            m_offset += 8;
            value = static_cast<long long>(raw);
            return true;
        }

        bool ReadBytes(size_t size, std::vector<unsigned char>& out)
        {
            if (m_offset + size > m_bytes.size())
            {
                return false;
            }

            out.assign(
                m_bytes.begin() + static_cast<std::ptrdiff_t>(m_offset),
                m_bytes.begin() + static_cast<std::ptrdiff_t>(m_offset + size));
            m_offset += size;
            return true;
        }

        bool Skip(size_t size)
        {
            if (m_offset + size > m_bytes.size())
            {
                return false;
            }
            m_offset += size;
            return true;
        }

    private:
        std::vector<unsigned char> const& m_bytes;
        size_t m_offset{ 0 };
    };

    std::wstring FormatEpochMillisAsUtc(long long epochMillis)
    {
        std::chrono::milliseconds millis(epochMillis);
        std::chrono::system_clock::time_point timePoint(millis);
        std::time_t time = std::chrono::system_clock::to_time_t(timePoint);
        std::tm utc{};
        if (gmtime_s(&utc, &time) != 0)
        {
            return L"n/a";
        }

        wchar_t buffer[64]{};
        StringCchPrintfW(
            buffer,
            ARRAYSIZE(buffer),
            L"%04d-%02d-%02dT%02d:%02d:%02dZ",
            utc.tm_year + 1900,
            utc.tm_mon + 1,
            utc.tm_mday,
            utc.tm_hour,
            utc.tm_min,
            utc.tm_sec);
        return buffer;
    }

    bool ParseBinaryManifestReleaseDate(
        std::vector<unsigned char> const& manifestBytes,
        std::wstring& releaseDate)
    {
        releaseDate = L"n/a";
        if (manifestBytes.size() < 31)
        {
            return false;
        }

        BinaryReader reader(manifestBytes);
        std::vector<unsigned char> magic;
        if (!reader.ReadBytes(12, magic))
        {
            return false;
        }

        std::string magicString(magic.begin(), magic.end());
        if (magicString != "MF-Churakov-")
        {
            return false;
        }

        unsigned short version = 0;
        unsigned char exportType = 0;
        long long generatedAtEpochMillis = 0;
        if (!reader.ReadU16(version) || !reader.ReadU8(exportType) || !reader.ReadI64(generatedAtEpochMillis))
        {
            return false;
        }

        releaseDate = FormatEpochMillisAsUtc(generatedAtEpochMillis);
        return true;
    }

    bool ParseBinaryDataPart(
        std::vector<unsigned char> const& dataBytes,
        AvBaseState& parsedState)
    {
        parsedState = {};
        BinaryReader reader(dataBytes);

        std::vector<unsigned char> magic;
        if (!reader.ReadBytes(12, magic))
        {
            return false;
        }
        std::string magicString(magic.begin(), magic.end());
        if (magicString != "DB-Churakov-")
        {
            return false;
        }

        unsigned short version = 0;
        unsigned long recordsCount = 0;
        if (!reader.ReadU16(version) || !reader.ReadU32(recordsCount))
        {
            return false;
        }
        (void)version;

        for (unsigned long i = 0; i < recordsCount; ++i)
        {
            unsigned long threatNameLength = 0;
            if (!reader.ReadU32(threatNameLength))
            {
                return false;
            }
            std::vector<unsigned char> threatNameBytes;
            if (!reader.ReadBytes(static_cast<size_t>(threatNameLength), threatNameBytes))
            {
                return false;
            }

            unsigned long prefixLength = 0;
            if (!reader.ReadU32(prefixLength) || prefixLength == 0)
            {
                return false;
            }
            std::vector<unsigned char> prefixBytes;
            if (!reader.ReadBytes(static_cast<size_t>(prefixLength), prefixBytes))
            {
                return false;
            }

            unsigned long hashLength = 0;
            if (!reader.ReadU32(hashLength) || hashLength == 0)
            {
                return false;
            }
            std::vector<unsigned char> remainderHash;
            if (!reader.ReadBytes(static_cast<size_t>(hashLength), remainderHash))
            {
                return false;
            }

            long long remainderLength = 0;
            if (!reader.ReadI64(remainderLength) || remainderLength < 0)
            {
                return false;
            }

            unsigned long fileTypeLength = 0;
            if (!reader.ReadU32(fileTypeLength))
            {
                return false;
            }
            std::vector<unsigned char> fileTypeBytes;
            if (!reader.ReadBytes(static_cast<size_t>(fileTypeLength), fileTypeBytes))
            {
                return false;
            }

            long long offsetStart = 0;
            long long offsetEnd = 0;
            if (!reader.ReadI64(offsetStart) || !reader.ReadI64(offsetEnd))
            {
                return false;
            }

            if (offsetEnd < offsetStart)
            {
                continue;
            }

            unsigned long long maxLen = (std::numeric_limits<unsigned long>::max)();
            unsigned long long fullSignatureLength =
                static_cast<unsigned long long>(prefixBytes.size()) + static_cast<unsigned long long>(remainderLength);
            if (fullSignatureLength > maxLen)
            {
                continue;
            }

            AvSignatureRecord record{};
            record.objectSignaturePrefixBytes = std::move(prefixBytes);
            if (record.objectSignaturePrefixBytes.size() >= 8)
            {
                record.objectSignaturePrefix = ReadPrefixAsU64(record.objectSignaturePrefixBytes.data());
            }
            record.remainderLength = static_cast<unsigned long>(remainderLength);
            record.objectSignatureLength = static_cast<unsigned long>(fullSignatureLength);
            record.objectSignature = std::move(remainderHash);
            record.offsetBegin = offsetStart;
            record.offsetEnd = offsetEnd;
            record.objectType = ParseObjectType(std::string(fileTypeBytes.begin(), fileTypeBytes.end()));
            record.threatName = Utf8ToWide(std::string(threatNameBytes.begin(), threatNameBytes.end()));
            if (record.threatName.empty())
            {
                record.threatName = L"UnknownThreat";
            }

            if (record.objectSignaturePrefixBytes.size() >= 8)
            {
                parsedState.recordsByPrefix[record.objectSignaturePrefix].push_back(std::move(record));
            }
            else
            {
                parsedState.shortPrefixRecordsByFirstByte[record.objectSignaturePrefixBytes.front()].push_back(std::move(record));
            }
            ++parsedState.recordsCount;
        }

        parsedState.loaded = true;
        parsedState.loadedAt = std::chrono::system_clock::now();
        if (parsedState.recordsCount == 0)
        {
            return false;
        }

        return true;
    }

    long LoadAvBaseFromServer(ApiUrlParts const& api, std::wstring const& accessToken, AvBaseState& avBaseState)
    {
        HttpResult response = SendBinaryRequest(api, L"GET", kApiBinarySignaturesFullPath, accessToken);
        if (!response.ok)
        {
            return ZIVPO_RPC_STATUS_NETWORK_ERROR;
        }

        if (response.statusCode == 401 || response.statusCode == 403)
        {
            return ZIVPO_RPC_STATUS_NOT_AUTHENTICATED;
        }

        if (response.statusCode == 404)
        {
            return ZIVPO_RPC_STATUS_NO_LICENSE;
        }

        if (response.statusCode != 200)
        {
            return ZIVPO_RPC_STATUS_INTERNAL_ERROR;
        }

        std::optional<std::string> boundary = ExtractMultipartBoundary(response.contentType);
        if (!boundary.has_value())
        {
            return ZIVPO_RPC_STATUS_INTERNAL_ERROR;
        }

        MultipartParts parts{};
        if (!ParseMultipartMixedParts(response.body, *boundary, parts))
        {
            return ZIVPO_RPC_STATUS_INTERNAL_ERROR;
        }

        AvBaseState parsedState{};
        if (!ParseBinaryDataPart(parts.data, parsedState))
        {
            return ZIVPO_RPC_STATUS_INTERNAL_ERROR;
        }

        std::wstring releaseDate = L"n/a";
        if (!parts.manifest.empty())
        {
            ParseBinaryManifestReleaseDate(parts.manifest, releaseDate);
        }
        parsedState.releaseDate = std::move(releaseDate);

        avBaseState = std::move(parsedState);
        return ZIVPO_RPC_STATUS_OK;
    }

    bool ReadFileBytes(std::wstring const& filePath, std::vector<unsigned char>& bytes)
    {
        bytes.clear();
        std::ifstream file(std::filesystem::path(filePath), std::ios::binary);
        if (!file.is_open())
        {
            return false;
        }

        file.seekg(0, std::ios::end);
        std::streamoff size = file.tellg();
        if (size < 0)
        {
            return false;
        }
        file.seekg(0, std::ios::beg);

        bytes.resize(static_cast<size_t>(size));
        if (!bytes.empty())
        {
            file.read(reinterpret_cast<char*>(bytes.data()), size);
            if (!file.good() && !file.eof())
            {
                return false;
            }
        }

        return true;
    }

    AvObjectType DetectObjectType(std::wstring const& path, std::vector<unsigned char> const& bytes)
    {
        std::wstring lowerPath = path;
        std::transform(lowerPath.begin(), lowerPath.end(), lowerPath.begin(), [](wchar_t ch)
        {
            return static_cast<wchar_t>(std::towlower(ch));
        });

        if (bytes.size() >= 2 && bytes[0] == 'M' && bytes[1] == 'Z')
        {
            return AvObjectType::PeFile;
        }

        if (lowerPath.size() >= 4 && lowerPath.substr(lowerPath.size() - 4) == L".ps1")
        {
            return AvObjectType::PowerShellScript;
        }

        return AvObjectType::Unknown;
    }

    bool ScanBytesWithBase(
        std::vector<unsigned char> const& bytes,
        AvObjectType objectType,
        AvBaseState const& baseState,
        std::wstring& detectedThreat)
    {
        detectedThreat.clear();
        if (bytes.empty())
        {
            return false;
        }

        auto tryMatchRecordAtPosition = [&](AvSignatureRecord const& record, size_t position) -> bool
        {
            if (record.objectType != AvObjectType::Unknown && record.objectType != objectType)
            {
                return false;
            }

            long long const currentOffset = static_cast<long long>(position);
            if (currentOffset < record.offsetBegin || currentOffset > record.offsetEnd)
            {
                return false;
            }

            size_t const prefixLength = record.objectSignaturePrefixBytes.size();
            if (prefixLength == 0 || position + prefixLength > bytes.size())
            {
                return false;
            }

            if (!std::equal(
                record.objectSignaturePrefixBytes.begin(),
                record.objectSignaturePrefixBytes.end(),
                bytes.begin() + static_cast<std::ptrdiff_t>(position)))
            {
                return false;
            }

            size_t remainderLength = static_cast<size_t>(record.remainderLength);
            if (position + prefixLength + remainderLength > bytes.size())
            {
                return false;
            }

            std::vector<unsigned char> remainderBytes;
            remainderBytes.assign(
                bytes.begin() + static_cast<std::ptrdiff_t>(position + prefixLength),
                bytes.begin() + static_cast<std::ptrdiff_t>(position + prefixLength + remainderLength));

            std::vector<unsigned char> fullSignatureBytes;
            fullSignatureBytes.assign(
                bytes.begin() + static_cast<std::ptrdiff_t>(position),
                bytes.begin() + static_cast<std::ptrdiff_t>(position + prefixLength + remainderLength));

            auto matchDetected = [&]()
            {
                detectedThreat = record.threatName.empty() ? L"Malware signature" : record.threatName;
                return true;
            };

            if (record.objectSignature == remainderBytes || record.objectSignature == fullSignatureBytes)
            {
                return matchDetected();
            }

            std::vector<unsigned char> computedHash;
            auto hashAndCompare = [&](wchar_t const* algorithm, std::vector<unsigned char> const& input) -> bool
            {
                if (!ComputeHash(algorithm, input, computedHash))
                {
                    return false;
                }

                if (computedHash == record.objectSignature)
                {
                    return true;
                }

                if (!record.objectSignature.empty() &&
                    record.objectSignature.size() < computedHash.size() &&
                    std::equal(
                        record.objectSignature.begin(),
                        record.objectSignature.end(),
                        computedHash.begin()))
                {
                    return true;
                }

                return false;
            };

            auto hashMatchesByLength = [&](std::vector<unsigned char> const& input) -> bool
            {
                size_t const signatureLen = record.objectSignature.size();
                if (signatureLen == 16)
                {
                    return hashAndCompare(BCRYPT_MD5_ALGORITHM, input);
                }
                if (signatureLen == 20)
                {
                    return hashAndCompare(BCRYPT_SHA1_ALGORITHM, input);
                }

                // Default and fallback path.
                return hashAndCompare(BCRYPT_SHA256_ALGORITHM, input);
            };

            // Primary backend path: remainder hash.
            if (hashMatchesByLength(remainderBytes))
            {
                return matchDetected();
            }

            // Compatibility path: hash over full signature bytes.
            if (hashMatchesByLength(fullSignatureBytes))
            {
                return matchDetected();
            }

            return false;
        };

        for (size_t position = 0; position + 8 <= bytes.size(); ++position)
        {
            unsigned long long prefix = ReadPrefixAsU64(bytes.data() + position);
            auto prefixIt = baseState.recordsByPrefix.find(prefix);
            if (prefixIt != baseState.recordsByPrefix.end())
            {
                auto const& records = prefixIt->second;
                for (auto const& record : records)
                {
                    if (tryMatchRecordAtPosition(record, position))
                    {
                        return true;
                    }
                }
            }

            unsigned int const firstByte = bytes[position];
            auto shortPrefixIt = baseState.shortPrefixRecordsByFirstByte.find(firstByte);
            if (shortPrefixIt == baseState.shortPrefixRecordsByFirstByte.end())
            {
                continue;
            }

            for (auto const& record : shortPrefixIt->second)
            {
                if (tryMatchRecordAtPosition(record, position))
                {
                    return true;
                }
            }
        }

        // Also scan tail where less than 8 bytes remain (needed for short prefixes).
        if (!baseState.shortPrefixRecordsByFirstByte.empty())
        {
            size_t const startTail = bytes.size() >= 8 ? bytes.size() - 7 : 0;
            for (size_t position = startTail; position < bytes.size(); ++position)
            {
                unsigned int const firstByte = bytes[position];
                auto shortPrefixIt = baseState.shortPrefixRecordsByFirstByte.find(firstByte);
                if (shortPrefixIt == baseState.shortPrefixRecordsByFirstByte.end())
                {
                    continue;
                }

                for (auto const& record : shortPrefixIt->second)
                {
                    if (tryMatchRecordAtPosition(record, position))
                    {
                        return true;
                    }
                }
            }
        }

        return false;
    }

    long ScanSingleFileWithState(
        std::wstring const& filePath,
        AvBaseState const& baseState,
        ScanAggregateResult& aggregate)
    {
        if (filePath.empty())
        {
            aggregate.details = L"File path is empty.";
            return ZIVPO_RPC_STATUS_INVALID_ARGUMENT;
        }

        std::error_code fsError;
        if (!std::filesystem::exists(std::filesystem::path(filePath), fsError) ||
            !std::filesystem::is_regular_file(std::filesystem::path(filePath), fsError))
        {
            aggregate.details = L"File was not found.";
            return ZIVPO_RPC_STATUS_INVALID_ARGUMENT;
        }

        std::vector<unsigned char> bytes;
        if (!ReadFileBytes(filePath, bytes))
        {
            aggregate.details = L"Failed to read file.";
            return ZIVPO_RPC_STATUS_SCAN_FAILED;
        }

        AvObjectType objectType = DetectObjectType(filePath, bytes);
        std::wstring threat;
        bool malicious = ScanBytesWithBase(bytes, objectType, baseState, threat);

        aggregate.completed = true;
        aggregate.malicious = malicious;
        aggregate.scannedObjects = 1;
        aggregate.infectedObjects = malicious ? 1 : 0;
        aggregate.targetPath = filePath;
        aggregate.detectedThreat = malicious ? threat : L"";
        aggregate.details = malicious ? L"Malicious signature detected." : L"No threats detected.";
        return ZIVPO_RPC_STATUS_OK;
    }

    long ScanDirectoryWithState(
        std::wstring const& directoryPath,
        AvBaseState const& baseState,
        ScanAggregateResult& aggregate)
    {
        if (directoryPath.empty())
        {
            aggregate.details = L"Directory path is empty.";
            return ZIVPO_RPC_STATUS_INVALID_ARGUMENT;
        }

        std::error_code fsError;
        std::filesystem::path root(directoryPath);
        if (!std::filesystem::exists(root, fsError) || !std::filesystem::is_directory(root, fsError))
        {
            aggregate.details = L"Directory was not found.";
            return ZIVPO_RPC_STATUS_INVALID_ARGUMENT;
        }

        aggregate.completed = true;
        aggregate.targetPath = directoryPath;
        aggregate.details = L"No threats detected.";

        for (auto const& entry : std::filesystem::recursive_directory_iterator(
            root,
            std::filesystem::directory_options::skip_permission_denied,
            fsError))
        {
            if (fsError)
            {
                fsError.clear();
                continue;
            }

            if (!entry.is_regular_file(fsError))
            {
                fsError.clear();
                continue;
            }

            std::wstring filePath = entry.path().wstring();
            std::vector<unsigned char> bytes;
            if (!ReadFileBytes(filePath, bytes))
            {
                continue;
            }

            ++aggregate.scannedObjects;
            AvObjectType objectType = DetectObjectType(filePath, bytes);
            std::wstring threat;
            if (ScanBytesWithBase(bytes, objectType, baseState, threat))
            {
                aggregate.malicious = true;
                ++aggregate.infectedObjects;
                if (aggregate.detectedThreat.empty())
                {
                    aggregate.detectedThreat = threat;
                }
            }
        }

        if (aggregate.malicious)
        {
            aggregate.details = L"Threats detected in directory.";
        }

        return ZIVPO_RPC_STATUS_OK;
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

    void StopAvWorkerThread();

    void StartAvWorkerThread()
    {
        if (g_avWorkerStopEvent == nullptr || g_avWorkerThread != nullptr)
        {
            return;
        }
        ResetEvent(g_avWorkerStopEvent);

        auto workerProc = [](LPVOID) -> DWORD
        {
            while (!g_serviceStopping.load())
            {
                ServiceAuthState authSnapshot{};
                {
                    std::lock_guard lock(g_authStateMutex);
                    authSnapshot = g_authState;
                }

                if (!authSnapshot.authenticated || !authSnapshot.hasLicenseTicket || authSnapshot.accessToken.empty())
                {
                    std::lock_guard lock(g_avBaseMutex);
                    ClearAvBaseStateLocked(g_avBaseState);
                    break;
                }

                ApiUrlParts api{};
                if (ParseApiUrlParts(GetApiBaseUrl(), api))
                {
                    AvBaseState refreshed{};
                    long refreshStatus = LoadAvBaseFromServer(api, authSnapshot.accessToken, refreshed);
                    if (refreshStatus == ZIVPO_RPC_STATUS_OK)
                    {
                        std::lock_guard lock(g_avBaseMutex);
                        g_avBaseState = std::move(refreshed);
                    }
                    else if (refreshStatus == ZIVPO_RPC_STATUS_NO_LICENSE || refreshStatus == ZIVPO_RPC_STATUS_NOT_AUTHENTICATED)
                    {
                        std::lock_guard authLock(g_authStateMutex);
                        g_authState.hasLicenseTicket = false;
                        g_authState.licenseBlocked = false;
                        g_authState.licenseExpirationDate.clear();
                        g_authState.licenseRefreshAt = {};
                        std::lock_guard avLock(g_avBaseMutex);
                        ClearAvBaseStateLocked(g_avBaseState);
                        break;
                    }
                }

                HANDLE waitHandles[2] = { g_avWorkerStopEvent, g_sessionMonitorStopEvent };
                DWORD waitResult = WaitForMultipleObjects(2, waitHandles, FALSE, kAvBaseRefreshIntervalMs);
                if (waitResult == WAIT_OBJECT_0 || waitResult == WAIT_OBJECT_0 + 1)
                {
                    break;
                }
            }

            return 0;
        };

        g_avWorkerThread = CreateThread(nullptr, 0, workerProc, nullptr, 0, nullptr);
    }

    void StopAvWorkerThread()
    {
        if (g_avWorkerStopEvent != nullptr)
        {
            SetEvent(g_avWorkerStopEvent);
        }

        if (g_avWorkerThread != nullptr)
        {
            WaitForSingleObject(g_avWorkerThread, 10000);
            CloseHandle(g_avWorkerThread);
            g_avWorkerThread = nullptr;
        }
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
                            std::lock_guard avLock(g_avBaseMutex);
                            ClearAvBaseStateLocked(g_avBaseState);
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
                                std::lock_guard avLock(g_avBaseMutex);
                                ClearAvBaseStateLocked(g_avBaseState);
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

                bool shouldRunAvWorker = false;
                {
                    std::lock_guard lock(g_authStateMutex);
                    shouldRunAvWorker = g_authState.authenticated && g_authState.hasLicenseTicket;
                }
                if (shouldRunAvWorker)
                {
                    StartAvWorkerThread();
                }
                else
                {
                    StopAvWorkerThread();
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
        g_avWorkerStopEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
        g_avWorkerThread = nullptr;
        {
            std::lock_guard lock(g_authStateMutex);
            ClearAuthStateLocked(g_authState);
        }
        {
            std::lock_guard lock(g_avBaseMutex);
            ClearAvBaseStateLocked(g_avBaseState);
        }

        UpdateServiceStatus(SERVICE_START_PENDING, NO_ERROR, 5000);

        RPC_STATUS rpcStatus = StartRpcServer();
        if (rpcStatus != RPC_S_OK)
        {
            g_serviceStopping = true;
            StopAvWorkerThread();
            if (g_avWorkerStopEvent != nullptr)
            {
                CloseHandle(g_avWorkerStopEvent);
                g_avWorkerStopEvent = nullptr;
            }
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
        StopAvWorkerThread();
        if (g_avWorkerStopEvent != nullptr)
        {
            CloseHandle(g_avWorkerStopEvent);
            g_avWorkerStopEvent = nullptr;
        }
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

    if (licenseStatus == ZIVPO_RPC_STATUS_OK)
    {
        std::wstring accessTokenSnapshot;
        {
            std::lock_guard lock(g_authStateMutex);
            accessTokenSnapshot = g_authState.accessToken;
        }

        AvBaseState loadedBase{};
        long loadStatus = LoadAvBaseFromServer(api, accessTokenSnapshot, loadedBase);
        if (loadStatus == ZIVPO_RPC_STATUS_OK)
        {
            std::lock_guard lock(g_avBaseMutex);
            g_avBaseState = std::move(loadedBase);
        }
        StartAvWorkerThread();
    }
    else
    {
        {
            std::lock_guard lock(g_avBaseMutex);
            ClearAvBaseStateLocked(g_avBaseState);
        }
        StopAvWorkerThread();
    }

    EnsureAuthWorkerStarted();
    return ZIVPO_RPC_STATUS_OK;
}

extern "C" long RpcLogout([[maybe_unused]] handle_t hBinding)
{
    {
        std::lock_guard lock(g_authStateMutex);
        ClearAuthStateLocked(g_authState);
    }
    {
        std::lock_guard avLock(g_avBaseMutex);
        ClearAvBaseStateLocked(g_avBaseState);
    }
    StopAvWorkerThread();
    return ZIVPO_RPC_STATUS_OK;
}

extern "C" long RpcGetLicenseInfo([[maybe_unused]] handle_t hBinding, ZIVPO_RPC_LICENSE_INFO* licenseInfo)
{
    bool authenticated = false;
    bool hasLicense = false;
    {
        std::lock_guard lock(g_authStateMutex);
        SetRpcLicenseInfo(licenseInfo, g_authState);
        authenticated = g_authState.authenticated;
        hasLicense = g_authState.hasLicenseTicket;
    }

    if (!authenticated)
    {
        return ZIVPO_RPC_STATUS_NOT_AUTHENTICATED;
    }

    if (!hasLicense)
    {
        {
            std::lock_guard avLock(g_avBaseMutex);
            ClearAvBaseStateLocked(g_avBaseState);
        }
        StopAvWorkerThread();
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

    AvBaseState loadedBase{};
    long loadStatus = LoadAvBaseFromServer(api, snapshot.accessToken, loadedBase);
    if (loadStatus == ZIVPO_RPC_STATUS_OK)
    {
        std::lock_guard lock(g_avBaseMutex);
        g_avBaseState = std::move(loadedBase);
    }
    StartAvWorkerThread();

    return ZIVPO_RPC_STATUS_OK;
}

extern "C" long RpcGetAvBaseInfo([[maybe_unused]] handle_t hBinding, ZIVPO_RPC_AV_BASE_INFO* avBaseInfo)
{
    if (avBaseInfo == nullptr)
    {
        return ZIVPO_RPC_STATUS_INVALID_ARGUMENT;
    }

    ServiceAuthState authSnapshot{};
    {
        std::lock_guard lock(g_authStateMutex);
        authSnapshot = g_authState;
    }

    if (!authSnapshot.authenticated)
    {
        return ZIVPO_RPC_STATUS_NOT_AUTHENTICATED;
    }

    if (!authSnapshot.hasLicenseTicket)
    {
        return ZIVPO_RPC_STATUS_NO_LICENSE;
    }

    {
        std::lock_guard lock(g_avBaseMutex);
        if (!g_avBaseState.loaded)
        {
            ApiUrlParts api{};
            if (!ParseApiUrlParts(GetApiBaseUrl(), api))
            {
                return ZIVPO_RPC_STATUS_NETWORK_ERROR;
            }

            AvBaseState loadedBase{};
            long loadStatus = LoadAvBaseFromServer(api, authSnapshot.accessToken, loadedBase);
            if (loadStatus != ZIVPO_RPC_STATUS_OK)
            {
                return loadStatus;
            }

            g_avBaseState = std::move(loadedBase);
            StartAvWorkerThread();
        }

        SetRpcAvBaseInfo(avBaseInfo, g_avBaseState);
    }

    return ZIVPO_RPC_STATUS_OK;
}

extern "C" long RpcScanFile(
    [[maybe_unused]] handle_t hBinding,
    const wchar_t* filePath,
    ZIVPO_RPC_SCAN_RESULT* scanResult)
{
    if (scanResult == nullptr || filePath == nullptr)
    {
        return ZIVPO_RPC_STATUS_INVALID_ARGUMENT;
    }

    ServiceAuthState authSnapshot{};
    {
        std::lock_guard lock(g_authStateMutex);
        authSnapshot = g_authState;
    }

    if (!authSnapshot.authenticated)
    {
        return ZIVPO_RPC_STATUS_NOT_AUTHENTICATED;
    }

    if (!authSnapshot.hasLicenseTicket)
    {
        return ZIVPO_RPC_STATUS_NO_LICENSE;
    }

    AvBaseState baseSnapshot{};
    {
        std::lock_guard lock(g_avBaseMutex);
        if (!g_avBaseState.loaded)
        {
            ApiUrlParts api{};
            if (!ParseApiUrlParts(GetApiBaseUrl(), api))
            {
                return ZIVPO_RPC_STATUS_NETWORK_ERROR;
            }

            AvBaseState loadedBase{};
            long loadStatus = LoadAvBaseFromServer(api, authSnapshot.accessToken, loadedBase);
            if (loadStatus != ZIVPO_RPC_STATUS_OK)
            {
                return loadStatus;
            }

            g_avBaseState = std::move(loadedBase);
            StartAvWorkerThread();
        }
        baseSnapshot = g_avBaseState;
    }

    ScanAggregateResult aggregate{};
    long scanStatus = ScanSingleFileWithState(filePath, baseSnapshot, aggregate);
    SetRpcScanResult(scanResult, aggregate);
    return scanStatus;
}

extern "C" long RpcScanDirectory(
    [[maybe_unused]] handle_t hBinding,
    const wchar_t* directoryPath,
    ZIVPO_RPC_SCAN_RESULT* scanResult)
{
    if (scanResult == nullptr || directoryPath == nullptr)
    {
        return ZIVPO_RPC_STATUS_INVALID_ARGUMENT;
    }

    ServiceAuthState authSnapshot{};
    {
        std::lock_guard lock(g_authStateMutex);
        authSnapshot = g_authState;
    }

    if (!authSnapshot.authenticated)
    {
        return ZIVPO_RPC_STATUS_NOT_AUTHENTICATED;
    }

    if (!authSnapshot.hasLicenseTicket)
    {
        return ZIVPO_RPC_STATUS_NO_LICENSE;
    }

    AvBaseState baseSnapshot{};
    {
        std::lock_guard lock(g_avBaseMutex);
        if (!g_avBaseState.loaded)
        {
            ApiUrlParts api{};
            if (!ParseApiUrlParts(GetApiBaseUrl(), api))
            {
                return ZIVPO_RPC_STATUS_NETWORK_ERROR;
            }

            AvBaseState loadedBase{};
            long loadStatus = LoadAvBaseFromServer(api, authSnapshot.accessToken, loadedBase);
            if (loadStatus != ZIVPO_RPC_STATUS_OK)
            {
                return loadStatus;
            }

            g_avBaseState = std::move(loadedBase);
            StartAvWorkerThread();
        }
        baseSnapshot = g_avBaseState;
    }

    ScanAggregateResult aggregate{};
    long scanStatus = ScanDirectoryWithState(directoryPath, baseSnapshot, aggregate);
    SetRpcScanResult(scanResult, aggregate);
    return scanStatus;
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

        void FillAvBaseInfo(ZIVPO_RPC_AV_BASE_INFO const& rpcAvBase, AvBaseInfo& avBaseInfo)
        {
            avBaseInfo.loaded = rpcAvBase.loaded != 0;
            avBaseInfo.recordsCount = static_cast<unsigned long long>(rpcAvBase.recordsCount);
            avBaseInfo.releaseDate = rpcAvBase.releaseDate;
        }

        void FillScanResult(ZIVPO_RPC_SCAN_RESULT const& rpcScanResult, ScanResult& scanResult)
        {
            scanResult.completed = rpcScanResult.completed != 0;
            scanResult.malicious = rpcScanResult.malicious != 0;
            scanResult.scannedObjects = static_cast<unsigned long long>(rpcScanResult.scannedObjects);
            scanResult.infectedObjects = static_cast<unsigned long long>(rpcScanResult.infectedObjects);
            scanResult.targetPath = rpcScanResult.targetPath;
            scanResult.detectedThreat = rpcScanResult.detectedThreat;
            scanResult.details = rpcScanResult.details;
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

        long RpcGetAvBaseInfoSafe(RPC_BINDING_HANDLE bindingHandle, ZIVPO_RPC_AV_BASE_INFO* avBaseInfo) noexcept
        {
            long rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            __try
            {
                rpcStatus = RpcGetAvBaseInfo(bindingHandle, avBaseInfo);
            }
            __except (EXCEPTION_EXECUTE_HANDLER)
            {
                rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            }
            return rpcStatus;
        }

        long RpcScanFileSafe(
            RPC_BINDING_HANDLE bindingHandle,
            wchar_t const* filePath,
            ZIVPO_RPC_SCAN_RESULT* scanResult) noexcept
        {
            long rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            __try
            {
                rpcStatus = RpcScanFile(bindingHandle, filePath, scanResult);
            }
            __except (EXCEPTION_EXECUTE_HANDLER)
            {
                rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            }
            return rpcStatus;
        }

        long RpcScanDirectorySafe(
            RPC_BINDING_HANDLE bindingHandle,
            wchar_t const* directoryPath,
            ZIVPO_RPC_SCAN_RESULT* scanResult) noexcept
        {
            long rpcStatus = ZIVPO_RPC_STATUS_NETWORK_ERROR;
            __try
            {
                rpcStatus = RpcScanDirectory(bindingHandle, directoryPath, scanResult);
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

    RpcCallResult GetAvBaseInfo(AvBaseInfo& avBaseInfo)
    {
        avBaseInfo = {};
        RPC_BINDING_HANDLE bindingHandle = nullptr;
        if (!CreateRpcBinding(bindingHandle))
        {
            return { RpcStatusCode::NetworkError, false };
        }

        ZIVPO_RPC_AV_BASE_INFO rpcAvBase{};
        long rpcStatus = RpcGetAvBaseInfoSafe(bindingHandle, &rpcAvBase);
        RpcBindingFree(&bindingHandle);
        FillAvBaseInfo(rpcAvBase, avBaseInfo);
        return BuildResultFromRpcStatus(rpcStatus);
    }

    RpcCallResult ScanFile(std::wstring_view filePath, ScanResult& scanResult)
    {
        scanResult = {};
        RPC_BINDING_HANDLE bindingHandle = nullptr;
        if (!CreateRpcBinding(bindingHandle))
        {
            return { RpcStatusCode::NetworkError, false };
        }

        ZIVPO_RPC_SCAN_RESULT rpcScanResult{};
        std::wstring path(filePath);
        long rpcStatus = RpcScanFileSafe(bindingHandle, path.c_str(), &rpcScanResult);
        RpcBindingFree(&bindingHandle);
        FillScanResult(rpcScanResult, scanResult);
        return BuildResultFromRpcStatus(rpcStatus);
    }

    RpcCallResult ScanDirectory(std::wstring_view directoryPath, ScanResult& scanResult)
    {
        scanResult = {};
        RPC_BINDING_HANDLE bindingHandle = nullptr;
        if (!CreateRpcBinding(bindingHandle))
        {
            return { RpcStatusCode::NetworkError, false };
        }

        ZIVPO_RPC_SCAN_RESULT rpcScanResult{};
        std::wstring path(directoryPath);
        long rpcStatus = RpcScanDirectorySafe(bindingHandle, path.c_str(), &rpcScanResult);
        RpcBindingFree(&bindingHandle);
        FillScanResult(rpcScanResult, scanResult);
        return BuildResultFromRpcStatus(rpcStatus);
    }
}
#endif
