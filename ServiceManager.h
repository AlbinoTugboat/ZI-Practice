#pragma once
#include <string>
#include <string_view>

namespace ZIVPO::Service
{
    enum class RpcStatusCode : int
    {
        Ok = 0,
        NotAuthenticated = 1,
        InvalidCredentials = 2,
        NoLicense = 3,
        ActivationFailed = 4,
        NetworkError = 5,
        InternalError = 6,
        InvalidArgument = 7,
        ScanFailed = 8
    };

    enum class GuiStartupDecision
    {
        Continue,
        Exit,
        Error
    };

    struct AuthUserInfo
    {
        bool authenticated{ false };
        std::wstring username;
    };

    struct LicenseInfo
    {
        bool hasLicense{ false };
        bool blocked{ false };
        std::wstring expirationDate;
    };

    struct RpcCallResult
    {
        RpcStatusCode status{ RpcStatusCode::InternalError };
        bool ok{ false };
    };

    struct AvBaseInfo
    {
        bool loaded{ false };
        unsigned long long recordsCount{ 0 };
        std::wstring releaseDate;
    };

    struct ScanResult
    {
        bool completed{ false };
        bool malicious{ false };
        unsigned long long scannedObjects{ 0 };
        unsigned long long infectedObjects{ 0 };
        std::wstring targetPath;
        std::wstring detectedThreat;
        std::wstring details;
    };

    int RunServiceMode();
    GuiStartupDecision PrepareGuiStartup();
    bool RequestServiceStop();

    RpcCallResult GetCurrentUser(AuthUserInfo& userInfo);
    RpcCallResult Login(std::wstring_view username, std::wstring_view password, AuthUserInfo& userInfo, LicenseInfo& licenseInfo);
    RpcCallResult Logout();
    RpcCallResult GetLicenseInfo(LicenseInfo& licenseInfo);
    RpcCallResult ActivateProduct(std::wstring_view activationKey, LicenseInfo& licenseInfo);
    RpcCallResult GetAvBaseInfo(AvBaseInfo& avBaseInfo);
    RpcCallResult ScanFile(std::wstring_view filePath, ScanResult& scanResult);
    RpcCallResult ScanDirectory(std::wstring_view directoryPath, ScanResult& scanResult);
}
