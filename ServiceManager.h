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
        InternalError = 6
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

    int RunServiceMode();
    GuiStartupDecision PrepareGuiStartup();
    bool RequestServiceStop();

    RpcCallResult GetCurrentUser(AuthUserInfo& userInfo);
    RpcCallResult Login(std::wstring_view username, std::wstring_view password, AuthUserInfo& userInfo, LicenseInfo& licenseInfo);
    RpcCallResult Logout();
    RpcCallResult GetLicenseInfo(LicenseInfo& licenseInfo);
    RpcCallResult ActivateProduct(std::wstring_view activationKey, LicenseInfo& licenseInfo);
}
