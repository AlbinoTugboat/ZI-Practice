#include "pch.h"
#include "App.xaml.h"
#include "ServiceManager.h"

#include <Aclapi.h>
#include <MddBootstrap.h>
#include <WindowsAppSDK-VersionInfo.h>
#include <array>
#include <cwctype>
#include <vector>

namespace
{
    constexpr std::wstring_view kHiddenSwitches[] = {
        L"--hidden",
        L"--background",
        L"/hidden",
        L"/background"
    };

    void TraceStartupMessage(std::wstring_view message) noexcept
    {
        OutputDebugStringW(message.data());
        OutputDebugStringW(L"\r\n");
    }

    bool ContainsSwitch(std::wstring_view commandLine, std::wstring_view value) noexcept
    {
        return commandLine.find(value) != std::wstring::npos;
    }

    bool IsEnvironmentFlagEnabled(wchar_t const* name) noexcept
    {
        wchar_t value[16]{};
        DWORD const length = GetEnvironmentVariableW(name, value, ARRAYSIZE(value));
        if (length == 0 || length >= ARRAYSIZE(value))
        {
            return false;
        }

        std::wstring normalized(value, value + length);
        std::transform(normalized.begin(), normalized.end(), normalized.begin(), [](wchar_t ch)
        {
            return static_cast<wchar_t>(::towlower(ch));
        });

        return normalized == L"1" || normalized == L"true" || normalized == L"yes" || normalized == L"on";
    }

    bool IsHiddenStartupCommandLine() noexcept
    {
        std::wstring_view commandLine = GetCommandLineW();
        for (auto const& token : kHiddenSwitches)
        {
            if (ContainsSwitch(commandLine, token))
            {
                return true;
            }
        }

        return false;
    }

    int FailStartup(HRESULT hr, std::wstring_view stage, bool showMessageBox) noexcept
    {
        wchar_t buffer[512]{};
        StringCchPrintfW(
            buffer,
            ARRAYSIZE(buffer),
            L"%s failed. HRESULT: 0x%08X",
            stage.data(),
            static_cast<unsigned int>(hr));

        TraceStartupMessage(buffer);
        if (showMessageBox)
        {
            MessageBoxW(nullptr, buffer, L"ZIVPO startup error", MB_OK | MB_ICONERROR);
        }
        return static_cast<int>(hr);
    }

    HRESULT InitializeBootstrapWithRetry() noexcept
    {
        PACKAGE_VERSION minVersion{};
        minVersion.Version = WINDOWSAPPSDK_RUNTIME_VERSION_UINT64;

        constexpr auto bootstrapOptions =
            MddBootstrapInitializeOptions_OnNoMatch_ShowUI |
            MddBootstrapInitializeOptions_OnPackageIdentity_NOOP;
        constexpr DWORD kRetryDelayMs = 2000;
        constexpr int kMaxClassNotRegisteredRetries = 30;

        for (int attempt = 0; attempt <= kMaxClassNotRegisteredRetries; ++attempt)
        {
            HRESULT hr = MddBootstrapInitialize2(
                WINDOWSAPPSDK_RELEASE_MAJORMINOR,
                WINDOWSAPPSDK_RELEASE_VERSION_TAG_W,
                minVersion,
                bootstrapOptions);

            if (SUCCEEDED(hr))
            {
                return hr;
            }

            if (hr != REGDB_E_CLASSNOTREG || attempt == kMaxClassNotRegisteredRetries)
            {
                return hr;
            }

            wchar_t retryMessage[256]{};
            StringCchPrintfW(
                retryMessage,
                ARRAYSIZE(retryMessage),
                L"MddBootstrapInitialize2 returned REGDB_E_CLASSNOTREG, retry %d/%d in %lu ms",
                attempt + 1,
                kMaxClassNotRegisteredRetries,
                kRetryDelayMs);
            TraceStartupMessage(retryMessage);

            Sleep(kRetryDelayMs);
        }

        return E_FAIL;
    }

    struct BootstrapGuard final
    {
        bool initialized{ false };
        ~BootstrapGuard()
        {
            if (initialized)
            {
                MddBootstrapShutdown();
            }
        }
    };

    void ApplyProcessMitigations() noexcept
    {
        PROCESS_MITIGATION_EXTENSION_POINT_DISABLE_POLICY extensionPointPolicy{};
        extensionPointPolicy.DisableExtensionPoints = 1;
        SetProcessMitigationPolicy(
            ProcessExtensionPointDisablePolicy,
            &extensionPointPolicy,
            sizeof(extensionPointPolicy));
    }

    void ApplyMicrosoftSignedOnlyMitigation() noexcept
    {
        PROCESS_MITIGATION_BINARY_SIGNATURE_POLICY signaturePolicy{};
        signaturePolicy.MicrosoftSignedOnly = 1;
        SetProcessMitigationPolicy(
            ProcessSignaturePolicy,
            &signaturePolicy,
            sizeof(signaturePolicy));
    }

    bool GetCurrentUserSid(std::vector<BYTE>& sidBytes) noexcept
    {
        HANDLE token = nullptr;
        if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &token))
        {
            return false;
        }

        DWORD required = 0;
        GetTokenInformation(token, TokenUser, nullptr, 0, &required);
        if (required == 0)
        {
            CloseHandle(token);
            return false;
        }

        std::vector<BYTE> tokenInfo(required);
        if (!GetTokenInformation(token, TokenUser, tokenInfo.data(), required, &required))
        {
            CloseHandle(token);
            return false;
        }
        CloseHandle(token);

        TOKEN_USER const* tokenUser = reinterpret_cast<TOKEN_USER const*>(tokenInfo.data());
        DWORD sidLength = GetLengthSid(tokenUser->User.Sid);
        if (sidLength == 0)
        {
            return false;
        }

        sidBytes.assign(sidLength, 0);
        return CopySid(sidLength, sidBytes.data(), tokenUser->User.Sid) != FALSE;
    }

    bool ApplyProcessTerminationProtection() noexcept
    {
        std::vector<BYTE> currentUserSid;
        if (!GetCurrentUserSid(currentUserSid))
        {
            return false;
        }

        std::array<BYTE, SECURITY_MAX_SID_SIZE> adminSidBuffer{};
        std::array<BYTE, SECURITY_MAX_SID_SIZE> systemSidBuffer{};
        DWORD adminSidSize = static_cast<DWORD>(adminSidBuffer.size());
        DWORD systemSidSize = static_cast<DWORD>(systemSidBuffer.size());
        if (!CreateWellKnownSid(WinBuiltinAdministratorsSid, nullptr, adminSidBuffer.data(), &adminSidSize) ||
            !CreateWellKnownSid(WinLocalSystemSid, nullptr, systemSidBuffer.data(), &systemSidSize))
        {
            return false;
        }

        EXPLICIT_ACCESSW entries[3]{};
        auto setEntry = [](EXPLICIT_ACCESSW& entry, PSID sid, DWORD permissions)
        {
            entry.grfAccessPermissions = permissions;
            entry.grfAccessMode = SET_ACCESS;
            entry.grfInheritance = NO_INHERITANCE;
            entry.Trustee.TrusteeForm = TRUSTEE_IS_SID;
            entry.Trustee.TrusteeType = TRUSTEE_IS_USER;
            entry.Trustee.ptstrName = static_cast<LPWSTR>(sid);
        };

        setEntry(entries[0], systemSidBuffer.data(), PROCESS_ALL_ACCESS);
        setEntry(entries[1], adminSidBuffer.data(), PROCESS_ALL_ACCESS);
        setEntry(
            entries[2],
            currentUserSid.data(),
            PROCESS_QUERY_LIMITED_INFORMATION |
            PROCESS_SET_LIMITED_INFORMATION |
            PROCESS_VM_READ |
            SYNCHRONIZE |
            READ_CONTROL);

        PACL dacl = nullptr;
        DWORD aclStatus = SetEntriesInAclW(static_cast<ULONG>(std::size(entries)), entries, nullptr, &dacl);
        if (aclStatus != ERROR_SUCCESS)
        {
            return false;
        }

        DWORD securityStatus = SetSecurityInfo(
            GetCurrentProcess(),
            SE_KERNEL_OBJECT,
            DACL_SECURITY_INFORMATION | PROTECTED_DACL_SECURITY_INFORMATION,
            nullptr,
            nullptr,
            dacl,
            nullptr);
        LocalFree(dacl);

        return securityStatus == ERROR_SUCCESS;
    }
}

int __stdcall wWinMain(HINSTANCE, HINSTANCE, PWSTR, int)
{
    if (IsEnvironmentFlagEnabled(L"ZIVPO_ENABLE_EXTENSION_POINT_MITIGATION"))
    {
        ApplyProcessMitigations();
    }
    bool const hiddenStartup = IsHiddenStartupCommandLine();

    switch (ZIVPO::Service::PrepareGuiStartup())
    {
    case ZIVPO::Service::GuiStartupDecision::Continue:
        break;
    case ZIVPO::Service::GuiStartupDecision::Exit:
        return 0;
    case ZIVPO::Service::GuiStartupDecision::Error:
    default:
        return FailStartup(E_FAIL, L"PrepareGuiStartup", !hiddenStartup);
    }

    BootstrapGuard bootstrapGuard{};
    const HRESULT bootstrapHr = InitializeBootstrapWithRetry();

    if (FAILED(bootstrapHr))
    {
        return FailStartup(bootstrapHr, L"MddBootstrapInitialize2", !hiddenStartup);
    }

    bootstrapGuard.initialized = true;

    if (IsEnvironmentFlagEnabled(L"ZIVPO_ENABLE_MICROSOFT_DLL_ONLY"))
    {
        ApplyMicrosoftSignedOnlyMitigation();
    }

    if (!ApplyProcessTerminationProtection())
    {
        TraceStartupMessage(L"Failed to apply process termination protection.");
    }

    try
    {
        winrt::init_apartment(winrt::apartment_type::single_threaded);
        winrt::Microsoft::UI::Xaml::Application::Start(
            [](auto&&)
            {
                winrt::make<winrt::ZIVPO::implementation::App>();
            });
    }
    catch (winrt::hresult_error const& ex)
    {
        return FailStartup(ex.code().value, L"Microsoft.UI.Xaml::Application::Start", !hiddenStartup);
    }
    catch (...)
    {
        return FailStartup(E_FAIL, L"Unhandled startup exception", !hiddenStartup);
    }

    return 0;
}
