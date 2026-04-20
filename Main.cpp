#include "pch.h"
#include "App.xaml.h"
#include "ServiceManager.h"

#include <MddBootstrap.h>
#include <WindowsAppSDK-VersionInfo.h>

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
}

int __stdcall wWinMain(HINSTANCE, HINSTANCE, PWSTR, int)
{
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
