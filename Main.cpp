#include "pch.h"
#include "App.xaml.h"

#include <MddBootstrap.h>
#include <WindowsAppSDK-VersionInfo.h>

namespace
{
    void TraceStartupMessage(std::wstring_view message) noexcept
    {
        OutputDebugStringW(message.data());
        OutputDebugStringW(L"\r\n");
    }

    int FailStartup(HRESULT hr, std::wstring_view stage) noexcept
    {
        wchar_t buffer[512]{};
        StringCchPrintfW(
            buffer,
            ARRAYSIZE(buffer),
            L"%s failed. HRESULT: 0x%08X",
            stage.data(),
            static_cast<unsigned int>(hr));

        TraceStartupMessage(buffer);
        MessageBoxW(nullptr, buffer, L"ZIVPO startup error", MB_OK | MB_ICONERROR);
        return static_cast<int>(hr);
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
    BootstrapGuard bootstrapGuard{};
    PACKAGE_VERSION minVersion{};
    minVersion.Version = WINDOWSAPPSDK_RUNTIME_VERSION_UINT64;

    constexpr auto bootstrapOptions =
        MddBootstrapInitializeOptions_OnNoMatch_ShowUI |
        MddBootstrapInitializeOptions_OnPackageIdentity_NOOP;

    const HRESULT bootstrapHr = MddBootstrapInitialize2(
        WINDOWSAPPSDK_RELEASE_MAJORMINOR,
        WINDOWSAPPSDK_RELEASE_VERSION_TAG_W,
        minVersion,
        bootstrapOptions);

    if (FAILED(bootstrapHr))
    {
        return FailStartup(bootstrapHr, L"MddBootstrapInitialize2");
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
        return FailStartup(ex.code().value, L"Microsoft.UI.Xaml::Application::Start");
    }
    catch (...)
    {
        return FailStartup(E_FAIL, L"Unhandled startup exception");
    }

    return 0;
}
