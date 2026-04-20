#include "pch.h"
#include "App.xaml.h"
#include "ServiceManager.h"

using namespace winrt;
using namespace winrt::Windows::Foundation;
using namespace Microsoft::UI::Xaml;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace winrt::ZIVPO::implementation
{
    namespace
    {
        App* g_appInstance = nullptr;
        constexpr wchar_t kSingleInstanceMutexPrefix[] = L"Local\\ZIVPO.SingleInstance.";
        constexpr wchar_t kAppInstancePropertyName[] = L"ZIVPO.App.Instance";
        constexpr UINT_PTR kMainMenuFileId = 3001;
        constexpr UINT_PTR kMainMenuExitId = 3002;
        constexpr std::wstring_view kHiddenSwitches[] = {
            L"--hidden",
            L"--background",
            L"/hidden",
            L"/background"
        };

        bool ContainsHiddenSwitch(std::wstring_view args)
        {
            for (auto const& token : kHiddenSwitches)
            {
                if (args.find(token) != std::wstring::npos)
                {
                    return true;
                }
            }

            return false;
        }

        void TraceMessage(std::wstring_view message) noexcept
        {
            OutputDebugStringW(message.data());
            OutputDebugStringW(L"\r\n");
        }
    }

    /// <summary>
    /// Initializes the singleton application object.  This is the first line of authored code
    /// executed, and as such is the logical equivalent of main() or WinMain().
    /// </summary>
    App::App()
    {
        g_appInstance = this;

        // Xaml objects should not call InitializeComponent during construction.
        // See https://github.com/microsoft/cppwinrt/tree/master/nuget#initializecomponent

        UnhandledException([this](IInspectable const&, UnhandledExceptionEventArgs const& e)
        {
            auto const errorMessage = e.Message();
            std::wstring message = L"Unhandled exception: ";
            message.append(errorMessage.c_str());
            TraceMessage(message);
            e.Handled(true);
            ExitApplication();

#if defined _DEBUG && !defined DISABLE_XAML_GENERATED_BREAK_ON_UNHANDLED_EXCEPTION
            if (IsDebuggerPresent())
            {
                __debugbreak();
            }
#endif
        });
    }

    App::~App()
    {
        if (g_appInstance == this)
        {
            g_appInstance = nullptr;
        }

        if (m_trayIcon)
        {
            m_trayIcon->Shutdown();
            m_trayIcon.reset();
        }

        if (m_mainWindowHwnd != nullptr && m_originalMainWindowProc != nullptr && IsWindow(m_mainWindowHwnd))
        {
            SetWindowLongPtrW(
                m_mainWindowHwnd,
                GWLP_WNDPROC,
                reinterpret_cast<LONG_PTR>(m_originalMainWindowProc));
            RemovePropW(m_mainWindowHwnd, kAppInstancePropertyName);
        }

        if (m_singleInstanceMutex != nullptr)
        {
            CloseHandle(m_singleInstanceMutex);
            m_singleInstanceMutex = nullptr;
        }
    }

    void RequestApplicationExit()
    {
        if (g_appInstance != nullptr)
        {
            g_appInstance->StopServiceAndExitApplication();
        }
    }

    /// <summary>
    /// Invoked when the application is launched.
    /// </summary>
    /// <param name="e">Details about the launch request and process.</param>
    void App::OnLaunched([[maybe_unused]] LaunchActivatedEventArgs const& e)
    {
        try
        {
            if (!AcquireSingleInstanceMutex())
            {
                Exit();
                return;
            }

            m_window = Window();
            m_window.Title(L"ZIVPO");

            m_trayIcon = std::make_unique<::ZIVPO::TrayIconManager>();
            if (!m_trayIcon->Initialize(
                L"ZIVPO",
                [this] { ShowMainWindow(); },
                [this] { StopServiceAndExitApplication(); }))
            {
                ExitApplication();
                return;
            }

            if (ShouldStartHidden(e.Arguments()))
            {
                return;
            }

            ShowMainWindow();
        }
        catch (winrt::hresult_error const& ex)
        {
            std::wstring message = L"OnLaunched failed with HRESULT 0x";
            wchar_t hrBuffer[16]{};
            StringCchPrintfW(hrBuffer, ARRAYSIZE(hrBuffer), L"%08X", static_cast<unsigned int>(ex.code().value));
            message.append(hrBuffer);
            message.append(L": ");
            message.append(ex.message().c_str());
            TraceMessage(message);
            ExitApplication();
        }
        catch (...)
        {
            TraceMessage(L"OnLaunched failed with unknown exception.");
            ExitApplication();
        }
    }

    bool App::AcquireSingleInstanceMutex()
    {
        std::wstring mutexName = kSingleInstanceMutexPrefix;
        std::vector<wchar_t> userName(256, L'\0');
        DWORD userNameSize = static_cast<DWORD>(userName.size());

        if (GetUserNameW(userName.data(), &userNameSize) && userNameSize > 1)
        {
            mutexName.append(userName.data());
        }
        else
        {
            mutexName.append(L"default");
        }

        m_singleInstanceMutex = CreateMutexW(nullptr, FALSE, mutexName.c_str());
        if (m_singleInstanceMutex == nullptr)
        {
            return false;
        }

        return GetLastError() != ERROR_ALREADY_EXISTS;
    }

    bool App::ShouldStartHidden(hstring const& arguments)
    {
        std::wstring activationArgs = arguments.c_str();
        if (ContainsHiddenSwitch(activationArgs))
        {
            return true;
        }

        std::wstring processArgs = GetCommandLineW();
        return ContainsHiddenSwitch(processArgs);
    }

    HWND App::MainWindowHandle() const
    {
        if (!m_window)
        {
            return nullptr;
        }

        auto windowNative = m_window.try_as<IWindowNative>();
        if (!windowNative)
        {
            return nullptr;
        }

        HWND hwnd = nullptr;
        HRESULT hr = windowNative->get_WindowHandle(&hwnd);
        if (FAILED(hr))
        {
            return nullptr;
        }

        return hwnd;
    }

    void App::EnsureMainWindowHooked()
    {
        if (m_mainWindowHwnd != nullptr)
        {
            return;
        }

        m_mainWindowHwnd = MainWindowHandle();
        if (m_mainWindowHwnd == nullptr)
        {
            return;
        }

        SetPropW(m_mainWindowHwnd, kAppInstancePropertyName, this);
        m_originalMainWindowProc = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(
            m_mainWindowHwnd,
            GWLP_WNDPROC,
            reinterpret_cast<LONG_PTR>(&App::MainWindowProc)));
        EnsureMainWindowMenu();
    }

    void App::EnsureMainWindowMenu()
    {
        if (m_mainWindowHwnd == nullptr)
        {
            return;
        }

        if (GetMenu(m_mainWindowHwnd) != nullptr)
        {
            return;
        }

        HMENU mainMenu = CreateMenu();
        HMENU fileSubMenu = CreatePopupMenu();
        if (mainMenu == nullptr || fileSubMenu == nullptr)
        {
            if (fileSubMenu != nullptr)
            {
                DestroyMenu(fileSubMenu);
            }
            if (mainMenu != nullptr)
            {
                DestroyMenu(mainMenu);
            }
            return;
        }

        AppendMenuW(fileSubMenu, MF_STRING, kMainMenuExitId, L"\x0412\x044B\x0445\x043E\x0434");
        AppendMenuW(mainMenu, MF_POPUP, reinterpret_cast<UINT_PTR>(fileSubMenu), L"\x0424\x0430\x0439\x043B");
        SetMenu(m_mainWindowHwnd, mainMenu);
        DrawMenuBar(m_mainWindowHwnd);
    }

    void App::ShowMainWindow()
    {
        if (!m_window)
        {
            return;
        }

        try
        {
            m_window.Activate();
        }
        catch (...)
        {
            return;
        }

        EnsureMainWindowHooked();

        auto hwnd = MainWindowHandle();
        if (hwnd == nullptr)
        {
            return;
        }

        ShowWindow(hwnd, SW_RESTORE);
        ShowWindow(hwnd, SW_SHOW);
        SetForegroundWindow(hwnd);
    }

    void App::HideMainWindow()
    {
        auto hwnd = MainWindowHandle();
        if (hwnd == nullptr)
        {
            return;
        }

        ShowWindow(hwnd, SW_HIDE);
    }

    void App::ExitApplication()
    {
        m_isExiting = true;

        if (m_trayIcon)
        {
            m_trayIcon->Shutdown();
        }

        Exit();
    }

    void App::StopServiceAndExitApplication()
    {
        ::ZIVPO::Service::RequestServiceStop();
        ExitApplication();
    }

    LRESULT CALLBACK App::MainWindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
    {
        auto* self = reinterpret_cast<App*>(GetPropW(hwnd, kAppInstancePropertyName));
        if (self == nullptr)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        return self->HandleMainWindowMessage(hwnd, message, wParam, lParam);
    }

    LRESULT App::HandleMainWindowMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
    {
        if (message == WM_CLOSE && !m_isExiting)
        {
            HideMainWindow();
            return 0;
        }

        if (message == WM_COMMAND && LOWORD(wParam) == kMainMenuExitId)
        {
            StopServiceAndExitApplication();
            return 0;
        }

        if (m_originalMainWindowProc != nullptr)
        {
            return CallWindowProcW(m_originalMainWindowProc, hwnd, message, wParam, lParam);
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }
}
