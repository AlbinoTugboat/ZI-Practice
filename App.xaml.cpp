#include "pch.h"
#include "App.xaml.h"

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
        constexpr std::wstring_view kHiddenSwitches[] = {
            L"--hidden",
            L"--background",
            L"/hidden",
            L"/background"
        };

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

        // When building without generated App.xaml glue, attach WinUI resources manually.
        try
        {
            ResourceDictionary resources{};
            resources.MergedDictionaries().Append(winrt::Microsoft::UI::Xaml::Controls::XamlControlsResources{});
            Resources(resources);
        }
        catch (...)
        {
            TraceMessage(L"Failed to initialize WinUI application resources.");
        }

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
            g_appInstance->ExitApplication();
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
            InitializeMainWindowContent();

            m_trayIcon = std::make_unique<::ZIVPO::TrayIconManager>();
            if (!m_trayIcon->Initialize(
                L"ZIVPO",
                [this] { ShowMainWindow(); },
                [this] { ExitApplication(); }))
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

    winrt::Microsoft::UI::Xaml::Markup::IXamlType App::GetXamlType(
        [[maybe_unused]] winrt::Windows::UI::Xaml::Interop::TypeName const& type)
    {
        return nullptr;
    }

    winrt::Microsoft::UI::Xaml::Markup::IXamlType App::GetXamlType(
        [[maybe_unused]] winrt::hstring const& fullName)
    {
        return nullptr;
    }

    winrt::com_array<winrt::Microsoft::UI::Xaml::Markup::XmlnsDefinition> App::GetXmlnsDefinitions()
    {
        return {};
    }

    void App::InitializeMainWindowContent()
    {
        namespace Controls = winrt::Microsoft::UI::Xaml::Controls;

        Controls::StackPanel root{};
        root.Padding(ThicknessHelper::FromUniformLength(12.0));
        root.Spacing(12.0);

        Controls::MenuBar menuBar{};
        Controls::MenuBarItem fileMenu{};
        fileMenu.Title(L"\x0424\x0430\x0439\x043B");

        Controls::MenuFlyoutItem exitItem{};
        exitItem.Text(L"\x0412\x044B\x0445\x043E\x0434");
        exitItem.Click([this](auto&&, auto&&)
        {
            ExitApplication();
        });

        fileMenu.Items().Append(exitItem);
        menuBar.Items().Append(fileMenu);
        root.Children().Append(menuBar);

        Controls::TextBlock header{};
        header.Text(L"The application keeps running in the tray.");
        header.FontSize(18.0);
        root.Children().Append(header);

        Controls::TextBlock details{};
        details.Text(L"Closing the window hides it to tray. Use Exit in menu or tray to stop the app.");
        root.Children().Append(details);

        m_window.Content(root);
        m_window.Title(L"ZIVPO");
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
        std::wstring args = arguments.c_str();

        for (auto const& token : kHiddenSwitches)
        {
            if (args.find(token) != std::wstring::npos)
            {
                return true;
            }
        }

        return false;
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

        if (m_originalMainWindowProc != nullptr)
        {
            return CallWindowProcW(m_originalMainWindowProc, hwnd, message, wParam, lParam);
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }
}
