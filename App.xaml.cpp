#include "pch.h"
#include "App.xaml.h"
#include "ServiceManager.h"
#include <chrono>
#include <cwctype>
#include <mutex>
#include <combaseapi.h>
#include <wincred.h>

#pragma comment(lib, "Credui.lib")

using namespace winrt;
using namespace winrt::Windows::Foundation;
using namespace Microsoft::UI::Xaml;
using namespace Microsoft::UI::Xaml::Controls;

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

        bool IsEnvironmentFlagEnabled(wchar_t const* name)
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

        void TraceMessage(std::wstring_view message) noexcept
        {
            static std::mutex s_logMutex;
            std::lock_guard lock(s_logMutex);

            SYSTEMTIME localTime{};
            GetLocalTime(&localTime);
            DWORD const processId = GetCurrentProcessId();
            DWORD const threadId = GetCurrentThreadId();

            wchar_t line[2048]{};
            StringCchPrintfW(
                line,
                ARRAYSIZE(line),
                L"[%04u-%02u-%02u %02u:%02u:%02u.%03u][pid:%lu tid:%lu] %s",
                localTime.wYear,
                localTime.wMonth,
                localTime.wDay,
                localTime.wHour,
                localTime.wMinute,
                localTime.wSecond,
                localTime.wMilliseconds,
                processId,
                threadId,
                message.data());

            OutputDebugStringW(message.data());
            OutputDebugStringW(L"\r\n");

            wchar_t logPath[MAX_PATH]{};
            DWORD const envLength = GetEnvironmentVariableW(L"LOCALAPPDATA", logPath, ARRAYSIZE(logPath));
            if (envLength == 0 || envLength >= ARRAYSIZE(logPath) - 24)
            {
                return;
            }

            StringCchCatW(logPath, ARRAYSIZE(logPath), L"\\ZIVPO");
            CreateDirectoryW(logPath, nullptr);
            StringCchCatW(logPath, ARRAYSIZE(logPath), L"\\zivpo_ui.log");

            HANDLE const file = CreateFileW(
                logPath,
                FILE_APPEND_DATA,
                FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                nullptr,
                OPEN_ALWAYS,
                FILE_ATTRIBUTE_NORMAL,
                nullptr);
            if (file == INVALID_HANDLE_VALUE)
            {
                return;
            }

            StringCchCatW(line, ARRAYSIZE(line), L"\r\n");
            DWORD bytesWritten = 0;
            WriteFile(
                file,
                line,
                static_cast<DWORD>(wcslen(line) * sizeof(wchar_t)),
                &bytesWritten,
                nullptr);
            CloseHandle(file);
        }

        hstring BuildRpcErrorText(::ZIVPO::Service::RpcStatusCode status)
        {
            switch (status)
            {
            case ::ZIVPO::Service::RpcStatusCode::InvalidCredentials:
                return L"Invalid username or password.";
            case ::ZIVPO::Service::RpcStatusCode::NotAuthenticated:
                return L"User is not authenticated.";
            case ::ZIVPO::Service::RpcStatusCode::NoLicense:
                return L"License is missing.";
            case ::ZIVPO::Service::RpcStatusCode::ActivationFailed:
                return L"Activation failed. Check activation key.";
            case ::ZIVPO::Service::RpcStatusCode::NetworkError:
                return L"Network error or service is unavailable.";
            default:
                return L"Internal service error.";
            }
        }

        bool PromptForCredentials(
            HWND ownerWindow,
            std::wstring_view caption,
            std::wstring_view message,
            bool requirePassword,
            std::wstring& username,
            std::wstring& password,
            hstring& errorText)
        {
            username.clear();
            password.clear();
            errorText = L"";

            CREDUI_INFOW promptInfo{};
            promptInfo.cbSize = sizeof(promptInfo);
            promptInfo.hwndParent = ownerWindow;
            promptInfo.pszCaptionText = caption.data();
            promptInfo.pszMessageText = message.data();

            wchar_t usernameBuffer[CREDUI_MAX_USERNAME_LENGTH + 1]{};
            wchar_t passwordBuffer[CREDUI_MAX_PASSWORD_LENGTH + 1]{};
            BOOL saveRequested = FALSE;

            constexpr DWORD promptFlags =
                CREDUI_FLAGS_GENERIC_CREDENTIALS |
                CREDUI_FLAGS_ALWAYS_SHOW_UI |
                CREDUI_FLAGS_DO_NOT_PERSIST |
                CREDUI_FLAGS_EXCLUDE_CERTIFICATES;

            DWORD const promptStatus = CredUIPromptForCredentialsW(
                &promptInfo,
                L"ZIVPO",
                nullptr,
                0,
                usernameBuffer,
                ARRAYSIZE(usernameBuffer),
                passwordBuffer,
                ARRAYSIZE(passwordBuffer),
                &saveRequested,
                promptFlags);

            if (promptStatus == ERROR_CANCELLED)
            {
                SecureZeroMemory(passwordBuffer, sizeof(passwordBuffer));
                return false;
            }

            if (promptStatus != NO_ERROR)
            {
                SecureZeroMemory(passwordBuffer, sizeof(passwordBuffer));
                errorText = L"Unable to open credentials prompt.";
                return false;
            }

            username.assign(usernameBuffer);
            password.assign(passwordBuffer);
            SecureZeroMemory(passwordBuffer, sizeof(passwordBuffer));

            if (username.empty() || (requirePassword && password.empty()))
            {
                errorText = requirePassword
                    ? L"Username and password are required."
                    : L"Activation key is required.";
                return false;
            }

            return true;
        }

        bool PromptForSecureDesktopConfirmation(
            HWND ownerWindow,
            std::wstring_view caption,
            std::wstring_view message,
            hstring& errorText)
        {
            errorText = L"";

            CREDUI_INFOW promptInfo{};
            promptInfo.cbSize = sizeof(promptInfo);
            promptInfo.hwndParent = ownerWindow;
            promptInfo.pszCaptionText = caption.data();
            promptInfo.pszMessageText = message.data();

            ULONG authPackage = 0;
            LPVOID authBuffer = nullptr;
            ULONG authBufferSize = 0;
            BOOL saveRequested = FALSE;

            constexpr DWORD promptFlags =
                CREDUIWIN_SECURE_PROMPT |
                CREDUIWIN_ENUMERATE_CURRENT_USER;

            DWORD const promptStatus = CredUIPromptForWindowsCredentialsW(
                &promptInfo,
                0,
                &authPackage,
                nullptr,
                0,
                &authBuffer,
                &authBufferSize,
                &saveRequested,
                promptFlags);

            if (promptStatus == ERROR_CANCELLED)
            {
                return false;
            }

            if (promptStatus != NO_ERROR)
            {
                errorText = L"Unable to open secure desktop confirmation.";
                return false;
            }

            if (authBuffer != nullptr)
            {
                if (authBufferSize > 0)
                {
                    SecureZeroMemory(authBuffer, authBufferSize);
                }
                CoTaskMemFree(authBuffer);
            }

            return true;
        }
    }

    /// <summary>
    /// Initializes the singleton application object.  This is the first line of authored code
    /// executed, and as such is the logical equivalent of main() or WinMain().
    /// </summary>
    App::App()
    {
        g_appInstance = this;
        TraceMessage(L"App::App begin");
        // Load App.xaml resources (XamlControlsResources and app-level dictionaries)
        // before creating controls in code.
        try
        {
            winrt::Windows::Foundation::Uri resourceLocator{ L"ms-appx:///App.xaml" };
            winrt::Microsoft::UI::Xaml::Application::LoadComponent(*this, resourceLocator);
            TraceMessage(L"App::App loaded App.xaml resources successfully");
        }
        catch (...)
        {
            TraceMessage(L"App::App: failed to load App.xaml resources.");
        }
        RequestedTheme(ApplicationTheme::Light);
        TraceMessage(L"App::App requested theme set");

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
        if (m_licensePollTimer != nullptr)
        {
            m_licensePollTimer.Stop();
            m_licensePollTimer = nullptr;
        }

        if (g_appInstance == this)
        {
            g_appInstance = nullptr;
        }

        if (m_trayIcon)
        {
            m_trayIcon->Shutdown();
            m_trayIcon.reset();
        }

        if (m_window && m_windowClosedHandlerAttached)
        {
            try
            {
                m_window.Closed(m_windowClosedToken);
            }
            catch (...)
            {
                // Ignore shutdown races.
            }
            m_windowClosedHandlerAttached = false;
        }

        if (m_mainWindowHwnd != nullptr && m_originalMainWindowProc != nullptr && IsWindow(m_mainWindowHwnd))
        {
            SetWindowLongPtrW(
                m_mainWindowHwnd,
                GWLP_WNDPROC,
                reinterpret_cast<LONG_PTR>(m_originalMainWindowProc));
            RemovePropW(m_mainWindowHwnd, kAppInstancePropertyName);
        }
        m_mainWindowHwnd = nullptr;
        m_originalMainWindowProc = nullptr;

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
            TraceMessage(L"OnLaunched entered");
            if (!AcquireSingleInstanceMutex())
            {
                TraceMessage(L"OnLaunched: single instance mutex acquisition failed");
                Exit();
                return;
            }

            bool const hiddenStartup = ShouldStartHidden(e.Arguments());
            m_uiDispatcherQueue = winrt::Microsoft::UI::Dispatching::DispatcherQueue::GetForCurrentThread();

            bool const disableTrayIcon = IsEnvironmentFlagEnabled(L"ZIVPO_DISABLE_TRAY_ICON");
            if (!disableTrayIcon)
            {
                m_trayIcon = std::make_unique<::ZIVPO::TrayIconManager>();
                if (!m_trayIcon->Initialize(
                    L"ZIVPO",
                    [this]
                    {
                        RunOnUiThread([this]
                        {
                            ShowMainWindow();
                        });
                    },
                    [this]
                    {
                        RunOnUiThread([this]
                        {
                            StopServiceAndExitApplication();
                        });
                    }))
                {
                    TraceMessage(L"OnLaunched: tray icon initialization failed");
                    ExitApplication();
                    return;
                }
                TraceMessage(L"OnLaunched: tray icon initialized");
            }
            else
            {
                TraceMessage(L"OnLaunched: tray icon disabled by environment flag");
            }

            if (hiddenStartup)
            {
                TraceMessage(L"OnLaunched: hidden startup, UI initialization skipped");
                return;
            }

            TraceMessage(L"OnLaunched: initializing UI for visible startup");
            EnsureUiInitialized(false);
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
        if (IsEnvironmentFlagEnabled(L"ZIVPO_ALLOW_MULTI_INSTANCE"))
        {
            return true;
        }

        DWORD sessionId = 0;
        ProcessIdToSessionId(GetCurrentProcessId(), &sessionId);

        std::wstring mutexName = kSingleInstanceMutexPrefix;
        mutexName.append(L"Session.");
        mutexName.append(std::to_wstring(sessionId));

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

    void App::EnsureWindowCreated()
    {
        if (m_window != nullptr)
        {
            return;
        }

        m_window = Window();
        m_window.Title(L"ZIVPO");
        m_windowClosedToken = m_window.Closed([this](auto const&, auto const&)
        {
            TraceMessage(L"Window.Closed handler invoked");
            if (!m_isExiting)
            {
                if (m_licensePollTimer != nullptr)
                {
                    m_licensePollTimer.Stop();
                    m_licensePollTimer = nullptr;
                }

                m_uiInitialized = false;
                m_mainWindowVisible = false;
                ResetUiElementHandles();
                m_window = nullptr;
            }
            m_mainWindowHwnd = nullptr;
            m_originalMainWindowProc = nullptr;
            m_windowClosedHandlerAttached = false;
        });
        m_windowClosedHandlerAttached = true;
    }

    void App::ResetUiElementHandles()
    {
        m_userStatusText = nullptr;
        m_licenseStatusText = nullptr;
        m_antivirusStatusText = nullptr;
        m_antivirusActionButton = nullptr;
        m_logoutButton = nullptr;
        m_authPanel = nullptr;
        m_usernameInput = nullptr;
        m_passwordInput = nullptr;
        m_authErrorText = nullptr;
        m_activationPanel = nullptr;
        m_activationKeyInput = nullptr;
        m_activationErrorText = nullptr;
    }

    void App::BuildMainUi()
    {
        TraceMessage(L"BuildMainUi begin");
        bool const minimalProbeEnabled = IsEnvironmentFlagEnabled(L"ZIVPO_MINIMAL_UI");

        auto root = StackPanel{};
        root.Spacing(8);
        root.Padding(Thickness{ 12, 12, 12, 12 });
        root.RequestedTheme(ElementTheme::Light);

        if (minimalProbeEnabled)
        {
            auto probeText = TextBlock{};
            probeText.Text(L"Minimal UI probe mode");
            root.Children().Append(probeText);
            m_window.Content(root);
            TraceMessage(L"BuildMainUi: minimal UI probe mode enabled");
            TraceMessage(L"BuildMainUi end");
            return;
        }

        auto content = StackPanel{};
        content.Spacing(10);

        m_userStatusText = TextBlock{};
        m_userStatusText.FontSize(16);
        content.Children().Append(m_userStatusText);

        m_licenseStatusText = TextBlock{};
        content.Children().Append(m_licenseStatusText);

        m_antivirusStatusText = TextBlock{};
        content.Children().Append(m_antivirusStatusText);

        m_antivirusActionButton = Button{};
        m_antivirusActionButton.Content(box_value(L"Run Scan (demo)"));
        m_antivirusActionButton.IsEnabled(false);
        content.Children().Append(m_antivirusActionButton);

        m_logoutButton = Button{};
        m_logoutButton.Content(box_value(L"Log Out"));
        m_logoutButton.Click({ this, &App::OnLogoutClick });
        m_logoutButton.Visibility(Visibility::Collapsed);
        content.Children().Append(m_logoutButton);

        m_authPanel = StackPanel{};
        m_authPanel.Spacing(8);

        auto authHeader = TextBlock{};
        authHeader.Text(L"Sign In");
        authHeader.FontSize(16);
        m_authPanel.Children().Append(authHeader);

        auto loginButton = Button{};
        loginButton.Content(box_value(L"Sign In..."));
        loginButton.Click({ this, &App::OnLoginClick });
        m_authPanel.Children().Append(loginButton);

        m_authErrorText = TextBlock{};
        m_authErrorText.TextWrapping(TextWrapping::Wrap);
        m_authPanel.Children().Append(m_authErrorText);

        content.Children().Append(m_authPanel);

        m_activationPanel = StackPanel{};
        m_activationPanel.Spacing(8);

        auto activationHeader = TextBlock{};
        activationHeader.Text(L"Product Activation");
        activationHeader.FontSize(16);
        m_activationPanel.Children().Append(activationHeader);

        auto activateButton = Button{};
        activateButton.Content(box_value(L"Activate..."));
        activateButton.Click({ this, &App::OnActivateClick });
        m_activationPanel.Children().Append(activateButton);

        m_activationErrorText = TextBlock{};
        m_activationErrorText.TextWrapping(TextWrapping::Wrap);
        m_activationPanel.Children().Append(m_activationErrorText);

        content.Children().Append(m_activationPanel);
        root.Children().Append(content);

        m_window.Content(root);
        TraceMessage(L"BuildMainUi end");
    }

    void App::EnsureUiInitialized(bool showErrorsOnRefresh)
    {
        (void)showErrorsOnRefresh;

        if (m_uiInitialized)
        {
            TraceMessage(L"EnsureUiInitialized: already initialized");
            return;
        }

        TraceMessage(L"EnsureUiInitialized: creating window and building UI");
        EnsureWindowCreated();
        try
        {
            m_window.Activate();
            TraceMessage(L"EnsureUiInitialized: window activated before BuildMainUi");
        }
        catch (...)
        {
            TraceMessage(L"EnsureUiInitialized: m_window.Activate failed before BuildMainUi");
        }
        BuildMainUi();
        EnsureMainWindowHooked();
        m_uiInitialized = true;
        TraceMessage(L"EnsureUiInitialized: completed");
    }

    void App::OnLoginClick(IInspectable const& /*sender*/, RoutedEventArgs const& /*args*/)
    {
        m_authErrorText.Text(L"");
        m_activationErrorText.Text(L"");

        std::wstring username;
        std::wstring password;
        hstring promptError;
        if (!PromptForCredentials(
            MainWindowHandle(),
            L"ZIVPO Sign In",
            L"Enter your username and password.",
            true,
            username,
            password,
            promptError))
        {
            if (!promptError.empty())
            {
                m_authErrorText.Text(promptError);
            }
            return;
        }

        ::ZIVPO::Service::AuthUserInfo userInfo{};
        ::ZIVPO::Service::LicenseInfo licenseInfo{};
        ::ZIVPO::Service::RpcCallResult result = ::ZIVPO::Service::Login(username, password, userInfo, licenseInfo);
        if (!password.empty())
        {
            SecureZeroMemory(password.data(), password.size() * sizeof(wchar_t));
            password.clear();
        }
        if (!result.ok)
        {
            m_authenticated = false;
            m_hasLicense = false;
            m_licenseBlocked = false;
            m_username.clear();
            m_licenseExpiration.clear();
            m_authErrorText.Text(BuildRpcErrorText(result.status));
            UpdateUiState();
            return;
        }

        m_authenticated = userInfo.authenticated;
        m_username = userInfo.username;
        m_hasLicense = licenseInfo.hasLicense;
        m_licenseBlocked = licenseInfo.blocked;
        m_licenseExpiration = licenseInfo.expirationDate;
        UpdateUiState();
    }

    void App::OnLogoutClick(IInspectable const& /*sender*/, RoutedEventArgs const& /*args*/)
    {
        m_authErrorText.Text(L"");
        m_activationErrorText.Text(L"");

        hstring confirmError;
        if (!PromptForSecureDesktopConfirmation(
            MainWindowHandle(),
            L"ZIVPO Secure Confirmation",
            L"Confirm log out.",
            confirmError))
        {
            if (!confirmError.empty())
            {
                m_authErrorText.Text(confirmError);
            }
            return;
        }

        ::ZIVPO::Service::RpcCallResult result = ::ZIVPO::Service::Logout();
        if (!result.ok)
        {
            m_authErrorText.Text(BuildRpcErrorText(result.status));
        }

        m_authenticated = false;
        m_hasLicense = false;
        m_licenseBlocked = false;
        m_username.clear();
        m_licenseExpiration.clear();
        UpdateUiState();
    }

    void App::OnActivateClick(IInspectable const& /*sender*/, RoutedEventArgs const& /*args*/)
    {
        m_activationErrorText.Text(L"");
        if (!m_authenticated)
        {
            m_activationErrorText.Text(L"Sign in first.");
            return;
        }

        std::wstring activationKey;
        std::wstring unusedPassword;
        hstring promptError;
        if (!PromptForCredentials(
            MainWindowHandle(),
            L"ZIVPO Activation",
            L"Enter activation key in the username field.",
            false,
            activationKey,
            unusedPassword,
            promptError))
        {
            if (!promptError.empty())
            {
                m_activationErrorText.Text(promptError);
            }
            return;
        }

        ::ZIVPO::Service::LicenseInfo licenseInfo{};
        ::ZIVPO::Service::RpcCallResult result = ::ZIVPO::Service::ActivateProduct(activationKey, licenseInfo);
        if (!result.ok)
        {
            m_hasLicense = false;
            m_licenseBlocked = false;
            m_licenseExpiration.clear();
            m_activationErrorText.Text(BuildRpcErrorText(result.status));
            UpdateUiState();
            return;
        }

        m_hasLicense = licenseInfo.hasLicense;
        m_licenseBlocked = licenseInfo.blocked;
        m_licenseExpiration = licenseInfo.expirationDate;
        UpdateUiState();
    }

    void App::RefreshStateFromService(bool showErrors)
    {
        TraceMessage(L"RefreshStateFromService begin");
        ::ZIVPO::Service::AuthUserInfo userInfo{};
        ::ZIVPO::Service::RpcCallResult userResult = ::ZIVPO::Service::GetCurrentUser(userInfo);
        TraceMessage(L"RefreshStateFromService: GetCurrentUser returned");
        if (userResult.status == ::ZIVPO::Service::RpcStatusCode::NotAuthenticated)
        {
            m_authenticated = false;
            m_hasLicense = false;
            m_licenseBlocked = false;
            m_username.clear();
            m_licenseExpiration.clear();
            if (showErrors)
            {
                m_authErrorText.Text(L"");
            }
            UpdateUiState();
            TraceMessage(L"RefreshStateFromService: not authenticated path end");
            return;
        }

        if (!userResult.ok)
        {
            if (showErrors)
            {
                m_authErrorText.Text(BuildRpcErrorText(userResult.status));
            }
            m_authenticated = false;
            m_hasLicense = false;
            m_licenseBlocked = false;
            m_username.clear();
            m_licenseExpiration.clear();
            UpdateUiState();
            TraceMessage(L"RefreshStateFromService: GetCurrentUser error path end");
            return;
        }

        m_authenticated = userInfo.authenticated;
        m_username = userInfo.username;

        ::ZIVPO::Service::LicenseInfo licenseInfo{};
        ::ZIVPO::Service::RpcCallResult licenseResult = ::ZIVPO::Service::GetLicenseInfo(licenseInfo);
        TraceMessage(L"RefreshStateFromService: GetLicenseInfo returned");
        if (licenseResult.ok)
        {
            m_hasLicense = licenseInfo.hasLicense;
            m_licenseBlocked = licenseInfo.blocked;
            m_licenseExpiration = licenseInfo.expirationDate;
        }
        else if (licenseResult.status == ::ZIVPO::Service::RpcStatusCode::NoLicense)
        {
            m_hasLicense = false;
            m_licenseBlocked = false;
            m_licenseExpiration.clear();
        }
        else if (licenseResult.status == ::ZIVPO::Service::RpcStatusCode::NotAuthenticated)
        {
            m_authenticated = false;
            m_hasLicense = false;
            m_licenseBlocked = false;
            m_username.clear();
            m_licenseExpiration.clear();
        }
        else if (showErrors)
        {
            m_activationErrorText.Text(BuildRpcErrorText(licenseResult.status));
        }

        UpdateUiState();
        TraceMessage(L"RefreshStateFromService end");
    }

    void App::SetAntivirusEnabled(bool enabled)
    {
        if (m_antivirusActionButton == nullptr || m_antivirusStatusText == nullptr)
        {
            return;
        }

        m_antivirusActionButton.IsEnabled(enabled);
        m_antivirusStatusText.Text(enabled ? L"Antivirus: enabled" : L"Antivirus: disabled");
    }

    void App::UpdateUiState()
    {
        TraceMessage(L"UpdateUiState begin");
        if (!m_uiInitialized ||
            m_userStatusText == nullptr ||
            m_licenseStatusText == nullptr ||
            m_authPanel == nullptr ||
            m_activationPanel == nullptr ||
            m_logoutButton == nullptr)
        {
            TraceMessage(L"UpdateUiState: controls are not initialized");
            return;
        }

        if (!m_authenticated)
        {
            m_userStatusText.Text(L"User: not authenticated");
            m_licenseStatusText.Text(L"License: missing");
            m_authPanel.Visibility(Visibility::Visible);
            m_activationPanel.Visibility(Visibility::Collapsed);
            m_logoutButton.Visibility(Visibility::Collapsed);
            SetAntivirusEnabled(false);
            TraceMessage(L"UpdateUiState end (unauthenticated)");
            return;
        }

        std::wstring userLine = L"User: ";
        userLine.append(m_username.empty() ? L"(unknown)" : m_username);
        m_userStatusText.Text(hstring(userLine));
        m_authPanel.Visibility(Visibility::Collapsed);
        m_logoutButton.Visibility(Visibility::Visible);

        if (!m_hasLicense)
        {
            m_licenseStatusText.Text(L"License: missing, activation required");
            m_activationPanel.Visibility(Visibility::Visible);
            SetAntivirusEnabled(false);
            TraceMessage(L"UpdateUiState end (no license)");
            return;
        }

        std::wstring licenseLine = L"License expires: ";
        licenseLine.append(m_licenseExpiration.empty() ? L"(unknown)" : m_licenseExpiration);
        if (m_licenseBlocked)
        {
            licenseLine.append(L" (blocked)");
        }

        m_licenseStatusText.Text(hstring(licenseLine));
        m_activationPanel.Visibility(Visibility::Collapsed);
        SetAntivirusEnabled(!m_licenseBlocked);
        TraceMessage(L"UpdateUiState end (licensed)");
    }

    void App::StartLicensePolling()
    {
        if (m_licensePollTimer != nullptr)
        {
            return;
        }

        m_licensePollTimer = DispatcherTimer();
        m_licensePollTimer.Interval(std::chrono::seconds(15));
        m_licensePollTimer.Tick([this](auto const&, auto const&)
        {
            RefreshStateFromService(false);
        });
        m_licensePollTimer.Start();
    }

    void App::OnMainMenuExitClick(IInspectable const& /*sender*/, RoutedEventArgs const& /*args*/)
    {
        StopServiceAndExitApplication();
    }

    void App::RunOnUiThread(std::function<void()> action)
    {
        if (!action)
        {
            return;
        }

        if (m_uiDispatcherQueue == nullptr)
        {
            action();
            return;
        }

        bool const enqueued = m_uiDispatcherQueue.TryEnqueue([action = std::move(action)]
        {
            TraceMessage(L"RunOnUiThread callback begin");
            action();
            TraceMessage(L"RunOnUiThread callback end");
        });

        if (!enqueued)
        {
            if (m_uiDispatcherQueue.HasThreadAccess())
            {
                action();
                return;
            }

            TraceMessage(L"RunOnUiThread: failed to enqueue callback from non-UI thread.");
        }
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
        TraceMessage(L"ShowMainWindow begin");
        if (m_uiDispatcherQueue != nullptr && !m_uiDispatcherQueue.HasThreadAccess())
        {
            TraceMessage(L"ShowMainWindow: not UI thread, re-dispatching");
            RunOnUiThread([this]
            {
                ShowMainWindow();
            });
            return;
        }

        if (m_showWindowInProgress)
        {
            return;
        }

        m_showWindowInProgress = true;
        struct ShowGuard final
        {
            explicit ShowGuard(bool& flagRef) noexcept : flag(flagRef) {}
            ~ShowGuard() { flag = false; }
            bool& flag;
        } showGuard(m_showWindowInProgress);

        EnsureUiInitialized(false);
        TraceMessage(L"ShowMainWindow: UI ensured");

        auto hwnd = MainWindowHandle();
        if (hwnd != nullptr && IsWindowVisible(hwnd) && !IsIconic(hwnd))
        {
            SetForegroundWindow(hwnd);
            m_mainWindowVisible = true;
            RunOnUiThread([this]
            {
                RefreshStateFromService(false);
                StartLicensePolling();
            });
            return;
        }

        try
        {
            m_window.Activate();
        }
        catch (...)
        {
            TraceMessage(L"ShowMainWindow: m_window.Activate failed.");
            return;
        }
        TraceMessage(L"ShowMainWindow: window activated");

        hwnd = MainWindowHandle();
        if (hwnd == nullptr)
        {
            TraceMessage(L"ShowMainWindow: unable to get main window handle.");
            m_mainWindowVisible = false;
            return;
        }

        EnsureMainWindowHooked();

        RECT windowRect{};
        if (GetWindowRect(hwnd, &windowRect))
        {
            int const currentWidth = windowRect.right - windowRect.left;
            int const currentHeight = windowRect.bottom - windowRect.top;
            if (currentWidth < 640 || currentHeight < 420)
            {
                SetWindowPos(
                    hwnd,
                    nullptr,
                    0,
                    0,
                    960,
                    640,
                    SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
            }
        }

        ShowWindow(hwnd, IsIconic(hwnd) ? SW_RESTORE : SW_SHOW);
        ShowWindow(hwnd, SW_SHOW);
        SetForegroundWindow(hwnd);
        m_mainWindowVisible = IsWindowVisible(hwnd) != FALSE;
        TraceMessage(L"ShowMainWindow: native window shown and focused");

        // Refresh service-backed UI state only after window is shown and UI is active.
        RunOnUiThread([this]
        {
            RefreshStateFromService(false);
            StartLicensePolling();
        });
        TraceMessage(L"ShowMainWindow end");
    }

    void App::HideMainWindow()
    {
        auto hwnd = MainWindowHandle();
        if (hwnd == nullptr)
        {
            return;
        }

        ShowWindow(hwnd, SW_HIDE);
        m_mainWindowVisible = false;
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
        if (!m_isExiting)
        {
            hstring confirmError;
            if (!PromptForSecureDesktopConfirmation(
                MainWindowHandle(),
                L"ZIVPO Secure Confirmation",
                L"Confirm application exit.",
                confirmError))
            {
                if (!confirmError.empty())
                {
                    TraceMessage(confirmError.c_str());
                    if (m_authErrorText != nullptr)
                    {
                        m_authErrorText.Text(confirmError);
                    }
                }
                return;
            }
        }

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
