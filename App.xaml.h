#pragma once

#include "TrayIconManager.h"

namespace winrt::ZIVPO::implementation
{
    void RequestApplicationExit();

    struct App : winrt::Microsoft::UI::Xaml::ApplicationT<App>
    {
        App();
        ~App();

        void OnLaunched(Microsoft::UI::Xaml::LaunchActivatedEventArgs const&);

    private:
        bool AcquireSingleInstanceMutex();
        static bool ShouldStartHidden(winrt::hstring const& arguments);
        void BuildMainUi();
        void EnsureWindowCreated();
        void EnsureUiInitialized(bool showErrorsOnRefresh);
        void ResetUiElementHandles();
        void OnLoginClick(winrt::Windows::Foundation::IInspectable const& sender, Microsoft::UI::Xaml::RoutedEventArgs const& args);
        void OnActivateClick(winrt::Windows::Foundation::IInspectable const& sender, Microsoft::UI::Xaml::RoutedEventArgs const& args);
        void OnLogoutClick(winrt::Windows::Foundation::IInspectable const& sender, Microsoft::UI::Xaml::RoutedEventArgs const& args);
        void RefreshStateFromService(bool showErrors);
        void UpdateUiState();
        void StartLicensePolling();
        void SetAntivirusEnabled(bool enabled);
        void OnMainMenuExitClick(winrt::Windows::Foundation::IInspectable const& sender, Microsoft::UI::Xaml::RoutedEventArgs const& args);
        void RunOnUiThread(std::function<void()> action);
        HWND MainWindowHandle() const;
        void ShowMainWindow();
        void HideMainWindow();
        void ExitApplication();
        void StopServiceAndExitApplication();
        friend void RequestApplicationExit();

        winrt::Microsoft::UI::Xaml::Window m_window{ nullptr };
        winrt::Microsoft::UI::Dispatching::DispatcherQueue m_uiDispatcherQueue{ nullptr };
        std::unique_ptr<::ZIVPO::TrayIconManager> m_trayIcon;
        HANDLE m_singleInstanceMutex{ nullptr };
        winrt::event_token m_windowClosedToken{};
        bool m_windowClosedHandlerAttached{ false };
        bool m_isExiting{ false };
        bool m_uiInitialized{ false };
        bool m_showWindowInProgress{ false };
        bool m_mainWindowVisible{ false };
        bool m_authenticated{ false };
        bool m_hasLicense{ false };
        bool m_licenseBlocked{ false };
        std::wstring m_username;
        std::wstring m_licenseExpiration;
        winrt::Microsoft::UI::Xaml::Controls::TextBlock m_userStatusText{ nullptr };
        winrt::Microsoft::UI::Xaml::Controls::TextBlock m_licenseStatusText{ nullptr };
        winrt::Microsoft::UI::Xaml::Controls::TextBlock m_antivirusStatusText{ nullptr };
        winrt::Microsoft::UI::Xaml::Controls::Button m_antivirusActionButton{ nullptr };
        winrt::Microsoft::UI::Xaml::Controls::Button m_logoutButton{ nullptr };
        winrt::Microsoft::UI::Xaml::Controls::StackPanel m_authPanel{ nullptr };
        winrt::Microsoft::UI::Xaml::Controls::TextBox m_usernameInput{ nullptr };
        winrt::Microsoft::UI::Xaml::Controls::PasswordBox m_passwordInput{ nullptr };
        winrt::Microsoft::UI::Xaml::Controls::TextBlock m_authErrorText{ nullptr };
        winrt::Microsoft::UI::Xaml::Controls::StackPanel m_activationPanel{ nullptr };
        winrt::Microsoft::UI::Xaml::Controls::TextBox m_activationKeyInput{ nullptr };
        winrt::Microsoft::UI::Xaml::Controls::TextBlock m_activationErrorText{ nullptr };
        winrt::Microsoft::UI::Xaml::DispatcherTimer m_licensePollTimer{ nullptr };
    };
}
