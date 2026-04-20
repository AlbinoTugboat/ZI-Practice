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
        static LRESULT CALLBACK MainWindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
        LRESULT HandleMainWindowMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);

        bool AcquireSingleInstanceMutex();
        static bool ShouldStartHidden(winrt::hstring const& arguments);
        HWND MainWindowHandle() const;
        void EnsureMainWindowHooked();
        void EnsureMainWindowMenu();
        void ShowMainWindow();
        void HideMainWindow();
        void ExitApplication();
        void StopServiceAndExitApplication();
        friend void RequestApplicationExit();

        winrt::Microsoft::UI::Xaml::Window m_window{ nullptr };
        std::unique_ptr<::ZIVPO::TrayIconManager> m_trayIcon;
        HANDLE m_singleInstanceMutex{ nullptr };
        HWND m_mainWindowHwnd{ nullptr };
        WNDPROC m_originalMainWindowProc{ nullptr };
        bool m_isExiting{ false };
    };
}
