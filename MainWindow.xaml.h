#pragma once

#include "MainWindow.g.h"

namespace winrt::ZIVPO::implementation
{
    struct MainWindow : MainWindowT<MainWindow>
    {
        MainWindow();
        void InitializeUi();

        void OnFileExitClick(
            winrt::Windows::Foundation::IInspectable const& sender,
            Microsoft::UI::Xaml::RoutedEventArgs const& args);
    };
}

namespace winrt::ZIVPO::factory_implementation
{
    struct MainWindow : MainWindowT<MainWindow, implementation::MainWindow>
    {
    };
}
