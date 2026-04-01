#pragma once

namespace ZIVPO
{
    class TrayIconManager final
    {
    public:
        TrayIconManager() = default;
        ~TrayIconManager();

        TrayIconManager(TrayIconManager const&) = delete;
        TrayIconManager& operator=(TrayIconManager const&) = delete;

        bool Initialize(std::wstring const& tooltip, std::function<void()> onOpen, std::function<void()> onExit);
        void Shutdown();

    private:
        static constexpr UINT kTrayIconId = 100;
        static constexpr UINT kTrayCallbackMessage = WM_APP + 1;
        static constexpr UINT kMenuOpenId = 2001;
        static constexpr UINT kMenuExitId = 2002;

        static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);
        LRESULT HandleWindowMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam);

        bool RegisterWindowClass();
        bool CreateMessageWindow();
        bool AddIcon();
        void RemoveIcon();
        void CreateContextMenu();
        void ShowContextMenu();

        HWND m_messageWindow{ nullptr };
        UINT m_taskbarCreatedMessage{ 0 };
        NOTIFYICONDATAW m_notifyIconData{};
        std::function<void()> m_onOpen;
        std::function<void()> m_onExit;
        HMENU m_contextMenu{ nullptr };
        bool m_isIconAdded{ false };
    };
}
