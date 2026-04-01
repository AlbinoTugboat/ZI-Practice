#include "pch.h"
#include "TrayIconManager.h"

namespace
{
    constexpr wchar_t kTrayWindowClassName[] = L"ZIVPO.TrayMessageWindow";

    void InvokeSafely(std::function<void()> const& callback) noexcept
    {
        if (!callback)
        {
            return;
        }

        try
        {
            callback();
        }
        catch (...)
        {
            // Swallow exceptions on the tray message pump thread to keep app alive.
        }
    }
}

namespace ZIVPO
{
    TrayIconManager::~TrayIconManager()
    {
        Shutdown();
    }

    bool TrayIconManager::Initialize(std::wstring const& tooltip, std::function<void()> onOpen, std::function<void()> onExit)
    {
        m_onOpen = std::move(onOpen);
        m_onExit = std::move(onExit);

        if (!RegisterWindowClass() || !CreateMessageWindow())
        {
            return false;
        }

        CreateContextMenu();
        m_taskbarCreatedMessage = RegisterWindowMessageW(L"TaskbarCreated");

        m_notifyIconData = {};
        m_notifyIconData.cbSize = sizeof(m_notifyIconData);
        m_notifyIconData.hWnd = m_messageWindow;
        m_notifyIconData.uID = kTrayIconId;
        m_notifyIconData.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
        m_notifyIconData.uCallbackMessage = kTrayCallbackMessage;
        m_notifyIconData.hIcon = LoadIconW(nullptr, IDI_APPLICATION);
        StringCchCopyW(m_notifyIconData.szTip, ARRAYSIZE(m_notifyIconData.szTip), tooltip.c_str());

        return AddIcon();
    }

    void TrayIconManager::Shutdown()
    {
        RemoveIcon();

        if (m_contextMenu != nullptr)
        {
            DestroyMenu(m_contextMenu);
            m_contextMenu = nullptr;
        }

        if (m_messageWindow != nullptr)
        {
            DestroyWindow(m_messageWindow);
            m_messageWindow = nullptr;
        }

        m_onOpen = nullptr;
        m_onExit = nullptr;
    }

    bool TrayIconManager::RegisterWindowClass()
    {
        WNDCLASSEXW windowClass{};
        windowClass.cbSize = sizeof(windowClass);
        windowClass.lpfnWndProc = &TrayIconManager::WindowProc;
        windowClass.hInstance = GetModuleHandleW(nullptr);
        windowClass.lpszClassName = kTrayWindowClassName;

        if (RegisterClassExW(&windowClass) != 0)
        {
            return true;
        }

        return GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
    }

    bool TrayIconManager::CreateMessageWindow()
    {
        m_messageWindow = CreateWindowExW(
            WS_EX_TOOLWINDOW,
            kTrayWindowClassName,
            L"",
            0,
            0,
            0,
            0,
            0,
            nullptr,
            nullptr,
            GetModuleHandleW(nullptr),
            this);

        return m_messageWindow != nullptr;
    }

    bool TrayIconManager::AddIcon()
    {
        if (m_isIconAdded)
        {
            return true;
        }

        if (!Shell_NotifyIconW(NIM_ADD, &m_notifyIconData))
        {
            return false;
        }

        NOTIFYICONDATAW versionData = m_notifyIconData;
        versionData.uVersion = NOTIFYICON_VERSION_4;
        Shell_NotifyIconW(NIM_SETVERSION, &versionData);
        m_isIconAdded = true;
        return true;
    }

    void TrayIconManager::RemoveIcon()
    {
        if (!m_isIconAdded)
        {
            return;
        }

        Shell_NotifyIconW(NIM_DELETE, &m_notifyIconData);
        m_isIconAdded = false;
    }

    void TrayIconManager::CreateContextMenu()
    {
        if (m_contextMenu != nullptr)
        {
            DestroyMenu(m_contextMenu);
        }

        m_contextMenu = CreatePopupMenu();
        AppendMenuW(
            m_contextMenu,
            MF_STRING,
            kMenuOpenId,
            L"\x041E\x0442\x043A\x0440\x044B\x0442\x044C");
        AppendMenuW(m_contextMenu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(
            m_contextMenu,
            MF_STRING,
            kMenuExitId,
            L"\x0412\x044B\x0445\x043E\x0434");
    }

    void TrayIconManager::ShowContextMenu()
    {
        if (m_messageWindow == nullptr || m_contextMenu == nullptr)
        {
            return;
        }

        POINT cursorPosition{};
        GetCursorPos(&cursorPosition);
        SetForegroundWindow(m_messageWindow);
        TrackPopupMenuEx(
            m_contextMenu,
            TPM_LEFTALIGN | TPM_BOTTOMALIGN | TPM_RIGHTBUTTON,
            cursorPosition.x,
            cursorPosition.y,
            m_messageWindow,
            nullptr);
        PostMessageW(m_messageWindow, WM_NULL, 0, 0);
    }

    LRESULT CALLBACK TrayIconManager::WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
    {
        TrayIconManager* self = nullptr;

        if (message == WM_NCCREATE)
        {
            auto createStruct = reinterpret_cast<CREATESTRUCTW*>(lParam);
            self = reinterpret_cast<TrayIconManager*>(createStruct->lpCreateParams);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
        else
        {
            self = reinterpret_cast<TrayIconManager*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
        }

        if (self == nullptr)
        {
            return DefWindowProcW(hwnd, message, wParam, lParam);
        }

        return self->HandleWindowMessage(hwnd, message, wParam, lParam);
    }

    LRESULT TrayIconManager::HandleWindowMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam)
    {
        if (message == m_taskbarCreatedMessage)
        {
            m_isIconAdded = false;
            AddIcon();
            return 0;
        }

        switch (message)
        {
        case kTrayCallbackMessage:
        {
            auto const notification = LOWORD(lParam);
            switch (notification)
            {
            case NIN_SELECT:
            case NIN_KEYSELECT:
            case WM_LBUTTONUP:
            case WM_LBUTTONDBLCLK:
                InvokeSafely(m_onOpen);
                return 0;
            case WM_RBUTTONUP:
            case WM_CONTEXTMENU:
                ShowContextMenu();
                return 0;
            default:
                break;
            }
            break;
        }
        case WM_COMMAND:
            switch (LOWORD(wParam))
            {
            case kMenuOpenId:
                InvokeSafely(m_onOpen);
                return 0;
            case kMenuExitId:
                InvokeSafely(m_onExit);
                return 0;
            default:
                break;
            }
            break;
        case WM_DESTROY:
            RemoveIcon();
            return 0;
        default:
            break;
        }

        return DefWindowProcW(hwnd, message, wParam, lParam);
    }
}
