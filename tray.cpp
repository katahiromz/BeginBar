#include <windows.h>
#include <commctrl.h>
#include <shlobj.h>
#include <tchar.h>
#include <string>
#include <vector>

#include "fakemenu.h"
#include "BeginBar.h"

static WCHAR s_szButtonTrayClass[] = L"ButtonTray";

static VOID
ButtonTray_ArrangeButton(HWND hwnd, LPSIZE psiz, LPSIZE psizTotal, BOOL bVertical)
{
    HWND hButton;
    WCHAR sz[64];
    SIZE siz;
    HFONT hFont;
    INT i, nButtonCount = (INT)g_buttons.size();

    if (bVertical)
    {
        INT y = 4;
        for (i = 0; i < nButtonCount; i++)
        {
            hButton = g_buttons[i].hButton;
            GetWindowTextW(hButton, sz, 64);
            hFont = (HFONT)SendMessageW(hButton, WM_GETFONT, 0, 0);

            HDC hdc = GetDC(hButton);
            SelectObject(hdc, hFont);
            GetTextExtentPoint32W(hdc, sz, lstrlenW(sz), &siz);
            ReleaseDC(hButton, hdc);
            if (siz.cy < 16)
                siz.cy = 16;
            siz.cx += 22 + 8;
            siz.cy += 16;
            MoveWindow(hButton, 4, y, siz.cx, siz.cy - 10, TRUE);
            y += siz.cy;
        }
    }
    else
    {
        INT x = 4;
        for (i = 0; i < nButtonCount; i++)
        {
            hButton = g_buttons[i].hButton;
            GetWindowTextW(hButton, sz, 64);
            hFont = (HFONT)SendMessageW(hButton, WM_GETFONT, 0, 0);

            HDC hdc = GetDC(hButton);
            SelectObject(hdc, hFont);
            GetTextExtentPoint32W(hdc, sz, lstrlenW(sz), &siz);
            ReleaseDC(hButton, hdc);
            if (siz.cy < 16)
                siz.cy = 16;
            siz.cx += 22 + 8;
            siz.cy += 16;
            MoveWindow(hButton, x, 4, siz.cx, psizTotal->cy - 10, TRUE);
            x += siz.cx + 8;
        }
    }

    MoveWindow(hwnd, 0, 0, psizTotal->cx, psizTotal->cy, TRUE);
    Pager_RecalcSize(g_hPager);

    RECT rc;
    DWORD style = GetWindowLongW(g_hPager, GWL_STYLE);
    style &= ~(PGS_VERT | PGS_HORZ);
    if (g_bVertical)
        style |= PGS_VERT;
    else
        style |= PGS_HORZ;
    SetWindowLongW(g_hPager, GWL_STYLE, style);

    GetWindowRect(hwnd, &rc);

    if (g_bVertical != bVertical)
    {
        RECT rcWorkArea;
        GetWorkAreaNearWindow(hwnd, &rcWorkArea);
        SIZE sizWorkArea;
        sizWorkArea.cx = rcWorkArea.right - rcWorkArea.left;
        sizWorkArea.cy = rcWorkArea.bottom - rcWorkArea.top;

        if (bVertical)
            rc.bottom = rc.top + sizWorkArea.cy * 2 / 3;
        else
            rc.right = rc.left + sizWorkArea.cx * 2 / 3;
    }
    else
    {
        RECT rcClient;
        GetClientRect(g_hMainWnd, &rcClient);
        if (bVertical)
           rc.bottom = rc.top + (rcClient.bottom - rcClient.top);
        else
            rc.right = rc.left + (rcClient.right - rcClient.left);
    }

    style = WS_POPUPWINDOW | WS_CLIPCHILDREN | WS_THICKFRAME;
    DWORD exstyle = WS_EX_TOPMOST | WS_EX_PALETTEWINDOW;
    AdjustWindowRectEx(&rc, style, FALSE, exstyle);

    siz.cx = rc.right - rc.left;
    siz.cy = rc.bottom - rc.top;
    if (bVertical)
    {
        if (g_nWindowCY != CW_USEDEFAULT)
            siz.cy = g_nWindowCY;
    }
    else
    {
        if (g_nWindowCX != CW_USEDEFAULT)
            siz.cx = g_nWindowCX;
    }
    g_bVertical = bVertical;
    SetWindowPos(g_hMainWnd, NULL, 0, 0, siz.cx, siz.cy,
                 SWP_NOACTIVATE | SWP_NOZORDER | SWP_NOMOVE);

    GetClientRect(g_hMainWnd, &rc);
    MoveWindow(g_hPager, rc.left, rc.top,
               rc.right - rc.left, rc.bottom - rc.top, TRUE);
    FitWindowInWorkArea(g_hMainWnd);
}

static VOID
ButtonTray_AddButtons(HWND hwnd, BOOL bVertical)
{
    HWND hButton;
    HDC hdc;
    HGDIOBJ hFontOld;
    SIZE siz, sizTotal = {0, 0};

    LOGFONTW lf;
    GetObjectW(GetStockObject(DEFAULT_GUI_FONT), sizeof(LOGFONTW), &lf);
    lf.lfWeight = FW_BOLD;
    HFONT hFont = CreateFontIndirect(&lf);

    INT i, nButtonCount = (INT)g_buttons.size();
    if (bVertical)
    {
        for (i = 0; i < nButtonCount; i++)
        {
            g_buttons[i].items[0].LoadIcon();
            hButton = CreateWindowW(L"BUTTON", g_buttons[i].items[0].text.c_str(),
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_OWNERDRAW,
                0, 0, 0, 0, hwnd, (HMENU)(i + 1), g_hInstance, NULL);
            g_buttons[i].hButton = hButton;
            SetWindowLongPtr(hButton, GWLP_USERDATA, (LONG_PTR)&g_buttons[i]);
            SendMessageW(hButton, WM_SETFONT, (WPARAM)hFont, TRUE);
            InvalidateRect(hButton, NULL, TRUE);
            hdc = GetDC(hButton);
            hFontOld = SelectObject(hdc, hFont);
            GetTextExtentPoint32(hdc, g_buttons[i].items[0].text.c_str(),
                                 g_buttons[i].items[0].text.size(), &siz);
            if (siz.cy < 16)
                siz.cy = 16;
            siz.cx += 22 + 8;
            siz.cy += 16;
            if (sizTotal.cx < siz.cx)
                sizTotal.cx = siz.cx;
            sizTotal.cy += siz.cy;
            SelectObject(hdc, hFontOld);
            ReleaseDC(hButton, hdc);
        }
        sizTotal.cx += 8;
    }
    else
    {
        for (i = 0; i < nButtonCount; i++)
        {
            g_buttons[i].items[0].LoadIcon();
            hButton = CreateWindowW(L"BUTTON", g_buttons[i].items[0].text.c_str(),
                WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_OWNERDRAW,
                0, 0, 0, 0, hwnd, (HMENU)(i + 1), g_hInstance, NULL);
            g_buttons[i].hButton = hButton;
            SetWindowLongPtr(hButton, GWLP_USERDATA, (LONG_PTR)&g_buttons[i]);
            SendMessageW(hButton, WM_SETFONT, (WPARAM)hFont, TRUE);
            InvalidateRect(hButton, NULL, TRUE);
            hdc = GetDC(hButton);
            hFontOld = SelectObject(hdc, hFont);
            GetTextExtentPoint32(hdc, g_buttons[i].items[0].text.c_str(),
                                 g_buttons[i].items[0].text.size(), &siz);
            if (siz.cy < 16)
                siz.cy = 16;
            siz.cx += 22 + 8;
            siz.cy += 16;
            if (sizTotal.cy < siz.cy)
                sizTotal.cy = siz.cy;
            sizTotal.cx += siz.cx;
            SelectObject(hdc, hFontOld);
            ReleaseDC(hButton, hdc);
        }
    }

    if (nButtonCount == 0)
        return;

    ButtonTray_ArrangeButton(hwnd, &siz, &sizTotal, bVertical);
}

static VOID
ButtonTray_DeleteButtons(HWND hwnd)
{
    INT i, nButtonCount = (INT)g_buttons.size();
    for (i = 0; i < nButtonCount; i++)
    {
        DestroyWindow(g_buttons[i].hButton);
        g_buttons[i].hButton = NULL;
    }
}

POINT s_ptHotSpot;

static VOID
ButtonTray_OnLButton(HWND hwnd, BOOL bDown)
{
    POINT pt;
    RECT rc;
    if (bDown)
    {
        SetCapture(hwnd);

        GetCursorPos(&pt);
        GetWindowRect(g_hMainWnd, &rc);
        pt.x -= rc.left;
        pt.y -= rc.top;
        s_ptHotSpot = pt;
    }
    else
    {
        if (GetCapture() == hwnd)
        {
            ReleaseCapture();
        }
    }
}

static VOID
ButtonTray_OnMouseMove(HWND hwnd, INT x, INT y)
{
    RECT rcWindow, rcDrag;
    SIZE sizWindow;

    if (GetCapture() != hwnd)
    {
        return;
    }

    POINT ptCursor = {x, y};
    MapWindowPoints(hwnd, NULL, &ptCursor, 1);

    GetWindowRect(g_hMainWnd, &rcWindow);
    sizWindow.cx = rcWindow.right - rcWindow.left;
    sizWindow.cy = rcWindow.bottom - rcWindow.top;

    rcDrag.left = ptCursor.x - s_ptHotSpot.x;
    rcDrag.top = ptCursor.y - s_ptHotSpot.y;
    SetWindowPos(g_hMainWnd, NULL,
                 rcDrag.left, rcDrag.top, 0, 0,
                 SWP_NOACTIVATE | SWP_NOSIZE);
}

BOOL MeasureButtonItem(HWND hwnd, UINT idCtl, LPMEASUREITEMSTRUCT pmis)
{
    return FALSE;
}

BOOL DrawButtonItem(HWND hwnd, UINT idCtl, LPDRAWITEMSTRUCT pdis)
{
    if (pdis->CtlType == ODT_BUTTON)
    {
        HWND hButton = pdis->hwndItem;
        WCHAR sz[256];
        GetWindowTextW(hButton, sz, 256);
        RECT rcItem = pdis->rcItem;
        BEGINBUTTON *pButton;
        pButton = (BEGINBUTTON *)GetWindowLongPtrW(hButton, GWLP_USERDATA);
        pButton->items[0].LoadIcon();
        HICON hIcon = pButton->items[0].hIcon;

        if (pdis->itemState & ODS_SELECTED)
        {
            DrawFrameControl(pdis->hDC, &rcItem,
                DFC_BUTTON, DFCS_BUTTONPUSH | DFCS_ADJUSTRECT | DFCS_PUSHED);
            OffsetRect(&rcItem, 1, 1);
        }
        else
        {
            DrawFrameControl(pdis->hDC, &rcItem,
                DFC_BUTTON, DFCS_BUTTONPUSH | DFCS_ADJUSTRECT);
        }

        if (pdis->itemState & ODS_FOCUS)
            DrawFocusRect(pdis->hDC, &rcItem);

        if (hIcon)
        {
            DrawIconEx(pdis->hDC,
                rcItem.left + 3, rcItem.top + (rcItem.bottom - rcItem.top - 16) / 2,
                hIcon,
                16, 16,
                0,
                NULL,
                DI_NORMAL);

            rcItem.left += 22;
        }

        SetBkMode(pdis->hDC, TRANSPARENT);
        if (hIcon)
            DrawTextW(pdis->hDC, sz, lstrlenW(sz), &rcItem,
                DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOCLIP);
        else
            DrawTextW(pdis->hDC, sz, lstrlenW(sz), &rcItem,
                DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOCLIP);
        return TRUE;
    }
    return FALSE;
}

static BOOL
ButtonTray_OnCreate(HWND hwnd)
{
    ButtonTray_AddButtons(hwnd, g_bVertical);
    return TRUE;
}

LRESULT CALLBACK
ButtonTray_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    HFONT hFont;
    std::vector<BEGINBUTTON> *pButtons;

    switch(uMsg)
    {
    case WM_CREATE:
        if (!ButtonTray_OnCreate(hwnd))
            return -1;
        break;

    case MYWM_UPDATEUI:
        pButtons = (std::vector<BEGINBUTTON> *)lParam;
        ButtonTray_DeleteButtons(hwnd);
        g_buttons = *pButtons;
        ButtonTray_AddButtons(hwnd, (BOOL)wParam);
        break;

    case WM_COMMAND:
    case WM_NOTIFY:
    case WM_DROPFILES:
        SendMessageW(GetParent(hwnd), uMsg, wParam, lParam);
        break;

    case WM_LBUTTONDOWN:
        ButtonTray_OnLButton(hwnd, TRUE);
        break;

    case WM_LBUTTONUP:
        ButtonTray_OnLButton(hwnd, FALSE);
        break;

    case WM_MOUSEMOVE:
        ButtonTray_OnMouseMove(hwnd, (SHORT)LOWORD(lParam), (SHORT)HIWORD(lParam));
        break;

    case WM_MEASUREITEM:
        return MeasureButtonItem(hwnd, (UINT)wParam, (LPMEASUREITEMSTRUCT)lParam);

    case WM_DRAWITEM:
        return DrawButtonItem(hwnd, (UINT)wParam, (LPDRAWITEMSTRUCT)lParam);

    case WM_DESTROY:
        hFont = (HFONT)SendMessageW(hwnd, WM_GETFONT, 0, 0);
        DeleteObject(hFont);
        break;

    case WM_CONTEXTMENU:
        SendMessageW(g_hMainWnd, uMsg, wParam, lParam);
        break;

    default:
        return DefWindowProcW(hwnd, uMsg, wParam, lParam);
    }
    return 0;
}

BOOL RegisterButtonTray(HINSTANCE hInstance)
{
    WNDCLASSW wc;
    wc.style = CS_DBLCLKS;
    wc.lpfnWndProc = ButtonTray_WndProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = 0;
    wc.hInstance = hInstance;
    wc.hIcon = NULL;
    wc.hCursor = LoadCursorW(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_3DFACE + 1);
    wc.lpszMenuName = NULL;
    wc.lpszClassName = s_szButtonTrayClass;
    return !!RegisterClassW(&wc);
}

HWND CreateButtonTray(HWND hwnd)
{
    return CreateWindowW(s_szButtonTrayClass, NULL,
                         WS_CHILD | WS_VISIBLE,
                         0, 0, 0, 0,
                         hwnd, NULL,
                         g_hInstance, NULL);
}
