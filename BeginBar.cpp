#include <windows.h>
#include <commctrl.h>
#include <shlobj.h>
#include <tchar.h>
#include <string>
#include <vector>

#include "fakemenu.h"
#include "BeginBar.h"
#include "resource.h"

static const TCHAR s_szWindowClass[] = TEXT("BeginBar");

HINSTANCE g_hInstance;
HWND g_hMainWnd = NULL;
HWND g_hPager;
HWND g_hTray;
INT g_iButton = -1;

WNDPROC g_fnPagerOldWndProc;

static LRESULT CALLBACK
Pager_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch(uMsg)
    {
    case WM_CONTEXTMENU:
    case WM_DROPFILES:
        SendMessageW(g_hMainWnd, uMsg, wParam, lParam);
        break;
    }
    return CallWindowProc(g_fnPagerOldWndProc, hwnd, uMsg, wParam, lParam);
}

static BOOL
OnCreate(HWND hwnd)
{
    DWORD style;

    // create a pager control
    style = WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN |
            WS_TABSTOP | PGS_AUTOSCROLL | PGS_DRAGNDROP;
    if (g_bVertical)
        style |= PGS_VERT;
    else
        style |= PGS_HORZ;
    g_hPager = CreateWindowW(WC_PAGESCROLLERW, NULL, style, 0, 0, 0, 0,
        hwnd, (HMENU)-1, g_hInstance, NULL);
    if (g_hPager == NULL)
        return FALSE;

    g_fnPagerOldWndProc = 
        (WNDPROC)SetWindowLongPtr(g_hPager, GWLP_WNDPROC, (LONG_PTR)Pager_WndProc);
    if (g_fnPagerOldWndProc == NULL)
        return FALSE;

    // create a button tray
    g_hTray = CreateButtonTray(g_hPager);
    if (g_hTray == NULL)
        return FALSE;

    RECT rc;
    GetWindowRect(g_hTray, &rc);
    SetParent(g_hTray, g_hPager);
    Pager_SetChild(g_hPager, g_hTray);
    ShowWindow(g_hTray, SW_SHOWNOACTIVATE);
    UpdateWindow(g_hTray);

    // calculate window size
    RECT rcWorkArea;
    GetWorkAreaNearWindow(hwnd, &rcWorkArea);
    SIZE sizWorkArea;
    sizWorkArea.cx = rcWorkArea.right - rcWorkArea.left;
    sizWorkArea.cy = rcWorkArea.bottom - rcWorkArea.top;

    if (g_bVertical)
    {
        if (g_nWindowCY == CW_USEDEFAULT)
            rc.bottom = rc.top + sizWorkArea.cy * 2 / 3;
    }
    else
    {
        if (g_nWindowCX == CW_USEDEFAULT)
            rc.right = rc.left + sizWorkArea.cx * 2 / 3;
    }

    style = WS_POPUPWINDOW | WS_CLIPCHILDREN | WS_THICKFRAME;
    DWORD exstyle = WS_EX_TOPMOST | WS_EX_PALETTEWINDOW;
    AdjustWindowRectEx(&rc, style, FALSE, exstyle);

    INT cx, cy;
    cx = rc.right - rc.left;
    cy = rc.bottom - rc.top;
    if (g_bVertical)
    {
        if (g_nWindowCY == CW_USEDEFAULT)
            g_nWindowCY = cy;
        else
            cy = g_nWindowCY;
    }
    else
    {
        if (g_nWindowCX == CW_USEDEFAULT)
            g_nWindowCX = cx;
        else
            cx = g_nWindowCX;
    }

    // calculate window position
    if (g_nWindowX == CW_USEDEFAULT)
        g_nWindowX = rcWorkArea.left + (sizWorkArea.cx - cx) * rand() / RAND_MAX;
    if (g_nWindowY == CW_USEDEFAULT)
        g_nWindowY = rcWorkArea.top + (sizWorkArea.cy - cy) * rand() / RAND_MAX;

    // move and resize window
    SetWindowPos(hwnd, NULL,
                 g_nWindowX, g_nWindowY, cx, cy,
                 SWP_NOACTIVATE | SWP_NOZORDER);

    // resize pager
    GetClientRect(hwnd, &rc);
    MoveWindow(g_hPager, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, TRUE);

    DragAcceptFiles(hwnd, TRUE);
    return TRUE;
}

static VOID
OnMove(HWND hwnd)
{
    if (g_hMainWnd != NULL)
    {
        RECT rc;
        GetWindowRect(hwnd, &rc);
        g_nWindowX = rc.left;
        g_nWindowY = rc.top;
    }
}

static VOID
OnSize(HWND hwnd)
{
    if (g_hMainWnd != NULL)
    {
        RECT rc;
        GetWindowRect(hwnd, &rc);
        if (g_bVertical)
            g_nWindowCY = rc.bottom - rc.top;
        else
            g_nWindowCX = rc.right - rc.left;

        GetClientRect(hwnd, &rc);
        MoveWindow(g_hPager, rc.left, rc.top,
                   rc.right - rc.left, rc.bottom - rc.top, TRUE);
    }
}

static VOID
OnFakeMenuSelect(HWND hwnd, UINT uIndex, UINT uFlags, HFAKEMENU hMenu)
{
    if (!(uFlags & MF_FAKEMENU))
        return;

    if (uFlags & MF_POPUP)
    {
        HFAKEMENU hSubMenu = GetSubFakeMenu(hMenu, uIndex);
        if (hSubMenu != NULL && GetFakeMenuItemCount(hSubMenu) == 0)
        {
            AddItemsToMenu(hSubMenu, hSubMenu->pidl);
        }
    }
    else if (hMenu->children[uIndex]->pidl)
    {
        SHELLEXECUTEINFOW sei;
        ZeroMemory(&sei, sizeof(sei));
        sei.cbSize = sizeof(sei);
        sei.fMask = SEE_MASK_IDLIST;
        sei.hwnd = hwnd;
        sei.lpIDList = hMenu->children[uIndex]->pidl;
        sei.nShow = SW_SHOWNORMAL;
        ShellExecuteExW(&sei);
    }
    else
    {
        SendMessageW(hwnd, WM_COMMAND,
                     MAKEWPARAM(hMenu->children[uIndex]->uID, 0xDEAD), 0xDEADFACE);
    }
}

static HRESULT
PidlFromPath(LPCWSTR lpszPath, LPITEMIDLIST& pidl)
{
    LPSHELLFOLDER pDesktopFolder;
    HRESULT hr = SHGetDesktopFolder(&pDesktopFolder);
    if (SUCCEEDED(hr))
    {
        ULONG chEaten;
        OLECHAR olePath[MAX_PATH];
        lstrcpynW(olePath, lpszPath, MAX_PATH);
        hr = pDesktopFolder->ParseDisplayName(
            GetDesktopWindow(),
            NULL,
            olePath,
            &chEaten,
            &pidl,
            NULL);

        pDesktopFolder->Release();
    }
    return hr;
}

INT CreateBeginBar(
    HWND hwnd,
    HFAKEMENU hMenu,
    BEGINBUTTON& button,
    INT iStart, INT iLevel)
{
    HFAKEMENU hSubMenu;
    LPITEMIDLIST pidl1, pidl2;
    std::wstring text;
    HICON hIcon;
    BOOL bGroup = FALSE;
    INT i, nCount = (INT)button.items.size();

    button.LoadIcons();
    for (i = iStart; i < nCount; i++)
    {
        if (iLevel > button.items[i].level)
        {
            if (bGroup)
            {
                hSubMenu = CreatePopupFakeMenu();
                hSubMenu->NoSort = TRUE;
                AppendFakeMenu(hMenu, MF_ENABLED | MF_STRING | MF_POPUP,
                               (UINT_PTR)hSubMenu, text.c_str(), NULL, hIcon);
            }
            return i;
        }
        else if (iLevel < button.items[i].level)
        {
            //assert(bGroup);
            hSubMenu = CreatePopupFakeMenu();
            hSubMenu->NoSort = TRUE;
            i = CreateBeginBar(hwnd, hSubMenu, button, i, button.items[i].level);
            AppendFakeMenu(hMenu, MF_ENABLED | MF_STRING | MF_POPUP,
                           (UINT_PTR)hSubMenu, text.c_str(), NULL, hIcon);
            i--;
            bGroup = FALSE;
        }
        else
        {
            bGroup = FALSE;
            switch(button.items[i].type)
            {
            case ITEMTYPE_SEP:
                AppendFakeMenu(hMenu, MF_ENABLED | MF_SEPARATOR, 0, NULL);
                break;

            case ITEMTYPE_GROUP:
                text = button.items[i].text;
                hIcon = button.items[i].hIcon;
                bGroup = TRUE;
                break;

            case ITEMTYPE_CSIDL1:
                hIcon = button.items[i].hIcon;
                if (button.items[i].expand)
                {
                    hSubMenu = CreatePopupFakeMenu();
                    SHGetSpecialFolderLocation(hwnd, button.items[i].csidl1, &pidl1);
                    if (button.items[i].csidl1 == CSIDL_DRIVES)
                        hSubMenu->NoSort = TRUE;
                    if (button.items[i].csidl1 == CSIDL_RECENT)
                        hSubMenu->Recent = TRUE;
                    AppendFakeMenu(hMenu, MF_ENABLED | MF_STRING | MF_POPUP, (UINT_PTR)hSubMenu,
                                   button.items[i].text.c_str(), pidl1, hIcon);
                }
                else
                {
                    SHGetSpecialFolderLocation(hwnd, button.items[i].csidl1, &pidl1);
                    AppendFakeMenu(hMenu, MF_ENABLED | MF_STRING, 0,
                                   button.items[i].text.c_str(), pidl1, hIcon);
                }
                break;

            case ITEMTYPE_CSIDL2:
                hIcon = button.items[i].hIcon;
                hSubMenu = CreatePopupFakeMenu();
                SHGetSpecialFolderLocation(hwnd, button.items[i].csidl1, &pidl1);
                SHGetSpecialFolderLocation(hwnd, button.items[i].csidl2, &pidl2);
                AddItemsToMenu(hSubMenu, pidl1);
                AddItemsToMenu(hSubMenu, pidl2);
                AppendFakeMenu(hMenu, MF_ENABLED | MF_STRING | MF_POPUP, (UINT_PTR)hSubMenu,
                               button.items[i].text.c_str(), pidl1, hIcon);
                CoTaskMemFree(pidl2);
                break;

            case ITEMTYPE_PATH1:
                hIcon = button.items[i].hIcon;
                if (button.items[i].expand)
                {
                    hSubMenu = CreatePopupFakeMenu();
                    PidlFromPath(button.items[i].path1.c_str(), pidl1);
                    AppendFakeMenu(hMenu, MF_ENABLED | MF_STRING | MF_POPUP, (UINT_PTR)hSubMenu,
                                   button.items[i].text.c_str(), pidl1, hIcon);
                }
                else
                {
                    PidlFromPath(button.items[i].path1.c_str(), pidl1);
                    AppendFakeMenu(hMenu, MF_ENABLED | MF_STRING, 0,
                                   button.items[i].text.c_str(), pidl1, hIcon);
                }
                break;

            case ITEMTYPE_PATH2:
                hIcon = button.items[i].hIcon;
                hSubMenu = CreatePopupFakeMenu();
                PidlFromPath(button.items[i].path1.c_str(), pidl1);
                PidlFromPath(button.items[i].path2.c_str(), pidl2);
                AddItemsToMenu(hSubMenu, pidl1);
                AddItemsToMenu(hSubMenu, pidl2);
                AppendFakeMenu(hMenu, MF_ENABLED | MF_STRING | MF_POPUP, (UINT_PTR)hSubMenu,
                               button.items[i].text.c_str(), pidl1, hIcon);
                CoTaskMemFree(pidl2);
                break;

            case ITEMTYPE_ACTION:
                hIcon = button.items[i].hIcon;
                AppendFakeMenu(hMenu, MF_ENABLED | MF_STRING, button.items[i].action,
                               button.items[i].text.c_str(), NULL, hIcon);
                break;
            }
        }
    }
    if (bGroup)
    {
        hSubMenu = CreatePopupFakeMenu();
        hSubMenu->NoSort = TRUE;
        AppendFakeMenu(hMenu, MF_ENABLED | MF_STRING | MF_POPUP,
                        (UINT_PTR)hSubMenu, text.c_str(), NULL, hIcon);
    }
    return i;
}

static VOID
ShowBeginBar(HWND hwnd, BEGINBUTTON& button)
{
    RECT rcButton;
    GetWindowRect(button.hButton, &rcButton);

    if (button.items.size() == 0)
        return;

    LPITEMIDLIST pidl1, pidl2;
    if (button.items.size() == 1)
    {
        switch(button.items[0].type)
        {
        case ITEMTYPE_CSIDL1:
        case ITEMTYPE_CSIDL2:
            if (!button.items[0].expand)
            {
                SHGetSpecialFolderLocation(hwnd, button.items[0].csidl1, &pidl1);

                SHELLEXECUTEINFOW sei;
                ZeroMemory(&sei, sizeof(sei));
                sei.cbSize = sizeof(sei);
                sei.fMask = SEE_MASK_IDLIST;
                sei.hwnd = hwnd;
                sei.lpIDList = pidl1;
                sei.nShow = SW_SHOWNORMAL;
                ShellExecuteExW(&sei);
                CoTaskMemFree(pidl1);
                return;
            }
            break;

        case ITEMTYPE_PATH1:
        case ITEMTYPE_PATH2:
            if (!button.items[0].expand)
            {
                PidlFromPath(button.items[0].path1.c_str(), pidl1);

                SHELLEXECUTEINFOW sei;
                ZeroMemory(&sei, sizeof(sei));
                sei.cbSize = sizeof(sei);
                sei.fMask = SEE_MASK_IDLIST;
                sei.hwnd = hwnd;
                sei.lpIDList = pidl1;
                sei.nShow = SW_SHOWNORMAL;
                ShellExecuteExW(&sei);
                CoTaskMemFree(pidl1);
                return;
            }
            break;

        case ITEMTYPE_ACTION:
            SendMessageW(hwnd, WM_COMMAND,
                         MAKEWPARAM(button.items[0].action, 0xDEAD), 0xDEADFACE);
            return;
        }
    }

    HFAKEMENU hMenu = CreatePopupFakeMenu();
    HFAKEMENU hChild;

    if (button.items.size() == 1)
    {
        switch (button.items[0].type)
        {
        case ITEMTYPE_CSIDL1:
            SHGetSpecialFolderLocation(hwnd, button.items[0].csidl1, &pidl1);
            AddItemsToMenu(hMenu, pidl1);
            CoTaskMemFree(pidl1);
            break;

        case ITEMTYPE_CSIDL2:
            SHGetSpecialFolderLocation(hwnd, button.items[0].csidl1, &pidl1);
            SHGetSpecialFolderLocation(hwnd, button.items[0].csidl2, &pidl2);
            AddItemsToMenu(hMenu, pidl1);
            AddItemsToMenu(hMenu, pidl2);
            CoTaskMemFree(pidl1);
            CoTaskMemFree(pidl2);
            break;

        case ITEMTYPE_PATH1:
            PidlFromPath(button.items[0].path1.c_str(), pidl1);
            AddItemsToMenu(hMenu, pidl1);
            CoTaskMemFree(pidl1);
            break;

        case ITEMTYPE_PATH2:
            PidlFromPath(button.items[0].path1.c_str(), pidl1);
            PidlFromPath(button.items[0].path2.c_str(), pidl2);
            AddItemsToMenu(hMenu, pidl1);
            AddItemsToMenu(hMenu, pidl2);
            CoTaskMemFree(pidl1);
            CoTaskMemFree(pidl2);
            break;
        }
        hChild = hMenu;
    }
    else
    {
        CreateBeginBar(hwnd, hMenu, button, 0, 0);
        hChild = hMenu->children[0];
        hChild->parent = NULL;
        hMenu->children.clear();
        DestroyFakeMenu(hMenu);
    }
    POINT pt;
    GetCursorPos(&pt);
    TrackPopupFakeMenuNoLock(hChild, TPM_BOTTOMALIGN, pt.x, pt.y, 0, hwnd, &rcButton);
}

BOOL EnableProcessPriviledge(LPCTSTR pszSE_)
{
    BOOL f;
    HANDLE hProcess;
    HANDLE hToken;
    LUID luid;
    TOKEN_PRIVILEGES tp;
    
    f = FALSE;
    hProcess = GetCurrentProcess();
    if (OpenProcessToken(hProcess, TOKEN_ADJUST_PRIVILEGES, &hToken))
    {
        if (LookupPrivilegeValue(NULL, pszSE_, &luid))
        {
            tp.PrivilegeCount = 1;
            tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;
            tp.Privileges[0].Luid = luid;
            f = AdjustTokenPrivileges(hToken, FALSE, &tp, 0, NULL, NULL);
        }
        CloseHandle(hToken);
    }
    return f;
}

static VOID
OnCommand(HWND hwnd, UINT uID, UINT uNotifyCode, HWND hwndCtl)
{
    if (IsWindow(hwndCtl) && uNotifyCode == BN_CLICKED)
    {
        UINT uIndex = uID - 1;
        if (uIndex < g_buttons.size())
            ShowBeginBar(hwnd, g_buttons[uIndex]);
        return;
    }

    if (uNotifyCode == 0xDEAD && (DWORD)hwndCtl == 0xDEADFACE)
    {
        switch(uID)
        {
        case ID_EXITAPP:
            DestroyWindow(hwnd);
            break;

        case ID_EXITWIN:
            EnableProcessPriviledge(SE_SHUTDOWN_NAME);
            ExitWindowsEx(EWX_SHUTDOWN, 0);
            break;

        case ID_LOGOFF:
            EnableProcessPriviledge(SE_SHUTDOWN_NAME);
            ExitWindowsEx(EWX_LOGOFF, 0);
            break;

        case ID_REBOOT:
            EnableProcessPriviledge(SE_SHUTDOWN_NAME);
            ExitWindowsEx(EWX_REBOOT, 0);
            break;
        }
    }
    else
    {
        std::vector<BEGINBUTTON> buttons;
        BEGINBUTTON button;
        ITEM item;
        switch(uID)
        {
        case ID_EXITAPP:
            DestroyWindow(hwnd);
            break;

        case ID_BUTTON_ADD:
            item.type = ITEMTYPE_GROUP;
            item.text = LoadStringDx1(IDS_NEWBUTTON);
            item.csidl1 = 0;
            item.csidl2 = 0;
            item.path1.clear();
            item.path2.clear();
            item.expand = TRUE;
            item.level = 0;
            item.action = 0;
            item.iconpath.clear();
            button.items.push_back(item);
            if (IDOK == DialogBoxParamW(g_hInstance, MAKEINTRESOURCEW(IDD_MENU), hwnd,
                                        Menu_DlgProc, (LPARAM)&button))
            {
                buttons = g_buttons;
                buttons.push_back(button);
                SendMessage(g_hTray, MYWM_UPDATEUI, g_bVertical, (LPARAM)&buttons);
            }
            break;

        case ID_BUTTON_DELETE:
            if (g_iButton != -1)
            {
                if (MessageBoxW(hwnd, LoadStringDx1(IDS_DELETEBUTTON),
                                LoadStringDx2(IDS_QUERYDELETE),
                                MB_ICONINFORMATION | MB_YESNOCANCEL) == IDYES)
                {
                    buttons = g_buttons;
                    std::vector<BEGINBUTTON>::iterator it, end = g_buttons.end();
                    it = buttons.begin();
                    for (INT i = 0; i < g_iButton; i++)
                        it++;
                    buttons.erase(it);
                    SendMessageW(g_hTray, MYWM_UPDATEUI, g_bVertical, (LPARAM)&buttons);
                }
            }
            break;

        case ID_BUTTON_PROPERTY:
            button = g_buttons[g_iButton];
            if (DialogBoxParamW(g_hInstance, MAKEINTRESOURCEW(IDD_MENU), hwnd,
                                Menu_DlgProc, (LPARAM)&button) == IDOK)
            {
                buttons = g_buttons;
                buttons[g_iButton] = button;
                SendMessage(g_hTray, MYWM_UPDATEUI, g_bVertical, (LPARAM)&buttons);
            }
            break;

        case ID_BAR_PROPERTY:
            buttons = g_buttons;
            if (DialogBoxParamW(g_hInstance, MAKEINTRESOURCEW(IDD_BUTTONS), hwnd,
                                Buttons_DlgProc, (LPARAM)&buttons) == IDOK)
            {
                SendMessageW(g_hTray, MYWM_UPDATEUI, g_bVertical, (LPARAM)&buttons);
            }
            break;

        case ID_HORIZONTAL:
            if (g_bVertical)
                SendMessageW(g_hTray, MYWM_UPDATEUI, FALSE, (LPARAM)&g_buttons);
            break;

        case ID_VERTICAL:
            if (!g_bVertical)
                SendMessageW(g_hTray, MYWM_UPDATEUI, TRUE, (LPARAM)&g_buttons);
            break;
        }
    }
}

static VOID
OnContextMenu(HWND hwnd, HWND hwndTarget, INT x, INT y)
{
    HMENU hMenu, hSubMenu;
    size_t i, nCount = g_buttons.size();
    for (i = 0; i < nCount; i++)
    {
        if (hwndTarget == g_buttons[i].hButton)
        {
            break;
        }
    }
    hMenu = LoadMenuW(g_hInstance, MAKEINTRESOURCEW(1));
    if (i < nCount)
    {
        // on button
        hSubMenu = GetSubMenu(hMenu, 0);
        g_iButton = i;
    }
    else
    {
        // on button tray
        hSubMenu = GetSubMenu(hMenu, 1);
        g_iButton = -1;
        if (g_bVertical)
            CheckMenuRadioItem(hSubMenu, 3, 4, 3, MF_BYPOSITION);
        else
            CheckMenuRadioItem(hSubMenu, 3, 4, 4, MF_BYPOSITION);
    }

    SetForegroundWindow(hwnd);
    TrackPopupMenu(hSubMenu, TPM_LEFTALIGN | TPM_RIGHTBUTTON,
        x, y, 0, hwnd, NULL);
    SendMessageW(hwnd, WM_NULL, 0, 0);
    DestroyMenu(hMenu);
}

static VOID
OnDropFiles(HWND hwnd, HDROP hDrop)
{
    WCHAR szFile[MAX_PATH], szTitle[MAX_PATH];
    DragQueryFileW(hDrop, 0, szFile, MAX_PATH);
    DragFinish(hDrop);

    DWORD attrs = GetFileAttributesW(szFile);
    if (attrs == 0xFFFFFFFF)
    {
        MessageBeep(0xFFFFFFFF);
        return;
    }
    GetFileTitleW(szFile, szTitle, MAX_PATH);

    ITEM item;
    item.type = ITEMTYPE_PATH1;
    item.text = szTitle;
    item.path1 = szFile;
    item.csidl1 = item.csidl2 = 0;
    item.expand = !!(attrs & FILE_ATTRIBUTE_DIRECTORY);
    item.level = 0;
    item.action = 0;

    BEGINBUTTON button;
    button.hButton = NULL;
    button.items.push_back(item);

    std::vector<BEGINBUTTON> buttons(g_buttons);
    buttons.push_back(button);
    SendMessageW(g_hTray, MYWM_UPDATEUI, g_bVertical, (LPARAM)&buttons);
}

static INT
OnNcHitTest(HWND hwnd, INT x, INT y)
{
    RECT rc;
    GetWindowRect(hwnd, &rc);

    POINT pt = {x, y};
    if (g_bVertical)
    {
        if (PtInRect(&rc, pt))
        {
            if (rc.top <= y && y <= rc.top + 4)
                return HTTOP;
            if (rc.bottom - 4 <= y && y <= rc.bottom)
                return HTBOTTOM;
            return HTCAPTION;
        }
    }
    else
    {
        if (PtInRect(&rc, pt))
        {
            if (rc.left <= x && x <= rc.left + 4)
                return HTLEFT;
            if (rc.right - 4 <= x && x <= rc.right)
                return HTRIGHT;
            return HTCAPTION;
        }
    }
    return HTNOWHERE;
}

static LRESULT CALLBACK
WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch(uMsg)
    {
    case WM_CREATE:
        if (!OnCreate(hwnd))
            return -1;
        break;

    case WM_MOVE:
        OnMove(hwnd);
        break;

    case WM_SIZE:
        OnSize(hwnd);
        break;

    case WM_FAKEMENUSELECT:
        OnFakeMenuSelect(hwnd, (UINT)LOWORD(wParam), (UINT)HIWORD(wParam), (HFAKEMENU)lParam);
        break;

    case WM_COMMAND:
        OnCommand(hwnd, LOWORD(wParam), HIWORD(wParam), (HWND)lParam);
        break;

    case WM_NCHITTEST:
        return OnNcHitTest(hwnd, (INT)LOWORD(lParam), (INT)HIWORD(lParam));

    case WM_DESTROY:
        PostQuitMessage(0);
        break;

    case WM_CONTEXTMENU:
        OnContextMenu(hwnd, (HWND)wParam,
                      (SHORT)LOWORD(lParam), (SHORT)HIWORD(lParam));
        break;

    case WM_DROPFILES:
        OnDropFiles(hwnd, (HDROP)wParam);
        break;

    default:
        return DefWindowProcW(hwnd, uMsg, wParam, lParam);
    }
    return 0;
}

VOID GetWorkAreaNearWindow(HWND hwnd, LPRECT prcWorkArea)
{
    RECT rc;
    GetWindowRect(hwnd, &rc);

    POINT pt = {(rc.left + rc.right) / 2, (rc.top + rc.bottom) / 2};
    HMONITOR hmon = MonitorFromPoint(pt, MONITOR_DEFAULTTONULL);
    if (hmon == NULL)
        hmon = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);

    MONITORINFO mi;
    mi.cbSize = sizeof(mi);
    if (GetMonitorInfo(hmon, &mi))
        *prcWorkArea = mi.rcWork;
    else
        SystemParametersInfo(SPI_GETWORKAREA, 0, prcWorkArea, 0);
}

VOID FitWindowInWorkArea(HWND hwnd)
{
    RECT rc, rcWorkArea;
    GetWindowRect(hwnd, &rc);

    GetWorkAreaNearWindow(hwnd, &rcWorkArea);

    HWND hwndParent = GetParent(hwnd);
    if (hwndParent)
        MapWindowPoints(NULL, hwndParent, (LPPOINT)&rcWorkArea, 2);

    SIZE siz = {rc.right - rc.left, rc.bottom - rc.top};
    if (rcWorkArea.right < rc.right)
    {
        rc.left = rcWorkArea.right - siz.cx;
        rc.right = rcWorkArea.right;
    }
    if (rc.left < rcWorkArea.left)
    {
        rc.left = rcWorkArea.left;
        rc.right = rcWorkArea.left + siz.cx;
    }
    if (rcWorkArea.bottom < rc.bottom)
    {
        rc.top = rcWorkArea.bottom - siz.cy;
        rc.bottom = rcWorkArea.bottom;
    }
    if (rc.top < rcWorkArea.top)
    {
        rc.top = rcWorkArea.top;
        rc.bottom = rcWorkArea.top + siz.cy;
    }

    SetWindowPos(hwnd, NULL, rc.left, rc.top, siz.cx, siz.cy,
                 SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOZORDER);
}

BOOL IsWin2kPlus(VOID)
{
    OSVERSIONINFOW osver;
    osver.dwOSVersionInfoSize = sizeof(osver);
    if (GetVersionExW(&osver))
    {
        return (osver.dwMajorVersion >= 5 ||
                (osver.dwMajorVersion == 4 &&
                 osver.dwMinorVersion >= 90));
    }
    return FALSE;
}

BOOL OpenFileDialog(
    HWND hwndParent,
    LPWSTR pszFile,
    DWORD dwOFN_flags,
    LPCWSTR pszTitle,
    LPCWSTR pszDefExt,
    LPCWSTR pszFilter,
    LPCWSTR pszInitialDir)
{
    OPENFILENAMEW ofn;
    WCHAR szFile[MAX_PATH];
    lstrcpynW(szFile, pszFile, MAX_PATH);

    ZeroMemory(&ofn, sizeof(ofn));

#ifdef OPENFILENAME_SIZE_VERSION_400W
    if (IsWin2kPlus())
        ofn.lStructSize     = sizeof(OPENFILENAMEW);
    else
        ofn.lStructSize     = OPENFILENAME_SIZE_VERSION_400W;
#else
    ofn.lStructSize         = sizeof(OPENFILENAMEW);
#endif

    ofn.hwndOwner           = hwndParent;
    ofn.hInstance           = g_hInstance;
    ofn.lpstrFilter         = pszFilter;
    ofn.lpstrFile           = szFile;
    ofn.nMaxFile            = MAX_PATH;
    ofn.lpstrInitialDir     = pszInitialDir;
    ofn.lpstrTitle          = pszTitle;
    ofn.Flags               = dwOFN_flags | OFN_ENABLESIZING | OFN_EXPLORER;
    ofn.lpstrDefExt         = pszDefExt;
    if (GetOpenFileNameW(&ofn))
    {
        lstrcpynW(pszFile, szFile, MAX_PATH);
        return TRUE;
    }
    return FALSE;
}

LPWSTR LoadStringDx1(UINT uID)
{
    static WCHAR s_szBuff[MAX_PATH];
    LoadStringW(g_hInstance, uID, s_szBuff, MAX_PATH);
    return s_szBuff;
}

LPWSTR LoadStringDx2(UINT uID)
{
    static WCHAR s_szBuff[MAX_PATH];
    LoadStringW(g_hInstance, uID, s_szBuff, MAX_PATH);
    return s_szBuff;
}

LPWSTR LoadFilterStringDx(UINT uID)
{
    static WCHAR s_szBuff[MAX_PATH];
    LoadStringW(g_hInstance, uID, s_szBuff, MAX_PATH);
    for (INT i = 0; s_szBuff[i] != L'\0'; i++)
    {
        if (s_szBuff[i] == L'|')
            s_szBuff[i] = L'\0';
    }
    return s_szBuff;
}

void trim(std::wstring& str)
{
    static WCHAR s_szSpace[] = L" \t\r\n";

    size_t i, j;
    i = str.find_first_not_of(s_szSpace);
    j = str.find_last_not_of(s_szSpace);
    if (j + 1 >= i)
        str = str.substr(i, j - i + 1);
    else
        str.clear();
}

extern "C"
INT WINAPI WinMain(
    HINSTANCE hInstance,
    HINSTANCE hPrevInstance,
    LPSTR lpCmdLine,
    INT nCmdShow)
{

    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    g_hInstance = hInstance;
    
    HWND hwndFind = FindWindowW(s_szWindowClass, NULL);
    if (hwndFind)
    {
        ShowWindow(hwndFind, SW_RESTORE);
        SetForegroundWindow(hwndFind);
        FitWindowInWorkArea(hwndFind);
        return 3;
    }

    InitCommonControls();

    LoadSettings();

    srand((unsigned)GetTickCount());

    WNDCLASSW wc;
    wc.style            = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc      = WindowProc;
    wc.cbClsExtra       = 0;
    wc.cbWndExtra       = 0;
    wc.hInstance        = hInstance;
    wc.hIcon            = LoadIconW(hInstance, MAKEINTRESOURCEW(1));
    wc.hCursor          = LoadCursorW(NULL, IDC_ARROW);
    wc.hbrBackground    = (HBRUSH)(COLOR_3DFACE + 1);
    wc.lpszMenuName     = NULL;
    wc.lpszClassName    = s_szWindowClass;
    if (!RegisterClassW(&wc) || !RegisterButtonTray(hInstance) ||
        !RegisterFakeMenu(hInstance))
    {
        MessageBoxW(NULL, L"RegisterClass failed", NULL, MB_ICONERROR);
        return 1;
    }

    g_hMainWnd = CreateWindowExW(WS_EX_TOPMOST | WS_EX_PALETTEWINDOW,
                                 s_szWindowClass, LoadStringDx1(IDS_BEGINBAR),
                                 WS_POPUPWINDOW | WS_CLIPCHILDREN | WS_THICKFRAME,
                                 CW_USEDEFAULT, CW_USEDEFAULT, 0, 0,
                                 NULL,
                                 NULL,
                                 hInstance,
                                 NULL);
    if (g_hMainWnd == NULL)
    {
        MessageBoxW(NULL, L"CreateWindowEx failed", NULL, MB_ICONERROR);
        return 2;
    }

    FitWindowInWorkArea(g_hMainWnd);
    ShowWindow(g_hMainWnd, nCmdShow);
    UpdateWindow(g_hMainWnd);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    SaveSettings();

    return (INT)msg.wParam;
}
