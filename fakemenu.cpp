#include <windows.h>
#include <commctrl.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <tchar.h>
#include <process.h>
#include <vector>

#include "fakemenu.h"
#include "BeginBar.h"
#include "resource.h"

const WCHAR g_szBaseBarClsName[] = L"Katayama Hirofumi MZ's BaseBar";
static IContextMenu2 *s_pcm2 = NULL;

HFAKEMENU FakeMenu_GetTop(HFAKEMENU hMenu)
{
    while (hMenu->parent != NULL)
        hMenu = hMenu->parent;
    return hMenu;
}

INT APIENTRY
GetFakeMenuItemCount(HFAKEMENU hMenu)
{
    if (hMenu->uFlags & MF_POPUP)
    {
        return (INT)hMenu->children.size();
    }
    return 0xFFFFFFFF;
}

HFAKEMENU APIENTRY
GetSubFakeMenu(HFAKEMENU hMenu, INT nPos)
{
    if (hMenu->uFlags & MF_POPUP)
    {
        return hMenu->children[nPos];
    }
    return NULL;
}

UINT APIENTRY
GetFakeMenuItemID(HFAKEMENU hMenu, INT nPos)
{
    if (hMenu->uFlags & MF_POPUP)
    {
        if (0 <= nPos && nPos < (INT)hMenu->children.size())
            return hMenu->children[nPos]->uID;
    }
    return 0xFFFFFFFF;
}

BOOL APIENTRY
AppendFakeMenu(HFAKEMENU hMenu,
               UINT uFlags,
               UINT_PTR uIDNewItem,
               LPCWSTR lpNewItem,
               LPITEMIDLIST pidl/* = NULL*/,
               HICON hIcon/* = NULL*/)
{
    HFAKEMENU hSubMenu;
    if (uFlags & MF_POPUP)
    {
        hSubMenu = (HFAKEMENU)uIDNewItem;
        hSubMenu->uID = 0xFFFFFFFF;
    }
    else
    {
        hSubMenu = CreatePopupFakeMenu();
        hSubMenu->uID = uIDNewItem;
    }
    if (lpNewItem)
        hSubMenu->text = lpNewItem;
    hSubMenu->uFlags = uFlags | MF_FAKEMENU;
    hSubMenu->parent = hMenu;
    if (pidl)
        hSubMenu->pidl = pidl;
    if (hIcon)
    {
        if (hSubMenu->hIcon)
            DestroyIcon(hSubMenu->hIcon);
        hSubMenu->hIcon = hIcon;
        hSubMenu->CustomIcon = TRUE;
    }
    hMenu->children.push_back(hSubMenu);
    return TRUE;
}

BOOL APIENTRY
DestroyFakeMenu(HFAKEMENU hMenu)
{
    if (hMenu->hIconThread)
    {
        hMenu->ThreadCancel = TRUE;
        WaitForSingleObject(hMenu->hIconThread, 1500);
        CloseHandle(hMenu->hIconThread);
        hMenu->hIconThread = NULL;
    }

    size_t i, size = hMenu->children.size();
    for (i = 0; i < size; i++)
        DestroyFakeMenu(hMenu->children[i]);
    if (hMenu->hIcon)
        DestroyIcon(hMenu->hIcon);
    if (hMenu->pidl)
        CoTaskMemFree(hMenu->pidl);

    delete hMenu;
    return TRUE;
}

static VOID
FakeMenu_ChoosePos(HWND hwnd,
                   INT x, INT y, INT cx, INT cy,
                   LPPOINT ppt,
                   LPSIZE psiz,
                   INT nReserved,
                   UINT uFlags,
                   LPCRECT prc)
{
    ppt->x = x;
    ppt->y = y;
    HMONITOR hmon = MonitorFromPoint(*ppt, MONITOR_DEFAULTTONULL);
    if (hmon == NULL)
    {
        hmon = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
    }

    MONITORINFO mi;
    mi.cbSize = sizeof(mi);
    GetMonitorInfo(hmon, &mi);

    INT cyScreen = mi.rcMonitor.bottom - mi.rcMonitor.top;
    psiz->cx = cx;
    psiz->cy = cy;

    RECT rc;
    if (nReserved == 0 && prc)
    {
        rc.left = prc->left;
        rc.right = prc->left + cx;
        if (mi.rcMonitor.right < rc.right)
        {
            rc.left = prc->right - cx;
            rc.right = prc->right;
        }
        if (uFlags & TPM_BOTTOMALIGN)
        {
            if (mi.rcMonitor.top <= prc->top - cy)
            {
                rc.top = prc->top - cy;
                rc.bottom = prc->top;
            }
            else
            {
                if ((prc->top - cy) - mi.rcMonitor.top <
                    mi.rcMonitor.bottom - (prc->bottom + cy))
                {
                    rc.top = prc->bottom;
                    rc.bottom = mi.rcMonitor.bottom;
                    if (rc.bottom - rc.top < cy)
                        psiz->cy = rc.bottom - rc.top;
                    else
                        rc.bottom = rc.top + cy;
                }
                else
                {
                    rc.top = mi.rcMonitor.top;
                    rc.bottom = prc->top;
                    if (rc.bottom - rc.top < cy)
                        psiz->cy = rc.bottom - rc.top;
                    else
                        rc.top = rc.bottom - cy;
                }
            }
        }
        else
        {
            if (mi.rcMonitor.bottom >= prc->bottom + cy)
            {
                rc.top = prc->bottom;
                rc.bottom = rc.top + cy;
            }
            else
            {
                if ((prc->top - cy) - mi.rcMonitor.top <
                    mi.rcMonitor.bottom - (prc->bottom + cy))
                {
                    rc.top = prc->bottom;
                    rc.bottom = mi.rcMonitor.bottom;
                    if (rc.bottom - rc.top < cy)
                        psiz->cy = rc.bottom - rc.top;
                    else
                        rc.bottom = rc.top + cy;
                }
                else
                {
                    rc.top = mi.rcMonitor.top;
                    rc.bottom = prc->top;
                    if (rc.bottom - rc.top < cy)
                        psiz->cy = rc.bottom - rc.top;
                    else
                        rc.top = rc.bottom - cy;
                }
            }
        }
        ppt->x = rc.left;
        ppt->y = rc.top;
    }
    else if (prc == NULL)
    {
        if (cy > cyScreen)
            psiz->cy = cy = cyScreen;

        if (uFlags & TPM_RIGHTALIGN)
        {
            ppt->x -= cx;
        }
        if (uFlags & TPM_BOTTOMALIGN)
        {
            ppt->y -= cy;
        }
        if (ppt->x > mi.rcMonitor.right - cx)
        {
            ppt->x = ppt->x - cx;
        }
        if (ppt->x < mi.rcMonitor.left)
        {
            ppt->x = mi.rcMonitor.left;
        }
        if (ppt->y > mi.rcMonitor.bottom - cy)
        {
            ppt->y = mi.rcMonitor.bottom - cy;
        }
        if (ppt->y < mi.rcMonitor.top)
        {
            ppt->y = mi.rcMonitor.top;
        }
    }
    else
    {
        if (cy > cyScreen)
            psiz->cy = cy = cyScreen;

        rc.left = prc->right;
        rc.top = prc->top;
        rc.right = prc->right + cx;
        rc.bottom = prc->top + cy;
        if (rc.right > mi.rcMonitor.right)
        {
            rc.left = prc->left - cx;
            rc.right = prc->left;
        }
        if (rc.left < mi.rcMonitor.left)
        {
            rc.left = prc->right;
            rc.right = prc->right + cx;
        }
        if (rc.bottom > mi.rcMonitor.bottom)
        {
            rc.top = mi.rcMonitor.bottom - cy;
            rc.bottom = mi.rcMonitor.bottom;
        }
        if (rc.top < mi.rcMonitor.top)
        {
            rc.top = mi.rcMonitor.top;
            rc.bottom = mi.rcMonitor.top + cy;
        }
        ppt->x = rc.left;
        ppt->y = rc.top;
    }
}

INT CalcToolBarSize(HWND hwndToolBar, LPSIZE psizToolBar)
{
    RECT rcItem;
    SIZE sizItem;
#if 1
    SendMessageW(hwndToolBar, TB_GETMAXSIZE, 0, (LPARAM)psizToolBar);
    SendMessageW(hwndToolBar, TB_GETITEMRECT, 0, (LPARAM)&rcItem);
    sizItem.cy = rcItem.bottom - rcItem.top;
#else
    psizToolBar->cx = psizToolBar->cy = 0;
    INT nCount = (INT)SendMessageW(hwndToolBar, TB_BUTTONCOUNT, 0, 0);
    for (INT i = 0; i < nCount; i++)
    {
        SendMessageW(hwndToolBar, TB_GETITEMRECT, (WPARAM)i, (LPARAM)&rcItem);
        sizItem.cx = rcItem.right - rcItem.left;
        sizItem.cy = rcItem.bottom - rcItem.top;
        if (psizToolBar->cx < sizItem.cx)
            psizToolBar->cx = sizItem.cx;
        psizToolBar->cy += sizItem.cy;
    }
#endif
    return sizItem.cy;
}

BOOL FakeMenu_OnCreate(HWND hwnd, HFAKEMENU hMenu)
{
    DWORD style;

    SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)hMenu);

    // create a pager control
    style = PGS_VERT | WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN |
            WS_TABSTOP | PGS_AUTOSCROLL | PGS_DRAGNDROP;
    hMenu->hPager = CreateWindowW(WC_PAGESCROLLERW, NULL, style, 0, 0, 0, 0,
        hwnd, (HMENU)-1, g_hInstance, NULL);
    if (hMenu->hPager == NULL)
    {
        return FALSE;
    }

    // create a toolbar
    style = WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN |
            TBSTYLE_REGISTERDROP | TBSTYLE_CUSTOMERASE | TBSTYLE_LIST |
            TBSTYLE_FLAT | TBSTYLE_BUTTON | CCS_NODIVIDER | CCS_VERT |
            CCS_NOPARENTALIGN | CCS_NORESIZE | CCS_TOP;
    hMenu->hToolBar = CreateWindowW(TOOLBARCLASSNAMEW, NULL, style,
        0, 0, 0, 0, hMenu->hPager, (HMENU)-1, g_hInstance, NULL);
    if (hMenu->hToolBar == NULL)
    {
        DestroyWindow(hMenu->hPager);
        return FALSE;
    }
    SendMessageW(hMenu->hToolBar, TB_BUTTONSTRUCTSIZE, (WPARAM)sizeof(TBBUTTON), 0);
    SendMessageW(hMenu->hToolBar, TB_SETBITMAPSIZE, 0,
                 MAKELPARAM(hMenu->nIconSize, hMenu->nIconSize));

    // add items to the toolbar
    std::wstring text;
    INT i, nCount = (INT)hMenu->children.size();
    if (nCount == 0)
        return FALSE;

    DWORD cb = sizeof(TBBUTTON) * nCount;
    HANDLE hHeap = GetProcessHeap();
    TBBUTTON *ptbb = (TBBUTTON *)HeapAlloc(hHeap, HEAP_ZERO_MEMORY, cb);
    if (ptbb)
    {
        for (i = 0; i < nCount; i++)
        {
            if (hMenu->children[i]->uFlags & MF_SEPARATOR)
            {
                ptbb[i].fsStyle = TBSTYLE_SEP;
                ptbb[i].iString = -1;
                ptbb[i].idCommand = 0;
            }
            else
            {
                text = hMenu->children[i]->text;
                text += L"    ";
                ptbb[i].iString = (INT)SendMessageW(hMenu->hToolBar, TB_ADDSTRINGW,
                                                    0, (LPARAM)text.c_str());
                ptbb[i].fsStyle = TBSTYLE_BUTTON | TBSTYLE_NOPREFIX;
                ptbb[i].idCommand = i;
            }

            if (hMenu->children[i]->uFlags & MF_GRAYED)
                ptbb[i].fsState = 0;
            else
                ptbb[i].fsState = TBSTATE_ENABLED;
        }
        SendMessageW(hMenu->hToolBar, TB_ADDBUTTONS, nCount, (LPARAM)ptbb);
        RECT rc;
        SendMessageW(hMenu->hToolBar, TB_SETROWS, MAKEWPARAM(nCount, FALSE), (LPARAM)&rc);
        HeapFree(hHeap, 0, ptbb);
    }
    else
    {
        return FALSE;
    }

    // calculate the size of the toolbar
    SIZE siz;
    hMenu->cyItem = CalcToolBarSize(hMenu->hToolBar, &siz);
    MoveWindow(hMenu->hToolBar, 0, 0, siz.cx, siz.cy, TRUE);
    hMenu->sizToolBar = siz;

    Pager_SetChild(hMenu->hPager, hMenu->hToolBar);

    hMenu->hBaseBar = hwnd;
    return TRUE;
}

#if 0
    LPITEMIDLIST WINAPI ILGetNext(LPCITEMIDLIST pidl)
    {
        WORD len;

        if (pidl)
        {
            len =  pidl->mkid.cb;
            if (len)
            {
                pidl = (LPCITEMIDLIST) (((const BYTE*)pidl) + len);
                return (LPITEMIDLIST)pidl;
            }
        }
        return NULL;
    }

    LPITEMIDLIST WINAPI ILFindLastID(LPCITEMIDLIST pidl)
    {
        LPCITEMIDLIST   pidlLast = pidl;

        if (pidl == NULL)
            return NULL;

        while (pidl->mkid.cb)
        {
            pidlLast = pidl;
            pidl = ILGetNext(pidl);
        }
        return (LPITEMIDLIST)pidlLast;
    }
#endif

static VOID
FakeMenu_OnContextMenu(HWND hwnd, INT x, INT y, HFAKEMENU hMenu)
{
    RECT rcItem;
    POINT pt;
    UINT uIndex, uSelectedOld;
    if (x == -1 && y == -1)
    {
        uIndex = hMenu->uSelected;
        if (uIndex == 0xFFFFFFFF)
            return;
        SendMessageW(hMenu->hToolBar, TB_GETITEMRECT, uIndex, (LPARAM)&rcItem);
        x = (rcItem.left + rcItem.right) / 2;
        y = (rcItem.top + rcItem.bottom) / 2;
        pt.x = x;
        pt.y = y;
        ClientToScreen(hMenu->hToolBar, &pt);
    }
    else
    {
        pt.x = x;
        pt.y = y;
        ScreenToClient(hMenu->hToolBar, &pt);
        uIndex = SendMessageW(hMenu->hToolBar, TB_HITTEST, 0, (LPARAM)&pt);
        uSelectedOld = hMenu->uSelected;
        hMenu->uSelected = uIndex;
        SendMessageW(hMenu->hToolBar, TB_GETITEMRECT, uSelectedOld, (LPARAM)&rcItem);
        InvalidateRect(hMenu->hToolBar, &rcItem, TRUE);
        SendMessageW(hMenu->hToolBar, TB_GETITEMRECT, uIndex, (LPARAM)&rcItem);
        InvalidateRect(hMenu->hToolBar, &rcItem, TRUE);
        pt.x = x;
        pt.y = y;
    }

    HFAKEMENU selected = hMenu->children[uIndex];
    if (selected->pidl == NULL)
        return;

    IShellFolder *psfFolder;
    if (FAILED(SHGetDesktopFolder(&psfFolder)))
        return;

    SHBindToParent(selected->pidl, IID_IShellFolder, reinterpret_cast<VOID **>(&psfFolder), NULL);
    LPITEMIDLIST pidlChild = ILFindLastID(selected->pidl);
    IContextMenu *pcm = NULL;
    HRESULT hr = psfFolder->GetUIObjectOf(hwnd, 1, (LPCITEMIDLIST *)&pidlChild,
                                          IID_IContextMenu, NULL, (VOID **)&pcm);
    if (SUCCEEDED(hr))
    {
        HMENU hPopup = CreatePopupMenu();
        hr = pcm->QueryInterface(IID_IContextMenu2, (VOID **)&s_pcm2);
        if (SUCCEEDED(hr) && s_pcm2)
        {
            hr = s_pcm2->QueryContextMenu(hPopup, 0, 1, 0x7FFF,
                                          CMF_NORMAL | CMF_EXPLORE);
        }
        else
        {
            hr = pcm->QueryContextMenu(hPopup, 0, 1, 0x7FFF,
                                       CMF_NORMAL | CMF_EXPLORE);
        }
        if (SUCCEEDED(hr))
        {
            UINT idCmd = TrackPopupMenu(hPopup,
                                        TPM_LEFTALIGN | TPM_RETURNCMD | TPM_RIGHTBUTTON,
                                        pt.x, pt.y, 0, hwnd, NULL);
            if (idCmd != 0)
            {
                CMINVOKECOMMANDINFO cmi;
                ZeroMemory(&cmi, sizeof(cmi));
                cmi.cbSize = sizeof(cmi);
                cmi.fMask = 0;
                cmi.hwnd = GetDesktopWindow();
                cmi.lpVerb = MAKEINTRESOURCEA(idCmd - 1);
                cmi.lpParameters = NULL;
                cmi.lpDirectory = NULL;
                cmi.nShow = SW_SHOWNORMAL;
                hr = pcm->InvokeCommand(&cmi);
                if (SUCCEEDED(hr))
                {
                    ;
                }
                SendMessageW(FakeMenu_GetTop(hMenu)->hBaseBar, WM_CLOSEBASEBAR, 1, 0);
            }
            if (s_pcm2)
            {
                s_pcm2->Release();
                s_pcm2 = NULL;
            }
        }
        DestroyMenu(hPopup);
        if (pcm)
            pcm->Release();
    }
    psfFolder->Release();
}

static LRESULT FakeMenu_OnNotify(HWND hwnd, WPARAM wParam, LPARAM lParam)
{
    HFAKEMENU hMenu = (HFAKEMENU)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
    LPNMHDR pnmhdr = (LPNMHDR)lParam;
    LPNMPGCALCSIZE pCalcSize;
    LPNMPGSCROLL pScroll;
    SIZE siz;
    LPNMTBCUSTOMDRAW pCustomDraw;
    switch(pnmhdr->code)
    {
    case PGN_CALCSIZE:
        SendMessageW(hMenu->hToolBar, TB_GETMAXSIZE, 0, (LPARAM)&siz);
        pCalcSize = (LPNMPGCALCSIZE)lParam;
        switch (pCalcSize->dwFlag)
        {
        case PGF_CALCWIDTH:
            pCalcSize->iWidth = siz.cx;
            break;

        case PGF_CALCHEIGHT:
            pCalcSize->iHeight = siz.cy;
            break;
        }
        break;

    case PGN_SCROLL:
        pScroll = (LPNMPGSCROLL)lParam;
        pScroll->iScroll = hMenu->cyItem;
        if (hMenu->popupped)
        {
            SendMessageW(hMenu->popupped->hBaseBar, WM_CLOSEBASEBAR, 2, 0);
            hMenu->popupped = NULL;
            hMenu->uPopupped = 0xFFFFFFFF;
        }
        break;

    case NM_RCLICK:
        {
            POINT pt;
            GetCursorPos(&pt);
            FakeMenu_OnContextMenu(hwnd, pt.x, pt.y, hMenu);
        }
        break;

    case NM_CUSTOMDRAW:
        pCustomDraw = (LPNMTBCUSTOMDRAW)lParam;
        if (pCustomDraw->nmcd.dwDrawStage == CDDS_PREERASE)
        {
            HDC hdc = pCustomDraw->nmcd.hdc;
            UINT i, nCount = (UINT)SendMessageW(hMenu->hToolBar, TB_BUTTONCOUNT, 0, 0);
            RECT rcItem;
            for (i = 0; i < nCount; i++)
            {
                SendMessageW(hMenu->hToolBar, TB_GETITEMRECT, i, (LPARAM)&rcItem);
                if (hMenu->uPopupped == i || hMenu->uSelected == i ||
                    (pCustomDraw->nmcd.uItemState & (CDIS_SELECTED | CDIS_HOT)))
                {
                    FillRect(hdc, &rcItem, GetSysColorBrush(COLOR_HIGHLIGHT));
                }
                else
                {
                    FillRect(hdc, &rcItem, GetSysColorBrush(COLOR_MENU));
                }
            }
            // we erased it so skip default
            return CDRF_SKIPDEFAULT;
        }
        else if (pCustomDraw->nmcd.dwDrawStage == CDDS_PREPAINT)
        {
            return CDRF_NOTIFYITEMDRAW;
        }
        else if (pCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)
        {
            UINT i = (UINT)pCustomDraw->nmcd.dwItemSpec;
            if (i == hMenu->uSelected || i == hMenu->uPopupped ||
                (pCustomDraw->nmcd.uItemState & (CDIS_SELECTED | CDIS_HOT)))
            {
                pCustomDraw->clrText = GetSysColor(COLOR_HIGHLIGHTTEXT);
                pCustomDraw->clrBtnFace = GetSysColor(COLOR_HIGHLIGHT);
            }
            else
            {
                pCustomDraw->clrText = GetSysColor(COLOR_MENUTEXT);
                pCustomDraw->clrBtnFace = GetSysColor(COLOR_MENU);
            }

            pCustomDraw->clrTextHighlight = GetSysColor(COLOR_HIGHLIGHTTEXT);
            pCustomDraw->clrHighlightHotTrack = GetSysColor(COLOR_HIGHLIGHT);
            pCustomDraw->nStringBkMode = TRANSPARENT;
            pCustomDraw->nHLStringBkMode = TRANSPARENT;

            return CDRF_NOTIFYPOSTPAINT | TBCDRF_HILITEHOTTRACK | TBCDRF_NOEDGES |
                   TBCDRF_NOOFFSET;
        }
        else if (pCustomDraw->nmcd.dwDrawStage == CDDS_POSTPAINT)
        {
            return CDRF_NOTIFYITEMDRAW;
        }
        else if (pCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPOSTPAINT)
        {
            UINT i = (UINT)pCustomDraw->nmcd.dwItemSpec;
            HDC hdc = pCustomDraw->nmcd.hdc;
            RECT rcItem;
            LOGFONTW lf;
            if (i < hMenu->children.size())
            {
                SendMessageW(hMenu->hToolBar, TB_GETITEMRECT, i, (LPARAM)&rcItem);

                // is the item popup?
                if (hMenu->children[i]->uFlags & MF_POPUP)
                {
                    // create a marlett font
                    ZeroMemory(&lf, sizeof(lf));
                    lf.lfHeight = hMenu->cyItem * 2 / 3;
                    lf.lfCharSet = DEFAULT_CHARSET;
                    lstrcpyW(lf.lfFaceName, L"Marlett");
                    HFONT hFont = CreateFontIndirectW(&lf);
                    if (hFont)
                    {
                        HGDIOBJ hFontOld = SelectObject(hdc, hFont);
                        SetBkMode(hdc, TRANSPARENT);

                        // draw a triangle
                        if (hMenu->uPopupped == i || hMenu->uSelected == i ||
                            (pCustomDraw->nmcd.uItemState & (CDIS_SELECTED | CDIS_HOT)))
                        {
                            SetTextColor(hdc, GetSysColor(COLOR_HIGHLIGHTTEXT));
                        }
                        else
                        {
                            SetTextColor(hdc, GetSysColor(COLOR_MENUTEXT));
                        }
                        TextOutW(hdc,
                                 rcItem.right - hMenu->cyItem * 2 / 3,
                                 rcItem.top + hMenu->cyItem / 6,
                                 L"8", 1);

                        // deselect and delete the font
                        SelectObject(hdc, hFontOld);
                        DeleteObject(hFont);
                    }
                }

                // draw an icon
                if (hMenu->children[i]->hIcon)
                {
                    INT padding = (rcItem.bottom - rcItem.top - hMenu->nIconSize) / 2;
                    DrawIconEx(hdc,
                               rcItem.left + padding,
                               rcItem.top + padding,
                               hMenu->children[i]->hIcon,
                               hMenu->nIconSize, hMenu->nIconSize, 0,
                               NULL, DI_NORMAL);
                }
            }
            return CDRF_NOTIFYPOSTPAINT;
        }
    }
    return 0;
}

VOID FakeMenu_OnSize(HWND hwnd, HFAKEMENU hMenu)
{
    RECT rc;
    GetClientRect(hwnd, &rc);
    MoveWindow(hMenu->hPager, rc.left, rc.top,
        rc.right - rc.left, rc.bottom - rc.top, TRUE);
}

static VOID
FakeMenu_OnCommand(HWND hwnd, WPARAM wParam, LPARAM lParam, HFAKEMENU hMenu)
{
    UINT uIndex = (UINT)LOWORD(wParam);
    HFAKEMENU hClicked = hMenu->children[uIndex];

    HFAKEMENU hTop = FakeMenu_GetTop(hMenu);
    SendMessageW(hTop->hwndOwner, WM_FAKEMENUSELECT,
                 MAKEWPARAM(uIndex, hClicked->uFlags), (LPARAM)hMenu);

    if (hClicked->uFlags & MF_POPUP)
    {
        if (hMenu->popupped)
        {
            SendMessageW(hMenu->popupped->hBaseBar, WM_CLOSEBASEBAR, 3, 0);
            hMenu->popupped = NULL;
            hMenu->uPopupped = 0xFFFFFFFF;
        }

        if (hMenu->uPopupped == uIndex)
        {
            hMenu->uPopupped = 0xFFFFFFFF;
            return;
        }

        RECT rcItem;
        SendMessageW(hMenu->hToolBar, TB_GETITEMRECT, (WPARAM)uIndex, (LPARAM)&rcItem);
        MapWindowPoints(hMenu->hToolBar, NULL, (LPPOINT)&rcItem, 2);
        TrackPopupFakeMenuNoLock(hClicked,
                                 0,
                                 rcItem.right, rcItem.top,
                                 1,
                                 hMenu->hBaseBar,
                                 &rcItem);
        hMenu->popupped = hClicked;
        hMenu->uPopupped = uIndex;
    }
    else
    {
        SendMessageW(hTop->hBaseBar, WM_CLOSEBASEBAR, 4, 0);
    }
}

static VOID
FakeMenu_OnClose(HWND hwnd, HFAKEMENU hMenu, WPARAM wParam)
{
    SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);

    HFAKEMENU popupped = hMenu->popupped;
    if (popupped)
    {
        FakeMenu_OnClose(popupped->hBaseBar, popupped, 12);
        hMenu->popupped = NULL;
        hMenu->uPopupped = 0xFFFFFFFF;
    }

    if (hMenu->parent)
    {
        HFAKEMENU parent = hMenu->parent;
        UINT i = parent->uPopupped;
        parent->uPopupped = 0xFFFFFFFF;
        parent->popupped = NULL;
        RECT rcItem;
        SendMessageW(parent->hToolBar, TB_GETITEMRECT, i, (LPARAM)&rcItem);
        InvalidateRect(parent->hToolBar, &rcItem, TRUE);
    }

    ShowWindow(hMenu->hBaseBar, SW_HIDE);

    INT i, nCount = (INT)hMenu->children.size();
    for (i = 0; i < nCount; i++)
    {
        if (!hMenu->children[i]->CustomIcon)
        {
            DestroyIcon(hMenu->children[i]->hIcon);
            hMenu->children[i]->hIcon = NULL;
        }
    }

    DestroyWindow(hMenu->hToolBar);
    hMenu->hToolBar = NULL;
    DestroyWindow(hMenu->hPager);
    hMenu->hPager = NULL;
    DestroyWindow(hMenu->hBaseBar);
    hMenu->hBaseBar = NULL;
    if (hMenu->parent == NULL)
    {
        SetActiveWindow(hMenu->hwndOwner);
    }
}

static VOID
FakeMenu_OnActivate(HWND hwnd, WPARAM wParam, LPARAM lParam, HFAKEMENU hMenu)
{
    if (LOWORD(wParam) == WA_INACTIVE)
    {
        HWND hwndNext = (HWND)lParam;
        WCHAR szClass[MAX_PATH];
        GetClassNameW(hwndNext, szClass, MAX_PATH);
        if (lstrcmpiW(szClass, g_szBaseBarClsName) == 0)
        {
            HFAKEMENU hMenu2, hTopMenu1, hTopMenu2;
            hMenu2 = (HFAKEMENU)GetWindowLongPtrW(hwndNext, GWLP_USERDATA);
            hTopMenu1 = FakeMenu_GetTop(hMenu);
            hTopMenu2 = FakeMenu_GetTop(hMenu2);
            if (hTopMenu1 != hTopMenu2)
            {
                SendMessageW(hTopMenu1->hBaseBar, WM_CLOSEBASEBAR, 5, 0);
            }
        }
        else
        {
            hMenu = FakeMenu_GetTop(hMenu);
            SendMessageW(hMenu->hBaseBar, WM_CLOSEBASEBAR, 6, 0);
        }
    }
}

static VOID
FixPagerPos(HFAKEMENU hMenu, LPCRECT prcItem)
{
    RECT rcPager;
    GetClientRect(hMenu->hPager, &rcPager);
    INT yPos = Pager_GetPos(hMenu->hPager);
    if (prcItem->top < yPos)
    {
        Pager_SetPos(hMenu->hPager, prcItem->top);
    }
    else if (yPos + (rcPager.bottom - rcPager.top) < prcItem->bottom)
    {
        Pager_SetPos(hMenu->hPager,
            prcItem->bottom - (rcPager.bottom - rcPager.top));
    }
}

static VOID
FakeMenu_OnKeyDown(HWND hwnd, INT vk, HFAKEMENU hMenu)
{
    UINT uSelectedOld = hMenu->uSelected;
    RECT rcItem, rcToolBar, rcPager;
    HFAKEMENU hTop, hClicked;

    GetWindowRect(hMenu->hToolBar, &rcToolBar);
    GetClientRect(hMenu->hPager, &rcPager);
    switch(vk)
    {
    case VK_UP:
        if (hMenu->uSelected == 0xFFFFFFFF)
        {
            hMenu->uSelected = (UINT)(hMenu->children.size() - 1);
        }
        else
        {
            hMenu->uSelected--;
            if (hMenu->uSelected == 0xFFFFFFFF)
            {
                hMenu->uSelected = (UINT)(hMenu->children.size() - 1);
            }
        }
        SendMessageW(hMenu->hToolBar, TB_GETITEMRECT, uSelectedOld, (LPARAM)&rcItem);
        InvalidateRect(hMenu->hToolBar, &rcItem, TRUE);
        SendMessageW(hMenu->hToolBar, TB_GETITEMRECT, hMenu->uSelected, (LPARAM)&rcItem);
        InvalidateRect(hMenu->hToolBar, &rcItem, TRUE);
        FixPagerPos(hMenu, &rcItem);
        break;

    case VK_DOWN:
        if (hMenu->uSelected == 0xFFFFFFFF)
        {
            hMenu->uSelected = 0;
        }
        else
        {
            hMenu->uSelected++;
            if (hMenu->uSelected == (UINT)hMenu->children.size())
            {
                hMenu->uSelected = 0;
            }
        }
        SendMessageW(hMenu->hToolBar, TB_GETITEMRECT, uSelectedOld, (LPARAM)&rcItem);
        InvalidateRect(hMenu->hToolBar, &rcItem, TRUE);
        SendMessageW(hMenu->hToolBar, TB_GETITEMRECT, hMenu->uSelected, (LPARAM)&rcItem);
        InvalidateRect(hMenu->hToolBar, &rcItem, TRUE);
        FixPagerPos(hMenu, &rcItem);
        break;

    case VK_ESCAPE:
        SendMessageW(hMenu->hBaseBar, WM_CLOSEBASEBAR, 7, 0);
        break;

    case VK_LEFT:
        if (hMenu->parent)
        {
            SendMessageW(hMenu->hBaseBar, WM_CLOSEBASEBAR, 8, 0);
            SetActiveWindow(hMenu->parent->hBaseBar);
        }
        break;

    case VK_RIGHT:
        if (uSelectedOld != 0xFFFFFFFF && hMenu->children[uSelectedOld]->uFlags & MF_POPUP)
        {
            hClicked = hMenu->children[uSelectedOld];

            hTop = FakeMenu_GetTop(hMenu);
            SendMessageW(hTop->hwndOwner, WM_FAKEMENUSELECT,
                         MAKEWPARAM(uSelectedOld, hClicked->uFlags), (LPARAM)hMenu);
            if (hMenu->popupped)
            {
                SendMessageW(hMenu->popupped->hBaseBar, WM_CLOSEBASEBAR, 9, 0);
                hMenu->popupped = NULL;
                hMenu->uPopupped = 0xFFFFFFFF;
            }

            RECT rcItem;
            SendMessageW(hMenu->hToolBar, TB_GETITEMRECT, (WPARAM)uSelectedOld, (LPARAM)&rcItem);
            MapWindowPoints(hMenu->hToolBar, NULL, (LPPOINT)&rcItem, 2);
            TrackPopupFakeMenuNoLock(hClicked,
                                     0,
                                     rcItem.right, rcItem.top,
                                     2,
                                     hMenu->hBaseBar,
                                     &rcItem);
            hMenu->popupped = hClicked;
            hMenu->uPopupped = hMenu->uSelected;
        }
        break;

    case VK_RETURN:
        hTop = FakeMenu_GetTop(hMenu);
        hClicked = hMenu->children[uSelectedOld];
        SendMessageW(hTop->hwndOwner, WM_FAKEMENUSELECT,
                     MAKEWPARAM(uSelectedOld, hClicked->uFlags), (LPARAM)hMenu);

        if (uSelectedOld != 0xFFFFFFFF && hMenu->children[uSelectedOld]->uFlags & MF_POPUP)
        {
            if (hMenu->popupped)
            {
                SendMessageW(hMenu->popupped->hBaseBar, WM_CLOSEBASEBAR, 10, 0);
                hMenu->popupped = NULL;
                hMenu->uPopupped = 0xFFFFFFFF;
            }

            RECT rcItem;
            SendMessageW(hMenu->hToolBar, TB_GETITEMRECT, (WPARAM)uSelectedOld, (LPARAM)&rcItem);
            MapWindowPoints(hMenu->hToolBar, NULL, (LPPOINT)&rcItem, 2);
            TrackPopupFakeMenuNoLock(hClicked,
                                     0,
                                     rcItem.right, rcItem.top,
                                     2,
                                     hMenu->hBaseBar,
                                     &rcItem);
            hMenu->popupped = hClicked;
            hMenu->uPopupped = hMenu->uSelected;
        }
        else
        {
            SendMessageW(hTop->hBaseBar, WM_CLOSEBASEBAR, 11, 0);
        }
        break;

    case VK_APPS:
        FakeMenu_OnContextMenu(hwnd, -1, -1, hMenu);
        break;
    }
}

static LRESULT CALLBACK
FakeMenu_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    LPCREATESTRUCT pcs;
    HFAKEMENU hMenu = (HFAKEMENU)GetWindowLongPtrW(hwnd, GWLP_USERDATA);

    switch(uMsg)
    {
    case WM_CREATE:
        pcs = (LPCREATESTRUCT)lParam;
        hMenu = (HFAKEMENU)pcs->lpCreateParams;
        if (!FakeMenu_OnCreate(hwnd, hMenu))
            return -1;
        return 0;

    case WM_COMMAND:
        if (hMenu == NULL)
            break;
        FakeMenu_OnCommand(hwnd, wParam, lParam, hMenu);
        break;

    case WM_NOTIFY:
        if (hMenu == NULL)
            break;
        return FakeMenu_OnNotify(hwnd, wParam, lParam);

    case WM_DESTROY:
        SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
        if (hMenu == NULL)
            break;
        hMenu->hBaseBar = NULL;
        hMenu->uSelected = 0xFFFFFFFF;
        hMenu->popupped = NULL;
        hMenu->uPopupped = 0xFFFFFFFF;
        hMenu->hBaseBar = NULL;
        if (hMenu->parent == NULL)
            DestroyFakeMenu(hMenu);
        break;

    case WM_SIZE:
        if (hMenu == NULL)
            break;
        FakeMenu_OnSize(hwnd, hMenu);
        break;

    case WM_CLOSEBASEBAR:
        if (hMenu == NULL)
            break;
        FakeMenu_OnClose(hwnd, hMenu, wParam);
        break;

    case WM_ACTIVATE:
        if (hMenu == NULL)
            break;
        FakeMenu_OnActivate(hwnd, wParam, lParam, hMenu);
        break;

    case WM_KEYDOWN:
        if (hMenu == NULL)
            break;
        FakeMenu_OnKeyDown(hwnd, (INT)wParam, hMenu);
        break;

    case WM_SYSCOMMAND:
        if (wParam == SC_CLOSE)
            break;
        return DefWindowProcW(hwnd, uMsg, wParam, lParam);

    case WM_INITMENUPOPUP:
    case WM_DRAWITEM:
    case WM_MEASUREITEM:
        if (s_pcm2)
            s_pcm2->HandleMenuMsg(uMsg, wParam, lParam);
        break;

    default:
        return DefWindowProcW(hwnd, uMsg, wParam, lParam);
    }
    return 0;
}

unsigned __stdcall IconThreadProc(void *param)
{
    HFAKEMENU hMenu = (HFAKEMENU)param;

    RECT rcItem;
    SHFILEINFOW shfi;
    HFAKEMENU hChild;
    UINT i, nCount = hMenu->children.size();
    for (i = 0; i < nCount; i++)
    {
        if (hMenu->ThreadCancel)
            return 1;

        hChild = hMenu->children[i];
        if (hChild->pidl && hChild->hIcon == NULL)
        {
            if (hMenu->nIconSize == 32)
            {
                SHGetFileInfoW((LPCWSTR)hChild->pidl, 0, &shfi, sizeof(shfi),
                    SHGFI_PIDL | SHGFI_LARGEICON | SHGFI_ICON);
            }
            else
            {
                SHGetFileInfoW((LPCWSTR)hChild->pidl, 0, &shfi, sizeof(shfi),
                    SHGFI_PIDL | SHGFI_SMALLICON | SHGFI_ICON);
            }
            if (shfi.hIcon == NULL)
                return 2;

            hChild->hIcon = shfi.hIcon;

            if (hMenu->ThreadCancel)
                return 1;

            SendMessageW(hMenu->hToolBar, TB_GETITEMRECT, i, (LPARAM)&rcItem);
            InvalidateRect(hMenu->hToolBar, &rcItem, FALSE);
        }
    }
    return 0;
}

#include <algorithm>

class LessFunctor
{
public:
    bool operator()(HFAKEMENU data1, HFAKEMENU data2)
    {
        if ((data1->uFlags & MF_POPUP))
        {
            if ((data2->uFlags & MF_POPUP))
                return lstrcmpW(data1->text.c_str(), data2->text.c_str()) < 0;
            else
                return true;
        }
        else
        {
            if ((data2->uFlags & MF_POPUP))
                return false;
            else
                return lstrcmpW(data1->text.c_str(), data2->text.c_str()) < 0;
        }
    }
};

VOID SortFakeMenu(HFAKEMENU hMenu)
{
    std::sort(hMenu->children.begin(), hMenu->children.end(), LessFunctor());
}

class RecentFunctor
{
public:
    bool operator()(HFAKEMENU data1, HFAKEMENU data2)
    {
        WCHAR szFileName1[MAX_PATH], szFileName2[MAX_PATH];
        HANDLE hFind1, hFind2;
        WIN32_FIND_DATAW find1, find2;
        FILETIME ft1, ft2;

        SHGetPathFromIDListW(data1->pidl, szFileName1);
        SHGetPathFromIDListW(data2->pidl, szFileName2);

        hFind1 = FindFirstFileW(szFileName1, &find1);
        ft1 = find1.ftLastWriteTime;
        FindClose(hFind1);
        hFind2 = FindFirstFileW(szFileName2, &find2);
        ft2 = find1.ftLastWriteTime;
        FindClose(hFind2);

        if (hFind1 == INVALID_HANDLE_VALUE && hFind2 == INVALID_HANDLE_VALUE)
            return false;
        if (hFind1 == INVALID_HANDLE_VALUE)
            return true;
        if (hFind2 == INVALID_HANDLE_VALUE)
            return false;
        return CompareFileTime(&ft1, &ft2) < 0;
    }
};

VOID SortFakeMenuByFileTime(HFAKEMENU hMenu)
{
    std::sort(hMenu->children.begin(), hMenu->children.end(), RecentFunctor());
}

BOOL APIENTRY
TrackPopupFakeMenuNoLock(HFAKEMENU hMenu,
                         UINT uFlags,
                         INT x, INT y,
                         INT nReserved,
                         HWND hwndOwner,
                         LPCRECT prc)
{
    if (nReserved == 2)
        hMenu->uSelected = 0;
    else
        hMenu->uSelected = 0xFFFFFFFF;
    hMenu->uPopupped = 0xFFFFFFFF;
    hMenu->popupped = NULL;

    if (hMenu->children.size() == 0)
    {
        AppendFakeMenu(hMenu, MF_GRAYED | MF_STRING, 0, LoadStringDx1(IDS_EMPTY));
    }
    else if (hMenu->Recent)
    {
        SortFakeMenuByFileTime(hMenu);
        size_t i, size = hMenu->children.size();
        if (size > MAX_RECENT)
        {
            for (i = MAX_RECENT; i < size; i++)
            {
                if (hMenu->children[i]->pidl)
                    CoTaskMemFree(hMenu->children[i]->pidl);
            }
            hMenu->children.resize(MAX_RECENT);
        }
    }
    if (!hMenu->NoSort)
        SortFakeMenu(hMenu);

    DWORD style = WS_POPUP | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | WS_DLGFRAME;
    DWORD exstyle = WS_EX_PALETTEWINDOW;

    hMenu->hwndOwner = hwndOwner;
    hMenu->hBaseBar = CreateWindowExW(exstyle,
                                      g_szBaseBarClsName,
                                      NULL,
                                      style,
                                      0, 0, 0, 0,
                                      hwndOwner,
                                      NULL,
                                      g_hInstance,
                                      hMenu);
    if (hMenu->hBaseBar == NULL)
    {
        return FALSE;
    }

    POINT pt;
    SIZE siz = hMenu->sizToolBar;
    RECT rc = {x, y, x + siz.cx, y + siz.cy};
    AdjustWindowRectEx(&rc, style, FALSE, exstyle);

    siz.cx = rc.right - rc.left;
    siz.cy = rc.bottom - rc.top;
    FakeMenu_ChoosePos(hMenu->hBaseBar,
                       x, y, siz.cx, siz.cy,
                       &pt,
                       &siz,
                       nReserved,
                       uFlags,
                       prc);
    MoveWindow(hMenu->hBaseBar, pt.x, pt.y, siz.cx, siz.cy, TRUE);

#ifdef BEGINMENU_ANIMATE
    DWORD dwFlags;
    DWORD dwDelay = 150;
    if (nReserved == 0)
    {
        if (pt.y <= prc->top)
            dwFlags = AW_SLIDE | AW_ACTIVATE | AW_VER_NEGATIVE;
        else
            dwFlags = AW_SLIDE | AW_ACTIVATE | AW_VER_POSITIVE;
    }
    else
    {
        if (pt.x <= prc->left)
            dwFlags = AW_SLIDE | AW_ACTIVATE | AW_HOR_NEGATIVE;
        else
            dwFlags = AW_SLIDE | AW_ACTIVATE | AW_HOR_POSITIVE;
    }
    AnimateWindow(hMenu->hBaseBar, dwDelay, dwFlags);
#endif

    ShowWindow(hMenu->hBaseBar, SW_SHOW);

    hMenu->ThreadCancel = FALSE;
    hMenu->hIconThread =
        (HANDLE)_beginthreadex(NULL, 0, IconThreadProc, hMenu, 0, NULL);

    return TRUE;
}

BOOL APIENTRY
RegisterFakeMenu(HINSTANCE hInstance)
{
    WNDCLASSW wc;
#ifndef CS_DROPSHADOW
    #define CS_DROPSHADOW 0x20000
#endif
    wc.style = CS_DROPSHADOW | CS_SAVEBITS;
    wc.lpfnWndProc = FakeMenu_WndProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = 0;
    wc.hInstance = hInstance;
    wc.hIcon = NULL;
    wc.hCursor = LoadCursorW(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_3DFACE + 1);
    wc.lpszMenuName = NULL;
    wc.lpszClassName = g_szBaseBarClsName;
    return !!RegisterClassW(&wc);
}

HFAKEMENU APIENTRY
CreatePopupFakeMenu(VOID)
{
    HFAKEMENU hMenu;
    hMenu = new FAKEMENU;
    if (hMenu)
    {
        hMenu->hwndOwner = NULL;
        hMenu->hBaseBar = hMenu->hPager = hMenu->hToolBar = NULL;
        hMenu->cyItem = 0;
        hMenu->uID = 0xFFFFFFFF;
        hMenu->uFlags = MF_ENABLED | MF_POPUP | MF_FAKEMENU;
        hMenu->parent = NULL;
        hMenu->uSelected = hMenu->uPopupped = 0xFFFFFFFF;
        hMenu->popupped = NULL;
        hMenu->hIcon = NULL;
        hMenu->pidl = NULL;
        hMenu->nIconSize = 16;
        hMenu->NoSort = FALSE;
        hMenu->hIconThread = NULL;
        hMenu->CustomIcon = FALSE;
        hMenu->Recent = FALSE;
    }
    return hMenu;
}

HRESULT GetFileList(LPITEMIDLIST pidl, std::vector<DATA>& list)
{
    HRESULT result;
    WCHAR szDisplayName[MAX_PATH], szPath[MAX_PATH];
    LPITEMIDLIST pidlChild;
    DATA data;

    IShellFolder *psfDesktop, *psf;
    IEnumIDList *pEnum;
    SHGetDesktopFolder(&psfDesktop);
    result = psfDesktop->BindToObject(pidl, NULL, IID_IShellFolder, (VOID **)&psf);
    if (SUCCEEDED(result))
    {
        const DWORD grfFlags = SHCONTF_FOLDERS | SHCONTF_NONFOLDERS;
        result = psf->EnumObjects(GetDesktopWindow(), grfFlags, &pEnum);
        if (SUCCEEDED(result))
        {
            while (!pEnum->Next(1, &pidlChild, NULL))
            {
                STRRET strret;
                psf->GetDisplayNameOf(pidlChild, SHGDN_NORMAL, &strret);
                StrRetToBufW(&strret, pidlChild, szDisplayName, MAX_PATH);
                if (strret.uType == STRRET_WSTR)
                    CoTaskMemFree(strret.pOleStr);

                LPITEMIDLIST pidlAbsolute = ILCombine(pidl, pidlChild);
                SHGetPathFromIDListW(pidlAbsolute, szPath);

                data.is_folder = !!(GetFileAttributesW(szPath) & FILE_ATTRIBUTE_DIRECTORY);
                data.path = szPath;
                data.name = szDisplayName;
                if (data.pidl)
                    CoTaskMemFree(data.pidl);
                data.pidl = ILClone(pidlAbsolute);
                list.push_back(data);

                CoTaskMemFree(pidlAbsolute);
                CoTaskMemFree(pidlChild);
            }
            pEnum->Release();
        }
        psf->Release();
    }
    psfDesktop->Release();
    return result;
}

BOOL GetPathOfShortcut(HWND hWnd, LPCWSTR pszLnkFile, LPWSTR pszPath)
{
    WCHAR           szPath[MAX_PATH];
    IShellLink*     pShellLink;
    IPersistFile*   pPersistFile;
    WIN32_FIND_DATA find;
    BOOL            bRes = FALSE;

    szPath[0] = '\0';
    HRESULT hRes = CoInitialize(NULL);
    if (SUCCEEDED(hRes))
    {
        if (SUCCEEDED(hRes = CoCreateInstance(CLSID_ShellLink, NULL, 
            CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID *)&pShellLink)))
        {
            if (SUCCEEDED(hRes = pShellLink->QueryInterface(IID_IPersistFile, 
                (VOID **)&pPersistFile)))
            {
                hRes = pPersistFile->Load(pszLnkFile,  STGM_READ);
                if (SUCCEEDED(hRes))
                {
                    if (SUCCEEDED(hRes = pShellLink->GetPath(szPath, 
                        MAX_PATH, &find, SLGP_SHORTPATH)))
                    {
                        if ('\0' != szPath[0])
                        {
                            lstrcpy(pszPath, szPath);
                            bRes = TRUE;
                        }
                    }
                }
                pPersistFile->Release();
            }
            pShellLink->Release();
        }
        CoUninitialize();
    }
    return bRes;
}

DWORD AddItemsToMenu(HFAKEMENU hMenu, LPITEMIDLIST pidl)
{
    std::vector<DATA> list;
    HRESULT result;
    WCHAR szPath[MAX_PATH], szTarget[MAX_PATH];
    INT cch;
    LPWSTR pch;
    INT nCount = 0;

    result = GetFileList(pidl, list);
    if (SUCCEEDED(result))
    {
        if (hMenu->Recent)
        {
            for (size_t i = 0; i < list.size(); i++)
            {
                SHGetPathFromIDListW(list[i].pidl, szPath);
                cch = lstrlenW(szPath);
                if (cch > 4)
                {
                    pch = szPath + cch - 4;
                    if (lstrcmpiW(pch, L".LNK") == 0 &&
                        GetPathOfShortcut(NULL, szPath, szTarget))
                    {
                        DWORD attrs = GetFileAttributesW(szTarget);
                        if (attrs == 0xFFFFFFFF || (attrs & FILE_ATTRIBUTE_DIRECTORY))
                            continue;
                    }
                    else
                        continue;
                }
                else
                    continue;

                AppendFakeMenu(hMenu, MF_ENABLED | MF_STRING, i,
                               list[i].name.c_str(), list[i].pidl);
                list[i].pidl = NULL;
                nCount++;
            }
        }
        else
        {
            for (size_t i = 0; i < list.size(); i++)
            {
                if (list[i].is_folder)
                {
                    HFAKEMENU hSubMenu = CreatePopupFakeMenu();
                    if (hSubMenu != NULL)
                    {
                        AppendFakeMenu(hMenu, MF_ENABLED | MF_POPUP, (UINT_PTR)hSubMenu,
                                       list[i].name.c_str(), list[i].pidl);
                        list[i].pidl = NULL;
                        nCount++;
                    }
                }
                else
                {
                    AppendFakeMenu(hMenu, MF_ENABLED | MF_STRING, i,
                                   list[i].name.c_str(), list[i].pidl);
                    list[i].pidl = NULL;
                    nCount++;
                }
            }
        }
        if (nCount > 0)
            return ERROR_SUCCESS;
        else
            return ERROR_MENU_ITEM_NOT_FOUND;
    }
    else
    {
        return ERROR_ACCESS_DENIED;
    }
}
