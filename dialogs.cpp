#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <commctrl.h>
#include <shlobj.h>
#include <tchar.h>
#include <string>
#include <vector>
#include <algorithm>

#include "fakemenu.h"
#include "BeginBar.h"
#include "resource.h"


BOOL ITEM::LoadIcon()
{
    WCHAR szPath[MAX_PATH];
    WORD w;

    if (hIcon)
    {
        DestroyIcon(hIcon);
        hIcon = NULL;
    }

    switch(type)
    {
    case ITEMTYPE_GROUP:
        if (iconpath.size())
        {
            lstrcpynW(szPath, iconpath.c_str(), MAX_PATH);
            w = 0;
            hIcon = ExtractAssociatedIconW(g_hInstance, szPath, &w);
        }
        break;

    case ITEMTYPE_SEP:
        break;

    case ITEMTYPE_CSIDL1:
    case ITEMTYPE_CSIDL2:
        if (iconpath.size())
        {
            lstrcpynW(szPath, iconpath.c_str(), MAX_PATH);
            w = 0;
            hIcon = ExtractAssociatedIconW(g_hInstance, szPath, &w);
        }
        else
        {
            LPITEMIDLIST pidl;
            if (SUCCEEDED(SHGetSpecialFolderLocation(NULL, csidl1, &pidl)))
            {
                SHFILEINFO shfi;
                ZeroMemory(&shfi, sizeof(shfi));
                SHGetFileInfoW((LPWSTR)pidl, 0, &shfi, sizeof(shfi),
                               SHGFI_PIDL | SHGFI_ICON);
                hIcon = shfi.hIcon;
                CoTaskMemFree(pidl);
            }
        }
        break;

    case ITEMTYPE_PATH1:
    case ITEMTYPE_PATH2:
        if (iconpath.size())
            lstrcpynW(szPath, iconpath.c_str(), MAX_PATH);
        else
            lstrcpynW(szPath, path1.c_str(), MAX_PATH);
        w = 0;
        hIcon = ExtractAssociatedIconW(g_hInstance, szPath, &w);
        break;

    case ITEMTYPE_ACTION:
        if (iconpath.size())
        {
            lstrcpynW(szPath, iconpath.c_str(), MAX_PATH);
            w = 0;
            hIcon = ExtractAssociatedIconW(g_hInstance, szPath, &w);
        }
        break;
    }
    return hIcon != NULL;
}

VOID BEGINBUTTON::LoadIcons()
{
    INT i, nCount = (INT)items.size();
    for (i = 0; i < nCount; i++)
        items[i].LoadIcon();
}

struct LISTBOXSEL
{
    INT nCount;
    INT nSelCount;
    LPINT pSelItems;

    LISTBOXSEL()
    {
        nCount = nSelCount = 0;
        pSelItems = NULL;
    }
    
    LISTBOXSEL(HWND hListBox)
    {
        pSelItems = NULL;
        SaveSel(hListBox);
    }

    ~LISTBOXSEL()
    {
        delete[] pSelItems;
    }

    BOOL IsItemSelected(INT i) const
    {
        for (INT j = 0; j < nSelCount; j++)
        {
            if (pSelItems[j] == i)
                return TRUE;
        }
        return FALSE;
    }

    VOID SelectItem(INT i)
    {
        if (0 <= i && i < nCount)
        {
            if (nSelCount >= 1)
            {
                nSelCount = 1;
                pSelItems[0] = i;
            }
            else
            {
                delete[] pSelItems;
                nSelCount = 1;
                pSelItems = new INT[1];
                pSelItems[0] = i;
            }
        }
        else
            nSelCount = 0;
    }

    VOID SaveSel(HWND hListBox)
    {
        delete[] pSelItems;
        nCount = (INT)SendMessageW(hListBox, LB_GETCOUNT, 0, 0);
        nSelCount = (INT)SendMessageW(hListBox, LB_GETSELCOUNT, 0, 0);
        pSelItems = new INT[nSelCount];
        SendMessageW(hListBox, LB_GETSELITEMS, nSelCount, (LPARAM)pSelItems);
    }

    VOID RestoreSel(HWND hListBox) const
    {
        INT i;
        for (i = 0; i < nCount; i++)
        {
            SendMessageW(hListBox, LB_SETSEL, FALSE, i);
        }
        for (i = 0; i < nSelCount; i++)
        {
            SendMessageW(hListBox, LB_SETSEL, TRUE, pSelItems[i]);
        }
    }

    VOID MoveUp()
    {
        INT i;
        if (nSelCount == 0 || pSelItems[0] == 0)
            return;

        for (i = 0; i < nSelCount; i++)
        {
            pSelItems[i]--;
        }
    }

    VOID MoveDown()
    {
        INT i;
        if (nSelCount == 0 || pSelItems[nSelCount - 1] == nCount - 1)
            return;

        for (i = 0; i < nSelCount; i++)
        {
            pSelItems[i]++;
        }
    }
};

static VOID
Buttons_UpdateUI(HWND hwnd, const std::vector<BEGINBUTTON>& buttons, const LISTBOXSEL& sel)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    INT i, nCount = (INT)buttons.size();
    SendMessageW(hListBox, WM_SETREDRAW, FALSE, 0);
    {
        SendMessageW(hListBox, LB_RESETCONTENT, 0, 0);
        for (i = 0; i < nCount; i++)
        {
            SendMessageW(hListBox, LB_ADDSTRING, 0,
                         (LPARAM)buttons[i].items[0].text.c_str());
        }
        for (i = 0; i < nCount; i++)
        {
            SendMessageW(hListBox, LB_SETITEMDATA, i, (LPARAM)&buttons[i]);
        }
        if (sel.nCount)
            sel.RestoreSel(hListBox);
    }
    SendMessageW(hListBox, WM_SETREDRAW, TRUE, 0);

    INT n, nSelCount = sel.nSelCount;

    n = (INT)SendMessageW(hListBox, LB_GETSEL, 0, 0);
    n = nSelCount && !(n != LB_ERR && n);
    EnableWindow(GetDlgItem(hwnd, psh3), n);

    n = (INT)SendMessageW(hListBox, LB_GETSEL, nCount - 1, 0);
    n = nSelCount && !(n != LB_ERR && n);
    EnableWindow(GetDlgItem(hwnd, psh4), n);

    EnableWindow(GetDlgItem(hwnd, psh5), nSelCount > 0);
    EnableWindow(GetDlgItem(hwnd, psh6), nSelCount > 0);
}

static VOID
Buttons_OnInitDialog(HWND hwnd, const std::vector<BEGINBUTTON>& buttons)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    LISTBOXSEL sel;
    sel.SelectItem(-1);
    Buttons_UpdateUI(hwnd, buttons, sel);
}

static VOID
Buttons_OnUp(HWND hwnd, std::vector<BEGINBUTTON>& buttons)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    INT n, i, nCount = (INT)buttons.size();
    n = (INT)SendMessageW(hListBox, LB_GETSEL, 0, 0);
    if (n != LB_ERR && n)
        return;

    LISTBOXSEL sel(hListBox);
    for (i = 1; i < nCount; i++)
    {
        n = (INT)SendMessageW(hListBox, LB_GETSEL, i, 0);
        if (n != LB_ERR && n)
            std::swap(buttons[i - 1], buttons[i]);
    }
    sel.MoveUp();
    Buttons_UpdateUI(hwnd, buttons, sel);
}

static VOID
Buttons_OnDown(HWND hwnd, std::vector<BEGINBUTTON>& buttons)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    INT n, i, nCount = (INT)buttons.size();

    n = (INT)SendMessageW(hListBox, LB_GETSEL, nCount - 1, 0);
    if (n != LB_ERR && n)
        return;

    LISTBOXSEL sel(hListBox);
    for (i = nCount - 2; i >= 0; i--)
    {
        n = (INT)SendMessageW(hListBox, LB_GETSEL, i, 0);
        if (n != LB_ERR && n)
            std::swap(buttons[i], buttons[i + 1]);
    }
    sel.MoveDown();
    Buttons_UpdateUI(hwnd, buttons, sel);
}

static VOID
Buttons_OnAddButton(HWND hwnd, std::vector<BEGINBUTTON>& buttons)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    LISTBOXSEL sel(hListBox);
    INT i;
    if (sel.nSelCount > 0)
        i = sel.pSelItems[sel.nSelCount - 1];
    else
        i = sel.nCount;

    ITEM item;
    BEGINBUTTON button;
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
        std::vector<BEGINBUTTON>::iterator it, end = buttons.end();
        it = buttons.begin();
        for (INT j = 0; j < i; j++)
            it++;
        buttons.insert(it, button);
        sel.nCount++;
        sel.SelectItem(i);
        Buttons_UpdateUI(hwnd, buttons, sel);
    }
}

static VOID
Buttons_OnDelete(HWND hwnd, std::vector<BEGINBUTTON>& buttons)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    LISTBOXSEL sel(hListBox);
    INT i, j, nCount = sel.nCount;
    for (i = nCount - 1; i >= 0; i--)
    {
        if (sel.IsItemSelected(i))
        {
            std::vector<BEGINBUTTON>::iterator it, end = buttons.end();
            j = 0;
            for (it = buttons.begin(); it != end; it++)
            {
                if (j == i)
                    break;
                j++;
            }
            buttons.erase(it);
        }
    }
    sel.SelectItem(-1);
    Buttons_UpdateUI(hwnd, buttons, sel);
}

static VOID
Buttons_OnEditButton(HWND hwnd, std::vector<BEGINBUTTON>& buttons)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    INT i = (INT)SendMessageW(hListBox, LB_GETCURSEL, 0, 0);
    INT nCount = (INT)buttons.size();
    if (i == LB_ERR || i < 0 || i >= nCount)
        return;

    LISTBOXSEL sel(hListBox);
    BEGINBUTTON button = buttons[i];
    if (IDOK == DialogBoxParamW(g_hInstance, MAKEINTRESOURCEW(IDD_MENU), hwnd,
                                Menu_DlgProc, (LPARAM)&button))
    {
        buttons[i] = button;
        Buttons_UpdateUI(hwnd, buttons, sel);
    }
}

static VOID
Buttons_OnChangeIcon(HWND hwnd, std::vector<BEGINBUTTON>& buttons)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    LISTBOXSEL sel(hListBox);

    INT i, j;
    WCHAR szPath[MAX_PATH];
    j = (INT)SendMessageW(hListBox, LB_GETCURSEL, 0, 0);
    lstrcpynW(szPath, buttons[j].items[0].iconpath.c_str(), MAX_PATH);
    if (OpenFileDialog(
        hwnd,
        szPath,
        OFN_DONTADDTORECENT | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
        LoadStringDx1(IDS_SELECTICON),
        L"ico",
        LoadFilterStringDx(IDS_ALLFILES),
        NULL))
    {
        for (i = 0; i < sel.nSelCount; i++)
        {
            j = sel.pSelItems[i];
            buttons[j].items[0].iconpath = szPath;
            buttons[j].items[0].LoadIcon();
        }
        InvalidateRect(hListBox, NULL, TRUE);
    }
}

static BOOL
Buttons_MeasureItem(
    HWND hwnd,
    LPMEASUREITEMSTRUCT pmis,
    std::vector<BEGINBUTTON>& buttons)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    HFONT hFont = (HFONT)SendMessageW(hListBox, WM_GETFONT, 0, 0);

    WCHAR sz[MAX_PATH];
    HDC hdc = GetDC(hListBox);
    SelectObject(hdc, hFont);
    SIZE siz, sizTotal = {0, 0};
    INT i, nCount = (INT)SendMessageW(hListBox, LB_GETCOUNT, 0, 0);
    for (i = 0; i < nCount; i++)
    {
        SendMessageW(hListBox, LB_GETTEXT, i, (LPARAM)sz);
        GetTextExtentPoint32W(hdc, sz, lstrlenW(sz), &siz);
        if (siz.cy < 16)
            siz.cy = 16;
        siz.cx += 16;
        if (sizTotal.cx < siz.cx)
            sizTotal.cx = siz.cx;
        if (sizTotal.cy < siz.cy)
            sizTotal.cy = siz.cy;
    }
    ReleaseDC(hListBox, hdc);

    pmis->itemWidth = sizTotal.cx;
    SendMessageW(hListBox, LB_SETHORIZONTALEXTENT, pmis->itemWidth, 0);
    pmis->itemHeight = sizTotal.cy;
    return TRUE;
}

static BOOL
Buttons_DrawItem(
    HWND hwnd,
    LPDRAWITEMSTRUCT pdis,
    std::vector<BEGINBUTTON>& buttons)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    if (pdis->itemID == (UINT)-1)
        return FALSE;

    WCHAR sz[MAX_PATH];
    SendMessageW(hListBox, LB_GETTEXT, pdis->itemID, (LPARAM)sz);
    BEGINBUTTON *pButton;
    pButton = (BEGINBUTTON *)SendMessageW(hListBox, LB_GETITEMDATA, pdis->itemID, 0);
    pButton->items[0].LoadIcon();

    HDC hdc = pdis->hDC;
    RECT rcItem = pdis->rcItem;
    SetBkMode(hdc, TRANSPARENT);
    if (pdis->itemState & ODS_SELECTED)
    {
        FillRect(hdc, &rcItem, (HBRUSH)(COLOR_HIGHLIGHT + 1));
        SetTextColor(hdc, GetSysColor(COLOR_HIGHLIGHTTEXT));
    }
    else
    {
        FillRect(hdc, &rcItem, (HBRUSH)(COLOR_WINDOW + 1));
        SetTextColor(hdc, GetSysColor(COLOR_WINDOWTEXT));
    }
    DrawIconEx(hdc,
        0,
        rcItem.top + (rcItem.bottom - rcItem.top - 16) / 2,
        pButton->items[0].hIcon,
        16,
        16,
        0,
        NULL,
        DI_NORMAL);
    rcItem.left += 16;
    DrawTextW(hdc, sz, lstrlenW(sz), &rcItem,
              DT_LEFT | DT_VCENTER | DT_SINGLELINE);
    return TRUE;
}

INT_PTR CALLBACK
Buttons_DlgProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    static std::vector<BEGINBUTTON> *pButtons = NULL;
    switch(uMsg)
    {
    case WM_INITDIALOG:
        pButtons = (std::vector<BEGINBUTTON> *)lParam;
        Buttons_OnInitDialog(hwnd, *pButtons);
        return TRUE;

    case WM_MEASUREITEM:
        if (pButtons == NULL)
            break;
        return Buttons_MeasureItem(hwnd, (LPMEASUREITEMSTRUCT)lParam, *pButtons);

    case WM_DRAWITEM:
        if (pButtons == NULL)
            break;
        return Buttons_DrawItem(hwnd, (LPDRAWITEMSTRUCT)lParam, *pButtons);

    case WM_COMMAND:
        switch(LOWORD(wParam))
        {
        case lst1:  // listbox
            if (HIWORD(wParam) == LBN_SELCHANGE)
            {
                HWND hListBox = GetDlgItem(hwnd, lst1);
                LISTBOXSEL sel(hListBox);
                Buttons_UpdateUI(hwnd, *pButtons, sel);
            }
            else if (HIWORD(wParam) == LBN_DBLCLK)
            {
                Buttons_OnEditButton(hwnd, *pButtons);
            }
            break;

        case psh1:  // add button
            Buttons_OnAddButton(hwnd, *pButtons);
            break;

        case psh2:  // edit button
            Buttons_OnEditButton(hwnd, *pButtons);
            break;

        case psh3:  // up
            Buttons_OnUp(hwnd, *pButtons);
            break;

        case psh4:  // down
            Buttons_OnDown(hwnd, *pButtons);
            break;

        case psh5:  // delete
            Buttons_OnDelete(hwnd, *pButtons);
            break;

        case psh6:  // change icon
            Buttons_OnChangeIcon(hwnd, *pButtons);
            break;

        case IDOK:
            EndDialog(hwnd, IDOK);
            break;

        case IDCANCEL:
            EndDialog(hwnd, IDCANCEL);
            break;
        }
    }
    return 0;
}

static BOOL
FixLevelsOfItems(BEGINBUTTON& button)
{
    INT i, nCount = (INT)button.items.size();
    if (nCount == 0)
        return FALSE;

    BOOL b = FALSE;
    button.items[0].level = 0;
    for (i = 1; i < nCount; i++)
    {
        if (button.items[i].level <= 0)
        {
            button.items[i].level = 1;
            b = TRUE;
        }
        if (button.items[i - 1].level + 1 < button.items[i].level)
        {
            button.items[i].level = button.items[i - 1].level + 1;
            b = TRUE;
        }
    }
    return b;
}

static VOID
Menu_UpdateUI(HWND hwnd, const BEGINBUTTON& button, const LISTBOXSEL& sel)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);

    INT i, n, nCount = (INT)button.items.size();
    std::wstring text;
    SendMessageW(hListBox, WM_SETREDRAW, FALSE, 0);
    {
        SendMessageW(hListBox, LB_RESETCONTENT, 0, 0);
        for (i = 0; i < nCount; i++)
        {
            text.clear();

            // "indents<TAB>text<TAB>content<TAB>expand?"

            for (INT j = 0; j < button.items[i].level; j++)
                text += LoadStringDx1(IDS_INDENT);
            text += L"\t";

            if (button.items[i].type == ITEMTYPE_SEP)
                text += LoadStringDx1(IDS_SEP);
            else
                text += button.items[i].text;

            switch(button.items[i].type)
            {
            case ITEMTYPE_GROUP:
                text += L"\t\t";
                break;

            case ITEMTYPE_SEP:
                text += L"\t\t";
                break;

            case ITEMTYPE_CSIDL1:
                text += L"\t";
                text += LoadStringDx1(IDS_CSIDL_0x0000 + button.items[i].csidl1);
                text += L"\t";
                if (button.items[i].expand)
                    text += LoadStringDx1(IDS_EXPAND);
                else
                    text += LoadStringDx1(IDS_NOEXPAND);
                break;

            case ITEMTYPE_CSIDL2:
                text += L"\t";
                text += LoadStringDx1(IDS_CSIDL_0x0000 + button.items[i].csidl1);
                text += L" + ";
                text += LoadStringDx1(IDS_CSIDL_0x0000 + button.items[i].csidl2);
                text += L"\t";
                text += LoadStringDx1(IDS_EXPAND);
                break;

            case ITEMTYPE_PATH1:
                text += L"\t";
                text += button.items[i].path1;
                text += L"\t";
                if (button.items[i].expand)
                    text += LoadStringDx1(IDS_EXPAND);
                else
                    text += LoadStringDx1(IDS_NOEXPAND);
                break;

            case ITEMTYPE_PATH2:
                text += L"\t";
                text += button.items[i].path1;
                text += L" + ";
                text += button.items[i].path2;
                text += L"\t";
                text += LoadStringDx1(IDS_EXPAND);
                break;

            case ITEMTYPE_ACTION:
                text += L"\t";
                text += LoadStringDx1(IDS_ACTION_0000 + button.items[i].action);
                text += L"\t";
                break;
            }
            SendMessageW(hListBox, LB_ADDSTRING, 0, (LPARAM)text.c_str());
        }
        if (sel.nCount)
            sel.RestoreSel(hListBox);
    }
    SendMessageW(hListBox, WM_SETREDRAW, TRUE, 0);

    EnableWindow(GetDlgItem(hwnd, psh2), (sel.nSelCount > 0));
    EnableWindow(GetDlgItem(hwnd, psh3), (sel.nSelCount > 0));
    EnableWindow(GetDlgItem(hwnd, psh4), (sel.nSelCount > 0));

    n = (INT)SendMessageW(hListBox, LB_GETSEL, 0, 0);
    n = sel.nSelCount && !(n != LB_ERR && n);
    EnableWindow(GetDlgItem(hwnd, psh5), n);

    n = (INT)SendMessageW(hListBox, LB_GETSEL, nCount - 1, 0);
    n = sel.nSelCount && !(n != LB_ERR && n);
    EnableWindow(GetDlgItem(hwnd, psh6), n);

    EnableWindow(GetDlgItem(hwnd, psh7), (sel.nSelCount > 0));
    EnableWindow(GetDlgItem(hwnd, psh8), (sel.nSelCount > 0));
}

static VOID
Menu_OnInitDialog(HWND hwnd, BEGINBUTTON& button)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    LISTBOXSEL sel;
    sel.SelectItem(-1);
    Menu_UpdateUI(hwnd, button, sel);
}

static VOID
Menu_OnAddItem(HWND hwnd, BEGINBUTTON& button)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    LISTBOXSEL sel(hListBox);

    INT i;
    if (sel.nSelCount > 0)
        i = sel.pSelItems[sel.nSelCount - 1];
    else
        i = sel.nCount;
    ITEM item;
    if (button.items.size())
        item.level = button.items[button.items.size() - 1].level;
    if (DialogBoxParamW(g_hInstance, MAKEINTRESOURCEW(IDD_ITEM), hwnd,
                        Item_DlgProc, (LPARAM)&item) == IDOK)
    {
        std::vector<ITEM>::iterator it, end = button.items.end();
        it = button.items.begin();
        for (INT j = 0; j < i; j++)
            it++;
        button.items.insert(it, item);
        sel.nCount++;
        sel.SelectItem(i);
        Menu_UpdateUI(hwnd, button, sel);
    }
}

static VOID
Menu_OnDeleteItems(HWND hwnd, BEGINBUTTON& button)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    LISTBOXSEL sel(hListBox);
    INT i, j, nCount = sel.nCount;
    for (i = nCount - 1; i >= 0; i--)
    {
        if (sel.IsItemSelected(i))
        {
            j = 0;
            std::vector<ITEM>::iterator it, end = button.items.end();
            for (it = button.items.begin(); it != end; it++)
            {
                if (j == i)
                    break;
                j++;
            }
            button.items.erase(it);
        }
    }
    sel.SelectItem(-1);
    Menu_UpdateUI(hwnd, button, sel);
}

static VOID
Menu_OnEditItem(HWND hwnd, BEGINBUTTON& button)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    LISTBOXSEL sel(hListBox);
    if (sel.nSelCount == 0)
        return;

    ITEM item = button.items[sel.pSelItems[0]];
    if (DialogBoxParamW(g_hInstance, MAKEINTRESOURCEW(IDD_ITEM), hwnd,
                        Item_DlgProc, (LPARAM)&item) == IDOK)
    {
        button.items[sel.pSelItems[0]] = item;
        Menu_UpdateUI(hwnd, button, sel);
    }
}

static VOID
Menu_OnLeft(HWND hwnd, BEGINBUTTON& button)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    LISTBOXSEL sel(hListBox);
    if (sel.nSelCount == 0)
        return;

    INT i, j, nCount = sel.nSelCount;
    for (i = 0; i < nCount; i++)
    {
        j = sel.pSelItems[i];
        button.items[j].level--;
    }
    FixLevelsOfItems(button);

    Menu_UpdateUI(hwnd, button, sel);
}

static VOID
Menu_OnRight(HWND hwnd, BEGINBUTTON& button)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    LISTBOXSEL sel(hListBox);
    if (sel.nSelCount == 0)
        return;

    INT i, j, nCount = sel.nSelCount;
    for (i = 0; i < nCount; i++)
    {
        j = sel.pSelItems[i];
        button.items[j].level++;
    }
    FixLevelsOfItems(button);

    Menu_UpdateUI(hwnd, button, sel);
}

static VOID
Menu_OnUp(HWND hwnd, BEGINBUTTON& button)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    INT n, i, nCount = (INT)button.items.size();
    n = (INT)SendMessageW(hListBox, LB_GETSEL, 0, 0);
    if (n != LB_ERR && n)
        return;

    LISTBOXSEL sel(hListBox);
    for (i = 1; i < nCount; i++)
    {
        n = (INT)SendMessageW(hListBox, LB_GETSEL, i, 0);
        if (n != LB_ERR && n)
            std::swap(button.items[i - 1], button.items[i]);
    }
    sel.MoveUp();
    FixLevelsOfItems(button);
    Menu_UpdateUI(hwnd, button, sel);
}

static VOID
Menu_OnDown(HWND hwnd, BEGINBUTTON& button)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    INT n, i, nCount = (INT)button.items.size();

    n = (INT)SendMessageW(hListBox, LB_GETSEL, nCount - 1, 0);
    if (n != LB_ERR && n)
        return;

    LISTBOXSEL sel(hListBox);
    for (i = nCount - 2; i >= 0; i--)
    {
        n = (INT)SendMessageW(hListBox, LB_GETSEL, i, 0);
        if (n != LB_ERR && n)
            std::swap(button.items[i], button.items[i + 1]);
    }
    sel.MoveDown();
    FixLevelsOfItems(button);
    Menu_UpdateUI(hwnd, button, sel);
}

static VOID
FixButton(BEGINBUTTON& button)
{
    FixLevelsOfItems(button);

    INT i, nCount = (INT)button.items.size();
    for (i = 0; i < nCount; i++)
    {
        if (button.items[i].type == ITEMTYPE_CSIDL2 ||
            button.items[i].type == ITEMTYPE_PATH2)
        {
            button.items[i].expand = TRUE;
        }
        if (button.items[i].type == ITEMTYPE_PATH1)
        {
            DWORD attrs;
            attrs = GetFileAttributesW(button.items[i].path1.c_str());
            if (!(attrs & FILE_ATTRIBUTE_DIRECTORY))
                button.items[i].expand = FALSE;
        }
    }
    for (i = 1; i < nCount; i++)
    {
        if (button.items[i - 1].level < button.items[i].level)
        {
            button.items[i - 1].type = ITEMTYPE_GROUP;
        }
    }

    button.items[0].LoadIcon();
    if (button.items[0].hIcon == NULL)
        button.items[0].iconpath = GetDefaultIconPath();
}

static VOID
Menu_OnOK(HWND hwnd, BEGINBUTTON& button)
{
    if (button.items.size() == 0)
    {
        MessageBoxW(hwnd, LoadStringDx1(IDS_NEEDSITEM), NULL,
                    MB_ICONERROR);
        return;
    }
    FixButton(button);
    EndDialog(hwnd, IDOK);
}

static VOID
Menu_OnChangeIcon(HWND hwnd, BEGINBUTTON& button)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    LISTBOXSEL sel(hListBox);

    INT i, j;
    WCHAR szPath[MAX_PATH];
    j = (INT)SendMessageW(hListBox, LB_GETCURSEL, 0, 0);
    lstrcpynW(szPath, button.items[j].iconpath.c_str(), MAX_PATH);
    if (OpenFileDialog(
        hwnd,
        szPath,
        OFN_DONTADDTORECENT | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
        LoadStringDx1(IDS_SELECTICON),
        L"ico",
        LoadFilterStringDx(IDS_ALLFILES),
        NULL))
    {
        for (i = 0; i < sel.nSelCount; i++)
        {
            j = sel.pSelItems[i];
            button.items[j].iconpath = szPath;
            button.items[j].LoadIcon();
        }
        InvalidateRect(hListBox, NULL, TRUE);
    }
}

static WCHAR s_szTab[] = L"\t";

static BOOL
Menu_MeasureItem(
    HWND hwnd,
    LPMEASUREITEMSTRUCT pmis,
    BEGINBUTTON& button)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    HFONT hFont = (HFONT)SendMessageW(hListBox, WM_GETFONT, 0, 0);

    WCHAR sz[MAX_PATH * 2 + 256];
    LPWSTR pchIndents, pchText, pchContent, pchExpand;
    HDC hdc = GetDC(hListBox);
    SelectObject(hdc, hFont);

    SIZE siz, siz2;
    INT cxIndentsAndText = 0, cxContent = 0, cxExpand = 0;
    INT cy = 0;

    INT i, nCount = (INT)SendMessageW(hListBox, LB_GETCOUNT, 0, 0);
    for (i = 0; i < nCount; i++)
    {
        SendMessageW(hListBox, LB_GETTEXT, i, (LPARAM)sz);

        // "indents<TAB>text<TAB>content<TAB>expand?"
        pchIndents = wcstok(sz, s_szTab);
        pchText = wcstok(NULL, s_szTab);
        pchContent = wcstok(NULL, s_szTab);
        pchExpand = wcstok(NULL, s_szTab);

        GetTextExtentPoint32W(hdc, pchIndents, lstrlenW(pchIndents), &siz);
        GetTextExtentPoint32W(hdc, pchText, lstrlenW(pchText), &siz2);
        siz.cx += siz2.cx + 16;
        if (cxIndentsAndText < siz.cx)
            cxIndentsAndText = siz.cx;
        if (cy < siz.cy)
            cy = siz.cy;
        if (cy < siz2.cy)
            cy = siz2.cy;

        GetTextExtentPoint32W(hdc, pchContent, lstrlenW(pchContent), &siz);
        if (cxContent < siz.cx)
            cxContent = siz.cx;
        if (cy < siz.cy)
            cy = siz.cy;

        GetTextExtentPoint32W(hdc, pchExpand, lstrlenW(pchExpand), &siz);
        if (cxExpand < siz.cx)
            cxExpand = siz.cx;
        if (cy < siz.cy)
            cy = siz.cy;
    }
    ReleaseDC(hListBox, hdc);

    cxIndentsAndText += 8;
    cxContent += 8;
    button.cxIndentsAndText = cxIndentsAndText;
#define CONTENT_MAX 200
    if (cxContent > CONTENT_MAX)
        cxContent = CONTENT_MAX;
    button.cxContent = cxContent;
    button.cxExpand = cxExpand;

    pmis->itemWidth = cxIndentsAndText + cxContent + cxExpand + 8 * 2;
    SendMessageW(hListBox, LB_SETHORIZONTALEXTENT, pmis->itemWidth, 0);
    pmis->itemHeight = cy;
    return TRUE;
}

static BOOL
Menu_DrawItem(
    HWND hwnd,
    LPDRAWITEMSTRUCT pdis,
    BEGINBUTTON& button)
{
    HWND hListBox = GetDlgItem(hwnd, lst1);
    if (pdis->itemID == (UINT)-1)
        return FALSE;

    button.items[pdis->itemID].LoadIcon();

    MEASUREITEMSTRUCT mis;
    Menu_MeasureItem(hwnd, &mis, button);

    WCHAR sz[MAX_PATH * 2 + 256];
    LPWSTR pchIndents, pchText, pchContent, pchExpand;
    SendMessageW(hListBox, LB_GETTEXT, pdis->itemID, (LPARAM)sz);
    // "indents<TAB>text<TAB>content<TAB>expand?"
    pchIndents = wcstok(sz, s_szTab);
    pchText = wcstok(NULL, s_szTab);
    pchContent = wcstok(NULL, s_szTab);
    pchExpand = wcstok(NULL, s_szTab);

    HDC hdc = pdis->hDC;
    RECT rc, rcItem = pdis->rcItem;
    SIZE siz;
    SetBkMode(hdc, TRANSPARENT);
    if (pdis->itemState & ODS_SELECTED)
    {
        FillRect(hdc, &rcItem, (HBRUSH)(COLOR_HIGHLIGHT + 1));
        SetTextColor(hdc, GetSysColor(COLOR_HIGHLIGHTTEXT));
    }
    else
    {
        FillRect(hdc, &rcItem, (HBRUSH)(COLOR_WINDOW + 1));
        SetTextColor(hdc, GetSysColor(COLOR_WINDOWTEXT));
    }

    rc = rcItem;
    DrawTextW(hdc, pchIndents, lstrlenW(pchIndents), &rc,
              DT_LEFT | DT_SINGLELINE | DT_VCENTER);
    if (lstrlenW(pchIndents) == 0)
        siz.cx = 0;
    else
        GetTextExtentPoint32(hdc, pchIndents, lstrlenW(pchIndents), &siz);

    DrawIconEx(hdc,
        siz.cx,
        rcItem.top + (rcItem.bottom - rcItem.top - 16) / 2,
        button.items[pdis->itemID].hIcon,
        16,
        16,
        0,
        NULL,
        DI_NORMAL);

    rc = rcItem;
    rc.left += siz.cx + 16;
    DrawTextW(hdc, pchText, lstrlenW(pchText), &rc,
              DT_LEFT | DT_SINGLELINE | DT_VCENTER);

    rc = rcItem;
    rc.left += button.cxIndentsAndText;
    rc.right = rc.left + CONTENT_MAX;
    DrawTextW(hdc, pchContent, lstrlenW(pchContent), &rc,
              DT_LEFT | DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);

    rc = rcItem;
    rc.left += button.cxIndentsAndText + button.cxContent;
    DrawTextW(hdc, pchExpand, lstrlenW(pchExpand), &rc,
              DT_LEFT | DT_SINGLELINE | DT_VCENTER);

    return TRUE;
}

INT_PTR CALLBACK
Menu_DlgProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    static BEGINBUTTON *pButton = NULL;
    switch(uMsg)
    {
    case WM_INITDIALOG:
        pButton = (BEGINBUTTON *)lParam;
        Menu_OnInitDialog(hwnd, *pButton);
        return TRUE;

    case WM_MEASUREITEM:
        if (pButton == NULL)
            break;
        return Menu_MeasureItem(hwnd, (LPMEASUREITEMSTRUCT)lParam, *pButton);

    case WM_DRAWITEM:
        if (pButton == NULL)
            break;
        return Menu_DrawItem(hwnd, (LPDRAWITEMSTRUCT)lParam, *pButton);

    case WM_COMMAND:
        switch(LOWORD(wParam))
        {
        case lst1:  // listbox
            if (HIWORD(wParam) == LBN_SELCHANGE)
            {
                HWND hListBox = GetDlgItem(hwnd, lst1);
                LISTBOXSEL sel(hListBox);
                Menu_UpdateUI(hwnd, *pButton, sel);
            }
            else if (HIWORD(wParam) == LBN_DBLCLK)
            {
                Menu_OnEditItem(hwnd, *pButton);
            }
            break;

        case psh1:  // add item
            Menu_OnAddItem(hwnd, *pButton);
            break;

        case psh2:  // edit item
            Menu_OnEditItem(hwnd, *pButton);
            break;

        case psh3:  // left
            Menu_OnLeft(hwnd, *pButton);
            break;

        case psh4:  // right
            Menu_OnRight(hwnd, *pButton);
            break;

        case psh5:  // up
            Menu_OnUp(hwnd, *pButton);
            break;

        case psh6:  // down
            Menu_OnDown(hwnd, *pButton);
            break;

        case psh7:  // delete
            Menu_OnDeleteItems(hwnd, *pButton);
            break;

        case psh8:  // change icon
            Menu_OnChangeIcon(hwnd, *pButton);
            break;

        case IDOK:
            Menu_OnOK(hwnd, *pButton);
            break;

        case IDCANCEL:
            EndDialog(hwnd, IDCANCEL);
            break;
        }
    }
    return 0;
}

static VOID
Item_OnInitDialog(HWND hwnd, ITEM& item)
{
    item.LoadIcon();
    SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);

    SetDlgItemTextW(hwnd, edt1, item.text.c_str());

    HWND hCmb1, hCmb2, hCmb3, hCmb4, hCmb5;
    hCmb1 = GetDlgItem(hwnd, cmb1);
    hCmb2 = GetDlgItem(hwnd, cmb2);
    hCmb3 = GetDlgItem(hwnd, cmb3);
    hCmb4 = GetDlgItem(hwnd, cmb4);
    hCmb5 = GetDlgItem(hwnd, cmb5);

    TCHAR sz[128];
    COMBOBOXEXITEM cbex;

    INT i1 = -1, i2 = -1;
    ZeroMemory(&cbex, sizeof(cbex));
    cbex.mask = CBEIF_TEXT | CBEIF_LPARAM;
    cbex.iItem = -1;
    cbex.pszText = LoadStringDx1(IDS_NONE);
    cbex.lParam = -1;
    SendMessageW(hCmb1, CBEM_INSERTITEM, 0, (LPARAM)&cbex);
    SendMessageW(hCmb2, CBEM_INSERTITEM, 0, (LPARAM)&cbex);
    for (INT i = IDS_CSIDL_0x0000; i <= IDS_CSIDL_0x0031; i++)
    {
        if (IDS_CSIDL_0x000C <= i && i <= IDS_CSIDL_0x000F)
            continue;
        LoadStringW(g_hInstance, i, sz, 128);
        cbex.pszText = sz;
        cbex.iItem = -1;
        cbex.lParam = i - IDS_CSIDL_0x0000;

        if (item.csidl1 == i - IDS_CSIDL_0x0000)
            i1 = (INT)SendMessageW(hCmb1, CBEM_INSERTITEM, 0, (LPARAM)&cbex);
        else
            SendMessageW(hCmb1, CBEM_INSERTITEM, 0, (LPARAM)&cbex);

        if (item.csidl2 == i - IDS_CSIDL_0x0000)
            i2 = (INT)SendMessageW(hCmb2, CBEM_INSERTITEM, 0, (LPARAM)&cbex);
        else
            SendMessageW(hCmb2, CBEM_INSERTITEM, 0, (LPARAM)&cbex);
    }
    SendMessageW(hCmb1, CB_SETCURSEL, 0, 0);
    SendMessageW(hCmb2, CB_SETCURSEL, 0, 0);

    std::vector<std::wstring> paths;
    if (LoadPathSettings(paths))
    {
        ZeroMemory(&cbex, sizeof(cbex));
        cbex.mask = CBEIF_TEXT;
        cbex.iItem = -1;
        INT i, nCount = (INT)paths.size();
        for (i = 0; i < nCount; i++)
        {
            lstrcpynW(sz, paths[i].c_str(), 128);
            cbex.pszText = sz;
            SendMessageW(hCmb3, CBEM_INSERTITEM, 0, (LPARAM)&cbex);
            SendMessageW(hCmb4, CBEM_INSERTITEM, 0, (LPARAM)&cbex);
        }
    }

    INT iAction = -1;
    ZeroMemory(&cbex, sizeof(cbex));
    cbex.mask = CBEIF_TEXT | CBEIF_LPARAM;
    cbex.iItem = -1;
    cbex.pszText = LoadStringDx1(IDS_NONE);
    cbex.lParam = -1;
    SendMessageW(hCmb5, CBEM_INSERTITEM, 0, (LPARAM)&cbex);
    for (INT i = IDS_ACTION_0000; i <= IDS_ACTION_0003; i++)
    {
        cbex.pszText = LoadStringDx1(i);
        cbex.lParam = ID_ACTION_FIRST + i - IDS_ACTION_0000;
        if (cbex.lParam == item.action)
            iAction = (INT)SendMessageW(hCmb5, CBEM_INSERTITEM, 0, (LPARAM)&cbex);
        else
            SendMessageW(hCmb5, CBEM_INSERTITEM, 0, (LPARAM)&cbex);
    }
    SendMessageW(hCmb5, CB_SETCURSEL, 0, 0);

    switch(item.type)
    {
    case ITEMTYPE_GROUP:
        CheckRadioButton(hwnd, rad1, rad5, rad1);
        EnableWindow(GetDlgItem(hwnd, chx1), FALSE);
        EnableWindow(GetDlgItem(hwnd, chx2), FALSE);
        break;

    case ITEMTYPE_CSIDL1:
    case ITEMTYPE_CSIDL2:
        CheckRadioButton(hwnd, rad1, rad5, rad2);
        if (item.expand)
            CheckDlgButton(hwnd, chx1, BST_CHECKED);
        SendMessageW(hCmb1, CB_SETCURSEL, i1, 0);
        if (item.type == ITEMTYPE_CSIDL2)
            SendMessageW(hCmb2, CB_SETCURSEL, i2, 0);
        else
            SendMessageW(hCmb2, CB_SETCURSEL, 0, 0);
        EnableWindow(GetDlgItem(hwnd, chx1), TRUE);
        EnableWindow(GetDlgItem(hwnd, chx2), FALSE);
        break;

    case ITEMTYPE_PATH1:
    case ITEMTYPE_PATH2:
        CheckRadioButton(hwnd, rad1, rad5, rad3);
        if (item.expand)
            CheckDlgButton(hwnd, chx2, BST_CHECKED);
        SetWindowTextW(hCmb3, item.path1.c_str());
        if (item.type == ITEMTYPE_PATH2)
            SetWindowTextW(hCmb4, item.path2.c_str());
        else
            SetWindowTextW(hCmb4, NULL);
        EnableWindow(GetDlgItem(hwnd, chx1), FALSE);
        EnableWindow(GetDlgItem(hwnd, chx2), TRUE);
        break;

    case ITEMTYPE_ACTION:
        CheckRadioButton(hwnd, rad1, rad5, rad4);
        EnableWindow(GetDlgItem(hwnd, chx1), FALSE);
        EnableWindow(GetDlgItem(hwnd, chx2), FALSE);
        SendMessageW(hCmb5, CB_SETCURSEL, iAction, 0);
        break;

    case ITEMTYPE_SEP:
        CheckRadioButton(hwnd, rad1, rad5, rad5);
        EnableWindow(GetDlgItem(hwnd, chx1), FALSE);
        EnableWindow(GetDlgItem(hwnd, chx2), FALSE);
        break;
    }
}

static BOOL
Item_SetData(HWND hwnd, ITEM& item)
{
    WCHAR szText[MAX_PATH];
    GetDlgItemTextW(hwnd, edt1, szText, MAX_PATH);
    std::wstring text(szText);
    trim(text);
    if (!IsDlgButtonChecked(hwnd, rad5) && text.size() == 0)
    {
        SetFocus(GetDlgItem(hwnd, edt1));
        MessageBoxW(hwnd, LoadStringDx1(IDS_INPUTITEMNAME), NULL, MB_ICONERROR);
        return FALSE;
    }
    item.text = text;

    if (IsDlgButtonChecked(hwnd, rad1))
    {
        item.type = ITEMTYPE_GROUP;
        item.csidl1 = item.csidl2 = 0;
        item.path1.clear();
        item.path2.clear();
        item.action = 0;
        item.expand = FALSE;
    }
    else if (IsDlgButtonChecked(hwnd, rad2))
    {
        INT i1 = (INT)SendDlgItemMessageW(hwnd, cmb1, CB_GETCURSEL, 0, 0);
        INT i2 = (INT)SendDlgItemMessageW(hwnd, cmb2, CB_GETCURSEL, 0, 0);

        COMBOBOXEXITEM cbex;

        ZeroMemory(&cbex, sizeof(cbex));
        cbex.mask = CBEIF_LPARAM;
        cbex.iItem = i1;
        cbex.lParam = -1;
        SendDlgItemMessageW(hwnd, cmb1, CBEM_GETITEM, i1, (LPARAM)&cbex);
        i1 = (INT)cbex.lParam;

        ZeroMemory(&cbex, sizeof(cbex));
        cbex.mask = CBEIF_LPARAM;
        cbex.iItem = i2;
        cbex.lParam = -1;
        SendDlgItemMessageW(hwnd, cmb2, CBEM_GETITEM, i2, (LPARAM)&cbex);
        i2 = (INT)cbex.lParam;

        if (i1 == -1 || i1 == CB_ERR)
        {
            SetFocus(GetDlgItem(hwnd, cmb1));
            MessageBoxW(hwnd, LoadStringDx1(IDS_INPUTCSIDL), NULL, MB_ICONERROR);
            return FALSE;
        }
        if (i2 == -1 || i2 == CB_ERR)
        {
            item.type = ITEMTYPE_CSIDL1;
            item.csidl1 = i1;
            item.csidl2 = 0;
        }
        else
        {
            item.type = ITEMTYPE_CSIDL2;
            item.csidl1 = i1;
            item.csidl2 = i2;
        }
        item.path1.clear();
        item.path2.clear();
        item.action = 0;
        item.expand = !!IsDlgButtonChecked(hwnd, chx1);
    }
    else if (IsDlgButtonChecked(hwnd, rad3))
    {
        std::vector<std::wstring> paths;
        WCHAR szPath1[MAX_PATH], szPath2[MAX_PATH];
        std::wstring path1, path2;
        GetDlgItemTextW(hwnd, cmb3, szPath1, MAX_PATH);
        GetDlgItemTextW(hwnd, cmb4, szPath2, MAX_PATH);
        path1 = szPath1;
        path2 = szPath2;
        trim(path1);
        trim(path2);
        if (path1.size() == 0)
        {
            SetFocus(GetDlgItem(hwnd, cmb3));
            MessageBoxW(hwnd, LoadStringDx1(IDS_INPUTPATH), NULL, MB_ICONERROR);
            return FALSE;
        }
        LoadPathSettings(paths);
        if (path2.size() == 0)
        {
            item.type = ITEMTYPE_PATH1;
            item.path1 = path1;
            item.path2.clear();
            AddRecentPath(paths, path1);
            SavePathSettings(paths);
        }
        else
        {
            item.type = ITEMTYPE_PATH2;
            item.path1 = path1;
            item.path2 = path2;
            AddRecentPath(paths, path1);
            AddRecentPath(paths, path2);
            SavePathSettings(paths);
        }
        item.csidl1 = item.csidl2 = 0;
        item.action = 0;
        item.expand = !!IsDlgButtonChecked(hwnd, chx2);
    }
    else if (IsDlgButtonChecked(hwnd, rad4))
    {
        INT i = (INT)SendDlgItemMessageW(hwnd, cmb5, CB_GETCURSEL, 0, 0);
        if (i == 0 || i == CB_ERR)
        {
            SetFocus(GetDlgItem(hwnd, cmb5));
            MessageBoxW(hwnd, LoadStringDx1(IDS_INPUTACTION), NULL, MB_ICONERROR);
            return FALSE;
        }

        COMBOBOXEXITEM cbex;
        ZeroMemory(&cbex, sizeof(cbex));
        cbex.mask = CBEIF_LPARAM;
        cbex.iItem = i;
        cbex.lParam = -1;
        SendDlgItemMessageW(hwnd, cmb1, CBEM_GETITEM, i, (LPARAM)&cbex);
        i = (INT)cbex.lParam;

        item.type = ITEMTYPE_ACTION;
        item.csidl1 = item.csidl2 = 0;
        item.path1.clear();
        item.path2.clear();
        item.action = i;
        item.expand = FALSE;
    }
    else if (IsDlgButtonChecked(hwnd, rad5))
    {
        item.type = ITEMTYPE_SEP;
        item.csidl1 = item.csidl2 = 0;
        item.path1.clear();
        item.path2.clear();
        item.action = 0;
        item.expand = FALSE;
    }
    else
        return FALSE;
    return TRUE;
}

static VOID
Item_OnChangeIcon(HWND hwnd, ITEM& item)
{
    WCHAR szPath[MAX_PATH];
    lstrcpynW(szPath, item.iconpath.c_str(), MAX_PATH);
    if (OpenFileDialog(
        hwnd,
        szPath,
        OFN_DONTADDTORECENT | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
        LoadStringDx1(IDS_SELECTICON),
        L"ico",
        LoadFilterStringDx(IDS_ALLFILES),
        NULL))
    {
        item.iconpath = szPath;
        item.LoadIcon();
        SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
    }
}

static VOID
Item_OnGroup(HWND hwnd, ITEM& item)
{
    EnableWindow(GetDlgItem(hwnd, chx1), FALSE);
    EnableWindow(GetDlgItem(hwnd, chx2), FALSE);
    item.type = ITEMTYPE_GROUP;
    item.LoadIcon();
    SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
}

static VOID
Item_OnSpecialFolder(HWND hwnd, ITEM& item)
{
    EnableWindow(GetDlgItem(hwnd, chx1), TRUE);
    EnableWindow(GetDlgItem(hwnd, chx2), FALSE);
    SetFocus(GetDlgItem(hwnd, cmb1));
    item.type = ITEMTYPE_CSIDL1;
    item.LoadIcon();
    SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
}

static VOID
Item_OnCSIDL1(HWND hwnd, ITEM& item)
{
    EnableWindow(GetDlgItem(hwnd, chx1), TRUE);
    EnableWindow(GetDlgItem(hwnd, chx2), FALSE);
    CheckRadioButton(hwnd, rad1, rad5, rad2);

    INT i1 = (INT)SendDlgItemMessageW(hwnd, cmb1, CB_GETCURSEL, 0, 0);
    COMBOBOXEXITEM cbex;
    ZeroMemory(&cbex, sizeof(cbex));
    cbex.mask = CBEIF_LPARAM;
    cbex.iItem = i1;
    cbex.lParam = -1;
    SendDlgItemMessageW(hwnd, cmb1, CBEM_GETITEM, i1, (LPARAM)&cbex);
    i1 = (INT)cbex.lParam;

    if (i1 != -1)
    {
        if (item.type != ITEMTYPE_CSIDL2)
            item.type = ITEMTYPE_CSIDL1;
        item.csidl1 = i1;

        item.LoadIcon();
        SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
    }
}

static VOID
Item_OnCSIDL2(HWND hwnd, ITEM& item)
{
    EnableWindow(GetDlgItem(hwnd, chx1), TRUE);
    EnableWindow(GetDlgItem(hwnd, chx2), FALSE);
    CheckRadioButton(hwnd, rad1, rad5, rad2);

    INT i2 = (INT)SendDlgItemMessageW(hwnd, cmb2, CB_GETCURSEL, 0, 0);
    COMBOBOXEXITEM cbex;
    ZeroMemory(&cbex, sizeof(cbex));
    cbex.mask = CBEIF_LPARAM;
    cbex.iItem = i2;
    cbex.lParam = -1;
    SendDlgItemMessageW(hwnd, cmb2, CBEM_GETITEM, i2, (LPARAM)&cbex);
    i2 = (INT)cbex.lParam;

    if (i2 == -1 || i2 == CB_ERR)
    {
        item.type = ITEMTYPE_CSIDL1;
        item.csidl2 = 0;
    }
    else
    {
        item.type = ITEMTYPE_CSIDL2;
        item.csidl2 = i2;
    }

    item.LoadIcon();
    SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
}

static VOID
Item_OnPath(HWND hwnd, ITEM& item)
{
    EnableWindow(GetDlgItem(hwnd, chx1), FALSE);
    EnableWindow(GetDlgItem(hwnd, chx2), TRUE);
    SetFocus(GetDlgItem(hwnd, cmb3));

    if (item.type != ITEMTYPE_PATH2)
        item.type = ITEMTYPE_PATH1;
    item.LoadIcon();
    SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
}

static VOID
Item_OnPath1Combo(HWND hwnd, ITEM& item)
{
    EnableWindow(GetDlgItem(hwnd, chx1), FALSE);
    EnableWindow(GetDlgItem(hwnd, chx2), TRUE);

    INT i1 = (INT)SendDlgItemMessageW(hwnd, cmb3, CB_GETCURSEL, 0, 0);
    if (i1 != -1)
    {
        WCHAR szPath1[MAX_PATH];
        SendDlgItemMessageW(hwnd, cmb3, CB_GETLBTEXT, i1, (LPARAM)szPath1);
        if (item.type != ITEMTYPE_PATH2)
            item.type = ITEMTYPE_PATH1;
        item.path1 = szPath1;
        trim(item.path1);
        CheckRadioButton(hwnd, rad1, rad5, rad3);
        item.LoadIcon();
        SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
    }
}

static VOID
Item_OnPath2Combo(HWND hwnd, ITEM& item)
{
    EnableWindow(GetDlgItem(hwnd, chx1), FALSE);
    EnableWindow(GetDlgItem(hwnd, chx2), TRUE);

    INT i2 = (INT)SendDlgItemMessageW(hwnd, cmb4, CB_GETCURSEL, 0, 0);
    if (i2 != -1)
    {
        WCHAR szPath2[MAX_PATH];
        SendDlgItemMessageW(hwnd, cmb4, CB_GETLBTEXT, i2, (LPARAM)szPath2);
        if (item.type != ITEMTYPE_PATH2)
            item.type = ITEMTYPE_PATH1;
        item.path2 = szPath2;
        trim(item.path2);
    }
    else
    {
        item.type = ITEMTYPE_PATH1;
    }
    CheckRadioButton(hwnd, rad1, rad5, rad3);
    item.LoadIcon();
    SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
}

static VOID
Item_OnBrowsePath1(HWND hwnd, ITEM& item)
{
    WCHAR szPath[MAX_PATH];
    lstrcpynW(szPath, item.path1.c_str(), MAX_PATH);
    if (OpenFileDialog(
        hwnd,
        szPath,
        OFN_DONTADDTORECENT | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
        LoadStringDx1(IDS_SELECTFILE),
        NULL, NULL, NULL))
    {
        CheckRadioButton(hwnd, rad1, rad5, rad3);
        if (item.type != ITEMTYPE_PATH2)
            item.type = ITEMTYPE_PATH1;
        item.path1 = szPath;
        SetDlgItemTextW(hwnd, cmb3, szPath);
        EnableWindow(GetDlgItem(hwnd, chx1), FALSE);
        EnableWindow(GetDlgItem(hwnd, chx2), TRUE);
        item.LoadIcon();
        SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
    }
}

static VOID
Item_OnBrowsePath2(HWND hwnd, ITEM& item)
{
    WCHAR szPath[MAX_PATH];
    lstrcpynW(szPath, item.path2.c_str(), MAX_PATH);
    if (OpenFileDialog(
        hwnd,
        szPath,
        OFN_DONTADDTORECENT | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY,
        LoadStringDx1(IDS_SELECTFILE),
        NULL, NULL, NULL))
    {
        CheckRadioButton(hwnd, rad1, rad5, rad3);
        item.type = ITEMTYPE_PATH2;
        item.path2 = szPath;
        SetDlgItemTextW(hwnd, cmb4, szPath);
        EnableWindow(GetDlgItem(hwnd, chx1), FALSE);
        EnableWindow(GetDlgItem(hwnd, chx2), TRUE);
        item.LoadIcon();
        SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
    }
}

static VOID
Item_OnAction(HWND hwnd, ITEM& item)
{
    EnableWindow(GetDlgItem(hwnd, chx1), FALSE);
    EnableWindow(GetDlgItem(hwnd, chx2), FALSE);
    SetFocus(GetDlgItem(hwnd, cmb5));
    item.type = ITEMTYPE_ACTION;
    item.LoadIcon();
    SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
}

static VOID
Item_OnActionCombo(HWND hwnd, ITEM& item)
{
    EnableWindow(GetDlgItem(hwnd, chx1), FALSE);
    EnableWindow(GetDlgItem(hwnd, chx2), FALSE);
    item.type = ITEMTYPE_ACTION;
    item.LoadIcon();
    SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
}

static VOID
Item_OnSep(HWND hwnd, ITEM& item)
{
    EnableWindow(GetDlgItem(hwnd, chx1), FALSE);
    EnableWindow(GetDlgItem(hwnd, chx2), FALSE);
    item.type = ITEMTYPE_SEP;
    item.LoadIcon();
    SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
}

static VOID
Item_OnNoIcon(HWND hwnd, ITEM& item)
{
    item.iconpath.clear();
    item.LoadIcon();
    SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
}

static VOID
Item_OnDropFiles(HWND hwnd, ITEM& item, HDROP hDrop)
{
    WCHAR szPath[MAX_PATH], szTitle[MAX_PATH];
    DragQueryFileW(hDrop, 0, szPath, MAX_PATH);
    DragFinish(hDrop);

    DWORD attrs = GetFileAttributesW(szPath);
    if (attrs == 0xFFFFFFFF)
    {
        MessageBeep(0xFFFFFFFF);
        return;
    }
    GetFileTitleW(szPath, szTitle, MAX_PATH);

    SetDlgItemTextW(hwnd, edt1, szTitle);
    CheckRadioButton(hwnd, rad1, rad5, rad3);
    item.type = ITEMTYPE_PATH1;
    item.text = szTitle;
    item.path1 = szPath;
    item.path2.clear();
    item.expand = !!(attrs & FILE_ATTRIBUTE_DIRECTORY);
    CheckDlgButton(hwnd, chx2, (item.expand ? BST_CHECKED : BST_UNCHECKED));
    SetDlgItemTextW(hwnd, cmb3, szPath);
    SetDlgItemTextW(hwnd, cmb4, NULL);
    EnableWindow(GetDlgItem(hwnd, chx1), FALSE);
    EnableWindow(GetDlgItem(hwnd, chx2), TRUE);
    item.LoadIcon();
    SendDlgItemMessageW(hwnd, ico1, STM_SETICON, (WPARAM)item.hIcon, 0);
}

INT_PTR CALLBACK
Item_DlgProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    static ITEM *pItem;
    switch(uMsg)
    {
    case WM_INITDIALOG:
        pItem = (ITEM *)lParam;
        Item_OnInitDialog(hwnd, *pItem);
        DragAcceptFiles(hwnd, TRUE);
        return TRUE;

    case WM_DROPFILES:
        Item_OnDropFiles(hwnd, *pItem, (HDROP)wParam);
        break;

    case WM_COMMAND:
        switch(LOWORD(wParam))
        {
        case psh1:  // change icon
            Item_OnChangeIcon(hwnd, *pItem);
            break;

        case rad1:  // group
            Item_OnGroup(hwnd, *pItem);
            break;

        case rad2:  // special folder
            Item_OnSpecialFolder(hwnd, *pItem);
            break;

        case cmb1:  // csidl1
            if (HIWORD(wParam) == CBN_SELCHANGE)
                Item_OnCSIDL1(hwnd, *pItem);
            break;

        case cmb2:  // csidl2
            if (HIWORD(wParam) == CBN_SELCHANGE)
                Item_OnCSIDL2(hwnd, *pItem);
            break;

        case rad3:  // path
            Item_OnPath(hwnd, *pItem);
            break;

        case cmb3:  // path1 combo
            if (HIWORD(wParam) == CBN_SELCHANGE)
                Item_OnPath1Combo(hwnd, *pItem);
            break;

        case psh2:  // browse path1
            Item_OnBrowsePath1(hwnd, *pItem);
            break;

        case cmb4:  // path2 combo
            if (HIWORD(wParam) == CBN_SELCHANGE)
                Item_OnPath2Combo(hwnd, *pItem);
            break;

        case psh3:  // browse path2
            Item_OnBrowsePath2(hwnd, *pItem);
            break;

        case rad4:  // action
            Item_OnAction(hwnd, *pItem);
            break;

        case cmb5:  // action combo
            Item_OnActionCombo(hwnd, *pItem);
            break;

        case rad5:  // separator
            Item_OnSep(hwnd, *pItem);
            break;

        case psh4:  // no icon
            Item_OnNoIcon(hwnd, *pItem);
            break;

        case IDOK:
            {
                ITEM item(*pItem);
                if (Item_SetData(hwnd, item))
                {
                    *pItem = item;
                    EndDialog(hwnd, IDOK);
                }
            }
            break;

        case IDCANCEL:
            EndDialog(hwnd, IDCANCEL);
            break;
        }
    }
    return 0;
}
