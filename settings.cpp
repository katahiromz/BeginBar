#include <windows.h>
#include <commctrl.h>
#include <shlobj.h>
#include <tchar.h>
#include <string>
#include <vector>

#include "fakemenu.h"
#include "BeginBar.h"
#include "resource.h"

INT g_nWindowX = CW_USEDEFAULT, g_nWindowY = CW_USEDEFAULT;
INT g_nWindowCX = CW_USEDEFAULT, g_nWindowCY = CW_USEDEFAULT;

std::vector<BEGINBUTTON> g_buttons;

BOOL g_bVertical = FALSE;

static const WCHAR
    s_szBeginBarKey[]  = L"Software\\Katayama Hirofumi MZ\\BeginBar";

static const WCHAR
    s_szWindowX[]                   = L"WindowX",
    s_szWindowY[]                   = L"WindowY",
    s_szWindowCX[]                  = L"WindowCX",
    s_szWindowCY[]                  = L"WindowCY",
    s_szButtonCount[]               = L"ButtonCount",
    s_szVertical[]                  = L"Vertical",
    s_szButtonItemCount[]           = L"Button%dItemCount",
    s_szButtonItemType[]            = L"Button%dItem%dType",
    s_szButtonItemText[]            = L"Button%dItem%dText",
    s_szButtonItemPath1[]           = L"Button%dItem%dPath1",
    s_szButtonItemPath2[]           = L"Button%dItem%dPath2",
    s_szButtonItemCSIDL1[]          = L"Button%dItem%dCSIDL1",
    s_szButtonItemCSIDL2[]          = L"Button%dItem%dCSIDL2",
    s_szButtonItemExpand[]          = L"Button%dItem%dExpand",
    s_szButtonItemLevel[]           = L"Button%dItem%dLevel",
    s_szButtonItemAction[]          = L"Button%dItem%dAction",
    s_szButtonItemIconPath[]        = L"Button%dItem%dIconPath";

static const WCHAR
    s_szPathCount[]                 = L"PathCount",
    s_szPath[]                      = L"Path%d";

static BOOL
CreateKey(HKEY key, LPCWSTR subkey, HKEY *newkey)
{
    LONG result;
    result = RegCreateKeyExW(key, subkey, 0, NULL, 0, KEY_WRITE, NULL, newkey, NULL);
    return ERROR_SUCCESS == result;
}

static BOOL
QueryLongValue(HKEY key, LPCWSTR name, LONG *value)
{
    DWORD len = sizeof(LONG);
    LONG result;
    result = RegQueryValueExW(key, name, NULL, NULL, (BYTE *)value, &len);
    return ERROR_SUCCESS == result;
}

static BOOL
SetLongValue(HKEY key, LPCWSTR name, LONG *value)
{
    LONG result;
    result = RegSetValueExW(key, name, 0, REG_DWORD, (BYTE *)value, sizeof(LONG));
    return ERROR_SUCCESS == result;
}

static BOOL
QueryStringValue(HKEY key, LPCWSTR name, std::wstring& value)
{
    DWORD len;
    LONG result;
    WCHAR sz[MAX_PATH];
    len = sizeof(sz);
    result = RegQueryValueExW(key, name, NULL, NULL, (BYTE *)sz, &len);
    value = sz;
    return ERROR_SUCCESS == result;
}

static BOOL
SetStringValue(HKEY key, LPCWSTR name, LPCWSTR value)
{
    LONG result;
    DWORD cb = (lstrlenW(value) + 1) * sizeof(WCHAR);
    result = RegSetValueExW(key, name, 0, REG_SZ, (BYTE *)value, cb);
    return ERROR_SUCCESS == result;
}

VOID LoadSettings(VOID)
{
    HKEY hkey;
    LONG value;

    RegOpenKeyExW(HKEY_CURRENT_USER, s_szBeginBarKey, 0, KEY_READ, &hkey);

    if (QueryLongValue(hkey, s_szWindowX, &value))
        g_nWindowX = value;

    if (QueryLongValue(hkey, s_szWindowY, &value))
        g_nWindowY = value;

    if (QueryLongValue(hkey, s_szWindowCX, &value))
        g_nWindowCX = value;

    if (QueryLongValue(hkey, s_szWindowCY, &value))
        g_nWindowCY = value;

    INT nButtonCount;
    if (QueryLongValue(hkey, s_szButtonCount, &value))
        nButtonCount = value;
    else
        nButtonCount = 0;

    if (QueryLongValue(hkey, s_szVertical, &value))
        g_bVertical = value;

    BEGINBUTTON button;
    ITEM item;
    WCHAR szName[MAX_PATH];
    INT i, j, nButtonItemCount;
    for (i = 0; i < nButtonCount; i++)
    {
        wsprintfW(szName, s_szButtonItemCount, i);
        if (QueryLongValue(hkey, szName, &value))
            nButtonItemCount = value;

        for (j = 0; j < nButtonItemCount; j++)
        {
            wsprintfW(szName, s_szButtonItemType, i, j);
            if (QueryLongValue(hkey, szName, &value))
                item.type = (ITEMTYPE)value;
            else
                break;

            wsprintfW(szName, s_szButtonItemText, i, j);
            QueryStringValue(hkey, szName, item.text);

            wsprintfW(szName, s_szButtonItemPath1, i, j);
            QueryStringValue(hkey, szName, item.path1);

            wsprintfW(szName, s_szButtonItemPath2, i, j);
            QueryStringValue(hkey, szName, item.path2);

            wsprintfW(szName, s_szButtonItemCSIDL1, i, j);
            if (QueryLongValue(hkey, szName, &value))
                item.csidl1 = value;

            wsprintfW(szName, s_szButtonItemCSIDL2, i, j);
            if (QueryLongValue(hkey, szName, &value))
                item.csidl2 = value;

            wsprintfW(szName, s_szButtonItemExpand, i, j);
            if (QueryLongValue(hkey, szName, &value))
                item.expand = value;

            wsprintfW(szName, s_szButtonItemLevel, i, j);
            if (QueryLongValue(hkey, szName, &value))
                item.level = value;

            wsprintfW(szName, s_szButtonItemAction, i, j);
            if (QueryLongValue(hkey, szName, &value))
                item.action = value;

            wsprintfW(szName, s_szButtonItemIconPath, i, j);
            QueryStringValue(hkey, szName, item.iconpath);

            button.items.push_back(item);
        }
        if (j != nButtonItemCount)
            break;

        g_buttons.push_back(button);
        button.items.clear();
    }

    if (g_buttons.size() == 0)
    {
        button.items.clear();

        item.type = ITEMTYPE_GROUP;
        item.text = LoadStringDx1(IDS_BEGIN);
        item.path1.clear();
        item.path2.clear();
        item.csidl1 = 0;
        item.csidl2 = 0;
        item.expand = FALSE;
        item.level = 0;
        item.action = 0;
        item.iconpath = GetDefaultIconPath();
        button.items.push_back(item);

        item.iconpath.clear();

        item.type = ITEMTYPE_CSIDL2;
        item.text = LoadStringDx1(IDS_PROGRAMS);
        item.path1.clear();
        item.path2.clear();
        item.csidl1 = CSIDL_COMMON_PROGRAMS;
        item.csidl2 = CSIDL_PROGRAMS;
        item.expand = TRUE;
        item.level = 1;
        item.action = 0;
        button.items.push_back(item);

        item.type = ITEMTYPE_CSIDL1;
        item.text = LoadStringDx1(IDS_FAVORITES);
        item.path1.clear();
        item.path2.clear();
        item.csidl1 = CSIDL_FAVORITES;
        item.csidl2 = 0;
        item.expand = TRUE;
        item.level = 1;
        item.action = 0;
        button.items.push_back(item);

        item.type = ITEMTYPE_CSIDL1;
        item.text = LoadStringDx1(IDS_COMPUTER);
        item.path1.clear();
        item.path2.clear();
        item.csidl1 = CSIDL_DRIVES;
        item.csidl2 = 0;
        item.expand = TRUE;
        item.level = 1;
        item.action = 0;
        button.items.push_back(item);

        item.type = ITEMTYPE_CSIDL1;
        item.text = LoadStringDx1(IDS_DOCUMENTS);
        item.path1.clear();
        item.path2.clear();
        item.csidl1 = CSIDL_PERSONAL;
        item.csidl2 = 0;
        item.expand = FALSE;
        item.level = 1;
        item.action = 0;
        button.items.push_back(item);

        item.type = ITEMTYPE_CSIDL1;
        item.text = LoadStringDx1(IDS_RECENTS);
        item.path1.clear();
        item.path2.clear();
        item.csidl1 = CSIDL_RECENT;
        item.csidl2 = 0;
        item.expand = TRUE;
        item.level = 1;
        item.action = 0;
        button.items.push_back(item);

        item.type = ITEMTYPE_CSIDL1;
        item.text = LoadStringDx1(IDS_CONTROLPANEL);
        item.path1.clear();
        item.path2.clear();
        item.csidl1 = CSIDL_CONTROLS;
        item.csidl2 = 0;
        item.expand = FALSE;
        item.level = 1;
        item.action = 0;
        button.items.push_back(item);

        item.type = ITEMTYPE_SEP;
        item.text.clear();
        item.path1.clear();
        item.path2.clear();
        item.csidl1 = 0;
        item.csidl2 = 0;
        item.expand = FALSE;
        item.level = 1;
        item.action = 0;
        button.items.push_back(item);

        item.type = ITEMTYPE_ACTION;
        item.text = LoadStringDx1(IDS_EXITBEGINBAR);
        item.path1.clear();
        item.path2.clear();
        item.csidl1 = 0;
        item.csidl2 = 0;
        item.expand = FALSE;
        item.level = 1;
        item.action = ID_EXITAPP;
        button.items.push_back(item);

        g_buttons.push_back(button);
    }

    RegCloseKey(hkey);
}

VOID SaveSettings(VOID)
{
    HKEY hkey;
    LONG value;

    CreateKey(HKEY_CURRENT_USER, s_szBeginBarKey, &hkey);

    value = g_nWindowX;
    SetLongValue(hkey, s_szWindowX, &value);

    value = g_nWindowY;
    SetLongValue(hkey, s_szWindowY, &value);

    value = g_nWindowCX;
    SetLongValue(hkey, s_szWindowCX, &value);

    value = g_nWindowCY;
    SetLongValue(hkey, s_szWindowCY, &value);

    INT nButtonCount = (INT)g_buttons.size();

    INT i, j, nButtonItemCount;
    ITEM item;
    WCHAR szName[128];
    for (i = 0; i < nButtonCount; i++)
    {
        nButtonItemCount = (INT)g_buttons[i].items.size();
        wsprintfW(szName, s_szButtonItemCount, i);
        value = nButtonItemCount;
        SetLongValue(hkey, szName, &value);

        for (j = 0; j < nButtonItemCount; j++)
        {
            item = g_buttons[i].items[j];
            wsprintfW(szName, s_szButtonItemType, i, j);
            value = item.type;
            if (!SetLongValue(hkey, szName, &value))
                break;

            wsprintfW(szName, s_szButtonItemText, i, j);
            SetStringValue(hkey, szName, item.text.c_str());

            wsprintfW(szName, s_szButtonItemPath1, i, j);
            SetStringValue(hkey, szName, item.path1.c_str());

            wsprintfW(szName, s_szButtonItemPath2, i, j);
            SetStringValue(hkey, szName, item.path2.c_str());

            wsprintfW(szName, s_szButtonItemCSIDL1, i, j);
            value = item.csidl1;
            SetLongValue(hkey, szName, &value);

            wsprintfW(szName, s_szButtonItemCSIDL2, i, j);
            value = item.csidl2;
            SetLongValue(hkey, szName, &value);

            wsprintfW(szName, s_szButtonItemExpand, i, j);
            value = item.expand;
            SetLongValue(hkey, szName, &value);

            wsprintfW(szName, s_szButtonItemLevel, i, j);
            value = item.level;
            SetLongValue(hkey, szName, &value);

            wsprintfW(szName, s_szButtonItemAction, i, j);
            value = item.action;
            SetLongValue(hkey, szName, &value);

            wsprintfW(szName, s_szButtonItemIconPath, i, j);
            SetStringValue(hkey, szName, item.iconpath.c_str());
        }
        if (j != nButtonItemCount)
            break;
    }

    value = nButtonCount;
    SetLongValue(hkey, s_szButtonCount, &value);

    value = g_bVertical;
    SetLongValue(hkey, s_szVertical, &value);

    RegCloseKey(hkey);
}

BOOL LoadPathSettings(std::vector<std::wstring>& paths)
{
    HKEY hkey;
    LONG value;
    INT i, nPathCount;
    WCHAR szName[MAX_PATH];
    std::wstring path;
    paths.clear();

    RegOpenKeyExW(HKEY_CURRENT_USER, s_szBeginBarKey, 0, KEY_READ, &hkey);

    if (QueryLongValue(hkey, s_szPathCount, &value))
        nPathCount = value;
    else
        nPathCount = 0;

    for (i = 0; i < nPathCount; i++)
    {
        wsprintfW(szName, s_szPath, i);
        if (!QueryStringValue(hkey, szName, path))
            break;
        paths.push_back(path);
    }

    return paths.size() > 0;
}

VOID SavePathSettings(const std::vector<std::wstring>& paths)
{
    HKEY hkey;
    LONG value;
    INT i, nPathCount;
    WCHAR szName[MAX_PATH];

    CreateKey(HKEY_CURRENT_USER, s_szBeginBarKey, &hkey);

    nPathCount = (INT)paths.size();
    for (i = 0; i < nPathCount; i++)
    {
        wsprintfW(szName, s_szPath, i);
        if (!SetStringValue(hkey, szName, paths[i].c_str()))
        {
            nPathCount = i;
            break;
        }
    }

    value = nPathCount;
    SetLongValue(hkey, s_szPathCount, &value);
}


VOID AddRecentPath(std::vector<std::wstring>& paths, const std::wstring& path)
{
    std::vector<std::wstring>::iterator it, end = paths.end();
    for (it = paths.begin(); it != end; it++)
    {
        if (lstrcmpiW(it->c_str(), path.c_str()) == 0)
        {
            paths.erase(it);
            break;
        }
    }
    paths.insert(paths.begin(), path);
}

LPWSTR GetDefaultIconPath(VOID)
{
    static WCHAR s_szPath[MAX_PATH];
    GetModuleFileNameW(NULL, s_szPath, MAX_PATH);
    return s_szPath;
}
