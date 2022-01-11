#pragma once

extern HINSTANCE g_hInstance;
extern HWND g_hMainWnd;
extern HWND g_hPager;
extern INT g_iButton;

enum ITEMTYPE
{
    ITEMTYPE_GROUP,
    ITEMTYPE_SEP,
    ITEMTYPE_CSIDL1,
    ITEMTYPE_CSIDL2,
    ITEMTYPE_PATH1,
    ITEMTYPE_PATH2,
    ITEMTYPE_ACTION
};

struct ITEM
{
    ITEMTYPE type;
    std::wstring text;
    std::wstring path1;
    std::wstring path2;
    INT csidl1;
    INT csidl2;
    BOOL expand;
    INT level;
    INT action;
    std::wstring iconpath;
    HICON hIcon;

    ITEM() : hIcon(NULL)
    {
    }

    ITEM(const ITEM& item) :
        type(item.type),
        text(item.text),
        path1(item.path1),
        path2(item.path2),
        csidl1(item.csidl1),
        csidl2(item.csidl2),
        expand(item.expand),
        level(item.level),
        action(item.action),
        iconpath(item.iconpath),
        hIcon(CopyIcon(item.hIcon))
    {
    }

    ~ITEM()
    {
        if (hIcon)
            DestroyIcon(hIcon);
    }

    VOID SetIcon(HICON hIcon_)
    {
        if (hIcon)
            DestroyIcon(hIcon);
        hIcon = CopyIcon(hIcon_);
    }

    ITEM& operator=(const ITEM& item)
    {
        type = item.type;
        text = item.text;
        path1 = item.path1;
        path2 = item.path2;
        csidl1 = item.csidl1;
        csidl2 = item.csidl2;
        expand = item.expand;
        level = item.level;
        action = item.action;
        iconpath = item.iconpath;
        SetIcon(item.hIcon);
        return *this;
    }

    BOOL LoadIcon();
};

struct BEGINBUTTON
{
    HWND hButton;
    std::vector<ITEM> items;
    INT cxIndentsAndText, cxContent, cxExpand;

    VOID LoadIcons();
};

struct DATA
{
    std::wstring path;
    std::wstring name;
    bool is_folder;
    LPITEMIDLIST pidl;

    DATA()
    {
        pidl = NULL;
    }
    DATA(const DATA& data)
    {
        path = data.path;
        name = data.name;
        is_folder = data.is_folder;
        pidl = ILClone(data.pidl);
    }
    DATA& operator=(const DATA& data)
    {
        path = data.path;
        name = data.name;
        is_folder = data.is_folder;
        pidl = ILClone(data.pidl);
        return *this;
    }
    ~DATA()
    {
        CoTaskMemFree(pidl);
    }
};

HRESULT GetFileList(LPITEMIDLIST pidl, std::vector<DATA>& list);
BOOL GetPathOfShortcut(HWND hWnd, LPCWSTR pszLnkFile, LPWSTR pszPath);
DWORD AddItemsToMenu(HFAKEMENU hMenu, LPITEMIDLIST pidl);
VOID FitWindowInWorkArea(HWND hwnd);
BOOL OpenFileDialog(
    HWND hwndParent,
    LPWSTR pszFile,
    DWORD dwOFN_flags,
    LPCWSTR pszTitle,
    LPCWSTR pszDefExt,
    LPCWSTR pszFilter,
    LPCWSTR pszInitialDir);
LPWSTR LoadStringDx1(UINT uID);
LPWSTR LoadStringDx2(UINT uID);
LPWSTR LoadFilterStringDx(UINT uID);
void trim(std::wstring& str);
VOID GetWorkAreaNearWindow(HWND hwnd, LPRECT prcWorkArea);

// settings.cpp
extern INT g_nWindowX, g_nWindowY;
extern INT g_nWindowCX, g_nWindowCY;
extern std::vector<BEGINBUTTON> g_buttons;
extern BOOL g_bVertical;
VOID LoadSettings(VOID);
VOID SaveSettings(VOID);
BOOL LoadPathSettings(std::vector<std::wstring>& paths);
VOID SavePathSettings(const std::vector<std::wstring>& paths);
VOID AddRecentPath(std::vector<std::wstring>& paths, const std::wstring& path);
LPWSTR GetDefaultIconPath(VOID);

// tray.cpp
extern HWND g_hTray;
#define MYWM_UPDATEUI (WM_USER + 200)
BOOL RegisterButtonTray(HINSTANCE hInstance);
HWND CreateButtonTray(HWND hwnd);

// dialogs.cpp
INT_PTR CALLBACK
Buttons_DlgProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK
Menu_DlgProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK
Item_DlgProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
