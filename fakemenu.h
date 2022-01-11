#include <stdio.h>
#include <string>

extern HINSTANCE g_hInstance;

struct FAKEMENU
{
    HWND                    hwndOwner;
    HWND                    hBaseBar;
    HWND                    hPager;
    HWND                    hToolBar;
    INT                     cyItem;
    UINT                    uID;
    UINT                    uFlags;
    FAKEMENU               *parent;
    UINT                    uSelected;
    UINT                    uPopupped;
    FAKEMENU               *popupped;
    std::wstring            text;
    std::vector<FAKEMENU *> children;
    HICON                   hIcon;
    LPITEMIDLIST            pidl;
    INT                     nIconSize;
    SIZE                    sizToolBar;
    BOOL                    ThreadCancel;
    BOOL                    NoSort;
    HANDLE                  hIconThread;
    BOOL                    CustomIcon;
    BOOL                    Recent;
};
typedef FAKEMENU *HFAKEMENU;

BOOL APIENTRY
RegisterFakeMenu(HINSTANCE hInstance);

HFAKEMENU APIENTRY
CreatePopupFakeMenu(VOID);

HFAKEMENU APIENTRY
GetSubFakeMenu(HFAKEMENU hMenu, INT nPos);

BOOL APIENTRY
DestroyFakeMenu(HFAKEMENU hMenu);

INT APIENTRY
GetFakeMenuItemCount(HFAKEMENU hMenu);

UINT APIENTRY
GetFakeMenuItemID(HFAKEMENU hMenu, INT nPos);

BOOL APIENTRY
AppendFakeMenu(HFAKEMENU hMenu,
               UINT uFlags,
               UINT_PTR uIDNewItem,
               LPCWSTR lpNewItem,
               LPITEMIDLIST pidl = NULL,
               HICON hIcon = NULL);

BOOL APIENTRY
TrackPopupFakeMenuNoLock(HFAKEMENU hMenu,
                         UINT uFlags,
                         INT x, INT y,
                         INT nReserved,
                         HWND hwndOwner,
                         LPCRECT prc);

#define MF_FAKEMENU MF_END

#define MAX_RECENT 15

#define WM_FAKEMENUSELECT (WM_USER + 888)
#define WM_CLOSEBASEBAR (WM_USER + 999)
