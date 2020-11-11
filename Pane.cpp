#include "Pane.h"
#include "Resource.h"


extern HINSTANCE hDllInst;
extern TCHAR szSetupKey[];
extern TCHAR szImsspKey[];
extern HWND hStartWnd;
extern HWND hRebarWnd;
extern HWND hTaskbarWnd;
extern HWND hSyslistWnd;
DWORD dwDisableMetro = 0;

HWND hPaneWnd = NULL;
HWND hForeWnd = NULL;
HICON hTabsIcon = NULL;
HICON hProgsIcon = NULL;
HFONT hFont = NULL;
TCHAR szProgs[25][50];
INT cchProgs[25];
INT nTab = 4;
INT nMouseMove = 0;
INT nLButtonDown = 0;
BYTE byFunc[10];
DWORD dwDisableRibbon = 0;
DWORD dwUseSmallFont = 0;

RECT rcPane[24] = {
	  0,   0,   0,   0,
	  0,   0, 300, 480,//1

	  0,   0,  48, 480,
	 48,   0, 300, 480,
	  0,   0,  48,  48,
	  0,  48,  48,  96,
	  0,  96,  48, 144,//6

	 50,   2, 298,  38,
	 50,  38, 298,  74,
	 50,  74, 298, 110,
	 50, 110, 298, 146,
	 50, 146, 298, 182,
	 50, 182, 298, 218,
	 50, 218, 298, 254,
	 50, 254, 298, 290,
	 50, 290, 298, 326,
	 50, 326, 298, 362,
	 50, 362, 298, 398,
	 50, 398, 298, 434,
	 50, 434,  50, 470,//19

	 48, 444, 300, 480,
	288,   0, 300, 444,//21

	 48,  36, 144, 432,
	 52,   0,  84,   0,//23
};


// Pane
static LRESULT CALLBACK PaneProc(HWND, UINT, WPARAM, LPARAM);
void OnPanePaint();
void OnPaneMouseMove(LPARAM);
void OnPaneLButtonDown(LPARAM);
void OnPaneLButtonUp(LPARAM);
void OpenExecute(BOOL, WPARAM);

// NewFunc
HMODULE WINAPI NewFunc(LPCWSTR, HANDLE, DWORD);
void Hook(PBYTE, PBYTE, PBYTE, DWORD);


HRESULT CreatePane()
{
	INT i;
	DWORD cb;
	TCHAR sz[2];
	WNDCLASSEX wcex;

	cb = sizeof(sz);
	RegGetValue(HKEY_CURRENT_USER, szImsspKey, NULL, RRF_RT_ANY, NULL, &sz, &cb);
	if (!lstrcmp(sz, TEXT("\\")))
		dwDisableMetro = 1;

	//float f=1.25;for(i=0;i<24;i++){rcPane[i].left*=f;rcPane[i].top*=f;rcPane[i].right*=f;rcPane[i].bottom*=f;}
	for (i=0; i<25; i++)
		cchProgs[i] = LoadString(hDllInst, i + IDS_PROG1, szProgs[i], 50);

	hTabsIcon = (HICON)LoadImage(hDllInst, MAKEINTRESOURCE(IDI_TABS),
		IMAGE_ICON, (int)rcPane[22].left, (int)rcPane[22].right, LR_DEFAULTCOLOR);
	if (!hTabsIcon)
		return E_FAIL;

	hProgsIcon = (HICON)LoadImage(hDllInst, MAKEINTRESOURCE(IDI_PROGS),
		IMAGE_ICON, (int)rcPane[22].top, (int)rcPane[22].bottom, LR_DEFAULTCOLOR);
	if (!hProgsIcon)
		return E_FAIL;

	wcex.cbClsExtra		= 0;
	wcex.cbSize			= sizeof(WNDCLASSEX);
	wcex.cbWndExtra		= 0;
	wcex.hbrBackground	= NULL;
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hIcon			= NULL;
	wcex.hIconSm		= NULL;
	wcex.hInstance		= hDllInst;
	wcex.lpfnWndProc	= PaneProc;
	wcex.lpszClassName	= TEXT("Pane");
	wcex.lpszMenuName	= NULL;
	wcex.style			= CS_VREDRAW | CS_HREDRAW;
	if (!RegisterClassEx(&wcex))
		return E_FAIL;

	hPaneWnd = CreateWindowEx(WS_EX_TOOLWINDOW | WS_EX_TOPMOST, TEXT("Pane"),
		NULL, WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		0, 0, 0, 0, hTaskbarWnd, NULL, hDllInst, NULL);
	if (!hPaneWnd)
		return E_FAIL;
	
	cb = sizeof(DWORD);
	RegGetValue(HKEY_CURRENT_USER, szSetupKey, TEXT("UseSmallFont"),
		RRF_RT_ANY, NULL, &dwUseSmallFont, &cb);

	cb = sizeof(DWORD);
	RegGetValue(HKEY_CURRENT_USER, szSetupKey, TEXT("DisableRibbon"),
		RRF_RT_ANY, NULL, &dwDisableRibbon, &cb);

	if (dwDisableRibbon)
		Hook((PBYTE)NewFunc, (PBYTE)LoadLibraryExW, byFunc, 1);

	return S_OK;
}

void ShowPane()
{
	RECT rc;
	LONG x, y;
	int w, h;
	HWND hwnd;

	if (hPaneWnd)
	{
		if (IsWindowVisible(hPaneWnd))
		{
			ShowWindow(hPaneWnd, SW_HIDE);
			SetForegroundWindow(hForeWnd);
			return;
		}

		x = y = 0;
		GetWindowRect(hTaskbarWnd, &rc);
		if (rc.left <= 0)
		{
			if (rc.top <= 0)
			{
				if (rc.right > rc.bottom) //taskbar on top
					y = rc.bottom;
				else //left
					x = rc.right;
			}
			else //bottom
				y = rc.top - rcPane[1].bottom;
		}
		else //right
			x = rc.left - rcPane[1].right;
		w = (int)rcPane[1].right;
		h = (int)rcPane[1].bottom;
		MoveWindow(hPaneWnd, (int)x, (int)y, w, h, TRUE);

		hwnd = hRebarWnd;
		if (dwUseSmallFont && hSyslistWnd)
			hwnd = hSyslistWnd;
		hFont = (HFONT)SendMessage(hwnd, WM_GETFONT, 0, 0);

		hForeWnd = GetForegroundWindow();
		ShowWindow(hPaneWnd, SW_SHOW);
		SetForegroundWindow(hPaneWnd);
	}
}

void ClosePane()
{
	if (dwDisableRibbon)
	{
		Hook((PBYTE)NewFunc, (PBYTE)LoadLibraryExW, byFunc, 0);
		dwDisableRibbon = 0;
	}

	if (hPaneWnd)
	{
		DestroyWindow(hPaneWnd);
		hPaneWnd = NULL;
	}

	if (hProgsIcon)
	{
		DestroyIcon(hProgsIcon);
		hProgsIcon = NULL;
	}

	if (hTabsIcon)
	{
		DestroyIcon(hTabsIcon);
		hTabsIcon = NULL;
	}
	
	UnregisterClass(TEXT("Pane"), hDllInst);
}

//
// Pane
//
static LRESULT CALLBACK PaneProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_PAINT:
		OnPanePaint();
		break;
	case WM_KILLFOCUS:
		nTab = 4;
		nMouseMove = 0;
		nLButtonDown = 0;
		ShowWindow(hWnd, SW_HIDE);
		break;
	case WM_SHOWWINDOW:
		EnableWindow(hStartWnd, !wParam);
		break;
	case WM_MOUSELEAVE:
		nMouseMove = 0;
		nLButtonDown = 0;
		InvalidateRect(hWnd, NULL, TRUE);
		break;
	case WM_MOUSEMOVE:
		OnPaneMouseMove(lParam);
		break;
	case WM_LBUTTONDOWN:
		OnPaneLButtonDown(lParam);
		break;
	case WM_LBUTTONUP:
		OnPaneLButtonUp(lParam);
		break;
	case WM_KEYDOWN:
		OpenExecute(FALSE, wParam);
		break;
	case WM_KEYUP:
		OpenExecute(TRUE, wParam);
		break;
	default:
		return DefWindowProc(hWnd, uMsg, wParam, lParam);
	}

	return 0;
}

void OnPanePaint()
{
	HDC hdc;
	PAINTSTRUCT ps;
	HDC hbufdc;
	HPAINTBUFFER hbufpt;
	int i, x, y, cx, cy;
	UINT u;

	hdc = BeginPaint(hPaneWnd, &ps);
	if (hdc)
	{
		hbufpt = BeginBufferedPaint(hdc, &rcPane[1], BPBF_TOPDOWNDIB, NULL, &hbufdc);
		if (hbufpt)
		{
			FillRect(hbufdc, &rcPane[2], (HBRUSH)(COLOR_3DDKSHADOW + 1));
			FillRect(hbufdc, &rcPane[3], (HBRUSH)(COLOR_BTNSHADOW + 1));
			if (nLButtonDown > 3)
				FillRect(hbufdc, &rcPane[nLButtonDown], (HBRUSH)(COLOR_SCROLLBAR + 1));
			else if (nMouseMove > 3)
				FillRect(hbufdc, &rcPane[nMouseMove], (HBRUSH)(COLOR_ACTIVEBORDER + 1));
			FillRect(hbufdc, &rcPane[nTab], (HBRUSH)(COLOR_BTNSHADOW + 1));

			cx = (int)rcPane[22].left;
			cy = (int)rcPane[22].right;
			DrawIconEx(hbufdc, 0, 0, hTabsIcon, cx, cy, 0, NULL, DI_NORMAL);

			SetBkMode(hbufdc, TRANSPARENT);
			SelectObject(hbufdc, hFont);
			u = DT_SINGLELINE | DT_VCENTER;
			if (nTab == 6)
			{
				u |= DT_CENTER;
				DrawText(hbufdc, szProgs[0], cchProgs[0], &rcPane[7], u);
				DrawText(hbufdc, szProgs[1], cchProgs[1], &rcPane[8], u);
				DrawText(hbufdc, szProgs[2], cchProgs[2], &rcPane[9], u);
				DrawText(hbufdc, szProgs[3], cchProgs[3], &rcPane[10], u);
				DrawText(hbufdc, szProgs[3], cchProgs[3], &rcPane[20], u);
			}
			else if (nTab == 5)
			{
				i = dwDisableMetro ? 19 : 20;
				DrawText(hbufdc, szProgs[i], cchProgs[i], &rcPane[7], u);
				i = dwDisableRibbon ? 21 : 22;
				DrawText(hbufdc, szProgs[i], cchProgs[i], &rcPane[8], u);
				i = dwUseSmallFont ? 23 : 24;
				DrawText(hbufdc, szProgs[i], cchProgs[i], &rcPane[9], u);
				u |= DT_CENTER;
				DrawText(hbufdc, szProgs[4], cchProgs[4], &rcPane[20], u);
			}
			else
			{
				cx = rcPane[7].bottom - rcPane[7].top;
				cy = cx - rcPane[7].top - rcPane[7].top;
				for(i=7; i<19; i++)
				{
					rcPane[23].top = rcPane[i].top + rcPane[7].top;
					rcPane[23].bottom = rcPane[23].top + cy;
					FillRect(hbufdc, &rcPane[23], (HBRUSH)(COLOR_3DDKSHADOW + 1));

					rcPane[i].left += cx;
					DrawText(hbufdc, szProgs[i], cchProgs[i], &rcPane[i], u);
					rcPane[i].left -= cx;
				}
				u |= DT_CENTER;
				DrawText(hbufdc, szProgs[5], cchProgs[5], &rcPane[20], u);

				x = (int)rcPane[7].left;
				y = (int)rcPane[7].top;
				cx = (int)rcPane[22].top;
				cy = (int)rcPane[22].bottom;
				DrawIconEx(hbufdc, x, y, hProgsIcon, cx, cy, 0, NULL, DI_NORMAL);
			}

			EndBufferedPaint(hbufpt, TRUE);
		}

		EndPaint(hPaneWnd, &ps);
	}
}

void OnPaneMouseMove(LPARAM lParam)
{
	TRACKMOUSEEVENT tme;
	POINT pt;
	int i, n;

	if (nMouseMove == 0)
	{
		tme.cbSize = sizeof(TRACKMOUSEEVENT);
		tme.dwFlags = TME_LEAVE;
		tme.dwHoverTime = 10;
		tme.hwndTrack = hPaneWnd;
		TrackMouseEvent(&tme);
	}
	
	pt.x = GET_X_LPARAM(lParam);
	pt.y = GET_Y_LPARAM(lParam);

	if (nTab == 6)
		n = 10;
	else if (nTab == 5)
		n = 9;
	else
		n = 20;
	for (i=n; i>2; i--)
		if (PtInRect(&rcPane[i], pt))
			break;

	if (i == nMouseMove)
		return;

	nMouseMove = i;
	InvalidateRect(hPaneWnd, NULL, TRUE);
}

void OnPaneLButtonDown(LPARAM lParam)
{
	POINT pt;
	int i, n;
	
	pt.x = GET_X_LPARAM(lParam);
	pt.y = GET_Y_LPARAM(lParam);

	if (nTab == 6)
		n = 10;
	else if (nTab == 5)
		n = 9;
	else
		n = 20;
	for (i=n; i>2; i--)
		if (PtInRect(&rcPane[i], pt))
			break;

	if (i>3 && i<7)
		nTab = i;

	nLButtonDown = i;
	InvalidateRect(hPaneWnd, NULL, TRUE);
}

void OnPaneLButtonUp(LPARAM lParam)
{
	POINT pt;
	int i, n;
	
	pt.x = GET_X_LPARAM(lParam);
	pt.y = GET_Y_LPARAM(lParam);

	if (nTab == 6)
		n = 10;
	else if (nTab == 5)
		n = 9;
	else
		n = 20;
	for (i=n; i>2; i--)
		if (PtInRect(&rcPane[i], pt))
			break;

	if (i != nLButtonDown)
		nLButtonDown = 0;

	OpenExecute(TRUE, 0);
}

void OpenExecute(BOOL fOpen, WPARAM wParam)
{
	WPARAM w;
	INT n;
	UINT u;
	DWORD dw;
	TCHAR *pszop;
	TCHAR *pszfile;
	TCHAR *pszparam;
	HANDLE h;
	TOKEN_PRIVILEGES tp;
	HWND hwnd;

	n = 0;
	u = 0;
	dw = 0;
	pszop = NULL;
	pszfile = NULL;
	pszparam = NULL;

	w = wParam >= 'A' ? wParam : nLButtonDown;
	if (nTab == 6)
	{
		switch (w)
		{
		case 'V':
			n = 4;
			break;
		case 'W':
			n = 5;
			break;
		case 'S':
		case 7:
			nLButtonDown = 7;
			u = 0xff;
			break;
		case 'L':
		case 8:
			nLButtonDown = 8;
			u = EWX_LOGOFF + 1;
			break;
		case 'R':
		case 9:
			nLButtonDown = 9;
			u = EWX_REBOOT + 1;
			break;
		case 'U':
		case 10:
			nLButtonDown = 10;
			u = EWX_SHUTDOWN + 1;
			break;
		default:
			break;
		}
	}
	else if (nTab == 5)
	{
		switch (w)
		{
		case 'V':
			n = 4;
			break;
		case 'U':
			n = 6;
			break;
		case 7:
			dw = 7;
			break;
		case 8:
			dw = 8;
			break;
		case 9:
			dw = 9;
			break;
		default:
			break;
		}
	}
	else
	{
		switch (w)
		{
		case 'W':
			n = 5;
			break;
		case 'U':
			n = 6;
			break;
		case 'T':
		case 7:
			nLButtonDown = 7;
			pszfile = TEXT("regedit.exe");
			break;
		case 'P':
		case 8:
			nLButtonDown = 8;
			pszfile = TEXT("control.exe");
			break;
		case 'C':
		case 9:
			nLButtonDown = 9;
			pszfile = TEXT("cmd.exe");
			break;
		case 'A':
		case 10:
			nLButtonDown = 10;
			pszop = TEXT("runas");
			pszfile = TEXT("cmd.exe");
			break;
		case 'E':
		case 11:
			nLButtonDown = 11;
			pszfile = TEXT("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}");
			break;
		case 'D':
		case 12:
			nLButtonDown = 12;
			pszfile = TEXT("devmgmt.msc");
			break;
		case 'M':
		case 13:
			nLButtonDown = 13;
			pszfile = TEXT("perfmon.exe");
			pszparam = TEXT("/res");
			break;
		case 'Y':
		case 14:
			nLButtonDown = 14;
			pszfile = TEXT("msconfig.exe");
			break;
		case 'I':
		case 15:
			nLButtonDown = 15;
			pszfile = TEXT("msinfo32.exe");
			break;
		case 'N':
		case 16:
			nLButtonDown = 16;
			pszfile = TEXT("shell:::{7007ACC7-3202-11D1-AAD2-00805FC1270E}");
			break;
		case 'S':
		case 17:
			nLButtonDown = 17;
			pszfile = TEXT("services.msc");
			break;
		case 'R':
		case 18:
			nLButtonDown = 18;
			pszfile = TEXT("shell:::{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}");
			break;
		case 'V':
		case 20:
			nLButtonDown = 20;
			break;
		default:
			break;
		}
	}

	if (!fOpen)
	{
		InvalidateRect(hPaneWnd, NULL, TRUE);
		return;
	}

	if (n)
	{
		nMouseMove = 0;
		nTab = n;
	}

	if (u)
	{
		if (OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES, &h))
		{
			tp.PrivilegeCount = 1;
			tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;
			LookupPrivilegeValue(NULL, SE_SHUTDOWN_NAME, &tp.Privileges[0].Luid);
			AdjustTokenPrivileges(h, FALSE, &tp, 0, NULL, NULL);
			CloseHandle(h);
		}

		if (u == 0xff)
			SetSystemPowerState(TRUE, FALSE);
		else
			ExitWindowsEx(u - 1, SHTDN_REASON_FLAG_PLANNED);
	}
	
	if (dw == 9)
	{
		dwUseSmallFont = !dwUseSmallFont;
		hwnd = hRebarWnd;
		if (dwUseSmallFont && hSyslistWnd)
			hwnd = hSyslistWnd;
		hFont = (HFONT)SendMessage(hwnd, WM_GETFONT, 0, 0);
		RegSetKeyValue(HKEY_CURRENT_USER, szSetupKey, TEXT("UseSmallFont"),
			REG_DWORD, &dwUseSmallFont, sizeof(DWORD));
	}
	else if (dw == 8)
	{
		dwDisableRibbon = !dwDisableRibbon;
		Hook((PBYTE)NewFunc, (PBYTE)LoadLibraryExW, byFunc, dwDisableRibbon);
		RegSetKeyValue(HKEY_CURRENT_USER, szSetupKey, TEXT("DisableRibbon"),
			REG_DWORD, &dwDisableRibbon, sizeof(DWORD));
	}
	else if (dw == 7)
	{
		dwDisableMetro = !dwDisableMetro;
		if (dwDisableMetro)
		{
			pszop = TEXT("\\");
			dw = lstrlen(pszop) * sizeof(TCHAR);
			RegSetKeyValue(HKEY_CURRENT_USER, szImsspKey, NULL, REG_SZ, pszop, dw);
		}
		else
		{
			szImsspKey[61] = TEXT('\0');
			RegDeleteTree(HKEY_CURRENT_USER, szImsspKey);
			szImsspKey[61] = TEXT('\\');
		}
	}

	if (pszfile)
		ShellExecute(NULL, pszop, pszfile, pszparam, NULL, SW_SHOW);

	nLButtonDown = 0;
	InvalidateRect(hPaneWnd, NULL, TRUE);
}

//
// NewFunc
//
HMODULE WINAPI NewFunc(LPCWSTR lpLibFileName, HANDLE hFile, DWORD dwFlags)
{
	HMODULE hmod;

	if (!lstrcmpiW(lpLibFileName, L"UIRibbonRes.dll"))
		return NULL;

	Hook((PBYTE)NewFunc, (PBYTE)LoadLibraryExW, byFunc, 0);
	hmod = LoadLibraryExW(lpLibFileName, hFile, dwFlags);
	Hook((PBYTE)NewFunc, (PBYTE)LoadLibraryExW, byFunc, 1);

	return hmod;
}

void Hook(PBYTE pbyNewFunc, PBYTE pbyOldFunc, PBYTE pbyFunc, DWORD dwHook)
{
	DWORD dw;

	if (pbyFunc[5] != 0xe9)
	{
		if (!dwHook)
			return;

		memcpy(&pbyFunc[0], pbyOldFunc, 5);
		pbyFunc[5] = 0xe9;
		*(INT32*)&pbyFunc[6] = (INT32)(pbyNewFunc - pbyOldFunc - 5);
	}

	VirtualProtect(pbyOldFunc, 5, PAGE_EXECUTE_READWRITE, &dw);
	memcpy(pbyOldFunc, dwHook ? &pbyFunc[5] : &pbyFunc[0], 5);
	VirtualProtect(pbyOldFunc, 5, dw, &dw);
}