#include "Pane.h"
#include "Resource.h"


extern HINSTANCE hDllInst;
extern TCHAR szSetupKey[];
extern TCHAR szImsspKey[];
extern HWND hStartWnd;
extern HWND hRebarWnd;
extern HWND hTaskbarWnd;
extern HWND hSyslistWnd;
extern HWND hTrayShowDesktopWnd;
BOOL fDisableMetro = FALSE;
BOOL fHideTrayShowDesktop = FALSE;

ATOM atomPane = 0;
HWND hPaneWnd = NULL;
HICON hTabsIcon = NULL;
HICON hProgsIcon = NULL;
HFONT hFont = NULL;
TCHAR szProgs[27][50];
INT cchProgs[27];
INT nTab = 4;
INT nMouseMove = 0;
INT nLButtonDown = 0;
BYTE byFunc[10];
BOOL fUseSmallFont = FALSE;
BOOL fDisableRibbon = FALSE;
FILELIST flList = {};
INT nFirstItem = 0;
INT nLastItem = 0;

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
	292,   2, 300, 434,//21

	 48,  36, 144, 432,
	 16,   0,   0,   0,//23
};


// Pane
static LRESULT CALLBACK PaneProc(HWND, UINT, WPARAM, LPARAM);
void OnPanePaint();
void OnPaneMouseWheel(WPARAM);
void OnPaneMouseMove(LPARAM);
void OnPaneLButtonDown(LPARAM);
void OnPaneLButtonUp(LPARAM);
void OpenExecute(BOOL, WPARAM);

// NewFunc
HMODULE WINAPI NewFunc(LPCWSTR, HANDLE, DWORD);
void HookApi(PBYTE, PBYTE, PBYTE, DWORD);

// FileList
HRESULT GetFileList(FILELIST *);
void DrawFileList(FILELIST *, HDC*, INT *, BOOL);
void FreeFileList(FILELIST *);


HRESULT CreatePane()
{
	INT i;
	DWORD cb;
	TCHAR sz[2];
	HRESULT hr;
	WNDCLASSEX wcex;

	cb = sizeof(sz);
	RegGetValue(HKEY_CURRENT_USER, szImsspKey, NULL, RRF_RT_ANY, NULL, &sz, &cb);
	if (!lstrcmp(sz, TEXT("\\")))
		fDisableMetro = TRUE;

	cb = sizeof(BOOL);
	RegGetValue(HKEY_CURRENT_USER, szSetupKey, TEXT("UseSmallFont"),
		RRF_RT_ANY, NULL, &fUseSmallFont, &cb);

	cb = sizeof(BOOL);
	RegGetValue(HKEY_CURRENT_USER, szSetupKey, TEXT("DisableRibbon"),
		RRF_RT_ANY, NULL, &fDisableRibbon, &cb);
	
	cb = sizeof(BOOL);
	RegGetValue(HKEY_CURRENT_USER, szSetupKey, TEXT("HideTrayShowDesktop"),
		RRF_RT_ANY, NULL, &fHideTrayShowDesktop, &cb);

	if (fDisableRibbon)
		HookApi((PBYTE)NewFunc, (PBYTE)LoadLibraryExW, byFunc, 1);

	if (fHideTrayShowDesktop)
		InvalidateRect(hTrayShowDesktopWnd, NULL, TRUE);

	//float f=1.5;for(i=0;i<24;i++){rcPane[i].left*=f;rcPane[i].top*=f;rcPane[i].right*=f;rcPane[i].bottom*=f;}
	for (i=0; i<27; i++)
		cchProgs[i] = LoadString(hDllInst, i + IDS_PROG1, szProgs[i], 50);

	hTabsIcon = (HICON)LoadImage(hDllInst, MAKEINTRESOURCE(IDI_TABS),
		IMAGE_ICON, (int)rcPane[22].left, (int)rcPane[22].right, LR_DEFAULTCOLOR);
	if (!hTabsIcon)
		return E_FAIL;

	hProgsIcon = (HICON)LoadImage(hDllInst, MAKEINTRESOURCE(IDI_PROGS),
		IMAGE_ICON, (int)rcPane[22].top, (int)rcPane[22].bottom, LR_DEFAULTCOLOR);
	if (!hProgsIcon)
		return E_FAIL;

	hr = SHGetKnownFolderPath(FOLDERID_Programs, KF_FLAG_DEFAULT,
		NULL, &flList.pszpath1);
	if (FAILED(hr))
		return hr;

	hr = SHGetKnownFolderPath(FOLDERID_CommonPrograms, KF_FLAG_DEFAULT,
		NULL, &flList.pszpath2);
	if (FAILED(hr))
		return hr;

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
	atomPane = RegisterClassEx(&wcex);
	if (!atomPane)
		return E_FAIL;

	hPaneWnd = CreateWindowEx(WS_EX_TOOLWINDOW | WS_EX_TOPMOST, TEXT("Pane"),
		NULL, WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		0, 0, 0, 0, hTaskbarWnd, NULL, hDllInst, NULL);
	if (!hPaneWnd)
		return E_FAIL;
	
	return S_OK;
}

void ShowPane()
{
	RECT rc;
	LONG x, y;
	int w, h;
	HWND hwnd;
	static HWND hforewnd = NULL;

	if (hPaneWnd)
	{
		if (IsWindowVisible(hPaneWnd))
		{
			ShowWindow(hPaneWnd, SW_HIDE);
			SetForegroundWindow(hforewnd);
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
		if (fUseSmallFont && hSyslistWnd)
			hwnd = hSyslistWnd;
		hFont = (HFONT)SendMessage(hwnd, WM_GETFONT, 0, 0);

		hforewnd = GetForegroundWindow();
		ShowWindow(hPaneWnd, SW_SHOW);
		SetForegroundWindow(hPaneWnd);
	}
}

void ClosePane()
{
	if (hPaneWnd)
	{
		DestroyWindow(hPaneWnd);
		hPaneWnd = NULL;
	}
	
	if (atomPane)
	{
		UnregisterClass(TEXT("Pane"), hDllInst);
		atomPane = 0;
	}

	if (flList.pszpath2)
	{
		CoTaskMemFree(flList.pszpath2);
		flList.pszpath2 = NULL;
	}

	if (flList.pszpath1)
	{
		CoTaskMemFree(flList.pszpath1);
		flList.pszpath1 = NULL;
	}

	FreeFileList(&flList);

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
	
	if (fDisableRibbon)
	{
		HookApi((PBYTE)NewFunc, (PBYTE)LoadLibraryExW, byFunc, 0);
		fDisableRibbon = FALSE;
	}
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
	case WM_MOUSEWHEEL:
		OnPaneMouseWheel(wParam);
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
	RECT rc;
	double d;

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
				i = fDisableMetro ? 19 : 20;
				DrawText(hbufdc, szProgs[i], cchProgs[i], &rcPane[7], u);
				i = fDisableRibbon ? 21 : 22;
				DrawText(hbufdc, szProgs[i], cchProgs[i], &rcPane[8], u);
				i = fUseSmallFont ? 23 : 24;
				DrawText(hbufdc, szProgs[i], cchProgs[i], &rcPane[9], u);
				i = fHideTrayShowDesktop ? 25 : 26;
				DrawText(hbufdc, szProgs[i], cchProgs[i], &rcPane[10], u);
				u |= DT_CENTER;
				DrawText(hbufdc, szProgs[4], cchProgs[4], &rcPane[20], u);
			}
			else
			{
				if (flList.bshow)
				{
					x = -rcPane[23].left;
					nLastItem = 0;
					DrawFileList(&flList, &hbufdc, &x, TRUE);
					for (i=nLastItem; i<12; i++)
						FillRect(hbufdc, &rcPane[i+7], (HBRUSH)(COLOR_BTNSHADOW + 1));

					if (nMouseMove && nLastItem > 12)
					{
						FillRect(hbufdc, &rcPane[21], (HBRUSH)(COLOR_ACTIVEBORDER + 1));
						rc.left = rcPane[21].left;
						rc.right = rcPane[21].right;
						x = rcPane[7].bottom - rcPane[7].top;
						y = rcPane[21].bottom - rcPane[21].top;
						d = y * y;
						d /= x;
						rc.bottom = (LONG)(d / nLastItem);

						rc.top = y - rc.bottom;
						d = nLastItem - 12;
						d = rc.top / d;
						rc.top = (LONG)(nFirstItem * d);
						rc.top += rcPane[7].top;
						rc.bottom += rc.top;
						FillRect(hbufdc, &rc, (HBRUSH)(COLOR_SCROLLBAR + 1));
					}
					u |= DT_CENTER;
					DrawText(hbufdc, szProgs[6], cchProgs[6], &rcPane[20], u);
				}
				else
				{
					for(i=7; i<19; i++)
					{
						rc.left = rcPane[i].left + rcPane[7].top;
						rc.top = rcPane[i].top + rcPane[7].top;
						rc.bottom = rcPane[i].bottom - rcPane[7].top;
						rc.right = rc.bottom - rc.top;
						rc.right += rc.left;
						FillRect(hbufdc, &rc, (HBRUSH)(COLOR_3DDKSHADOW + 1));

						rc.left = rc.right + rcPane[7].top;
						rc.right = rcPane[7].right - rcPane[7].top;
						DrawText(hbufdc, szProgs[i], cchProgs[i], &rc, u);
					}
					u |= DT_CENTER;
					DrawText(hbufdc, szProgs[5], cchProgs[5], &rcPane[20], u);

					x = (int)rcPane[7].left;
					y = (int)rcPane[7].top;
					cx = (int)rcPane[22].top;
					cy = (int)rcPane[22].bottom;
					DrawIconEx(hbufdc, x, y, hProgsIcon, cx, cy, 0, NULL, DI_NORMAL);
				}
			}

			EndBufferedPaint(hbufpt, TRUE);
		}

		EndPaint(hPaneWnd, &ps);
	}
}

void OnPaneMouseWheel(WPARAM wParam)
{
	int n;

	if (nTab != 4 || !flList.bshow || !nMouseMove)
		return;

	n = GET_WHEEL_DELTA_WPARAM(wParam) / WHEEL_DELTA;
	nFirstItem -= n;
	if (nFirstItem > nLastItem - 12)
		nFirstItem = nLastItem - 12;
	if (nFirstItem < 0)
		nFirstItem = 0;

	InvalidateRect(hPaneWnd, NULL, TRUE);
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
		n = 10;
	else if (flList.bshow && nLastItem > 12)
		n = 21;
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
		n = 10;
	else if (flList.bshow && nLastItem > 12)
		n = 21;
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
		n = 10;
	else if (flList.bshow && nLastItem > 12)
		n = 21;
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
	INT n;
	UINT u;
	DWORD dw;
	TCHAR *pszop;
	TCHAR *pszfile;
	TCHAR *pszparam;
	TCHAR szdir[MAX_PATH];
	DWORD nsize;
	WPARAM w;
	HANDLE h;
	TOKEN_PRIVILEGES tp;
	HWND hwnd;

	n = 0;
	u = 0;
	dw = 0;
	pszop = NULL;
	pszfile = NULL;
	pszparam = NULL;
	szdir[0] = TEXT('\0');

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
		case 10:
			dw = 10;
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
			nsize = MAX_PATH * sizeof(TCHAR);
			ExpandEnvironmentStrings(TEXT("%HOMEDRIVE%%HOMEPATH%"), szdir, nsize);
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
	
	if (dw == 10)
	{
		fHideTrayShowDesktop = !fHideTrayShowDesktop;
		InvalidateRect(hTrayShowDesktopWnd, NULL, TRUE);
		RegSetKeyValue(HKEY_CURRENT_USER, szSetupKey, TEXT("HideTrayShowDesktop"),
			REG_DWORD, &fHideTrayShowDesktop, sizeof(BOOL));
	}
	else if (dw == 9)
	{
		fUseSmallFont = !fUseSmallFont;
		hwnd = hRebarWnd;
		if (fUseSmallFont && hSyslistWnd)
			hwnd = hSyslistWnd;
		hFont = (HFONT)SendMessage(hwnd, WM_GETFONT, 0, 0);
		RegSetKeyValue(HKEY_CURRENT_USER, szSetupKey, TEXT("UseSmallFont"),
			REG_DWORD, &fUseSmallFont, sizeof(BOOL));
	}
	else if (dw == 8)
	{
		fDisableRibbon = !fDisableRibbon;
		HookApi((PBYTE)NewFunc, (PBYTE)LoadLibraryExW, byFunc, fDisableRibbon);
		RegSetKeyValue(HKEY_CURRENT_USER, szSetupKey, TEXT("DisableRibbon"),
			REG_DWORD, &fDisableRibbon, sizeof(BOOL));
	}
	else if (dw == 7)
	{
		fDisableMetro = !fDisableMetro;
		if (fDisableMetro)
		{
			pszop = TEXT("\\");
			nsize = lstrlen(pszop) * sizeof(TCHAR);
			RegSetKeyValue(HKEY_CURRENT_USER, szImsspKey, NULL, REG_SZ, pszop, nsize);
		}
		else
		{
			szImsspKey[61] = TEXT('\0');
			RegDeleteTree(HKEY_CURRENT_USER, szImsspKey);
			szImsspKey[61] = TEXT('\\');
		}
	}

	if (nTab == 4)
	{
		if (nLButtonDown == 20)
			flList.bshow = !flList.bshow;

		if (flList.bshow)
		{
			GetFileList(&flList);
			if (nLButtonDown > 6 && nLButtonDown < 19)
			{
				n = 0;
				nLastItem = 0;
				nLButtonDown -= 7;
				DrawFileList(&flList, NULL, &n, FALSE);
				if (nFirstItem > nLastItem - 12)
					nFirstItem = nLastItem - 12;
				if (nFirstItem < 0)
					nFirstItem = 0;
			}
		}
		else
		{
			nFirstItem = 0;
			FreeFileList(&flList);
			if (pszfile)
				ShellExecute(NULL, pszop, pszfile, pszparam, szdir, SW_SHOW);
		}
	}

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

	HookApi((PBYTE)NewFunc, (PBYTE)LoadLibraryExW, byFunc, 0);
	hmod = LoadLibraryExW(lpLibFileName, hFile, dwFlags);
	HookApi((PBYTE)NewFunc, (PBYTE)LoadLibraryExW, byFunc, 1);

	return hmod;
}

void HookApi(PBYTE pbyNewFunc, PBYTE pbyOldFunc, PBYTE pbyFunc, DWORD dwHook)
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

//
// FileList
//
HRESULT GetFileList(FILELIST *pfl)
{
	HRESULT hr;
	IImageList *pimgl;
	IShellItem *psi;
	IShellItem *psisub;
	IEnumShellItems *pesi;
	SHFILEINFO sfi;
	SFGAOF sfg;
	TCHAR *psz;
	int i, m, n, t, cmp;

	if (!pfl)
		return E_FAIL;
	if (pfl->pfl)
		return S_OK;

	pimgl = NULL;
	psi = NULL;
	psisub = NULL;
	pesi = NULL;
	psz = pfl->pszpath1;

	t = 100;
	pfl->pfl = (FILELIST *)CoTaskMemAlloc(t * sizeof(FILELIST));
	if (!pfl->pfl)
		return E_OUTOFMEMORY;
	ZeroMemory(pfl->pfl, t * sizeof(FILELIST));
	t--;

	hr = SHGetImageList(SHIL_LARGE, IID_PPV_ARGS(&pimgl));//SHIL_EXTRALARGE
	if (FAILED(hr))
		goto end;
start:
	hr = SHCreateItemFromParsingName(psz, NULL, IID_PPV_ARGS(&psi));
	if (FAILED(hr))
		goto end;

	hr = psi->BindToHandler(NULL, BHID_EnumItems, IID_PPV_ARGS(&pesi));
	if (FAILED(hr))
		goto end;

	while (pesi->Next(1, &psisub, NULL) == S_OK)
	{
		if (pfl->nfldr + pfl->nfile >= t)
			break;
		hr = psisub->GetAttributes(SFGAO_FOLDER | SFGAO_HIDDEN, &sfg);
		if (FAILED(hr))
			break;
		if (sfg & SFGAO_HIDDEN)
		{
			psisub->Release();
			psisub = NULL;
			continue;
		}
		hr = psisub->GetDisplayName(SIGDN_NORMALDISPLAY, &pfl->pfl[t].pszdisp);
		if (FAILED(hr))
			break;
		hr = psisub->GetDisplayName(SIGDN_FILESYSPATH, &pfl->pfl[t].pszpath1);
		if (FAILED(hr))
			break;
		if (!SHGetFileInfo(pfl->pfl[t].pszpath1, 0, &sfi, sizeof(SHFILEINFO),
			SHGFI_SYSICONINDEX))
			break;
		hr = pimgl->GetIcon(sfi.iIcon, ILD_NORMAL, &pfl->pfl[t].hicon);
		if (FAILED(hr))
			break;

		n = pfl->nfldr;
		m = pfl->nfldr + pfl->nfile;
		if (sfg & SFGAO_FOLDER)
		{
			pfl->pfl[t].bfldr = TRUE;
			memcpy(&pfl->pfl[m], &pfl->pfl[n], sizeof(FILELIST));
			memcpy(&pfl->pfl[n], &pfl->pfl[t], sizeof(FILELIST));
			for (i=0; i<n; i++)
			{
				cmp = lstrcmp(pfl->pfl[n].pszdisp, pfl->pfl[i].pszdisp);
				if (cmp < 0)
				{
					memcpy(&pfl->pfl[t], &pfl->pfl[i], sizeof(FILELIST));
					memcpy(&pfl->pfl[i], &pfl->pfl[n], sizeof(FILELIST));
					memcpy(&pfl->pfl[n], &pfl->pfl[t], sizeof(FILELIST));
				}
				else if (cmp == 0)
				{
					DestroyIcon(pfl->pfl[n].hicon);
					CoTaskMemFree(pfl->pfl[n].pszdisp);
					pfl->pfl[i].pszpath2 = pfl->pfl[n].pszpath1;
					memcpy(&pfl->pfl[n], &pfl->pfl[m], sizeof(FILELIST));
					m = pfl->nfldr;
					pfl->nfldr--;
					break;
				}
			}
			pfl->nfldr++;
			n = pfl->nfldr;
		}
		else
		{
			pfl->pfl[t].bfldr = FALSE;
			memcpy(&pfl->pfl[m], &pfl->pfl[t], sizeof(FILELIST));
			pfl->nfile++;
		}

		for (i=n; i<m; i++)
		{
			cmp = lstrcmp(pfl->pfl[m].pszdisp, pfl->pfl[i].pszdisp);
			if (cmp < 0)
			{
				memcpy(&pfl->pfl[t], &pfl->pfl[i], sizeof(FILELIST));
				memcpy(&pfl->pfl[i], &pfl->pfl[m], sizeof(FILELIST));
				memcpy(&pfl->pfl[m], &pfl->pfl[t], sizeof(FILELIST));
			}
		}

		ZeroMemory(&pfl->pfl[t], sizeof(FILELIST));
		psisub->Release();
		psisub = NULL;
	}

end:
	if (pfl->pfl[t].hicon)
	{
		DestroyIcon(pfl->pfl[t].hicon);
		pfl->pfl[t].hicon = NULL;
	}
	if (pfl->pfl[t].pszdisp)
	{
		CoTaskMemFree(pfl->pfl[t].pszdisp);
		pfl->pfl[t].pszdisp = NULL;
	}
	if (pfl->pfl[t].pszpath1)
	{
		CoTaskMemFree(pfl->pfl[t].pszpath1);
		pfl->pfl[t].pszpath1 = NULL;
	}
	if (psisub)
	{
		psisub->Release();
		psisub = NULL;
	}
	if (pesi)
	{
		pesi->Release();
		pesi = NULL;
	}
	if (psi)
	{
		psi->Release();
		psi = NULL;
	}
	if (psz == pfl->pszpath1 && lstrlen(psz))
	{
		psz = pfl->pszpath2;
		if (lstrlen(psz))
			goto start;
	}
	if (pimgl)
	{
		pimgl->Release();
		pimgl = NULL;
	}

	return hr;
}

void DrawFileList(FILELIST *pfl, HDC *phdc, INT *px, BOOL fdraw)
{
	RECT rc;
	int i, x, y, cx, cy, n;

	if (!pfl || !pfl->pfl)
		return;

	*px += rcPane[23].left;
	for (i=0; i<pfl->nfldr+pfl->nfile; i++)
	{
		nLastItem++;
		if (fdraw)
		{
			if (nLastItem > nFirstItem && nLastItem < nFirstItem + 13)
			{
				n = nLastItem - nFirstItem + 6;
				rc.top = rcPane[n].top + rcPane[7].top;
				rc.bottom = rcPane[n].bottom - rcPane[7].top;
				rc.left = rcPane[n].left + rcPane[7].top + *px;
				rc.right = rc.bottom - rc.top;
				rc.right += rc.left;
				FillRect(*phdc, &rc, (HBRUSH)(COLOR_3DDKSHADOW + 1));

				x = (int)rc.left;
				y = (int)rc.top;
				cx = (int)(rc.right - rc.left);
				cy = (int)(rc.bottom - rc.top);
				DrawIconEx(*phdc, x, y, pfl->pfl[i].hicon, cx, cy, 0, NULL, DI_NORMAL);

				rc.left = rc.right + rcPane[7].top;
				rc.right = rcPane[7].right - rcPane[7].top;
				n = lstrlen(pfl->pfl[i].pszdisp);
				DrawText(*phdc, pfl->pfl[i].pszdisp, n, &rc, DT_SINGLELINE | DT_VCENTER);
			}
		}
		else
		{
			if (nLastItem == nLButtonDown + nFirstItem + 1)
			{
				if (pfl->pfl[i].bfldr)
				{
					pfl->pfl[i].bshow = !pfl->pfl[i].bshow;
					if (pfl->pfl[i].bshow)
						GetFileList(&pfl->pfl[i]);
					else
						FreeFileList(&pfl->pfl[i]);
				}
				else
					ShellExecute(NULL, NULL, pfl->pfl[i].pszpath1, NULL, NULL, SW_SHOW);
			}
		}

		if (pfl->pfl[i].bshow)
			DrawFileList(&pfl->pfl[i], phdc, px, fdraw);
	}
	*px -= rcPane[23].left;
}

void FreeFileList(FILELIST *pfl)
{
	int i;

	if (!pfl || !pfl->pfl)
		return;

	for (i=0; i<pfl->nfldr+pfl->nfile; i++)
	{
		if (pfl->pfl[i].hicon)
		{
			DestroyIcon(pfl->pfl[i].hicon);
			pfl->pfl[i].hicon = NULL;
		}
		if (pfl->pfl[i].pszdisp)
		{
			CoTaskMemFree(pfl->pfl[i].pszdisp);
			pfl->pfl[i].pszdisp = NULL;
		}
		if (pfl->pfl[i].pszpath1)
		{
			CoTaskMemFree(pfl->pfl[i].pszpath1);
			pfl->pfl[i].pszpath1 = NULL;
		}
		if (pfl->pfl[i].pszpath2)
		{
			CoTaskMemFree(pfl->pfl[i].pszpath2);
			pfl->pfl[i].pszpath2 = NULL;
		}
		FreeFileList(&pfl->pfl[i]);
	}
	pfl->nfldr = 0;
	pfl->nfile = 0;
	CoTaskMemFree(pfl->pfl);
	pfl->pfl = NULL;
}