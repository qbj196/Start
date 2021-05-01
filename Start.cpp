#include "Start.h"
#include "Pane.h"
#include "Resource.h"


extern HINSTANCE hDllInst;
extern BOOL fDisableMetro;
extern BOOL fHideTrayShowDesktop;
HWND hStartWnd = NULL;
HWND hRebarWnd = NULL;
HWND hTaskbarWnd = NULL;
HWND hSyslistWnd = NULL;
HWND hTrayShowDesktopWnd = NULL;

ATOM atomStart = 0;
HWND hProgmanWnd = NULL;
HWND hNetworkWnd = NULL;
HWND hOrigStartWnd = NULL;
BOOL fEnterDesktop = FALSE;
BOOL fStartEnabled = TRUE;
BOOL fMouseLeftWnd = TRUE;
BOOL fPreShowDesktop = TRUE;
HICON hStartIcon16 = NULL;
HICON hStartIcon32 = NULL;
WNDPROC wpOrigRebarProc = NULL;
WNDPROC wpOrigTaskbarProc = NULL;
WNDPROC wpOrigProgmanProc = NULL;
WNDPROC wpOrigNetworkProc = NULL;
WNDPROC wpOrigTrayShowDesktopProc = NULL;


// Start
static LRESULT CALLBACK StartProc(HWND, UINT, WPARAM, LPARAM);
void OnStartPaint();
void OnStartMouseMove();
void OnStartRButtonUp();

// Rebar
static LRESULT CALLBACK RebarProc(HWND, UINT, WPARAM, LPARAM);
void OnRebarWindowPosChanging(LPARAM);
void OnRebarWindowPosChanged(LPARAM);

// Taskbar
static LRESULT CALLBACK TaskbarProc(HWND, UINT, WPARAM, LPARAM);
void OnTaskbarEnterDesktop();
void OnTaskbarSetWindowProc();
void OnTaskbarShowEdgeui(BOOL);

// Progman
static LRESULT CALLBACK ProgmanProc(HWND, UINT, WPARAM, LPARAM);

// Network
static LRESULT CALLBACK NetworkProc(HWND, UINT, WPARAM, LPARAM);

// TrayShowDesktop
static LRESULT CALLBACK TrayShowDesktopProc(HWND, UINT, WPARAM, LPARAM);


HRESULT CreateStart(HWND hWnd)
{
	HWND htray;
	WNDCLASSEX wcex;

	hRebarWnd = hWnd;
	hTaskbarWnd = GetParent(hWnd);

	htray = FindWindowEx(hTaskbarWnd, NULL, TEXT("TrayNotifyWnd"), NULL);
	if (!htray)
		return E_FAIL;

	hTrayShowDesktopWnd = FindWindowEx(htray, NULL,
		TEXT("TrayShowDesktopButtonWClass"), NULL);
	if (!hTrayShowDesktopWnd)
		return E_FAIL;

	wpOrigTrayShowDesktopProc = (WNDPROC)SetWindowLongPtr(hTrayShowDesktopWnd,
		GWLP_WNDPROC, (LONG_PTR)TrayShowDesktopProc);
	if (!wpOrigTrayShowDesktopProc)
		return E_FAIL;

	hOrigStartWnd = FindWindowEx(hTaskbarWnd, NULL, TEXT("Start"), NULL);
	if (hOrigStartWnd)
		ShowWindow(hOrigStartWnd, SW_HIDE);

	wpOrigRebarProc = (WNDPROC)SetWindowLongPtr(hRebarWnd, GWLP_WNDPROC,
		(LONG_PTR)RebarProc);
	if (!wpOrigRebarProc)
		return E_FAIL;

	wpOrigTaskbarProc = (WNDPROC)SetWindowLongPtr(hTaskbarWnd, GWLP_WNDPROC,
		(LONG_PTR)TaskbarProc);
	if (!wpOrigTaskbarProc)
		return E_FAIL;

	hStartIcon16 = (HICON)LoadImage(hDllInst, MAKEINTRESOURCE(IDI_START),
		IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR);
	if (!hStartIcon16)
		return E_FAIL;

	hStartIcon32 = (HICON)LoadImage(hDllInst, MAKEINTRESOURCE(IDI_START),
		IMAGE_ICON, 32, 32, LR_DEFAULTCOLOR);
	if (!hStartIcon32)
		return E_FAIL;

	wcex.cbClsExtra		= 0;
	wcex.cbSize			= sizeof(WNDCLASSEX);
	wcex.cbWndExtra		= 0;
	wcex.hbrBackground	= NULL;
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hIcon			= NULL;
	wcex.hIconSm		= NULL;
	wcex.hInstance		= hDllInst;
	wcex.lpfnWndProc	= StartProc;
	wcex.lpszClassName	= TEXT("Start");
	wcex.lpszMenuName	= NULL;
	wcex.style			= CS_VREDRAW | CS_HREDRAW;
	atomStart = RegisterClassEx(&wcex);
	if (!atomStart)
		return E_FAIL;

	hStartWnd = CreateWindowEx(0, TEXT("Start"), NULL,
		WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		0, 0, 0, 0, hTaskbarWnd, NULL, hDllInst, NULL);
	if (!hStartWnd)
		return E_FAIL;
	
	return CreatePane();
}

void ShowStart(BOOL fShow, DWORD dwBandID)
{
	LRESULT index;

	index = SendMessage(hRebarWnd, RB_IDTOINDEX, dwBandID, 0);
	SendMessage(hRebarWnd, RB_SHOWBAND, index, FALSE); //hide this band
	SendMessage(hTaskbarWnd, WM_SIZE, 0, 0);

	if (hStartWnd)
		ShowWindow(hStartWnd, fShow ? SW_SHOW : SW_HIDE);
}

void UpdateStart()
{
	if (hStartWnd)
	{
		InvalidateRect(hStartWnd, NULL, TRUE);
		UpdateWindow(hStartWnd);
	}
}

void CloseStart()
{
	ClosePane();

	if (hStartWnd)
	{
		DestroyWindow(hStartWnd);
		hStartWnd = NULL;
	}

	if (atomStart)
	{
		UnregisterClass(TEXT("Start"), hDllInst);
		atomStart = 0;
	}

	if (hStartIcon32)
	{
		DestroyIcon(hStartIcon32);
		hStartIcon32 = NULL;
	}

	if (hStartIcon16)
	{
		DestroyIcon(hStartIcon16);
		hStartIcon16 = NULL;
	}

	if (wpOrigNetworkProc)
	{
		SetWindowLongPtr(hNetworkWnd, GWLP_WNDPROC,
			(LONG_PTR)wpOrigNetworkProc);
		wpOrigNetworkProc = NULL;
	}

	if (wpOrigProgmanProc)
	{
		SetWindowLongPtr(hProgmanWnd, GWLP_WNDPROC,
			(LONG_PTR)wpOrigProgmanProc);
		wpOrigProgmanProc = NULL;
	}

	if (wpOrigTaskbarProc)
	{
		SetWindowLongPtr(hTaskbarWnd, GWLP_WNDPROC,
			(LONG_PTR)wpOrigTaskbarProc);
		wpOrigTaskbarProc = NULL;
	}

	if (wpOrigRebarProc)
	{
		SetWindowLongPtr(hRebarWnd, GWLP_WNDPROC,
			(LONG_PTR)wpOrigRebarProc);
		wpOrigRebarProc = NULL;
	}

	if (hOrigStartWnd)
	{
		ShowWindow(hOrigStartWnd, SW_SHOW);
		hOrigStartWnd = NULL;
	}
	
	if (wpOrigTrayShowDesktopProc)
	{
		SetWindowLongPtr(hTrayShowDesktopWnd, GWLP_WNDPROC,
			(LONG_PTR)wpOrigTrayShowDesktopProc);
		wpOrigTrayShowDesktopProc = NULL;

		InvalidateRect(hTrayShowDesktopWnd, NULL, TRUE);
	}

	OnTaskbarShowEdgeui(TRUE);

	PostMessage(hTaskbarWnd, WM_SIZE, 0, 0);
	PostMessage(hTaskbarWnd, WM_TIMER, 24, 0); //explorer release unuse dll
}

//
// Start
//
static LRESULT CALLBACK StartProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_PAINT:
		OnStartPaint();
		break;
	case WM_THEMECHANGED:
		InvalidateRect(hWnd, NULL, TRUE);
		break;
	case WM_ENABLE:
		fStartEnabled = (BOOL)wParam;
		InvalidateRect(hWnd, NULL, TRUE);
		break;
	case WM_MOUSELEAVE:
		fMouseLeftWnd = TRUE;
		InvalidateRect(hWnd, NULL, TRUE);
		break;
	case WM_MOUSEMOVE:
		OnStartMouseMove();
		break;
	case WM_LBUTTONDOWN:
		ShowPane();
		break;
	case WM_RBUTTONUP:
		OnStartRButtonUp();
		break;
	default:
		return DefWindowProc(hWnd, uMsg, wParam, lParam);
	}

	return 0;
}

void OnStartPaint()
{
	HDC hdc;
	PAINTSTRUCT ps;
	RECT rc;
	HDC hbufdc;
	HPAINTBUFFER hbufpt;
	int x, y, w;
	HICON hicon;

	hdc = BeginPaint(hStartWnd, &ps);
	if (hdc)
	{
		GetClientRect(hStartWnd, &rc);
		hbufpt = BeginBufferedPaint(hdc, &rc, BPBF_TOPDOWNDIB, NULL, &hbufdc);
		if (hbufpt)
		{
			DrawThemeParentBackground(hStartWnd, hbufdc, &rc);

			if (!fStartEnabled)
				BufferedPaintSetAlpha(hbufpt, &rc, 96);
			else if (!fMouseLeftWnd)
				BufferedPaintSetAlpha(hbufpt, &rc, 0);

			w = rc.bottom < 32 ? 16 : 32;
			x = (rc.right - w) / 2;
			y = (rc.bottom - w) / 2;
			hicon = rc.bottom < 32 ? hStartIcon16 : hStartIcon32;
			DrawIconEx(hbufdc, x, y, hicon, w, w, 0, NULL, DI_NORMAL);

			EndBufferedPaint(hbufpt, TRUE);
		}

		EndPaint(hStartWnd, &ps);
	}
}

void OnStartMouseMove()
{
	TRACKMOUSEEVENT tme;

	if (fMouseLeftWnd)
	{
		fMouseLeftWnd = FALSE;
		InvalidateRect(hStartWnd, NULL, TRUE);

		tme.cbSize = sizeof(TRACKMOUSEEVENT);
		tme.dwFlags = TME_LEAVE;
		tme.dwHoverTime = 10;
		tme.hwndTrack = hStartWnd;
		TrackMouseEvent(&tme);
	}
}

void OnStartRButtonUp()
{
	UINT umsg;
	HWND hwnd;

	umsg = RegisterWindowMessage(TEXT("SHELLHOOK"));
	if (!umsg)
		return;

	hwnd = FindWindowEx(hRebarWnd, NULL, TEXT("MSTaskSwWClass"), NULL);
	if (!hwnd)
		return;

	PostMessage(hwnd, umsg, HSHELL_TASKMAN, 0); //show start screen or start menu
}

//
// Rebar
//
static LRESULT CALLBACK RebarProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_WINDOWPOSCHANGING:
		OnRebarWindowPosChanging(lParam);
		break;
	case WM_WINDOWPOSCHANGED:
		OnRebarWindowPosChanged(lParam);
		break;
	default:
		break;
	}

	return CallWindowProc(wpOrigRebarProc, hWnd, uMsg, wParam, lParam);
}

void OnRebarWindowPosChanging(LPARAM lParam)
{
	WINDOWPOS *p;

	p = (WINDOWPOS *)lParam;
	if (p->x == 0)
	{
		p->cy += p->y;
		p->y = 48;
		p->cy -= p->y;
	}
	else
	{
		p->cx += p->x;
		p->x = 48;
		if (p->cy < 40)
			p->x = 36;
		p->cx -= p->x;
	}
}

void OnRebarWindowPosChanged(LPARAM lParam)
{
	WINDOWPOS *p;
	int w, h;

	p = (WINDOWPOS *)lParam;
	if (p->x == 0)
	{
		w = p->cx;
		h = p->y;
	}
	else
	{
		w = p->x;
		h = p->cy;
	}

	MoveWindow(hStartWnd, 0, 0, w, h, TRUE);
}

//
// Taskbar
//
static LRESULT CALLBACK TaskbarProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case 0x0565: //win8.1 explorer launch will receive this
		fPreShowDesktop = FALSE;
		break;
	case 0x0574:
		if (lParam == 0x10000) //desktop has show will receive this
	case 0x0550: //this band launch will receive this
	case 0x0590: //win8.0 and 8.1 explorer launch will receive this
		OnTaskbarSetWindowProc();
		break;
	case 0x05b3: //win8.0 explorer launch will receive this
		fEnterDesktop = !fDisableMetro;
		break;
	case 0x0579: //switch to start screen will receive this
		OnTaskbarEnterDesktop();
		break;
	case 0x05ba: //show or hide pc settings will receive this
	case 0x05bb: //show or hide start screen will receive this
		OnTaskbarShowEdgeui((BOOL)wParam);
		break;
	default:
		break;
	}

	return CallWindowProc(wpOrigTaskbarProc, hWnd, uMsg, wParam, lParam);
}

void OnTaskbarEnterDesktop()
{
	if (fEnterDesktop)
	{
		fEnterDesktop = FALSE;
		SetForegroundWindow(hTaskbarWnd);
	}
}

void OnTaskbarSetWindowProc()
{
	HWND hwnd;

	if (!wpOrigProgmanProc)
	{
		hProgmanWnd = GetShellWindow();
		if (!hProgmanWnd)
			return;
		
		wpOrigProgmanProc = (WNDPROC)SetWindowLongPtr(hProgmanWnd,
			GWLP_WNDPROC, (LONG_PTR)ProgmanProc);

		hwnd = FindWindowEx(hProgmanWnd, NULL, TEXT("SHELLDLL_DefView"), NULL);
		if (hwnd)
			hSyslistWnd = FindWindowEx(hwnd, NULL, TEXT("SysListView32"), NULL);

		if (fPreShowDesktop)
			PostMessage(hProgmanWnd, 0x045c, 3, 7); //win8.0 pre-show desktop

		if (fDisableMetro)
		{
			PostMessage(hTaskbarWnd, 0x0574, 2, 0); //show taskbar and desktop
			PostMessage(hTaskbarWnd, 0x05ba, 0, 0); //enable show desktop and open windows
		}
	}

	if (fDisableMetro && !wpOrigNetworkProc)
	{
		hNetworkWnd = FindWindowEx(NULL, NULL, NULL, TEXT("Network Flyout"));
		if (!hNetworkWnd)
			return;

		wpOrigNetworkProc = (WNDPROC)SetWindowLongPtr(hNetworkWnd,
			GWLP_WNDPROC, (LONG_PTR)NetworkProc);
	}
	
	OnTaskbarShowEdgeui(FALSE);
}

void OnTaskbarShowEdgeui(BOOL fShow)
{
	HWND hwnd;
	HWND hpar;
	TCHAR *psz;
	LONG_PTR l;

	if (!hProgmanWnd)
		return;

	hwnd = NULL;
	hpar = hProgmanWnd;
	psz = TEXT("WorkerW");
	for (;;)
	{
		hwnd = FindWindowEx(hpar, hwnd, psz, NULL);
		if (!hwnd)
			break;

		l = GetWindowLongPtr(hwnd, GWL_EXSTYLE);
		if (l)
		{
			if (fShow)
				l ^= WS_EX_TRANSPARENT;
			else
				l |= WS_EX_TRANSPARENT;
			SetWindowLongPtr(hwnd, GWL_EXSTYLE, l);
		}

		if (hpar == hProgmanWnd)
		{
			hwnd = NULL;
			hpar = NULL;
			psz = TEXT("EdgeUiInputWndClass");
		}
	}
}

//
// Progman
//
static LRESULT CALLBACK ProgmanProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	if (uMsg == WM_SYSCOMMAND && wParam == SC_TASKLIST) //press win key
	{
		ShowPane();
		return 0;
	}

	return CallWindowProc(wpOrigProgmanProc, hWnd, uMsg, wParam, lParam);
}

//
// Network
//
static LRESULT CALLBACK NetworkProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	DWORD dw;
	static TCHAR *psz = NULL;

	dw = LOWORD(lParam);
	if (dw == WM_LBUTTONDOWN)
	{
		psz = TEXT("shell:::{7007ACC7-3202-11D1-AAD2-00805FC1270E}");
		return 0;
	}
	if (dw == WM_LBUTTONUP)
	{
		if (psz)
		{
			ShellExecute(NULL, NULL, psz, NULL, NULL, SW_SHOW);
			psz = NULL;
		}
		return 0;
	}

	return CallWindowProc(wpOrigNetworkProc, hWnd, uMsg, wParam, lParam);
}

//
// TrayShowDesktop
//
static LRESULT CALLBACK TrayShowDesktopProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	HDC hdc;
	PAINTSTRUCT ps;
	HWND htray;

	if (uMsg == WM_PAINT && fHideTrayShowDesktop)
	{
		hdc = BeginPaint(hWnd, &ps);
		htray = GetParent(hWnd);
		SendMessage(htray, WM_ERASEBKGND, (WPARAM)hdc, 0);
		SendMessage(htray, WM_PRINTCLIENT, (WPARAM)hdc, PRF_CLIENT);
		EndPaint(hWnd, &ps);
		return 0;
	}

	return CallWindowProc(wpOrigTrayShowDesktopProc, hWnd, uMsg, wParam, lParam);
}