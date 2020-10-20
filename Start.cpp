#include "Start.h"
#include "Pane.h"
#include "Resource.h"


extern HINSTANCE g_hDllInst;

HWND hOrigStartWnd = NULL;
HWND hStartWnd = NULL;
HWND hRebarWnd = NULL;
HWND hTaskSwWnd = NULL;
HWND hTaskbarWnd = NULL;
HWND hProgmanWnd = NULL;
HWND hWorkervWnd = NULL;
HWND hWorkerwWnd = NULL;
BOOL fEnterDesktop = FALSE;
BOOL fStartEnabled = TRUE;
BOOL fMouseLeftWnd = TRUE;
BOOL fInterceptMsg = FALSE;
UINT uShellHookMsg = 0;
HICON hStartIcon16 = NULL;
HICON hStartIcon32 = NULL;
WNDPROC wpOrigRebarProc = NULL;
WNDPROC wpOrigTaskbarProc = NULL;
WNDPROC wpOrigProgmanProc = NULL;
WNDPROC wpOrigWorkerwProc = NULL;


// Start
static LRESULT CALLBACK StartProc(HWND, UINT, WPARAM, LPARAM);
void OnStartPaint();
void OnStartMouseMove();

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

// Workerw
static LRESULT CALLBACK WorkerwProc(HWND, UINT, WPARAM, LPARAM);


HRESULT CreateStart(HWND hWnd)
{
	WNDCLASSEX wcex;

	hRebarWnd = hWnd;
	hTaskbarWnd = GetParent(hWnd);

	uShellHookMsg = RegisterWindowMessage(TEXT("SHELLHOOK"));
	if (!uShellHookMsg)
		return E_FAIL;

	hTaskSwWnd = FindWindowEx(hRebarWnd, NULL, TEXT("MSTaskSwWClass"), NULL);
	if (!hTaskSwWnd)
		return E_FAIL;
	
	hOrigStartWnd = FindWindowEx(hTaskbarWnd, NULL, TEXT("Start"), NULL);
	if (hOrigStartWnd)
		ShowWindow(hOrigStartWnd, SW_HIDE);

	hStartIcon16 = (HICON)LoadImage(g_hDllInst, MAKEINTRESOURCE(IDI_START),
		IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR);
	if (!hStartIcon16)
		return E_FAIL;

	hStartIcon32 = (HICON)LoadImage(g_hDllInst, MAKEINTRESOURCE(IDI_START),
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
	wcex.hInstance		= g_hDllInst;
	wcex.lpfnWndProc	= StartProc;
	wcex.lpszClassName	= TEXT("Start");
	wcex.lpszMenuName	= NULL;
	wcex.style			= CS_VREDRAW | CS_HREDRAW;
	if (!RegisterClassEx(&wcex))
		return E_FAIL;

	hStartWnd = CreateWindowEx(0, TEXT("Start"), NULL,
		WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		0, 0, 0, 0, hTaskbarWnd, NULL, g_hDllInst, NULL);
	if (!hStartWnd)
		return E_FAIL;

	wpOrigRebarProc = (WNDPROC)SetWindowLongPtr(hRebarWnd,
		GWLP_WNDPROC, (LONG_PTR)RebarProc);
	if (!wpOrigRebarProc)
		return E_FAIL;

	wpOrigTaskbarProc = (WNDPROC)SetWindowLongPtr(hTaskbarWnd,
		GWLP_WNDPROC, (LONG_PTR)TaskbarProc);
	if (!wpOrigTaskbarProc)
		return E_FAIL;

	return CreatePane();
}

void ShowStart(BOOL fShow, DWORD dwBandID)
{
	LRESULT index;

	if (hStartWnd)
	{
		index = SendMessage(hRebarWnd, RB_IDTOINDEX, dwBandID, 0);
		SendMessage(hRebarWnd, RB_SHOWBAND, index, FALSE); //hide this band
		SendMessage(hTaskbarWnd, WM_SIZE, 0, 0);
		ShowWindow(hStartWnd, fShow ? SW_SHOW : SW_HIDE);
	}
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

	if (wpOrigWorkerwProc)
	{
		SetWindowLongPtr(hWorkerwWnd, GWLP_WNDPROC,
			(LONG_PTR)wpOrigWorkerwProc);
		wpOrigWorkerwProc = NULL;
	}

	if (hWorkerwWnd)
	{
		OnTaskbarShowEdgeui(TRUE);
		hWorkerwWnd = NULL;
	}

	if (hWorkervWnd)
	{
		ShowWindow(hWorkervWnd, SW_SHOW);
		hWorkervWnd = NULL;
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

	if (hStartWnd)
	{
		DestroyWindow(hStartWnd);
		hStartWnd = NULL;
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

	if (hOrigStartWnd)
	{
		ShowWindow(hOrigStartWnd, SW_SHOW);
		hOrigStartWnd = NULL;
	}

	UnregisterClass(TEXT("Start"), g_hDllInst);

	PostMessage(hTaskbarWnd, WM_SIZE, 0, 0);
	PostMessage(hTaskbarWnd, WM_TIMER, 24, 0); //release dll message
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
	case WM_RBUTTONUP: //show start screen or start menu
		PostMessage(hTaskSwWnd, uShellHookMsg, HSHELL_TASKMAN, 0);
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
	case 0x0579: //switch window will receive this
		OnTaskbarEnterDesktop();
		break;
	case 0x05b3: //win8.0 explorer launch will receive this
		fEnterDesktop = TRUE;
	case 0x0550: //this band launch will receive this
		OnTaskbarSetWindowProc();
		break;
	case 0x05bb: //switch to start screen or desktop will receive this
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
	HMODULE hdll;
	HMODULE hproc;
	LONG_PTR lproc;
	HWND hwnd;

	if (!wpOrigProgmanProc)
	{
		hProgmanWnd = GetShellWindow();
		if (!hProgmanWnd)
			return;
		
		wpOrigProgmanProc = (WNDPROC)SetWindowLongPtr(hProgmanWnd,
			GWLP_WNDPROC, (LONG_PTR)ProgmanProc);
	}

	if (!hWorkervWnd)
	{
		hWorkervWnd = FindWindowEx(hProgmanWnd, NULL, TEXT("WorkerW"), NULL);
		if (hWorkervWnd)
			ShowWindow(hWorkervWnd, SW_HIDE);
	}

	if (!wpOrigWorkerwProc)
	{
		hdll = GetModuleHandle(TEXT("twinui.appcore.dll")); //win8.1
		if (!hdll)
		{
			hdll = GetModuleHandle(TEXT("twinui.dll")); //win8.0
			if (!hdll)
				return;
		}

		hwnd = NULL;
		for (;;)
		{
			hwnd = FindWindowEx(NULL, hwnd, TEXT("WorkerW"), NULL);
			if (!hwnd)
				break;

			lproc = GetWindowLongPtr(hwnd, GWLP_WNDPROC);
			if (GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
				(LPCWSTR)lproc, &hproc))
			{
				if (hproc == hdll)
					hWorkerwWnd = hwnd;
				FreeLibrary(hproc);
			}
		}
		if (!hWorkerwWnd)
			return;
		
		wpOrigWorkerwProc = (WNDPROC)SetWindowLongPtr(hWorkerwWnd,
			GWLP_WNDPROC, (LONG_PTR)WorkerwProc);
		if (!wpOrigWorkerwProc)
			return;

		if (!fEnterDesktop)
			OnTaskbarShowEdgeui(FALSE);
	}
}

void OnTaskbarShowEdgeui(BOOL fShow)
{
	LPARAM lparam;

	if (hWorkerwWnd)
	{
		lparam = (LPARAM)hTaskbarWnd;
		fInterceptMsg = FALSE; //0x35|0x36: enter|exit fullscreen edgeui will be hide|show
		SendMessage(hWorkerwWnd, uShellHookMsg, fShow ? 0x36 : 0x35, lparam);
		SendMessage(hWorkerwWnd, uShellHookMsg, HSHELL_RUDEAPPACTIVATED, lparam);
		fInterceptMsg = TRUE;
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
// Workerw
//
static LRESULT CALLBACK WorkerwProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	if (uMsg == uShellHookMsg)
	{
		switch (wParam)
		{
		case 0x35:
		case 0x36:
		case HSHELL_RUDEAPPACTIVATED:
			if (fInterceptMsg)
				return 0;
			break;
		default:
			break;
		}
	}

	return CallWindowProc(wpOrigWorkerwProc, hWnd, uMsg, wParam, lParam);
}