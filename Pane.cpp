#include "Pane.h"
#include "Resource.h"


extern HINSTANCE g_hDllInst;
extern HWND hStartWnd;
extern HWND hRebarWnd;
extern HWND hTaskbarWnd;

HWND hPaneWnd = NULL;
HWND hForeWnd = NULL;
HICON hTabsIcon = NULL;
HICON hProgsIcon = NULL;
HFONT hFont = NULL;
TCHAR szProgs[13][50];
INT cchProgs[13];
INT nMouseMove = 0;
INT nLButtonDown = 0;

RECT rcPane[23] = {
	  0,   0,   0,   0,
	  0,   0, 300, 480,
	  0,   0,  48, 480,
	 48,   0, 300, 480,

	  0,   0,  48,  48,
	  0,  48,  48,  96,
	  0,  96,  48, 144,

	 50,   2, 286,  38,
	 50,  38, 286,  74,
	 50,  74, 286, 110,
	 50, 110, 286, 146,
	 50, 146, 286, 182,
	 50, 182, 286, 218,
	 50, 218, 286, 254,
	 50, 254, 286, 290,
	 50, 290, 286, 326,
	 50, 326, 286, 362,
	 50, 362, 286, 398,
	 50, 398, 286, 434,
	 48, 444, 300, 480,
	288,   0, 300, 440,

	 48,  36, 144, 432,
	 52,   0,  84,   0,
};


// Pane
static LRESULT CALLBACK PaneProc(HWND, UINT, WPARAM, LPARAM);
void OnPanePaint();
void OnPaneMouseMove(WPARAM, LPARAM);
void OnPaneLButtonDown(LPARAM);
void OnPaneLButtonUp();
void OpenExecute(BOOL, WPARAM);


HRESULT CreatePane()
{
	WNDCLASSEX wcex;
	int i;
	//float f=1.25;for(i=0;i<23;i++){rcPane[i].left*=f;rcPane[i].top*=f;rcPane[i].right*=f;rcPane[i].bottom*=f;}

	for (i=0; i<13; i++)
		cchProgs[i] = LoadString(g_hDllInst, i + IDS_PROG1, szProgs[i], 50);

	hTabsIcon = (HICON)LoadImage(g_hDllInst, MAKEINTRESOURCE(IDI_TABS),
		IMAGE_ICON, (int)rcPane[21].left, (int)rcPane[21].right, LR_DEFAULTCOLOR);
	if (!hTabsIcon)
		return E_FAIL;

	hProgsIcon = (HICON)LoadImage(g_hDllInst, MAKEINTRESOURCE(IDI_PROGS),
		IMAGE_ICON, (int)rcPane[21].top, (int)rcPane[21].bottom, LR_DEFAULTCOLOR);
	if (!hProgsIcon)
		return E_FAIL;

	wcex.cbClsExtra		= 0;
	wcex.cbSize			= sizeof(WNDCLASSEX);
	wcex.cbWndExtra		= 0;
	wcex.hbrBackground	= NULL;
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hIcon			= NULL;
	wcex.hIconSm		= NULL;
	wcex.hInstance		= g_hDllInst;
	wcex.lpfnWndProc	= PaneProc;
	wcex.lpszClassName	= TEXT("Pane");
	wcex.lpszMenuName	= NULL;
	wcex.style			= CS_VREDRAW | CS_HREDRAW;
	if (!RegisterClassEx(&wcex))
		return E_FAIL;

	hPaneWnd = CreateWindowEx(WS_EX_TOOLWINDOW | WS_EX_TOPMOST, TEXT("Pane"),
		NULL, WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		0, 0, 0, 0, hTaskbarWnd, NULL, g_hDllInst, NULL);
	if (!hPaneWnd)
		return E_FAIL;

	return S_OK;
}

void ShowPane()
{
	RECT rc;
	LONG x, y;
	int w, h;

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
		if (rc.left == 0)
		{
			if (rc.top == 0)
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

		hForeWnd = GetForegroundWindow();
		hFont = (HFONT)SendMessage(hRebarWnd, WM_GETFONT, 0, 0);
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

	UnregisterClass(TEXT("Pane"), g_hDllInst);
}

//
// Start
//
static LRESULT CALLBACK PaneProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_PAINT:
		OnPanePaint();
		break;
	case WM_KILLFOCUS:
		nMouseMove = 0;
		nLButtonDown = 0;
		ShowWindow(hWnd, SW_HIDE);
		break;
	case WM_SHOWWINDOW:
		EnableWindow(hStartWnd, !wParam);
		break;
	case WM_MOUSEMOVE:
		OnPaneMouseMove(wParam, lParam);
		break;
	case WM_MOUSELEAVE:
		nMouseMove = 0;
		InvalidateRect(hWnd, NULL, TRUE);
		break;
	case WM_LBUTTONDOWN:
		SetCapture(hWnd);
		OnPaneLButtonDown(lParam);
		break;
	case WM_LBUTTONUP:
		ReleaseCapture();
		OpenExecute(TRUE, 0);
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
	int i, cx, cy;

	hdc = BeginPaint(hPaneWnd, &ps);
	if (hdc)
	{
		hbufpt = BeginBufferedPaint(hdc, &rcPane[1], BPBF_TOPDOWNDIB, NULL, &hbufdc);
		if (hbufpt)
		{
			FillRect(hbufdc, &rcPane[2], (HBRUSH)(COLOR_GRAYTEXT + 1));
			FillRect(hbufdc, &rcPane[3], (HBRUSH)(COLOR_BTNSHADOW + 1));
			if (nMouseMove > 3)
				FillRect(hbufdc, &rcPane[nMouseMove], (HBRUSH)(COLOR_ACTIVEBORDER + 1));
			if (nLButtonDown > 3)
				FillRect(hbufdc, &rcPane[nLButtonDown], (HBRUSH)(COLOR_SCROLLBAR + 1));

			SetBkMode(hbufdc, TRANSPARENT);
			SelectObject(hbufdc, hFont);
			cx = rcPane[7].bottom - rcPane[7].top;
			cy = cx - rcPane[7].top - rcPane[7].top;
			for(i=7; i<19; i++)
			{
				rcPane[22].top = rcPane[i].top + rcPane[7].top;
				rcPane[22].bottom = rcPane[22].top + cy;
				FillRect(hbufdc, &rcPane[22], (HBRUSH)(COLOR_GRAYTEXT + 1));

				rcPane[i].left = rcPane[i].left + cx;
				DrawText(hbufdc, szProgs[i-7], cchProgs[i-7], &rcPane[i], DT_SINGLELINE | DT_VCENTER);
				rcPane[i].left = rcPane[i].left - cx;
			}
			DrawText(hbufdc, szProgs[i-7], cchProgs[i-7], &rcPane[i], DT_SINGLELINE | DT_VCENTER | DT_CENTER);

			DrawIconEx(hbufdc, 0, 0, hTabsIcon, (int)rcPane[21].left, (int)rcPane[21].right, 0, NULL, DI_NORMAL);
			DrawIconEx(hbufdc, (int)rcPane[7].left, (int)rcPane[7].top, hProgsIcon,
				(int)rcPane[21].top, (int)rcPane[21].bottom, 0, NULL, DI_NORMAL);

			EndBufferedPaint(hbufpt, TRUE);
		}

		EndPaint(hPaneWnd, &ps);
	}
}

void OnPaneMouseMove(WPARAM wParam, LPARAM lParam)
{
	TRACKMOUSEEVENT tme;
	POINT pt;
	int i;

	if (wParam)
		return;

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
	for (i=20; i>0; i--)
	{
		if (PtInRect(&rcPane[i], pt))
		{
			if (i != nMouseMove)
			{
				nMouseMove = i;
				InvalidateRect(hPaneWnd, NULL, TRUE);
			}
			break;
		}
	}
}

void OnPaneLButtonDown(LPARAM lParam)
{
	POINT pt;
	int i;

	pt.x = GET_X_LPARAM(lParam);
	pt.y = GET_Y_LPARAM(lParam);
	for (i=20; i>0; i--)
	{
		if (PtInRect(&rcPane[i], pt))
		{
			nLButtonDown = i;
			InvalidateRect(hPaneWnd, NULL, TRUE);
			break;
		}
	}
}

void OpenExecute(BOOL fOpen, WPARAM wParam)
{
	TCHAR *pszop;
	TCHAR *pszfile;
	TCHAR *pszparam;

	pszop = NULL;
	pszfile = NULL;
	pszparam = NULL;

	switch (wParam >= 'A' ? wParam : nLButtonDown)
	{
	case 'T':
		nLButtonDown = 7;
	case 7:
		pszfile = TEXT("regedit.exe");
		break;
	case 'P':
		nLButtonDown = 8;
	case 8:
		pszfile = TEXT("control.exe");
		break;
	case 'C':
		nLButtonDown = 9;
	case 9:
		pszfile = TEXT("cmd.exe");
		break;
	case 'A':
		nLButtonDown = 10;
	case 10:
		pszop = TEXT("runas");
		pszfile = TEXT("cmd.exe");
		break;
	case 'E':
		nLButtonDown = 11;
	case 11:
		pszfile = TEXT("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}");
		break;
	case 'D':
		nLButtonDown = 12;
	case 12:
		pszfile = TEXT("devmgmt.msc");
		break;
	case 'M':
		nLButtonDown = 13;
	case 13:
		pszfile = TEXT("perfmon.exe");
		pszparam = TEXT("/res");
		break;
	case 'Y':
		nLButtonDown = 14;
	case 14:
		pszfile = TEXT("msconfig.exe");
		break;
	case 'I':
		nLButtonDown = 15;
	case 15:
		pszfile = TEXT("msinfo32.exe");
		break;
	case 'N':
		nLButtonDown = 16;
	case 16:
		pszfile = TEXT("shell:::{7007ACC7-3202-11D1-AAD2-00805FC1270E}");
		break;
	case 'S':
		nLButtonDown = 17;
	case 17:
		pszfile = TEXT("services.msc");
		break;
	case 'R':
		nLButtonDown = 18;
	case 18:
		pszfile = TEXT("shell:::{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}");
		break;
	case 'U':
		nLButtonDown = 19;
	case 19:
		if (fOpen)
			PostMessage(hTaskbarWnd, 0x556, 0, 0); //launch shut down windows dialog
		break;
	default:
		nLButtonDown = 0;
		break;
	}

	if (fOpen)
	{
		if (pszfile)
			ShellExecute(NULL, pszop, pszfile, pszparam, NULL, SW_SHOW);

		nLButtonDown = 0;
	}

	InvalidateRect(hPaneWnd, NULL, TRUE);
}