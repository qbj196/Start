#include "Start.h"


extern HINSTANCE g_hDllInst;
extern DWORD g_dwBandID;

HWND m_hStart = NULL;
HWND m_hRebar = NULL;
HWND m_hTaskbar = NULL;
BOOL m_fShowDesktop = TRUE;
WNDPROC m_wpOrigTaskbarProc = NULL;


// Start
static LRESULT CALLBACK StartProc(HWND, UINT, WPARAM, LPARAM);
void OnStartShowWindow(HWND, BOOL);

// Taskbar
static LRESULT CALLBACK TaskbarProc(HWND, UINT, WPARAM, LPARAM);
void OnTaskbarShowDesktop(HWND);


HRESULT CreateStart(HWND hWndParent)
{
	HRESULT hr;
	WNDCLASSEX wcex;

	hr = E_FAIL;
	m_hRebar = hWndParent;
	m_hTaskbar = GetParent(hWndParent);

	wcex.cbClsExtra		= 0;
	wcex.cbSize			= sizeof(WNDCLASSEX);
	wcex.cbWndExtra		= 0;
	wcex.hbrBackground	= NULL;
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hIcon			= NULL;
	wcex.hIconSm		= NULL;
	wcex.hInstance		= g_hDllInst;
	wcex.lpfnWndProc	= StartProc;
	wcex.lpszClassName	= _T("Start");
	wcex.lpszMenuName	= NULL;
	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	if (!RegisterClassEx(&wcex))
		goto end;

	m_hStart = CreateWindowEx(0, wcex.lpszClassName, NULL,
		WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		0, 0, 0, 0, hWndParent, NULL, wcex.hInstance, NULL);
	if (!m_hStart)
		goto end;

	m_wpOrigTaskbarProc = (WNDPROC)SetWindowLongPtr(
		m_hTaskbar, GWLP_WNDPROC, (LONG_PTR)TaskbarProc);
	if (!m_wpOrigTaskbarProc)
		goto end;

	hr = S_OK;

end:
	return hr;
}

void CloseStart()
{
	if (m_wpOrigTaskbarProc)
	{
		SetWindowLongPtr(m_hTaskbar, GWLP_WNDPROC, (LONG_PTR)m_wpOrigTaskbarProc);
		m_wpOrigTaskbarProc = NULL;
	}

	if (m_hStart)
	{
		ShowWindow(m_hStart, SW_HIDE);
		DestroyWindow(m_hStart);
		m_hStart = NULL;
	}

	UnregisterClass(_T("Start"), g_hDllInst);

	PostMessage(m_hTaskbar, WM_TIMER, 24, 0);
}

//
// Start
//
static LRESULT CALLBACK StartProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_SHOWWINDOW:
		OnStartShowWindow(hWnd, (BOOL)wParam);
		break;
	default:
		break;
	}

	return DefWindowProc(hWnd, uMsg, wParam, lParam);
}

void OnStartShowWindow(HWND hWnd, BOOL fShow)
{
	LRESULT id;

	if (fShow)
	{
		id = SendMessage(m_hRebar, RB_IDTOINDEX, g_dwBandID, 0);
		SendMessage(m_hRebar, RB_SHOWBAND, (WPARAM)id, FALSE);
	}
}

//
// Taskbar
//
static LRESULT CALLBACK TaskbarProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case 0x0579:
		OnTaskbarShowDesktop(hWnd);
		break;
	default:
		break;
	}

	return CallWindowProc(m_wpOrigTaskbarProc, hWnd, uMsg, wParam, lParam);
}

void OnTaskbarShowDesktop(HWND hWnd)
{
	HRESULT hr;
	IAppVisibility *pav;
	BOOL flv;

	pav = NULL;

	if (!m_fShowDesktop)
		return;

	hr = CoCreateInstance(CLSID_AppVisibility, NULL, CLSCTX_INPROC_SERVER,
		IID_IAppVisibility, (LPVOID *)&pav);
	if (SUCCEEDED(hr))
	{
		hr = pav->IsLauncherVisible(&flv);
		if (SUCCEEDED(hr))
		{
			m_fShowDesktop = FALSE;
			if (flv)
			{
				SetForegroundWindow(hWnd);
			}
		}
	}

	if (pav)
	{
		pav->Release();
		pav = NULL;
	}
}