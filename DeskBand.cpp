#include "DeskBand.h"
#include "Start.h"


extern LONG g_cDllRef;
extern CLSID CLSID_Start;


CDeskBand::CDeskBand()
{
	m_cRef = 1;
	m_dwBandID = 0;
	m_fCompositionEnabled = FALSE;

	InterlockedIncrement(&g_cDllRef);
}

CDeskBand::~CDeskBand()
{
	InterlockedDecrement(&g_cDllRef);
}

//
// IUnknown
//
STDMETHODIMP CDeskBand::QueryInterface(REFIID riid, void **ppvObject)
{
	HRESULT hr;

	hr = S_OK;

	if (IsEqualIID(IID_IUnknown, riid)			||
		IsEqualIID(IID_IOleWindow, riid)		||
		IsEqualIID(IID_IDockingWindow, riid)	||
		IsEqualIID(IID_IDeskBand, riid)			||
		IsEqualIID(IID_IDeskBand2, riid))
	{
		*ppvObject = static_cast<IOleWindow *>(this);
	}
	else if (IsEqualIID(IID_IPersist, riid) ||
		IsEqualIID(IID_IPersistStream, riid))
	{
		*ppvObject = static_cast<IPersist *>(this);
	}
	else if (IsEqualIID(IID_IObjectWithSite, riid))
	{
		*ppvObject = static_cast<IObjectWithSite *>(this);
	}
	else
	{
		hr = E_NOINTERFACE;
		*ppvObject = NULL;
	}

	if (*ppvObject)
		AddRef();

	return hr;
}

STDMETHODIMP_(ULONG) CDeskBand::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CDeskBand::Release()
{
	ULONG cref;

	cref = InterlockedDecrement(&m_cRef);
	if (cref == 0)
		delete this;

	return cref;
}

//
// IOleWindow
//
STDMETHODIMP CDeskBand::GetWindow(HWND *phWnd)
{
	*phWnd = NULL;

	return S_OK;
}

STDMETHODIMP CDeskBand::ContextSensitiveHelp(BOOL fEnterMode)
{
	return E_NOTIMPL;
}

//
// IDockingWindow
//
STDMETHODIMP CDeskBand::ShowDW(BOOL fShow)
{
	ShowStart(fShow, m_dwBandID);

	return S_OK;
}

STDMETHODIMP CDeskBand::CloseDW(DWORD dwReserved)
{
	CloseStart();

	return S_OK;
}

STDMETHODIMP CDeskBand::ResizeBorderDW(LPCRECT prcBorder, IUnknown *punkToolbarSite, BOOL fReserved)
{
	return E_NOTIMPL;
}

//
// IDeskBand
//
STDMETHODIMP CDeskBand::GetBandInfo(DWORD dwBandID, DWORD dwViewMode, DESKBANDINFO *pdbi)
{
	m_dwBandID = dwBandID;

	return E_NOTIMPL;
}

//
// IDeskBand2
//
STDMETHODIMP CDeskBand::CanRenderComposited(BOOL *pfCanRenderComposited)
{
	*pfCanRenderComposited = TRUE;

	return S_OK;
}

STDMETHODIMP CDeskBand::SetCompositionState(BOOL fCompositionEnabled)
{
	m_fCompositionEnabled = fCompositionEnabled;

	UpdateStart();

	return S_OK;
}

STDMETHODIMP CDeskBand::GetCompositionState(BOOL *pfCompositionEnabled)
{
	*pfCompositionEnabled = m_fCompositionEnabled;

	return S_OK;
}

//
// IPersist
//
STDMETHODIMP CDeskBand::GetClassID(CLSID *pClassID)
{
	*pClassID = CLSID_Start;

	return S_OK;
}

//
// IPersistStream
//
STDMETHODIMP CDeskBand::IsDirty()
{
	return S_FALSE;
}

STDMETHODIMP CDeskBand::Load(IStream *pStm)
{
	return S_OK;
}

STDMETHODIMP CDeskBand::Save(IStream *pStm, BOOL fClearDirty)
{
	return S_OK;
}

STDMETHODIMP CDeskBand::GetSizeMax(ULARGE_INTEGER *pcbSize)
{
	return E_NOTIMPL;
}

//
// IObjectWithSite
//
STDMETHODIMP CDeskBand::SetSite(IUnknown *pUnkSite)
{
	HRESULT hr;
	IOleWindow *pow;
	HWND hwnd;

	hr = E_FAIL;

	if (pUnkSite)
	{
		hr = pUnkSite->QueryInterface(IID_IOleWindow,
			reinterpret_cast<void **>(&pow));
		if (SUCCEEDED(hr))
		{
			hr = pow->GetWindow(&hwnd);
			if (SUCCEEDED(hr))
				hr = CreateStart(hwnd);

			pow->Release();
		}
	}

	return hr;
}

STDMETHODIMP CDeskBand::GetSite(REFIID riid, void **ppvSite)
{
	*ppvSite = NULL;

	return E_FAIL;
}