#include "DeskBand.h"
#include "Start.h"


extern LONG g_cDllRef;
extern CLSID g_StartID;

DWORD g_dwBandID = 0;


CDeskBand::CDeskBand()
{
	m_cRef = 1;
	m_fIsDirty = FALSE;
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
	{
		AddRef();
	}

	return hr;
}

STDMETHODIMP_(ULONG) CDeskBand::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CDeskBand::Release()
{
	ULONG cRef;

	cRef = InterlockedDecrement(&m_cRef);
	if (cRef == 0)
	{
		delete this;
	}

	return cRef;
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
	g_dwBandID = dwBandID;

	return S_OK;
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
	*pClassID = g_StartID;

	return S_OK;
}

//
// IPersistStream
//
STDMETHODIMP CDeskBand::IsDirty()
{
	return m_fIsDirty ? S_OK : S_FALSE;
}

STDMETHODIMP CDeskBand::Load(IStream *pStm)
{
	return S_OK;
}

STDMETHODIMP CDeskBand::Save(IStream *pStm, BOOL fClearDirty)
{
	if (fClearDirty)
	{
		m_fIsDirty = FALSE;
	}

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
	HWND hWnd;

	hr = E_FAIL;
	pow = NULL;
	hWnd = NULL;

	if (pUnkSite)
	{
		hr = pUnkSite->QueryInterface(IID_IOleWindow, reinterpret_cast<void **>(&pow));
		if (SUCCEEDED(hr))
		{
			hr = pow->GetWindow(&hWnd);
			if (SUCCEEDED(hr))
			{
				hr = CreateStart(hWnd);
			}
		}
	}

	if (pow)
	{
		pow->Release();
		pow = NULL;
	}

	return hr;
}

STDMETHODIMP CDeskBand::GetSite(REFIID riid, void **ppvSite)
{
	*ppvSite = NULL;

	return E_FAIL;
}