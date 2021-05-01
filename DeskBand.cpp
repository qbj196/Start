#include "DeskBand.h"
#include "Start.h"


extern LONG cDllRef;
extern CLSID CLSID_Start;


CDeskBand::CDeskBand()
{
	m_cRef = 1;
	m_dwBandID = 0;
	m_fCompositionEnabled = FALSE;

	InterlockedIncrement(&cDllRef);
}

CDeskBand::~CDeskBand()
{
	InterlockedDecrement(&cDllRef);
}

//
// IUnknown
//
STDMETHODIMP CDeskBand::QueryInterface(REFIID riid, void **ppvObject)
{
	if (IsEqualIID(riid, IID_IUnknown))
		*ppvObject = this;
	else if (IsEqualIID(riid, IID_IOleWindow))
		*ppvObject = (IOleWindow *)this;
	else if (IsEqualIID(riid, IID_IDockingWindow))
		*ppvObject = (IDockingWindow *)this;
	else if (IsEqualIID(riid, IID_IDeskBand))
		*ppvObject = (IDeskBand *)this;
	else if (IsEqualIID(riid, IID_IDeskBand2))
		*ppvObject = (IDeskBand2 *)this;
	else if (IsEqualIID(riid, IID_IPersist))
		*ppvObject = (IPersist *)this;
	else if (IsEqualIID(riid, IID_IPersistStream))
		*ppvObject = (IPersistStream *)this;
	else if (IsEqualIID(riid, IID_IObjectWithSite))
		*ppvObject = (IObjectWithSite *)this;
	else
		*ppvObject = NULL;

	if (*ppvObject)
	{
		AddRef();
		return S_OK;
	}

	return E_NOINTERFACE;
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
		hr = pUnkSite->QueryInterface(IID_IOleWindow, (void **)&pow);
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

STDMETHODIMP CDeskBand::GetSite(REFIID, void **ppvSite)
{
	*ppvSite = NULL;

	return E_NOTIMPL;
}