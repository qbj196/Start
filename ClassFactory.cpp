#include "ClassFactory.h"
#include "DeskBand.h"


extern LONG g_cDllRef;


CClassFactory::CClassFactory()
{
	m_cRef = 1;

	InterlockedIncrement(&g_cDllRef);
}

CClassFactory::~CClassFactory()
{
	InterlockedDecrement(&g_cDllRef);
}

//
// IUnknown
//
STDMETHODIMP CClassFactory::QueryInterface(REFIID riid, void **ppvObject)
{
	HRESULT hr;

	hr = S_OK;

	if (IsEqualIID(IID_IUnknown, riid) || IsEqualIID(IID_IClassFactory, riid))
	{
		*ppvObject = static_cast<IUnknown *>(this);
		AddRef();
	}
	else
	{
		hr = E_NOINTERFACE;
		*ppvObject = NULL;
	}

	return hr;
}

STDMETHODIMP_(ULONG) CClassFactory::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CClassFactory::Release()
{
	ULONG cref;

	cref = InterlockedDecrement(&m_cRef);
	if (cref == 0)
		delete this;

	return cref;
}

//
// IClassFactory
//
STDMETHODIMP CClassFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject)
{
	HRESULT hr;
	CDeskBand *pdb;

	hr = CLASS_E_NOAGGREGATION;
	*ppvObject = NULL;

	if (!pUnkOuter)
	{
		hr = E_OUTOFMEMORY;

		pdb = new CDeskBand();
		if (pdb)
		{
			hr = pdb->QueryInterface(riid, ppvObject);
			pdb->Release();
		}
	}

	return hr;
}

STDMETHODIMP CClassFactory::LockServer(BOOL fLock)
{
	if (fLock)
		InterlockedIncrement(&g_cDllRef);
	else
		InterlockedDecrement(&g_cDllRef);

	return S_OK;
}