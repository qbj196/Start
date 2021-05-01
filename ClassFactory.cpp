#include "ClassFactory.h"
#include "DeskBand.h"


extern LONG cDllRef;


CClassFactory::CClassFactory()
{
	m_cRef = 1;

	InterlockedIncrement(&cDllRef);
}

CClassFactory::~CClassFactory()
{
	InterlockedDecrement(&cDllRef);
}

//
// IUnknown
//
STDMETHODIMP CClassFactory::QueryInterface(REFIID riid, void **ppvObject)
{
	if (IsEqualIID(riid, IID_IUnknown))
		*ppvObject = this;
	else if (IsEqualIID(riid, IID_IClassFactory))
		*ppvObject = (IClassFactory *)this;
	else
		*ppvObject = NULL;

	if (*ppvObject)
	{
		AddRef();
		return S_OK;
	}

	return E_NOINTERFACE;
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

	*ppvObject = NULL;
	hr = CLASS_E_NOAGGREGATION;
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
		InterlockedIncrement(&cDllRef);
	else
		InterlockedDecrement(&cDllRef);

	return S_OK;
}