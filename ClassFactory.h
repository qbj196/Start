#pragma once

#include <Windows.h>


class CClassFactory : public IClassFactory
{
public:
	// IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	// IClassFactory
	STDMETHODIMP CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject);
	STDMETHODIMP LockServer(BOOL fLock);

	CClassFactory();
protected:
	~CClassFactory();
private:
	LONG m_cRef;
};