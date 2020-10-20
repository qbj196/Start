#pragma once

#include <Windows.h>
#include <ShlObj.h>


class CDeskBand : public IDeskBand2, IPersistStream, IObjectWithSite
{
public:
	// IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	// IOleWindow
	STDMETHODIMP GetWindow(HWND *phWnd);
	STDMETHODIMP ContextSensitiveHelp(BOOL fEnterMode);

	// IDockingWindow
	STDMETHODIMP ShowDW(BOOL fShow);
	STDMETHODIMP CloseDW(DWORD dwReserved);
	STDMETHODIMP ResizeBorderDW(LPCRECT prcBorder, IUnknown *punkToolbarSite, BOOL fReserved);

	// IDeskBand
	STDMETHODIMP GetBandInfo(DWORD dwBandID, DWORD dwViewMode, DESKBANDINFO *pdbi);

	// IDeskBand2
	STDMETHODIMP CanRenderComposited(BOOL *pfCanRenderComposited);
	STDMETHODIMP SetCompositionState(BOOL fCompositionEnabled);
	STDMETHODIMP GetCompositionState(BOOL *pfCompositionEnabled);

	// IPersist
	STDMETHODIMP GetClassID(CLSID *pClassID);

	// IPersistStream
	STDMETHODIMP IsDirty();
	STDMETHODIMP Load(IStream *pStm);
	STDMETHODIMP Save(IStream *pStm, BOOL fClearDirty);
	STDMETHODIMP GetSizeMax(ULARGE_INTEGER *pcbSize);

	// IObjectWithSite
	STDMETHODIMP SetSite(IUnknown *pUnkSite);
	STDMETHODIMP GetSite(REFIID riid, void **ppvSite);

	CDeskBand();
protected:
	~CDeskBand();
private:
	LONG m_cRef;
	DWORD m_dwBandID;
	BOOL m_fCompositionEnabled;
};