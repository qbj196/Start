#include <Windows.h>
#include <ShlObj.h>
#include "ClassFactory.h"
#include "Resource.h"


HINSTANCE	g_hDllInst = NULL;
LONG		g_cDllRef = 0;
CLSID		CLSID_Start = {0x3f6953f0, 0x5359, 0x47fc, {0xbd, 0x99, 0x9f, 0x2c, 0xb9, 0x5a, 0x62, 0xff}};
TCHAR		szStartKey[] = TEXT("CLSID\\{3F6953F0-5359-47FC-BD99-9F2CB95A62FF}");


STDAPI_(BOOL) DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	if (dwReason == DLL_PROCESS_ATTACH)
	{
		g_hDllInst = hInstance;
		DisableThreadLibraryCalls(g_hDllInst);
	}

	return TRUE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
	HRESULT hr;
	CClassFactory *pcf;

	hr = CLASS_E_CLASSNOTAVAILABLE;
	*ppv = NULL;

	if (IsEqualCLSID(CLSID_Start, rclsid))
	{
		hr = E_OUTOFMEMORY;

		pcf = new CClassFactory();
		if (pcf)
		{
			hr = pcf->QueryInterface(riid, ppv);
			pcf->Release();
		}
	}

	return hr;
}

STDAPI DllCanUnloadNow()
{
	return g_cDllRef > 0 ? S_FALSE : S_OK;
}

STDAPI DllRegisterServer()
{
	HRESULT hr;
	LSTATUS ls;
	HKEY hkey;
	TCHAR sz[MAX_PATH];
	DWORD cb;
	ICatRegister *pcr;
	CATID catid;

	hr = E_FAIL;
	hkey = NULL;
	pcr = NULL;

	ls = RegCreateKeyEx(HKEY_CLASSES_ROOT, szStartKey, 0, NULL,
		REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hkey, NULL);
	if (ls != ERROR_SUCCESS)
		goto end;

	cb = LoadString(g_hDllInst, IDS_START, sz, MAX_PATH);
	if (GetLastError() != ERROR_SUCCESS)
		goto end;

	cb *= sizeof(TCHAR);
	ls = RegSetKeyValue(hkey, NULL, NULL, REG_SZ, sz, cb);
	if (ls != ERROR_SUCCESS)
		goto end;
	
	cb = GetModuleFileName(g_hDllInst, sz, MAX_PATH);
	if (GetLastError() != ERROR_SUCCESS)
		goto end;

	cb *= sizeof(TCHAR);
	ls = RegSetKeyValue(hkey, TEXT("InprocServer32"), NULL, REG_SZ, sz, cb);
	if (ls != ERROR_SUCCESS)
		goto end;

	lstrcpy(sz, TEXT("Apartment"));
	cb = lstrlen(sz) * sizeof(TCHAR);
	ls = RegSetKeyValue(hkey, TEXT("InprocServer32"), TEXT("ThreadingModel"),
		REG_SZ, sz, cb);
	if (ls != ERROR_SUCCESS)
		goto end;

	hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr, NULL,
		CLSCTX_INPROC_SERVER, IID_ICatRegister, (LPVOID *)&pcr);
	if (FAILED(hr))
		goto end;

	catid = CATID_DeskBand;
	hr = pcr->RegisterClassImplCategories(CLSID_Start, 1, &catid);

end:
	if (pcr)
		pcr->Release();
	if (hkey)
		RegCloseKey(hkey);

	return hr;
}

STDAPI DllUnregisterServer()
{
	HRESULT hr;
	ICatRegister *pcr;
	CATID catid;
	LSTATUS ls;

	hr = E_FAIL;
	pcr = NULL;

	hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr, NULL,
		CLSCTX_INPROC_SERVER, IID_ICatRegister, (LPVOID *)&pcr);
	if (FAILED(hr))
		goto end;

	catid = CATID_DeskBand;
	hr = pcr->UnRegisterClassImplCategories(CLSID_Start, 1, &catid);
	if (FAILED(hr))
		goto end;

	ls = RegDeleteTree(HKEY_CLASSES_ROOT, szStartKey);
	if (ls != ERROR_SUCCESS)
		hr = E_FAIL;

end:
	if (pcr)
		pcr->Release();

	return hr;
}