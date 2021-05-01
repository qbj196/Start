#include <Windows.h>
#include <ShlObj.h>
#include "ClassFactory.h"
#include "Resource.h"


HINSTANCE	hDllInst = NULL;
LONG		cDllRef = 0;
CLSID		CLSID_Start = {0x3f6953f0, 0x5359, 0x47fc, {0xbd, 0x99, 0x9f, 0x2c, 0xb9, 0x5a, 0x62, 0xff}};
TCHAR		szSetupKey[] = TEXT("Software\\Start");
TCHAR		szImsspKey[] = TEXT("Software\\Classes\\CLSID\\{23650F94-13B8-4F39-B2C3-817E6564A756}\\InProcServer32");
TCHAR		szStartKey[] = TEXT("Software\\Classes\\CLSID\\{3F6953F0-5359-47FC-BD99-9F2CB95A62FF}\\InprocServer32");


STDAPI_(BOOL) DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	if (dwReason == DLL_PROCESS_ATTACH)
	{
		hDllInst = hInstance;
		DisableThreadLibraryCalls(hDllInst);
	}

	return TRUE;
}

STDAPI DllCanUnloadNow()
{
	return cDllRef > 0 ? S_FALSE : S_OK;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
	HRESULT hr;
	CClassFactory *pcf;

	*ppv = NULL;
	hr = CLASS_E_CLASSNOTAVAILABLE;
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

STDAPI DllRegisterServer()
{
	HRESULT hr;
	TCHAR sz[MAX_PATH];
	DWORD cb;
	LSTATUS ls;
	TCHAR *psz;
	ICatRegister *pcr;
	CATID catid;

	hr = E_FAIL;
	pcr = NULL;

	cb = LoadString(hDllInst, IDS_START, sz, MAX_PATH);
	if (!cb)
		goto end;

	szStartKey[61] = TEXT('\0');
	cb *= sizeof(TCHAR);
	ls = RegSetKeyValue(HKEY_LOCAL_MACHINE, szStartKey, NULL, REG_SZ, sz, cb);
	if (ls != ERROR_SUCCESS)
		goto end;
	szStartKey[61] = TEXT('\\');
	
	cb = GetModuleFileName(hDllInst, sz, MAX_PATH);
	if (!cb)
		goto end;

	cb *= sizeof(TCHAR);
	ls = RegSetKeyValue(HKEY_LOCAL_MACHINE, szStartKey, NULL, REG_SZ, sz, cb);
	if (ls != ERROR_SUCCESS)
		goto end;

	psz = TEXT("ThreadingModel");
	lstrcpy(sz, TEXT("Apartment"));
	cb = lstrlen(sz) * sizeof(TCHAR);
	ls = RegSetKeyValue(HKEY_LOCAL_MACHINE, szStartKey, psz, REG_SZ, sz, cb);
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

	szImsspKey[61] = TEXT('\0');
	RegDeleteTree(HKEY_CURRENT_USER, szImsspKey);
	RegDeleteTree(HKEY_CURRENT_USER, szSetupKey);
	szStartKey[61] = TEXT('\0');
	ls = RegDeleteTree(HKEY_LOCAL_MACHINE, szStartKey);
	if (ls != ERROR_SUCCESS)
		hr = E_FAIL;

end:
	if (pcr)
		pcr->Release();

	return hr;
}