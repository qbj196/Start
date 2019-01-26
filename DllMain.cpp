#include <Windows.h>
#include <tchar.h>
#include <ShlObj.h>
#include "ClassFactory.h"


HINSTANCE	g_hDllInst = NULL;
LONG		g_cDllRef = NULL;
CLSID		g_StartID = {0x3f6953f0, 0x5359, 0x47fc, {0xbd, 0x99, 0x9f, 0x2c, 0xb9, 0x5a, 0x62, 0xff}};


STDAPI_(BOOL) DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	if (dwReason == DLL_PROCESS_ATTACH)
	{
		g_hDllInst = hInstance;
		DisableThreadLibraryCalls(hInstance);
	}

	return TRUE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
	HRESULT hr;
	CClassFactory *pcf;

	hr = CLASS_E_CLASSNOTAVAILABLE;
	pcf = NULL;

	if (IsEqualCLSID(g_StartID, rclsid))
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
	TCHAR szSubKey[MAX_PATH];
	TCHAR szData[MAX_PATH];
	DWORD cbData;
	HKEY hKey;
	ICatRegister *pcr;
	CATID catid;

	hr = E_FAIL;
	hKey = NULL;
	pcr = NULL;

	wsprintf(szSubKey,
		_T("CLSID\\{%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}"),
		g_StartID.Data1, g_StartID.Data2, g_StartID.Data3,
		g_StartID.Data4[0], g_StartID.Data4[1], g_StartID.Data4[2],
		g_StartID.Data4[3], g_StartID.Data4[4], g_StartID.Data4[5],
		g_StartID.Data4[6], g_StartID.Data4[7]);

	ls = RegCreateKeyEx(HKEY_CLASSES_ROOT, szSubKey, 0, NULL,
		REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
	if (ls != ERROR_SUCCESS)
	{
		goto end;
	}

	lstrcpy(szData, _T("开始(&S)"));
	cbData = lstrlen(szData) * sizeof(TCHAR);
	ls = RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE *)szData, cbData);
	if (ls != ERROR_SUCCESS)
	{
		goto end;
	}

	ls = RegCloseKey(hKey);
	if (ls != ERROR_SUCCESS)
	{
		goto end;
	}
	hKey = NULL;

	lstrcat(szSubKey, _T("\\InprocServer32"));
	ls = RegCreateKeyEx(HKEY_CLASSES_ROOT, szSubKey, 0, NULL,
		REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL);
	if (ls != ERROR_SUCCESS)
	{
		goto end;
	}
	
	cbData = GetModuleFileName(g_hDllInst, szData, ARRAYSIZE(szData));
	if (0 != GetLastError())
	{
		goto end;
	}
	cbData *= sizeof(TCHAR);

	ls = RegSetValueEx(hKey, NULL, 0, REG_SZ, (BYTE *)szData, cbData);
	if (ls != ERROR_SUCCESS)
	{
		goto end;
	}

	lstrcpy(szSubKey, _T("ThreadingModel"));
	lstrcpy(szData, _T("Apartment"));
	cbData = lstrlen(szData) + 1;
	cbData *= sizeof(TCHAR);
	ls = RegSetValueEx(hKey, szSubKey, 0, REG_SZ, (BYTE *)szData, cbData);
	if (ls != ERROR_SUCCESS)
	{
		goto end;
	}

	hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr, NULL,
		CLSCTX_INPROC_SERVER, IID_ICatRegister, (LPVOID *)&pcr);
	if (FAILED(hr))
	{
		goto end;
	}

	catid = CATID_DeskBand;
	hr = pcr->RegisterClassImplCategories(g_StartID, 1, &catid);

end:
	if (pcr)
	{
		pcr->Release();
		pcr = NULL;
	}
	if (hKey)
	{
		RegCloseKey(hKey);
		hKey = NULL;
	}

	return hr;
}

STDAPI DllUnregisterServer()
{
	HRESULT hr;
	LSTATUS ls;
	TCHAR szSubKey[MAX_PATH];
	ICatRegister *pcr;
	CATID catid;

	hr = E_FAIL;
	pcr = NULL;

	wsprintf(szSubKey,
		_T("CLSID\\{%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}"),
		g_StartID.Data1, g_StartID.Data2, g_StartID.Data3,
		g_StartID.Data4[0], g_StartID.Data4[1], g_StartID.Data4[2],
		g_StartID.Data4[3], g_StartID.Data4[4], g_StartID.Data4[5],
		g_StartID.Data4[6], g_StartID.Data4[7]);

	hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr, NULL,
		CLSCTX_INPROC_SERVER, IID_ICatRegister, (LPVOID *)&pcr);
	if (FAILED(hr))
	{
		goto end;
	}

	catid = CATID_DeskBand;
	hr = pcr->UnRegisterClassImplCategories(g_StartID, 1, &catid);
	if (FAILED(hr))
	{
		goto end;
	}

	ls = RegDeleteTree(HKEY_CLASSES_ROOT, szSubKey);
	if (ls != ERROR_SUCCESS)
	{
		hr = E_FAIL;
	}

end:
	if (pcr)
	{
		pcr->Release();
		pcr = NULL;
	}

	return hr;
}