#pragma once

#include <Windows.h>
#include <windowsx.h>
#include <Uxtheme.h>
#include <ShlObj.h>
#include <commoncontrols.h>

#pragma comment(lib, "uxtheme.lib")


struct FILELIST {
	HICON hicon;
	TCHAR *pszdisp;
	TCHAR *pszpath1;
	TCHAR *pszpath2;

	INT nfldr;
	INT nfile;
	BOOL bfldr;
	BOOL bshow;
	FILELIST *pfl;
};

HRESULT CreatePane();
void ShowPane();
void ClosePane();