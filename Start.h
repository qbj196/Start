#pragma once

#include <Windows.h>
#include <ShlObj.h>
#include <Uxtheme.h>

#pragma comment(lib, "uxtheme.lib")


HRESULT CreateStart(HWND hWnd);
void ShowStart(BOOL fShow, DWORD dwBandID);
void UpdateStart();
void CloseStart();