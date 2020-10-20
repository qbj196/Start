#pragma once

#include <Windows.h>
#include <windowsx.h>
#include <Uxtheme.h>

#pragma comment(lib, "uxtheme.lib")


HRESULT CreatePane();
void ShowPane();
void ClosePane();