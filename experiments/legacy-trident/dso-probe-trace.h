// Private sample diagnostics: 48 in-memory phase/HRESULT pairs, no file I/O.
#pragma once
#include <windows.h>
extern "C" void __stdcall OieTrace(LONG stage, HRESULT hr);
extern "C" void __stdcall OieRememberWindow(HWND window);
