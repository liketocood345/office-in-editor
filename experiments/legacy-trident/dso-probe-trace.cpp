#include "dso-probe-trace.h"
namespace { LONG stages[48] = {}, results[48] = {}, used = 0; HWND documentWindow = NULL; }
extern "C" void __stdcall OieRememberWindow(HWND window) { documentWindow = window; }
extern "C" HWND __stdcall OieProbeDocumentWindow() { return documentWindow; }
extern "C" void __stdcall OieTrace(LONG stage, HRESULT hr) {
    if (used < 48) { stages[used] = stage; results[used] = hr; ++used; }
}
extern "C" LONG __stdcall OieProbeTrace(LONG index, LONG* stage, LONG* hr) {
    if (!stage || !hr || index < 0 || index >= used) return 0;
    *stage = stages[index]; *hr = results[index]; return 1;
}
