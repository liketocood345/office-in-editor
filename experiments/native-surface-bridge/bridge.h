#pragma once
#include <windows.h>
#include <stdint.h>

#ifdef OIE_BRIDGE_BUILD
#define OIE_API __declspec(dllexport)
#else
#define OIE_API __declspec(dllimport)
#endif

// Experimental host-side ABI. This is not a VS Code API.
// Every call belongs on the native parent's owning UI thread.
extern "C" {
enum OieResult {
    OIE_OK = 0, OIE_INVALID = 1, OIE_WRONG_THREAD = 2, OIE_STALE = 3,
    OIE_BUSY = 4, OIE_LIMIT = 5, OIE_WINDOW_ERROR = 6, OIE_UNSUPPORTED = 7
};
enum OieOverlayMode { OIE_NATIVE_CHILD_OVERLAYS = 1 };
struct OieSlotToken { uint64_t id; uint64_t generation; };
struct OieSlotSnapshot {
    HWND window;
    DWORD host_pid;
    DWORD host_tid;
    UINT dpi;
    RECT bounds;
    uint64_t revision;
    BOOL visible;
    BOOL occupied;
};

OIE_API unsigned oie_bridge_abi_version();
OIE_API OieResult oie_host_create(HWND parent, unsigned overlay_mode, uint64_t* host);
OIE_API OieResult oie_host_close(uint64_t host);
OIE_API OieResult oie_slot_create(uint64_t host, const char* opaque_tab_id, OieSlotToken* token);
OIE_API OieResult oie_slot_layout(uint64_t host, OieSlotToken token, RECT bounds, BOOL visible);
OIE_API OieResult oie_slot_snapshot(uint64_t host, OieSlotToken token, OieSlotSnapshot* snapshot);
// Refuses to destroy a slot while a document container is still attached.
OIE_API OieResult oie_slot_release(uint64_t host, OieSlotToken token);
// Only a host-owned native overlay is accepted. HTML DOM nodes cannot be passed here.
OIE_API OieResult oie_host_raise_overlay(uint64_t host, HWND native_overlay);
}
