#define OIE_BRIDGE_BUILD
#include "bridge.h"
#include <atomic>
#include <memory>
#include <mutex>
#include <string>
#include <unordered_map>

namespace {
constexpr wchar_t ClassName[] = L"OfficeInEditor.ExperimentalNativeSlot.1";
constexpr wchar_t HostProperty[] = L"OfficeInEditor.ExperimentalNativeHostLease.1";
constexpr size_t MaxHosts = 8, MaxSlots = 32;
std::atomic<uint64_t> sequence{1};
std::mutex registration_mutex;
size_t host_count{};
struct Slot {
    OieSlotToken token{};
    HWND window{};
    std::string tab;
    uint64_t revision{1};
};
struct Host {
    HWND parent{};
    DWORD thread{};
    uint64_t identity{};
    std::unordered_map<uint64_t, std::unique_ptr<Slot>> slots;
};
// No worker, hooks, polling, disk state, or cross-thread mutation.
thread_local std::unordered_map<uint64_t, Host> hosts;

LRESULT CALLBACK SlotProc(HWND window, UINT message, WPARAM w, LPARAM l) {
    auto slot = reinterpret_cast<Slot*>(GetWindowLongPtrW(window, GWLP_USERDATA));
    if (message == WM_NCCREATE) {
        slot = static_cast<Slot*>(reinterpret_cast<CREATESTRUCTW*>(l)->lpCreateParams);
        SetWindowLongPtrW(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(slot));
        slot->window = window;
    }
    if (slot && (message == WM_SIZE || message == WM_SHOWWINDOW || message == WM_DPICHANGED))
        ++slot->revision;
    if (message == WM_NCDESTROY && slot) {
        slot->window = nullptr;
        ++slot->revision;
        SetWindowLongPtrW(window, GWLP_USERDATA, 0);
    }
    return DefWindowProcW(window, message, w, l);
}

Host* find_host(uint64_t id) {
    auto found = hosts.find(id);
    return found == hosts.end() ? nullptr : &found->second;
}
Slot* find_slot(Host* host, OieSlotToken token) {
    if (!host) return nullptr;
    auto found = host->slots.find(token.id);
    return found != host->slots.end() && found->second->token.generation == token.generation
        ? found->second.get() : nullptr;
}
bool valid_parent(Host* host) {
    DWORD pid{};
    return host && IsWindow(host->parent)
        && GetWindowThreadProcessId(host->parent, &pid) == GetCurrentThreadId()
        && pid == GetCurrentProcessId()
        && reinterpret_cast<uintptr_t>(GetPropW(host->parent, HostProperty)) == host->identity;
}
bool same_rect(RECT a, RECT b) {
    return a.left == b.left && a.top == b.top && a.right == b.right && a.bottom == b.bottom;
}
bool slot_rect(HWND window, HWND parent, RECT* rect) {
    if (!GetWindowRect(window, rect)) return false;
    SetLastError(0);
    auto result = MapWindowPoints(nullptr, parent, reinterpret_cast<POINT*>(rect), 2);
    return result || GetLastError() == 0;
}
}

unsigned oie_bridge_abi_version() { return 1; }

OieResult oie_host_create(HWND parent, unsigned mode, uint64_t* out) {
    if (!out || !IsWindow(parent)) return OIE_INVALID;
    *out = 0;
    if (mode != OIE_NATIVE_CHILD_OVERLAYS) return OIE_UNSUPPORTED;
    DWORD pid{};
    auto thread = GetWindowThreadProcessId(parent, &pid);
    if (pid != GetCurrentProcessId() || thread != GetCurrentThreadId()) return OIE_WRONG_THREAD;
    if (GetPropW(parent, HostProperty)) return OIE_BUSY;
    if (hosts.size() >= MaxHosts) return OIE_LIMIT;
    std::lock_guard lock(registration_mutex);
    WNDCLASSEXW cls{sizeof(cls)};
    cls.lpfnWndProc = SlotProc;
    cls.hInstance = GetModuleHandleW(nullptr);
    cls.lpszClassName = ClassName;
    cls.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    if (!RegisterClassExW(&cls) && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) return OIE_WINDOW_ERROR;
    WNDCLASSEXW registered{sizeof(registered)};
    if (!GetClassInfoExW(cls.hInstance, ClassName, &registered) || registered.lpfnWndProc != SlotProc)
        return OIE_WINDOW_ERROR;
    auto id = sequence.fetch_add(1);
    if (!SetPropW(parent, HostProperty, reinterpret_cast<HANDLE>(static_cast<uintptr_t>(id)))) {
        if (host_count == 0) UnregisterClassW(ClassName, GetModuleHandleW(nullptr));
        return OIE_WINDOW_ERROR;
    }
    hosts.emplace(id, Host{parent, thread, id, {}});
    ++host_count;
    *out = id;
    return OIE_OK;
}

OieResult oie_host_close(uint64_t id) {
    auto host = find_host(id);
    if (!host) return OIE_INVALID;
    if (!host->slots.empty()) return OIE_BUSY;
    if (reinterpret_cast<uintptr_t>(GetPropW(host->parent, HostProperty)) == id)
        RemovePropW(host->parent, HostProperty);
    hosts.erase(id);
    std::lock_guard lock(registration_mutex);
    if (--host_count == 0) UnregisterClassW(ClassName, GetModuleHandleW(nullptr));
    return OIE_OK;
}

OieResult oie_slot_create(uint64_t id, const char* tab, OieSlotToken* out) {
    auto host = find_host(id);
    if (!out || !tab || !valid_parent(host)) return OIE_INVALID;
    *out = {};
    auto length = strnlen_s(tab, 129);
    if (length == 0 || length > 128) return OIE_INVALID;
    for (size_t i = 0; i < length; ++i)
        if (!(tab[i] >= 'a' && tab[i] <= 'z') && !(tab[i] >= 'A' && tab[i] <= 'Z')
            && !(tab[i] >= '0' && tab[i] <= '9') && tab[i] != '-' && tab[i] != '_') return OIE_INVALID;
    for (auto& [key, slot] : host->slots) if (slot->tab == tab) return OIE_BUSY;
    if (host->slots.size() >= MaxSlots) return OIE_LIMIT;
    auto slot = std::make_unique<Slot>();
    slot->tab = tab;
    slot->token = {sequence.fetch_add(1), sequence.fetch_add(1)};
    auto window = CreateWindowExW(0, ClassName, L"", WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
        0, 0, 0, 0, host->parent, nullptr, GetModuleHandleW(nullptr), slot.get());
    if (!window) return OIE_WINDOW_ERROR;
    *out = slot->token;
    host->slots.emplace(slot->token.id, std::move(slot));
    return OIE_OK;
}

OieResult oie_slot_snapshot(uint64_t id, OieSlotToken token, OieSlotSnapshot* out) {
    if (!out) return OIE_INVALID;
    *out = {};
    auto host = find_host(id);
    auto slot = find_slot(host, token);
    if (!slot || !IsWindow(slot->window) || !valid_parent(host)) return OIE_STALE;
    DWORD pid{};
    out->host_tid = GetWindowThreadProcessId(host->parent, &pid);
    out->host_pid = pid;
    out->window = slot->window;
    out->dpi = GetDpiForWindow(slot->window);
    out->revision = slot->revision;
    out->visible = (GetWindowLongPtrW(slot->window, GWL_STYLE) & WS_VISIBLE) != 0;
    out->occupied = GetWindow(slot->window, GW_CHILD) != nullptr;
    if (!slot_rect(slot->window, host->parent, &out->bounds)) return OIE_WINDOW_ERROR;
    return OIE_OK;
}

OieResult oie_slot_layout(uint64_t id, OieSlotToken token, RECT bounds, BOOL visible) {
    auto host = find_host(id);
    auto slot = find_slot(host, token);
    if (!slot || !IsWindow(slot->window) || !valid_parent(host)) return OIE_STALE;
    RECT parent{};
    if (!GetClientRect(host->parent, &parent) || bounds.left < 0 || bounds.top < 0
        || bounds.right <= bounds.left || bounds.bottom <= bounds.top
        || bounds.right > parent.right || bounds.bottom > parent.bottom) return OIE_INVALID;
    RECT current{};
    if (!slot_rect(slot->window, host->parent, &current)) return OIE_WINDOW_ERROR;
    bool shown = (GetWindowLongPtrW(slot->window, GWL_STYLE) & WS_VISIBLE) != 0;
    if (same_rect(current, bounds) && shown == !!visible) return OIE_OK;
    return SetWindowPos(slot->window, nullptr, bounds.left, bounds.top,
        bounds.right - bounds.left, bounds.bottom - bounds.top,
        SWP_NOZORDER | SWP_NOACTIVATE | (visible ? SWP_SHOWWINDOW : SWP_HIDEWINDOW))
        ? OIE_OK : OIE_WINDOW_ERROR;
}

OieResult oie_slot_release(uint64_t id, OieSlotToken token) {
    auto host = find_host(id);
    auto slot = find_slot(host, token);
    if (!slot) return OIE_STALE;
    if (IsWindow(slot->window)) {
        if (GetWindow(slot->window, GW_CHILD)) return OIE_BUSY;
        if (!DestroyWindow(slot->window)) return OIE_WINDOW_ERROR;
    }
    host->slots.erase(token.id);
    return OIE_OK;
}

OieResult oie_host_raise_overlay(uint64_t id, HWND overlay) {
    auto host = find_host(id);
    DWORD pid{};
    if (!valid_parent(host) || !IsWindow(overlay) || GetParent(overlay) != host->parent
        || GetWindowThreadProcessId(overlay, &pid) != GetCurrentThreadId()
        || pid != GetCurrentProcessId()) return OIE_INVALID;
    for (auto& [key, slot] : host->slots) if (overlay == slot->window) return OIE_INVALID;
    return SetWindowPos(overlay, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE)
        ? OIE_OK : OIE_WINDOW_ERROR;
}
