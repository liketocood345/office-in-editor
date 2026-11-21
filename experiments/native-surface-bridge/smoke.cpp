#include "bridge.h"
#include <iostream>
#include <stdexcept>
#include <string>
#include <string>
#include <thread>
#include <vector>

static int checks = 0;
void check(bool value, const char* name) {
    if (!value) throw std::runtime_error(name);
    ++checks;
    std::cout << "PASS " << name << '\n';
}
int main() {
    SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    HWND root = CreateWindowExW(0, L"STATIC", L"Synthetic native host", WS_POPUP | WS_CLIPCHILDREN,
        40, 40, 800, 600, nullptr, nullptr, GetModuleHandleW(nullptr), nullptr);
    if (!root) return 2;
    ShowWindow(root, SW_SHOWNOACTIVATE);
    try {
        uint64_t host{};
        check(oie_bridge_abi_version() == 1, "ABI is explicit");
        check(oie_host_create(root, 0, &host) == OIE_UNSUPPORTED,
            "unspecified overlay mechanism is rejected");
        OieResult other_thread{};
        std::thread other([&] { uint64_t unused{}; other_thread = oie_host_create(root, 1, &unused); });
        other.join();
        check(other_thread == OIE_WRONG_THREAD, "another UI thread cannot impersonate the host");
        check(oie_host_create(root, OIE_NATIVE_CHILD_OVERLAYS, &host) == OIE_OK, "native host creates a bridge");
        uint64_t duplicate_host{};
        check(oie_host_create(root, OIE_NATIVE_CHILD_OVERLAYS, &duplicate_host) == OIE_BUSY,
            "one native parent cannot be registered by two hosts");
        OieSlotToken first{}, second{}, duplicate{};
        check(oie_slot_create(host, "tab_a", &first) == OIE_OK
            && oie_slot_create(host, "tab_b", &second) == OIE_OK, "two independent tab slots are created");
        check(oie_slot_create(host, "tab_a", &duplicate) == OIE_BUSY, "duplicate tab ownership is refused");
        check(oie_slot_layout(host, first, {10, 50, 390, 450}, TRUE) == OIE_OK
            && oie_slot_layout(host, second, {410, 50, 790, 450}, TRUE) == OIE_OK, "slots accept parent client coordinates");
        OieSlotSnapshot a{}, b{};
        oie_slot_snapshot(host, first, &a); oie_slot_snapshot(host, second, &b);
        check(GetParent(a.window) == root && GetParent(b.window) == root
            && (GetWindowLongPtrW(a.window, GWL_STYLE) & WS_CHILD), "slots are actual Win32 children");
        auto changed = first; ++changed.generation;
        check(oie_slot_layout(host, changed, {0, 0, 100, 100}, TRUE) == OIE_STALE,
            "stale generation cannot move a live slot");
        check(oie_slot_layout(host, first, {-1, 0, 100, 100}, TRUE) == OIE_INVALID,
            "layout outside the parent is refused");
        uint64_t revision = a.revision;
        for (int i = 0; i < 1000; ++i)
            if (oie_slot_layout(host, first, {10, 50, 390, 450}, TRUE) != OIE_OK) throw std::runtime_error("repeat");
        oie_slot_snapshot(host, first, &a);
        check(a.revision == revision, "unchanged layout does not generate native size/show work");

        // Native HWND ordering only: this does not assert Chromium DOM composition.
        HWND overlay = CreateWindowExW(0, L"STATIC", L"Synthetic menu", WS_CHILD | WS_VISIBLE,
            30, 60, 450, 100, root, nullptr, GetModuleHandleW(nullptr), nullptr);
        check(oie_host_raise_overlay(host, overlay) == OIE_OK
            && ChildWindowFromPointEx(root, {40, 80}, CWP_SKIPINVISIBLE) == overlay,
            "host-owned native overlay orders above both slots");
        check(oie_slot_layout(host, first, {10, 50, 390, 450}, FALSE) == OIE_OK
            && oie_slot_layout(host, first, {10, 50, 390, 450}, TRUE) == OIE_OK
            && ChildWindowFromPointEx(root, {40, 80}, CWP_SKIPINVISIBLE) == overlay,
            "tab visibility updates do not raise content above a host overlay");
        check(oie_host_raise_overlay(host, a.window) == OIE_INVALID,
            "a content slot cannot be relabelled as an overlay");

        RECT before{}, after{};
        GetWindowRect(a.window, &before);
        SetWindowPos(root, nullptr, 140, 100, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
        GetWindowRect(a.window, &after);
        check(after.left - before.left == 100 && after.top - before.top == 60,
            "native parenting follows host movement without a positioning loop");

        HWND content = CreateWindowExW(0, L"STATIC", L"Synthetic document", WS_CHILD | WS_VISIBLE,
            0, 0, 100, 100, a.window, nullptr, GetModuleHandleW(nullptr), nullptr);
        check(oie_slot_release(host, first) == OIE_BUSY && IsWindow(content),
            "release refuses to destroy a slot containing a document");
        check(oie_host_close(host) == OIE_BUSY, "host close requires explicit slot release");
        DestroyWindow(content);
        check(oie_slot_release(host, first) == OIE_OK && IsWindow(b.window),
            "releasing one tab leaves another intact");
        check(oie_slot_snapshot(host, first, &a) == OIE_STALE, "released token stays invalid");
        std::vector<OieSlotToken> extra;
        for (int i = 0; i < 31; ++i) {
            OieSlotToken token{};
            if (oie_slot_create(host, ("bounded_" + std::to_string(i)).c_str(), &token) != OIE_OK)
                throw std::runtime_error("bounded slot creation");
            extra.push_back(token);
        }
        check(oie_slot_create(host, "overflow", &duplicate) == OIE_LIMIT, "slot count is bounded");
        for (auto token : extra) oie_slot_release(host, token);
        oie_slot_release(host, second);
        DestroyWindow(overlay);
        check(oie_host_close(host) == OIE_OK, "explicit cleanup succeeds after content detach");
        check(GetPropW(root, L"OfficeInEditor.ExperimentalNativeHostLease.1") == nullptr,
            "closing the host removes its native property");
        DestroyWindow(root);
        std::cout << "Completed " << checks << " checks. No Office or IDE was started.\n";
        return 0;
    } catch (const std::exception& e) {
        std::cerr << "FAIL " << e.what() << '\n';
        DestroyWindow(root);
        return 1;
    }
}
