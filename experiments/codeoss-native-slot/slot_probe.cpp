#define OIE_BRIDGE_BUILD
#include "../native-surface-bridge/bridge.h"
#include <dwmapi.h>
#include <array>
#include <cstring>
#include <string>

namespace {
// Resolve the stable Node-API entry points from the lab Electron executable.
// This keeps the addon independent of Electron's private C++ ABI.
using Env = void*;
using Value = void*;
using Info = void*;
using Callback = Value (*)(Env, Info);
struct NodeApi {
    template<typename Fn> static Fn resolve(const char* name) {
        const auto address = GetProcAddress(GetModuleHandleW(nullptr), name);
        Fn function{};
        static_assert(sizeof(function) == sizeof(address));
        std::memcpy(&function, &address, sizeof(function));
        return function;
    }
    decltype(resolve<int(*)(Env, Info, size_t*, Value*, Value*, void**)>("")) args =
        resolve<int(*)(Env, Info, size_t*, Value*, Value*, void**)>("napi_get_cb_info");
    int (*string_value)(Env, Value, char*, size_t, size_t*) =
        resolve<decltype(string_value)>("napi_get_value_string_utf8");
    int (*make_string)(Env, const char*, size_t, Value*) = resolve<decltype(make_string)>("napi_create_string_utf8");
    int (*buffer)(Env, Value, void**, size_t*) = resolve<decltype(buffer)>("napi_get_buffer_info");
    int (*integer)(Env, Value, int32_t*) = resolve<decltype(integer)>("napi_get_value_int32");
    int (*function)(Env, const char*, size_t, Callback, void*, Value*) = resolve<decltype(function)>("napi_create_function");
    int (*property)(Env, Value, const char*, Value) = resolve<decltype(property)>("napi_set_named_property");
} api;

struct Fixture {
    HWND parent{}, slot{}, edit{}, palette{}, button{};
    uint64_t host{};
    OieSlotToken token{};
    bool selected{};
    bool painted{};
} fixture;
constexpr wchar_t PaletteClass[] = L"OieCodeOssLabPalette";
constexpr int Choose = 90;
constexpr COLORREF PaletteMarker = RGB(24, 121, 165);

LRESULT CALLBACK PaletteProc(HWND hwnd, UINT msg, WPARAM w, LPARAM l) {
    if (msg == WM_PAINT) {
        PAINTSTRUCT paint{};
        const auto dc = BeginPaint(hwnd, &paint);
        if (dc) {
            SetDCBrushColor(dc, PaletteMarker);
            fixture.painted = FillRect(dc, &paint.rcPaint, reinterpret_cast<HBRUSH>(GetStockObject(DC_BRUSH))) != 0;
        }
        EndPaint(hwnd, &paint);
        return 0;
    }
    if (msg == WM_COMMAND && LOWORD(w) == Choose && HIWORD(w) == BN_CLICKED && reinterpret_cast<HWND>(l) == fixture.button) {
        fixture.selected = true;
        ShowWindow(hwnd, SW_HIDE);
        SetFocus(fixture.edit);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, w, l);
}

const char* hit(POINT point) {
    if (!fixture.parent || !ClientToScreen(fixture.parent, &point)) return "none";
    const auto hwnd = WindowFromPoint(point);
    if (!hwnd) return "none";
    if (hwnd == fixture.button) return "palette-button";
    if (hwnd == fixture.palette) return "palette";
    if (hwnd == fixture.edit) return "native-document";
    if (hwnd == fixture.slot) return "native-slot";
    if (hwnd == fixture.parent || IsChild(fixture.parent, hwnd)) return "workbench-window";
    return "outside-lab-window";
}

void clear_overlay() {
    if (IsWindow(fixture.palette)) DestroyWindow(fixture.palette);
    UnregisterClassW(PaletteClass, GetModuleHandleW(nullptr));
    fixture.palette = nullptr;
    fixture.button = nullptr;
    fixture.selected = false;
    fixture.painted = false;
}

void reset() {
    clear_overlay();
    if (IsWindow(fixture.edit)) DestroyWindow(fixture.edit);
    if (fixture.host) {
        oie_slot_release(fixture.host, fixture.token);
        oie_host_close(fixture.host);
    }
    fixture = {};
}

const char* attach(HWND parent, RECT bounds) {
    if (fixture.host || !IsWindow(parent)) return "invalid-parent";
    if (oie_host_create(parent, OIE_NATIVE_CHILD_OVERLAYS, &fixture.host) != OIE_OK) return "host-create-failed";
    fixture.parent = parent;
    OieSlotSnapshot snapshot{};
    if (oie_slot_create(fixture.host, "synthetic-tab", &fixture.token) != OIE_OK ||
        oie_slot_layout(fixture.host, fixture.token, bounds, TRUE) != OIE_OK ||
        oie_slot_snapshot(fixture.host, fixture.token, &snapshot) != OIE_OK) {
        reset();
        return "slot-create-failed";
    }
    fixture.slot = snapshot.window;
    fixture.edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"Synthetic native document. No Office was opened.",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE, 0, 0, bounds.right - bounds.left, bounds.bottom - bounds.top,
        fixture.slot, nullptr, GetModuleHandleW(nullptr), nullptr);
    if (!fixture.edit) {
        reset();
        return "document-create-failed";
    }
    // The owning Electron host establishes the sibling order when it mounts its own slot.
    // A later host overlay can go above it. Layout updates continue to preserve Z order.
    if (!SetWindowPos(fixture.slot, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE)) {
        reset();
        return "slot-order-failed";
    }
    return "attached";
}

const char* overlay(RECT bounds, bool ownedPopup) {
    if (!fixture.edit || fixture.palette) return "invalid-state";
    WNDCLASSEXW cls{sizeof(cls)};
    cls.lpfnWndProc = PaletteProc;
    cls.hInstance = GetModuleHandleW(nullptr);
    cls.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    cls.lpszClassName = PaletteClass;
    if (!RegisterClassExW(&cls)) return "palette-class-failed";
    const auto width = bounds.right - bounds.left;
    const auto height = bounds.bottom - bounds.top;
    POINT origin{bounds.left, bounds.top};
    if (ownedPopup && !ClientToScreen(fixture.parent, &origin)) return "palette-position-failed";
    const DWORD style = (ownedPopup ? WS_POPUP : WS_CHILD) | WS_CLIPSIBLINGS | WS_CLIPCHILDREN;
    fixture.palette = CreateWindowExW(ownedPopup ? WS_EX_TOOLWINDOW : 0, PaletteClass, L"", style,
        origin.x, origin.y, width, height, fixture.parent, nullptr, cls.hInstance, nullptr);
    if (!fixture.palette) return "palette-create-failed";
    fixture.button = CreateWindowExW(0, L"BUTTON", L"Run synthetic palette action", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        0, 12, width, height - 12, fixture.palette, reinterpret_cast<HMENU>(static_cast<INT_PTR>(Choose)), cls.hInstance, nullptr);
    if (!fixture.button || (!ownedPopup && oie_host_raise_overlay(fixture.host, fixture.palette) != OIE_OK)) return "palette-button-failed";
    ShowWindow(fixture.palette, SW_SHOWNA);
    RedrawWindow(fixture.palette, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
    DwmFlush();
    SetFocus(fixture.button);
    return "overlay-shown";
}

Value operation(Env env, Info info) {
    std::array<Value, 6> argv{};
    size_t argc = argv.size();
    std::array<char, 24> op{};
    size_t length{};
    const char* result = "invalid-arguments";
    if (api.args(env, info, &argc, argv.data(), nullptr, nullptr) == 0 && argc == argv.size() &&
        api.string_value(env, argv[0], op.data(), op.size(), &length) == 0) {
        void* bytes{};
        size_t size{};
        std::array<int32_t, 4> data{};
        bool valid = true;
        for (size_t i = 0; i < data.size(); ++i) valid = valid && api.integer(env, argv[i + 2], &data[i]) == 0;
        if (valid && api.buffer(env, argv[1], &bytes, &size) == 0 && size == sizeof(HWND)) {
            HWND parent{};
            std::memcpy(&parent, bytes, size);
            const std::string command(op.data(), length);
            const RECT rect{data[0], data[1], data[2], data[3]};
            const POINT point{data[0], data[1]};
            if (command == "attach") result = attach(parent, rect);
            else if (parent != fixture.parent || !IsWindow(parent)) result = "unexpected-parent";
            else if (command == "overlay") result = overlay(rect, false);
            else if (command == "popup-control") result = overlay(rect, true);
            else if (command == "clear-overlay") { clear_overlay(); result = "overlay-cleared"; }
            else if (command == "sample") result = hit(point);
            else if (command == "paint-state") result = fixture.painted ? "paint-completed" : "paint-not-observed";
            else if (command == "paint" && std::strcmp(hit(point), "palette") == 0) {
                // One known color in our synthetic palette. Keep no bitmap or pixel value.
                auto screenPoint = point;
                ClientToScreen(parent, &screenPoint);
                const auto dc = GetDC(nullptr);
                if (dc) {
                    const auto pixel = GetPixel(dc, screenPoint.x, screenPoint.y);
                    result = pixel == CLR_INVALID ? "desktop-read-unavailable"
                        : pixel == PaletteMarker ? "marker-painted" : "marker-not-visible";
                    ReleaseDC(nullptr, dc);
                }
            }
            else if (command == "activate" && std::strcmp(hit(point), "palette-button") == 0) {
                // Target only our own button. BM_CLICK drives its real action without moving the system pointer.
                SendMessageW(fixture.button, BM_CLICK, 0, 0);
                result = fixture.selected && GetFocus() == fixture.edit && IsWindowVisible(fixture.edit)
                    ? "action-fired-focus-restored" : "action-or-focus-failed";
            } else if (command == "focus") result = GetFocus() == fixture.edit ? "document-focused" : "other-focus";
            else if (command == "detach") { reset(); result = "detached"; }
        }
    }
    Value value{};
    api.make_string(env, result, std::strlen(result), &value);
    return value;
}
}

extern "C" __declspec(dllexport) Value napi_register_module_v1(Env env, Value exports) {
    if (!api.args || !api.string_value || !api.make_string || !api.buffer || !api.integer || !api.function || !api.property) return exports;
    Value fn{};
    if (api.function(env, "operation", 9, operation, nullptr, &fn) == 0) api.property(env, exports, "operation", fn);
    return exports;
}
