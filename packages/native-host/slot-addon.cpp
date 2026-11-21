#define OIE_BRIDGE_BUILD
#include "bridge.h"
#include <array>
#include <cstring>
#include <string>
#include <unordered_map>

namespace {
using Env = void*; using Value = void*; using Info = void*;
using Callback = Value (*)(Env, Info);
template<typename T> T resolve(const char* name) {
    auto address = GetProcAddress(GetModuleHandleW(nullptr), name);
    T fn{}; static_assert(sizeof(fn) == sizeof(address)); std::memcpy(&fn, &address, sizeof(fn)); return fn;
}
struct Api {
    int (*args)(Env, Info, size_t*, Value*, Value*, void**) = resolve<decltype(args)>("napi_get_cb_info");
    int (*string)(Env, Value, char*, size_t, size_t*) = resolve<decltype(string)>("napi_get_value_string_utf8");
    int (*buffer)(Env, Value, void**, size_t*) = resolve<decltype(buffer)>("napi_get_buffer_info");
    int (*integer)(Env, Value, int32_t*) = resolve<decltype(integer)>("napi_get_value_int32");
    int (*make_string)(Env, const char*, size_t, Value*) = resolve<decltype(make_string)>("napi_create_string_utf8");
    int (*function)(Env, const char*, size_t, Callback, void*, Value*) = resolve<decltype(function)>("napi_create_function");
    int (*property)(Env, Value, const char*, Value) = resolve<decltype(property)>("napi_set_named_property");
} api;
struct Slot { HWND parent{}, hwnd{}; uint64_t host{}; OieSlotToken token{}; std::wstring property; };
std::unordered_map<HWND, uint64_t> hosts;
std::unordered_map<std::string, Slot> slots;
void release_empty_host(HWND parent, uint64_t host) {
    for (const auto& entry : slots) if (entry.second.parent == parent) return;
    if (oie_host_close(host) == OIE_OK) hosts.erase(parent);
}
std::string error(unsigned code) { return "{\"error\":\"native_slot_" + std::to_string(code) + "\"}"; }
std::string snapshot(const Slot& slot) {
    return "{\"hwnd\":\"" + std::to_string(reinterpret_cast<uintptr_t>(slot.hwnd)) +
        "\",\"generation\":\"" + std::to_string(slot.token.generation) +
        "\",\"pid\":" + std::to_string(GetCurrentProcessId()) + "}";
}
unsigned office_application(DWORD pid) {
    HANDLE process = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
    if (!process) return 0;
    wchar_t filename[1024]{};
    DWORD length = static_cast<DWORD>(std::size(filename));
    const BOOL read = QueryFullProcessImageNameW(process, 0, filename, &length);
    CloseHandle(process);
    if (!read) return 0;
    const wchar_t* leaf = wcsrchr(filename, L'\\');
    leaf = leaf ? leaf + 1 : filename;
    if (_wcsicmp(leaf, L"WINWORD.EXE") == 0) return 1;
    if (_wcsicmp(leaf, L"EXCEL.EXE") == 0) return 2;
    if (_wcsicmp(leaf, L"POWERPNT.EXE") == 0) return 4;
    return 0;
}
Value operate(Env env, Info info) {
    std::array<Value, 8> argv{}; size_t argc = argv.size();
    std::array<char, 32> command{}; std::array<char, 80> token{}; size_t used{};
    std::string result = error(OIE_INVALID);
    void* bytes{}; size_t count{}; std::array<int32_t, 5> data{};
    bool valid = api.args(env, info, &argc, argv.data(), nullptr, nullptr) == 0 && argc == argv.size() &&
        api.string(env, argv[0], command.data(), command.size(), &used) == 0 &&
        api.string(env, argv[2], token.data(), token.size(), &used) == 0 && used == 44 &&
        api.buffer(env, argv[1], &bytes, &count) == 0 && count == sizeof(HWND);
    for (size_t i = 0; i < data.size(); ++i) valid = valid && api.integer(env, argv[i + 3], &data[i]) == 0;
    if (valid) {
        HWND parent{}; std::memcpy(&parent, bytes, count);
        const std::string id(token.data());
        const std::string op(command.data());
        if (id.rfind("oie-surface-", 0) != 0 || id.find_first_not_of("0123456789abcdef", 12) != std::string::npos) valid = false;
        if (valid && op == "create" && slots.find(id) == slots.end()) {
            uint64_t host{};
            auto registered = hosts.find(parent);
            unsigned code = OIE_OK;
            if (registered == hosts.end()) {
                code = oie_host_create(parent, OIE_NATIVE_CHILD_OVERLAYS, &host);
                if (code == OIE_OK) hosts.emplace(parent, host);
            } else host = registered->second;
            Slot slot{}; slot.parent = parent; slot.host = host;
            if (code == OIE_OK) code = oie_slot_create(host, id.c_str(), &slot.token);
            OieSlotSnapshot state{};
            if (code == OIE_OK) code = oie_slot_snapshot(host, slot.token, &state);
            slot.hwnd = state.window;
            slot.property = L"OfficeInEditor.NativeSurface." + std::wstring(id.begin(), id.end());
            if (code == OIE_OK && !SetPropW(slot.hwnd, slot.property.c_str(), reinterpret_cast<HANDLE>(static_cast<uintptr_t>(slot.token.generation)))) code = OIE_WINDOW_ERROR;
            if (code == OIE_OK) {
                code = oie_slot_layout(host, slot.token, {data[0], data[1], data[0] + data[2], data[1] + data[3]}, data[4]);
                if (code == OIE_OK) {
                    SetWindowPos(slot.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);
                    slots.emplace(id, slot); result = snapshot(slot);
                }
            }
            if (code != OIE_OK) {
                if (slot.token.id) oie_slot_release(host, slot.token);
                if (host) release_empty_host(parent, host);
                result = error(code);
            }
        } else if (valid) {
            auto found = slots.find(id);
            if (found != slots.end() && found->second.parent == parent) {
                auto& slot = found->second;
                if (op == "layout") {
                    auto code = oie_slot_layout(slot.host, slot.token, {data[0], data[1], data[0] + data[2], data[1] + data[3]}, data[4]);
                    result = code == OIE_OK ? snapshot(slot) : error(code);
                } else if (op == "inspect") {
                    struct Check {
                        bool word{};
                        bool visible{};
                        unsigned applications{};
                        std::unordered_map<DWORD, unsigned> known;
                    } state;
                    EnumChildWindows(slot.hwnd, [](HWND hwnd, LPARAM raw) -> BOOL {
                        auto check = reinterpret_cast<Check*>(raw);
                        wchar_t cls[160]{}; GetClassNameW(hwnd, cls, 160);
                        const bool visible = IsWindowVisible(hwnd) != FALSE;
                        if (std::wstring(cls).rfind(L"_Ww", 0) == 0) {
                            check->word = true;
                            check->visible = check->visible || visible;
                        }
                        RECT bounds{};
                        if (visible && GetClientRect(hwnd, &bounds) && bounds.right >= 32 && bounds.bottom >= 32) {
                            DWORD pid{};
                            GetWindowThreadProcessId(hwnd, &pid);
                            if (pid) {
                                auto known = check->known.find(pid);
                                if (known == check->known.end() && check->known.size() < 16)
                                    known = check->known.emplace(pid, office_application(pid)).first;
                                if (known != check->known.end()) check->applications |= known->second;
                            }
                        }
                        return TRUE;
                    }, reinterpret_cast<LPARAM>(&state));
                    unsigned samples{}, inSlot{}, officeHits{};
                    RECT client{};
                    if (GetClientRect(slot.hwnd, &client) && client.right >= 32 && client.bottom >= 32) {
                        const POINT weights[] = {{2, 2}, {1, 1}, {3, 1}, {1, 3}, {3, 3}};
                        for (const auto weight : weights) {
                            POINT point{client.right * weight.x / 4, client.bottom * weight.y / 4};
                            if (!ClientToScreen(slot.hwnd, &point)) continue;
                            samples++;
                            const HWND hit = WindowFromPoint(point);
                            if (hit != slot.hwnd && !IsChild(slot.hwnd, hit)) continue;
                            inSlot++;
                            DWORD pid{};
                            GetWindowThreadProcessId(hit, &pid);
                            const auto application = state.known.find(pid);
                            if (application != state.known.end() && application->second) officeHits++;
                        }
                    }
                    const bool foreground = GetForegroundWindow() == GetAncestor(slot.hwnd, GA_ROOT);
                    const auto style = GetWindowLongPtrW(slot.hwnd, GWL_STYLE);
                    const bool slotIsVerifiedChild = GetParent(slot.hwnd) == slot.parent &&
                        (style & WS_CHILD) != 0 && (style & WS_POPUP) == 0 && GetWindow(slot.hwnd, GW_OWNER) == nullptr;
                    result = "{\"wordChild\":" + std::string(state.word ? "true" : "false") +
                        ",\"visible\":" + (state.visible ? "true" : "false") +
                        ",\"slotIsVerifiedChild\":" + (slotIsVerifiedChild ? "true" : "false") +
                        ",\"officeApplications\":" + std::to_string(state.applications) +
                        ",\"samples\":" + std::to_string(samples) +
                        ",\"inSlot\":" + std::to_string(inSlot) +
                        ",\"officeHits\":" + std::to_string(officeHits) +
                        ",\"foreground\":" + (foreground ? "true" : "false") + "}";
                } else if (op == "hide") {
                    if (IsWindow(slot.hwnd)) {
                        ShowWindow(slot.hwnd, SW_HIDE);
                        result = snapshot(slot);
                    }
                } else if (op == "release") {
                    const auto host = slot.host;
                    const auto code = oie_slot_release(host, slot.token);
                    result = code == OIE_OK ? "{\"released\":true}" : error(code);
                    if (code == OIE_OK) {
                        slots.erase(found);
                        release_empty_host(parent, host);
                    }
                }
            }
        }
    }
    Value value{}; api.make_string(env, result.c_str(), result.size(), &value); return value;
}
}
extern "C" __declspec(dllexport) Value napi_register_module_v1(Env env, Value exports) {
    if (!api.args || !api.string || !api.buffer || !api.integer || !api.make_string || !api.function || !api.property) return exports;
    Value function{}; if (api.function(env, "operate", 7, operate, nullptr, &function) == 0) api.property(env, exports, "operate", function);
    return exports;
}
