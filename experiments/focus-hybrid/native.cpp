#define NOMINMAX
#include <windows.h>
#include <d3d11.h>
#include <dwmapi.h>
#include <roapi.h>
#include <wincodec.h>
#include <wincrypt.h>
#include <UIAutomation.h>
#include <windows.graphics.capture.interop.h>
#include <windows.graphics.directx.direct3d11.interop.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Graphics.Capture.h>
#include <winrt/Windows.Graphics.DirectX.Direct3D11.h>
#include <algorithm>
#include <atomic>
#include <chrono>
#include <iostream>
#include <mutex>
#include <string>
#include <thread>
#include <vector>

using namespace winrt;
using namespace winrt::Windows::Graphics::Capture;
using namespace winrt::Windows::Graphics::DirectX;
using namespace winrt::Windows::Graphics::DirectX::Direct3D11;

namespace {
constexpr UINT ForegroundChanged = WM_APP + 1;
constexpr UINT EnterOffice = WM_APP + 2;
constexpr UINT EndProbe = WM_APP + 3;
constexpr UINT ReturnToEditor = WM_APP + 4;
constexpr UINT LayoutChanged = WM_APP + 5;
constexpr UINT CapturePolicyChanged = WM_APP + 6;
constexpr UINT SmokeClickPreview = WM_APP + 7;
constexpr UINT SmokeClickEditor = WM_APP + 8;
constexpr UINT SmokeMovePointer = WM_APP + 9;
constexpr UINT SmokeRaiseParked = WM_APP + 10;
constexpr UINT AdoptPolicyChanged = WM_APP + 11;
enum class CapturePolicy : WPARAM { Switch, Always, Mouse };
std::mutex outputLock;
std::atomic_bool stopped = false;
UINT sequence = 0;
DWORD controlThread;
const auto nativeClockStart = std::chrono::steady_clock::now();

struct MtaLifetime {
    HRESULT status = RoInitialize(RO_INIT_MULTITHREADED);
    ~MtaLifetime() { if (SUCCEEDED(status)) RoUninitialize(); }
};

void emit(const std::string& line) {
    std::lock_guard<std::mutex> lock(outputLock);
    const auto us = std::chrono::duration_cast<std::chrono::microseconds>(
        std::chrono::steady_clock::now() - nativeClockStart).count();
    std::cout.write(line.data(), static_cast<std::streamsize>(line.size() - 1));
    std::cout << ",\"nativeUs\":" << us << "}\n" << std::flush;
}

void stage(const char* name) { emit(std::string("{\"stage\":\"") + name + "\"}"); }

void watchCapturePolicy(HANDLE quit) {
    HKEY key{};
    if (RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\OfficeInEditor\\FocusHybrid", 0, nullptr, 0,
            KEY_QUERY_VALUE | KEY_NOTIFY, nullptr, &key, nullptr) != ERROR_SUCCESS) {
        stage("capture_policy_watch_unavailable");
        return;
    }
    handle changed{CreateEventW(nullptr, FALSE, FALSE, nullptr)};
    if (changed) {
        const HANDLE events[]{quit, changed.get()};
        do {
            if (RegNotifyChangeKeyValue(key, FALSE, REG_NOTIFY_CHANGE_LAST_SET, changed.get(), TRUE) != ERROR_SUCCESS) break;
            DWORD policy{}, type{}, bytes = sizeof(policy);
            if (RegQueryValueExW(key, L"PixelPolicy", nullptr, &type, reinterpret_cast<BYTE*>(&policy), &bytes) == ERROR_SUCCESS &&
                type == REG_DWORD && bytes == sizeof(policy) && policy <= static_cast<DWORD>(CapturePolicy::Mouse)) {
                PostThreadMessageW(controlThread, CapturePolicyChanged, policy, 0);
            }
        } while (WaitForMultipleObjects(2, events, FALSE, INFINITE) == WAIT_OBJECT_0 + 1);
    }
    RegCloseKey(key);
}

HRESULT requestAccessibleFocus(HWND window) {
    const HRESULT initialized = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(initialized)) return initialized;
    HRESULT result;
    {
        com_ptr<IUIAutomation> automation;
        result = CoCreateInstance(__uuidof(CUIAutomation), nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(automation.put()));
        com_ptr<IUIAutomationElement> target;
        if (SUCCEEDED(result)) result = automation->ElementFromHandle(window, target.put());
        if (SUCCEEDED(result) && target) result = target->SetFocus();
    }
    CoUninitialize();
    return result;
}

bool previewBounds(HWND editor, const std::wstring& surfaceName, RECT& bounds) {
    const HRESULT initialized = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(initialized)) return false;
    bool found = false;
    {
        com_ptr<IUIAutomation> automation;
        com_ptr<IUIAutomationElement> window, surface;
        com_ptr<IUIAutomationCondition> matches;
        VARIANT name{};
        name.vt = VT_BSTR;
        name.bstrVal = SysAllocString(surfaceName.c_str());
        RECT editorRect{};
        found = name.bstrVal &&
            SUCCEEDED(CoCreateInstance(__uuidof(CUIAutomation), nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(automation.put()))) &&
            SUCCEEDED(automation->ElementFromHandle(editor, window.put())) && window &&
            SUCCEEDED(automation->CreatePropertyCondition(UIA_NamePropertyId, name, matches.put())) &&
            SUCCEEDED(window->FindFirst(TreeScope_Descendants, matches.get(), surface.put())) && surface &&
            SUCCEEDED(surface->get_CurrentBoundingRectangle(&bounds)) && GetWindowRect(editor, &editorRect) &&
            bounds.right > bounds.left && bounds.bottom > bounds.top && !IsIconic(editor) &&
            bounds.left >= editorRect.left && bounds.top >= editorRect.top &&
            bounds.right <= editorRect.right && bounds.bottom <= editorRect.bottom;
        VariantClear(&name);
    }
    CoUninitialize();
    return found;
}

struct WindowBounds { RECT outer, visible; };

WindowBounds windowBounds(HWND window) {
    WindowBounds bounds{};
    GetWindowRect(window, &bounds.outer);
    if (FAILED(DwmGetWindowAttribute(window, DWMWA_EXTENDED_FRAME_BOUNDS, &bounds.visible, sizeof(RECT)))) {
        bounds.visible = bounds.outer;
    }
    return bounds;
}

bool fitVisibleWindow(HWND window, const RECT& target, const char* event) {
    if (IsZoomed(window)) ShowWindow(window, SW_SHOWNOACTIVATE);
    const auto before = windowBounds(window);
    const LONG left = before.visible.left - before.outer.left;
    const LONG top = before.visible.top - before.outer.top;
    const LONG right = before.outer.right - before.visible.right;
    const LONG bottom = before.outer.bottom - before.visible.bottom;
    if (!SetWindowPos(window, nullptr, target.left - left, target.top - top,
            target.right - target.left + left + right, target.bottom - target.top + top + bottom,
            SWP_NOACTIVATE | SWP_NOZORDER)) return false;
    DwmFlush();
    const RECT actual = windowBounds(window).visible;
    const LONG error = std::max({std::abs(actual.left - target.left), std::abs(actual.top - target.top),
        std::abs(actual.right - target.right), std::abs(actual.bottom - target.bottom)});
    emit(std::string("{\"stage\":\"") + event + "\",\"width\":" + std::to_string(target.right - target.left) +
        ",\"height\":" + std::to_string(target.bottom - target.top) + ",\"alignmentErrorPx\":" + std::to_string(error) + "}");
    return error <= 2;
}

DWORD windowPid(HWND window) {
    DWORD pid{};
    GetWindowThreadProcessId(window, &pid);
    return pid;
}

bool clickOwnedWindow(HWND window, POINT point) {
    HWND underPointer = WindowFromPoint(point);
    if (!underPointer || GetAncestor(underPointer, GA_ROOTOWNER) != window) {
        stage(underPointer && windowPid(underPointer) == windowPid(window)
            ? "smoke_click_hit_sibling_root" : "smoke_click_hit_other_window");
        return false;
    }
    const int screenX = GetSystemMetrics(SM_XVIRTUALSCREEN);
    const int screenY = GetSystemMetrics(SM_YVIRTUALSCREEN);
    const int width = GetSystemMetrics(SM_CXVIRTUALSCREEN);
    const int height = GetSystemMetrics(SM_CYVIRTUALSCREEN);
    if (width < 2 || height < 2 || point.x < screenX || point.y < screenY ||
        point.x >= screenX + width || point.y >= screenY + height) {
        stage("smoke_click_outside_desktop");
        return false;
    }
    INPUT input[3]{};
    for (auto& item : input) item.type = INPUT_MOUSE;
    input[0].mi.dx = MulDiv(point.x - screenX, 65535, width - 1);
    input[0].mi.dy = MulDiv(point.y - screenY, 65535, height - 1);
    input[0].mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK;
    input[1].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
    input[2].mi.dwFlags = MOUSEEVENTF_LEFTUP;
    if (SendInput(3, input, sizeof(INPUT)) == 3) return true;
    stage("smoke_click_sendinput_rejected");
    return false;
}

struct WindowSearch { DWORD pid; std::wstring prefix; const wchar_t* className; HWND result{}; };
BOOL CALLBACK findWindow(HWND window, LPARAM parameter) {
    auto& search = *reinterpret_cast<WindowSearch*>(parameter);
    if (windowPid(window) != search.pid || !IsWindowVisible(window) || GetWindow(window, GW_OWNER)) return TRUE;
    wchar_t title[260]{};
    GetWindowTextW(window, title, 260);
    if (!title[0] || std::wstring(title).rfind(search.prefix, 0) != 0) return TRUE;
    if (search.className) {
        wchar_t actualClass[64]{};
        GetClassNameW(window, actualClass, 64);
        if (wcscmp(search.className, actualClass) != 0) return TRUE;
    }
    search.result = window;
    return FALSE;
}

void CALLBACK onForeground(HWINEVENTHOOK, DWORD, HWND, LONG, LONG, DWORD, DWORD) {
    PostThreadMessageW(controlThread, ForegroundChanged, 0, 0);
}

std::string pngFrame(IWICImagingFactory* factory, const std::vector<BYTE>& pixels, UINT w, UINT h) {
    com_ptr<IStream> stream;
    check_hresult(CreateStreamOnHGlobal(nullptr, TRUE, stream.put()));
    com_ptr<IWICBitmapEncoder> encoder;
    check_hresult(factory->CreateEncoder(GUID_ContainerFormatPng, nullptr, encoder.put()));
    check_hresult(encoder->Initialize(stream.get(), WICBitmapEncoderNoCache));
    com_ptr<IWICBitmapFrameEncode> frame;
    check_hresult(encoder->CreateNewFrame(frame.put(), nullptr));
    check_hresult(frame->Initialize(nullptr));
    check_hresult(frame->SetSize(w, h));
    auto format = GUID_WICPixelFormat32bppBGRA;
    check_hresult(frame->SetPixelFormat(&format));
    if (format != GUID_WICPixelFormat32bppBGRA) throw hresult_error(E_FAIL);
    check_hresult(frame->WritePixels(h, w * 4, static_cast<UINT>(pixels.size()), const_cast<BYTE*>(pixels.data())));
    check_hresult(frame->Commit());
    check_hresult(encoder->Commit());
    STATSTG size{};
    check_hresult(stream->Stat(&size, STATFLAG_NONAME));
    if (size.cbSize.QuadPart > 700000) return {};
    HGLOBAL memory{};
    check_hresult(GetHGlobalFromStream(stream.get(), &memory));
    auto bytes = static_cast<const BYTE*>(GlobalLock(memory));
    if (!bytes) throw hresult_error(E_FAIL);
    DWORD n{};
    const auto flags = CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF;
    CryptBinaryToStringA(bytes, size.cbSize.LowPart, flags, nullptr, &n);
    std::string encoded(n, '\0');
    const BOOL success = CryptBinaryToStringA(bytes, size.cbSize.LowPart, flags, encoded.data(), &n);
    GlobalUnlock(memory);
    if (!success) throw hresult_error(HRESULT_FROM_WIN32(GetLastError()));
    encoded.resize(n);
    return encoded;
}

void captureWindow(HWND hwnd, HANDLE prepared, HANDLE offscreen, bool diagnostics) {
    const char* failureStage = "wgc_device_failed";
    try {
        MtaLifetime apartment;
        check_hresult(apartment.status);
        stage("wgc_preparing_device");
        if (!GraphicsCaptureSession::IsSupported()) { stage("wgc_unsupported"); SetEvent(prepared); return; }
        com_ptr<ID3D11Device> device;
        com_ptr<ID3D11DeviceContext> context;
        check_hresult(D3D11CreateDevice(nullptr, D3D_DRIVER_TYPE_HARDWARE, nullptr, D3D11_CREATE_DEVICE_BGRA_SUPPORT,
            nullptr, 0, D3D11_SDK_VERSION, device.put(), nullptr, context.put()));
        auto dxgi = device.as<IDXGIDevice>();
        com_ptr<IInspectable> inspectable;
        check_hresult(CreateDirect3D11DeviceFromDXGIDevice(dxgi.get(), inspectable.put()));
        auto direct = inspectable.as<IDirect3DDevice>();
        failureStage = "wgc_item_failed";
        stage("wgc_preparing_item");
        auto factory = get_activation_factory<GraphicsCaptureItem, IGraphicsCaptureItemInterop>();
        GraphicsCaptureItem item{ nullptr };
        check_hresult(factory->CreateForWindow(hwnd, guid_of<GraphicsCaptureItem>(), put_abi(item)));
        failureStage = "wgc_frame_pool_failed";
        auto size = item.Size();
        if (size.Width <= 0 || size.Height <= 0) {
            const auto bounds = windowBounds(hwnd).visible;
            size = {bounds.right - bounds.left, bounds.bottom - bounds.top};
        }
        auto pool = Direct3D11CaptureFramePool::CreateFreeThreaded(direct, DirectXPixelFormat::B8G8R8A8UIntNormalized, 2, size);
        failureStage = "wgc_session_failed";
        auto session = pool.CreateCaptureSession(item);
        try { session.IsCursorCaptureEnabled(false); } catch (const hresult_error&) { stage("wgc_cursor_setting_unavailable"); }
        handle arrival{ CreateEventW(nullptr, FALSE, FALSE, nullptr) };
        auto subscribed = pool.FrameArrived(auto_revoke, [&arrival](auto const&, auto const&) { SetEvent(arrival.get()); });
        com_ptr<IWICImagingFactory> imaging;
        check_hresult(CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(imaging.put())));
        stage("wgc_session_prepared");
        SetEvent(prepared);
        WaitForSingleObject(offscreen, INFINITE);
        if (stopped) return;
        failureStage = "wgc_start_failed";
        session.StartCapture();
        stage("wgc_started_while_parked");
        failureStage = "wgc_frame_failed";
        auto nextCopy = std::chrono::steady_clock::now();
        com_ptr<ID3D11Texture2D> staging;
        std::vector<BYTE> previous;
        UINT prevWidth = 0, prevHeight = 0;
        bool loggedParkedFrame = false;
        bool loggedParkedUniform = false;
        LONG previousCropWidth = 0, previousCropHeight = 0;
        auto statsSince = std::chrono::steady_clock::now();
        UINT arrived = 0, encodedCount = 0, unchanged = 0, throttled = 0, oversize = 0;
        size_t bytes = 0;
        auto readCpuMs = [] {
            FILETIME created{}, ended{}, kernel{}, user{};
            if (!GetProcessTimes(GetCurrentProcess(), &created, &ended, &kernel, &user)) return 0ULL;
            return (((static_cast<unsigned long long>(kernel.dwHighDateTime) << 32) | kernel.dwLowDateTime) +
                    ((static_cast<unsigned long long>(user.dwHighDateTime) << 32) | user.dwLowDateTime)) / 10000ULL;
        };
        auto lastCpuMs = diagnostics ? readCpuMs() : 0;
        while (!stopped && IsWindow(hwnd)) {
            auto waited = WaitForSingleObject(arrival.get(), 100);
            const auto observed = std::chrono::steady_clock::now();
            const auto intervalMs = std::chrono::duration_cast<std::chrono::milliseconds>(observed - statsSince).count();
            if (diagnostics && intervalMs >= 1000) {
                auto cpuMs = readCpuMs();
                emit("{\"stage\":\"capture_stats_1s\",\"intervalMs\":" + std::to_string(intervalMs) +
                    ",\"framesRead\":" + std::to_string(arrived) + ",\"framesSent\":" + std::to_string(encodedCount) +
                    ",\"unchanged\":" + std::to_string(unchanged) + ",\"throttled\":" + std::to_string(throttled) +
                    ",\"oversize\":" + std::to_string(oversize) + ",\"bytes\":" + std::to_string(bytes) +
                    ",\"helperCpuMs\":" + std::to_string(cpuMs >= lastCpuMs ? cpuMs - lastCpuMs : 0) + "}");
                lastCpuMs = cpuMs;
                statsSince = observed;
                arrived = encodedCount = unchanged = throttled = oversize = 0;
                bytes = 0;
            }
            if (waited != WAIT_OBJECT_0 || stopped) continue;
            auto frame = pool.TryGetNextFrame();
            if (!frame) continue;
            ++arrived;
            const auto current = frame.ContentSize();
            if (current.Width <= 0 || current.Height <= 0) continue;
            if (size.Width != current.Width || size.Height != current.Height) {
                frame.Close();
                size = current;
                staging = nullptr;
                pool.Recreate(direct, DirectXPixelFormat::B8G8R8A8UIntNormalized, 2, size);
                continue;
            }
            const auto now = std::chrono::steady_clock::now();
            if (now < nextCopy) { ++throttled; continue; }
            nextCopy = now + std::chrono::milliseconds(300);
            const auto bounds = windowBounds(hwnd);
            const LONG visibleWidth = bounds.visible.right - bounds.visible.left;
            const LONG visibleHeight = bounds.visible.bottom - bounds.visible.top;
            if (visibleWidth <= 0 || visibleHeight <= 0 || bounds.outer.right > GetSystemMetrics(SM_XVIRTUALSCREEN)) continue;
            // WGC can provide either the visible rectangle or the full Win32 rectangle with resize margins.
            const LONG padX = size.Width - visibleWidth;
            const LONG padY = size.Height - visibleHeight;
            if (padX < 0 || padY < 0 || padX > 64 || padY > 64) continue;
            const LONG cropLeft = size.Width == bounds.outer.right - bounds.outer.left
                ? std::clamp(bounds.visible.left - bounds.outer.left, 0L, padX) : padX / 2;
            const LONG cropTop = size.Height == bounds.outer.bottom - bounds.outer.top
                ? std::clamp(bounds.visible.top - bounds.outer.top, 0L, padY) : padY / 2;
            if (visibleWidth != previousCropWidth || visibleHeight != previousCropHeight) {
                emit("{\"stage\":\"preview_frame_bounds\",\"width\":" + std::to_string(visibleWidth) +
                    ",\"height\":" + std::to_string(visibleHeight) + ",\"cropWidthPx\":" + std::to_string(padX) +
                    ",\"cropHeightPx\":" + std::to_string(padY) + "}");
                previousCropWidth = visibleWidth;
                previousCropHeight = visibleHeight;
            }
            auto surface = frame.Surface().as<::Windows::Graphics::DirectX::Direct3D11::IDirect3DDxgiInterfaceAccess>();
            com_ptr<ID3D11Texture2D> texture;
            check_hresult(surface->GetInterface(IID_PPV_ARGS(texture.put())));
            if (!staging) {
                D3D11_TEXTURE2D_DESC description{};
                texture->GetDesc(&description);
                description.BindFlags = description.MiscFlags = 0;
                description.Usage = D3D11_USAGE_STAGING;
                description.CPUAccessFlags = D3D11_CPU_ACCESS_READ;
                check_hresult(device->CreateTexture2D(&description, nullptr, staging.put()));
            }
            context->CopyResource(staging.get(), texture.get());
            const LONG scale = std::max({960L, visibleWidth, visibleHeight});
            const UINT width = std::max(1L, 960L * visibleWidth / scale);
            const UINT height = std::max(1L, 960L * visibleHeight / scale);
            std::vector<BYTE> pixels(width * height * 4);
            D3D11_MAPPED_SUBRESOURCE mapped{};
            check_hresult(context->Map(staging.get(), 0, D3D11_MAP_READ, 0, &mapped));
            for (UINT y = 0; y < height; ++y) {
                auto row = static_cast<const BYTE*>(mapped.pData) + (cropTop + y * visibleHeight / height) * mapped.RowPitch;
                for (UINT x = 0; x < width; ++x) {
                    size_t offset = (static_cast<size_t>(y) * width + x) * 4;
                    auto source = row + 4 * (cropLeft + x * visibleWidth / width);
                    std::copy_n(source, 3, pixels.data() + offset);
                    pixels[offset + 3] = 255;
                }
            }
            context->Unmap(staging.get(), 0);
            BYTE minGreen = 255, maxGreen = 0;
            for (size_t offset = 1; offset < pixels.size(); offset += 4) {
                minGreen = std::min(minGreen, pixels[offset]);
                maxGreen = std::max(maxGreen, pixels[offset]);
            }
            if (minGreen == maxGreen) {
                if (!loggedParkedUniform) stage("wgc_uniform_parked");
                loggedParkedUniform = true;
                continue;
            }
            if (!loggedParkedFrame) {
                stage("wgc_parked_frame_sampled");
                loggedParkedFrame = true;
            }
            if (width == prevWidth && height == prevHeight && previous == pixels) { ++unchanged; continue; }
            auto encoded = pngFrame(imaging.get(), pixels, width, height);
            if (encoded.empty()) { ++oversize; continue; }
            ++encodedCount;
            bytes += encoded.size();
            prevWidth = width;
            prevHeight = height;
            previous = std::move(pixels);
            emit("{\"sequence\":" + std::to_string(++sequence) + ",\"parked\":true" +
                ",\"png\":\"" + encoded + "\"}");
        }
        subscribed.revoke();
        session.Close();
        pool.Close();
    } catch (const hresult_error& error) {
        emit(std::string("{\"stage\":\"") + failureStage + "\",\"hresult\":" + std::to_string(static_cast<int>(error.code())) + "}");
    }
    SetEvent(prepared);
}
}

int wmain(int argc, wchar_t* argv[]) {
    const bool smoke = argc == 6 && wcscmp(argv[5], L"--smoke") == 0;
    if (argc != 5 && !smoke) return 2;
    const std::wstring surfaceName = std::wstring(L"Office Focus Preview Surface ") + argv[4];
    // Keep the process MTA alive across per-edit WGC sessions and UIA calls.
    MtaLifetime apartment;
    if (FAILED(apartment.status)) { stage("com_initialize_failed"); return 1; }
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    const DWORD officePid = wcstoul(argv[1], nullptr, 10);
    const DWORD codePid = wcstoul(argv[2], nullptr, 10);
    WindowSearch target{officePid, argv[3], nullptr};
    for (int i = 0; i < 60 && !target.result; ++i) { EnumWindows(findWindow, reinterpret_cast<LPARAM>(&target)); if (!target.result) Sleep(200); }
    HWND officeWindow = target.result;
    if (!officeWindow || !officePid || !codePid) { stage("office_window_not_found"); return 1; }
    RECT original{};
    GetWindowRect(officeWindow, &original);
    HWND code = nullptr;
    // Use the exact editor process, independent of other open editor windows.
    WindowSearch editor{codePid, L"", L"Chrome_WidgetWin_1"};
    for (int i = 0; i < 30 && !editor.result; ++i) { EnumWindows(findWindow, reinterpret_cast<LPARAM>(&editor)); if (!editor.result) Sleep(150); }
    code = editor.result;
    if (!code) { stage("editor_window_not_found"); return 1; }
    controlThread = GetCurrentThreadId();
    MSG message{};
    PeekMessageW(&message, nullptr, 0, 0, PM_NOREMOVE);
    auto events = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, nullptr, onForeground, 0, 0, WINEVENT_OUTOFCONTEXT);
    if (!events) { stage("foreground_hook_failed"); return 1; }
    handle capturePrepared{CreateEventW(nullptr, TRUE, FALSE, nullptr)};
    handle captureOffscreen{CreateEventW(nullptr, TRUE, FALSE, nullptr)};
    handle policyQuit{CreateEventW(nullptr, TRUE, FALSE, nullptr)};
    if (!capturePrepared || !captureOffscreen || !policyQuit) { UnhookWinEvent(events); return 1; }
    std::thread policyWatch;
    if (!smoke) policyWatch = std::thread(watchCapturePolicy, policyQuit.get());
    std::thread captureThread;
    auto stopCapture = [&] {
        if (!captureThread.joinable()) return;
        stopped = true;
        SetEvent(captureOffscreen.get());
        captureThread.join();
        stage("wgc_stopped");
    };
    std::thread([] {
        std::string line;
        while (std::getline(std::cin, line)) {
            if (line == "activate") PostThreadMessageW(controlThread, EnterOffice, 0, 0);
            if (line == "return-to-editor") PostThreadMessageW(controlThread, ReturnToEditor, 0, 0);
            if (line == "layout") PostThreadMessageW(controlThread, LayoutChanged, 0, 0);
            if (line == "clock") stage("native_clock");
            if (line == "smoke-click-preview") PostThreadMessageW(controlThread, SmokeClickPreview, 0, 0);
            if (line == "smoke-click-editor") PostThreadMessageW(controlThread, SmokeClickEditor, 0, 0);
            if (line == "smoke-move-pointer") PostThreadMessageW(controlThread, SmokeMovePointer, 0, 0);
            if (line == "smoke-raise-parked") PostThreadMessageW(controlThread, SmokeRaiseParked, 0, 0);
            if (line == "adopt off") PostThreadMessageW(controlThread, AdoptPolicyChanged, 0, 0);
            if (line == "adopt on") PostThreadMessageW(controlThread, AdoptPolicyChanged, 1, 0);
            if (line == "policy switch") PostThreadMessageW(controlThread, CapturePolicyChanged, static_cast<WPARAM>(CapturePolicy::Switch), 0);
            if (line == "policy always") PostThreadMessageW(controlThread, CapturePolicyChanged, static_cast<WPARAM>(CapturePolicy::Always), 0);
            if (line == "policy mouse") PostThreadMessageW(controlThread, CapturePolicyChanged, static_cast<WPARAM>(CapturePolicy::Mouse), 0);
            if (line == "exit") break;
        }
        PostThreadMessageW(controlThread, EndProbe, 0, 0);
    }).detach();
    bool editing = false;
    RECT pane{};
    bool parked = false;
    CapturePolicy policy = CapturePolicy::Always;
    UINT_PTR pointerTimer = 0;
    POINT pointerBaseline{};
    auto stopPointerWatch = [&] {
        if (pointerTimer) KillTimer(nullptr, pointerTimer);
        pointerTimer = 0;
    };
    auto prepareCapture = [&](const char* reason) {
        if (captureThread.joinable()) return;
        stage(reason);
        ResetEvent(capturePrepared.get());
        ResetEvent(captureOffscreen.get());
        stopped = false;
        captureThread = std::thread(captureWindow, officeWindow, capturePrepared.get(), captureOffscreen.get(), smoke);
    };
    auto prepareWhileEditing = [&] {
        stopPointerWatch();
        if (policy == CapturePolicy::Always) prepareCapture("capture_prewarm_always");
        if (policy == CapturePolicy::Mouse && GetCursorPos(&pointerBaseline)) {
            pointerTimer = SetTimer(nullptr, 0, 80, nullptr);
        }
    };
    auto park = [&] {
        stopPointerWatch();
        prepareCapture("capture_prepare_on_switch");
        if (WaitForSingleObject(capturePrepared.get(), 1500) != WAIT_OBJECT_0) stage("wgc_prepare_delayed");
        const LONG width = pane.right - pane.left;
        const LONG right = GetSystemMetrics(SM_XVIRTUALSCREEN) - std::max(128L, original.right - original.left);
        const RECT target{right - width, pane.top, right, pane.bottom};
        fitVisibleWindow(officeWindow, target, "native_park_alignment");
        RECT outer{};
        parked = GetWindowRect(officeWindow, &outer) && outer.right <= GetSystemMetrics(SM_XVIRTUALSCREEN);
        stage(parked ? "native_parked" : "native_park_unconfirmed");
        if (parked) SetEvent(captureOffscreen.get());
        else stopCapture();
    };
    auto focusOwner = [&] {
        const DWORD pid = windowPid(GetForegroundWindow());
        if (pid == officePid) return "office";
        if (pid == codePid) return "editor";
        return pid ? "other" : "none";
    };
    auto reportFocus = [&](const char* name) {
        emit(std::string("{\"stage\":\"") + name + "\",\"owner\":\"" + focusOwner() + "\"}");
    };
    auto requestFocus = [&](HWND wanted) {
        auto confirmed = [&] { return GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == wanted; };
        if (SetForegroundWindow(wanted)) {
            for (int i = 0; i < 20 && !confirmed(); ++i) Sleep(10);
        }
        if (confirmed()) return true;
        const DWORD owner = windowPid(GetForegroundWindow());
        if (owner != codePid && owner != officePid) return false;
        const HRESULT result = requestAccessibleFocus(wanted);
        emit(std::string("{\"stage\":\"uia_focus_result\",\"hresult\":") + std::to_string(result) + "}");
        for (int i = 0; SUCCEEDED(result) && i < 20 && !confirmed(); ++i) Sleep(10);
        return confirmed();
    };
    // Wait for the Webview layout and any transient Office startup dialog.
    const auto earliestPark = std::chrono::steady_clock::now() + std::chrono::milliseconds(1300);
    const UINT_PTR startupPoll = SetTimer(nullptr, 0, 250, nullptr);
    if (!startupPoll) stage("startup_timer_failed");
    BOOL officeWow64 = FALSE;
    handle officeProcess{OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, officePid)};
    const bool bitnessKnown = officeProcess && IsWow64Process(officeProcess.get(), &officeWow64);
    emit("{\"stage\":\"native_environment\",\"displayCount\":" + std::to_string(GetSystemMetrics(SM_CMONITORS)) +
        ",\"editorDpi\":" + std::to_string(GetDpiForWindow(code)) +
        ",\"officeBits\":" + std::to_string(bitnessKnown ? (officeWow64 ? 32 : 64) : 0) + "}");
    reportFocus("office_window_verified");
    bool initialized = false;
    bool modalReported = false;
    bool layoutReported = false;
    bool observedOffscreenFocus = false;
    bool firstSmokeClick = true;
    // D1: a parked Office window that takes the foreground (Alt+Tab, taskbar, its own dialog) is adopted into editing.
    bool adoptForeground = true;
    UINT_PTR raiseTimer = 0;
    UINT_PTR parkedWatch = 0;
    std::chrono::steady_clock::time_point raiseStarted;
    bool raiseObservedForeground = false;
    auto officeOffscreen = [&] {
        RECT outer{};
        return GetWindowRect(officeWindow, &outer) && outer.right <= GetSystemMetrics(SM_XVIRTUALSCREEN);
    };
    auto recenterOwnedDialog = [&](const RECT& area) {
        const HWND foreground = GetForegroundWindow();
        RECT dialog{};
        if (foreground == officeWindow || GetAncestor(foreground, GA_ROOTOWNER) != officeWindow ||
            !GetWindowRect(foreground, &dialog) || dialog.right > GetSystemMetrics(SM_XVIRTUALSCREEN)) return;
        const LONG width = dialog.right - dialog.left, height = dialog.bottom - dialog.top;
        const LONG left = std::max(area.left, (area.left + area.right - width) / 2);
        const LONG top = std::max(area.top, (area.top + area.bottom - height) / 2);
        const bool moved = SetWindowPos(foreground, nullptr, left, top, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
        stage(moved ? "office_dialog_recentered" : "office_dialog_recenter_failed");
    };
    auto adoptOfficeForeground = [&](bool watchdog) {
        stopPointerWatch();
        stopCapture();
        RECT current{};
        const bool docked = !IsIconic(code) && previewBounds(code, surfaceName, current);
        bool aligned = false;
        if (docked) {
            pane = current;
            aligned = fitVisibleWindow(officeWindow, pane, "native_edit_alignment");
        }
        if (!aligned) {
            // The tab is hidden or the editor is minimized: show Office where it was before docking.
            SetWindowPos(officeWindow, nullptr, original.left, original.top, original.right - original.left,
                original.bottom - original.top, SWP_NOZORDER | SWP_NOACTIVATE);
        }
        parked = false;
        editing = true;
        RECT area{};
        if (aligned) area = pane; else GetWindowRect(officeWindow, &area);
        recenterOwnedDialog(area);
        emit(std::string("{\"stage\":\"office_foreground_adopted\",\"docked\":") + (aligned ? "1" : "0") +
            ",\"watchdog\":" + (watchdog ? "1" : "0") +
            ",\"owner\":\"" + focusOwner() + "\"}");
        GUITHREADINFO focus{sizeof(GUITHREADINFO)};
        if (GetGUIThreadInfo(GetWindowThreadProcessId(officeWindow, nullptr), &focus) && windowPid(focus.hwndFocus) == officePid) {
            stage("office_keyboard_focus_confirmed");
        }
        prepareWhileEditing();
    };
    // The scheduled smoke runner may start behind the previous isolated editor.
    auto prepareEditorForeground = [&](const RECT& preview) {
        if (!requestFocus(code)) {
            const HRESULT focus = requestAccessibleFocus(code);
            emit("{\"stage\":\"smoke_editor_uia_focus\",\"hresult\":" + std::to_string(focus) + "}");
            for (int i = 0; SUCCEEDED(focus) && i < 12 && GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) != code; ++i) Sleep(10);
        }
        if (GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) != code &&
            SetWindowPos(code, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE)) {
            DwmFlush();
            const POINT header{preview.left + (preview.right - preview.left) * 2 / 3, preview.top - 12};
            const bool clicked = clickOwnedWindow(code, header);
            stage(clicked ? "smoke_editor_warm_click_sent" : "smoke_editor_warm_click_blocked");
            for (int i = 0; clicked && i < 24 && GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) != code; ++i) Sleep(10);
        }
        reportFocus("smoke_editor_focus_prepared");
        return GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == code;
    };
    auto finishRaise = [&] {
        KillTimer(nullptr, raiseTimer);
        raiseTimer = 0;
        const bool owns = GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == officeWindow;
        const bool offscreenAtDeadline = owns && officeOffscreen();
        const auto elapsed = std::chrono::duration_cast<std::chrono::milliseconds>(std::chrono::steady_clock::now() - raiseStarted).count();
        const bool recovered = owns && editing && !offscreenAtDeadline;
        // Baseline variant: hand the foreground back so the next round starts from the editor.
        if (owns && !editing) requestFocus(code);
        emit("{\"stage\":\"smoke_parked_raise_result\",\"raised\":" + std::to_string(raiseObservedForeground ? 1 : 0) +
            ",\"recovered\":" + std::to_string(recovered ? 1 : 0) +
            ",\"offscreenAtDeadline\":" + std::to_string(offscreenAtDeadline ? 1 : 0) +
            ",\"elapsedMs\":" + std::to_string(elapsed) +
            ",\"editorRestored\":" + std::to_string(GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == code ? 1 : 0) + "}");
    };
    while (GetMessageW(&message, nullptr, 0, 0) > 0) {
        if (message.message == EndProbe) break;
        if (message.message == WM_TIMER && message.wParam == startupPoll && !initialized &&
            std::chrono::steady_clock::now() >= earliestPark) {
            if (!IsWindow(officeWindow) || windowPid(officeWindow) != officePid) break;
            if (!IsWindowEnabled(officeWindow)) {
                if (!modalReported) reportFocus("office_modal_blocking");
                modalReported = true;
                continue;
            }
            if (!previewBounds(code, surfaceName, pane)) {
                if (!layoutReported) stage("preview_bounds_unavailable");
                layoutReported = true;
                continue;
            }
            KillTimer(nullptr, startupPoll);
            initialized = true;
            parkedWatch = SetTimer(nullptr, 0, 100, nullptr);
            if (!parkedWatch) stage("parked_watch_timer_failed");
            if (modalReported) reportFocus("office_modal_cleared");
            if (windowPid(GetForegroundWindow()) == officePid) requestFocus(code);
            reportFocus("initial_foreground");
            if (windowPid(GetForegroundWindow()) == officePid) {
                editing = true;
                prepareWhileEditing();
                stage("park_deferred_while_office_focused");
            } else park();
        }
        if (message.message == CapturePolicyChanged) {
            const auto requested = static_cast<CapturePolicy>(message.wParam);
            if (requested == policy) continue;
            policy = requested;
            stage(policy == CapturePolicy::Always ? "capture_policy_always" :
                policy == CapturePolicy::Mouse ? "capture_policy_mouse" : "capture_policy_switch");
            if (editing) {
                stopCapture();
                prepareWhileEditing();
            }
        }
        if (message.message == WM_TIMER && pointerTimer && message.wParam == pointerTimer && editing &&
            GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == officeWindow) {
            POINT point{};
            if (GetCursorPos(&point) && (std::abs(point.x - pointerBaseline.x) >= 4 || std::abs(point.y - pointerBaseline.y) >= 4)) {
                stopPointerWatch();
                prepareCapture("capture_prewarm_mouse");
            }
        }
        if (message.message == LayoutChanged && initialized) {
            RECT updated{};
            if (previewBounds(code, surfaceName, updated) && !EqualRect(&pane, &updated)) {
                pane = updated;
                if (editing && GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == officeWindow) {
                    fitVisibleWindow(officeWindow, pane, "native_edit_alignment");
                } else if (parked) park();
            }
        }
        if (smoke && message.message == SmokeClickPreview) {
            RECT preview{};
            const bool valid = initialized && parked && !IsIconic(code) && windowPid(code) == codePid &&
                previewBounds(code, surfaceName, preview);
            if (firstSmokeClick && valid && GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) != code) {
                prepareEditorForeground(preview);
            }
            firstSmokeClick = false;
            const bool ownsForeground = GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == code;
            bool sent = false;
            if (!valid || !ownsForeground) {
                stage(valid ? "smoke_click_editor_root_not_foreground" : "smoke_click_preview_geometry_unavailable");
            } else {
                DwmFlush();
                if (previewBounds(code, surfaceName, preview)) {
                    sent = clickOwnedWindow(code, {(preview.left + preview.right) / 2, (preview.top + preview.bottom) / 2});
                } else stage("smoke_click_preview_geometry_changed");
            }
            stage(sent ? "smoke_preview_click_sent" : "smoke_preview_click_blocked");
        }
        if (smoke && message.message == SmokeClickEditor) {
            const POINT header{pane.left + (pane.right - pane.left) * 2 / 3, pane.top - 12};
            const bool sent = editing && GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == officeWindow &&
                windowPid(code) == codePid && clickOwnedWindow(code, header);
            stage(sent ? "smoke_editor_click_sent" : "smoke_editor_click_blocked");
        }
        if (smoke && message.message == SmokeMovePointer && editing &&
            GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == officeWindow) {
            POINT point{};
            if (GetCursorPos(&point) && PtInRect(&pane, point)) SetCursorPos(point.x + 8, point.y);
        }
        if (message.message == AdoptPolicyChanged) {
            adoptForeground = message.wParam != 0;
            stage(adoptForeground ? "adopt_foreground_on" : "adopt_foreground_off");
        }
        if (smoke && message.message == SmokeRaiseParked) {
            if (!initialized || !parked || editing || raiseTimer) { stage("smoke_raise_precondition_failed"); continue; }
            // A real Alt+Tab or taskbar raise starts from the editor in the foreground and on top of the z-order.
            RECT preview{};
            if (GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) != code &&
                (!previewBounds(code, surfaceName, preview) || !prepareEditorForeground(preview))) {
                stage("smoke_raise_precondition_failed");
                continue;
            }
            raiseStarted = std::chrono::steady_clock::now();
            raiseObservedForeground = false;
            // Simulate an external raise (Alt+Tab, taskbar) without going through the preview click path.
            const bool direct = SetForegroundWindow(officeWindow) != FALSE;
            const HRESULT uia = direct ? S_OK : requestAccessibleFocus(officeWindow);
            emit(std::string("{\"stage\":\"smoke_parked_raise_sent\",\"hresult\":") + std::to_string(uia) +
                ",\"direct\":" + (direct ? "1" : "0") + "}");
            raiseTimer = SetTimer(nullptr, 0, 25, nullptr);
            if (!raiseTimer) stage("smoke_raise_timer_failed");
        }
        if (message.message == WM_TIMER && raiseTimer && message.wParam == raiseTimer) {
            const bool owns = GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == officeWindow;
            if (owns) raiseObservedForeground = true;
            const auto elapsed = std::chrono::steady_clock::now() - raiseStarted;
            if ((owns && editing && !officeOffscreen()) || elapsed >= std::chrono::milliseconds(850)) finishRaise();
        }
        if (message.message == EnterOffice) {
            const HWND foreground = GetForegroundWindow();
            if (!IsWindow(officeWindow) || windowPid(officeWindow) != officePid || windowPid(foreground) != codePid) {
                reportFocus("activation_precondition_failed");
                continue;
            }
            if (!IsWindowEnabled(officeWindow)) {
                reportFocus("office_modal_blocking");
                continue;
            }
            if (!initialized || !previewBounds(code, surfaceName, pane)) {
                stage("preview_bounds_unavailable");
                continue;
            }
            stopCapture();
            editing = requestFocus(officeWindow);
            if (editing) {
                parked = false;
                if (!fitVisibleWindow(officeWindow, pane, "native_edit_alignment")) {
                    stage("native_alignment_unconfirmed");
                    requestFocus(code);
                    editing = GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == officeWindow;
                    if (!editing) park();
                    continue;
                }
                GUITHREADINFO focus{sizeof(GUITHREADINFO)};
                if (GetGUIThreadInfo(GetWindowThreadProcessId(officeWindow, nullptr), &focus) && windowPid(focus.hwndFocus) == officePid) {
                    stage("office_keyboard_focus_confirmed");
                }
                prepareWhileEditing();
            }
            reportFocus(editing ? "native_foreground_confirmed" : "native_activation_denied");
            if (!editing) park();
        }
        if (message.message == ReturnToEditor && smoke && editing && GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == officeWindow && windowPid(code) == codePid) {
            stage(requestFocus(code) ? "smoke_editor_foreground" : "smoke_return_denied");
        }
        if (message.message == ForegroundChanged && adoptForeground && initialized && !editing &&
            GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == officeWindow) {
            adoptOfficeForeground(false);
        }
        // Some foreground changes (seen with UIA SetFocus) never deliver EVENT_SYSTEM_FOREGROUND, so poll while not editing.
        if (message.message == WM_TIMER && parkedWatch && message.wParam == parkedWatch && adoptForeground && initialized &&
            !editing && GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == officeWindow) {
            adoptOfficeForeground(true);
        }
        if (message.message == ForegroundChanged) {
            const bool officeOwnsForeground = GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) == officeWindow;
            RECT outer{};
            const bool offscreenFocus = officeOwnsForeground && parked && GetWindowRect(officeWindow, &outer) &&
                outer.right <= GetSystemMetrics(SM_XVIRTUALSCREEN);
            if (offscreenFocus && !observedOffscreenFocus) stage("screen_office_foreground_observed");
            observedOffscreenFocus = offscreenFocus;
            if (smoke || editing || offscreenFocus) reportFocus("foreground_observed");
        }
        if (message.message == ForegroundChanged && editing && GetAncestor(GetForegroundWindow(), GA_ROOTOWNER) != officeWindow) {
            editing = false;
            stage("office_foreground_released");
            RECT currentPane{};
            if (previewBounds(code, surfaceName, currentPane)) pane = currentPane;
            park();
        }
    }
    if (raiseTimer) KillTimer(nullptr, raiseTimer);
    if (parkedWatch) KillTimer(nullptr, parkedWatch);
    stopCapture();
    stopPointerWatch();
    SetEvent(policyQuit.get());
    if (policyWatch.joinable()) policyWatch.join();
    if (!initialized && startupPoll) KillTimer(nullptr, startupPoll);
    UnhookWinEvent(events);
    if (IsWindow(officeWindow) && windowPid(officeWindow) == officePid) {
        SetWindowPos(officeWindow, nullptr, original.left, original.top, original.right - original.left, original.bottom - original.top,
            SWP_NOZORDER | SWP_NOACTIVATE);
        wchar_t closingTitle[260]{};
        GetWindowTextW(officeWindow, closingTitle, 260);
        if (IsWindowEnabled(officeWindow) && std::wstring(closingTitle).rfind(target.prefix, 0) == 0 && PostMessageW(officeWindow, WM_CLOSE, 0, 0)) {
            stage("office_close_posted");
        }
    }
    stage("probe_exited");
}
