// oie-focus-capture: Windows.Graphics.Capture helper for the focus-hybrid route.
// Core starts one per focus session and owns every window decision; this process only captures and encodes.
//
// Usage: oie-focus-capture.exe <hwnd> <pid> [--any-position]
//   --any-position  publish frames wherever the window is (lab experiments only; Core never passes it)
// stdin (one command per line):
//   prepare          create the capture session without starting it (pre-warm)
//   publish <id>     the window is parked: start capture and send frames tagged with park id <id>
//   stop             end the capture session (window is being edited); the capture item is kept for the next one
//   fps <n>          frame-rate cap, 0..60; 0 pauses sending (the tab is hidden). Default 30.
//   format <f>       jpeg (default, fast enough for 30 fps) or png (lossless)
//   exit / EOF       stop and exit; Core closing stdin always ends the helper
// stdout (binary): u32 LE length of (kind + body), u8 kind, body
//   kind 1: status JSON {"stage":...} (no titles, paths or text); "capture_stats" once a second while publishing
//   kind 2: frame: u16 LE header length, header JSON {"parkId","width","height","format"}, image bytes
// Frames are sent only while the window is off-screen (unless --any-position), only when pixels change, and at
// most at the fps cap. The first frame after each publish, or after leaving fps 0, is always sent.
#define NOMINMAX
#include <windows.h>
#include <d3d11.h>
#include <dwmapi.h>
#include <roapi.h>
#include <shobjidl.h>
#include <wincodec.h>
#include <windows.graphics.capture.interop.h>
#include <windows.graphics.directx.direct3d11.interop.h>
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.Graphics.Capture.h>
#include <winrt/Windows.Graphics.DirectX.Direct3D11.h>
#include <algorithm>
#include <atomic>
#include <chrono>
#include <cstdio>
#include <iostream>
#include <mutex>
#include <string>
#include <thread>
#include <vector>
#include <fcntl.h>
#include <io.h>

using namespace winrt;
using namespace winrt::Windows::Graphics::Capture;
using namespace winrt::Windows::Graphics::DirectX;
using namespace winrt::Windows::Graphics::DirectX::Direct3D11;

namespace {
constexpr size_t MaxFrameBytes = 30u * 1024u * 1024u;
constexpr float JpegQuality = 0.9f;
std::mutex outputLock;
HWND target{};
DWORD targetPid{};
bool anyPosition = false;

void writeMessage(unsigned char kind, const std::string& head, const std::vector<BYTE>* tail = nullptr) {
    std::lock_guard<std::mutex> lock(outputLock);
    const size_t body = head.size() + (tail ? tail->size() : 0);
    const uint32_t length = static_cast<uint32_t>(1 + body);
    unsigned char prefix[5]{static_cast<unsigned char>(length), static_cast<unsigned char>(length >> 8),
        static_cast<unsigned char>(length >> 16), static_cast<unsigned char>(length >> 24), kind};
    fwrite(prefix, 1, sizeof(prefix), stdout);
    fwrite(head.data(), 1, head.size(), stdout);
    if (tail && !tail->empty()) fwrite(tail->data(), 1, tail->size(), stdout);
    fflush(stdout);
}

void stage(const std::string& name, const std::string& extra = "") {
    writeMessage(1, "{\"stage\":\"" + name + "\"" + extra + "}");
}

void frameMessage(int parkId, UINT width, UINT height, bool jpeg, const std::vector<BYTE>& image) {
    const std::string header = "{\"parkId\":" + std::to_string(parkId) + ",\"width\":" + std::to_string(width) +
        ",\"height\":" + std::to_string(height) + ",\"format\":\"" + (jpeg ? "jpeg" : "png") + "\"}";
    std::string head(2, '\0');
    head[0] = static_cast<char>(header.size() & 0xff);
    head[1] = static_cast<char>((header.size() >> 8) & 0xff);
    head += header;
    writeMessage(2, head, &image);
}

struct Bounds { RECT outer{}, visible{}; };
Bounds windowBounds(HWND window) {
    Bounds bounds{};
    GetWindowRect(window, &bounds.outer);
    if (FAILED(DwmGetWindowAttribute(window, DWMWA_EXTENDED_FRAME_BOUNDS, &bounds.visible, sizeof(RECT)))) bounds.visible = bounds.outer;
    return bounds;
}

bool targetAlive() {
    DWORD pid{};
    return IsWindow(target) && GetWindowThreadProcessId(target, &pid) && pid == targetPid;
}

std::vector<BYTE> encodeImage(IWICImagingFactory* factory, const std::vector<BYTE>& pixels, UINT width, UINT height, bool jpeg) {
    com_ptr<IStream> stream;
    check_hresult(CreateStreamOnHGlobal(nullptr, TRUE, stream.put()));
    com_ptr<IWICBitmapEncoder> encoder;
    check_hresult(factory->CreateEncoder(jpeg ? GUID_ContainerFormatJpeg : GUID_ContainerFormatPng, nullptr, encoder.put()));
    check_hresult(encoder->Initialize(stream.get(), WICBitmapEncoderNoCache));
    com_ptr<IWICBitmapFrameEncode> frame;
    com_ptr<IPropertyBag2> options;
    check_hresult(encoder->CreateNewFrame(frame.put(), options.put()));
    if (jpeg && options) {
        PROPBAG2 option{};
        option.pstrName = const_cast<LPOLESTR>(L"ImageQuality");
        VARIANT value{};
        value.vt = VT_R4;
        value.fltVal = JpegQuality;
        check_hresult(options->Write(1, &option, &value));
    }
    check_hresult(frame->Initialize(options.get()));
    check_hresult(frame->SetSize(width, height));
    // JPEG has no alpha: hand it 24-bit pixels; PNG keeps the 32-bit buffer as is.
    auto format = jpeg ? GUID_WICPixelFormat24bppBGR : GUID_WICPixelFormat32bppBGRA;
    check_hresult(frame->SetPixelFormat(&format));
    if (jpeg) {
        if (format != GUID_WICPixelFormat24bppBGR) throw hresult_error(E_FAIL);
        std::vector<BYTE> rgb(static_cast<size_t>(width) * height * 3);
        for (size_t p = 0, q = 0; p < pixels.size(); p += 4, q += 3) { rgb[q] = pixels[p]; rgb[q + 1] = pixels[p + 1]; rgb[q + 2] = pixels[p + 2]; }
        check_hresult(frame->WritePixels(height, width * 3, static_cast<UINT>(rgb.size()), rgb.data()));
    } else {
        if (format != GUID_WICPixelFormat32bppBGRA) throw hresult_error(E_FAIL);
        check_hresult(frame->WritePixels(height, width * 4, static_cast<UINT>(pixels.size()), const_cast<BYTE*>(pixels.data())));
    }
    check_hresult(frame->Commit());
    check_hresult(encoder->Commit());
    STATSTG size{};
    check_hresult(stream->Stat(&size, STATFLAG_NONAME));
    if (size.cbSize.QuadPart > MaxFrameBytes) return {};
    HGLOBAL memory{};
    check_hresult(GetHGlobalFromStream(stream.get(), &memory));
    auto bytes = static_cast<const BYTE*>(GlobalLock(memory));
    if (!bytes) throw hresult_error(E_FAIL);
    std::vector<BYTE> result(bytes, bytes + size.cbSize.LowPart);
    GlobalUnlock(memory);
    return result;
}

// One capture session: prepared on construction, started by publish, ended by stop.
class Capture {
public:
    void prepare() {
        // A session that ended on its own (it failed) is replaced: the next park tries again.
        if (worker.joinable() && ended) worker.join();
        if (worker.joinable()) return;
        stopping = false;
        started = false;
        ended = false;
        worker = std::thread([this] { run(); ended = true; });
    }

    void publish(int id) {
        parkId = id;
        fresh = true;
        prepare();
        started = true;
        SetEvent(wake.get());
    }

    void setFps(int value) {
        const int previous = fps.exchange(std::clamp(value, 0, 60));
        if (previous == 0 && value > 0) fresh = true;
    }

    void setJpeg(bool value) { jpeg = value; fresh = true; }

    void stop() {
        if (!worker.joinable()) return;
        stopping = true;
        SetEvent(wake.get());
        worker.join();
        stage("capture_stopped");
    }

private:
    std::thread worker;
    std::atomic_bool stopping{false}, started{false}, fresh{false}, jpeg{true}, ended{false};
    std::atomic_int parkId{0}, fps{30};
    handle wake{CreateEventW(nullptr, FALSE, FALSE, nullptr)};
    // Created once and reused by every session: a window without a taskbar button (Core hides it while the
    // window is hosted) cannot get a new capture item, but an existing item keeps working.
    GraphicsCaptureItem item{nullptr};

    static GraphicsCaptureItem createItem() {
        auto interop = get_activation_factory<GraphicsCaptureItem, IGraphicsCaptureItemInterop>();
        GraphicsCaptureItem created{nullptr};
        const HRESULT first = interop->CreateForWindow(target, guid_of<GraphicsCaptureItem>(), put_abi(created));
        if (SUCCEEDED(first)) return created;
        if (first != E_INVALIDARG) throw_hresult(first);
        // E_INVALIDARG: the window has no taskbar button. Put it back just long enough to create the item. Core
        // removes the button again on its own schedule and may do so in between, so try a few times.
        com_ptr<ITaskbarList> taskbar;
        check_hresult(CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(taskbar.put())));
        check_hresult(taskbar->HrInit());
        HRESULT retry = first;
        for (int attempt = 1; attempt <= 5; ++attempt) {
            check_hresult(taskbar->AddTab(target));
            retry = interop->CreateForWindow(target, guid_of<GraphicsCaptureItem>(), put_abi(created));
            taskbar->DeleteTab(target);
            if (SUCCEEDED(retry)) {
                stage("capture_item_tab_restored", ",\"attempts\":" + std::to_string(attempt));
                return created;
            }
            if (retry != E_INVALIDARG) break;
            Sleep(150);
        }
        throw_hresult(retry);
    }

    void run() {
        const char* failure = "capture_device_failed";
        try {
            check_hresult(RoInitialize(RO_INIT_MULTITHREADED));
            struct Uninit { ~Uninit() { RoUninitialize(); } } uninit;
            if (!GraphicsCaptureSession::IsSupported()) { stage("capture_unsupported"); return; }
            com_ptr<ID3D11Device> device;
            com_ptr<ID3D11DeviceContext> context;
            check_hresult(D3D11CreateDevice(nullptr, D3D_DRIVER_TYPE_HARDWARE, nullptr, D3D11_CREATE_DEVICE_BGRA_SUPPORT,
                nullptr, 0, D3D11_SDK_VERSION, device.put(), nullptr, context.put()));
            com_ptr<IInspectable> inspectable;
            check_hresult(CreateDirect3D11DeviceFromDXGIDevice(device.as<IDXGIDevice>().get(), inspectable.put()));
            auto direct = inspectable.as<IDirect3DDevice>();
            failure = "capture_item_failed";
            if (!item) item = createItem();
            failure = "capture_session_failed";
            auto size = item.Size();
            if (size.Width <= 0 || size.Height <= 0) {
                const auto bounds = windowBounds(target).visible;
                size = {bounds.right - bounds.left, bounds.bottom - bounds.top};
            }
            auto pool = Direct3D11CaptureFramePool::CreateFreeThreaded(direct, DirectXPixelFormat::B8G8R8A8UIntNormalized, 2, size);
            auto session = pool.CreateCaptureSession(item);
            try { session.IsCursorCaptureEnabled(false); } catch (const hresult_error&) {}
            try { session.IsBorderRequired(false); } catch (const hresult_error&) {}
            handle arrival{CreateEventW(nullptr, FALSE, FALSE, nullptr)};
            auto subscribed = pool.FrameArrived(auto_revoke, [&arrival](auto const&, auto const&) { SetEvent(arrival.get()); });
            com_ptr<IWICImagingFactory> imaging;
            check_hresult(CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(imaging.put())));
            stage("capture_prepared");
            while (!stopping && !started) WaitForSingleObject(wake.get(), INFINITE);
            if (stopping) return;
            failure = "capture_start_failed";
            session.StartCapture();
            stage("capture_started");
            failure = "capture_frame_failed";
            com_ptr<ID3D11Texture2D> staging;
            std::vector<BYTE> previous;
            UINT previousWidth = 0, previousHeight = 0;
            auto nextFrame = std::chrono::steady_clock::now();
            auto statsSince = nextFrame;
            unsigned arrived = 0, sent = 0, unchanged = 0;
            long long encodeUs = 0;
            size_t bytesSent = 0;
            bool reportedUniform = false;
            while (!stopping && targetAlive()) {
                const auto waited = WaitForSingleObject(arrival.get(), 100);
                const auto now = std::chrono::steady_clock::now();
                if (now - statsSince >= std::chrono::seconds(1)) {
                    stage("capture_stats", ",\"arrived\":" + std::to_string(arrived) + ",\"sent\":" + std::to_string(sent) +
                        ",\"unchanged\":" + std::to_string(unchanged) + ",\"encodeMsAvg\":" + std::to_string(sent ? encodeUs / sent / 1000 : 0) +
                        ",\"bytes\":" + std::to_string(bytesSent) + ",\"fps\":" + std::to_string(fps.load()));
                    statsSince = now; arrived = sent = unchanged = 0; encodeUs = 0; bytesSent = 0;
                }
                if (waited != WAIT_OBJECT_0 || stopping) continue;
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
                const int cap = fps.load();
                if (cap <= 0) continue;  // Hidden tab: keep the session warm, send nothing.
                if (now < nextFrame && !fresh) continue;
                const auto bounds = windowBounds(target);
                const LONG visibleWidth = bounds.visible.right - bounds.visible.left;
                const LONG visibleHeight = bounds.visible.bottom - bounds.visible.top;
                // Only a parked window is published: frames never show the window while it is on screen.
                if (visibleWidth <= 0 || visibleHeight <= 0) continue;
                if (!anyPosition && bounds.outer.right > GetSystemMetrics(SM_XVIRTUALSCREEN)) continue;
                // WGC returns either the visible rectangle or the full Win32 rectangle with resize margins.
                const LONG padX = size.Width - visibleWidth, padY = size.Height - visibleHeight;
                if (padX < 0 || padY < 0 || padX > 64 || padY > 64) continue;
                const LONG cropLeft = size.Width == bounds.outer.right - bounds.outer.left
                    ? std::clamp(bounds.visible.left - bounds.outer.left, 0L, padX) : padX / 2;
                const LONG cropTop = size.Height == bounds.outer.bottom - bounds.outer.top
                    ? std::clamp(bounds.visible.top - bounds.outer.top, 0L, padY) : padY / 2;
                auto access = frame.Surface().as<::Windows::Graphics::DirectX::Direct3D11::IDirect3DDxgiInterfaceAccess>();
                com_ptr<ID3D11Texture2D> texture;
                check_hresult(access->GetInterface(IID_PPV_ARGS(texture.put())));
                if (!staging) {
                    D3D11_TEXTURE2D_DESC description{};
                    texture->GetDesc(&description);
                    description.BindFlags = description.MiscFlags = 0;
                    description.Usage = D3D11_USAGE_STAGING;
                    description.CPUAccessFlags = D3D11_CPU_ACCESS_READ;
                    check_hresult(device->CreateTexture2D(&description, nullptr, staging.put()));
                }
                context->CopyResource(staging.get(), texture.get());
                const UINT width = static_cast<UINT>(visibleWidth), height = static_cast<UINT>(visibleHeight);
                std::vector<BYTE> pixels(static_cast<size_t>(width) * height * 4);
                D3D11_MAPPED_SUBRESOURCE mapped{};
                check_hresult(context->Map(staging.get(), 0, D3D11_MAP_READ, 0, &mapped));
                for (UINT y = 0; y < height; ++y) {
                    auto row = static_cast<const BYTE*>(mapped.pData) + static_cast<size_t>(cropTop + y) * mapped.RowPitch + 4 * cropLeft;
                    auto out = pixels.data() + static_cast<size_t>(y) * width * 4;
                    std::copy_n(row, static_cast<size_t>(width) * 4, out);
                    for (UINT x = 0; x < width; ++x) out[x * 4 + 3] = 255;
                }
                context->Unmap(staging.get(), 0);
                // A uniform frame means DWM has not painted the parked window yet.
                BYTE low = 255, high = 0;
                for (size_t offset = 1; offset < pixels.size(); offset += 4) { low = std::min(low, pixels[offset]); high = std::max(high, pixels[offset]); }
                if (low == high) { if (!reportedUniform) stage("capture_uniform_frame"); reportedUniform = true; continue; }
                const bool first = fresh.exchange(false);
                if (!first && width == previousWidth && height == previousHeight && previous == pixels) { ++unchanged; continue; }
                const auto encodeStart = std::chrono::steady_clock::now();
                const bool asJpeg = jpeg.load();
                auto image = encodeImage(imaging.get(), pixels, width, height, asJpeg);
                encodeUs += std::chrono::duration_cast<std::chrono::microseconds>(std::chrono::steady_clock::now() - encodeStart).count();
                if (image.empty()) { stage("capture_frame_oversize"); continue; }
                nextFrame = now + std::chrono::microseconds(1000000 / cap);
                frameMessage(parkId, width, height, asJpeg, image);
                ++sent; bytesSent += image.size();
                previousWidth = width; previousHeight = height;
                previous = std::move(pixels);
            }
            subscribed.revoke();
            session.Close();
            pool.Close();
            if (!targetAlive()) stage("capture_target_gone");
        } catch (const hresult_error& error) {
            stage(failure, ",\"hresult\":" + std::to_string(static_cast<int>(error.code())));
        }
    }
};
}

int wmain(int argc, wchar_t* argv[]) {
    if (argc != 3 && !(argc == 4 && wcscmp(argv[3], L"--any-position") == 0)) return 2;
    anyPosition = argc == 4;
    _setmode(_fileno(stdout), _O_BINARY);
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    target = reinterpret_cast<HWND>(static_cast<INT_PTR>(_wcstoi64(argv[1], nullptr, 10)));
    targetPid = wcstoul(argv[2], nullptr, 10);
    if (!targetAlive()) { stage("capture_target_invalid"); return 1; }
    stage("capture_ready");
    Capture capture;
    std::string line;
    while (std::getline(std::cin, line)) {
        if (!line.empty() && line.back() == '\r') line.pop_back();
        if (line == "prepare") capture.prepare();
        else if (line.rfind("publish ", 0) == 0) {
            const int id = std::atoi(line.c_str() + 8);
            if (id > 0) capture.publish(id);
        }
        else if (line.rfind("fps ", 0) == 0) capture.setFps(std::atoi(line.c_str() + 4));
        else if (line == "format jpeg") capture.setJpeg(true);
        else if (line == "format png") capture.setJpeg(false);
        else if (line == "stop") capture.stop();
        else if (line == "exit") break;
    }
    capture.stop();
    return 0;
}
