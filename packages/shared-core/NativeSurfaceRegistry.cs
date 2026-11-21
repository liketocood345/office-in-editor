using System.Collections.Concurrent;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;

namespace OfficeInEditor.Core;

// Explicit main-process slot leases. No UIA discovery or screen-position polling.
internal static class NativeSurfaceRegistry
{
    private sealed record Entry(string Token, IntPtr Window, int Pid, long StartTicks, long Generation, int Requester);
    private static readonly ConcurrentDictionary<string, Entry> entries = new(StringComparer.Ordinal);
    internal static JsonNode Register(JsonNode? value, int requester)
    {
        var token = Wire.Text(value, "surfaceToken", 80);
        if (!WebviewSurface.IsToken(token) || !long.TryParse(Wire.Text(value, "hwnd", 32), CultureInfo.InvariantCulture, out var raw) ||
            !long.TryParse(Wire.Text(value, "generation", 32), CultureInfo.InvariantCulture, out var generation) || generation <= 0)
            throw new CoreFault("invalid_native_surface", "Native surface identity is invalid.");
        var pid = Wire.Number(value, "pid");
        using var process = Process.GetProcessById(pid);
        if (process.ProcessName is not ("Code" or "Cursor" or "VSCodium"))
            throw new CoreFault("invalid_native_surface", "The native surface is not owned by the editor.");
        var entry = new Entry(token, new IntPtr(raw), pid, process.StartTime.ToUniversalTime().Ticks, generation, requester);
        if (!Valid(entry)) throw new CoreFault("invalid_native_surface", "The native slot lease could not be verified.");
        foreach (var stale in entries.Values.Where(item => !Valid(item))) entries.TryRemove(stale.Token, out _);
        if (entries.TryGetValue(token, out var prior) && prior != entry)
            throw new CoreFault("identity_conflict", "A native slot cannot be rebound while live.");
        if (!entries.ContainsKey(token) && entries.Count >= 32) throw new CoreFault("quota_exceeded", "Native surface limit reached.");
        entries[token] = entry;
        return Wire.Node(new { registered = true })!;
    }
    internal static JsonNode Unregister(JsonNode? value, int requester)
    {
        var token = Wire.Text(value, "surfaceToken", 80);
        if (!entries.TryGetValue(token, out var entry)) return Wire.Node(new { released = true })!;
        if (entry.Requester != requester) throw new CoreFault("unauthorized", "The native surface belongs to another host.");
        if (Valid(entry) && GetWindow(entry.Window, 5) != IntPtr.Zero)
            throw new CoreFault("document_still_attached", "Close the document before releasing its slot.");
        entries.TryRemove(token, out _);
        return Wire.Node(new { released = true })!;
    }
    internal static bool TryPlacement(string token, out WebviewSurface.Placement placement)
    {
        placement = default;
        if (!entries.TryGetValue(token, out var entry) || !Valid(entry) || !NativeWindow.GetClientRect(entry.Window, out var rect)) return false;
        var point = new NativeWindow.Point();
        if (!NativeWindow.ClientToScreen(entry.Window, ref point) || rect.Right < 32 || rect.Bottom < 32) return false;
        placement = new(entry.Window, point.X, point.Y, rect.Right, rect.Bottom, "native-slot");
        return true;
    }
    internal static bool IsRegistered(string token) => entries.ContainsKey(token);
    internal static void VerifyRequester(string token, int requester)
    {
        if (entries.TryGetValue(token, out var entry) && (entry.Requester != requester || !Valid(entry)))
            throw new CoreFault("invalid_native_surface", "The native slot is unavailable to this editor session.");
    }
    private static bool Valid(Entry entry)
    {
        if (!NativeWindow.IsWindow(entry.Window) || WebviewSurface.ClassName(entry.Window) != "OfficeInEditor.ExperimentalNativeSlot.1" ||
            GetProp(entry.Window, "OfficeInEditor.NativeSurface." + entry.Token).ToInt64() != entry.Generation) return false;
        GetWindowThreadProcessId(entry.Window, out var pid);
        if (pid != entry.Pid) return false;
        try { using var process = Process.GetProcessById(entry.Pid); return process.StartTime.ToUniversalTime().Ticks == entry.StartTicks; }
        catch (ArgumentException) { return false; }
    }
    [DllImport("user32.dll", CharSet=CharSet.Unicode, EntryPoint="GetPropW")] private static extern IntPtr GetProp(IntPtr hwnd, string key);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hwnd, uint command);
}
