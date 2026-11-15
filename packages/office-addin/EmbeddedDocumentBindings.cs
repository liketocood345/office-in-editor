using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeInEditor.OfficeAddin
{
    // Metadata is pulled from the authenticated Core by this Office process.
    // OBJID_NATIVEOM ties the source to a precise Word document, not a PID-wide alias.
    internal sealed class EmbeddedDocumentBindings
    {
        internal sealed class Binding
        {
            internal object Document;
            internal object NativeWindow;
            internal string Path, SessionId;
            internal long Window;
            internal bool SourceVerified, Periodic, SaveOnClose = true;
            internal int Seconds;
        }
        private List<Binding> bindings = new List<Binding>();
        private readonly Func<long, object> nativeDocument;
        private readonly Func<long, bool> isWindow;
        private readonly Dictionary<long, object> nativeWindows = new Dictionary<long, object>();
        internal EmbeddedDocumentBindings(Func<long, object> nativeDocument = null, Func<long, bool> isWindow = null)
        { this.nativeDocument = nativeDocument ?? NativeDocument; this.isWindow = isWindow ?? (value => IsWindow(new IntPtr(value))); }
        internal List<Binding> Update(object value)
        {
            var next = new List<Binding>();
            object values = Json.Get(value, "bindings");
            if (values == null) values = Json.IsTrue(Json.Get(value, "bound")) ? new[] { value } : new object[0];
            foreach (object item in Json.Array(values))
            {
                if (next.Count >= 128) throw new BridgeException("TOO_MANY_BOUND_DOCUMENTS");
                string path = Json.Text(item, "sourcePath"), session = Json.Text(item, "sessionId");
                if (!LocalAutoSave.LocalPath(path, "word") || !File.Exists(path) || string.IsNullOrEmpty(session) ||
                    !session.StartsWith("ole-", StringComparison.Ordinal) || !Guid.TryParseExact(session.Substring(4), "N", out _))
                    throw new BridgeException("INVALID_SOURCE_BINDING");
                if (!long.TryParse(Convert.ToString(Json.Get(item, "hwnd"), CultureInfo.InvariantCulture), NumberStyles.None,
                    CultureInfo.InvariantCulture, out long window) || window <= 0 || !isWindow(window))
                    throw new BridgeException("INVALID_SOURCE_BINDING");
                object document = nativeDocument(window);
                if (document == null || next.Exists(entry => entry.SessionId == session || DocumentRegistry.SameObject(entry.Document, document)))
                    throw new BridgeException("INVALID_SOURCE_BINDING");
                int seconds = 60;
                object interval = Json.Get(item, "saveIntervalSeconds");
                if (interval != null) seconds = Convert.ToInt32(interval, CultureInfo.InvariantCulture);
                if (seconds < 5 || seconds > 3600) throw new BridgeException("INVALID_SAVE_INTERVAL");
                nativeWindows.TryGetValue(window, out object nativeWindow);
                next.Add(new Binding { Document = document, NativeWindow = nativeWindow, Path = Path.GetFullPath(path), SessionId = session, Window = window,
                    SourceVerified = Json.IsTrue(Json.Get(item, "sourceVerified")), Periodic = Json.IsTrue(Json.Get(item, "periodicSave")), Seconds = seconds,
                    SaveOnClose = Json.Get(item, "saveOnClose") == null || Json.IsTrue(Json.Get(item, "saveOnClose")) });
            }
            bindings = next;
            var expired = new List<long>();
            foreach (long window in nativeWindows.Keys) if (!next.Exists(entry => entry.Window == window)) expired.Add(window);
            foreach (long window in expired) nativeWindows.Remove(window);
            return new List<Binding>(bindings);
        }
        internal object Active()
        {
            object result = null;
            foreach (var binding in bindings)
            {
                if (!isWindow(binding.Window) || binding.NativeWindow == null) continue;
                bool active;
                try { active = Convert.ToBoolean(((dynamic)binding.NativeWindow).Active, CultureInfo.InvariantCulture); }
                catch { continue; }
                if (!active) continue;
                if (result != null) return null;
                result = binding.Document;
            }
            return result;
        }
        internal object SelectionRange(object document)
        {
            var binding = Find(document);
            return binding?.NativeWindow == null ? null : (object)((dynamic)binding.NativeWindow).Selection.Range;
        }
        internal Binding Find(object document) => document == null ? null : bindings.Find(entry =>
            DocumentRegistry.SameObject(entry.Document, document) && isWindow(entry.Window));
        internal string Resolve(object document) => Find(document)?.Path;
        internal List<object> Documents()
        {
            var result = new List<object>();
            foreach (var entry in bindings) if (isWindow(entry.Window)) result.Add(entry.Document);
            return result;
        }
        internal void Clear() { bindings.Clear(); nativeWindows.Clear(); }
        private object NativeDocument(long handle)
        {
            var window = new IntPtr(handle);
            GetWindowThreadProcessId(window, out uint pid);
            var name = new StringBuilder(256); GetClassName(window, name, name.Capacity);
            if (pid != Process.GetCurrentProcess().Id || name.ToString() != "_WwG") throw new BridgeException("INVALID_SOURCE_BINDING");
            Guid dispatch = new Guid("00020400-0000-0000-C000-000000000046");
            Marshal.ThrowExceptionForHR(AccessibleObjectFromWindow(window, 0xFFFFFFF0, ref dispatch, out object native));
            nativeWindows[handle] = native;
            return ((dynamic)native).Document;
        }
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int size);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr window, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IDispatch)] out object value);
    }
}
