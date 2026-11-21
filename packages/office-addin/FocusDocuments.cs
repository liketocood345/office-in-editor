using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace OfficeInEditor.OfficeAddin
{
    /// <summary>
    /// Add-in side of the focus-hybrid route. Core asks which top-level frame shows a document (D4: from the object
    /// model, never by searching window titles across processes), whether it has unsaved changes, and to close it.
    /// Paths arrive only from Core over the current-user pipe and are never logged or returned.
    /// </summary>
    internal sealed class FocusDocuments
    {
        private readonly dynamic application;
        private readonly string kind;
        private readonly DocumentRegistry documents;
        private readonly Func<IntPtr, IntPtr> frameOf;
        private readonly Func<string, IntPtr> frameByCaption;

        internal FocusDocuments(object application, string kind, DocumentRegistry documents,
            Func<IntPtr, IntPtr> frameOf = null, Func<string, IntPtr> frameByCaption = null)
        {
            this.application = application;
            this.kind = kind;
            this.documents = documents;
            this.frameOf = frameOf ?? (window => NativeFrame(window, kind));
            this.frameByCaption = frameByCaption ?? (caption => NativeFrameByCaption(caption, kind));
        }

        internal static bool Handles(string method) =>
            method == "resolveDocumentWindow" || method == "documentState" || method == "closeDocument";

        internal object Dispatch(string method, object parameters)
        {
            try
            {
                switch (method)
                {
                    case "resolveDocumentWindow": return Resolve(Json.Text(parameters, "path"));
                    case "documentState": return State(Json.Text(parameters, "documentId"));
                    case "closeDocument": return Close(Json.Text(parameters, "documentId"), Json.IsTrue(Json.Get(parameters, "save")));
                    default: throw new BridgeException("METHOD_NOT_FOUND");
                }
            }
            catch (BridgeException) { throw; }
            catch { throw new BridgeException("OFFICE_OPERATION_FAILED"); }
        }

        private object Resolve(string path)
        {
            if (string.IsNullOrEmpty(path) || path.Length > 32767 || !Path.IsPathRooted(path) || path.StartsWith("\\\\", StringComparison.Ordinal))
                throw new BridgeException("INVALID_PATH");
            string wanted = Path.GetFullPath(path);
            string wantedId = FileIdentity(wanted);
            foreach (object item in documents.OpenDocuments())
            {
                // Documents hosted in the OLE route are not top-level frames.
                if (documents.Embedded.Find(item) != null) continue;
                dynamic document = item;
                string fullName;
                try { fullName = Convert.ToString(document.FullName, CultureInfo.InvariantCulture); } catch { continue; }
                if (!SamePath(fullName, wanted, wantedId)) continue;
                IntPtr frame = FrameOf(document);
                if (frame == IntPtr.Zero) return Json.Object("found", false, "reason", "window_unavailable");
                return Json.Object("found", true, "hwnd", frame.ToInt64(), "documentId", documents.Bind(item));
            }
            return Json.Object("found", false);
        }

        private object State(string documentId)
        {
            dynamic document = documents.Resolve(documentId);
            bool dirty = true;
            try { dirty = !Convert.ToBoolean(document.Saved, CultureInfo.InvariantCulture); } catch { }
            return Json.Object("documentId", documentId, "dirty", dirty);
        }

        private object Close(string documentId, bool save)
        {
            dynamic document = documents.Resolve(documentId);
            try
            {
                switch (kind)
                {
                    case "word": document.Close(save ? -1 : 0); break;          // WdSaveOptions: wdSaveChanges / wdDoNotSaveChanges
                    case "excel": document.Close(save); break;
                    case "powerpoint":
                        // Presentation.Close has no save argument and never prompts, so save first when asked.
                        // Saved is never written (RISK_BOUNDARIES S06): it would hide Office's own unsaved-changes state.
                        if (save) document.Save();
                        document.Close();
                        break;
                    default: throw new BridgeException("UNSUPPORTED_APPLICATION");
                }
            }
            catch (BridgeException) { throw; }
            catch { throw new BridgeException(save ? "SAVE_FAILED" : "CLOSE_FAILED"); }
            return Json.Object("documentId", documentId, "closed", true, "saved", save);
        }

        private IntPtr FrameOf(dynamic document)
        {
            IntPtr window = IntPtr.Zero;
            try
            {
                switch (kind)
                {
                    case "word": window = new IntPtr(Convert.ToInt64(document.ActiveWindow.Hwnd, CultureInfo.InvariantCulture)); break;
                    case "excel": window = new IntPtr(Convert.ToInt64(document.Windows.Item(1).Hwnd, CultureInfo.InvariantCulture)); break;
                    case "powerpoint":
                        // DocumentWindow has no HWND on every build; fall back to this process's frame with the exact caption.
                        dynamic presentationWindow = document.Windows.Item(1);
                        try { window = new IntPtr(Convert.ToInt64(presentationWindow.HWND, CultureInfo.InvariantCulture)); }
                        catch { window = frameByCaption(Convert.ToString(presentationWindow.Caption, CultureInfo.InvariantCulture)); }
                        break;
                }
            }
            catch { return IntPtr.Zero; }
            return window == IntPtr.Zero ? IntPtr.Zero : frameOf(window);
        }

        private static bool SamePath(string fullName, string wanted, string wantedId)
        {
            if (string.IsNullOrEmpty(fullName) || fullName.Length < 3 || !Path.IsPathRooted(fullName)) return false;
            string candidate;
            try { candidate = Path.GetFullPath(fullName); } catch { return false; }
            if (string.Equals(candidate, wanted, StringComparison.OrdinalIgnoreCase)) return true;
            return wantedId != null && wantedId == FileIdentity(candidate);
        }

        internal static string FileIdentity(string path)
        {
            try
            {
                using (var handle = CreateFile(path, 0x80 /* FILE_READ_ATTRIBUTES */, 7 /* share all */, IntPtr.Zero, 3 /* OPEN_EXISTING */, 0x02000000 /* backup semantics */, IntPtr.Zero))
                {
                    if (handle.IsInvalid || !GetFileInformationByHandle(handle, out var info)) return null;
                    return info.VolumeSerialNumber.ToString("x8", CultureInfo.InvariantCulture) + ":" +
                        info.FileIndexHigh.ToString("x8", CultureInfo.InvariantCulture) + info.FileIndexLow.ToString("x8", CultureInfo.InvariantCulture);
                }
            }
            catch { return null; }
        }

        // Normalizes to the top-level frame and checks it belongs to this Office process and application.
        private static IntPtr NativeFrame(IntPtr window, string kind)
        {
            if (!IsWindow(window)) return IntPtr.Zero;
            IntPtr root = GetAncestor(window, 2 /* GA_ROOT */);
            if (root == IntPtr.Zero) return IntPtr.Zero;
            GetWindowThreadProcessId(root, out uint pid);
            if (pid != (uint)System.Diagnostics.Process.GetCurrentProcess().Id) return IntPtr.Zero;
            return ClassName(root) == FrameClass(kind) ? root : IntPtr.Zero;
        }

        private static IntPtr NativeFrameByCaption(string caption, string kind)
        {
            if (string.IsNullOrEmpty(caption)) return IntPtr.Zero;
            uint self = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
            string frameClass = FrameClass(kind);
            IntPtr found = IntPtr.Zero;
            int matches = 0;
            EnumWindows((window, _) =>
            {
                GetWindowThreadProcessId(window, out uint pid);
                if (pid != self || ClassName(window) != frameClass) return true;
                var title = new StringBuilder(512);
                GetWindowText(window, title, title.Capacity);
                string text = title.ToString();
                if (text == caption || text.StartsWith(caption + " - ", StringComparison.Ordinal)) { found = window; matches++; }
                return true;
            }, IntPtr.Zero);
            // Two frames with the same caption cannot be told apart: report none rather than guess.
            return matches == 1 ? found : IntPtr.Zero;
        }

        private static string FrameClass(string kind) =>
            kind == "word" ? "OpusApp" : kind == "excel" ? "XLMAIN" : kind == "powerpoint" ? "PPTFrameClass" : "";

        private static string ClassName(IntPtr window)
        {
            var name = new StringBuilder(64);
            GetClassName(window, name, name.Capacity);
            return name.ToString();
        }

        private delegate bool EnumProc(IntPtr window, IntPtr data);
        [StructLayout(LayoutKind.Sequential)]
        private struct FileInformation
        {
            public uint Attributes;
            public long Created, Accessed, Written;
            public uint VolumeSerialNumber, SizeHigh, SizeLow, Links, FileIndexHigh, FileIndexLow;
        }
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc callback, IntPtr data);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int size);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int size);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(string name, uint access, uint share, IntPtr security, uint disposition, uint flags, IntPtr template);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool GetFileInformationByHandle(SafeFileHandle file, out FileInformation information);
    }
}
