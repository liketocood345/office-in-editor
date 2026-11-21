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

        private readonly Func<IntPtr, int[]> moveOutOfView;

        internal FocusDocuments(object application, string kind, DocumentRegistry documents,
            Func<IntPtr, IntPtr> frameOf = null, Func<string, IntPtr> frameByCaption = null, Func<IntPtr, int[]> moveOutOfView = null)
        {
            this.application = application;
            this.kind = kind;
            this.documents = documents;
            this.frameOf = frameOf ?? (window => NativeFrame(window, kind));
            this.frameByCaption = frameByCaption ?? (caption => NativeFrameByCaption(caption, kind));
            this.moveOutOfView = moveOutOfView ?? NativeMoveOutOfView;
        }

        internal static bool Handles(string method) =>
            method == "resolveDocumentWindow" || method == "openDocumentWindow" || method == "documentState" || method == "closeDocument";

        internal object Dispatch(string method, object parameters)
        {
            try
            {
                switch (method)
                {
                    case "resolveDocumentWindow": return Resolve(Json.Text(parameters, "path"));
                    case "openDocumentWindow": return Open(Json.Text(parameters, "path"));
                    case "documentState": return State(Json.Text(parameters, "documentId"));
                    case "closeDocument": return Close(Json.Text(parameters, "documentId"), Json.IsTrue(Json.Get(parameters, "save")));
                    default: throw new BridgeException("METHOD_NOT_FOUND");
                }
            }
            catch (BridgeException) { throw; }
            catch { throw new BridgeException("OFFICE_OPERATION_FAILED"); }
        }

        private static string CheckedPath(string path)
        {
            if (string.IsNullOrEmpty(path) || path.Length > 32767 || !Path.IsPathRooted(path) || path.StartsWith("\\\\", StringComparison.Ordinal))
                throw new BridgeException("INVALID_PATH");
            return Path.GetFullPath(path);
        }

        /// <summary>
        /// Opens a document for a focus tab without it ever showing on screen: it opens hidden, its frame moves out of
        /// view, and only then is it shown. The placement Office chose is returned so Core can put the frame back there.
        /// Word only for now; other applications answer "unsupported" and Core starts them the usual way.
        /// </summary>
        private object Open(string path)
        {
            string wanted = CheckedPath(path);
            object existing = Resolve(wanted);
            if (Json.IsTrue(Json.Get(existing, "found"))) return existing;
            if (kind != "word") return Json.Object("found", false, "reason", "unsupported");
            if (!File.Exists(wanted)) throw new BridgeException("FILE_NOT_FOUND");
            object missing = Type.Missing;
            dynamic document;
            // Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate,
            //   Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible)
            try { document = application.Documents.Open(wanted, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, false); }
            catch { throw new BridgeException("OPEN_FAILED"); }
            dynamic window = document.Windows.Item(1);
            IntPtr frame = IntPtr.Zero;
            try { frame = frameOf(new IntPtr(Convert.ToInt64(window.Hwnd, CultureInfo.InvariantCulture))); } catch { }
            int[] placement = frame == IntPtr.Zero ? null : moveOutOfView(frame);
            window.Visible = true;
            if (placement == null) return Json.Object("found", false, "reason", "window_unavailable");
            return Json.Object("found", true, "hwnd", frame.ToInt64(), "documentId", documents.Bind(document),
                "placement", Json.Object("left", placement[0], "top", placement[1], "right", placement[2], "bottom", placement[3], "show", placement[4]));
        }

        private object Resolve(string path)
        {
            string wanted = CheckedPath(path);
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
            // Core closes the frame Office may leave behind (empty) only when no document is left in this Office.
            int remaining = -1;
            try { remaining = documents.OpenDocuments().Count; } catch { }
            return Json.Object("documentId", documentId, "closed", true, "saved", save, "remaining", remaining);
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

        // Moves a hidden frame left of the virtual screen (it stays hidden) and returns the placement it had:
        // normal left, top, right, bottom and show command. Runs on Office's own thread, so the move is immediate.
        private static int[] NativeMoveOutOfView(IntPtr frame)
        {
            var placement = new WindowPlacement { Length = Marshal.SizeOf(typeof(WindowPlacement)) };
            if (!GetWindowPlacement(frame, ref placement) || !GetWindowRect(frame, out var rect)) return null;
            int show = placement.ShowCmd == 2 /* SW_SHOWMINIMIZED */ ? 1 : placement.ShowCmd;
            int left = GetSystemMetrics(76 /* SM_XVIRTUALSCREEN */) - (rect.Right - rect.Left) - 128;
            if (!SetWindowPos(frame, IntPtr.Zero, left, rect.Top, 0, 0, 0x0001 | 0x0004 | 0x0010 /* NOSIZE | NOZORDER | NOACTIVATE */)) return null;
            return new[] { placement.NormalLeft, placement.NormalTop, placement.NormalRight, placement.NormalBottom, show };
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct WindowPlacement
        {
            public int Length, Flags, ShowCmd, MinX, MinY, MaxX, MaxY, NormalLeft, NormalTop, NormalRight, NormalBottom;
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct Rect { public int Left, Top, Right, Bottom; }
        [DllImport("user32.dll")] private static extern bool GetWindowPlacement(IntPtr window, ref WindowPlacement placement);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out Rect rect);
        [DllImport("user32.dll")] private static extern int GetSystemMetrics(int index);
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr window, IntPtr after, int x, int y, int width, int height, uint flags);

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
