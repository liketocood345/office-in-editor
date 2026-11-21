using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

// Exact prefix of the pinned dual interface's native vtable. The four
// IDispatch slots precede Activate. PreserveSig exposes the real HRESULT,
// independently of the dynamic binder/type-library invocation path.
[ComImport, Guid("70B8EE71-2B79-4D45-9A77-AE395CFEA071"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IProbeFramer
{
    [PreserveSig] int GetTypeInfoCount(out uint count);
    [PreserveSig] int GetTypeInfo(uint index, uint locale, out IntPtr info);
    [PreserveSig] int GetIDsOfNames(ref Guid iid, IntPtr names, uint count, uint locale, IntPtr ids);
    [PreserveSig] int Invoke(int id, ref Guid iid, uint locale, ushort flags, IntPtr parameters, IntPtr result, IntPtr exception, IntPtr argument);
    [PreserveSig] int Activate();
    [PreserveSig] int ActiveDocument([MarshalAs(UnmanagedType.IDispatch)] out object document);
    [PreserveSig] int CreateNew([MarshalAs(UnmanagedType.BStr)] string progId);
    [PreserveSig] int Open([MarshalAs(UnmanagedType.Struct)] object document, [MarshalAs(UnmanagedType.Struct)] object readOnly,
        [MarshalAs(UnmanagedType.Struct)] object progId, [MarshalAs(UnmanagedType.Struct)] object user, [MarshalAs(UnmanagedType.Struct)] object password);
    [PreserveSig] int Save([MarshalAs(UnmanagedType.Struct)] object target, [MarshalAs(UnmanagedType.Struct)] object overwrite,
        [MarshalAs(UnmanagedType.Struct)] object user, [MarshalAs(UnmanagedType.Struct)] object password);
    [PreserveSig] int PrintOutOld([MarshalAs(UnmanagedType.Struct)] object prompt);
    [PreserveSig] int Close();
}

[ComImport, Guid("0000011B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IProbeOleContainer
{
    [PreserveSig] int ParseDisplayName(IntPtr context, [MarshalAs(UnmanagedType.LPWStr)] string name, out uint eaten, out IntPtr moniker);
    [PreserveSig] int EnumObjects(uint flags, out IProbeEnumUnknown objects);
    [PreserveSig] int LockContainer([MarshalAs(UnmanagedType.Bool)] bool locked);
}
[ComImport, Guid("00000100-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IProbeEnumUnknown
{
    [PreserveSig] int Next(uint count, out IntPtr item, out uint fetched);
    [PreserveSig] int Skip(uint count);
    [PreserveSig] int Reset();
    [PreserveSig] int Clone(out IProbeEnumUnknown copy);
}

internal sealed class DsoDocumentProbe
{
    internal static object GetHostedControl(object htmlDocument)
    {
        // Native host control access, not DOM scripting. IObjectSafety and
        // browser policy stay unchanged; the page receives no automation API.
        IProbeEnumUnknown objects;
        Marshal.ThrowExceptionForHR(((IProbeOleContainer)htmlDocument).EnumObjects(1, out objects));
        try
        {
            for (int i = 0; i < 4; i++)
            {
                IntPtr pointer; uint fetched;
                int hr = objects.Next(1, out pointer, out fetched);
                Marshal.ThrowExceptionForHR(hr);
                if (hr != 0 || fetched != 1) break;
                object candidate;
                try { candidate = Marshal.GetObjectForIUnknown(pointer); } finally { Marshal.Release(pointer); }
                if (candidate is IProbeFramer) return candidate;
                ProbeRun.Release(candidate);
            }
        }
        finally { ProbeRun.Release(objects); }
        throw new InvalidOperationException("Private hosted control not found");
    }
    private sealed class Framer : AxHost
    {
        internal Framer() : base(TridentProbe.ControlId) { }
        internal object Instance { get { return GetOcx(); } }
    }
    private readonly Form form;
    private readonly Dictionary<string, object> report;
    private readonly string document, root, operation;
    private readonly Func<object> htmlFactory;
    private readonly IntPtr htmlHost;
    private readonly bool manual;
    internal DsoDocumentProbe(Form form, Dictionary<string, object> report, string document, string root, string operation, Func<object> htmlFactory = null, bool manual = false, IntPtr htmlHost = default(IntPtr))
    { this.form = form; this.report = report; this.document = document; this.root = root; this.operation = operation; this.htmlFactory = htmlFactory; this.manual = manual; this.htmlHost = htmlHost; }

    internal async Task ExecuteAsync()
    {
        ProbeRun run = null;
        Framer framer = null;
        object control = null, docObject = null, appObject = null;
        bool opened = false, closed = false, cleanup = false, loaded = false;
        int appPid = 0;
        IProbeFramer native = null;
        try
        {
            run = new ProbeRun(root, report, Path.GetExtension(document));
            if (run.Application != "word" && (operation != "open" || manual || htmlFactory == null))
                throw new InvalidOperationException("This format requires an HTML open probe");
            run.ValidateDocument(document);
            report["operation"] = operation;
            report["progId"] = run.ProgId;
            run.Stage("create-private-control");
            string ocxFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "oie-legacy-framer.ocx");
            report["controlSha256"] = ProbeRun.Hash(ocxFile);
            int initialStage, initialHr;
            OieProbeTrace(-1, out initialStage, out initialHr); // retain the diagnostics module
            form.Activate();
            if (htmlFactory != null) control = htmlFactory();
            else
            {
                framer = new Framer { Dock = DockStyle.Fill };
                form.Controls.Add(framer);
                framer.CreateControl();
                framer.Select();
                framer.Focus();
                control = framer.Instance;
            }
            report["presentation"] = htmlFactory != null ? "html-object" : "native-control";
            native = (IProbeFramer)control;
            report["controlInstance"] = true;
            run.Stage(operation == "create" ? "native-create-new" : "native-open");
            // One operation per fresh run, with no silent fallback or blind retry.
            int hr = operation == "create" ? native.CreateNew(run.ProgId) :
                native.Open(run.Document, false, run.ProgId, null, null);
            Check(hr, "loadHresult");
            loaded = true;
            await Task.Delay(250); // return to the STA message pump without reentrant DoEvents
            Check(native.Activate(), "activateHresult");
            run.Stage("native-active-document");
            hr = native.ActiveDocument(out docObject);
            Check(hr, "activeDocumentHresult");
            if (docObject == null) throw new InvalidOperationException("No active document after native load");
            opened = true;
            report["documentOpened"] = true;
            dynamic doc = docObject;
            IntPtr docWindow = OieProbeDocumentWindow();
            if (docWindow == IntPtr.Zero) throw new InvalidOperationException("No in-place document HWND");
            appPid = run.BindNativeWindow(docWindow);
            IntPtr hostWindow = htmlFactory != null ? htmlHost : form.Handle;
            bool nativeChild = IsChild(hostWindow, docWindow);
            bool visible = IsWindowVisible(docWindow);
            report["nativeOfficeChildVerified"] = nativeChild;
            report["documentWindowVisible"] = visible;
            report["officeHitSamples"] = CountDocumentHits(docWindow);
            if (run.Application == "word") report["nativeWordChildVerified"] = nativeChild;
            try { appObject = doc.Application; }
            catch (COMException) { report["applicationAutomationUnavailable"] = true; }
            var classes = new List<string>();
            EnumChildWindows(hostWindow, delegate(IntPtr child, IntPtr unused) {
                var name = new StringBuilder(160);
                GetClassName(child, name, name.Capacity);
                string cls = name.ToString();
                if (classes.Count < 32 && !classes.Contains(cls)) classes.Add(cls);
                return true;
            }, IntPtr.Zero);
            report["childWindowClasses"] = classes;
            report["hostDpiAwareness"] = GetAwarenessFromDpiAwarenessContext(GetWindowDpiAwarenessContext(form.Handle));
            string manualFingerprint = null;
            if (manual)
            {
                if (!nativeChild) throw new InvalidOperationException("Manual document is not hosted");
                manualFingerprint = await ManualWordSession.RunAsync(form, run, native, docObject);
            }
            else
            {
            run.Stage("write-marker");
            WriteMarker(run.Application, docObject, run.Marker);
            report["markerWritten"] = true;
            run.Stage("native-save");
            // CreateNew has no source. Open must save back to its existing source.
            Check(native.Save(operation == "create" ? run.Document : null, true, null, null), "saveHresult");
            report["savedByControl"] = true;
            for (int tick = 1; tick <= 2; tick++)
            {
                await Task.Delay(1000); // bounded periodic work on the same STA and document
                run.Stage("periodic-save-" + tick);
                WriteMarker(run.Application, docObject, run.Marker + " periodic-" + tick);
                Check(native.Save(null, true, null, null), "periodicSaveHresult" + tick);
                if (OieProbeDocumentWindow() != docWindow) throw new InvalidOperationException("Document HWND changed during save");
            }
            report["periodicSaveRequests"] = 2;
            report["sameDocumentThroughout"] = true;
            }
            run.Stage("native-close");
            Check(native.Close(), "closeHresult");
            closed = true;
            report["controlClosed"] = true;
            if (run.Application == "excel") CloseExcelWorkbook(run, appObject);
            report["documentClosed"] = true;
            ProbeRun.Release(docObject); docObject = null;
            if (framer != null) { framer.Dispose(); framer = null; }
            ProbeRun.Release(control); control = null;
            // Excel can stay alive until Trident releases its remaining OLE site.
            // The supervisor verifies the exact process identity after this host exits.
            cleanup = run.QuitIfOwnedAndEmpty(appObject, appPid, run.Application == "excel");
            run.Stage("verify-saved-file");
            report["documentVerified"] = manual ? run.SavedTextFingerprint() == manualFingerprint : run.VerifySavedMarker(run.Marker + " periodic-2");
            if (!(bool)report["documentVerified"] || !cleanup) throw new InvalidOperationException("Saved marker or cleanup failed");
            if (!nativeChild || !visible) {
                report["failureReason"] = "NativeOfficeWindowNotVisibleInHost";
                throw new InvalidOperationException("No visible Office document window in the host");
            }
        }
        catch (Exception e)
        {
            if (run != null) run.Error(Convert.ToString(report.ContainsKey("stage") ? report["stage"] : "preflight"), e);
            else { report["errorType"] = e.GetType().FullName; report["errorHresult"] = "0x" + Marshal.GetHRForException(e).ToString("X8"); }
            if (loaded && !closed && native != null)
            {
                // Recover this container's synthetic data before closing. This
                // remains a failed test even if the recovery save succeeds.
                try
                {
                    if (opened && (manual || report.ContainsKey("markerWritten")))
                        Check(native.Save(operation == "create" ? run.Document : null, true, null, null), "recoverySaveHresult");
                    Check(native.Close(), "recoveryCloseHresult");
                    closed = true;
                    report["failedTestRecoveredSafely"] = true;
                }
                catch (Exception recovery) { report["recoveryErrorType"] = recovery.GetType().FullName; }
            }
        }
        finally
        {
            var trace = new List<object>();
            try
            {
                int stage, hr;
                for (int i = 0; i < 48 && OieProbeTrace(i, out stage, out hr) != 0; i++)
                    trace.Add(new Dictionary<string, object> { { "stage", stage }, { "hresult", "0x" + hr.ToString("X8") } });
            }
            catch (EntryPointNotFoundException) { report["nativeTraceUnavailable"] = true; }
            catch (DllNotFoundException) { report["nativeTraceUnavailable"] = true; }
            report["nativeTrace"] = trace;
            // Do not call Quit from failure cleanup. A failed load has no proven
            // application identity. A failed save must not be silently discarded.
            report["cleanupNeedsReview"] = loaded && !closed;
            if (framer != null && (!loaded || closed)) framer.Dispose();
            if (loaded && !closed)
            {
                // Keep the owned container alive for recovery instead of closing
                // its HWND under an unsaved document. External runner will report
                // the retained process rather than killing it.
                bool recovering = false;
                form.FormClosing += delegate(object sender, FormClosingEventArgs e) {
                    e.Cancel = true;
                    if (recovering) return;
                    recovering = true;
                    try {
                        Check(native.Save(operation == "create" ? run.Document : null, true, null, null), "recoverySaveHresult");
                        Check(native.Close(), "recoveryCloseHresult");
                        e.Cancel = false;
                    }
                    catch (Exception recovery) { report["recoveryErrorType"] = recovery.GetType().FullName; }
                    finally { recovering = false; }
                };
                form.Text = "Office probe - save failed; recovery required";
            }
            if (!loaded || closed)
            {
                ProbeRun.Release(docObject);
                ProbeRun.Release(appObject); ProbeRun.Release(control);
            }
            if (run != null) run.Finish(opened && closed && cleanup);
        }
    }

    private static void CloseExcelWorkbook(ProbeRun run, object application)
    {
        if (application == null) throw new InvalidOperationException("Excel application was not bound");
        run.Stage("close-excel-workbook");
        object workbooks = ((dynamic)application).Workbooks;
        try
        {
            int count = Convert.ToInt32(((dynamic)workbooks).Count);
            for (int i = 1; i <= count; i++)
            {
                object workbook = ((dynamic)workbooks).Item(i);
                try
                {
                    dynamic book = workbook;
                    if (!String.Equals(Convert.ToString(book.FullName), run.Document, StringComparison.OrdinalIgnoreCase)) continue;
                    if (!Convert.ToBoolean(book.Saved)) throw new InvalidOperationException("Excel retained a dirty synthetic workbook");
                    book.Close(false);
                    run.Report["excelWorkbookClosedAfterControl"] = true;
                    return;
                }
                finally { ProbeRun.Release(workbook); }
            }
            run.Report["excelWorkbookClosedAfterControl"] = false;
        }
        finally { ProbeRun.Release(workbooks); }
    }

    private static int CountDocumentHits(IntPtr window)
    {
        Rect rect;
        if (!GetWindowRect(window, out rect) || rect.Right - rect.Left < 32 || rect.Bottom - rect.Top < 32) return 0;
        int count = 0;
        foreach (int x in new[] { 1, 2, 3 })
        {
            var point = new System.Drawing.Point(rect.Left + (rect.Right - rect.Left) * x / 4, rect.Top + (rect.Bottom - rect.Top) / 2);
            IntPtr hit = WindowFromPoint(point);
            if (hit == window || IsChild(window, hit)) count++;
        }
        return count;
    }

    private static void WriteMarker(string application, object document, string marker)
    {
        dynamic doc = document;
        object collection = null, item = null, target = null, frame = null, range = null;
        try
        {
            if (application == "excel")
            {
                collection = doc.Worksheets;
                item = ((dynamic)collection).Item(1);
                range = ((dynamic)item).Range["A4"];
                ((dynamic)range).Value2 = marker;
            }
            else if (application == "powerpoint")
            {
                collection = doc.Slides;
                item = ((dynamic)collection).Item(1);
                target = ((dynamic)item).Shapes;
                frame = ((dynamic)target).Item(1);
                object textFrame = ((dynamic)frame).TextFrame;
                try {
                    range = ((dynamic)textFrame).TextRange;
                    ((dynamic)range).Text = marker;
                } finally { ProbeRun.Release(textFrame); }
            }
            else
            {
                range = doc.Content;
                ((dynamic)range).Text = marker;
            }
        }
        finally
        {
            ProbeRun.Release(range); ProbeRun.Release(frame); ProbeRun.Release(target);
            ProbeRun.Release(item); ProbeRun.Release(collection);
        }
    }

    private void Check(int hr, string name)
    {
        report[name] = "0x" + hr.ToString("X8");
        if (hr < 0) Marshal.ThrowExceptionForHR(hr);
    }
    [DllImport("oie-legacy-framer.ocx", CallingConvention=CallingConvention.StdCall)]
    private static extern int OieProbeTrace(int index, out int stage, out int hr);
    [DllImport("oie-legacy-framer.ocx", CallingConvention=CallingConvention.StdCall)]
    private static extern IntPtr OieProbeDocumentWindow();
    private delegate bool WindowCallback(IntPtr child, IntPtr unused);
    [StructLayout(LayoutKind.Sequential)] private struct Rect { internal int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
    [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(System.Drawing.Point point);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, WindowCallback callback, IntPtr unused);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr child, StringBuilder name, int count);
    [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern IntPtr GetWindowDpiAwarenessContext(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern int GetAwarenessFromDpiAwarenessContext(IntPtr context);
}
