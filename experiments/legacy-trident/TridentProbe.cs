using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;

// Development probe: G0 observes synthetic HTML/empty controls. Explicit G1
// modes open only an owned synthetic fixture. Never registers COM.
internal static class TridentProbe
{
    internal const string ControlId = "70B8EE72-2B79-4D45-9A77-AE395CFEA071";
    private const string OriginalControlId = "00460182-9E5E-11D5-B7C8-B8269041DD57";
    [STAThread]
    private static int Main(string[] args)
    {
        string mode = "html", output = null, docPath = null, ownRoot = null, operation = "open";
        bool htmlObject = false, manual = false, loopback = false;
        foreach (string arg in args)
        {
            if (arg == "--dso-object") mode = "dso-object";
            else if (arg == "--dso-native") mode = "dso-native";
            else if (arg.StartsWith("--dso-doc=")) { mode = "dso-doc"; docPath = arg.Substring(10); }
            else if (arg == "--create-new") operation = "create";
            else if (arg == "--html-object") htmlObject = true;
            else if (arg == "--loopback") loopback = true;
            else if (arg == "--manual") manual = true;
            else if (arg.StartsWith("--own-root=")) ownRoot = arg.Substring(11);
            else if (arg.StartsWith("--output=")) output = Path.GetFullPath(arg.Substring(9));
            else { Console.Error.WriteLine("Unknown argument"); return 2; }
        }
        if (manual && (mode != "dso-doc" || !htmlObject)) return 2;
        if (loopback && (mode != "dso-doc" || !htmlObject)) return 2;
        var report = new Dictionary<string, object>();
        report["mode"] = mode;
        report["bitness"] = IntPtr.Size * 8;
        report["officeStartedByProbe"] = false;
        report["documentOpened"] = false;
        report["ideEmbeddingVerified"] = false;
        report["policyChanged"] = false;
        report["globalRegistrationRequested"] = false;
        try
        {
            // A private CLSID separates this test from installed controls; it must
            // never evade a kill bit applied to the original control identity.
            if (mode == "html" || KillBitPreflight(report))
            {
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
                Application.EnableVisualStyles();
                using (var form = new ProbeForm(mode, report, docPath, ownRoot, operation, htmlObject, manual, loopback)) Application.Run(form);
            }
        }
        catch (Exception error)
        {
            report["errorType"] = error.GetType().FullName;
            report["errorHresult"] = "0x" + Marshal.GetHRForException(error).ToString("X8");
        }
        string json = new JavaScriptSerializer().Serialize(report);
        if (mode != "dso-doc")
        {
            bool html = report.ContainsKey("tridentHtmlLoaded") && (bool)report["tridentHtmlLoaded"];
            report["passed"] = html && (mode == "html" || (report.ContainsKey("framerNativeWindowPresent") && (bool)report["framerNativeWindowPresent"])) && !report.ContainsKey("errorType") && !report.ContainsKey("timedOut") && !report.ContainsKey("activationBlocked");
            json = new JavaScriptSerializer().Serialize(report);
            if (output != null) { ProbeRun.CheckTree(output); using (var writer = new StreamWriter(new FileStream(output, FileMode.CreateNew))) writer.Write(json); }
        }
        Console.WriteLine(new JavaScriptSerializer().Serialize(report));
        return report.ContainsKey("activationBlocked") ? 3 : report.ContainsKey("timedOut") ? 4 : report.ContainsKey("passed") && (bool)report["passed"] ? 0 : 1;
    }

    private static bool KillBitPreflight(Dictionary<string, object> report)
    {
        var blocked = new List<string>();
        var views = Environment.Is64BitOperatingSystem ? new[] { RegistryView.Registry32, RegistryView.Registry64 } : new[] { RegistryView.Registry32 };
        var paths = new[] {
            @"SOFTWARE\Microsoft\Internet Explorer\ActiveX Compatibility\",
            @"SOFTWARE\Microsoft\Office\Common\COM Compatibility\",
            @"SOFTWARE\Microsoft\Office\12.0\Common\COM Compatibility\",
            @"SOFTWARE\Microsoft\Office\14.0\Common\COM Compatibility\",
            @"SOFTWARE\Microsoft\Office\15.0\Common\COM Compatibility\",
            @"SOFTWARE\Microsoft\Office\16.0\Common\COM Compatibility\"
        };
        report["policyPreflightScope"] = "Read-only original/private ActiveX and Office COM kill-bit keys; additional host policy remains enforced";
        try
        {
            foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
            foreach (var view in views)
            using (var root = RegistryKey.OpenBaseKey(hive, view))
            foreach (string path in paths)
            foreach (string id in new[] { OriginalControlId, ControlId })
            using (var key = root.OpenSubKey(path + "{" + id + "}", false))
            {
                if (key == null) continue;
                object raw = key.GetValue("Compatibility Flags", null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                if (raw != null && (Convert.ToInt64(raw) & 0x400) != 0)
                    blocked.Add(hive.ToString() + "/" + view.ToString() + "/" + path + "{" + id + "}");
            }
            report["killBitPreflightCompleted"] = true;
            report["blockingKillBitKeys"] = blocked;
            if (blocked.Count == 0) return true;
            report["activationBlocked"] = "Original or private control has an ActiveX/Office kill bit; no activation attempted";
        }
        catch (Exception error)
        {
            report["activationBlocked"] = "Policy preflight could not complete; no activation attempted";
            report["policyErrorType"] = error.GetType().FullName;
        }
        return false;
    }

    private sealed class EmptyFramer : AxHost
    {
        internal EmptyFramer() : base(ControlId) { }
        internal object Instance { get { return GetOcx(); } }
    }

    private sealed class ProbeForm : Form
    {
        private readonly string mode;
        private readonly Dictionary<string, object> report;
        private readonly string docPath;
        private readonly string ownRoot;
        private readonly string operation;
        private readonly bool htmlObject, manual;
        private LoopbackPage page;
        private readonly WebBrowser browser = new WebBrowser();
        private readonly Timer deadline = new Timer();
        private readonly Timer settle = new Timer();
        private readonly Timer docWork = new Timer();
        private EmptyFramer framer;
        protected override bool ShowWithoutActivation { get { return mode != "dso-doc"; } }

        internal ProbeForm(string mode, Dictionary<string, object> report, string docPath, string ownRoot, string operation, bool htmlObject, bool manual, bool loopback)
        {
            this.mode = mode; this.report = report; this.docPath = docPath; this.ownRoot = ownRoot; this.operation = operation;
            this.htmlObject = htmlObject; this.manual = manual;
            Text = "Office in Editor - synthetic Trident probe";
            Width = 820; Height = 560; ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual; Left = 100; Top = 100;
            if (manual) {
                Text = "Word 网页手测 — 合成文档"; ShowInTaskbar = true;
                Width = Math.Min(1100, Screen.PrimaryScreen.WorkingArea.Width - 80);
                Height = Math.Min(800, Screen.PrimaryScreen.WorkingArea.Height - 80);
                StartPosition = FormStartPosition.CenterScreen;
            }
            if (mode == "dso-doc")
            {
                // The external supervisor owns the deadline; a UI timer cannot
                // interrupt a synchronous COM call. No nested DoEvents loop.
                docWork.Interval = 800;
                docWork.Tick += DocWorkTick;
                if (htmlObject)
                {
                    browser.Dock = DockStyle.Fill;
                    browser.AllowWebBrowserDrop = false;
                    browser.IsWebBrowserContextMenuEnabled = false;
                    browser.WebBrowserShortcutsEnabled = false;
                    browser.Navigating += delegate(object sender, WebBrowserNavigatingEventArgs e) {
                        if (e.Url == null || (e.Url.AbsoluteUri != "about:blank" && (page == null || e.Url != page.Url))) e.Cancel = true;
                    };
                    browser.NewWindow += delegate(object sender, System.ComponentModel.CancelEventArgs e) { e.Cancel = true; };
                    bool scheduled = false;
                    browser.DocumentCompleted += delegate(object sender, WebBrowserDocumentCompletedEventArgs e) {
                        if (!scheduled && (!loopback || (page != null && e.Url == page.Url))) { scheduled = true; docWork.Start(); }
                    };
                    Controls.Add(browser);
                    Shown += delegate {
                        string html = "<!doctype html><html style='height:100%'><head><meta http-equiv='X-UA-Compatible' content='IE=edge'></head>" +
                            "<body style='margin:0;height:100%'><object id='framer' classid='clsid:" + ControlId + "' style='width:100%;height:100%'></object></body></html>";
                        report["pageTransport"] = loopback ? "localhost" : "document-text";
                        if (loopback)
                        {
                            page = new LoopbackPage(html);
                            report["port"] = page.Url.Port;
                            var preflight = new ProbeRun(ownRoot, report, Path.GetExtension(docPath));
                            preflight.ValidateDocument(docPath);
                            preflight.Stage("wait-localhost-page");
                            browser.Navigate(page.Url);
                        }
                        else browser.DocumentText = html;
                    };
                }
                else Shown += delegate { docWork.Start(); };
                return;
            }
            browser.Dock = DockStyle.Fill;
            browser.AllowWebBrowserDrop = false;
            browser.IsWebBrowserContextMenuEnabled = false;
            browser.WebBrowserShortcutsEnabled = false;
            browser.ScriptErrorsSuppressed = false;
            browser.Navigating += delegate(object sender, WebBrowserNavigatingEventArgs e)
            {
                if (e.Url == null || e.Url.AbsoluteUri != "about:blank")
                { e.Cancel = true; report["externalNavigationBlocked"] = true; }
            };
            browser.NewWindow += delegate(object sender, System.ComponentModel.CancelEventArgs e) { e.Cancel = true; };
            browser.DocumentCompleted += delegate
            {
                if (settle.Enabled) return;
                settle.Interval = 600; settle.Start();
            };
            settle.Tick += delegate
            {
                settle.Stop();
                try { Observe(); }
                catch (Exception error)
                {
                    report["errorType"] = error.GetType().FullName;
                    report["errorHresult"] = "0x" + Marshal.GetHRForException(error).ToString("X8");
                }
                Close();
            };
            deadline.Interval = 10000;
            deadline.Tick += delegate { deadline.Stop(); report["timedOut"] = true; Close(); };
            Controls.Add(browser);
            Shown += delegate
            {
                deadline.Start();
                string body = "<div id='synthetic' style='height:300px;background:#f4a340'>Synthetic HTML content</div>";
                if (mode == "dso-object") body = "<object id='framer' classid='clsid:" + ControlId + "' width='650' height='350'><p>ActiveX fallback</p></object>";
                browser.DocumentText = "<!doctype html><html><head><meta http-equiv='X-UA-Compatible' content='IE=edge'><title>Synthetic Trident probe</title>" +
                    "<script>function probeData(){return [document.documentMode||0,typeof ActiveXObject,!!document.getElementById('framer')].join('|');}" +
                    "function overlayProbe(){var r=document.getElementById('html-overlay-marker').getBoundingClientRect();var x=Math.round((r.left+r.right)/2),y=Math.round((r.top+r.bottom)/2);var e=document.elementFromPoint(x,y);return [x,y,e?e.id:'',e?e.tagName:'',document.documentElement.clientWidth,document.documentElement.clientHeight].join('|');}</script>" +
                    "</head><body style='font:18px sans-serif'><h1>Local synthetic probe</h1>" + body +
                    "<div id='html-overlay-marker' style='position:absolute;left:80px;top:180px;padding:25px;background:#68a8ed;z-index:99999'>HTML overlay marker</div></body></html>";
                if (mode == "dso-native")
                {
                    framer = new EmptyFramer(); framer.Left = 30; framer.Top = 90; framer.Width = 650; framer.Height = 350;
                    Controls.Add(framer); framer.BringToFront();
                }
            };
        }

        private void Observe()
        {
            report["tridentHtmlLoaded"] = browser.Document != null && browser.ReadyState == WebBrowserReadyState.Complete;
            report["engineVersion"] = browser.Version.ToString();
            if (browser.Document != null)
                report["scriptProbe"] = Convert.ToString(browser.Document.InvokeScript("probeData"));
            if (mode == "dso-native") report["nativeControlObjectAvailable"] = framer != null && framer.Instance != null;
            // DOM-script invocation of the framer is deliberately not attempted:
            // its source marks untrusted initialization safe but scripting unsafe.
            var classes = new List<string>();
            EnumChildWindows(Handle, delegate(IntPtr window, IntPtr unused)
            {
                var name = new StringBuilder(160); GetClassName(window, name, name.Capacity);
                string value = name.ToString(); if (!classes.Contains(value)) classes.Add(value);
                return true;
            }, IntPtr.Zero);
            report["childWindowClasses"] = classes;
            report["tridentNativeWindowPresent"] = classes.Contains("Internet Explorer_Server");
            report["framerNativeWindowPresent"] = classes.Contains("DSOFramerOCXWnd");
            ObserveOverlay();
        }

        private void ObserveOverlay()
        {
            try
            {
                string raw = Convert.ToString(browser.Document.InvokeScript("overlayProbe"));
                report["overlayDomProbe"] = raw;
                string[] parts = raw.Split('|');
                int viewportWidth = Int32.Parse(parts[4]), viewportHeight = Int32.Parse(parts[5]);
                if (viewportWidth <= 0 || viewportHeight <= 0) throw new InvalidOperationException();
                int x = (int)Math.Round(Int32.Parse(parts[0]) * browser.ClientSize.Width / (double)viewportWidth);
                int y = (int)Math.Round(Int32.Parse(parts[1]) * browser.ClientSize.Height / (double)viewportHeight);
                Point screen = browser.PointToScreen(new Point(x, y));
                // No input injection. Inspect the OS hit only when it belongs to
                // this probe, and do not collect data about covering windows.
                IntPtr actual = WindowFromPoint(screen);
                bool own = actual == Handle || IsChild(Handle, actual);
                report["overlayNativeHitWithinProbe"] = own;
                if (own) report["overlayNativeHitClass"] = WindowClass(actual);
                else report["overlayNativeHitInconclusive"] = "Probe is covered or outside the visible desktop";
                // Own-subtree geometry is useful when an unrelated window covers
                // the probe, but is not proof of the actual desktop hit result.
                var route = new List<string>();
                IntPtr current = Handle;
                for (int depth = 0; depth < 24; depth++)
                {
                    route.Add(WindowClass(current));
                    Point local = screen; ScreenToClient(current, ref local);
                    IntPtr next = ChildWindowFromPointEx(current, local, 1); // skip invisible children
                    if (next == IntPtr.Zero || next == current) break;
                    current = next;
                }
                report["overlayOwnChildHitRoute"] = route;
                report["overlayPixelCompositionVerified"] = false;
            }
            catch (Exception error) { report["overlayObservationErrorType"] = error.GetType().FullName; }
        }
        private static string WindowClass(IntPtr window)
        {
            var value = new StringBuilder(160); GetClassName(window, value, value.Capacity); return value.ToString();
        }

        private async void DocWorkTick(object sender, EventArgs e)
        {
            docWork.Stop();
            Func<object> factory = null;
            if (htmlObject) factory = delegate {
                report["htmlObjectLoaded"] = browser.ReadyState == WebBrowserReadyState.Complete;
                if (page != null)
                {
                    report["localhostPageServed"] = page.Served;
                    if (!page.Served || browser.Url != page.Url) throw new InvalidOperationException("Loopback page identity not verified");
                }
                browser.Focus();
                report["controlAccess"] = "native-ole-container";
                return DsoDocumentProbe.GetHostedControl(browser.Document.DomDocument);
            };
            await new DsoDocumentProbe(this, report, docPath, ownRoot, operation, factory, manual,
                htmlObject ? browser.Handle : IntPtr.Zero).ExecuteAsync();
            Close();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { deadline.Dispose(); settle.Dispose(); docWork.Dispose(); if (page != null) page.Dispose(); }
            base.Dispose(disposing);
        }
    }
    private delegate bool WindowEnum(IntPtr window, IntPtr param);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, WindowEnum callback, IntPtr param);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int max);
    [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(Point point);
    [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr window);
    [DllImport("user32.dll")] private static extern bool ScreenToClient(IntPtr window, ref Point point);
    [DllImport("user32.dll")] private static extern IntPtr ChildWindowFromPointEx(IntPtr parent, Point point, uint flags);
}
