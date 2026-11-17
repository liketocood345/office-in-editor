using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using OfficeInEditor.OfficeAddin;

public static class Program
{
    private static int assertions;
    [STAThread]
    public static int Main()
    {
        try
        {
            CheckRibbon(); CheckReadOnlyCapture(); CheckSecurity(); CheckWire(); CheckCancelledDispatch(); CheckCallback(); CheckEdits(); CheckWindowHost(); CheckLocalAutoSave(); CheckEmbeddedBindings(); CheckRouteCache();
            Console.WriteLine("Office add-in smoke passed: " + assertions + " assertions; no Office activation or registration.");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine("Smoke failed: " + error.GetType().Name + ": " + error.Message);
            if (error.InnerException != null)
                Console.Error.WriteLine("Inner: " + error.InnerException.GetType().Name + ": " + error.InnerException.Message);
            Console.Error.WriteLine(error.StackTrace);
            return 1;
        }
    }
    private static void Assert(bool ok, [System.Runtime.CompilerServices.CallerLineNumber] int line = 0)
    {
        assertions++;
        if (!ok) throw new Exception("Assertion failed at line " + line);
    }
    private static void CheckRibbon()
    {
        XNamespace ns = "http://schemas.microsoft.com/office/2009/07/customui";
        foreach (string app in new[] { "word", "excel", "powerpoint" })
        {
            var doc = XDocument.Parse(RibbonMarkup.Build(app));
            var ids = new HashSet<string>();
            foreach (var element in doc.Descendants())
            {
                var id = element.Attribute("id"); if (id != null) Assert(ids.Add(id.Value));
            }
            foreach (var menu in doc.Descendants(ns + "contextMenu"))
            {
                var button = menu.Element(ns + "button");
                Assert(button.Attribute("label").Value == "Add to Chat");
                Assert(button.Attribute("getEnabled").Value == "GetSendEnabled");
                Assert(button.Attribute("insertBeforeMso").Value == "Cut");
            }
        }
        Assert(OfficeAddin.DetectApplication("Microsoft Word") == "word");
        Assert(OfficeAddin.DetectApplication("Microsoft Excel") == "excel");
        Assert(OfficeAddin.DetectApplication("Microsoft PowerPoint") == "powerpoint");
        Assert(OfficeAddin.DetectApplication("Other app") == "unsupported");
        Assert(RibbonMarkup.FromRibbonId("Microsoft.Word.Document") == "word");
        Assert(RibbonMarkup.FromRibbonId("Microsoft.Excel.Workbook") == "excel");
        Assert(RibbonMarkup.FromRibbonId("Microsoft.PowerPoint.Presentation") == "powerpoint");
        Assert(RibbonMarkup.FromRibbonId("") == "unsupported");
        var word = XDocument.Parse(RibbonMarkup.Build("word"));
        var wordMenus = new List<string>();
        foreach (var menu in word.Descendants(ns + "contextMenu"))
            wordMenus.Add((string)menu.Attribute("idMso"));
        Assert(wordMenus.Count == 1 && wordMenus[0] == "ContextMenuText");
        Assert(RibbonMarkup.Build("word").Contains("OieTab"));
        Assert(RibbonMarkup.Build("word").Contains("OieSaveOnClose"));
        Assert(RibbonMarkup.Build("word").Contains("关标签保存"));
        Assert(RibbonMarkup.Build("word").Contains("OieAckAlerts"));
        Assert(RibbonMarkup.Build("word").Contains("自动确认警告"));
        Assert(RibbonMarkup.Build("excel").Contains("OieSaveOnClose"));
        Assert(RibbonMarkup.Build("excel").Contains("OieAckAlerts"));
        Assert(!RibbonMarkup.Build("word").Contains("ContextMenuHeading"));
        var addin = new OfficeAddin();
        var fallback = addin.GetCustomUI(null);
        Assert(fallback.Contains("OieTab"));
        Assert(!fallback.Contains("ContextMenuHeading"));
        var byRibbonId = addin.GetCustomUI("Microsoft.Word.Document");
        Assert(byRibbonId.Contains("OieTab"));
        Assert(byRibbonId.Contains("ContextMenuText"));
        Assert(!byRibbonId.Contains("ContextMenuHeading"));
        Assert(byRibbonId.Contains("onAction=\"RefreshChats\""));
        Assert(byRibbonId.Contains("onAction=\"AddToChat\""));
        Assert(byRibbonId.Contains("idMso=\"FileSave\""));
        Assert(byRibbonId.Contains("onAction=\"RepurposeFileSave\""));
        Assert(!byRibbonId.Contains("AutoSave"));
        Assert(!byRibbonId.Contains("MCP App"));
        Assert(!byRibbonId.Contains("OieSaveOnEdit"));
        Assert(byRibbonId.Contains("invalidateContentOnDrop=\"true\""));
        Assert(addin.GetStatusLabel(null) == "");
        addin.ChooseChat("missing-chat-id");
        Assert(addin.GetStatusLabel(null) == "");
        addin.RefreshChats(null);
        addin.AddToChat(null);
    }
    private static void CheckReadOnlyCapture()
    {
        var app = new FakeWord();
        var capture = new SelectionCapture(app, "word", "smoke-session");
        var first = Json.Map(capture.Capture());
        Assert((string)first["text"] == "sample");
        Assert(first["sourcePath"] == null);
        Assert(first["autoSaveAvailable"] == null);
        Assert((string)first["autoSaveState"] == "unsupported");
        Assert((bool)first["dirty"]);
        string identity = (string)first["documentId"];
        app.Selection.Range.Text = "changed";
        Assert((string)first["text"] == "sample"); // immutable captured string, not live selection
        Assert((string)Json.Map(capture.Capture())["documentId"] == identity);
        Assert((string)Json.Map(capture.Capture(identity))["documentId"] == identity);
        var another = new FakeDocument(); app.Documents.Add(another); app.ActiveDocument = another; app.Selection.Range.Document = another;
        Expect("DOCUMENT_CHANGED", () => capture.Capture(identity));
        app.ActiveDocument = app.Documents.Item(1); app.Selection.Range.Document = app.ActiveDocument;
        app.Selection.Range.End = 0;
        Expect("NO_SELECTION", () => capture.Capture());
        app.Selection.Range.End = SelectionCapture.MaxSelectionBytes;
        app.Selection.Range.ThrowOnText = true;
        Expect("SELECTION_TOO_LARGE", () => capture.Capture());
        Assert(!app.ActiveDocument.Mutated);
    }
    private static void CheckSecurity()
    {
        var sid = WindowsIdentity.GetCurrent().User;
        var security = Wire.CurrentUserSecurity();
        Assert(security.AreAccessRulesProtected);
        int allows = 0;
        foreach (PipeAccessRule rule in security.GetAccessRules(true, true, typeof(SecurityIdentifier)))
            if (rule.AccessControlType == AccessControlType.Allow) { allows++; Assert(rule.IdentityReference.Equals(sid)); }
        Assert(allows == 1);
        Assert(Wire.CorePipeName().StartsWith("office-in-editor-v1-"));
    }
    private static void CheckWire()
    {
        var memory = new MemoryStream(Wire.Utf8.GetBytes("{\"text\":\"字\"}\n"));
        Assert(Wire.ReadLineAsync(memory, CancellationToken.None).GetAwaiter().GetResult() == "{\"text\":\"字\"}");
        var oversized = new MemoryStream(new byte[Wire.MaxLineBytes + 1]);
        Expect("REQUEST_TOO_LARGE", () => Wire.ReadLineAsync(oversized, CancellationToken.None).GetAwaiter().GetResult());
    }
    private static void CheckCancelledDispatch()
    {
        int calls = 0;
        using (var dispatcher = new OfficeDispatcher())
        using (var cancel = new CancellationTokenSource())
        {
            cancel.Cancel();
            for (int index = 0; index < 100; index++)
            {
                var task = Task.Run(() => dispatcher.InvokeAsync(() => ++calls, cancel.Token));
                try { task.GetAwaiter().GetResult(); } catch (OperationCanceledException) { }
                Assert(task.IsCanceled);
                Assert(dispatcher.PendingCount == 0);
            }
        }
        // Any queued native APC now targets the process-rooted no-op thunk.
        Application.DoEvents(); Thread.Sleep(1);
        Assert(calls == 0);
    }

    private static void CheckLocalAutoSave()
    {
        string fixture = Path.Combine(Path.GetTempPath(), "oie-autosave-" + Guid.NewGuid().ToString("N") + ".docx");
        File.WriteAllText(fixture, "Fake Office save policy fixture");
        try
        {
            var application = new AutoSaveApplication();
            var doc = application.ActiveDocument;
            doc.FullName = fixture;
            long now = 0; int reports = 0;
            using (var dispatcher = new OfficeDispatcher())
            using (var save = new LocalAutoSave(application, "word", dispatcher, _ => reports++, () => now, false))
            {
                save.Check(); Assert(doc.SaveCalls == 0);
                Assert(!save.Periodic); save.Configure(false, 60); save.Check();
                now = 2000; save.Check(); Assert(doc.SaveCalls == 0);
                doc.Saved = false; save.Configure(true, 5); save.Check();
                now += 4999; save.Check(); Assert(doc.SaveCalls == 0);
                now++; save.Check(); Assert(doc.SaveCalls == 1);
                save.Configure(false, 5); doc.Saved = false;
                now += 60000; save.Check(); Assert(doc.SaveCalls == 1);
                save.Configure(true, 60); save.Check();
                now += 60_000; save.Check(); Assert(doc.SaveCalls == 2);
                doc.ReadOnly = false; doc.Saved = false; application.FailCommand = true;
                save.Configure(true, 5); save.Check(); now += 5000; save.Check(); save.Check();
                Assert(doc.SaveCalls == 2 && reports == 1); // failed FileSave does not fall back to Save()
                application.FailCommand = false;
                save.Configure(true, 5); doc.Saved = false; save.Check(); now += 5000; save.Check();
                Assert(doc.SaveCalls == 3);
                doc.Saved = false; File.AppendAllText(fixture, " external change");
                now += 60_000; save.Check(); save.Check();
                Assert(doc.SaveCalls == 3); Assert(reports == 2);
                save.Configure(true, 60); doc.ReadOnly = true; doc.Saved = false; save.Check();
                now += 60_000; save.Check(); Assert(doc.SaveCalls == 3);
            }
            Assert(!LocalAutoSave.LocalPath("https://example.test/file.docx", "word"));
            Assert(!LocalAutoSave.LocalPath("Unsaved.docx", "word"));
            Assert(!LocalAutoSave.LocalPath(fixture, "excel"));
            foreach (string kind in new[] { "excel", "powerpoint" })
            {
                string other = Path.ChangeExtension(fixture, kind == "excel" ? ".xlsx" : ".pptx");
                File.WriteAllText(other, "Fake Office format-specific policy fixture");
                try
                {
                    var app = new AutoSaveApplication(); app.ActiveDocument.FullName = other;
                    long elapsed = 0;
                    using (var dispatcher = new OfficeDispatcher())
                    using (var save = new LocalAutoSave(app, kind, dispatcher, _ => { }, () => elapsed, false))
                    {
                        save.Configure(true, 60); save.Check();
                        elapsed = 60_000; save.Check(); Assert(app.ActiveDocument.SaveCalls == 1);
                    }
                }
                finally { File.Delete(other); }
            }
        }
        finally { File.Delete(fixture); }
    }

    private static void CheckCallback()
    {
        string name = "office-in-editor-smoke-" + Guid.NewGuid().ToString("N");
        int uiThread = Thread.CurrentThread.ManagedThreadId;
        int calls = 0;
        using (var dispatcher = new OfficeDispatcher())
        using (var listener = new OfficeCallbackServer(name, dispatcher, (method, parameters) =>
            {
                if (method != "captureSelection") throw new BridgeException("METHOD_NOT_FOUND");
                calls++; return Json.Object("thread", Thread.CurrentThread.ManagedThreadId);
            }))
        {
            listener.Start();
            var valid = Task.Run(() => Request(name, "captureSelection"));
            PumpUntil(valid);
            var result = valid.GetAwaiter().GetResult();
            Assert(Json.IsTrue(Json.Get(result, "ok")));
            Assert(Convert.ToInt32(Json.Get(Json.Get(result, "result"), "thread")) == uiThread);
            var rejected = Task.Run(() => Request(name, "saveDocument"));
            PumpUntil(rejected);
            Assert(!Json.IsTrue(Json.Get(rejected.Result, "ok")));
            Assert(Json.Text(Json.Get(rejected.Result, "error"), "code") == "METHOD_NOT_FOUND");
            Assert(calls == 1);
        }
        using (var dispatcher = new OfficeDispatcher())
        {
            int windowCalls = 0;
            using (var windows = new OfficeCallbackServer(name + "-window", dispatcher, (method, parameters) =>
                {
                    windowCalls++;
                    if (method == "describeWindow") return Json.Object("hwnd", "1", "processId", 1, "attached", false);
                    if (method == "attachWindow")
                    {
                        Assert(Convert.ToInt32(Json.Get(parameters, "width")) >= 32);
                        return Json.Object("attached", true, "hwnd", "1");
                    }
                    throw new BridgeException("METHOD_NOT_FOUND");
                }))
            {
                windows.Start();
                var described = Task.Run(() => Request(name + "-window", "describeWindow"));
                PumpUntil(described);
                Assert(Json.IsTrue(Json.Get(described.Result, "ok")));
                var attached = Task.Run(() => Request(name + "-window", "attachWindow", Json.Object("x", 0, "y", 0, "width", 640, "height", 480)));
                PumpUntil(attached);
                Assert(Json.Text(Json.Get(attached.Result, "error"), "code") == "METHOD_NOT_FOUND" || Json.IsTrue(Json.Get(attached.Result, "ok")));
            }
            Assert(windowCalls >= 1);
        }
    }
    private static void CheckWindowHost()
    {
        var host = new WindowHost();
        var app = new WindowedWord();
        var described = Json.Map(host.DescribeWindow(app));
        Assert(described["hwnd"] is string);
        Assert(!(bool)described["attached"]);
        Expect("INVALID_HOST_WINDOW", () => host.Attach(Json.Object("x", 0, "y", 0, "width", 1, "height", 1), app));
        Expect("INVALID_HOST_WINDOW", () => host.Attach(Json.Object("x", 0, "y", 0, "width", "wide", "height", 100), app));
        var detached = Json.Map(host.Detach(app));
        Assert(!(bool)detached["attached"]);
        host.RestoreAll();
    }
    private static object Request(string pipe, string method, object parameters = null)
    {
        using (var cancel = new CancellationTokenSource(5000))
        using (var client = new NamedPipeClientStream(".", pipe, PipeDirection.InOut, PipeOptions.Asynchronous))
        {
            client.Connect(3000);
            Wire.WriteLineAsync(client, Json.Object("v", 1, "id", "smoke", "method", method,
                "params", parameters ?? Json.Object()), cancel.Token).GetAwaiter().GetResult();
            return Json.Serializer().DeserializeObject(Wire.ReadLineAsync(client, cancel.Token).GetAwaiter().GetResult());
        }
    }
    private static void PumpUntil(Task task)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(7);
        while (!task.IsCompleted && DateTime.UtcNow < deadline) { Application.DoEvents(); Thread.Sleep(5); }
        if (!task.IsCompleted) throw new TimeoutException();
    }
    private static void CheckEdits()
    {
        var wordApp = new EditableWord();
        var wordDocs = new SelectionCapture(wordApp, "word", "word-smoke").Documents;
        var word = new DocumentOperations(wordApp, "word", wordDocs);
        var inspected = Json.Map(word.Dispatch("inspectDocument", Json.Object()));
        string id = (string)inspected["documentId"];
        var caps = Json.Map(inspected["capabilities"]);
        Assert(((string[])caps["writeKinds"])[0] == "word.insert_text");
        Assert(!(bool)caps["manualSave"]);
        Assert(!(bool)caps["nativeAutoSaveLocalRedirect"]);
        Expect("MANUAL_SAVE_UNAVAILABLE", () => word.Dispatch("saveDocument", Json.Object("documentId", id)));
        object insert = Json.Object("kind", "word.insert_text", "args", Json.Object("offset", 0, "text", "X"));
        var prepared = Json.Map(word.Dispatch("prepareEdit", Json.Object("documentId", id, "operation", insert)));
        Assert(wordApp.ActiveDocument.Body == "Hello");
        var applied = Json.Map(word.Dispatch("applyEdit", Json.Object("documentId", id, "operation", insert,
            "prepared", prepared["prepared"], "recovery", prepared["recovery"], "beforeHash", prepared["beforeHash"])));
        Assert(wordApp.ActiveDocument.Body == "XHello");
        word.Dispatch("revertEdit", Json.Object("documentId", id, "recovery", prepared["recovery"],
            "afterHash", applied["afterHash"], "result", applied["result"]));
        Assert(wordApp.ActiveDocument.Body == "Hello");
        object again = Json.Object("kind", "word.insert_text", "args", Json.Object("offset", 0, "text", "Y"));
        var second = Json.Map(word.Dispatch("prepareEdit", Json.Object("documentId", id, "operation", again)));
        wordApp.ActiveDocument.Body += "!";
        Expect("DOCUMENT_CONFLICT", () => word.Dispatch("applyEdit", Json.Object("documentId", id, "operation", again,
            "prepared", second["prepared"], "recovery", second["recovery"], "beforeHash", second["beforeHash"])));
        Assert(wordApp.ActiveDocument.Body == "Hello!");
        wordApp.ActiveDocument.FullName = Path.Combine(Path.GetPathRoot(Path.GetTempPath()), "oie-smoke.docx");
        word.Dispatch("saveDocument", Json.Object("documentId", id));
        Assert(wordApp.ActiveDocument.SaveCalls == 1);
        wordApp.ActiveDocument.Protection = 3;
        var protectedInfo = Json.Map(word.Dispatch("inspectDocument", Json.Object("documentId", id)));
        Assert(((string[])Json.Get(protectedInfo["capabilities"], "writeKinds")).Length == 0);
        Expect("DOCUMENT_PROTECTED", () => word.Dispatch("prepareEdit", Json.Object("documentId", id, "operation", insert)));
        wordApp.ActiveDocument.UnknownProtection = true;
        Expect("DOCUMENT_PROTECTION_UNKNOWN", () => word.Dispatch("prepareEdit", Json.Object("documentId", id, "operation", insert)));
        wordApp.ActiveDocument.UnknownProtection = false; wordApp.ActiveDocument.Protection = -1;

        var excelApp = new EditableExcel();
        var excel = new DocumentOperations(excelApp, "excel", new SelectionCapture(excelApp, "excel", "excel-smoke").Documents);
        var excelInfo = Json.Map(excel.Dispatch("inspectDocument", Json.Object()));
        string workbookId = (string)excelInfo["documentId"];
        object values = Json.Object("kind", "excel.set_range_values", "args", Json.Object("sheet", "Sheet1", "address", "A1", "values", new object[] { new object[] { 42 } }));
        var excelPrepared = Json.Map(excel.Dispatch("prepareEdit", Json.Object("documentId", workbookId, "operation", values)));
        var excelApplied = Json.Map(excel.Dispatch("applyEdit", Json.Object("documentId", workbookId, "operation", values,
            "prepared", excelPrepared["prepared"], "recovery", excelPrepared["recovery"], "beforeHash", excelPrepared["beforeHash"])));
        Assert(Equals(excelApp.ActiveWorkbook.Sheet.Cell.Value2, 42d) || Equals(excelApp.ActiveWorkbook.Sheet.Cell.Value2, 42));
        excel.Dispatch("revertEdit", Json.Object("documentId", workbookId, "recovery", excelPrepared["recovery"],
            "afterHash", excelApplied["afterHash"], "result", excelApplied["result"]));
        Assert(excelApp.ActiveWorkbook.Sheet.Cell.Value2 == null);

        var pptApp = new EditablePowerPoint();
        var ppt = new DocumentOperations(pptApp, "powerpoint", new SelectionCapture(pptApp, "powerpoint", "ppt-smoke").Documents);
        var pptInfo = Json.Map(ppt.Dispatch("inspectDocument", Json.Object()));
        string presentationId = (string)pptInfo["documentId"];
        object addSlide = Json.Object("kind", "powerpoint.add_slide", "args", Json.Object());
        var slidePrepared = Json.Map(ppt.Dispatch("prepareEdit", Json.Object("documentId", presentationId, "operation", addSlide)));
        var slideApplied = Json.Map(ppt.Dispatch("applyEdit", Json.Object("documentId", presentationId, "operation", addSlide,
            "prepared", slidePrepared["prepared"], "recovery", slidePrepared["recovery"], "beforeHash", slidePrepared["beforeHash"])));
        Assert(pptApp.ActivePresentation.Slides.Count == 1);
        int slideId = Convert.ToInt32(Json.Get(Json.Get(slideApplied, "result"), "slideId"));
        object addBox = Json.Object("kind", "powerpoint.add_textbox", "args", Json.Object("slideId", slideId, "text", "Title", "left", 10, "top", 20, "width", 100, "height", 40));
        var boxPrepared = Json.Map(ppt.Dispatch("prepareEdit", Json.Object("documentId", presentationId, "operation", addBox)));
        var boxApplied = Json.Map(ppt.Dispatch("applyEdit", Json.Object("documentId", presentationId, "operation", addBox,
            "prepared", boxPrepared["prepared"], "recovery", boxPrepared["recovery"], "beforeHash", boxPrepared["beforeHash"])));
        Assert(pptApp.ActivePresentation.Slides.Only.Shapes.Count == 1);
        ppt.Dispatch("revertEdit", Json.Object("documentId", presentationId, "recovery", boxPrepared["recovery"],
            "afterHash", boxApplied["afterHash"], "result", boxApplied["result"]));
        Assert(pptApp.ActivePresentation.Slides.Only.Shapes.Count == 0);
        ppt.Dispatch("revertEdit", Json.Object("documentId", presentationId, "recovery", slidePrepared["recovery"],
            "afterHash", slideApplied["afterHash"], "result", slideApplied["result"]));
        Assert(pptApp.ActivePresentation.Slides.Count == 0);
    }
    private static void CheckEmbeddedBindings()
    {
        string firstPath = Path.Combine(Path.GetTempPath(), "oie-binding-" + Guid.NewGuid().ToString("N") + ".docx");
        string secondPath = Path.ChangeExtension(firstPath, ".second.docx");
        File.WriteAllText(firstPath, "first fake"); File.WriteAllText(secondPath, "second fake");
        try
        {
            var app = new EditableWord(); var first = app.ActiveDocument; var second = new EditableDocument();
            string firstSession = "ole-" + Guid.NewGuid().ToString("N"), secondSession = "ole-" + Guid.NewGuid().ToString("N");
            bool live = true;
            var embedded = new EmbeddedDocumentBindings(hwnd => hwnd == 1 ? first : second, _ => live);
            var registry = new DocumentRegistry(app, "word", "binding-smoke", embedded);
            embedded.Update(Json.Object("bindings", new[] {
                Json.Object("sessionId", firstSession, "sourcePath", firstPath, "hwnd", "1", "sourceVerified", true),
                Json.Object("sessionId", secondSession, "sourcePath", secondPath, "hwnd", "2", "sourceVerified", true) }));
            Assert(registry.OpenDocuments().Count == 2);
            string firstId = registry.Bind(first), secondId = registry.Bind(second);
            Assert(firstId != secondId && registry.Bind(first) == firstId);
            Assert(ReferenceEquals(registry.Resolve(secondId), second));
            Assert(Json.Text(registry.Describe(first), "sourcePath") == firstPath);
            Assert(Json.Text(registry.Describe(second), "sessionId") == secondSession);
            var operations = new DocumentOperations(app, "word", registry);
            var requested = operations.Dispatch("saveDocument", Json.Object("documentId", secondId));
            Assert(Json.IsTrue(Json.Get(requested, "requiresHostSave")) && !Json.IsTrue(Json.Get(requested, "saved")));
            Assert(Json.Text(requested, "sessionId") == secondSession && second.SaveCalls == 0 && first.SaveCalls == 0);
            using (var dispatcher = new OfficeDispatcher())
            using (var autosave = new LocalAutoSave(app, "word", dispatcher, _ => { }, () => 60000, false, registry))
            { autosave.Configure(true, 5); autosave.Check(); Assert(first.SaveCalls == 0); }
            live = false;
            Expect("DOCUMENT_CLOSED", () => registry.Resolve(secondId));
            Assert(embedded.Find(first) == null);
        }
        finally { File.Delete(firstPath); File.Delete(secondPath); }
    }
    private static void CheckRouteCache()
    {
        var cache = new ChatRouteCache();
        Assert(cache.Bind("doc-a", "session-a", "chat-a", "A", "epoch-a", 2, "broker-a"));
        Assert(cache.Bind("doc-b", "session-b", "chat-b", "B", "epoch-b", 1, "broker-a"));
        Assert(!cache.Bind("doc-a", "session-a", "stale", "stale", "old", 1, "broker-a"));
        Assert(cache.For("doc-a").Id == "chat-a" && cache.For("doc-b").Id == "chat-b");
        cache.For("doc-a").Manual = true; cache.For("doc-a").Id = null;
        Assert(cache.For("doc-a").Manual && cache.For("doc-b").Id == "chat-b");
        Assert(cache.Bind("doc-a", "session-a", "new", "New", "new-epoch", 3, "broker-a"));
        Assert(!cache.For("doc-a").Manual);
        Assert(cache.Bind("doc-a", "session-a", "restart", "Restart", "restart-epoch", 1, "broker-b"));
        Assert(cache.For("doc-b").Id == null);
        Assert(!cache.Bind("doc-a", "session-a", "late", "Late", "late-epoch", 99, "broker-a"));
        Assert(cache.For("doc-a").Id == "restart");
    }
    private static void Expect(string code, Action action)
    {
        try { action(); throw new Exception("Expected rejection"); }
        catch (BridgeException error) { Assert(error.Code == code); }
    }
}

// Test doubles only. No COM activation, Office process access or document writes.
public sealed class AutoSaveApplication
{
    public AutoSaveDocument ActiveDocument { get; } = new AutoSaveDocument();
    public AutoSaveDocument ActiveWorkbook => ActiveDocument;
    public AutoSaveDocument ActivePresentation => ActiveDocument;
    public bool FailCommand;
    public AutoSaveCommands CommandBars => new AutoSaveCommands(ActiveDocument, FailCommand);
}
public sealed class AutoSaveCommands
{
    private readonly AutoSaveDocument document;
    private readonly bool fail;
    public AutoSaveCommands(AutoSaveDocument document, bool fail) { this.document = document; this.fail = fail; }
    public void ExecuteMso(string id) { if (id != "FileSave" || fail) throw new InvalidOperationException(); document.Save(); }
}
public sealed class AutoSaveDocument
{
    public string FullName { get; set; }
    public bool Saved { get; set; }
    public bool ReadOnly { get; set; }
    public int ProtectionType => -1;
    public int SaveCalls { get; private set; }
    public void Save() { SaveCalls++; Saved = true; }
}
public sealed class WindowedWord
{
    public int Hwnd => 1;
}

public sealed class FakeWord
{
    public FakeCount Documents { get; }
    public FakeDocument ActiveDocument { get; set; } = new FakeDocument();
    public FakeSelection Selection { get; } = new FakeSelection();
    public FakeWord() { Documents = new FakeCount(ActiveDocument); Selection.Range.Document = ActiveDocument; }
}
public sealed class FakeCount
{
    private readonly List<FakeDocument> documents = new List<FakeDocument>();
    public FakeCount(FakeDocument document) { documents.Add(document); }
    public int Count => documents.Count;
    public void Add(FakeDocument document) { documents.Add(document); }
    public FakeDocument Item(int index) => documents[index - 1];
}
public sealed class FakeSelection { public FakeRange Range { get; } = new FakeRange(); }
public sealed class FakeRange
{
    private string text = "sample";
    public int Start => 0;
    public int End { get; set; } = 6;
    public int StoryType => 1;
    public bool ThrowOnText { get; set; }
    public object Document { get; set; }
    public string Text { get { if (ThrowOnText) throw new Exception(); return text; } set { text = value; } }
}
public sealed class FakeDocument
{
    public string Name => "Unsaved test document";
    public string FullName => "Unsaved test document";
    public bool Mutated { get; private set; }
    public bool Saved { get => false; set { Mutated = true; throw new Exception(); } }
    public bool AutoSaveOn { get => throw new NotSupportedException(); set { Mutated = true; throw new Exception(); } }
}

public sealed class EditableWord
{
    public EditableDocument ActiveDocument { get; } = new EditableDocument();
    public EditableDocuments Documents { get; }
    public EditableCommands CommandBars { get; }
    public EditableWord()
    {
        Documents = new EditableDocuments(ActiveDocument);
        CommandBars = new EditableCommands(ActiveDocument);
    }
}
public sealed class EditableCommands
{
    private readonly EditableDocument document;
    public EditableCommands(EditableDocument document) { this.document = document; }
    public void ExecuteMso(string id)
    {
        if (id != "FileSave") throw new InvalidOperationException();
        document.Save();
    }
}
public sealed class EditableDocuments
{
    private readonly EditableDocument document;
    public EditableDocuments(EditableDocument document) { this.document = document; }
    public int Count => 1;
    public EditableDocument Item(int index) => index == 1 ? document : throw new InvalidOperationException();
}
public sealed class EditableDocument
{
    public string Body = "Hello";
    public string Name => "Notes";
    public string FullName { get; set; } = "Unsaved notes";
    public bool Saved { get; set; } = true;
    public bool ReadOnly => false;
    public int Protection = -1;
    public bool UnknownProtection;
    public int ProtectionType => UnknownProtection ? throw new NotSupportedException() : Protection;
    public int SaveCalls { get; private set; }
    public EditableContent Content { get; }
    public bool AutoSaveOn { get => throw new NotSupportedException(); }
    public EditableDocument() { Content = new EditableContent(this); }
    public EditableStoryRange Range(int start, int end) => new EditableStoryRange(this, start, end);
    public void Save() { SaveCalls++; Saved = true; }
}
public sealed class EditableContent
{
    private readonly EditableDocument document;
    public EditableContent(EditableDocument document) { this.document = document; }
    public int Start => 0;
    public int End => document.Body.Length;
}
public sealed class EditableStoryRange
{
    private readonly EditableDocument document;
    public int Start { get; private set; }
    public int End { get; private set; }
    public EditableStoryRange(EditableDocument document, int start, int end)
    {
        if (start < 0 || end < start || end > document.Body.Length) throw new InvalidOperationException();
        this.document = document; Start = start; End = end;
    }
    public string Text
    {
        get => document.Body.Substring(Start, End - Start);
        set
        {
            string next = value ?? "";
            document.Body = document.Body.Remove(Start, End - Start).Insert(Start, next);
            End = Start + next.Length;
        }
    }
    public void InsertBefore(string value)
    {
        string next = value ?? "";
        document.Body = document.Body.Insert(Start, next);
        End += next.Length;
    }
}
public sealed class EditableExcel
{
    public EditableWorkbook ActiveWorkbook { get; } = new EditableWorkbook();
    public EditableWorkbooks Workbooks { get; }
    public EditableExcel() { Workbooks = new EditableWorkbooks(ActiveWorkbook); }
}
public sealed class EditableWorkbooks
{
    private readonly EditableWorkbook workbook;
    public EditableWorkbooks(EditableWorkbook workbook) { this.workbook = workbook; }
    public int Count => 1;
    public EditableWorkbook Item(int index) => index == 1 ? workbook : throw new InvalidOperationException();
}
public sealed class EditableWorkbook
{
    public EditableSheet Sheet { get; } = new EditableSheet();
    public string Name => "Book";
    public string FullName { get; set; } = "Unsaved book";
    public bool Saved { get; set; } = true;
    public bool ReadOnly => false;
    public int ProtectionType => -1;
    public bool AutoSaveOn { get => throw new NotSupportedException(); }
    public EditableSheets Worksheets { get; }
    public EditableSheet ActiveSheet => Sheet;
    public EditableWorkbook() { Worksheets = new EditableSheets(Sheet); }
    public void Save() { Saved = true; }
}
public sealed class EditableSheets
{
    private readonly EditableSheet sheet;
    public EditableSheets(EditableSheet sheet) { this.sheet = sheet; }
    public EditableSheet Item(string name) => name == sheet.Name ? sheet : throw new InvalidOperationException();
}
public sealed class EditableSheet
{
    public string Name => "Sheet1";
    public bool ProtectContents => false;
    public EditableCell Cell { get; } = new EditableCell();
    public EditableSheetRange Range(string address) => address == "A1" ? new EditableSheetRange(this) : throw new InvalidOperationException();
}
public sealed class EditableSheetRange
{
    private readonly EditableSheet sheet;
    public EditableSheetRange(EditableSheet sheet) { this.sheet = sheet; }
    public EditableSheet Worksheet => sheet;
    public EditableAxis Rows => new EditableAxis(1);
    public EditableAxis Columns => new EditableAxis(1);
    public EditableCells Cells => new EditableCells(sheet.Cell);
}
public sealed class EditableAxis { public int Count { get; } public EditableAxis(int count) { Count = count; } }
public sealed class EditableCells
{
    private readonly EditableCell cell;
    public EditableCells(EditableCell cell) { this.cell = cell; }
    public EditableCell Item(int row, int column) => row == 1 && column == 1 ? cell : throw new InvalidOperationException();
}
public sealed class EditableCell
{
    private string formula;
    public object Value2 { get; set; }
    public object Formula
    {
        get => formula ?? (Value2 == null ? "" : Convert.ToString(Value2, System.Globalization.CultureInfo.InvariantCulture));
        set
        {
            formula = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
            if (formula != null && formula.StartsWith("=", StringComparison.Ordinal)) Value2 = 0d;
            else { Value2 = formula; formula = null; }
        }
    }
}
public sealed class EditablePowerPoint
{
    public EditablePresentation ActivePresentation { get; } = new EditablePresentation();
    public EditablePresentations Presentations { get; }
    public EditablePowerPoint() { Presentations = new EditablePresentations(ActivePresentation); }
}
public sealed class EditablePresentations
{
    private readonly EditablePresentation presentation;
    public EditablePresentations(EditablePresentation presentation) { this.presentation = presentation; }
    public int Count => 1;
    public EditablePresentation Item(int index) => index == 1 ? presentation : throw new InvalidOperationException();
}
public sealed class EditablePresentation
{
    public EditableSlides Slides { get; } = new EditableSlides();
    public string Name => "Deck";
    public string FullName { get; set; } = "Unsaved deck";
    public bool Saved { get; set; } = true;
    public bool ReadOnly => false;
    public int ProtectionType => -1;
    public bool AutoSaveOn { get => throw new NotSupportedException(); }
    public void Save() { Saved = true; }
}
public sealed class EditableSlides
{
    private readonly List<EditableSlide> slides = new List<EditableSlide>();
    private int nextId = 256;
    public int Count => slides.Count;
    public EditableSlide Only => slides[0];
    public EditableSlide Item(int index) => slides[index - 1];
    public EditableSlide Add(int index, int layout)
    {
        if (index != slides.Count + 1 || layout != 12) throw new InvalidOperationException();
        var slide = new EditableSlide(this, nextId++);
        slides.Add(slide);
        return slide;
    }
    public void Remove(EditableSlide slide) => slides.Remove(slide);
}
public sealed class EditableSlide
{
    private readonly EditableSlides owner;
    public int SlideID { get; }
    public EditableShapes Shapes { get; } = new EditableShapes();
    public EditableSlide(EditableSlides owner, int id) { this.owner = owner; SlideID = id; }
    public void Delete() => owner.Remove(this);
}
public sealed class EditableShapes
{
    private readonly List<EditableShape> shapes = new List<EditableShape>();
    private int nextId = 1;
    public int Count => shapes.Count;
    public EditableShape Item(int index) => shapes[index - 1];
    public EditableShape AddTextbox(int orientation, double left, double top, double width, double height)
    {
        if (orientation != 1 || width < 1 || height < 1) throw new InvalidOperationException();
        var shape = new EditableShape(this, nextId++);
        shapes.Add(shape);
        return shape;
    }
    public void Remove(EditableShape shape) => shapes.Remove(shape);
}
public sealed class EditableShape
{
    private readonly EditableShapes owner;
    public int Id { get; }
    public EditableTextFrame TextFrame { get; } = new EditableTextFrame();
    public EditableShape(EditableShapes owner, int id) { this.owner = owner; Id = id; }
    public void Delete() => owner.Remove(this);
}
public sealed class EditableTextFrame { public EditableTextRange TextRange { get; } = new EditableTextRange(); }
public sealed class EditableTextRange { public string Text { get; set; } }
