using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using System.Xml;

// Shared by development probes only. No registry changes, process termination,
// global input, timers, background work or document discovery.
internal sealed class ProbeRun
{
    internal readonly string Root, Document, Marker;
    internal readonly string Application, ProcessName, ProgId;
    internal readonly Dictionary<string, object> Report;
    internal readonly Stopwatch Clock = Stopwatch.StartNew();
    private readonly List<object> stages = new List<object>();
    private readonly DateTime startedUtc = DateTime.UtcNow;
    private readonly HashSet<int> preexisting = new HashSet<int>();
    private readonly Dictionary<int, long> identities = new Dictionary<int, long>();
    private readonly JavaScriptSerializer json = new JavaScriptSerializer();

    internal ProbeRun(string root, Dictionary<string, object> report, string extension = ".docx")
    {
        if (String.IsNullOrWhiteSpace(root)) throw new ArgumentException("Run root required");
        Root = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar);
        CheckTree(Root);
        var owner = ReadJson(Path.Combine(Root, "ownership.json"));
        object project, id, purpose;
        if (!owner.TryGetValue("project", out project) || (string)project != "office-in-editor" ||
            !owner.TryGetValue("run", out id) || !System.Text.RegularExpressions.Regex.IsMatch((string)id, "^[a-f0-9]{32}$") ||
            !owner.TryGetValue("purpose", out purpose) ||
            ((string)purpose != "word-save-probe-p0" && (string)purpose != "trident-g1" &&
             (string)purpose != "word-browser-local"))
            throw new InvalidDataException("Invalid run ownership");
        Report = report;
        Marker = "OIE probe " + (string)id;
        switch (extension)
        {
            case ".docx": Application = "word"; ProcessName = "WINWORD"; ProgId = "Word.Document.12"; break;
            case ".xlsx": Application = "excel"; ProcessName = "EXCEL"; ProgId = "Excel.Sheet.12"; break;
            case ".pptx": Application = "powerpoint"; ProcessName = "POWERPNT"; ProgId = "PowerPoint.Show.12"; break;
            default: throw new InvalidDataException("Unsupported synthetic document format");
        }
        if (Application != "word" && (string)purpose != "trident-g1") throw new InvalidDataException("Unexpected fixture purpose");
        Document = Path.Combine(Root, (Application == "excel" ? "synthetic-" + (string)id : "synthetic") + extension);
        foreach (var p in Process.GetProcessesByName(ProcessName))
            using (p) preexisting.Add(p.Id);
        report["run"] = (string)id;
        report["application"] = Application;
        report["preexistingOfficeCount"] = preexisting.Count;
        if (Application == "word") report["preexistingWordCount"] = preexisting.Count;
        report["officeStartedByProbe"] = false;
        report["passed"] = false;
        report["stages"] = stages;
        report["executableSha256"] = Hash(Process.GetCurrentProcess().MainModule.FileName);
    }

    internal static void CheckTree(string path)
    {
        string full = Path.GetFullPath(path);
        if (full.StartsWith(@"\\", StringComparison.Ordinal)) throw new InvalidDataException("UNC run paths refused");
        for (string current = full; !String.IsNullOrEmpty(current); current = Path.GetDirectoryName(current))
            if ((File.Exists(current) || Directory.Exists(current)) &&
                (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                throw new InvalidDataException("Reparse points refused");
    }

    internal void ValidateDocument(string candidate)
    {
        if (!String.Equals(Path.GetFullPath(candidate), Document, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Only the run fixture is accepted");
        CheckTree(Document);
        if (Application != "word")
        {
            var fixtures = ReadJson(Path.Combine(Root, "fixtures.json"));
            if (!String.Equals(Convert.ToString(fixtures[Path.GetFileName(Document)]), Hash(Document), StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Fixture identity or digest mismatch");
            return;
        }
        var manifest = ReadJson(Path.Combine(Root, "synthetic.json"));
        if (Convert.ToString(manifest["marker"]) != Marker ||
            Convert.ToString(manifest["file"]) != "synthetic.docx" ||
            Convert.ToString(manifest["sha256"]) != Hash(Document))
            throw new InvalidDataException("Fixture identity or digest mismatch");
        if (!ContainsMarker("OIE synthetic seed")) throw new InvalidDataException("Not a synthetic seed");
    }

    internal void Prepare()
    {
        if (File.Exists(Document) || File.Exists(Path.Combine(Root, "synthetic.json")))
            throw new IOException("Fresh run required");
        using (var stream = new FileStream(Document, FileMode.CreateNew, FileAccess.ReadWrite))
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create))
        {
            Entry(zip, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
            Entry(zip, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
            Entry(zip, "word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>OIE synthetic seed</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
        }
        WriteNew(Path.Combine(Root, "synthetic.json"), new Dictionary<string, object> {
            { "file", "synthetic.docx" }, { "marker", Marker }, { "sha256", Hash(Document) } });
    }

    private static void Entry(ZipArchive zip, string name, string text)
    {
        using (var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false))) writer.Write(text);
    }

    internal bool ContainsMarker(string marker)
    {
        using (var zip = ZipFile.OpenRead(Document))
        {
            var entry = zip.GetEntry("word/document.xml");
            if (entry == null || entry.Length > 1048576) return false;
            using (var reader = new StreamReader(entry.Open())) return reader.ReadToEnd().Contains(marker);
        }
    }

    internal bool VerifySavedMarker(string marker)
    {
        if (Application == "word") return ContainsMarker(marker);
        using (var zip = ZipFile.OpenRead(Document))
        {
            if (Application == "excel")
            {
                var sheet = ReadPart(zip, "xl/worksheets/sheet1.xml");
                var ns = new XmlNamespaceManager(sheet.NameTable);
                ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var cell = sheet.SelectSingleNode("//s:c[@r='A4']", ns);
                if (cell == null) return false;
                if (cell.Attributes["t"] != null && cell.Attributes["t"].Value == "s")
                {
                    int index;
                    var value = cell.SelectSingleNode("s:v", ns);
                    if (value == null || !Int32.TryParse(value.InnerText, out index) || index < 0) return false;
                    var strings = ReadPart(zip, "xl/sharedStrings.xml").DocumentElement;
                    return index < strings.ChildNodes.Count && strings.ChildNodes[index].InnerText == marker;
                }
                return cell.InnerText == marker;
            }
            var slide = ReadPart(zip, "ppt/slides/slide1.xml");
            var slideNs = new XmlNamespaceManager(slide.NameTable);
            slideNs.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            slideNs.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            var combined = new StringBuilder();
            foreach (XmlNode run in slide.SelectNodes("/p:sld/p:cSld/p:spTree/p:sp[1]/p:txBody//a:t", slideNs)) combined.Append(run.InnerText);
            return combined.ToString() == marker;
        }
    }

    private static XmlDocument ReadPart(ZipArchive zip, string name)
    {
        var entry = zip.GetEntry(name);
        if (entry == null || entry.Length > 1048576) throw new InvalidDataException("Invalid synthetic document XML");
        var xml = new XmlDocument { XmlResolver = null };
        using (var stream = entry.Open())
        using (var reader = XmlReader.Create(stream, new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 1048576 })) xml.Load(reader);
        return xml;
    }

    internal static string TextFingerprint(string text)
    {
        var compact = new StringBuilder();
        foreach (char value in text) if (!Char.IsWhiteSpace(value) && !Char.IsControl(value)) compact.Append(value);
        using (var hash = SHA256.Create()) return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(compact.ToString()))).Replace("-", "");
    }
    internal string SavedTextFingerprint()
    {
        using (var zip = ZipFile.OpenRead(Document))
        {
            var entry = zip.GetEntry("word/document.xml");
            if (entry == null || entry.Length > 1048576) throw new InvalidDataException("Invalid document XML");
            var xml = new System.Xml.XmlDocument { XmlResolver = null };
            using (var stream = entry.Open())
            using (var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 1048576 })) xml.Load(reader);
            var text = new StringBuilder();
            foreach (System.Xml.XmlNode node in xml.GetElementsByTagName("t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")) text.Append(node.InnerText);
            return TextFingerprint(text.ToString());
        }
    }

    internal void Stage(string name)
    {
        if (stages.Count >= 32) throw new InvalidOperationException("Stage budget exhausted");
        Report["stage"] = name;
        stages.Add(new Dictionary<string, object> { { "name", name }, { "atMs", Clock.ElapsedMilliseconds } });
        // At most one small overwrite per COM phase, never a polling log.
        CheckTree(Path.Combine(Root, "checkpoint.json"));
        File.WriteAllText(Path.Combine(Root, "checkpoint.json"), json.Serialize(Report), new UTF8Encoding(false));
    }

    internal int BindApplication(object application, object document)
    {
        // Word exposes HWND on Window, unlike Excel's Application.Hwnd.
        // Use this synthetic document's own Windows collection, never the
        // globally active window or title matching.
        object windows = ((dynamic)document).Windows;
        object window = null;
        IntPtr hwnd;
        try
        {
            window = ((dynamic)windows).Item(1);
            hwnd = new IntPtr(unchecked((long)(uint)Convert.ToInt32(((dynamic)window).Hwnd)));
        }
        finally { Release(window); Release(windows); }
        return BindNativeWindow(hwnd);
    }

    internal int BindNativeWindow(IntPtr hwnd)
    {
        uint pid;
        if (GetWindowThreadProcessId(hwnd, out pid) == 0 || pid == 0)
            throw new InvalidOperationException("Office HWND unavailable");
        using (var process = Process.GetProcessById((int)pid))
        {
            if (!String.Equals(process.ProcessName, ProcessName, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Unexpected HWND process");
            long ticks = process.StartTime.ToUniversalTime().Ticks;
            bool own = !preexisting.Contains((int)pid) && ticks >= startedUtc.AddSeconds(-1).Ticks;
            if (own) identities[(int)pid] = ticks;
            Report["applicationIsOwned"] = own;
            Report["officeStartedByProbe"] = own;
            Report["applicationPid"] = (int)pid;
            Report["applicationStartUtcTicks"] = ticks;
            Report["documentWindowHwnd"] = hwnd.ToInt64();
            Report["documentWindowDpiAwareness"] = GetAwarenessFromDpiAwarenessContext(GetWindowDpiAwarenessContext(hwnd));
            return (int)pid;
        }
    }

    // Must only be called after this probe has closed its synthetic document.
    // Shared instances are never asked to quit. PID identity is checked again.
    internal bool QuitIfOwnedAndEmpty(object application, int pid, bool verifyExitAfterHostCloses = false)
    {
        long ticks;
        if (!identities.TryGetValue(pid, out ticks)) { Report["applicationQuitSkippedShared"] = true; return true; }
        Process owned;
        try { owned = Process.GetProcessById(pid); }
        catch (ArgumentException) { Report["ownedProcessExited"] = true; return true; }
        using (var process = owned)
        {
        if (process.StartTime.ToUniversalTime().Ticks != ticks) throw new InvalidOperationException("Office process identity changed");
        if (application == null)
        {
            bool alreadyExited = process.WaitForExit(5000);
            Report["ownedProcessExited"] = alreadyExited;
            return alreadyExited;
        }
            dynamic app = application;
            object documents;
            if (Application == "excel") documents = app.Workbooks;
            else if (Application == "powerpoint") documents = app.Presentations;
            else documents = app.Documents;
            try
            {
                if (Convert.ToInt32(((dynamic)documents).Count) != 0)
                { Report["applicationQuitSkippedNonempty"] = true; return false; }
            }
            finally { Release(documents); }
            Stage("quit-owned-empty-instance");
            if (Application == "word") app.Quit(0);
            else app.Quit();
            if (verifyExitAfterHostCloses)
            {
                Report["ownedProcessExitDeferred"] = true;
                return true;
            }
            bool exited = process.WaitForExit(5000);
            Report["ownedProcessExited"] = exited;
            return exited;
        }
    }

    internal void Finish(bool passed)
    {
        Report["passed"] = passed && !Report.ContainsKey("errorType");
        Report["elapsedMs"] = Clock.ElapsedMilliseconds;
        Report["finishedUtc"] = DateTime.UtcNow.ToString("o");
        WriteNew(Path.Combine(Root, "result.json"), Report);
    }

    internal void Error(string stage, Exception exception)
    {
        Report["errorType"] = exception.GetType().FullName;
        Report["errorStage"] = stage;
        Report["errorHresult"] = "0x" + Marshal.GetHRForException(exception).ToString("X8");
    }

    internal static string Hash(string file)
    {
        using (var hash = SHA256.Create())
        using (var stream = File.OpenRead(file)) return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "");
    }
    private Dictionary<string, object> ReadJson(string file)
    { CheckTree(file); return json.Deserialize<Dictionary<string, object>>(File.ReadAllText(file)); }
    internal void WriteNew(string file, object value)
    {
        CheckTree(file);
        using (var writer = new StreamWriter(new FileStream(file, FileMode.CreateNew, FileAccess.Write), new UTF8Encoding(false)))
            writer.Write(json.Serialize(value));
    }
    internal static void Release(object value)
    { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll")] private static extern IntPtr GetWindowDpiAwarenessContext(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern int GetAwarenessFromDpiAwarenessContext(IntPtr context);
}
