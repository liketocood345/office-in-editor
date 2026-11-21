using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;

// One STA and one owned synthetic document. No browser window or UI automation.
internal static class WordWorker
{
    private const int MaxText = 4096;
    private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = 32768 };

    [STAThread]
    private static int Main(string[] args)
    {
        Console.InputEncoding = new UTF8Encoding(false);
        Console.OutputEncoding = new UTF8Encoding(false);
        var report = new Dictionary<string, object> { { "mode", "word-browser-local" } };
        ProbeRun run = null;
        object app = null, documents = null, document = null;
        int pid = 0, revision = 0;
        bool owned = false, canClose = true, closeRequested = false;
        try
        {
            if ((args.Length != 1 && (args.Length != 2 || args[1] != "--prepare")) ||
                System.Diagnostics.Process.GetCurrentProcess().SessionId == 0) throw new InvalidOperationException();
            run = new ProbeRun(args[0], report);
            if (args.Length == 2)
            {
                if (File.Exists(run.Document)) run.ValidateDocument(run.Document);
                else run.Prepare();
                return 0;
            }
            if (!File.Exists(run.Document)) run.Prepare();
            run.ValidateDocument(run.Document);
            run.Stage("create-word");
            app = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application", true));
            documents = ((dynamic)app).Documents;
            run.Stage("open-synthetic");
            document = ((dynamic)documents).Open(FileName: run.Document, ConfirmConversions: false,
                ReadOnly: false, AddToRecentFiles: false, Visible: false);
            pid = run.BindApplication(app, document);
            owned = Convert.ToBoolean(report["applicationIsOwned"]);
            if (!owned) throw new InvalidOperationException("An independent Word instance is required.");
            ((dynamic)app).Visible = false;
            ((dynamic)app).AutomationSecurity = 3;
            AssertDocument(document, run);
            string text = ReadText(document);
            if (text != "OIE synthetic seed") throw new InvalidDataException();
            var stamp = SourceStamp(run.Document);
            bool hidden = !IsWindowVisible(new IntPtr(Convert.ToInt64(report["documentWindowHwnd"])));
            if (!hidden) throw new InvalidOperationException("The Word view must remain hidden.");
            report["wordWindowHidden"] = hidden;
            run.Stage("ready");
            Respond(new { ok = true, text, revision, engine = "Microsoft Word", wordWindowHidden = hidden });

            string line;
            while ((line = Console.ReadLine()) != null)
            {
                if (line.Length > 32768) { Respond(new { ok = false, error = "invalid_request" }); continue; }
                var message = Json.Deserialize<Dictionary<string, object>>(line);
                string op = message.ContainsKey("op") ? message["op"] as string : null;
                if (op == "read")
                {
                    Respond(new { ok = true, text, revision, engine = "Microsoft Word" });
                    continue;
                }
                if (op == "close")
                {
                    if (canClose) { closeRequested = true; break; }
                    Respond(new { ok = false, error = "save_unconfirmed" });
                    continue;
                }
                object next, expected;
                if (op != "save" || !message.TryGetValue("text", out next) || !IsText(next as string) ||
                    !message.TryGetValue("revision", out expected) || !(expected is int))
                { Respond(new { ok = false, error = "invalid_request" }); continue; }
                if (!canClose) { Respond(new { ok = false, error = "save_unconfirmed" }); continue; }
                if ((int)expected != revision) { Respond(new { ok = false, error = "edit_conflict" }); continue; }
                string newText = (string)next;
                if (newText == text) { Respond(new { ok = true, text, revision, changed = false }); continue; }
                AssertDocument(document, run);
                if (!SourceStamp(run.Document).Equals(stamp) || ReadText(document) != text)
                { Respond(new { ok = false, error = "document_conflict" }); continue; }

                canClose = false;
                try
                {
                    ReplaceText(document, newText);
                    ((dynamic)document).Save();
                    if (!Convert.ToBoolean(((dynamic)document).Saved) || ReadText(document) != newText)
                        throw new InvalidOperationException("Word did not confirm the save.");
                    stamp = SourceStamp(run.Document);
                    text = newText;
                    revision++;
                    report["confirmedSaves"] = revision;
                    canClose = true;
                    Respond(new { ok = true, text, revision, changed = true, wordSaved = true });
                }
                catch (Exception error)
                {
                    run.Error("save", error);
                    Respond(new { ok = false, error = "save_unconfirmed" });
                }
            }
            if (!canClose) Application.Run(); // Retain the Word session for recovery if its caller disappears.
            AssertDocument(document, run);
            ((dynamic)document).Close(0);
            ProbeRun.Release(document); document = null;
            ProbeRun.Release(documents); documents = null;
            bool clean = run.QuitIfOwnedAndEmpty(app, pid);
            run.Finish(clean);
            if (closeRequested) Respond(new { ok = clean, error = clean ? null : "close_unconfirmed" });
            return clean ? 0 : 1;
        }
        catch (Exception error)
        {
            if (run != null)
            {
                run.Error("worker", error);
                if (document != null && canClose)
                {
                    try
                    {
                        if (String.Equals((string)((dynamic)document).FullName, run.Document, StringComparison.OrdinalIgnoreCase))
                        {
                            ((dynamic)document).Close(0); // The worker never edited this clean document.
                            ProbeRun.Release(document); document = null;
                            ProbeRun.Release(documents); documents = null;
                            if (owned) run.QuitIfOwnedAndEmpty(app, pid);
                        }
                    }
                    catch { /* No process termination fallback. */ }
                }
                run.Finish(false);
            }
            Respond(new { ok = false, error = "word_unavailable" });
            return 1;
        }
        finally { ProbeRun.Release(document); ProbeRun.Release(documents); ProbeRun.Release(app); }
    }

    private static void AssertDocument(object document, ProbeRun run)
    {
        dynamic doc = document;
        if (!String.Equals((string)doc.FullName, run.Document, StringComparison.OrdinalIgnoreCase) ||
            Convert.ToBoolean(doc.ReadOnly) || Convert.ToInt32(doc.ProtectionType) != -1)
            throw new InvalidOperationException("Document identity or write permission changed.");
    }

    private static string ReadText(object document)
    {
        object range = ((dynamic)document).Content;
        try
        {
            string value = (string)((dynamic)range).Text;
            if (!value.EndsWith("\r", StringComparison.Ordinal)) throw new InvalidDataException();
            value = value.Substring(0, value.Length - 1);
            if (!IsText(value)) throw new InvalidDataException();
            return value;
        }
        finally { ProbeRun.Release(range); }
    }

    private static bool IsText(string text)
    {
        if (text == null || text.Length > MaxText) return false;
        foreach (char c in text) if (Char.IsControl(c)) return false;
        return true;
    }

    private static void ReplaceText(object document, string text)
    {
        object range = ((dynamic)document).Content;
        try
        {
            dynamic editable = range;
            editable.End = Convert.ToInt32(editable.End) - 1;
            editable.Text = text;
        }
        finally { ProbeRun.Release(range); }
    }

    private static Tuple<long, long> SourceStamp(string filename)
    {
        ProbeRun.CheckTree(filename);
        var file = new FileInfo(filename);
        if (!file.Exists || file.IsReadOnly || file.Length > 1048576) throw new IOException();
        return Tuple.Create(file.Length, file.LastWriteTimeUtc.Ticks);
    }

    private static void Respond(object response) { Console.WriteLine(Json.Serialize(response)); Console.Out.Flush(); }
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
}
