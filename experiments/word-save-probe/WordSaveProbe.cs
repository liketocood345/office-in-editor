using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

internal static class WordSaveProbe
{
    [STAThread]
    private static int Main(string[] args)
    {
        var report = new Dictionary<string, object> { { "mode", "word-save" }, { "passed", false } };
        ProbeRun run = null;
        object appObject = null, docObject = null, documents = null, content = null;
        bool closed = false, cleanup = false;
        int appPid = 0;
        try
        {
            if (args.Length < 1 || args.Length > 2) throw new ArgumentException("Run root and optional --prepare/--visible required");
            run = new ProbeRun(args[0], report);
            if (args.Length == 2 && args[1] == "--prepare") { run.Prepare(); return 0; }
            if (args.Length == 2 && args[1] == "--recover") return RecoverOwnedDocument.Run(run);
            bool visible = args.Length == 2 && args[1] == "--visible";
            if (args.Length == 2 && !visible) throw new ArgumentException("Unknown mode");
            run.ValidateDocument(run.Document);
            run.Stage("create-application");
            appObject = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application", true));
            dynamic app = appObject;
            documents = app.Documents;
            run.Stage("open-synthetic-document");
            docObject = ((dynamic)documents).Open(FileName: run.Document, ConfirmConversions: false,
                ReadOnly: false, AddToRecentFiles: false, Visible: visible);
            run.Stage("bind-document-window");
            appPid = run.BindApplication(appObject, docObject);
            if (!Convert.ToBoolean(report["applicationIsOwned"])) throw new InvalidOperationException("New owned instance required");
            app.Visible = visible;
            app.DisplayAlerts = 0;
            app.AutomationSecurity = 3;
            dynamic doc = docObject;
            content = doc.Content;
            ((dynamic)content).Text = run.Marker;
            ProbeRun.Release(content); content = null;
            run.Stage("saveas-single-argument");
            doc.SaveAs2(run.Document);
            report["singleArgumentSaveReturned"] = true;

            // A timer requests the same official save command as the UI button.
            // Exactly two ticks, with no close/reopen or continuous disk polling.
            int saves = 0;
            report["periodicSaveMechanism"] = visible ? "FileSave-command" : "Document.Save";
            Exception saveError = null;
            string finalMarker = run.Marker + " periodic-2";
            using (var timer = new Timer { Interval = 1500 })
            {
                timer.Tick += delegate {
                    timer.Stop();
                    try
                    {
                        saves++;
                        content = doc.Content;
                        ((dynamic)content).Text = run.Marker + " periodic-" + saves;
                        ProbeRun.Release(content); content = null;
                        run.Stage("periodic-save-" + saves);
                        if (visible)
                        {
                            object active = app.ActiveDocument;
                            try
                            {
                                if (Convert.ToInt32(((dynamic)documents).Count) != 1 ||
                                    !String.Equals(Path.GetFullPath(Convert.ToString(((dynamic)active).FullName)), run.Document, StringComparison.OrdinalIgnoreCase))
                                    throw new InvalidOperationException("Save command target changed");
                            }
                            finally { ProbeRun.Release(active); }
                            object bars = app.CommandBars;
                            try { ((dynamic)bars).ExecuteMso("FileSave"); } finally { ProbeRun.Release(bars); }
                        }
                        else doc.Save(); // hidden windows do not have an enabled UI command
                        if (saves < 2) timer.Start();
                    }
                    catch (Exception e) { saveError = e; saves = 2; }
                };
                timer.Start();
                while (saves < 2) { Application.DoEvents(); System.Threading.Thread.Sleep(25); }
            }
            if (saveError != null) throw saveError;
            report["periodicSaveRequests"] = saves;
            report["sameDocumentThroughout"] = String.Equals(Path.GetFullPath(Convert.ToString(doc.FullName)), run.Document, StringComparison.OrdinalIgnoreCase);
            run.Stage("verify-before-close");
            // Word may lock the file against ZIP readers while editing; verify
            // persisted XML after normal close, and require Saved before close.
            if (!Convert.ToBoolean(doc.Saved)) throw new InvalidOperationException("Save command left document dirty");
            if (!(bool)report["sameDocumentThroughout"]) throw new InvalidOperationException("Document changed");
            run.Stage("close-synthetic-document");
            doc.Close(0);
            closed = true;
            ProbeRun.Release(docObject); docObject = null;
            ProbeRun.Release(documents); documents = null;
            cleanup = run.QuitIfOwnedAndEmpty(appObject, appPid);
            report["documentClosed"] = closed;
            report["markerInSavedFile"] = run.ContainsMarker(finalMarker);
            report["documentVerified"] = report["markerInSavedFile"];
            if (!(bool)report["markerInSavedFile"] || !cleanup) throw new InvalidOperationException("Save verification or cleanup failed");
        }
        catch (Exception e)
        {
            if (run == null) { Console.Error.WriteLine(e.GetType().FullName); return 2; }
            run.Error(Convert.ToString(report.ContainsKey("stage") ? report["stage"] : "preflight"), e);
            // Preserve the synthetic edits before closing an owned failed test.
            // A recovery save never turns the failed experiment into a pass.
            if (docObject != null && appPid != 0 && report.ContainsKey("applicationIsOwned") && (bool)report["applicationIsOwned"])
            {
                try
                {
                    dynamic failed = docObject;
                    if (!String.Equals(Path.GetFullPath(Convert.ToString(failed.FullName)), run.Document, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("Recovery document identity mismatch");
                    failed.Save();
                    failed.Close(0);
                    closed = true;
                    ProbeRun.Release(docObject); docObject = null;
                    ProbeRun.Release(documents); documents = null;
                    cleanup = run.QuitIfOwnedAndEmpty(appObject, appPid);
                    report["failedTestRecoveredSafely"] = cleanup;
                }
                catch (Exception recovery) { report["recoveryErrorType"] = recovery.GetType().FullName; }
            }
        }
        finally
        {
            // No force-kill and no speculative Quit on error. Preserve a failed
            // instance for diagnosis; never discard unsaved data in a catch.
            ProbeRun.Release(content); ProbeRun.Release(docObject);
            ProbeRun.Release(documents); ProbeRun.Release(appObject);
        }
        run.Finish(closed && cleanup);
        Console.WriteLine(new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(report));
        return (bool)report["passed"] ? 0 : 1;
    }
}
