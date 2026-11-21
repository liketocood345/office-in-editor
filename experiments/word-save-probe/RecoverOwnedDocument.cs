using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Web.Script.Serialization;

internal static class RecoverOwnedDocument
{
    internal static int Run(ProbeRun run)
    {
        var result = new Dictionary<string, object> { { "passed", false } };
        object document = null, app = null, documents = null;
        IRunningObjectTable rot = null;
        IMoniker moniker = null;
        try
        {
            var previous = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(Path.Combine(run.Root, "result.json")));
            if (!previous.ContainsKey("applicationIsOwned") || !Convert.ToBoolean(previous["applicationIsOwned"]))
                throw new InvalidOperationException("No recorded owned application");
            int expectedPid = Convert.ToInt32(previous["applicationPid"]);
            using (var process = Process.GetProcessById(expectedPid))
            {
                long identity = process.StartTime.ToUniversalTime().Ticks;
                DateTime finished = DateTime.Parse(Convert.ToString(previous["finishedUtc"]), null, System.Globalization.DateTimeStyles.RoundtripKind).ToUniversalTime();
                DateTime earliest = finished.AddMilliseconds(-Convert.ToInt64(previous["elapsedMs"])).AddSeconds(-1);
                if (process.ProcessName != "WINWORD" || identity < earliest.Ticks || identity > finished.Ticks ||
                    (previous.ContainsKey("applicationStartUtcTicks") && identity != Convert.ToInt64(previous["applicationStartUtcTicks"])))
                    throw new InvalidOperationException("Recorded process identity no longer matches");
                Marshal.ThrowExceptionForHR(GetRunningObjectTable(0, out rot));
                Marshal.ThrowExceptionForHR(CreateFileMoniker(run.Document, out moniker));
                // ROT lookup cannot activate or bind to a new server. Never use
                // GetActiveObject, enumeration by title or a new Word instance.
                if (rot.IsRunning(moniker) != 0) throw new InvalidOperationException("Owned file is not registered in ROT");
                rot.GetObject(moniker, out document);
                if (document == null) throw new InvalidOperationException("No running file object");
                dynamic doc = document;
                app = doc.Application;
                if (run.BindApplication(app, document) != expectedPid ||
                    !String.Equals(Path.GetFullPath(Convert.ToString(doc.FullName)), run.Document, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Running document identity mismatch");
                object content = doc.Content;
                try { if (!Convert.ToString(((dynamic)content).Text).Contains(run.Marker)) throw new InvalidDataException("Owned document marker missing"); }
                finally { ProbeRun.Release(content); }
                doc.Save();
                doc.Close(0);
                ProbeRun.Release(document); document = null;
                documents = ((dynamic)app).Documents;
                if (Convert.ToInt32(((dynamic)documents).Count) != 0) throw new InvalidOperationException("Additional documents retained; application not quit");
                ProbeRun.Release(documents); documents = null;
                if (process.StartTime.ToUniversalTime().Ticks != identity) throw new InvalidOperationException("Process changed before quit");
                ((dynamic)app).Quit(0);
                result["passed"] = process.WaitForExit(5000);
                result["savedAndClosedOwnedDocument"] = true;
            }
        }
        catch (Exception e) { result["errorType"] = e.GetType().FullName; result["hresult"] = "0x" + Marshal.GetHRForException(e).ToString("X8"); }
        finally { ProbeRun.Release(documents); ProbeRun.Release(document); ProbeRun.Release(app); ProbeRun.Release(moniker); ProbeRun.Release(rot); }
        run.WriteNew(Path.Combine(run.Root, "recovery-" + Guid.NewGuid().ToString("N") + ".json"), result);
        Console.WriteLine(new JavaScriptSerializer().Serialize(result));
        return (bool)result["passed"] ? 0 : 1;
    }
    [DllImport("ole32.dll")] private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable rot);
    [DllImport("ole32.dll", CharSet=CharSet.Unicode)] private static extern int CreateFileMoniker(string path, out IMoniker moniker);
}
