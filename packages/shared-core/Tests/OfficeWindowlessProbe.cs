using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Win32;

// An opt-in, single-application probe. Query the actual Office document OLE
// server without loading a file, creating a site, or activating any view.
internal static class OfficeWindowlessProbe
{
    private const uint LocalServer = 4;
    private const int NoInterface = unchecked((int)0x80004002);

    internal static int Run(string application)
    {
        var (progId, processName) = application switch
        {
            "word" => ("Word.Document.12", "WINWORD"),
            "excel" => ("Excel.Sheet.12", "EXCEL"),
            "powerpoint" => ("PowerPoint.Show.12", "POWERPNT"),
            _ => (null, null)
        };
        if (progId is null || processName is null) return 2;
        int result = 1;
        var thread = new Thread(() =>
        {
            try { result = QueryOfficeClass(application, progId, processName); }
            catch (Exception error)
            {
                Write(new { application, outcome = "probe_failed", hresult = HResult(error.HResult) });
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        return result;
    }

    private static int QueryOfficeClass(string application, string progId, string processName)
    {
        using var key = Registry.ClassesRoot.OpenSubKey(progId + @"\CLSID", writable: false);
        if (!Guid.TryParse(key?.GetValue(null) as string, out var classId))
        {
            Write(new { application, progId, outcome = "class_not_registered" });
            return 1;
        }
        if (AppRunning(processName))
        {
            Write(new { application, progId, outcome = "existing_office_process" });
            return 1;
        }

        var unknownId = new Guid("00000000-0000-0000-C000-000000000046");
        var created = CoCreateInstance(ref classId, IntPtr.Zero, LocalServer, ref unknownId, out var instance);
        if (created < 0 || instance == IntPtr.Zero)
        {
            Write(new { application, progId, outcome = "ole_class_create_failed", hresult = HResult(created) });
            return 1;
        }
        try
        {
            var oleObject = Query(instance, "00000112-0000-0000-C000-000000000046");
            var oleDocument = Query(instance, "B722BCC5-4E68-101B-A2BC-00AA00404770");
            var inPlaceObject = Query(instance, "00000113-0000-0000-C000-000000000046");
            var windowless = Query(instance, "1C2056CC-5EF4-101B-8BC8-00AA003E3B29");
            var viewObjectEx = Query(instance, "3AF24292-0C96-11CE-A0CF-00AA00600AB8");
            string outcome = windowless switch
            {
                0 => "windowless_interface_reported",
                NoInterface => "windowless_interface_absent_on_ole_class",
                _ => "windowless_query_inconclusive"
            };
            Write(new
            {
                application, progId, outcome, oleObject = HResult(oleObject), oleDocument = HResult(oleDocument),
                inPlaceObject = HResult(inPlaceObject), windowless = HResult(windowless), viewObjectEx = HResult(viewObjectEx),
                documentLoaded = false, inPlaceActivationAttempted = false
            });
            return windowless == 0 || windowless == NoInterface ? 0 : 1;
        }
        finally { Marshal.Release(instance); }
    }

    private static bool AppRunning(string processName)
    {
        var processes = Process.GetProcessesByName(processName);
        foreach (var process in processes) process.Dispose();
        return processes.Length != 0;
    }

    private static int Query(IntPtr instance, string iid)
    {
        var interfaceId = new Guid(iid);
        int hr = Marshal.QueryInterface(instance, ref interfaceId, out var pointer);
        if (pointer != IntPtr.Zero) Marshal.Release(pointer);
        return hr;
    }

    private static string HResult(int value) => $"0x{value:X8}";

    private static void Write<T>(T report) => Console.WriteLine(JsonSerializer.Serialize(report));

    [DllImport("ole32.dll", ExactSpelling = true)]
    private static extern int CoCreateInstance(ref Guid clsid, IntPtr outer, uint context, ref Guid iid, out IntPtr instance);
}
