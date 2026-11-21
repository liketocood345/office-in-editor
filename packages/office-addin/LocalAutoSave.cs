using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeInEditor.OfficeAddin
{
    // Runs only while opted in. The heartbeat reads the in-memory Saved flag;
    // clean documents cause no file access, content serialization or Save calls.
    internal sealed class LocalAutoSave : IDisposable
    {
        private sealed class Stamp
        {
            internal string Path;
            internal long Length, LastWrite, LastSave;
            internal bool Paused;
            internal bool WasDirty;
        }
        private readonly dynamic application;
        private readonly string kind;
        private readonly OfficeDispatcher dispatcher;
        private readonly Action<string> report;
        private readonly DocumentRegistry documents;
        private ConditionalWeakTable<object, Stamp> stamps = new ConditionalWeakTable<object, Stamp>();
        private readonly Stopwatch clock = Stopwatch.StartNew();
        private readonly Func<long> elapsed;
        private readonly bool automaticTicks;
        private long Now => elapsed == null ? clock.ElapsedMilliseconds : elapsed();
        private Timer timer;
        private int running;
        private volatile bool disposed;
        internal bool Periodic { get; private set; }
        internal int Seconds { get; private set; } = 60;
        internal LocalAutoSave(object application, string kind, OfficeDispatcher dispatcher, Action<string> report, Func<long> elapsed = null, bool automaticTicks = true, DocumentRegistry documents = null)
        { this.application = application; this.kind = kind; this.dispatcher = dispatcher; this.report = report; this.elapsed = elapsed; this.automaticTicks = automaticTicks; this.documents = documents; }

        internal void Configure(bool periodic, int seconds)
        {
            if (seconds < 5 || seconds > 3600) throw new BridgeException("INVALID_SAVE_INTERVAL");
            Periodic = periodic; Seconds = seconds;
            stamps = new ConditionalWeakTable<object, Stamp>();
            timer?.Dispose(); timer = null;
            if (!disposed && automaticTicks && Periodic) timer = new Timer(Tick, null, 0, 1000);
        }
        private async void Tick(object ignored)
        {
            if (disposed || Interlocked.CompareExchange(ref running, 1, 0) != 0) return;
            try
            {
                using (var cancel = new CancellationTokenSource(2000))
                    await dispatcher.InvokeAsync(() => { if (!disposed) Check(); return true; }, cancel.Token).ConfigureAwait(false);
            }
            catch { /* Busy Office is retried by the one bounded heartbeat, never a new queue. */ }
            finally { Interlocked.Exchange(ref running, 0); }
        }
        internal void Check()
        {
            if (disposed || !Periodic) return;
            object target;
            try
            {
                switch (kind)
                {
                    case "word": target = application.ActiveDocument; break;
                    case "excel": target = application.ActiveWorkbook; break;
                    case "powerpoint": target = application.ActivePresentation; break;
                    default: return;
                }
                if (target == null) return;
                // Core alone owns each embedded session's timer and persistence.
                if (documents?.Embedded.Find(target) != null) return;
                dynamic document = target;
                bool dirty = !Convert.ToBoolean(document.Saved, CultureInfo.InvariantCulture);
                var stamp = stamps.GetValue(target, _ => new Stamp { LastSave = Now });
                if (stamp.Paused) return;
                bool due = Periodic && Now - stamp.LastSave >= Seconds * 1000L;
                // Snapshot once when opting in, including for initially clean documents.
                if (stamp.Path == null)
                {
                    string name = Convert.ToString(document.FullName, CultureInfo.InvariantCulture);
                    if (!LocalPath(name, kind)) { stamp.Paused = true; return; }
                    var info = new FileInfo(name);
                    if (!info.Exists || info.IsReadOnly) { stamp.Paused = true; return; }
                    stamp.Path = info.FullName; stamp.Length = info.Length; stamp.LastWrite = info.LastWriteTimeUtc.Ticks;
                }
                if (!dirty)
                {
                    if (stamp.WasDirty)
                    {
                        var manual = new FileInfo(stamp.Path);
                        stamp.Length = manual.Length; stamp.LastWrite = manual.LastWriteTimeUtc.Ticks;
                        stamp.LastSave = Now; stamp.WasDirty = false;
                    }
                    return;
                }
                stamp.WasDirty = true;
                if (!due) return;
                // Fail closed: an unknown read-only/protection state never permits saving.
                if (Convert.ToBoolean(document.ReadOnly, CultureInfo.InvariantCulture)) return;
                string current = Convert.ToString(document.FullName, CultureInfo.InvariantCulture);
                if (!string.Equals(Path.GetFullPath(current), stamp.Path, StringComparison.OrdinalIgnoreCase)) { stamp.Paused = true; return; }
                var before = new FileInfo(stamp.Path);
                if (!before.Exists || before.Length != stamp.Length || before.LastWriteTimeUtc.Ticks != stamp.LastWrite)
                { stamp.Paused = true; report("自动保存已暂停：文件被外部修改。切换开关可重新启用。"); return; }
                try
                {
                    if (kind == "word") RequireWordWritable(target);
                    ClickSave(application, target, stamp.Path, kind);
                    if (!Convert.ToBoolean(document.Saved, CultureInfo.InvariantCulture)) throw new BridgeException("SAVE_NOT_CONFIRMED");
                    var after = new FileInfo(stamp.Path);
                    stamp.Length = after.Length; stamp.LastWrite = after.LastWriteTimeUtc.Ticks;
                    stamp.LastSave = Now;
                    stamp.WasDirty = false;
                }
                catch { stamp.Paused = true; report("自动保存已暂停，请手动保存后重新启用。"); }
            }
            catch { /* No active document / Office busy / protected view: no popup or write. */ }
        }
        internal static bool LocalPath(string name, string kind)
        {
            if (string.IsNullOrEmpty(name) || name.Length < 3 || !char.IsLetter(name[0]) || name[1] != ':' || (name[2] != '\\' && name[2] != '/')) return false;
            string extension = Path.GetExtension(name).ToLowerInvariant();
            return (kind == "word" && extension == ".docx") || (kind == "excel" && extension == ".xlsx") || (kind == "powerpoint" && extension == ".pptx");
        }
        internal static void ClickSave(object application, object target, string path, string kind = "word")
        {
            if (kind == "word")
            {
                RequireWordWritable(target);
                if (!DocumentRegistry.SameObject((object)((dynamic)application).ActiveDocument, target))
                    throw new BridgeException("DOCUMENT_CHANGED");
                ((dynamic)application).CommandBars.ExecuteMso("FileSave");
                return;
            }
            // Retain the existing native save foundation for the other applications.
            Save(target, path);
        }
        internal static void Save(object target, string path)
        {
            ((dynamic)target).Save();
        }
        internal static void RequireWordWritable(object target)
        {
            dynamic document = target;
            try
            {
                if (Convert.ToBoolean(document.ReadOnly, CultureInfo.InvariantCulture) ||
                    Convert.ToInt32(document.ProtectionType, CultureInfo.InvariantCulture) != -1)
                    throw new BridgeException("DOCUMENT_PROTECTED");
            }
            catch (BridgeException) { throw; }
            catch { throw new BridgeException("DOCUMENT_PROTECTION_UNKNOWN"); }
        }
        public void Dispose() { disposed = true; timer?.Dispose(); timer = null; }
    }
}
