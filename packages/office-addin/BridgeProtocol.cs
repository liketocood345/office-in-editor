using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace OfficeInEditor.OfficeAddin
{
    internal sealed class BridgeException : Exception
    {
        internal string Code { get; }
        internal BridgeException(string code) : base(code) { Code = code; }
    }

    internal static class Json
    {
        internal static Dictionary<string, object> Object(params object[] pairs)
        {
            var result = new Dictionary<string, object>(StringComparer.Ordinal);
            for (int i = 0; i < pairs.Length; i += 2) result.Add((string)pairs[i], pairs[i + 1]);
            return result;
        }
        internal static JavaScriptSerializer Serializer() => new JavaScriptSerializer
            { MaxJsonLength = Wire.MaxLineBytes, RecursionLimit = 40 };
        internal static string Encode(object value) => Serializer().Serialize(value);
        internal static Dictionary<string, object> Map(object value) => value as Dictionary<string, object>;
        internal static object Get(object value, string name)
        {
            var map = Map(value);
            return map != null && map.TryGetValue(name, out object result) ? result : null;
        }
        internal static string Text(object value, string name) => Get(value, name) as string;
        internal static IEnumerable Array(object value) => value is string || value is IDictionary
            ? new object[0] : value as IEnumerable ?? new object[0];
        internal static bool IsTrue(object value) => value is bool b && b;
    }

    internal static class Wire
    {
        internal const int MaxLineBytes = 2 * 1024 * 1024;
        internal static readonly Encoding Utf8 = new UTF8Encoding(false, true);
        internal static async Task<string> ReadLineAsync(Stream stream, CancellationToken cancel)
        {
            using (var data = new MemoryStream())
            {
                var buffer = new byte[4096];
                while (true)
                {
                    int count = await stream.ReadAsync(buffer, 0, buffer.Length, cancel).ConfigureAwait(false);
                    if (count == 0) throw new BridgeException("INCOMPLETE_REQUEST");
                    int end = System.Array.IndexOf(buffer, (byte)'\n', 0, count);
                    int used = end >= 0 ? end : count;
                    if (data.Length + used > MaxLineBytes) throw new BridgeException("REQUEST_TOO_LARGE");
                    data.Write(buffer, 0, used);
                    if (end >= 0) return Utf8.GetString(data.ToArray()).TrimEnd('\r');
                }
            }
        }
        internal static async Task WriteLineAsync(Stream stream, object value, CancellationToken cancel)
        {
            byte[] bytes = Utf8.GetBytes(Json.Encode(value) + "\n");
            if (bytes.Length - 1 > MaxLineBytes) throw new BridgeException("RESPONSE_TOO_LARGE");
            await stream.WriteAsync(bytes, 0, bytes.Length, cancel).ConfigureAwait(false);
            await stream.FlushAsync(cancel).ConfigureAwait(false);
        }
        internal static string CorePipeName()
        {
            string sid = WindowsIdentity.GetCurrent().User.Value;
            using (var sha = SHA256.Create())
                return "office-in-editor-v1-" + BitConverter.ToString(sha.ComputeHash(Utf8.GetBytes(sid)))
                    .Replace("-", "").ToLowerInvariant().Substring(0, 16);
        }
        internal static PipeSecurity CurrentUserSecurity()
        {
            var sid = WindowsIdentity.GetCurrent().User;
            var security = new PipeSecurity();
            security.SetAccessRuleProtection(true, false);
            security.SetOwner(sid);
            security.AddAccessRule(new PipeAccessRule(new SecurityIdentifier(WellKnownSidType.NetworkSid, null),
                PipeAccessRights.FullControl, AccessControlType.Deny));
            security.AddAccessRule(new PipeAccessRule(sid, PipeAccessRights.FullControl, AccessControlType.Allow));
            return security;
        }
    }

    internal sealed class CoreClient
    {
        internal async Task<object> CallAsync(string method, object arguments, bool ensureBroker = false)
        {
            if (ensureBroker) await EnsureBrokerAsync().ConfigureAwait(false);
            using (var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(8)))
            using (var client = new NamedPipeClientStream(".", Wire.CorePipeName(), PipeDirection.InOut,
                PipeOptions.Asynchronous, TokenImpersonationLevel.Identification))
            {
                // A single connection attempt; no retries or write-operation replay.
                await client.ConnectAsync(ensureBroker ? 3000 : 250, timeout.Token).ConfigureAwait(false);
                string id = Guid.NewGuid().ToString("N");
                await Wire.WriteLineAsync(client, Json.Object("v", 1, "id", id, "method", method,
                    "params", arguments ?? Json.Object()), timeout.Token).ConfigureAwait(false);
                var reply = Json.Serializer().DeserializeObject(await Wire.ReadLineAsync(client, timeout.Token)
                    .ConfigureAwait(false));
                if (Json.Text(reply, "id") != id) throw new BridgeException("INVALID_RESPONSE");
                if (!Json.IsTrue(Json.Get(reply, "ok")))
                    throw new BridgeException(SafeError(Json.Text(Json.Get(reply, "error"), "code")));
                return Json.Get(reply, "result");
            }
        }

        internal static string SafeError(string code)
        {
            if (string.IsNullOrEmpty(code) || code.Length > 64) return "REQUEST_FAILED";
            foreach (char c in code) if (!(c >= 'A' && c <= 'Z') && c != '_') return "REQUEST_FAILED";
            return code;
        }

        private static async Task EnsureBrokerAsync()
        {
            string executable = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Programs", "OfficeInEditor", "current", "core", "OfficeInEditor.Core.exe");
            if (!File.Exists(executable)) throw new BridgeException("CORE_NOT_INSTALLED");
            using (var process = new Process())
            {
                process.StartInfo = new ProcessStartInfo(executable, "--ensure-broker")
                { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden };
                if (!process.Start()) throw new BridgeException("CORE_NOT_AVAILABLE");
                // No output is logged or accumulated; this runs off Office's STA.
                if (!await Task.Run(() => process.WaitForExit(6000)).ConfigureAwait(false))
                    throw new BridgeException("CORE_START_TIMEOUT");
                if (process.ExitCode != 0)
                    throw new BridgeException("CORE_NOT_AVAILABLE");
            }
        }
    }

    internal sealed class OfficeDispatcher : IDisposable
    {
        private const uint ThreadSetContext = 0x0010;
        private readonly Dictionary<int, Action> pending = new Dictionary<int, Action>();
        private readonly int threadId;
        private readonly uint nativeThreadId;
        private readonly IntPtr threadHandle;
        private static readonly ApcCallback apc = ExecuteApc;
        private static readonly TimerCallback timerCallback = ExecuteTimer;
        private static readonly Dictionary<int, OfficeDispatcher> owners = new Dictionary<int, OfficeDispatcher>();
        private static readonly Dictionary<(uint, UIntPtr), OfficeDispatcher> timers = new Dictionary<(uint, UIntPtr), OfficeDispatcher>();
        private static int nextOwner;
        private readonly int ownerToken;
        private readonly UIntPtr timer;
        private bool wakeQueued;
        internal int PendingCount { get { lock (pending) return pending.Count; } }
        private int next;
        private volatile bool disposed;
        internal OfficeDispatcher()
        {
            threadId = Thread.CurrentThread.ManagedThreadId;
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new BridgeException("OFFICE_STA_REQUIRED");
            nativeThreadId = GetCurrentThreadId();
            threadHandle = OpenThread(ThreadSetContext, false, nativeThreadId);
            if (threadHandle == IntPtr.Zero) throw new BridgeException("OFFICE_STA_REQUIRED");
            ownerToken = Interlocked.Increment(ref nextOwner);
            // An idle Office STA can pump window messages without making an alertable wait.
            timer = SetTimer(IntPtr.Zero, UIntPtr.Zero, 250, timerCallback);
            lock (owners)
            {
                owners[ownerToken] = this;
                if (timer != UIntPtr.Zero) timers[(nativeThreadId, timer)] = this;
            }
        }
        internal void Execute(int token)
        {
            Action work;
            lock (pending)
            {
                if (!pending.TryGetValue(token, out work)) return;
                pending.Remove(token);
            }
            work();
        }
        private static void ExecuteApc(UIntPtr token)
        {
            OfficeDispatcher owner;
            lock (owners) if (!owners.TryGetValue(unchecked((int)token.ToUInt32()), out owner)) return;
            owner.Drain();
        }
        private static void ExecuteTimer(IntPtr window, uint message, UIntPtr timerId, uint time)
        {
            OfficeDispatcher owner;
            lock (owners) timers.TryGetValue((GetCurrentThreadId(), timerId), out owner);
            if (owner == null) KillTimer(IntPtr.Zero, timerId);
            else owner.Drain();
        }
        private void Drain()
        {
            int[] work;
            lock (pending)
            {
                wakeQueued = false;
                if (pending.Count == 0) return;
                work = new int[pending.Count];
                pending.Keys.CopyTo(work, 0);
            }
            foreach (int item in work) try { Execute(item); } catch { }
        }
        internal async Task<T> InvokeAsync<T>(Func<T> callback, CancellationToken cancel)
        {
            if (disposed) throw new BridgeException("OFFICE_DISCONNECTED");
            if (Thread.CurrentThread.ManagedThreadId == threadId) return callback();
            var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            int token;
            bool schedule;
            lock (pending)
            {
                    if (disposed) throw new BridgeException("OFFICE_DISCONNECTED");
                    if (pending.Count >= 32) throw new BridgeException("OFFICE_BUSY");
                    token = ++next;
                    pending[token] = () =>
                    {
                        if (disposed || cancel.IsCancellationRequested) { completion.TrySetCanceled(); return; }
                        try { completion.TrySetResult(callback()); }
                        catch (Exception error) { completion.TrySetException(error); }
                    };
                    schedule = !wakeQueued;
                    wakeQueued = true;
            }
            using (cancel.Register(() => { lock (pending) pending.Remove(token); completion.TrySetCanceled(); }))
            {
                if (schedule && QueueUserAPC(apc, threadHandle, new UIntPtr(unchecked((uint)ownerToken))) == 0)
                {
                    lock (pending) { pending.Remove(token); wakeQueued = false; }
                    throw new BridgeException("OFFICE_DISCONNECTED");
                }
                return await completion.Task.ConfigureAwait(false);
            }
        }
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            lock (owners)
            {
                owners.Remove(ownerToken);
                timers.Remove((nativeThreadId, timer));
            }
            if (timer != UIntPtr.Zero && GetCurrentThreadId() == nativeThreadId) KillTimer(IntPtr.Zero, timer);
            if (threadHandle != IntPtr.Zero) try { CloseHandle(threadHandle); } catch { }
            List<Action> abandoned;
            lock (pending) { abandoned = new List<Action>(pending.Values); pending.Clear(); }
            foreach (var work in abandoned) try { work(); } catch { }
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate void ApcCallback(UIntPtr parameter);
        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate void TimerCallback(IntPtr window, uint message, UIntPtr timer, uint time);
        [DllImport("user32.dll")]
        private static extern UIntPtr SetTimer(IntPtr window, UIntPtr timer, uint milliseconds, TimerCallback callback);
        [DllImport("user32.dll")]
        private static extern bool KillTimer(IntPtr window, UIntPtr timer);
        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr OpenThread(uint access, bool inherit, uint threadId);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern uint QueueUserAPC(ApcCallback callback, IntPtr thread, UIntPtr data);
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr handle);
    }

    internal sealed class OfficeCallbackServer : IDisposable
    {
        private readonly string name;
        private readonly OfficeDispatcher dispatcher;
        private readonly Func<string, object, object> handle;
        private readonly CancellationTokenSource shutdown = new CancellationTokenSource();
        private readonly object gate = new object();
        private NamedPipeServerStream active;
        internal OfficeCallbackServer(string name, OfficeDispatcher dispatcher, Func<string, object, object> handle)
        { this.name = name; this.dispatcher = dispatcher; this.handle = handle; }
        internal void Start() => Task.Run(ListenAsync);
        private async Task ListenAsync()
        {
            while (!shutdown.IsCancellationRequested)
            {
                bool connected = false;
                try
                {
                    using (var pipe = new NamedPipeServerStream(name, PipeDirection.InOut, 1,
                        PipeTransmissionMode.Byte, PipeOptions.Asynchronous, 4096, 4096, Wire.CurrentUserSecurity()))
                    {
                        lock (gate) { if (shutdown.IsCancellationRequested) return; active = pipe; }
                        await pipe.WaitForConnectionAsync(shutdown.Token).ConfigureAwait(false);
                        connected = true;
                        using (var timeout = CancellationTokenSource.CreateLinkedTokenSource(shutdown.Token))
                        {
                            timeout.CancelAfter(TimeSpan.FromSeconds(12));
                            // Disposing on timeout also interrupts Framework pipe I/O that ignores cancellation.
                            using (timeout.Token.Register(() => { try { pipe.Dispose(); } catch { } }))
                            {
                                object id = null;
                                object reply;
                                try
                                {
                                    object request = Json.Serializer().DeserializeObject(await Wire.ReadLineAsync(pipe,
                                        timeout.Token).ConfigureAwait(false));
                                    id = Json.Get(request, "id");
                                    if (!(id is string) || ((string)id).Length > 128 ||
                                        Convert.ToInt32(Json.Get(request, "v")) != 1)
                                        throw new BridgeException("INVALID_REQUEST");
                                    string method = Json.Text(request, "method");
                                    if (string.IsNullOrEmpty(method) || method.Length > 64)
                                        throw new BridgeException("METHOD_NOT_FOUND");
                                    object parameters = Json.Get(request, "params");
                                    object result = await dispatcher.InvokeAsync(() => handle(method, parameters), timeout.Token)
                                        .ConfigureAwait(false);
                                    reply = Json.Object("id", id, "ok", true, "result", result);
                                }
                                catch (BridgeException error)
                                { reply = Json.Object("id", id, "ok", false, "error", Json.Object("code", error.Code,
                                    "message", "The Office request was rejected.")); }
                                catch
                                { reply = Json.Object("id", id, "ok", false, "error", Json.Object("code", "OFFICE_BUSY_OR_UNAVAILABLE",
                                    "message", "The Office request was rejected.")); }
                                await Wire.WriteLineAsync(pipe, reply, timeout.Token).ConfigureAwait(false);
                            }
                        }
                    }
                }
                catch
                {
                    // Bind/accept failure disables this listener instead of a background retry loop.
                    if (!connected) return;
                }
                finally { lock (gate) active = null; }
                // No failed-listener retry storm: only accept errors from an existing client are recoverable.
                if (shutdown.IsCancellationRequested) return;
            }
        }
        public void Dispose()
        {
            shutdown.Cancel();
            lock (gate) { try { active?.Dispose(); } catch { } }
        }
    }
}
