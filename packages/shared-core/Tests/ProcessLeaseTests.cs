using System.Diagnostics;
using OfficeInEditor.Core;

/// <summary>
/// Regression for the lab Core hang: a host lease and its chat lease watch the same process. When it exits, both exit
/// callbacks run at once; the first takes the Core lock and disposes the other lease while that lease's own callback
/// (which holds its Process lock) waits for the Core lock. Disposing a lease must not wait for that Process lock.
/// </summary>
static class ProcessLeaseTests
{
    public static void Run(Action<bool, string> check)
    {
        var shared = new object();
        using var owner = Process.Start(new ProcessStartInfo("cmd.exe", "/c ping -n 2 127.0.0.1 >nul") { UseShellExecute = false, CreateNoWindow = true })!;
        ProcessLease? chat = null;
        var hostDone = new ManualResetEventSlim(false);
        var chatDone = new ManualResetEventSlim(false);
        var host = new ProcessLease(owner.Id, () =>
        {
            lock (shared)
            {
                // Let the chat lease's exit callback start and block on the shared lock first.
                Thread.Sleep(300);
                chat!.Dispose();
            }
            hostDone.Set();
        });
        chat = new ProcessLease(owner.Id, () =>
        {
            Thread.Sleep(100);
            lock (shared) { }
            chatDone.Set();
        });
        var finished = hostDone.Wait(TimeSpan.FromSeconds(10));
        check(finished, "process lease disposed from another lease's exit callback does not deadlock");
        chatDone.Wait(TimeSpan.FromSeconds(5));
        host.Dispose();
    }
}
