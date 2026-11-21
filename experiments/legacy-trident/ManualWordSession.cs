using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

internal static class ManualWordSession
{
    internal static async Task<string> RunAsync(Form form, ProbeRun run, IProbeFramer native, object document)
    {
        var done = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var body = form.Controls[0];
        body.Dock = DockStyle.None;
        var bar = new ToolStrip { Dock = DockStyle.None, GripStyle = ToolStripGripStyle.Hidden, AutoSize = false, Height = 36 };
        var save = new ToolStripButton("保存");
        var periodic = new ToolStripButton("周期保存（10 秒）") { CheckOnClick = true };
        var finish = new ToolStripButton("保存并结束");
        var status = new ToolStripLabel("合成文档 · 等待手测");
        bar.Items.AddRange(new ToolStripItem[] { save, periodic, finish, status });
        form.Controls.Add(bar);
        Action layout = delegate {
            bar.SetBounds(0, 0, form.ClientSize.Width, 36);
            body.SetBounds(0, 36, form.ClientSize.Width, Math.Max(1, form.ClientSize.Height - 36));
        };
        EventHandler resized = delegate { layout(); };
        form.Resize += resized;
        layout();
        bool finishing = false, saving = false;
        int manualSaves = 0, automaticSaves = 0;
        Func<bool, bool> saveDocument = delegate(bool automatic) {
            if (saving || finishing) return false;
            saving = true;
            try
            {
                if (automatic && Convert.ToBoolean(((dynamic)document).Saved)) return true;
                int hr = native.Save(null, true, null, null);
                if (hr < 0) System.Runtime.InteropServices.Marshal.ThrowExceptionForHR(hr);
                if (automatic) automaticSaves++; else manualSaves++;
                run.Report["manualSaveCount"] = manualSaves;
                run.Report["automaticSaveCount"] = automaticSaves;
                status.Text = "已保存 " + DateTime.Now.ToString("HH:mm:ss");
                File.WriteAllText(Path.Combine(run.Root, "manual-status.json"),
                    new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(new Dictionary<string, object> {
                        { "manualSaves", manualSaves }, { "automaticSaves", automaticSaves }, { "savedUtc", DateTime.UtcNow.ToString("o") } }));
                return true;
            }
            catch (Exception error)
            {
                periodic.Checked = false;
                status.Text = "保存失败，窗口保留，请重试";
                run.Report["lastManualSaveError"] = "0x" + System.Runtime.InteropServices.Marshal.GetHRForException(error).ToString("X8");
                return false;
            }
            finally { saving = false; }
        };
        Action end = delegate {
            if (!finishing && saveDocument(false)) { finishing = true; done.TrySetResult(true); }
        };
        save.Click += delegate { saveDocument(false); };
        finish.Click += delegate { end(); };
        FormClosingEventHandler closing = delegate(object sender, FormClosingEventArgs e) {
            if (!finishing) { e.Cancel = true; end(); }
        };
        form.FormClosing += closing;
        using (var timer = new Timer { Interval = 10000 })
        {
            timer.Tick += delegate { if (periodic.Checked) saveDocument(true); };
            timer.Start();
            run.Report["manualMode"] = true;
            run.Stage("manual-ready");
            var pending = Path.Combine(run.Root, "manual-ready.pending");
            File.WriteAllText(pending, new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(run.Report));
            File.Move(pending, Path.Combine(run.Root, "manual-ready.json"));
            try { await done.Task; }
            finally
            {
                timer.Stop();
                form.FormClosing -= closing;
                form.Resize -= resized;
                form.Controls.Remove(bar);
                bar.Dispose();
            }
        }
        object content = ((dynamic)document).Content;
        try { return ProbeRun.TextFingerprint(Convert.ToString(((dynamic)content).Text)); }
        finally { ProbeRun.Release(content); }
    }
}
