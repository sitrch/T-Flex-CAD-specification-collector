using System;
using System.Windows.Forms;

namespace SpecCollector
{
    public partial class DebugProgressForm : Form
    {
        private TextBox txtLog;
        private ProgressBar progressBar;

        public DebugProgressForm()
        {
            this.Text = "Debug Monitor";
            this.Size = new System.Drawing.Size(400, 300);

            progressBar = new ProgressBar { Dock = DockStyle.Top, Height = 20 };
            txtLog = new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };

            this.Controls.Add(txtLog);
            this.Controls.Add(progressBar);
        }

        public void AppendLog(string message)
        {
            if (this.InvokeRequired)
            {
                 this.BeginInvoke(new Action(() => AppendLog(message)));
                return;
            }

            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
        }

        public void SetProgress(int percent)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => SetProgress(percent)));
                return;
            }

            int val = Math.Max(0, Math.Min(100, percent));
            progressBar.Value = val;
        }
    }
}