using System;
using System.Windows.Forms;

namespace MyDebugLibrary
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

        public void AddMessage(string message)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => AddMessage(message)));
                return;
            }

            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
            // Автоматическая прокрутка к последней строке
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

            // Ограничение 0-100, чтобы не "уронить" ProgressBar
            int val = Math.Max(0, Math.Min(100, percent));
            progressBar.Value = val;
        }

    }
}
