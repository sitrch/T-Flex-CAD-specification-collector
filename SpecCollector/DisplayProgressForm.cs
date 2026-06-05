using System;
using System.Drawing;
using System.Windows.Forms;

namespace SpecCollector
{
    public class DisplayProgressForm : Form
    {
        private TextBox txtLog;
        private ProgressBar progressBar;
        private Label lblStatus;

        public DisplayProgressForm()
        {
            this.Text = "Собиратель — монитор";
            this.Size = new Size(650, 450);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;

            progressBar = new ProgressBar { Dock = DockStyle.Top, Height = 20 };
            lblStatus = new Label { Dock = DockStyle.Top, Height = 20, Text = "Ожидание...", TextAlign = ContentAlignment.MiddleLeft };
            txtLog = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9f)
            };

            this.Controls.Add(txtLog);
            this.Controls.Add(lblStatus);
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
            lblStatus.Text = $"Прогресс: {val}%";
        }
    }
}