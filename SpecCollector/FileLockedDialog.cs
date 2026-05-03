using System;
using System.Windows.Forms;

namespace SpecCollector
{
    public partial class FileLockedDialog : Form
    {
        public enum FileAction
        {
            Retry,
            SaveAs,
            Cancel
        }

        public FileAction Result { get; private set; }

        public FileLockedDialog(string filePath)
        {
            InitializeComponent();
            labelMessage.Text = $"Файл занят другим приложением:\n{filePath}\n\nЗакройте файл в другой программе или выберите действие:";
        }

        private void InitializeComponent()
        {
            this.labelMessage = new System.Windows.Forms.Label();
            this.btnRetry = new System.Windows.Forms.Button();
            this.btnSaveAs = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // labelMessage
            // 
            this.labelMessage.AutoSize = true;
            this.labelMessage.Location = new System.Drawing.Point(12, 9);
            this.labelMessage.MaximumSize = new System.Drawing.Size(350, 0);
            this.labelMessage.Name = "labelMessage";
            this.labelMessage.Size = new System.Drawing.Size(350, 60);
            this.labelMessage.TabIndex = 0;
            // 
            // btnRetry
            // 
            this.btnRetry.Location = new System.Drawing.Point(12, 80);
            this.btnRetry.Name = "btnRetry";
            this.btnRetry.Size = new System.Drawing.Size(110, 30);
            this.btnRetry.TabIndex = 1;
            this.btnRetry.Text = "Повторить";
            this.btnRetry.UseVisualStyleBackColor = true;
            this.btnRetry.Click += new System.EventHandler(this.btnRetry_Click);
            // 
            // btnSaveAs
            // 
            this.btnSaveAs.Location = new System.Drawing.Point(128, 80);
            this.btnSaveAs.Name = "btnSaveAs";
            this.btnSaveAs.Size = new System.Drawing.Size(140, 30);
            this.btnSaveAs.TabIndex = 2;
            this.btnSaveAs.Text = "Сохранить под другим именем";
            this.btnSaveAs.UseVisualStyleBackColor = true;
            this.btnSaveAs.Click += new System.EventHandler(this.btnSaveAs_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(274, 80);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(88, 30);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "Отмена";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // FileLockedDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(374, 120);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSaveAs);
            this.Controls.Add(this.btnRetry);
            this.Controls.Add(this.labelMessage);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FileLockedDialog";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Файл занят";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private System.Windows.Forms.Label labelMessage;
        private System.Windows.Forms.Button btnRetry;
        private System.Windows.Forms.Button btnSaveAs;
        private System.Windows.Forms.Button btnCancel;

        private void btnRetry_Click(object sender, EventArgs e)
        {
            Result = FileAction.Retry;
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void btnSaveAs_Click(object sender, EventArgs e)
        {
            Result = FileAction.SaveAs;
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Result = FileAction.Cancel;
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }
    }
}
