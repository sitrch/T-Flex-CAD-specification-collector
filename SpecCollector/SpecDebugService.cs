using System;
using System.Windows.Forms;

namespace SpecCollector
{
    public static class SpecDebugService
    {
        private static DebugProgressForm _debugForm;

        public static void Show()
        {
            if (_debugForm == null || _debugForm.IsDisposed)
            {
                _debugForm = new DebugProgressForm();
                _debugForm.Show();
            }
            else
            {
                _debugForm.Activate();
            }
        }

        public static void Log(string message)
        {
            if (_debugForm != null && !_debugForm.IsDisposed)
            {
                _debugForm.AppendLog(message);
            }
        }
    }
}
