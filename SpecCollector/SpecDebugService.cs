using System;
using System.Threading;
using System.Windows.Forms;

namespace SpecCollector
{
    public static class SpecDebugService
    {
        private static DebugProgressForm _form;
        private static Thread _thread;
        private static readonly object _lock = new object();

        public static void Show()
        {
            lock (_lock)
            {
                if (_thread != null && _thread.IsAlive) return;

                _thread = new Thread(() =>
                {
                    _form = new DebugProgressForm();
                    Application.Run(_form);
                });

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.IsBackground = true;
                _thread.Start();

                while (_form == null || !_form.IsHandleCreated) Thread.Sleep(10);
            }
        }

        public static void Log(string message)
        {
            if (_form == null || _form.IsDisposed) Show();
            _form.AddMessage(message);
        }

        public static void SetProgress(int percent)
        {
            if (_form == null || _form.IsDisposed) Show();
            _form.SetProgress(percent);
        }

        public static void Close()
        {
            if (_form != null && !_form.IsDisposed)
            {
                _form.Invoke(new Action(() => _form.Close()));
            }
        }
    }
}