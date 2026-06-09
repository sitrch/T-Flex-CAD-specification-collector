using System;
using System.Threading;
using System.Windows.Forms;

namespace SpecCollector
{
    public static class ___DisplayService
    {
        private static DisplayProgressForm _form;
        private static Thread _thread;
        private static readonly object _lock = new object();

        public static void Show()
        {
            lock (_lock)
            {
                if (_thread != null && _thread.IsAlive) return;

                _form = null;

                _thread = new Thread(() =>
                {
                    _form = new DisplayProgressForm();
                    Application.Run(_form);
                });

                _thread.SetApartmentState(ApartmentState.STA);
                _thread.IsBackground = true;
                _thread.Start();

                // Ждём инициализации формы
                while (_form == null || !_form.IsHandleCreated)
                    Thread.Sleep(10);
            }
        }

        public static void Log(string message)
        {
            if (_form == null || _form.IsDisposed) Show();
            try { _form.AddMessage(message); } catch { /* ignored: форма могла быть закрыта */ }
        }

        public static void SetProgress(int percent)
        {
            if (_form == null || _form.IsDisposed) Show();
            try { _form.SetProgress(percent); } catch { /* ignored: форма могла быть закрыта */ }
        }

        public static void Close()
        {
            if (_form != null && !_form.IsDisposed)
            {
                try
                {
                    _form.Invoke(new Action(() => _form.Close()));
                }
                catch { /* ignored: форма могла быть закрыта */ }
            }
        }
    }
}