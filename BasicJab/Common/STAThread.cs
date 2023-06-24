using System;
using System.Threading;
using System.Windows.Forms;

namespace BasicJab.Common
{
    public class STAThread : IDisposable
    {
        private Thread _thread;
        private SynchronizationContext _ctx;
        private ManualResetEvent _mre;

        public STAThread()
        {
            using (_mre = new ManualResetEvent(false))
            {
                _thread = new Thread(() => {
                    Application.Idle += Initialize;
                    Application.Run();  //run message loop
                });
                _thread.IsBackground = true;
                _thread.SetApartmentState(ApartmentState.STA);
                _thread.Start();
                _mre.WaitOne();
            }
        }
        public void BeginInvoke(Delegate dlg, params Object[] args)
        {
            if (_ctx == null) throw new ObjectDisposedException("STAThread");
            _ctx.Post((_) => dlg.DynamicInvoke(args), null);
        }
        public object Invoke(Delegate dlg, params Object[] args)
        {
            if (_ctx == null) throw new ObjectDisposedException("STAThread");
            object result = null;
            _ctx.Send((_) => result = dlg.DynamicInvoke(args), null);
            return result;
        }
        protected virtual void Initialize(object sender, EventArgs e)
        {
            _ctx = SynchronizationContext.Current;
            _mre.Set();
            Application.Idle -= Initialize;
        }
        public void Dispose()
        {
            if (_ctx == null) return;
            _ctx.Send((_) => Application.ExitThread(), null);
            _ctx = null;
        }

    }
}
