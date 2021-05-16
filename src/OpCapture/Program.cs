using Gma.System.MouseKeyHook;
using OpCapture.Services;
using OpCapture.Util;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpCapture
{
    class Program
    {
        private IKeyboardMouseEvents m_GlobalHook;

        static void Main(string[] args)
        {
            var util = new ScreenshotService().TakeScreenshot(@".\");
            var program = new Program();
            var result = Task.Run(() =>
            {
                program.InternalMain();
                while (true)
                {
                    Task.Delay(1000);
                }
            });
            result.Wait();
            
        }
        internal void InternalMain()
        {
            this.Subscribe();
        }

        public void Subscribe()
        {
            // Note: for the application hook, use the Hook.AppEvents() instead
            m_GlobalHook = Hook.GlobalEvents();

            m_GlobalHook.MouseDownExt += GlobalHookMouseDownExt;
            m_GlobalHook.KeyPress += GlobalHookKeyPress;
        }

        private void GlobalHookKeyPress(object sender, KeyPressEventArgs e)
        {
            Console.WriteLine("KeyPress: \t{0}", e.KeyChar);
        }

        private void GlobalHookMouseDownExt(object sender, MouseEventExtArgs e)
        {
            Console.WriteLine("MouseDown: \t{0}; \t System Timestamp: \t{1}", e.Button, e.Timestamp);

            // uncommenting the following line will suppress the middle mouse button click
            // if (e.Buttons == MouseButtons.Middle) { e.Handled = true; }
        }

        public void Unsubscribe()
        {
            m_GlobalHook.MouseDownExt -= GlobalHookMouseDownExt;
            m_GlobalHook.KeyPress -= GlobalHookKeyPress;

            //It is recommened to dispose it
            m_GlobalHook.Dispose();
        }
    }
}
