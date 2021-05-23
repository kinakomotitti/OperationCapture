namespace OperationCapture.Core
{
    #region using
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    #endregion

    public class GlobalHookManager
    {
        static GlobalHooks.MouseHook mouseHook = new GlobalHooks.MouseHook();
        static GlobalHooks.KeyboardHook keyboardHook = new GlobalHooks.KeyboardHook();

        #region public

        public static void InitializeHooks()
        {
            // register events
            //mouseHook.MouseMove += new GlobalHooks.MouseHook.MouseHookCallback(MouseHook_MouseMove);
            mouseHook.LeftButtonDown += new GlobalHooks.MouseHook.MouseHookCallback(MouseHook_LeftButtonDown);
            //mouseHook.LeftButtonUp += new GlobalHooks.MouseHook.MouseHookCallback(MouseHook_LeftButtonUp);
            mouseHook.RightButtonDown += new GlobalHooks.MouseHook.MouseHookCallback(MouseHook_RightButtonDown);
            //mouseHook.RightButtonUp += new GlobalHooks.MouseHook.MouseHookCallback(MouseHook_RightButtonUp);
            mouseHook.MiddleButtonDown += new GlobalHooks.MouseHook.MouseHookCallback(MouseHook_MiddleButtonDown);
            //mouseHook.MiddleButtonUp += new GlobalHooks.MouseHook.MouseHookCallback(MouseHook_MiddleButtonUp);
            mouseHook.MouseWheel += new GlobalHooks.MouseHook.MouseHookCallback(MouseHook_MouseWheel);

            keyboardHook.KeyDown += new GlobalHooks.KeyboardHook.KeyboardHookCallback(KeyboardHook_KeyDown);
            //keyboardHook.KeyUp += new GlobalHooks.KeyboardHook.KeyboardHookCallback(KeyboardHook_KeyUp);

            //開始
            mouseHook.Install();
            keyboardHook.Install();
        }

        public static void FinalizeHooks()
        {
            //開始
            mouseHook.Uninstall();
            keyboardHook.Uninstall();
        }

        #endregion

        #region private

        private static void CallCoreMethod(string operation)
        {
            Task.Run(() =>
            {
                EventCore core = new EventCore();
                core.TakeScreenShot(operation);
            });
        }

        #endregion

        #region Event

        static void MouseHook_MouseWheel(GlobalHooks.MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            CallCoreMethod("Mouse Wheel Move");
        }

        static void MouseHook_MiddleButtonDown(GlobalHooks.MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            CallCoreMethod("Mouse Middle Button Down");
        }

        static void MouseHook_RightButtonDown(GlobalHooks.MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            CallCoreMethod("Mouse Right Button Down");
        }

        static void MouseHook_LeftButtonDown(GlobalHooks.MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            CallCoreMethod("Mouse Left Button Down");
        }

        static void KeyboardHook_KeyDown(GlobalHooks.KeyboardHook.VKeys key)
        {

            if (key == GlobalHooks.KeyboardHook.VKeys.ESCAPE)
            {
                System.Environment.Exit(0);
                //this.ClosingProcess();
            }
            //鬱陶しいからEnter以外は無視しようと思います。
            else if (key != GlobalHooks.KeyboardHook.VKeys.RETURN)
            {
                //No Action
                return;
            }
            CallCoreMethod(key.ToString());
        }

        private static void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            //Hookの撤収
            mouseHook.Uninstall();
            keyboardHook.Uninstall();
        }

        #endregion
    }
}
