using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OperationCapture.Core
{
    public class GlobalHookManager
    {
        /// <summary>
        /// Class declarations
        /// </summary>
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

        #region CallCoreMethod

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

        #region MouseHook_MouseWheel

        static void MouseHook_MouseWheel(GlobalHooks.MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            CallCoreMethod("Mouse Wheel Move");
        }

        #endregion

        #region MouseHook_MiddleButtonDown

        static void MouseHook_MiddleButtonDown(GlobalHooks.MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            CallCoreMethod("Mouse Middle Button Down");
        }

        #endregion

        #region MouseHook_RightButtonDown

        static void MouseHook_RightButtonDown(GlobalHooks.MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            CallCoreMethod("Mouse Right Button Down");
        }

        #endregion

        #region MouseHook_LeftButtonDown

        static void MouseHook_LeftButtonDown(GlobalHooks.MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            CallCoreMethod("Mouse Left Button Down");
        }

        #endregion

        #region KeyboardHook_KeyDown

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

        #endregion

        #region Window_Closing:Window閉じるイベント

        private static void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            //Hookの撤収
            mouseHook.Uninstall();
            keyboardHook.Uninstall();
        }

        #endregion

        #endregion


    }
}
