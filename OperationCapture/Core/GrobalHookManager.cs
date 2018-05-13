using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OperationCapture.Core
{
    public  class GrobalHookManager
    {
        /// <summary>
        /// Class declarations
        /// </summary>
        static GrobalHooks.MouseHook mouseHook = new GrobalHooks.MouseHook();
        static GrobalHooks.KeyboardHook keyboardHook = new GrobalHooks.KeyboardHook();

        #region public

        public static void InitializeHooks()
        {
            // register events
            //mouseHook.MouseMove += new GrobalHooks.MouseHook.MouseHookCallback(MouseHook_MouseMove);
            mouseHook.LeftButtonDown += new GrobalHooks.MouseHook.MouseHookCallback(MouseHook_LeftButtonDown);
            //mouseHook.LeftButtonUp += new GrobalHooks.MouseHook.MouseHookCallback(MouseHook_LeftButtonUp);
            mouseHook.RightButtonDown += new GrobalHooks.MouseHook.MouseHookCallback(MouseHook_RightButtonDown);
            //mouseHook.RightButtonUp += new GrobalHooks.MouseHook.MouseHookCallback(MouseHook_RightButtonUp);
            mouseHook.MiddleButtonDown += new GrobalHooks.MouseHook.MouseHookCallback(MouseHook_MiddleButtonDown);
            //mouseHook.MiddleButtonUp += new GrobalHooks.MouseHook.MouseHookCallback(MouseHook_MiddleButtonUp);
            mouseHook.MouseWheel += new GrobalHooks.MouseHook.MouseHookCallback(MouseHook_MouseWheel);

            keyboardHook.KeyDown += new GrobalHooks.KeyboardHook.KeyboardHookCallback(KeyboardHook_KeyDown);
            //keyboardHook.KeyUp += new GrobalHooks.KeyboardHook.KeyboardHookCallback(KeyboardHook_KeyUp);

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

        private static async Task CallCoreMethod(string operation)
        {
            EventCore core = new EventCore();
            core.TakeScreenShot(operation);
        }

        #endregion

        #region Event

        #region MouseHook_MouseWheel

        async static void MouseHook_MouseWheel(GrobalHooks.MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            await CallCoreMethod("Mouse Wheel Move");
        }

        #endregion

        #region MouseHook_MiddleButtonDown

        async static void MouseHook_MiddleButtonDown(GrobalHooks.MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            await CallCoreMethod("Mouse Middle Button Down");
        }

        #endregion

        #region MouseHook_RightButtonDown

        async static void MouseHook_RightButtonDown(GrobalHooks.MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            await CallCoreMethod("Mouse Right Button Down");
        }

        #endregion

        #region MouseHook_LeftButtonDown

        async static void MouseHook_LeftButtonDown(GrobalHooks.MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            await CallCoreMethod("Mouse Left Button Down");
        }

        #endregion

        #region KeyboardHook_KeyDown

        async static void KeyboardHook_KeyDown(GrobalHooks.KeyboardHook.VKeys key)
        {
            await CallCoreMethod(key.ToString());
            if (key == GrobalHooks.KeyboardHook.VKeys.ESCAPE)
            {
                System.Environment.Exit(0);
                //this.ClosingProcess();
            }
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
