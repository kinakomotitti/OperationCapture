using Capture.GlobalHook;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Capture
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        static MouseHook mouseHook = new MouseHook();
        static KeyboardHook keyboardHook = new KeyboardHook();

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            // register events
            mouseHook.LeftButtonDown += new MouseHook.MouseHookCallback(MouseHook_LeftButtonDown);
            mouseHook.RightButtonDown += new MouseHook.MouseHookCallback(MouseHook_RightButtonDown);
            mouseHook.MiddleButtonDown += new MouseHook.MouseHookCallback(MouseHook_MiddleButtonDown);
            mouseHook.MouseWheel += new MouseHook.MouseHookCallback(MouseHook_MouseWheel);
            keyboardHook.KeyDown += new KeyboardHook.KeyboardHookCallback(KeyboardHook_KeyDown);

            //開始
            mouseHook.Install();
            keyboardHook.Install();

            while (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                await Task.Delay(10000, stoppingToken);
            }
        }

        #region Event

        #region MouseHook_MouseWheel

        static void MouseHook_MouseWheel(MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            Console.WriteLine("Mouse Wheel Move");
        }

        #endregion

        #region MouseHook_MiddleButtonDown

        static void MouseHook_MiddleButtonDown(MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            Console.WriteLine("Mouse Middle Button Down");
        }

        #endregion

        #region MouseHook_RightButtonDown

        static void MouseHook_RightButtonDown(MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            Console.WriteLine("Mouse Right Button Down");
        }

        #endregion

        #region MouseHook_LeftButtonDown

        static void MouseHook_LeftButtonDown(MouseHook.MSLLHOOKSTRUCT mouseStruct)
        {
            Console.WriteLine("Mouse Left Button Down");
        }

        #endregion

        #region KeyboardHook_KeyDown

        static void KeyboardHook_KeyDown(KeyboardHook.VKeys key)
        {

            if (key == KeyboardHook.VKeys.ESCAPE)
            {
                System.Environment.Exit(0);
            }
            else if (key != KeyboardHook.VKeys.RETURN)
            {
                return;
            }
            Console.WriteLine(key.ToString());
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
