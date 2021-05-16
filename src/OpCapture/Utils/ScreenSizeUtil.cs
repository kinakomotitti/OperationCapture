using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace OpCapture.Util
{
    public static class ScreenSizeUtil
    {
        public static (int Width, int Height) GetDisplaySize()
        {
            return (Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
        }
    }
}
