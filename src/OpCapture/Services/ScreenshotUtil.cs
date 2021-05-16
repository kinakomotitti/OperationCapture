using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using OpCapture.Util;

namespace OpCapture.Services
{
    public class ScreenshotService
    {
        public bool TakeScreenshot(string filePath)
        {
            var screenSize = ScreenSizeUtil.GetDisplaySize();
            using var bitmap = new Bitmap(screenSize.Width, screenSize.Height);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.CopyFromScreen(0, 0, 0, 0,
                bitmap.Size, CopyPixelOperation.SourceCopy);
                bitmap.Save(Path.Combine(filePath, "filename.jpg"), ImageFormat.Jpeg);
            }
            return true;
        }
    }
}
