using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

namespace OperationCapture.Core
{
    public class EventCore
    {
        #region 変数/定数

        private int ActiveCellRow = 1;
        private long CellRow = 24;
        private static int SheetId = 1;
        public static Dictionary<string, string> Operations = new Dictionary<string, string>();

        #endregion

        #region public

        #region SaveEvidenceToExcelFile :　エクセルに画像データを保存

        /// <summary>
        /// ClosedXMLを利用してExcelに画像を貼り付けて保存
        /// </summary>
        public void SaveEvidenceToExcelFile()
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet(DateTime.Now.ToString("yyyyMMdd_hhmmss"));
                ws.Cell($"A{ActiveCellRow.ToString()}").Value = "操作開始！";
                ActiveCellRow++;
                foreach (var item in Operations)
                {
                    var image = ws.AddPicture(item.Key).MoveTo(ws.Cell($"B{ActiveCellRow}").Address);
                    ActiveCellRow += (int)(image.Height / CellRow)+1;
                    ws.Cell($"A{ActiveCellRow.ToString()}").Value = $"↑の画像は、{item.Value} した結果です。";
                    ActiveCellRow++;
                }
                wb.SaveAs("Created.xlsx");
            }
        }

        #endregion

        #region CreateDirectory : ファイル格納フォルダ作成

        public void CreateDirectory(string dirPath)
        {
            if (Directory.Exists(dirPath) == false)
            {
                Directory.CreateDirectory(dirPath);
            }
        }

        #endregion

        #region TakeScreenShot   :   スクリーンショットの取得

        /// <summary>
        /// StackOverFlowのAnswerからコア部分のロジックを拝借しています
        /// </summary>
        public void TakeScreenShot()
        {
            string dirPath = ".\\screenshot\\";
            this.CreateDirectory(dirPath);
            string shotName = $"{dirPath}\\{DateTime.Now.ToString("yyyyMMdd_hhmmss")}.jpeg";
            //this.ScreenShot(shotName);
            this.CaptureScreen(true, shotName);
        }
        public void TakeScreenShot(string operation, string dirPath = ".\\screenshot\\")
        {
            this.CreateDirectory(dirPath);
            string shotName = $"{dirPath}\\{DateTime.Now.ToString("yyyyMMdd_HHmmss.fffffff")}.jpeg";
            //this.ScreenShot(shotName);
            this.CaptureScreen(true, shotName);
            EventCore.Operations.Add(shotName, operation);
        }

        #endregion

        #endregion

        #region private

        #region スクリーンショット
        //https://stackoverflow.com/questions/6750056/how-to-capture-the-screen-and-mouse-pointer-using-windows-apis
        //より拝借（一部修正）
        [StructLayout(LayoutKind.Sequential)]
        struct CURSORINFO
        {
            public Int32 cbSize;
            public Int32 flags;
            public IntPtr hCursor;
            public POINTAPI ptScreenPos;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct POINTAPI
        {
            public int x;
            public int y;
        }

        [DllImport("user32.dll")]
        static extern bool GetCursorInfo(out CURSORINFO pci);

        [DllImport("user32.dll")]
        static extern bool DrawIcon(IntPtr hDC, int X, int Y, IntPtr hIcon);

        const Int32 CURSOR_SHOWING = 0x00000001;

        private void CaptureScreen(bool CaptureMouse, string filePath)
        {
            Bitmap result = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, PixelFormat.Format24bppRgb);

            try
            {
                using (Graphics g = Graphics.FromImage(result))
                {
                    g.CopyFromScreen(0, 0, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);

                    if (CaptureMouse)
                    {
                        CURSORINFO pci;
                        pci.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(CURSORINFO));

                        if (GetCursorInfo(out pci))
                        {
                            if (pci.flags == CURSOR_SHOWING)
                            {
                                DrawIcon(g.GetHdc(), pci.ptScreenPos.x, pci.ptScreenPos.y, pci.hCursor);
                                g.ReleaseHdc();
                            }
                        }
                    }
                }
                result?.Save(filePath);
            }
            catch
            {
                result = null;
            }
        }
        #endregion

        #endregion
    }
}
