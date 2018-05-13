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

namespace OperationCapture.Core
{
    public class EventCore
    {
        #region 変数/定数

        #region Closed
        private XLWorkbook workBook;
        private IXLWorksheet workSheet;
        #endregion

        #region Open
        private static SpreadsheetDocument document;
        private static WorkbookPart wbpart;
        private static WorksheetPart wspart;
        private static SheetData sheetData;
        private static Sheets sheets;
        private static Sheet sheet;
        #endregion

        #region 共通

        private int ActiveCellRow = 1;
        private long ActiveImageCellRow = 230000L;
        private long CellRow = 230000L;
        private static int SheetId = 1;
        public static Dictionary<string, string> Operations = new Dictionary<string, string>();

        #endregion

        #endregion

        #region プロパティ

        public string BookName { get; set; } = "Default.xlsx";

        public string SheetName
        {
            get
            {
                return $"DefaultSheet{SheetId}";
            }
        }

        public bool HasSheet
        {
            get
            {
                return SheetId > 1;
            }

        }

        #endregion

        #region Enum

        public enum Oss
        {
            Interop = 99,
            Open = 0,
            Closed = 1
        }

        public enum ScreenshotMode
        {
            FullWindow = 0,
            ActiveWindow = 1
        }

        #endregion

        #region コンストラクタ

        public EventCore(Oss param = Oss.Open)
        {
            //this.InitializeThisClassCore(param);
        }

        #endregion

        #region public

        #region SaveEvidenceToExcelFile

        public void SaveEvidenceToExcelFile()
        {
            // スプレッドシートドキュメントを作成
            document = SpreadsheetDocument.Create(this.BookName, SpreadsheetDocumentType.Workbook);

            // WorkbookPart をドキュメントに追加
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // WorksheetPart を WorkbookPart に追加
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            // Sheets をワークブックに追加
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());

            // ワークシートをワークブックに追加
            var sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = DateTime.Now.ToString("yyyyMMdd_hhmmss"),
            };
            sheets.Append(sheet);

            foreach (var item in Operations)
            {
                InsertText(item.Value, "A", ActiveCellRow.ToString());
                var dt = InsertImage(worksheetPart.Worksheet, 520000L, ActiveImageCellRow, null, null, item.Key);
                ActiveImageCellRow += dt;

                //4行分下げてみる
                ActiveImageCellRow += CellRow;
                ActiveImageCellRow += CellRow;
                ActiveImageCellRow += CellRow;
                ActiveCellRow += (int)(dt / CellRow) + 4; //43行分
            }

            // ワークブックを保存
            worksheetPart.Worksheet.Save();

            // ドキュメントを閉じる
            document.Close();
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
        /// 
        /// </summary>
        public void TakeScreenShot()
        {
            string dirPath = ".\\screenshot\\";
            this.CreateDirectory(dirPath);
            string shotName = $"{dirPath}\\{DateTime.Now.ToString("yyyyMMdd_hhmmss")}.jpeg";
            this.ScreenShot(shotName);
        }
        public void TakeScreenShot(string operation,string dirPath = ".\\screenshot\\")
        {
            this.CreateDirectory(dirPath);
            string shotName = $"{dirPath}\\{DateTime.Now.ToString("yyyyMMdd_HHmmss.fffffff")}.jpeg";
            this.ScreenShot(shotName);
            EventCore.Operations.Add(shotName, operation);
        }

        #endregion

        #endregion

        #region private

        #region InsertText

        // Given a document name and text, 
        // inserts a new work sheet and writes the text to cell "A1" of the new worksheet.

        public void InsertText(string text, string positionX = "A", string positionY = "1")
        {
            // Get the SharedStringTablePart. If it does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = document.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            // Insert the text into the SharedStringTablePart.
            int index = InsertSharedStringItem(text, shareStringPart);

            // Insert a new worksheet.
            //WorksheetPart worksheetPart = InsertWorksheet(document.WorkbookPart);
            WorksheetPart worksheetPart = document.WorkbookPart.WorksheetParts.First();

            // Insert cell A1 into the new worksheet.
            Cell cell = InsertCellInWorksheet(positionX, uint.Parse(positionY), worksheetPart);

            // Set the value of cell A1.
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            // Save the new worksheet.
            worksheetPart.Worksheet.Save();
        }
        #endregion

        #region InsertSharedStringItem

        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        private int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
        #endregion

        #region InsertCellInWorksheet
        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        #endregion

        #region InsertImage :画像の埋め込み

        protected static long InsertImage(Worksheet ws, long x, long y, long? width, long? height, string sImagePath)
        {
            try
            {
                WorksheetPart wsp = ws.WorksheetPart;
                DrawingsPart dp;
                ImagePart imgp;
                WorksheetDrawing wsd;

                ImagePartType ipt;
                switch (sImagePath.Substring(sImagePath.LastIndexOf('.') + 1).ToLower())
                {
                    case "png":
                        ipt = ImagePartType.Png;
                        break;
                    case "jpg":
                    case "jpeg":
                        ipt = ImagePartType.Jpeg;
                        break;
                    case "gif":
                        ipt = ImagePartType.Gif;
                        break;
                    default:
                        return 0;
                }

                if (wsp.DrawingsPart == null)
                {
                    //----- no drawing part exists, add a new one

                    dp = wsp.AddNewPart<DrawingsPart>();
                    imgp = dp.AddImagePart(ipt, wsp.GetIdOfPart(dp));
                    wsd = new WorksheetDrawing();
                }
                else
                {
                    //----- use existing drawing part

                    dp = wsp.DrawingsPart;
                    imgp = dp.AddImagePart(ipt);
                    dp.CreateRelationshipToPart(imgp);
                    wsd = dp.WorksheetDrawing;
                }

                using (FileStream fs = new FileStream(sImagePath, FileMode.Open))
                {
                    imgp.FeedData(fs);
                }

                int imageNumber = dp.ImageParts.Count<ImagePart>();
                if (imageNumber == 1)
                {
                    Drawing drawing = new Drawing();
                    drawing.Id = dp.GetIdOfPart(imgp);
                    ws.Append(drawing);
                }

                DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties nvdp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties();
                nvdp.Id = new UInt32Value((uint)(1024 + imageNumber));
                nvdp.Name = "Picture " + imageNumber.ToString();
                nvdp.Description = "";
                DocumentFormat.OpenXml.Drawing.PictureLocks picLocks = new DocumentFormat.OpenXml.Drawing.PictureLocks();
                picLocks.NoChangeAspect = true;
                picLocks.NoChangeArrowheads = true;
                DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties nvpdp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties();
                nvpdp.PictureLocks = picLocks;
                DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties nvpp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties();
                nvpp.NonVisualDrawingProperties = nvdp;
                nvpp.NonVisualPictureDrawingProperties = nvpdp;

                DocumentFormat.OpenXml.Drawing.Stretch stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
                stretch.FillRectangle = new DocumentFormat.OpenXml.Drawing.FillRectangle();

                DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill blipFill = new DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill();
                DocumentFormat.OpenXml.Drawing.Blip blip = new DocumentFormat.OpenXml.Drawing.Blip();
                blip.Embed = dp.GetIdOfPart(imgp);
                blip.CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print;
                blipFill.Blip = blip;
                blipFill.SourceRectangle = new DocumentFormat.OpenXml.Drawing.SourceRectangle();
                blipFill.Append(stretch);

                DocumentFormat.OpenXml.Drawing.Transform2D t2d = new DocumentFormat.OpenXml.Drawing.Transform2D();
                DocumentFormat.OpenXml.Drawing.Offset offset = new DocumentFormat.OpenXml.Drawing.Offset();
                offset.X = 0;
                offset.Y = 0;
                t2d.Offset = offset;
                Bitmap bm = new Bitmap(sImagePath);

                DocumentFormat.OpenXml.Drawing.Extents extents = new DocumentFormat.OpenXml.Drawing.Extents();

                if (width == null)
                    extents.Cx = (long)bm.Width * (long)((float)914400 / bm.HorizontalResolution);
                else
                    extents.Cx = width;

                if (height == null)
                    extents.Cy = (long)bm.Height * (long)((float)914400 / bm.VerticalResolution);
                else
                    extents.Cy = height;

                bm.Dispose();
                t2d.Extents = extents;
                DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties sp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties();
                sp.BlackWhiteMode = DocumentFormat.OpenXml.Drawing.BlackWhiteModeValues.Auto;
                sp.Transform2D = t2d;
                DocumentFormat.OpenXml.Drawing.PresetGeometry prstGeom = new DocumentFormat.OpenXml.Drawing.PresetGeometry();
                prstGeom.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle;
                prstGeom.AdjustValueList = new DocumentFormat.OpenXml.Drawing.AdjustValueList();
                sp.Append(prstGeom);
                sp.Append(new DocumentFormat.OpenXml.Drawing.NoFill());

                DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture picture = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture();
                picture.NonVisualPictureProperties = nvpp;
                picture.BlipFill = blipFill;
                picture.ShapeProperties = sp;

                DocumentFormat.OpenXml.Drawing.Spreadsheet.Position pos = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Position();
                pos.X = x;
                pos.Y = y;
                Extent ext = new Extent();
                ext.Cx = extents.Cx;
                ext.Cy = extents.Cy;
                AbsoluteAnchor anchor = new AbsoluteAnchor();
                anchor.Position = pos;
                anchor.Extent = ext;
                anchor.Append(picture);
                anchor.Append(new ClientData());
                wsd.Append(anchor);
                wsd.Save(dp);
                return extents.Cy;
            }
            catch (Exception ex)
            {
                throw ex; // or do something more interesting if you want
            }
        }
        #endregion

        #region InitializeThisClassCore　:　コンストラクタの処理の中核（WorkBookの作成）

        private void InitializeThisClassCore(Oss param = Oss.Open)
        {
            if (param == Oss.Closed)
            {
                //Closed
                workBook = new XLWorkbook();
            }
            else if (param == Oss.Open)
            {
                //Open
                // 新しいxlsxドキュメントを作成
                document = SpreadsheetDocument.Create(this.BookName, SpreadsheetDocumentType.Workbook, true);

                // ドキュメントのワークブックパートに、ワークブックを設定
                wbpart = document.AddWorkbookPart();
                wbpart.Workbook = new Workbook();
            }
        }

        #endregion

        #region スクリーンショット

        private const int SRCCOPY = 13369376;
        private const int CAPTUREBLT = 1073741824;

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hwnd);

        [DllImport("gdi32.dll")]
        private static extern int BitBlt(IntPtr hDestDC,
            int x,
            int y,
            int nWidth,
            int nHeight,
            IntPtr hSrcDC,
            int xSrc,
            int ySrc,
            int dwRop);


        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        private struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        [DllImport("User32.Dll")]
        static extern int GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll")]
        extern static IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr ReleaseDC(IntPtr hwnd, IntPtr hdc);

        private void ScreenShot(string filePath, ScreenshotMode shotMode = ScreenshotMode.FullWindow)
        {
            Bitmap bmp;
            System.Drawing.Rectangle rect;
            if (shotMode == ScreenshotMode.FullWindow)
            {
                //フルスクリーンモード
                rect = System.Windows.Forms.Screen.PrimaryScreen.Bounds;
                bmp = new Bitmap(rect.Width, rect.Height, PixelFormat.Format32bppArgb);
            }
            else
            {
                //アクティブWindowモード
                RECT r;
                IntPtr active = GetForegroundWindow();
                GetWindowRect(active, out r);
                rect = new System.Drawing.Rectangle(r.left, r.top, r.right - r.left, r.bottom - r.top);
                bmp = new Bitmap(rect.Width, rect.Height, PixelFormat.Format32bppArgb);
            }
            //Bitmapの作成
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(rect.X, rect.Y, 0, 0, rect.Size, CopyPixelOperation.SourceCopy);
            }

            bmp.Save(filePath, ImageFormat.Jpeg);

        }

        #endregion

        #endregion
    }
}
