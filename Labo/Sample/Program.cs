using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            var cellHeight = 24;
            var row = 1;
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                var imagePath = @"./Resource/image.jpg";
                var image = ws.AddPicture(imagePath).MoveTo(ws.Cell($"A{row}").Address);
                row += (int)(image.Height / cellHeight)+2;
                ws.Cell($"A{row}").Value = row.ToString();
                wb.SaveAs("file.xlsx");
            }
        }

    }
}
