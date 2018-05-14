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
            var cellHeight = 18;
            var row = 1;
            using (var wb = new XLWorkbook())
            {
                //ワークシートの設定
                var ws = wb.AddWorksheet("Sheet1");

                //貼り付ける画像を指定
                var imagePath = @"./Resource/image.jpg";

                //タイトルとしてA１セルに文字を出力
                ws.Cell($"A{row}").Value = "これ食べたい";
                row++;

                //ワークシートの変数のAddPictureメソッドを呼び出す
                //左端がB２のセルになるように画像を移動する
                var image = ws.AddPicture(imagePath).MoveTo(ws.Cell($"B{row}").Address);

                //画像の左端のセル番号を計算してみる・・・
                row += (int)(image.Height / cellHeight);

                //締めの一言を画像にかぶらないように出力する
                ws.Cell($"A{row}").Value = "以上です。";

                wb.SaveAs("file.xlsx");
            }
        }
    }
}
