using OpenCvSharp;
using OpenCvSharp.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace OpenCVSharpSample
{
    class Program
    {
        static void Main(string[] args)
        {
            // 画像１を読み込む
            using (Mat image1 = Cv2.ImRead("test01.jpeg"))
            // 画像２を読み込む
            using (Mat image2 = Cv2.ImRead("test02.jpeg"))
            // 差分画像を保存する領域を確保する
            using (Mat diff = new Mat(new OpenCvSharp.Size(image1.Cols, image1.Rows), MatType.CV_8UC3))
            {
                // 画像１と画像２の差分をとる
                Cv2.Absdiff(image1, image2, diff);
                // BitmapSourceConverterを利用するとMatをBitmapSourceに変換できる
                var bitmap = BitmapConverter.ToBitmap(diff);
                // Sourceに画像を割り当てる
                bitmap.Save("result.jpeg");

                Mat merged = new Mat();
                Cv2.Add(image2, diff, merged);
                
                merged.ToBitmap().Save("result2.jpeg");
                
            }
        }
    }
}
