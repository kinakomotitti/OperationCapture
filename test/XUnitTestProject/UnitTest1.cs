using AnimatedGif;
using System;
using System.Drawing;
using System.IO;
using Xunit;

namespace XUnitTestProject
{
    public class UnitTest1
    {
        [Fact]
        public void Test1()
        {
            using (var gif = AnimatedGif.AnimatedGif.Create(@"C:\Users\Kinakomotitti\Desktop\Screenshot\mygif.gif", 1000))
            {
                var list = Directory.GetFiles(@"C:\Users\Kinakomotitti\Desktop\Screenshot");
                foreach (var item in list)
                {
                    if (new FileInfo(item).Extension == ".jpeg")
                    {
                        var img = Image.FromFile(item);
                        gif.AddFrame(img, delay: -1, quality: GifQuality.Bit8);
                    }
                }
            }
        }
    }
}
