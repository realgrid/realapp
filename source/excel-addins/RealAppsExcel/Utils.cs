using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RealAppsExcel
{
    class Utils
    {
        public static void ShowMessage(string message)
        {
            System.Windows.Forms.MessageBox.Show(message);
        }

        public static string Indent(int spaces, int tabcount)
        {
            return new string(' ', spaces * tabcount);
        }

        public static int WidthToPixel(double width)
        {

            //Range.Width -> Pixel : Range.Width / 0.75
            return Convert.ToInt32(Math.Round(width / 0.75));
            //return Convert.ToInt32(Math.Round(width * 96.0 / 72.0));
        }
        public static double PixelToColumnWidth(int pixel)
        {
            //Pixel -> ColumnWidth : Pixel * 0.076875 + (Pixel - 13) * 0.048135
            return Math.Round(pixel * 0.076875 + (pixel - 13) * 0.048135, 2);
            //return Math.Round(width / 96.0 * 72.0, 2);
        }
    }
}
