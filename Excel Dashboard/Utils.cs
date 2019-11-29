using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;

namespace Excel_Dashboard
{
    public static class Utils
    {
        public static Font HEADER_FONT = new System.Drawing.Font("Arial Narrow", 12, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        public static Font CONTENT_FONT = new System.Drawing.Font("Arial Narrow", 12, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        public static Color HEADER_COLOR = System.Drawing.Color.Red;
        public static Color CONTENT_COLOR = System.Drawing.Color.Black;
        public static int COL_NUMBERS = 8;
        public static int ROW_HEIGHT = 40;
    }
}
