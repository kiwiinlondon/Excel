using XL=Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class CellStyling
    {
        public CellStyling(string styleName, Color? backgroundColour, bool isBold, XL.XlHAlign? justification, XL.XlLineStyle? topBorder, XL.XlLineStyle? bottomBorder,bool isItalic)
        {
            StyleName = styleName;
            BackgroundColour = backgroundColour;
            IsBold = isBold;
            Justification = justification;
            TopBorder = topBorder;
            BottomBorder = bottomBorder;
            IsItalic = isItalic;
        }
        public string StyleName { get; set; }
        public Color? BackgroundColour { get; set; }

        public bool IsBold { get; set; }

        public XL.XlHAlign? Justification { get; set; }

        public XL.XlLineStyle? TopBorder { get; set; }

        public XL.XlLineStyle? BottomBorder { get; set; }

        public bool IsItalic { get; set; }

    }
}
