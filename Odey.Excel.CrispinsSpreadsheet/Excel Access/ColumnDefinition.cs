using XL=Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class ColumnDefinition
    {


        public ColumnDefinition(int columnNumber,string columnLabel, string headerLabel, string style, bool isHidden, decimal width, string bloombergMneumonic, Color? backgroundColour,string fxStyleName, XL.XlHAlign headerJustification, bool totalIsBold, bool isSummable,bool hasRightHandBorder)
        {
            ColumnNumber = columnNumber;
            ColumnLabel = columnLabel;
            HeaderLabel = headerLabel;
            BloombergMneumonic = bloombergMneumonic;
            StyleName = style;
            IsHidden = isHidden;
            Width = width;
            BackgroundColour = backgroundColour;
            FXStyleName = fxStyleName;
            HeaderJustification = headerJustification;
            TotalIsBold = totalIsBold;
            IsSummable = isSummable;
            HasRightHandBorder = hasRightHandBorder;
        }




        public int ColumnNumber { get; set; }
        public string HeaderLabel { get; set; }
        public string StyleName { get; set; }


        
        public bool IsHidden { get; set; }
        public decimal Width { get; set; }
        public string BloombergMneumonic { get; set; }
        public Color? BackgroundColour { get; set; }

        public string ColumnLabel { get; set; }

        public string FXStyleName { get; set; }

        public XL.XlHAlign HeaderJustification { get; set; }

        public bool TotalIsBold { get; set; }

        public bool IsSummable { get; set; }



        public bool HasRightHandBorder { get; set; }
    }
}
