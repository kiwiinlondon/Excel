using XL=Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class CellStyler
    {
        private static readonly CellStyler instance = new CellStyler();

        private CellStyler()
        {

        }

        public static CellStyler Instance
        {
            get
            {
                return instance;
            }
        }


        public static readonly string StyleNormal = "Normal";
        public static readonly string StylePrice = "CO_Price";
        public static readonly string StylePriceChange = "CO_PriceChange";
        public static readonly string StylePercentageChange = "CO_PercentageChange";
        public static readonly string StyleUnits = "CO_Units";
        public static readonly string StyleFXRate = "CO_FXRate";
        public static readonly string StylePNL = "CO_PNL";
        public static readonly string StyleContribution = "CO_ContributionPercentage";
        public static readonly string StyleExposure = "CO_Exposure";
        public static readonly string StyleExposurePercentage = "CO_ExposurePercentage";
        public static Color PreviousSectionGrey = Color.FromArgb(242, 242, 242);
        public static Color HeaderGrey = Color.FromArgb(217, 217, 217);

        private CellStyling GetCellStyling(RowType rowType, ColumnDefinition column)
        {
            string style = column.StyleName;
            bool isBold = false;
            Color? backgroundColour = column.BackgroundColour;
            XL.XlHAlign? justification = null;
            XL.XlLineStyle? topBorder = null;
            XL.XlLineStyle? bottomBorder = null;
            bool isItalic = false;
            switch (rowType)
            {
                case RowType.FXPosition:
                    if (!string.IsNullOrWhiteSpace(column.FXStyleName))
                    {
                        style = column.FXStyleName;
                    };
                    break;
                case RowType.Header:
                    style = StyleNormal;
                    isBold = true;
                    justification = column.HeaderJustification;
                    backgroundColour = HeaderGrey;
                    topBorder = XL.XlLineStyle.xlContinuous;
                    break;
                case RowType.MainBookOrAssetClassTotal:
                    isBold = column.TotalIsBold;
                    topBorder = XL.XlLineStyle.xlContinuous;
                    bottomBorder = XL.XlLineStyle.xlContinuous;
                    backgroundColour = HeaderGrey;
                    break;
                case RowType.Total:
                    isBold = column.TotalIsBold;
                    topBorder = XL.XlLineStyle.xlContinuous;
                    bottomBorder = XL.XlLineStyle.xlContinuous;
                    break;
                case RowType.FundTotal:
                    isBold = column.TotalIsBold;
                    topBorder = XL.XlLineStyle.xlContinuous;
                    bottomBorder = XL.XlLineStyle.xlDouble;
                    backgroundColour = HeaderGrey;
                    break;
            }

            return new CellStyling(style, backgroundColour, isBold, justification, topBorder, bottomBorder, isItalic);
        }

        public void ApplyStyle(Row row, ColumnDefinition column)
        {
            var cell = row.Range.Cells[1, column.ColumnNumber];
            ApplyStyle(cell, row.RowType, column);
        }

        public void ApplyStyle(XL.Range cell, RowType rowType, ColumnDefinition column)
        {
            if (rowType != RowType.AdditionalFundTotal)
            {
                CellStyling cellStyling = GetCellStyling(rowType, column);

                cell.Style = cellStyling.StyleName;

                if (cellStyling.BackgroundColour.HasValue)
                {
                    cell.Interior.Color = cellStyling.BackgroundColour;
                }
                cell.Font.Bold = cellStyling.IsBold;
                cell.Font.Italic = cellStyling.IsItalic;
                if (cellStyling.Justification.HasValue)
                {
                    cell.HorizontalAlignment = cellStyling.Justification.Value;
                }
                if (cellStyling.TopBorder.HasValue)
                {
                    cell.Borders[XL.XlBordersIndex.xlEdgeTop].LineStyle = cellStyling.TopBorder.Value;
                }
                if (cellStyling.BottomBorder.HasValue)
                {
                    cell.Borders[XL.XlBordersIndex.xlEdgeBottom].LineStyle = cellStyling.BottomBorder.Value;
                }
            }
        }
    }
}
