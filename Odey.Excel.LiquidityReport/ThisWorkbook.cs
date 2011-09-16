using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Exc = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Odey.LiquidityCalculator.Clients;
using Odey.LiquidityCalculator.Contracts;

namespace Odey.Excel.LiquidityReport
{
    public partial class ThisWorkbook
    {
        

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            LiquidityCalculatorClient client = new LiquidityCalculatorClient();
            Dictionary<string,LiquidityCalculatorOutput> outputs = client.Calculate(new LiquidityCalculatorInput(20, 20, 2, 50, 5), 4);
            Exc.Worksheet mainWorkSheet = (Exc.Worksheet)Globals.ThisWorkbook.Worksheets[1];
            int maxColumn = 9;
            Exc.Style headingStyle = Globals.ThisWorkbook.Styles.Add("HeadingStyle", missing);
            headingStyle.Interior.Color = Exc.XlRgbColor.rgbCornflowerBlue;
            headingStyle.Interior.Pattern = Exc.XlPattern.xlPatternSolid;
            headingStyle.Borders.Color = Exc.XlRgbColor.rgbBlack;
            headingStyle.Borders.LineStyle = Exc.XlLineStyle.xlContinuous;
            headingStyle.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Exc.XlLineStyle.xlLineStyleNone;
            headingStyle.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Exc.XlLineStyle.xlLineStyleNone;
            headingStyle.Font.Color = Exc.XlRgbColor.rgbWhite;
            headingStyle.Font.Bold = true;
            headingStyle.WrapText = true;
            headingStyle.HorizontalAlignment = Exc.XlHAlign.xlHAlignCenter;
            mainWorkSheet.Range[mainWorkSheet.Cells[1, 1], mainWorkSheet.Cells[1, maxColumn]].Style = "HeadingStyle";

            Exc.Style mainGridStyle = Globals.ThisWorkbook.Styles.Add("MainGridStyle", missing);
            mainGridStyle.Borders.Color = Exc.XlRgbColor.rgbBlack;
            mainGridStyle.Borders.LineStyle = Exc.XlLineStyle.xlContinuous;
            mainGridStyle.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Exc.XlLineStyle.xlLineStyleNone;
            mainGridStyle.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Exc.XlLineStyle.xlLineStyleNone;

            //mainWorkSheet.Rows[1].Interior.Color = Exc.XlRgbColor.rgbLightGray;
            mainWorkSheet.Columns[4].ColumnWidth = 14;
            mainWorkSheet.Columns[5].ColumnWidth = 14;
            mainWorkSheet.Columns[8].ColumnWidth = 14;
            mainWorkSheet.Columns[9].ColumnWidth = 17;

            mainWorkSheet.Cells[1, 1] = "Fund";
            mainWorkSheet.Cells[1, 2] = "GrossNav";
            mainWorkSheet.Columns[2].NumberFormat = "#,###";
            mainWorkSheet.Cells[1, 3] = "Nav";
            mainWorkSheet.Columns[3].NumberFormat = "#,###";
            mainWorkSheet.Cells[1, 4] = "Number Nonliquidated Positions";
            mainWorkSheet.Cells[1, 5] = "Number Of Positions";
            mainWorkSheet.Cells[1, 6] = "Total Fine";
            mainWorkSheet.Columns[6].NumberFormat = "#,###";
            mainWorkSheet.Cells[1, 7] = "Unsold Value";
            mainWorkSheet.Columns[7].NumberFormat = "#,###";
            mainWorkSheet.Cells[1, 8] = "Number Of Exceptions";
            mainWorkSheet.Cells[1, 9] = "Weighted Number of Days to 100% Liquidated ";
            mainWorkSheet.Columns[9].NumberFormat = "#,###";


            
            int maxFundColumn = 13;
            int row = 2;
            int sheetNumber = 2;
            Exc.Worksheet fundWorksheet = null;
            foreach (KeyValuePair<string, LiquidityCalculatorOutput> output in outputs.OrderByDescending(a=>a.Value.UnsoldValue))
            {
                mainWorkSheet.Cells[row, 1] = output.Key;
                mainWorkSheet.Cells[row, 2] = output.Value.GrossNav;                
                mainWorkSheet.Cells[row, 3] = output.Value.Nav;
                mainWorkSheet.Cells[row, 4] = output.Value.NonLiquidatedPortfolio.Count;
                mainWorkSheet.Cells[row, 5] = output.Value.NumberOfPositions;
                mainWorkSheet.Cells[row, 6] = output.Value.TotalFine;
                mainWorkSheet.Cells[row, 7] = output.Value.UnsoldValue;
                mainWorkSheet.Cells[row, 8] = output.Value.Exceptions.Count;
                mainWorkSheet.Cells[row, 9] = output.Value.WeightedDaysToLiquidatePortfolio;
                if (row % 2 != 0)
                {
                    mainWorkSheet.Range[mainWorkSheet.Cells[row, 1], mainWorkSheet.Cells[row, maxColumn]].Interior.Color = Exc.XlRgbColor.rgbLightGray;
                }
                if (Globals.ThisWorkbook.Worksheets.Count < sheetNumber)
                {
                    fundWorksheet = (Exc.Worksheet)Globals.ThisWorkbook.Worksheets.Add(missing, fundWorksheet, missing, missing);
                }
                else
                {
                    fundWorksheet = (Exc.Worksheet)Globals.ThisWorkbook.Worksheets[sheetNumber];
                }


                fundWorksheet.Cells[1, 1] = "Non Liquid Positions";

                fundWorksheet.Cells[2, 1] = "Instrument Name";
                fundWorksheet.Cells[2, 2] = "Amount To Liquidate";
                fundWorksheet.Cells[2, 3] = "Amount Unable To Be Liquidated";
                fundWorksheet.Cells[2, 4] = "Daily Available Liquidity";
                fundWorksheet.Cells[2, 5] = "Daily Liquidity";
                fundWorksheet.Cells[2, 6] = "Days To Liquidate Complete Position";
                fundWorksheet.Cells[2, 7] = "Delta Market Value";
                fundWorksheet.Cells[2, 8] = "Excess Days To Liquidate";
                fundWorksheet.Cells[2, 9] = "Market Value";
                fundWorksheet.Cells[2, 10] = "Net Position";
                fundWorksheet.Cells[2, 11] = "Value Of Amount Unable To Be Liquidated"; 
                fundWorksheet.Cells[2, 12] = "Weighted Days To Liquidate Portfolio";
                fundWorksheet.Cells[2, 13] = "Write Down With Fine";
                fundWorksheet.Range[fundWorksheet.Cells[1, 1], fundWorksheet.Cells[2, maxFundColumn]].Style = "HeadingStyle";
                fundWorksheet.Range[fundWorksheet.Cells[1, 1], fundWorksheet.Cells[1, maxFundColumn]].Merge();
                fundWorksheet.Rows[1].AutoFit();
                int fundRow = 3;
                foreach (LiquidityCalculatorNonLiquidatedPosition nonLiquidPosition in output.Value.NonLiquidatedPortfolio.Values.OrderByDescending(a=>a.ValueOfAmountUnableTobeLiquidated))
                {
                    if (fundRow % 2 != 0)
                    {
                        fundWorksheet.Range[fundWorksheet.Cells[fundRow, 1], fundWorksheet.Cells[fundRow, maxFundColumn]].Interior.Color = Exc.XlRgbColor.rgbLightGray;
                    }
                    fundWorksheet.Cells[fundRow, 1] = nonLiquidPosition.InstrumentName;
                    fundWorksheet.Cells[fundRow, 2] = nonLiquidPosition.AmountToLiquidate;
                    fundWorksheet.Cells[fundRow, 3] = nonLiquidPosition.AmountUnableToBeLiquidated;
                    fundWorksheet.Cells[fundRow, 4] = nonLiquidPosition.DailyAvailableLiquidity;
                    fundWorksheet.Cells[fundRow, 5] = nonLiquidPosition.DailyLiquidity;
                    fundWorksheet.Cells[fundRow, 6] = nonLiquidPosition.DaysToLiquidateCompletePosition;
                    fundWorksheet.Cells[fundRow, 7] = nonLiquidPosition.DeltaMarketValue;
                    fundWorksheet.Cells[fundRow, 8] = nonLiquidPosition.ExcessDaysToLiquidate;
                    fundWorksheet.Cells[fundRow, 9] = nonLiquidPosition.MarketValue;
                    fundWorksheet.Cells[fundRow, 10] = nonLiquidPosition.NetPosition;
                    fundWorksheet.Cells[fundRow, 11] = nonLiquidPosition.ValueOfAmountUnableTobeLiquidated;
                    fundWorksheet.Cells[fundRow, 12] = nonLiquidPosition.WeightedDaysToLiquidatePortfolio;
                    fundWorksheet.Cells[fundRow, 13] = nonLiquidPosition.WriteDownWithFine;
                    fundRow++;
                }
                if (output.Value.NonLiquidatedPortfolio.Values.Count > 0)
                {
                    fundWorksheet.Cells[fundRow, 11].Formula = String.Format("=Sum(K3:K{0})", fundRow - 1);
                    fundWorksheet.Cells[fundRow, 12].Formula = String.Format("=Average(L3:L{0})", fundRow - 1);
                    fundWorksheet.Cells[fundRow, 13].Formula = String.Format("=Sum(M3:M{0})", fundRow - 1);
                    fundRow++;
                }
                
                fundRow++;
                fundWorksheet.Columns.AutoFit();
                fundWorksheet.Columns.NumberFormat = "#,###";
                fundWorksheet.Cells[fundRow, 1] = "Exceptions";
                fundWorksheet.Range[fundWorksheet.Cells[fundRow, 1], fundWorksheet.Cells[fundRow, maxFundColumn]].Style = "HeadingStyle";
                fundWorksheet.Range[fundWorksheet.Cells[fundRow, 1], fundWorksheet.Cells[fundRow, maxFundColumn]].Merge();
                fundRow++;
                foreach (KeyValuePair<string,string> exception in output.Value.Exceptions)
                {
                    if (fundRow % 2 != 0)
                    {
                        fundWorksheet.Range[fundWorksheet.Cells[fundRow, 1], fundWorksheet.Cells[fundRow, maxFundColumn]].Interior.Color = Exc.XlRgbColor.rgbLightGray;
                    }
                    fundWorksheet.Cells[fundRow, 1] = exception.Key;
                    fundWorksheet.Cells[fundRow, 2] = exception.Value;
                    fundRow++;
                }
                fundWorksheet.Name = output.Key;
                sheetNumber++;
                row++;
            }
            mainWorkSheet.Activate();
            mainWorkSheet.Name = "Results";
            mainWorkSheet.Columns[2].AutoFit();
            mainWorkSheet.Columns[3].AutoFit();            
            mainWorkSheet.Columns[6].AutoFit();
            mainWorkSheet.Columns[7].AutoFit();

            Exc.Range mainGridRange =  mainWorkSheet.Range[mainWorkSheet.Cells[2, 1], mainWorkSheet.Cells[row - 1, maxColumn]];
            mainGridRange.Borders.Color = Exc.XlRgbColor.rgbBlack;
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
