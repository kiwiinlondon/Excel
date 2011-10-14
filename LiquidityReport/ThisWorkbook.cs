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

namespace LiquidityReport
{
    public partial class ThisWorkbook
    {
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            InitialBuild();
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
            
        }

        #region Initial Build
        private void InitialBuild()
        {
            decimal percentageOfPortfolioToLiquidate = 20;
            decimal percentageOfDailyVolume = 20;
            decimal fine = 2;
            decimal fineCap = 50;
            decimal numberOfDays = 5;
            int daysToLiquidateNonTradingAssets = 200;
            
            mainWorkSheet = (Exc.Worksheet)Globals.ThisWorkbook.Worksheets[1];
           
            BuildMainPageStructure();
            AddParameterStructure();
            Refresh(DateTime.Now.Date, percentageOfPortfolioToLiquidate, percentageOfDailyVolume, fine, fineCap, numberOfDays, daysToLiquidateNonTradingAssets);
            var buttonRange = Globals.Sheet1.Range["N13:P13"];
            var button = Globals.Sheet1.Controls.AddButton(buttonRange, "Refresh Button");
            button.Text = "Refresh";
            button.Click += new EventHandler(button_Click);
        }
        #endregion

        

        #region Build Main Page Structure
        
        private int maxColumn = 12;
        private Exc.Worksheet mainWorkSheet = null;

        private void BuildMainPageStructure()
        {
            mainWorkSheet.Name = "Results";
            AddStyles();
            
            mainWorkSheet.Range[mainWorkSheet.Cells[1, 1], mainWorkSheet.Cells[1, maxColumn]].Style = "HeadingStyle";

            mainWorkSheet.Cells[1, 1] = "Fund";
            mainWorkSheet.Cells[1, 2] = "Gross Nav";
            mainWorkSheet.Columns[2].NumberFormat = "#,###";
            mainWorkSheet.Cells[1, 3] = "Nav";
            mainWorkSheet.Columns[3].NumberFormat = "#,###";
            mainWorkSheet.Cells[1, 4] = "Number Nonliquidated Positions";
            mainWorkSheet.Cells[1, 5] = "Number Of Positions";
            mainWorkSheet.Cells[1, 6] = "Total Fine";
            mainWorkSheet.Columns[6].NumberFormat = "#,###";
            mainWorkSheet.Cells[1, 7] = "Unsold Value";
            mainWorkSheet.Columns[7].NumberFormat = "#,###";
            mainWorkSheet.Cells[1, 8] = "Unsold Value % of Nav";
            mainWorkSheet.Columns[8].NumberFormat = "0.00%";
            mainWorkSheet.Cells[1, 9] = "Number Of Exceptions";
            mainWorkSheet.Cells[1, 10] = "Value Of Exceptions";
            mainWorkSheet.Columns[10].NumberFormat = "#,###";
            mainWorkSheet.Cells[1, 11] = "Exceptions % of Nav";
            mainWorkSheet.Columns[11].NumberFormat = "0.00%";
            mainWorkSheet.Cells[1, 12] = "Weighted Number of Days to 100% Liquidated ";
            mainWorkSheet.Columns[12].NumberFormat = "#,###";
        }
        #endregion

        #region Build Fund Page Structure
        private void BuildFundPageStructure(Exc.Worksheet fundWorkSheet, int maxFundColumn, string name)
        {
            fundWorkSheet.Cells[1, 1] = "Non Liquid Positions";

            fundWorkSheet.Cells[2, 1] = "Instrument Name";
            fundWorkSheet.Cells[2, 2] = "Amount To Liquidate";
            fundWorkSheet.Cells[2, 3] = "Amount Unable To Be Liquidated";
            fundWorkSheet.Cells[2, 4] = "Daily Available Liquidity";
            fundWorkSheet.Cells[2, 5] = "Daily Liquidity";
            fundWorkSheet.Cells[2, 6] = "Days To Liquidate Complete Position";
            fundWorkSheet.Cells[2, 7] = "Delta Market Value";
            fundWorkSheet.Cells[2, 8] = "Excess Days To Liquidate";
            fundWorkSheet.Cells[2, 9] = "Market Value";
            fundWorkSheet.Cells[2, 10] = "Net Position";
            fundWorkSheet.Cells[2, 11] = "Value Of Amount Unable To Be Liquidated";
            fundWorkSheet.Cells[2, 12] = "Weighted Days To Liquidate Portfolio";
            fundWorkSheet.Cells[2, 13] = "Write Down With Fine";
            fundWorkSheet.Cells[2, 14] = "CIS Days Between Dealing Days";
            fundWorkSheet.Cells[2, 15] = "Listing Status";
            fundWorkSheet.Range[fundWorkSheet.Cells[1, 1], fundWorkSheet.Cells[2, maxFundColumn]].Style = "HeadingStyle";
            fundWorkSheet.Range[fundWorkSheet.Cells[1, 1], fundWorkSheet.Cells[1, maxFundColumn]].Merge();
            fundWorkSheet.Name = name;
        }
        #endregion

        #region Build Fund Page Structure
        private void BuildFundExcpetionStructure(Exc.Worksheet fundWorkSheet, ref int fundRow)
        {
            fundWorkSheet.Columns.NumberFormat = "#,###";
            fundWorkSheet.Cells[fundRow++, 1] = "Exceptions";
            fundWorkSheet.Cells[fundRow, 1] = "Instrument Name";
            fundWorkSheet.Cells[fundRow, 2] = "Net Position";
            fundWorkSheet.Cells[fundRow, 3] = "Market Value";
            fundWorkSheet.Cells[fundRow, 4] = "Delta Market Value";
            fundWorkSheet.Range[fundWorkSheet.Cells[fundRow - 1, 1], fundWorkSheet.Cells[fundRow, 4]].Style = "HeadingStyle";
            fundWorkSheet.Range[fundWorkSheet.Cells[fundRow - 1, 1], fundWorkSheet.Cells[fundRow - 1, 4]].Merge();
            fundWorkSheet.Columns[2].ColumnWidth = 10;
            fundWorkSheet.Columns[3].ColumnWidth = 10;
            fundWorkSheet.Columns[4].ColumnWidth = 10;
            fundWorkSheet.Columns[6].ColumnWidth = 10;
            fundWorkSheet.Columns[8].ColumnWidth = 10;
            fundRow++;
        }
        #endregion

        #region Build Fund Page Structure
        private void AddFundExceptionValues(Exc.Worksheet fundWorkSheet, int maxFundColumn, LiquidityCalculatorOutput output, ref int fundRow)
        {
            if (output.Exceptions.Count > 0)
            {
                BuildFundExcpetionStructure(fundWorkSheet, ref fundRow);
            }
            foreach (KeyValuePair<string, LiquidityCalculatorException> exception in output.Exceptions.OrderByDescending(a => Math.Abs(a.Value.MarketValue)))
            {
                if (fundRow % 2 != 0)
                {
                    fundWorkSheet.Range[fundWorkSheet.Cells[fundRow, 1], fundWorkSheet.Cells[fundRow, 4]].Interior.Color = Exc.XlRgbColor.rgbLightGray;
                }
                fundWorkSheet.Cells[fundRow, 1] = exception.Key;
                fundWorkSheet.Cells[fundRow, 2] = exception.Value.NetPosition;
                fundWorkSheet.Cells[fundRow, 3] = exception.Value.MarketValue;
                fundWorkSheet.Cells[fundRow, 4] = exception.Value.DeltaMarketValue;
                fundRow++;
            }
        }
        #endregion

        #region Build Fund Page Structure
        private void AddFundPageValues(Exc.Worksheet fundWorkSheet, int maxFundColumn, LiquidityCalculatorOutput output, ref int fundRow)
        {
            foreach (LiquidityCalculatorNonLiquidatedPosition nonLiquidPosition in output.NonLiquidatedPortfolio.Values.OrderByDescending(a => a.ValueOfAmountUnableTobeLiquidated))
            {
                if (fundRow % 2 != 0)
                {
                    fundWorkSheet.Range[fundWorkSheet.Cells[fundRow, 1], fundWorkSheet.Cells[fundRow, maxFundColumn]].Interior.Color = Exc.XlRgbColor.rgbLightGray;
                }
                fundWorkSheet.Cells[fundRow, 1] = nonLiquidPosition.InstrumentName;
                fundWorkSheet.Cells[fundRow, 2] = nonLiquidPosition.AmountToLiquidate;
                fundWorkSheet.Cells[fundRow, 3] = nonLiquidPosition.AmountUnableToBeLiquidated;
                fundWorkSheet.Cells[fundRow, 4] = nonLiquidPosition.DailyAvailableLiquidity;
                fundWorkSheet.Cells[fundRow, 5] = nonLiquidPosition.DailyLiquidity;
                fundWorkSheet.Cells[fundRow, 6] = nonLiquidPosition.DaysToLiquidateCompletePosition;
                fundWorkSheet.Cells[fundRow, 7] = nonLiquidPosition.DeltaMarketValue;
                fundWorkSheet.Cells[fundRow, 8] = nonLiquidPosition.ExcessDaysToLiquidate;
                fundWorkSheet.Cells[fundRow, 9] = nonLiquidPosition.MarketValue;
                fundWorkSheet.Cells[fundRow, 10] = nonLiquidPosition.NetPosition;
                fundWorkSheet.Cells[fundRow, 11] = nonLiquidPosition.ValueOfAmountUnableTobeLiquidated;
                fundWorkSheet.Cells[fundRow, 12] = nonLiquidPosition.WeightedDaysToLiquidatePortfolio;
                fundWorkSheet.Cells[fundRow, 13] = nonLiquidPosition.WriteDownWithFine;
                fundWorkSheet.Cells[fundRow, 14] = nonLiquidPosition.CISDaysBetweenDealingDays;
                fundWorkSheet.Cells[fundRow, 15] = nonLiquidPosition.ListedStatus;
                fundRow++;
            }
            if (output.NonLiquidatedPortfolio.Values.Count > 0)
            {
                fundWorkSheet.Cells[fundRow, 11].Formula = String.Format("=Sum(K3:K{0})", fundRow - 1);
                fundWorkSheet.Cells[fundRow, 12].Formula = String.Format("=Average(L3:L{0})", fundRow - 1);
                fundWorkSheet.Cells[fundRow, 13].Formula = String.Format("=Sum(M3:M{0})", fundRow - 1);
                fundRow++;
            }
        }
        #endregion

        #region Add Styles
        private void AddStyles()
        {
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

            Exc.Style mainGridStyle = Globals.ThisWorkbook.Styles.Add("MainGridStyle", missing);
            mainGridStyle.Borders.Color = Exc.XlRgbColor.rgbBlack;
            mainGridStyle.Borders.LineStyle = Exc.XlLineStyle.xlContinuous;
            mainGridStyle.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Exc.XlLineStyle.xlLineStyleNone;
            mainGridStyle.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Exc.XlLineStyle.xlLineStyleNone;
        }
        #endregion

        #region Add Parameter Values
        private void AddParameterValues(
            DateTime dateToFind,
            decimal percentageOfPortfolioToLiquidate,
            decimal percentageOfDailyVolume,
            decimal fine,
            decimal fineCap,
            decimal numberOfDays,
            int daysToLiquidateNonTradingAssets)
        {
            mainWorkSheet.Cells[5, maxColumn + 3] = dateToFind;
            mainWorkSheet.Cells[6, maxColumn + 3] = percentageOfPortfolioToLiquidate;
            mainWorkSheet.Cells[7, maxColumn + 3] = percentageOfDailyVolume;
            mainWorkSheet.Cells[8, maxColumn + 3] = fine;
            mainWorkSheet.Cells[9, maxColumn + 3] = fineCap;
            mainWorkSheet.Cells[10, maxColumn + 3] = numberOfDays;
            mainWorkSheet.Cells[11, maxColumn + 3] = daysToLiquidateNonTradingAssets;
        }
        #endregion

        #region Add Parameter Structure
        private void AddParameterStructure()
        {
            mainWorkSheet.Columns[4].ColumnWidth = 13;
            mainWorkSheet.Columns[5].ColumnWidth = 8.1;
            mainWorkSheet.Columns[9].ColumnWidth = 10;
            mainWorkSheet.Columns[10].ColumnWidth = 10;
            mainWorkSheet.Columns[11].ColumnWidth = 10;
            mainWorkSheet.Columns[12].ColumnWidth = 12;
            mainWorkSheet.Columns[13].ColumnWidth = 2;
            mainWorkSheet.Cells[4, maxColumn + 2] = "Parameters";
            mainWorkSheet.Cells[5, maxColumn + 2] = "Date";            
            mainWorkSheet.Cells[6, maxColumn + 2] = "Percentage To Liquidate";            
            mainWorkSheet.Cells[6, maxColumn + 4] = "%";
            mainWorkSheet.Cells[7, maxColumn + 2] = "Daily Volume";
            mainWorkSheet.Cells[7, maxColumn + 4] = "%";
            mainWorkSheet.Cells[8, maxColumn + 2] = "Fine";   
            mainWorkSheet.Cells[8, maxColumn + 4] = "%";
            mainWorkSheet.Cells[9, maxColumn + 2] = "Fine Cap";            
            mainWorkSheet.Cells[9, maxColumn + 4] = "%";
            mainWorkSheet.Cells[10, maxColumn + 2] = "Number of Days";            
            mainWorkSheet.Cells[11, maxColumn + 2] = "Days To Liquidate Non-Trading";
            
            mainWorkSheet.Range[mainWorkSheet.Cells[4, maxColumn + 2], mainWorkSheet.Cells[4, maxColumn + 4]].Style = "HeadingStyle";
            mainWorkSheet.Range[mainWorkSheet.Cells[4, maxColumn + 2], mainWorkSheet.Cells[4, maxColumn + 4]].Merge();
            mainWorkSheet.Range[mainWorkSheet.Cells[5, maxColumn + 2], mainWorkSheet.Cells[5, maxColumn + 4]].Interior.Color = Exc.XlRgbColor.rgbLightGray;
            mainWorkSheet.Range[mainWorkSheet.Cells[5, maxColumn + 3], mainWorkSheet.Cells[5, maxColumn + 4]].Merge();
            mainWorkSheet.Range[mainWorkSheet.Cells[5, maxColumn + 3], mainWorkSheet.Cells[5, maxColumn + 4]].HorizontalAlignment = Exc.XlHAlign.xlHAlignCenter;
            mainWorkSheet.Range[mainWorkSheet.Cells[7, maxColumn + 2], mainWorkSheet.Cells[7, maxColumn + 4]].Interior.Color = Exc.XlRgbColor.rgbLightGray;
            mainWorkSheet.Range[mainWorkSheet.Cells[9, maxColumn + 2], mainWorkSheet.Cells[9, maxColumn + 4]].Interior.Color = Exc.XlRgbColor.rgbLightGray;
            mainWorkSheet.Range[mainWorkSheet.Cells[11, maxColumn + 2], mainWorkSheet.Cells[11, maxColumn + 4]].Interior.Color = Exc.XlRgbColor.rgbLightGray;
            mainWorkSheet.Range[mainWorkSheet.Cells[5, maxColumn + 2], mainWorkSheet.Cells[11, maxColumn + 4]].Borders.Color = Exc.XlRgbColor.rgbBlack;
            mainWorkSheet.Range[mainWorkSheet.Cells[5, maxColumn + 3], mainWorkSheet.Cells[11, maxColumn + 3]].Borders[Exc.XlBordersIndex.xlEdgeRight].LineStyle = Exc.XlLineStyle.xlLineStyleNone;
        }
        #endregion

        #region Button Click
        private void button_Click(object sender, EventArgs e)
        {

            DateTime dateToFind = (DateTime)Globals.Sheet1.Cells[5, 15].Value;
            decimal percentageOfPortfolioToLiquidate = (decimal)Globals.Sheet1.Cells[6,15].Value;
            decimal percentageOfDailyVolume = (decimal)Globals.Sheet1.Cells[7, 15].Value;
            decimal fine = (decimal)Globals.Sheet1.Cells[8, 15].Value;
            decimal fineCap = (decimal)Globals.Sheet1.Cells[9, 15].Value;
            decimal numberOfDays = (decimal)Globals.Sheet1.Cells[10, 15].Value;
            int daysToLiquidateNonTradingAssets = (int)Globals.Sheet1.Cells[11, 15].Value;
            Exc.Worksheet mainWorkSheet = (Exc.Worksheet)Globals.ThisWorkbook.Worksheets[1];
            Refresh(dateToFind, percentageOfPortfolioToLiquidate, percentageOfDailyVolume, fine, fineCap, numberOfDays, daysToLiquidateNonTradingAssets);
        }
        #endregion

        #region Add Main Page Value
        private void AddMainPageValues(string fundName, LiquidityCalculatorOutput output, Exc.Worksheet mainWorkSheet,int row,int maxColumn)
        {
            mainWorkSheet.Cells[row, 1] = fundName;
            mainWorkSheet.Cells[row, 2] = output.GrossNav;
            mainWorkSheet.Cells[row, 3] = output.Nav;
            mainWorkSheet.Cells[row, 4] = output.NonLiquidatedPortfolio.Count;
            mainWorkSheet.Cells[row, 5] = output.NumberOfPositions;
            mainWorkSheet.Cells[row, 6] = output.TotalFine;
            mainWorkSheet.Cells[row, 7] = output.UnsoldValue;
            mainWorkSheet.Cells[row, 8].Formula = String.Format("=G{0}/B{0}", row);
            mainWorkSheet.Cells[row, 9] = output.Exceptions.Count;
            mainWorkSheet.Cells[row, 10] = output.TotalMarketValueExceptions;
            mainWorkSheet.Cells[row, 11].Formula = String.Format("=J{0}/B{0}", row);
            mainWorkSheet.Cells[row, 12] = output.WeightedDaysToLiquidatePortfolio;
            if (row % 2 != 0)
            {
                mainWorkSheet.Range[mainWorkSheet.Cells[row, 1], mainWorkSheet.Cells[row, maxColumn]].Interior.Color = Exc.XlRgbColor.rgbLightGray;
            }
        }
        #endregion

        #region Format Main Work Sheet After Data
        private void FormatMainWorkSheetAfterData(int row)
        {
            mainWorkSheet.Activate();

            mainWorkSheet.Columns[2].AutoFit();
            mainWorkSheet.Columns[3].AutoFit();
            mainWorkSheet.Columns[6].AutoFit();
            mainWorkSheet.Columns[7].AutoFit();
            mainWorkSheet.Columns[8].AutoFit();
            mainWorkSheet.Columns[maxColumn + 2].AutoFit();
            mainWorkSheet.Columns[maxColumn + 4].AutoFit();
            mainWorkSheet.Rows[4].Autofit();
            mainWorkSheet.Rows[1].Autofit();
            Exc.Range mainGridRange = mainWorkSheet.Range[mainWorkSheet.Cells[2, 1], mainWorkSheet.Cells[row - 1, maxColumn]];
            mainGridRange.Borders.Color = Exc.XlRgbColor.rgbBlack;
            mainWorkSheet.Cells[1, 1].Select();
        }
        #endregion

        #region Format Fund Work Sheet After Data
        private void FormatFundWorkSheetAfterData(Exc.Worksheet fundWorkSheet)
        {
            fundWorkSheet.Rows[1].AutoFit();
            fundWorkSheet.Columns.AutoFit();
        }
        #endregion 


        #region Refresh
        private void Refresh(           
            DateTime dateToFind,
            decimal percentageOfPortfolioToLiquidate,
            decimal percentageOfDailyVolume,
            decimal fine,
            decimal fineCap,
            decimal numberOfDays,
            int daysToLiquidateNonTradingAssets)
        {
            LiquidityCalculatorClient client = new LiquidityCalculatorClient();

            TimeSpan difference = DateTime.Now.Date.Subtract(dateToFind);

            Dictionary<string, LiquidityCalculatorOutput> outputs = client.Calculate(new LiquidityCalculatorInput(
                percentageOfPortfolioToLiquidate, percentageOfDailyVolume, fine, fineCap, numberOfDays, daysToLiquidateNonTradingAssets), difference.Days);

            Application.DisplayAlerts = false;
            //for (int i = 2; i <= Globals.ThisWorkbook.Worksheets.Count; i++)

            
            foreach(Exc.Worksheet worksheet in Globals.ThisWorkbook.Worksheets)
            {
                if (worksheet.Index != 1)
                {
                    worksheet.Delete();
                }
            }
            Application.DisplayAlerts = true;

            AddParameterValues(dateToFind, percentageOfPortfolioToLiquidate, percentageOfDailyVolume, fine, fineCap, numberOfDays, daysToLiquidateNonTradingAssets);
            int maxFundColumn = 15;
            int row = 2;
            mainWorkSheet.Range[mainWorkSheet.Cells[row, 1], mainWorkSheet.Cells[1000, maxColumn]].ClearContents();
            int sheetNumber = 2;
            Exc.Worksheet fundWorkSheet = mainWorkSheet;
            foreach (KeyValuePair<string, LiquidityCalculatorOutput> output in outputs.OrderByDescending(a => a.Value.UnsoldValue))
            {
                AddMainPageValues(output.Key, output.Value, mainWorkSheet,row, maxColumn);
                                
                if (Globals.ThisWorkbook.Worksheets.Count < sheetNumber)
                {
                    fundWorkSheet = (Exc.Worksheet)Globals.ThisWorkbook.Worksheets.Add(missing, fundWorkSheet, missing, missing);
                }
                else
                {
                    fundWorkSheet = (Exc.Worksheet)Globals.ThisWorkbook.Worksheets[sheetNumber];
                }

                BuildFundPageStructure(fundWorkSheet, maxFundColumn, output.Key);
                
                
                int fundRow = 3;
                AddFundPageValues(fundWorkSheet, maxFundColumn, output.Value, ref fundRow);
                
                fundRow++;

                AddFundExceptionValues(fundWorkSheet, maxFundColumn, output.Value, ref fundRow);
                FormatFundWorkSheetAfterData(fundWorkSheet);      
                sheetNumber++;
                row++;
            }

            FormatMainWorkSheetAfterData(row);
        }
        #endregion

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
