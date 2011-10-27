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
            int daysToLiquidateCap = 200;
            
            mainWorkSheet = (Exc.Worksheet)Globals.ThisWorkbook.Worksheets[1];
           
            //mainWorkSheet.Cells.Clear();
            BuildMainPageStructure();
            AddParameterStructure();
            Refresh(DateTime.Now.Date, percentageOfPortfolioToLiquidate, percentageOfDailyVolume, fine, fineCap, numberOfDays, daysToLiquidateCap);
            var buttonRange = Globals.Sheet1.Range["O13:Q14"];
            var button = Globals.Sheet1.Controls.AddButton(buttonRange, "Refresh Button");
            button.Text = "Refresh";
            button.Click += new EventHandler(button_Click);
        }
        #endregion

        

        #region Build Main Page Structure
        
        private int maxColumn = 13;
        private Exc.Worksheet mainWorkSheet = null;
        private string headingStyleLabel = "HeadingStyle";
        private void BuildMainPageStructure()
        {
            mainWorkSheet.Name = "Results";
            AddStyles();

            mainWorkSheet.Range[mainWorkSheet.Cells[1, 1], mainWorkSheet.Cells[1, maxColumn]].Style = GetStyle(headingStyleLabel);
            mainWorkSheet.Columns[1].ColumnWidth = 20;
            mainWorkSheet.Cells[1, 1] = "Fund";
            mainWorkSheet.Cells[1, 2] = "Nav";
            mainWorkSheet.Columns[2].ColumnWidth = 15;
            mainWorkSheet.Columns[2].NumberFormat = "#,###";
            mainWorkSheet.Cells[1, 3] = "Number Non-liquidated Positions";
            mainWorkSheet.Columns[3].ColumnWidth = 13;
            mainWorkSheet.Cells[1, 4] = "Unsold Value";           
            mainWorkSheet.Columns[4].NumberFormat = "#,###";
            mainWorkSheet.Cells[1, 5] = "Unsold Value % of Nav";
            mainWorkSheet.Columns[5].NumberFormat = "0.00%";
            mainWorkSheet.Columns[5].ColumnWidth = 13;
            mainWorkSheet.Cells[1, 6] = "Total Fine"; 
            mainWorkSheet.Columns[6].NumberFormat = "#,###";
            mainWorkSheet.Cells[1, 7] = "Total Fine % of Nav";
            mainWorkSheet.Columns[7].NumberFormat = "0.00%";
            mainWorkSheet.Columns[7].ColumnWidth = 20;

            mainWorkSheet.Cells[1, 8] = "Value Greater than Max Day Limit";
            mainWorkSheet.Columns[8].NumberFormat = "#,###";
            mainWorkSheet.Cells[1, 9] = "Value Greater Max Day Limit % of Nav";
            mainWorkSheet.Columns[9].NumberFormat = "0.00%";
            mainWorkSheet.Columns[9].ColumnWidth =20;

            mainWorkSheet.Cells[1, 10] = "Number Of Exceptions";
            mainWorkSheet.Columns[10].ColumnWidth = 10;
            mainWorkSheet.Cells[1, 11] = "Value Of Exceptions";
            mainWorkSheet.Columns[11].NumberFormat = "#,###";
            mainWorkSheet.Columns[11].ColumnWidth = 10;
            mainWorkSheet.Cells[1, 12] = "Exceptions % of Nav";
            mainWorkSheet.Columns[12].NumberFormat = "0.00%";
            mainWorkSheet.Columns[12].ColumnWidth = 10;
            mainWorkSheet.Cells[1, 13] = "Weighted Number of Days to 100% Liquidated ";
            mainWorkSheet.Columns[13].NumberFormat = "#,###";
            
            mainWorkSheet.Columns[13].ColumnWidth = 13;
        }
        #endregion

        #region Build Fund Page Structure



        private void BuildFundPageStructure(Exc.Worksheet fundWorkSheet, int maxFundColumn, string name)
        {
            fundWorkSheet.Cells[1, 1] = "Non Liquid Positions";
            fundWorkSheet.Cells[2, 1] = "Instrument Name";
            fundWorkSheet.Cells[2, 2] = "Market Value % Nav";
            fundWorkSheet.Columns[2].NumberFormat ="0.00%";
            fundWorkSheet.Cells[2, 3] = "Delta Market Value % Nav";
            fundWorkSheet.Columns[3].NumberFormat = "0.00%";
            fundWorkSheet.Cells[2, 4] = "Value To Be Liquidated";
            fundWorkSheet.Columns[4].NumberFormat = "#,###";
            fundWorkSheet.Cells[2, 5] = "Total Value Daily Liquidity";
            fundWorkSheet.Columns[5].NumberFormat = "#,###";
            fundWorkSheet.Cells[2, 6] = "Value Available Daily Liquidity";
            fundWorkSheet.Columns[6].NumberFormat = "#,###";
            fundWorkSheet.Cells[2, 7] = "Amount Unsold % Value To Liquidate";
            fundWorkSheet.Columns[7].NumberFormat = "0.00%";
            fundWorkSheet.Cells[2, 8] = "Excess Days To Liquidate";
            fundWorkSheet.Columns[8].NumberFormat = "#,###.00";
            fundWorkSheet.Cells[2, 9] = "Fine % Delta Market Value of Position";
            fundWorkSheet.Columns[9].NumberFormat = "0.00%";
            fundWorkSheet.Cells[2, 10] = "Days To Liquidate Complete Position";
            fundWorkSheet.Columns[10].NumberFormat = "#,###.00";
            fundWorkSheet.Cells[2, 11] = "Weighted Days To Liquidate Portfolio";
            fundWorkSheet.Columns[11].NumberFormat = "0.00";
            fundWorkSheet.Cells[2, 12] = "CIS Days Between Dealing Days";
            fundWorkSheet.Cells[2, 13] = "Listing Status";
            fundWorkSheet.Cells[2, 14] = "Net Position";
            fundWorkSheet.Columns[14].NumberFormat = "#,###";
            fundWorkSheet.Cells[2, 15] = "Amount To Liquidate";
            fundWorkSheet.Columns[15].NumberFormat = "#,###";
            fundWorkSheet.Cells[2, 16] = "Daily Liquidity";
            fundWorkSheet.Columns[16].NumberFormat = "#,###";
            fundWorkSheet.Cells[2, 17] = "Daily Available Liquidity";
            fundWorkSheet.Columns[17].NumberFormat = "#,###";
            fundWorkSheet.Cells[2, 18] = "Amount Unable To Be Liquidated";
            fundWorkSheet.Columns[18].NumberFormat = "#,###";
            fundWorkSheet.Cells[2, 19] = "Market Value";
            fundWorkSheet.Columns[19].NumberFormat = "#,###";
            fundWorkSheet.Cells[2, 20] = "Delta Market Value";
            fundWorkSheet.Columns[20].NumberFormat = "#,###";
            fundWorkSheet.Cells[2, 21] = "Value Of Amount Unsold";
            fundWorkSheet.Columns[21].NumberFormat = "#,###";
            fundWorkSheet.Cells[2, 22] = "Write Down With Fine";
            fundWorkSheet.Columns[22].NumberFormat = "#,###";
            fundWorkSheet.Range[fundWorkSheet.Cells[1, 1], fundWorkSheet.Cells[2, maxFundColumn]].Style = "HeadingStyle";
            fundWorkSheet.Range[fundWorkSheet.Cells[1, 1], fundWorkSheet.Cells[1, maxFundColumn]].Merge();
            fundWorkSheet.Name = name;
        }
        #endregion

        #region Build Fund Exception Structure
        private void BuildFundExceptionStructure(Exc.Worksheet fundWorkSheet, ref int fundRow)
        {
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

        #region Add Fund Exception Values
        private void AddFundExceptionValues(Exc.Worksheet fundWorkSheet, int maxFundColumn, LiquidityCalculatorOutput output, ref int fundRow)
        {
            if (output.Exceptions.Count > 0)
            {
                BuildFundExceptionStructure(fundWorkSheet, ref fundRow);
            }
            foreach (KeyValuePair<string, LiquidityCalculatorException> exception in output.Exceptions.OrderByDescending(a => Math.Abs(a.Value.MarketValue)))
            {
                if (fundRow % 2 != 0)
                {
                    fundWorkSheet.Range[fundWorkSheet.Cells[fundRow, 1], fundWorkSheet.Cells[fundRow, 4]].Interior.Color = Exc.XlRgbColor.rgbLightGray;
                }
                fundWorkSheet.Cells[fundRow, 1] = exception.Key;
                fundWorkSheet.Cells[fundRow, 2] = exception.Value.NetPosition;
                fundWorkSheet.Cells[fundRow, 2].NumberFormat = "#,###";
                fundWorkSheet.Cells[fundRow, 3] = exception.Value.MarketValue;
                fundWorkSheet.Cells[fundRow, 3].NumberFormat = "#,###";
                fundWorkSheet.Cells[fundRow, 4] = exception.Value.DeltaMarketValue;
                fundWorkSheet.Cells[fundRow, 4].NumberFormat = "#,###";
                fundRow++;
            }
        }
        #endregion

        #region Add Fund PageValues
        private void AddFundPageValues(Exc.Worksheet fundWorkSheet, int maxFundColumn, LiquidityCalculatorOutput output, ref int fundRow, int mainSheetRow, int daysToLiquidateCap,decimal numberOfDays)
        {
            List<LiquidityCalculatorNonLiquidatedPosition> nonLiquidPositions = output.NonLiquidatedPortfolio.Values.OrderByDescending(a => (a.ExcessDaysToLiquidate >= daysToLiquidateCap-numberOfDays ? 0 : 1)).ThenByDescending(a => a.DeltaMarketValue/ a.NetPosition * a.AmountUnableToBeLiquidated).ToList(); ;
            //.ThenByDescending(a => a.AmountUnableToBeLiquidated).ToList();
            foreach (LiquidityCalculatorNonLiquidatedPosition nonLiquidPosition in nonLiquidPositions)
            {
                if (fundRow % 2 != 0)
                {
                    fundWorkSheet.Range[fundWorkSheet.Cells[fundRow, 1], fundWorkSheet.Cells[fundRow, maxFundColumn]].Interior.Color = Exc.XlRgbColor.rgbLightGray;
                }
                fundWorkSheet.Cells[fundRow, 1] = nonLiquidPosition.InstrumentName;
                fundWorkSheet.Cells[fundRow, 2].Formula = String.Format("=S{0}/Results!B{1}", fundRow, mainSheetRow);
                fundWorkSheet.Cells[fundRow, 3].Formula = String.Format("=T{0}/Results!B{1}", fundRow, mainSheetRow);
               // fundWorkSheet.Cells[fundRow, 2] = nonLiquidPosition.MarketValue;

             //   fundWorkSheet.Cells[fundRow, 3] = nonLiquidPosition.DeltaMarketValue;
                
                fundWorkSheet.Cells[fundRow, 4].Formula = String.Format("=T{0}/N{0}*O{0}", fundRow);

                fundWorkSheet.Cells[fundRow, 5].Formula = String.Format("=T{0}/N{0}*P{0}", fundRow);
                fundWorkSheet.Cells[fundRow, 6].Formula = String.Format("=T{0}/N{0}*Q{0}", fundRow);

                fundWorkSheet.Cells[fundRow, 7].Formula = String.Format("=if(D{0}=0,0,U{0}/D{0})", fundRow);


                fundWorkSheet.Cells[fundRow, 8] = nonLiquidPosition.ExcessDaysToLiquidate;
                fundWorkSheet.Cells[fundRow, 9].Formula = String.Format("=if(V{0}=0,0,V{0}/T{0})", fundRow);
                fundWorkSheet.Cells[fundRow, 10] = nonLiquidPosition.DaysToLiquidateCompletePosition;
                fundWorkSheet.Cells[fundRow, 11].Formula = String.Format("={0}/Results!B{1}", nonLiquidPosition.WeightedDaysToLiquidatePortfolio, mainSheetRow);
                fundWorkSheet.Cells[fundRow, 12] = nonLiquidPosition.CISDaysBetweenDealingDays;
                fundWorkSheet.Cells[fundRow, 13] = nonLiquidPosition.ListedStatus;

                fundWorkSheet.Cells[fundRow, 14] = nonLiquidPosition.NetPosition;
                fundWorkSheet.Cells[fundRow, 15] = nonLiquidPosition.AmountToLiquidate;
                fundWorkSheet.Cells[fundRow, 16] = nonLiquidPosition.DailyLiquidity;
                fundWorkSheet.Cells[fundRow, 17] = nonLiquidPosition.DailyAvailableLiquidity;                
                fundWorkSheet.Cells[fundRow, 18] = nonLiquidPosition.AmountUnableToBeLiquidated;
                fundWorkSheet.Cells[fundRow, 19] = nonLiquidPosition.MarketValue;

                fundWorkSheet.Cells[fundRow, 20] = nonLiquidPosition.DeltaMarketValue;
                fundWorkSheet.Cells[fundRow, 21] = nonLiquidPosition.ValueOfAmountUnableTobeLiquidated;
                fundWorkSheet.Cells[fundRow, 22] = nonLiquidPosition.WriteDownWithFine;
                
                
                
                
                
                
                fundRow++;
            }
            if (output.NonLiquidatedPortfolio.Values.Count > 0)
            {
             //   fundWorkSheet.Cells[fundRow, 7].Formula = String.Format("=Sum(G3:G{0})", fundRow - 1);
              //  fundWorkSheet.Cells[fundRow, 11].Formula = String.Format("=Average(K3:K{0})", fundRow - 1);
              //  fundWorkSheet.Cells[fundRow, 9].Formula = String.Format("=Sum(I3:I{0})", fundRow - 1);
                fundRow++;
            }
        }
        #endregion

        #region Get Style
        private Exc.Style GetStyle(string name)
        {
            try
            {
                return Globals.ThisWorkbook.Styles[name];
            }
            catch
            {
                return Globals.ThisWorkbook.Styles.Add(name, missing);
            }
            
        }
        #endregion 

        #region Add Styles
        private void AddStyles()
        {

            Exc.Style headingStyle = GetStyle(headingStyleLabel);

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

            Exc.Style mainGridStyle = GetStyle("MainGridStyle");
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

            mainWorkSheet.Columns[maxColumn+1].ColumnWidth = 2;
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
            mainWorkSheet.Cells[11, maxColumn + 2] = "Max Days To Liquidate";
            
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
            DateTime dateToFind = Globals.Sheet1.Cells[5, maxColumn + 3].Value;
            //DateTime dateToFind = DateTime.Parse(dateToFindAsString);
            decimal percentageOfPortfolioToLiquidate = (decimal)Globals.Sheet1.Cells[6, maxColumn+3].Value;
            decimal percentageOfDailyVolume = (decimal)Globals.Sheet1.Cells[7, maxColumn + 3].Value;
            decimal fine = (decimal)Globals.Sheet1.Cells[8, maxColumn + 3].Value;
            decimal fineCap = (decimal)Globals.Sheet1.Cells[9, maxColumn + 3].Value;
            decimal numberOfDays = (decimal)Globals.Sheet1.Cells[10, maxColumn + 3].Value;
            int daysToLiquidateNonTradingAssets = (int)Globals.Sheet1.Cells[11, maxColumn + 3].Value;
            Exc.Worksheet mainWorkSheet = (Exc.Worksheet)Globals.ThisWorkbook.Worksheets[1];
            Refresh(dateToFind, percentageOfPortfolioToLiquidate, percentageOfDailyVolume, fine, fineCap, numberOfDays, daysToLiquidateNonTradingAssets);
        }
        #endregion

        #region Add Main Page Values
        private void AddMainPageValues(string fundName, LiquidityCalculatorOutput output, Exc.Worksheet mainWorkSheet,int row,int maxColumn)
        {
            mainWorkSheet.Cells[row, 1] = fundName;
            mainWorkSheet.Cells[row, 2] = output.Nav;
            mainWorkSheet.Cells[row, 3] = output.NonLiquidatedPortfolio.Count;
            mainWorkSheet.Cells[row, 4] = output.UnsoldValue; 
            mainWorkSheet.Cells[row, 5].Formula = String.Format("=D{0}/B{0}", row);
            mainWorkSheet.Cells[row, 6] = output.TotalFine;
            mainWorkSheet.Cells[row, 7].Formula = String.Format("=F{0}/B{0}", row);
            mainWorkSheet.Cells[row, 8] = output.ValueOutsideLiquidityLimit;
            mainWorkSheet.Cells[row, 9].Formula = String.Format("=H{0}/B{0}", row);
            mainWorkSheet.Cells[row, 10] = output.Exceptions.Count;
            mainWorkSheet.Cells[row, 11] = output.TotalMarketValueExceptions;
            mainWorkSheet.Cells[row, 12].Formula = String.Format("=K{0}/B{0}", row);
            if (output.WeightedDaysToLiquidatePortfolio >= 1)
            {
                mainWorkSheet.Cells[row, 13] = output.WeightedDaysToLiquidatePortfolio;
            }
            else
            {
                mainWorkSheet.Cells[row, 13] = "<1";
                mainWorkSheet.Cells[row, 13].HorizontalAlignment = Exc.XlHAlign.xlHAlignRight;
            }

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

            mainWorkSheet.Columns[4].AutoFit();
            mainWorkSheet.Columns[6].AutoFit();
            mainWorkSheet.Columns[8].AutoFit();
            mainWorkSheet.Columns[maxColumn + 2].AutoFit();
            mainWorkSheet.Columns[maxColumn + 4].AutoFit();
            mainWorkSheet.Rows[4].Autofit();
            mainWorkSheet.Rows[1].Autofit();
            Exc.Range mainGridRange = mainWorkSheet.Range[mainWorkSheet.Cells[2, 1], mainWorkSheet.Cells[row - 1, maxColumn]];
            mainGridRange.Borders.Color = Exc.XlRgbColor.rgbBlack;
            mainWorkSheet.Cells[1, 1].Select();
            mainWorkSheet.Columns[4].Hidden = true;
            mainWorkSheet.Columns[6].Hidden = true;
            mainWorkSheet.Columns[8].Hidden = true;
            mainWorkSheet.Columns[10].Hidden = true;
            mainWorkSheet.Columns[11].Hidden = true;
            mainWorkSheet.Columns[12].Hidden = true;
        }
        #endregion

        #region Format Fund Work Sheet After Data
        private void FormatFundWorkSheetAfterData(Exc.Worksheet fundWorkSheet)
        {
            fundWorkSheet.Rows[1].AutoFit();
            

            fundWorkSheet.Columns.AutoFit();
            fundWorkSheet.Columns[2].ColumnWidth = 13;
            fundWorkSheet.Columns[3].ColumnWidth = 13;
            fundWorkSheet.Columns[4].Hidden = true;
            fundWorkSheet.Columns[5].Hidden = true;
            fundWorkSheet.Columns[6].Hidden = true;            
            fundWorkSheet.Columns[7].ColumnWidth = 13;
            fundWorkSheet.Columns[8].ColumnWidth = 13;
            fundWorkSheet.Columns[9].ColumnWidth = 13;
            fundWorkSheet.Columns[10].ColumnWidth = 13;
            fundWorkSheet.Columns[11].ColumnWidth = 13;
            fundWorkSheet.Columns[14].Hidden = true;
            fundWorkSheet.Columns[15].Hidden = true;
            fundWorkSheet.Columns[16].Hidden = true;
            fundWorkSheet.Columns[17].Hidden = true;
            fundWorkSheet.Columns[18].Hidden = true;
            fundWorkSheet.Columns[19].Hidden = true;
            fundWorkSheet.Columns[20].Hidden = true;
            fundWorkSheet.Columns[21].Hidden = true;
            fundWorkSheet.Columns[22].Hidden = true;
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
            int daysToLiquidateCap)
        {
            LiquidityCalculatorClient client = new LiquidityCalculatorClient();

            TimeSpan difference = DateTime.Now.Date.Subtract(dateToFind);

            Dictionary<string, LiquidityCalculatorOutput> outputs = client.Calculate(new LiquidityCalculatorInput(
                percentageOfPortfolioToLiquidate, percentageOfDailyVolume, fine, fineCap, numberOfDays, daysToLiquidateCap), difference.Days);

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

            AddParameterValues(dateToFind, percentageOfPortfolioToLiquidate, percentageOfDailyVolume, fine, fineCap, numberOfDays, daysToLiquidateCap);
            int maxFundColumn = 22;
            int row = 2;
            mainWorkSheet.Range[mainWorkSheet.Cells[row, 1], mainWorkSheet.Cells[29, maxColumn]].ClearContents();
            int sheetNumber = 2;
            Exc.Worksheet fundWorkSheet = mainWorkSheet;
            foreach (KeyValuePair<string, LiquidityCalculatorOutput> output in outputs.OrderByDescending(a => a.Value.UnsoldValue))
            {
                AddMainPageValues(output.Key, output.Value, mainWorkSheet,row, maxColumn);

                if (output.Value.NonLiquidatedPortfolio.Count > 0)
                {

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
                    AddFundPageValues(fundWorkSheet, maxFundColumn, output.Value, ref fundRow, row, daysToLiquidateCap, numberOfDays);

                    fundRow++;

                    AddFundExceptionValues(fundWorkSheet, maxFundColumn, output.Value, ref fundRow);
                    FormatFundWorkSheetAfterData(fundWorkSheet);
                    sheetNumber++;
                }
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
