using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Odey.Framework.Keeley.Entities.Enums;
using Odey.PortfolioCache.Clients;
using Odey.PortfolioCache.Entities;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace Odey.ExcelAddin
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private static FundIds[] funds = new[] { FundIds.ARFF, FundIds.BVFF, FundIds.DEVM, FundIds.FDXH, FundIds.OUAR };

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
           
            return GetResourceText("Odey.ExcelAddin.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnActionCallback(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;

            app.StatusBar = "Loading portfolio weightings...";

            //var Funds = new StaticDataClient().GetAllFunds().ToDictionary(f => f.EntityId);
            var data = new PortfolioCacheClient().GetPortfolioExposures(new PortfolioRequestObject
            {
                FundIds = funds.Cast<int>().ToArray(),
                ReferenceDates = new[] { DateTime.Today },
            });


            WriteWatchList(app, data);
            WriteWeightings(app, data);

            app.StatusBar = "Refreshing queries...";

            // Trigger refresh all
            Globals.ThisAddIn.Application.ActiveWorkbook.RefreshAll();

            app.StatusBar = null;
        }

        #endregion

        #region Helpers

        private void WriteWatchList(Excel.Application app, List<PortfolioDTO> data)
        {
            const string sheetName = "Watch List";
            const string tickerColumnName = "Ticker";
            const int headerRow = 5;

            Excel.Worksheet sheet;
            try
            {
                // Get exisiting sheet
                sheet = app.Sheets[sheetName];
            }
            catch
            {
                // Create sheet
                sheet = app.Sheets.Add(Before: app.Sheets[1]);
                sheet.Name = sheetName;

                // Set header
                Excel.Range cell = sheet.Cells[headerRow, 1];
                cell.Value = tickerColumnName;
            }
            
            // Read existing tickers
            var currentTickers = new List<string>();
            var i = 1;
            Excel.Range x = sheet.Cells[headerRow + i, 1];
            while (!string.IsNullOrWhiteSpace(x.Text))
            {
                currentTickers.Add(x.Text);
                ++i;
                x = sheet.Cells[headerRow + i, 1];
            }

            // Add new tickers
            var newTickers = data.Select(p => p.BloombergTicker).Where(t => t != null).Distinct().Except(currentTickers).OrderBy(t => t);
            foreach (var newTicker in newTickers)
            {
                sheet.Cells[headerRow + i, 1] = newTicker;
                ++i;
            }
        }

        private void WriteWeightings(Excel.Application app, List<PortfolioDTO> data)
        {
            Excel.Worksheet sheet;
            try
            {
                sheet = app.Sheets["Weightings"];
            }
            catch
            {
                sheet = app.Sheets.Add(After: app.Sheets[app.Sheets.Count]);
                sheet.Name = "Weightings";
            }

            // Clear content
            sheet.UsedRange.ClearContents();

            // Headers
            sheet.Cells[1, 1] = "Ticker";
            sheet.Cells[1, 2] = "Name";
            sheet.Cells[1, 3] = "Manager";
            sheet.Cells[1, 4] = "Strategy";
            var startColumn = 5;
            foreach (var fund in funds)
            {
                sheet.Cells[1, startColumn + Array.IndexOf(funds, fund)] = fund.ToString();
            }

            // Rows
            var row = 2;

            var instruments = data.Where(p => p.ExposureTypeId == ExposureTypeIds.Primary).OrderBy(p => p.BloombergTicker);
            foreach (var instrument in instruments.ToLookup(p => new { p.EquivalentInstrumentMarketId, p.BloombergTicker, p.ManagerName, p.StrategyName }))
            {
                // Ticker
                var ticker = instrument.Key.BloombergTicker;
                sheet.Cells[row, 1] = ticker;

                // Name
                var nonCFD = instrument.FirstOrDefault(p => p.InstrumentClassId != (int)InstrumentClassIds.ContractForDifference);
                sheet.Cells[row, 2] = (nonCFD ?? instrument.First()).InstrumentName;

                // Manager
                sheet.Cells[row, 3] = instrument.Key.ManagerName;

                // Strategy
                if (instrument.Key.StrategyName != "None")
                {
                    sheet.Cells[row, 4] = instrument.Key.StrategyName;
                }


                // Exposure %NAV
                foreach (var fund in instrument.ToLookup(p => p.FundId))
                {
                    var fundNAV = fund.Select(p => p.FundNAV).Distinct().Single();
                    var column = (startColumn + Array.IndexOf(funds, (FundIds)fund.Key));
                    sheet.Cells[row, column] = fund.Sum(p => p.Exposure) / fundNAV;
                }

                ++row;
            }

            //++row;
            //var currencies = data.Where(p => p.ExposureTypeId == ExposureTypeIds.Currency).OrderBy(p => p.PositionCurrency).ToLookup(p => p.PositionCurrencyId);
            //foreach (var currency in currencies)
            //{
            //    sheet.Cells[row, 2] = currency.Select(p => p.PositionCurrency).Distinct().Single();
            //    foreach (var fund in currency.ToLookup(p => p.FundId))
            //    {
            //        var fundNAV = fund.Select(p => p.FundNAV).Distinct().Single();
            //        var column = (startColumn + Array.IndexOf(funds, (FundIds)fund.Key));
            //        var exposure = fund.Sum(p => p.Exposure);
            //        if (Funds[fund.Key].CurrencyId == currency.Key)
            //        {
            //            exposure -= fundNAV;
            //            Excel.Range cell = sheet.Cells[row, column];
            //            cell.Font.Bold = true;
            //        }
            //        sheet.Cells[row, column] = exposure / fundNAV;
            //    }
            //    ++row;
            //}

            // Set number format
            Excel.Range numbers = sheet.Range[sheet.Cells[2, startColumn], sheet.Cells[row - 1, startColumn + funds.Count()]];
            numbers.NumberFormat = "0.00%";

            // Set column widths
            SetColumnWidth(sheet, 1, 22);
            SetColumnWidth(sheet, 2, 35);
            SetColumnWidth(sheet, 3, 20);
            SetColumnWidth(sheet, 4, 20);
        }

        private void SetColumnWidth(Excel.Worksheet sheet, int columnIndex, int width)
        {
            Excel.Range column = sheet.Range[sheet.Cells[1, columnIndex], sheet.Cells[1, columnIndex]];
            column.ColumnWidth = width;
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
