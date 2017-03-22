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
using System.Diagnostics;
using Microsoft.Office.Tools.Excel;

namespace Odey.ExcelAddin
{
    public class WatchListItem
    {
        public string Ticker { get; set; }
        public string JHManagerOverride { get; set; }
    }

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
            
            //try
            //{
            app.StatusBar = "Loading portfolio weightings...";
            //var Funds = new StaticDataClient().GetAllFunds().ToDictionary(f => f.EntityId);
            var data = new PortfolioCacheClient().GetPortfolioExposures(new PortfolioRequestObject
            {
                FundIds = funds.Cast<int>().ToArray(),
                ReferenceDates = new[] { DateTime.Today },
            });

            app.StatusBar = "Reading watch list...";
            var watchList = GetWatchList(app, data);

            app.StatusBar = "Writing weightings...";
            WriteWeightings(app, data, watchList);

            // Refresh all
            app.StatusBar = "Refreshing queries...";
            Globals.ThisAddIn.Application.ActiveWorkbook.RefreshAll();

            app.StatusBar = null;
            //}
            //catch (Exception e)
            //{
            //    app.StatusBar = e.Message;
            //}
        }

        #endregion

        #region Helpers

        private Dictionary<string, WatchListItem> GetWatchList(Excel.Application app, List<PortfolioDTO> data)
        {
            const string sheetName = "Watch List";

            const int tickerColumn = 1;
            const string tickerColumnName = "TICKER";
            const int managerColumn = 50;
            const string managerColumnName = "Portfolio Manager";

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
                sheet.Cells[headerRow, tickerColumn] = tickerColumnName;
                sheet.Cells[headerRow, managerColumn] = managerColumnName;
            }

            EnsureText(sheet, headerRow, tickerColumn, tickerColumnName);
            EnsureText(sheet, headerRow, managerColumn, managerColumnName);

            // Read existing tickers
            var watchList = new Dictionary<string, WatchListItem>();
            var row = headerRow + 1;
            var ticker = ReadText(sheet, row, tickerColumn);
            while (ticker != null)
            {
                watchList.Add(ticker, new WatchListItem
                {
                    Ticker = ticker,
                    JHManagerOverride = ReadText(sheet, row, managerColumn),
                });
                ticker = ReadText(sheet, ++row, tickerColumn);
            }

            // Add new tickers
            var newTickers = data.Select(p => p.BloombergTicker).Distinct().Where(t => t != null).Except(watchList.Keys).OrderBy(t => t);
            foreach (var newTicker in newTickers)
            {
                sheet.Cells[row, tickerColumn] = newTicker;
                ++row;
            }

            return watchList;
        }

        private string ReadText(Excel.Worksheet sheet, int row, int column)
        {
            var ret = sheet.Cells[row, column].Text;
            if (string.IsNullOrWhiteSpace(ret))
            {
                return null;
            }
            return ret;
        }

        private void EnsureText(Excel.Worksheet sheet, int row, int column, string text)
        {
            Excel.Range c = sheet.Cells[row, column];
            if (c.Text != text)
            {
                throw new Exception($"Unexpected value '{c.Text}'. Expected '{text}'");
            }
        }

        private void WriteWeightings(Excel.Application app, List<PortfolioDTO> data, Dictionary<string, WatchListItem> watchList)
        {
            Worksheet sheet;
            try
            {
                sheet = Globals.Factory.GetVstoObject(app.Sheets["Weightings"]);
            }
            catch
            {
                sheet = Globals.Factory.GetVstoObject(app.Sheets.Add(After: app.Sheets[app.Sheets.Count]));
                sheet.Name = "Weightings";
            }

            // Fund NAVs
            var navs = data.GroupBy(p => p.FundId).ToDictionary(g => (FundIds)g.Key, g => g.Select(p => p.FundNAV).Distinct().Single());

            // Get table
            ListObject lov = null;
            foreach (Excel.ListObject lo in sheet.ListObjects)
            {
                if (lo.Name == "weightings")
                {
                    Debug.WriteLine("Found existing list object");
                    lov = Globals.Factory.GetVstoObject(lo);
                }
            }

            if (lov == null)
            {
                Debug.WriteLine("Creating new list object");
                Excel.Range range = sheet.Cells[1, 1];
                lov = sheet.Controls.AddListObject(range, "weightings");
            }

            var rows = data
                .Where(p => p.ExposureTypeId == ExposureTypeIds.Primary)
                .OrderBy(p => p.BloombergTicker)
                .ToLookup(p => new { p.EquivalentInstrumentMarketId, p.BloombergTicker, p.ManagerName, p.StrategyName })
                .Select(g => {
                    var ticker = g.Key.BloombergTicker;
                    var nonCFD = g.FirstOrDefault(p => p.InstrumentClassId != (int)InstrumentClassIds.ContractForDifference);

                    // Manager
                    var manager = g.Key.ManagerName;
                    if (manager == "James Hanbury" && ticker != null && watchList.ContainsKey(ticker))
                    {
                        var item = watchList[ticker];
                        if (item.JHManagerOverride != null)
                        {
                            // Manager override
                            manager = item.JHManagerOverride;
                        }
                    }

                    // Strategy
                    var strategy = g.Key.StrategyName;
                    if (strategy == "None")
                    {
                        strategy = null;
                    }

                    var items = g.ToLookup(p => (FundIds)p.FundId);

                    return new
                    {
                        Ticker = ticker,
                        Name = (nonCFD ?? g.First()).InstrumentName,
                        Manager = manager,
                        Strategy = strategy,
                        ARFF = items[FundIds.ARFF].Any() ? items[FundIds.ARFF].Sum(p => p.Exposure) / navs[FundIds.ARFF] : null,
                        BVFF = items[FundIds.BVFF].Any() ? items[FundIds.BVFF].Sum(p => p.Exposure) / navs[FundIds.BVFF] : null,
                        DEVM = items[FundIds.DEVM].Any() ? items[FundIds.DEVM].Sum(p => p.Exposure) / navs[FundIds.DEVM] : null,
                        FDXH = items[FundIds.FDXH].Any() ? items[FundIds.FDXH].Sum(p => p.Exposure) / navs[FundIds.FDXH] : null,
                        OAR  = items[FundIds.OUAR].Any() ? items[FundIds.OUAR].Sum(p => p.Exposure) / navs[FundIds.OUAR] : null,
                    };
                })
                .ToList();

            lov.AutoSetDataBoundColumnHeaders = true;
            lov.SetDataBinding(rows);
            lov.ListColumns["Ticker"].Range.ColumnWidth = 22;
            lov.ListColumns["Name"].Range.ColumnWidth = 35;
            lov.ListColumns["Manager"].Range.ColumnWidth = 20;
            lov.ListColumns["Strategy"].Range.ColumnWidth = 20;
            lov.ListColumns["ARFF"].Range.NumberFormat = "0.00%";
            lov.ListColumns["BVFF"].Range.NumberFormat = "0.00%";
            lov.ListColumns["DEVM"].Range.NumberFormat = "0.00%";
            lov.ListColumns["FDXH"].Range.NumberFormat = "0.00%";
            lov.ListColumns["OAR"].Range.NumberFormat = "0.00%";
            lov.Disconnect();

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
