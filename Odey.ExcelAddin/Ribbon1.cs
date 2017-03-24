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
        private Dictionary<FundIds, decimal?> NAVs;

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

            // Cache fund NAVs
            NAVs = data.GroupBy(p => p.FundId).ToDictionary(g => (FundIds)g.Key, g => g.Select(p => p.FundNAV).Distinct().Single());

            app.StatusBar = "Reading watch list...";
            var watchList = GetWatchList(app, data);

            // Apply watch list
            foreach (var row in data)
            {
                if (row.StrategyName == "None")
                {
                    row.StrategyName = null;
                }

                if (row.BloombergTicker != null && row.ManagerId == (int)ApplicationUserIds.JamesHanbury)
                {
                    // Automatic DEVM & FDXH manager override
                    if (row.FundId == (int)FundIds.DEVM || row.FundId == (int)FundIds.FDXH)
                    {
                        var others = data.Where(p => p.BloombergTicker == row.BloombergTicker && p.ManagerId != (int)ApplicationUserIds.JamesHanbury && p.FundId != (int)FundIds.DEVM && p.FundId != (int)FundIds.FDXH).ToList();
                        var ids = others.Select(p => p.ManagerId).Distinct();
                        if (ids.Count() == 1)
                        {
                            row.ManagerId = ids.Single();
                            row.ManagerName = others.Select(p => p.ManagerName).First();
                        }
                    }

                    // Manual manager override
                    if (watchList.ContainsKey(row.BloombergTicker))
                    {
                        var item = watchList[row.BloombergTicker];
                        if (item.JHManagerOverride != null)
                        {
                            row.ManagerName = item.JHManagerOverride;
                            row.ManagerId = -1;
                        }
                    }
                }
            }

            // Write sheets
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
            //    throw;
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
            
            //for (var i = 1; i < 100; ++i)
            //{
            //    Excel.Range cell = sheet.Cells[headerRow, i];
            //    Excel.Range cell2 = sheet.Cells[headerRow + 1, i];

            //    Debug.WriteLine($"{cell.Address} {cell.Text} {cell2.Formula}");
            //}

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
            var newTickers = data.Select(p => p.BloombergTicker).Distinct().Where(t => t != null).Except(watchList.Keys).OrderBy(t => t).ToList();
            foreach (var newTicker in newTickers)
            {
                watchList.Add(newTicker, new WatchListItem
                {
                    Ticker = newTicker,
                });
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

        private void WritePortfolio(Excel.Application app, List<PortfolioDTO> data, Dictionary<string, WatchListItem> watchList, FundIds fundId)
        {
            var rows = data
                .Where(p => p.ExposureTypeId == ExposureTypeIds.Primary)
                .Where(p => p.FundId == (int)fundId)
                .OrderBy(p => p.BloombergTicker)
                .ToLookup(p => new { p.EquivalentInstrumentMarketId, p.BloombergTicker })
                .Select(g => new {
                    Ticker = g.Key.BloombergTicker,
                    Name = g.Select(p => p.Issuer).Distinct().Single(),
                    Exposure = g.Sum(p => p.Exposure) / NAVs[fundId],
                })
                .ToList();

            var sheet = app.GetOrCreateVstoWorksheet($"Portfolio - {fundId}");
            var lo = sheet.GetOrCreateListObject($"Portfolio - {fundId}");

            lo.AutoSetDataBoundColumnHeaders = true;
            lo.SetDataBinding(rows);
            lo.ListColumns["Ticker"].Range.ColumnWidth = 22;
            lo.ListColumns["Name"].Range.ColumnWidth = 35;
            lo.ListColumns["Exposure"].Range.NumberFormat = "0.00%";
            lo.Disconnect();
        }


        private void WriteWeightings(Excel.Application app, List<PortfolioDTO> data, Dictionary<string, WatchListItem> watchList)
        {
            var rows = data
                .Where(p => p.ExposureTypeId == ExposureTypeIds.Primary)
                .OrderBy(p => p.BloombergTicker)
                .ToLookup(p => new { p.EquivalentInstrumentMarketId, p.BloombergTicker, p.Issuer, p.ManagerName })
                .Select(g => {
                    var nonCFD = g.FirstOrDefault(p => p.InstrumentClassId != (int)InstrumentClassIds.ContractForDifference);
                    var items = g.ToLookup(p => (FundIds)p.FundId);
                    return new
                    {
                        Ticker = g.Key.BloombergTicker,
                        Issuer = g.Key.Issuer,
                        Manager = g.Key.ManagerName,
                        ARFF = items[FundIds.ARFF].Any() ? items[FundIds.ARFF].Sum(p => p.Exposure) / NAVs[FundIds.ARFF] : null,
                        BVFF = items[FundIds.BVFF].Any() ? items[FundIds.BVFF].Sum(p => p.Exposure) / NAVs[FundIds.BVFF] : null,
                        DEVM = items[FundIds.DEVM].Any() ? items[FundIds.DEVM].Sum(p => p.Exposure) / NAVs[FundIds.DEVM] : null,
                        FDXH = items[FundIds.FDXH].Any() ? items[FundIds.FDXH].Sum(p => p.Exposure) / NAVs[FundIds.FDXH] : null,
                        OAR  = items[FundIds.OUAR].Any() ? items[FundIds.OUAR].Sum(p => p.Exposure) / NAVs[FundIds.OUAR] : null,
                    };
                })
                .ToList();

            var sheet = app.GetOrCreateVstoWorksheet("Weightings");
            var lo = sheet.GetOrCreateListObject("weightings", 1, 1);
            lo.AutoSetDataBoundColumnHeaders = true;
            lo.SetDataBinding(rows);
            lo.ListColumns["Ticker"].Range.ColumnWidth = 22;
            lo.ListColumns["Issuer"].Range.ColumnWidth = 35;
            lo.ListColumns["Manager"].Range.ColumnWidth = 20;
            lo.ListColumns["ARFF"].Range.NumberFormat = "0.00%";
            lo.ListColumns["BVFF"].Range.NumberFormat = "0.00%";
            lo.ListColumns["DEVM"].Range.NumberFormat = "0.00%";
            lo.ListColumns["FDXH"].Range.NumberFormat = "0.00%";
            lo.ListColumns["OAR"].Range.NumberFormat = "0.00%";
            lo.Disconnect();
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
