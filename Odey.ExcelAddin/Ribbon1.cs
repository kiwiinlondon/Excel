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
using System.Windows.Forms;

namespace Odey.ExcelAddin
{
    public class WatchListItem
    {
        public int RowIndex { get; set; }
        public string Ticker { get; set; }
        public string Quality { get; set; }
        public string JHManagerOverride { get; set; }
        public string Conviction { get; set; }
        public double? Upside { get; set; }
    }

    public static class WatchListColumns
    {
        public static ColumnDef Ticker = new ColumnDef { Index = 1, Name = "TICKER" };
        public static ColumnDef Upside = new ColumnDef { Index = 20, Name = "TICKER" };
        public static ColumnDef Quality = new ColumnDef { Index = 49, Name = "TICKER" };
        public static ColumnDef Manager = new ColumnDef { Index = 50, Name = "TICKER" };
        public static ColumnDef Conviction = new ColumnDef { Index = 51, Name = "TICKER" };
    }

    public class ExposureItem
    {
        public string Ticker { get; set; }
        public string Issuer { get; set; }
        public string Manager { get; set; }
        public string Fund { get; set; }
        public int? FundId { get; set; }
        public decimal? PercentNAV { get; set; }
        public decimal? NetPosition { get; set; }
    }

    public class ColumnDef
    {
        public int? Index { get; set; }
        public string Name { get; set; }
        public string Formula { get; set; }
        public string NumberFormat { get; set; }
        public double Width { get; set; }
    }

    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private static FundIds[] funds = new[] { FundIds.ARFF, FundIds.BVFF, FundIds.DEVM, FundIds.FDXH, FundIds.OUAR };

        private static Dictionary<string, string> managerInitials = new Dictionary<string, string>
        {
            { "Adrian Courtenay", "AC" },
            { "Jamie Grimston", "JG" },
            { "James Hanbury", "JH" },
        };

        private static List<ColumnDef> PortfolioColumns = new List<ColumnDef>
        {
            new ColumnDef
            {
                Name = "NAME",
                Formula = "=BDP([Ticker],\"SHORT_NAME\",\"Fill=B\")",
                Width = 18.29,
            },
            new ColumnDef
            {
                Name = "SECTOR",
                Formula = "=BDP([Ticker],\"GICS_SECTOR_NAME\",\"Fill=B\")",
                Width = 26.14,
            },
            new ColumnDef
            {
                Name = "COUNTRY",
                Formula = "=BDP([Ticker],\"COUNTRY_FULL_NAME\",\"Fill=B\")",
                Width = 16.14,
            },
            //new ColumnDef
            //{
            //    Name = "UPSIDE",
            //    Formula = "=(H15-I15)/I15",
            //    Width = 12.29,
            //    NumberFormat = "0%",
            //},
            //new ColumnDef
            //{
            //    Name = "BASIS for TARGET PRICE",
            //    Formula = "",
            //    Width = 15.43,
            //},
            //new ColumnDef
            //{
            //    Name = "TARGET PRICE",
            //    Formula = "=VLOOKUP([Ticker], Watch_List_Table, 5, FALSE) & \"\"",
            //    Width = 12.29,
            //    NumberFormat = "#,##0.00",
            //},
            new ColumnDef
            {
                Name = "PRICE",
                Formula = "=BDP([Ticker],\"PX_LAST\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "#,##0.00",
            },
            //new ColumnDef
            //{
            //    Name = "UPSIDE/Weight(%)",
            //    Formula = "=IFERROR([UPSIDE]/[PercentNAV], \"\")",
            //    Width = 12.29,
            //    NumberFormat = "0.00",
            //},
            new ColumnDef
            {
                Name = "MARKET CAP $ (Mln)",
                Formula = "=BDP([Ticker],\"CRNCY_ADJ_MKT_CAP\",\"EQY_FUND_CRNCY\",\"USD\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "#,##0",
            },
            new ColumnDef
            {
                Name = "NET DEBT Inc PENSIONS (Mln)",
                Formula = "=BDP([Ticker],\"NET_DEBT_ADJ_FOR_PENSION_PR_LIAB\",\"SCALING_FORMAT\",\"MLN\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "#,##0",
            },
            new ColumnDef
            {
                Name = "ENTERPRISE VALUE",
                Formula = "=BDP([Ticker],\"CURR_ENTP_VAL\",\"SCALING_FORMAT\",\"MLN\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "#,##0",
            },
            new ColumnDef
            {
                Name = "ENTERPRISE VALUE EST",
                Formula = "=BDP([Ticker],\"BEST_EV\",\"BEST_FPERIOD_OVERRIDE\",\"1BF\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "#,##0",
            },
            new ColumnDef
            {
                Name = "DIVIDEND YIELD",
                Formula = "=BDP([Ticker],\"EQY_DVD_YLD_IND\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "#,##0.00",
            },
            new ColumnDef
            {
                Name = "DIVIDEND YIELD EST",
                Formula = "=BDP([Ticker],\"BEST_DIV_YLD\",\"BEST_FPERIOD_OVERRIDE\",\"1BF\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "0.00",
            },
            new ColumnDef
            {
                Name = "EBIT (Mln)",
                Formula = "=BDP([Ticker],\"EBIT\",\"SCALING_FORMAT\",\"MLN\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "0",
            },
            new ColumnDef
            {
                Name = "EBIT EST",
                Formula = "=BDP([Ticker],\"BEST_EBIT\",\"BEST_FPERIOD_OVERRIDE=1BF\",\"SCALING_FORMAT\",\"MLN\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "0",
            },
            //new ColumnDef
            //{
            //    Name = "EV/EBIT",
            //    Formula = "=IFERROR(M15/Q15,\"\")",
            //    Width = 12.29,
            //    NumberFormat = "#,##0.00",
            //},
            new ColumnDef
            {
                Name = "EV/EBIT EST",
                Formula = "=BDP([Ticker],\"BEST_EV_TO_BEST_EBIT\",\"BEST_FPERIOD_OVERRIDE=1BF\",\"SCALING_FORMAT\",\"MLN\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "#,##0.00",
            },
            new ColumnDef
            {
                Name = "Sales",
                Formula = "=BDP([Ticker],\"SALES_REV_TURN\",\"SCALING_FORMAT\",\"MLN\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "0",
            },
            new ColumnDef
            {
                Name = "Sales EST",
                Formula = "=BDP([Ticker],\"BEST_SALES\", \"BEST_FPERIOD_OVERRIDE=1BF\",\"SCALING_FORMAT\",\"MLN\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "0",
            },
            //new ColumnDef
            //{
            //    Name = "EV/Sales",
            //    Formula = "=IFERROR(M15/U15,\"\")",
            //    Width = 12.29,
            //    NumberFormat = "#,##0.00",
            //},
            //new ColumnDef
            //{
            //    Name = "EV/Sales EST",
            //    Formula = "=IFERROR(N15/V15,\"\")",
            //    Width = 12.29,
            //    NumberFormat = "#,##0.00",
            //},
            new ColumnDef
            {
                Name = "TRAIL 12M EPS",
                Formula = "=BDP([Ticker],\"TRAIL_12M_EPS_BEF_XO_ITEM\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "0.0",
            },
            new ColumnDef
            {
                Name = "EPS EST",
                Formula = "=BDP([Ticker],\"BEST_EPS\",\"BEST_FPERIOD_OVERRIDE=1BF\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "0.0",
            },
            new ColumnDef
            {
                Name = "P/E Ratio",
                Formula = "=BDP([Ticker],\"PE_RATIO\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "0.0",
            },
            new ColumnDef
            {
                Name = "P/E Ratio EST",
                Formula = "=BDP([Ticker],\"BEST_PE_RATIO\",\"BEST_FPERIOD_OVERRIDE=1BF\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "0.0",
            },
            new ColumnDef
            {
                Name = "Book Value Per SH",
                Formula = "=BDP([Ticker],\"BOOK_VAL_PER_SH\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "0.0",
            },
            new ColumnDef
            {
                Name = "Book Value Per SH EST",
                Formula = "=BDP([Ticker],\"BEST_BPS\",\"BEST_FPERIOD_OVERRIDE=1Bf\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "0.0",
            },
            new ColumnDef
            {
                Name = "P/NAV",
                Formula = "=BDP([Ticker],\"PX_TO_BOOK_RATIO\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "0.0",
            },
            new ColumnDef
            {
                Name = "P/NAV EST",
                Formula = "=BDP([Ticker],\"BEST_PX_BPS_RATIO\",\"BEST_FPERIOD_OVERRIDE=1BF\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "0.0",
            },
            new ColumnDef
            {
                Name = "Tang Book Value Per SH",
                Formula = "=BDP([Ticker],\"TANG_BOOK_VAL_PER_SH\",\"Fill=B\")",
                Width = 0,
                NumberFormat = "0.0",
            },
            new ColumnDef
            {
                Name = "P/TNAV",
                Formula = "=BDP([Ticker],\"PX_TO_TANG_BV_PER_SH\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "0.0",
            },
            //new ColumnDef
            //{
            //    Name = "EV/EBITDA",
            //    Formula = "",
            //    Width = 9.43,
            //    NumberFormat = "General",
            //},
            new ColumnDef
            {
                Name = "60-day beta (MSCI world TR relevant currency for fund)",
                Formula = "=BDP([Ticker],\"BETA_ADJ_OVERRIDABLE\",\"BETA_OVERRIDE_REL_INDEX=gdduwi index\",\"BETA_OVERRIDE_START_DT\",TEXT('Watch List'!BO6,\"YYYYMMDD\"),\"BETA_OVERRIDE_PERIOD=d\",\"Fill=B\")",
                Width = 10.71,
                NumberFormat = "#,##0",
            },
        };

        public Ribbon1()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Odey.ExcelAddin.Ribbon1.xml");
        }

        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnActionCallback(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;

            var prevScreenUpdating = app.ScreenUpdating;
            var prevEvents = app.EnableEvents;
            var prevCalculation = app.Calculation;
            app.ScreenUpdating = false;
            app.EnableEvents = false;
            app.Calculation = Excel.XlCalculation.xlCalculationManual;

            try
            {
                app.StatusBar = "Loading portfolio weightings...";
                var data = new PortfolioCacheClient().GetPortfolioExposures(new PortfolioRequestObject
                {
                    FundIds = funds.Cast<int>().ToArray(),
                    ReferenceDates = new[] { DateTime.Today },
                });
                //var Funds = new StaticDataClient().GetAllFunds().ToDictionary(f => f.EntityId);

                var watchList = GetWatchList(app, data);
                ApplyManagerOverrides(data, watchList);
                WriteWatchList(app, watchList, "Watch List Top", true);
                WriteWatchList(app, watchList, "Watch List Bottom", false);
                WriteWatchList(app, watchList, "Watch List High Quality", true, "H");
                WriteWatchList(app, watchList, "Watch List Low Quality", false, "L");
                foreach (var fund in funds)
                {
                    //WriteExposureSheet(app, fund, data, watchList);
                }
                foreach (var fund in funds)
                {
                    WritePortfolioSheet(app, fund, data, watchList);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, e.GetType().Name);
            }
            finally
            {
                app.StatusBar = null;
                app.EnableEvents = prevEvents;
                app.ScreenUpdating = prevScreenUpdating;
                app.Calculation = prevCalculation;
            }
        }

        private void ApplyManagerOverrides(List<PortfolioDTO> data, Dictionary<string, WatchListItem> watchList)
        {
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
        }

        private void WriteWatchList(Excel.Application app, Dictionary<string, WatchListItem> watchList, string sheetName, bool descending, string onlyQuality = null)
        {
            const int topX = 30;

            var rows = watchList.Values.Where(w => w.Upside.HasValue);
            if (onlyQuality != null)
            {
                rows = rows.Where(w => w.Quality == onlyQuality);
            }
            if (descending)
            {
                rows = rows.OrderByDescending(w => w.Upside);
            }
            else
            {
                rows = rows.OrderBy(w => w.Upside);
            }
            rows = rows.Take(topX);

            var sheet = app.GetOrCreateVstoWorksheet(sheetName);

            // Clear data
            var columnList = new[] { "B", "E", "F", "S", "T", "U", "W", "Z", "AD", "AI", "AM", "AQ", "AS", "AT", "AW", "BK" };
            var y = 14;
            Excel.Range r = sheet.Range[sheet.Cells[y, 1], sheet.Cells[y + topX, 1 + columnList.Length]];
            r.ClearContents();

            // Write header
            Excel.Range headerRange = sheet.Range[sheet.Cells[y, 1], sheet.Cells[y, 1 + columnList.Length]];
            headerRange.WrapText = true;
            headerRange.RowHeight = 75;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Excel.Range cell = sheet.Cells[y, 1];
            cell.Value = "Ticker";
            cell.ColumnWidth = 14;
            var x = 2;
            foreach (var columnIndex in columnList)
            {
                cell = sheet.Cells[y, x];
                cell.Formula = $"='Watch List'!{columnIndex}{5}";
                cell.ColumnWidth = 14;
                ++x;
            }
            ++y;

            // Write content
            foreach (var row in rows)
            {
                sheet.Cells[y, 1] = row.Ticker;
                x = 2;
                foreach (var columnIndex in columnList)
                {
                    sheet.Cells[y, x].Formula = $"='Watch List'!{columnIndex}{row.RowIndex}";
                    ++x;
                }
                ++y;
            }
        }

        private void WriteExposureSheet(Excel.Application app, FundIds fundId, List<PortfolioDTO> weightings, Dictionary<string, WatchListItem> watchList)
        {
            app.StatusBar = $"Writing {fundId} exposure sheet...";

            var rows = weightings
                .Where(p => p.ExposureTypeId == ExposureTypeIds.Primary && p.BloombergTicker != null)
                .ToLookup(p => new { p.EquivalentInstrumentMarketId, p.BloombergTicker, p.ManagerName, p.FundId })
                .Select(g => new
                {
                    Ticker = g.Key.BloombergTicker,
                    Manager = managerInitials.ContainsKey(g.Key.ManagerName) ? managerInitials[g.Key.ManagerName] : g.Key.ManagerName,
                    FundId = g.Key.FundId,
                    PercentNAV = g.Sum(p => p.Exposure) / g.Select(p => p.FundNAV).Distinct().Single(),
                    NetPosition = g.Sum(p => p.NetPosition),
                })
                .ToList();
            var sheet = app.GetOrCreateVstoWorksheet($"Exposure {fundId}");

            var managers = new Dictionary<string, int>
            {
                { "JH", 34 },
                { "AC", 10 },
                { "JG", 12 },
            };

            var row = 1;
            foreach (var manager in managers)
            {
                sheet.GetCell(row, 1).Value = managerInitials.Single(x => x.Value == manager.Key).Key;
                ++row;

                var fund = rows.Where(x => x.Manager == manager.Key && x.FundId == (int)fundId);
                var longs = fund.Where(x => x.PercentNAV > 0).OrderByDescending(x => x.PercentNAV).Take(manager.Value).Select((x, j) => new
                {
                    Rank = j + 1,
                    Long = x.Ticker,
                    PercentNAV = x.PercentNAV,
                    NetPosition = x.NetPosition
                }).ToList();
                var shorts = fund.Where(x => x.PercentNAV < 0).OrderBy(x => x.PercentNAV).Take(manager.Value).Select((x, j) => new
                {
                    Rank = j + 1,
                    Short = x.Ticker,
                    PercentNAV = x.PercentNAV,
                    NetPosition = x.NetPosition
                }).ToList();


                var column = 1;

                if (longs.Any())
                {
                    var tName = $"Exposure_{fundId}_{manager.Key}_Long";
                    var longTable = sheet.GetListObject(tName);
                    if (longTable == null)
                    {
                        longTable = sheet.CreateListObject(tName, row, column);
                        longTable.ShowTableStyleRowStripes = false;
                    }
                    longTable.AutoSetDataBoundColumnHeaders = true;
                    longTable.SetDataBinding(longs);
                    longTable.ListColumns["Long"].Range.ColumnWidth = 22;
                    longTable.ListColumns["PercentNAV"].Range.ColumnWidth = 14;
                    longTable.ListColumns["PercentNAV"].Range.NumberFormat = "0.00%";
                    longTable.ListColumns["NetPosition"].Range.NumberFormat = "#,###";
                    longTable.ListColumns["NetPosition"].Range.ColumnWidth = 15;
                    longTable.Disconnect();
                    column += longTable.ListColumns.Count + 1;
                }

                if (shorts.Any())
                {
                    var tName = $"Exposure_{fundId}_{manager.Key}_Short";
                    var shortTable = sheet.GetListObject(tName);
                    if (shortTable == null)
                    {
                        shortTable = sheet.CreateListObject(tName, row, column);
                        shortTable.ShowTableStyleRowStripes = false;
                    }
                    shortTable.AutoSetDataBoundColumnHeaders = true;
                    shortTable.SetDataBinding(shorts);
                    shortTable.ListColumns["Short"].DataBodyRange.ColumnWidth = 22;
                    shortTable.ListColumns["PercentNAV"].DataBodyRange.ColumnWidth = 14;
                    shortTable.ListColumns["PercentNAV"].DataBodyRange.NumberFormat = "0.00%";
                    shortTable.ListColumns["NetPosition"].DataBodyRange.NumberFormat = "#,###";
                    shortTable.ListColumns["NetPosition"].Range.EntireColumn.Hidden = true;
                    shortTable.Disconnect();

                    // Daily Vol
                    var col = shortTable.ListColumns.Add();
                    col.Name = "% Daily Volume";
                    var AverageVolumeAllExchanges3M = "BDP([Short], \"INTERVAL_AVG\", \"MARKET_DATA_OVERRIDE=pq718\", \"CALC_INTERVAL=3m\")";
                    col.DataBodyRange.Formula = $"=[NetPosition]/{AverageVolumeAllExchanges3M}";
                    col.DataBodyRange.NumberFormat = "0.00%";
                    col.Range.ColumnWidth = 15;

                    // Conviction
                    var col2 = shortTable.ListColumns.Add();
                    col2.Name = "Conviction";
                    var convictionColumn = 51;
                    col2.DataBodyRange.Formula = $"=VLOOKUP([Short], Watch_List_Table, {convictionColumn}, FALSE) & \"\"";

                    column += shortTable.ListColumns.Count;
                }

                if (column > 1)
                {
                    Excel.Range header = sheet.Range[sheet.Cells[row - 1, 1], sheet.Cells[row - 1, column - 1]];
                    //header.Merge();
                }

                row += manager.Value + 5;
            }
        }
        
        private void WritePortfolioSheet(Excel.Application app, FundIds fundId, List<PortfolioDTO> weightings, Dictionary<string, WatchListItem> watchList)
        {
            app.StatusBar = $"Writing {fundId} portfolio sheet...";

            var rows = weightings
                .Where(p => p.ExposureTypeId == ExposureTypeIds.Primary && p.BloombergTicker != null && p.FundId == (int)fundId)
                .ToLookup(p => new { p.EquivalentInstrumentMarketId, p.BloombergTicker, p.ManagerName })
                .Select(g => new
                {
                    Ticker = g.Key.BloombergTicker,
                    Manager = managerInitials.ContainsKey(g.Key.ManagerName) ? managerInitials[g.Key.ManagerName] : g.Key.ManagerName,
                    PercentNAV = g.Sum(p => p.Exposure) / g.Select(p => p.FundNAV).Distinct().Single(),
                })
                .ToList();

            var sheet = app.GetOrCreateVstoWorksheet($"Portfolio {fundId}");

            var tName = $"Portfolio_{fundId}";
            var table = sheet.GetListObject(tName);
            if (table == null)
            {
                table = sheet.CreateListObject(tName, 14, 1);
                table.ShowTableStyleRowStripes = false;
                table.ShowTableStyleFirstColumn = true;
                table.AutoSetDataBoundColumnHeaders = true;
                table.HeaderRowRange.WrapText = true;
                table.HeaderRowRange.RowHeight = 75;
                table.HeaderRowRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                table.HeaderRowRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }

            table.SetDataBinding(rows);
            table.ListColumns["Ticker"].DataBodyRange.ColumnWidth = 22;
            table.ListColumns["PercentNAV"].DataBodyRange.ColumnWidth = 14;
            table.ListColumns["PercentNAV"].DataBodyRange.NumberFormat = "0.00%";
            table.Disconnect();

            foreach (var column in PortfolioColumns)
            {
                var col = table.ListColumns.Add();
                col.Name = column.Name;
                col.Range.ColumnWidth = column.Width;
                if (column.NumberFormat != null)
                {
                    col.Range.NumberFormat = column.NumberFormat;
                }
                col.DataBodyRange.Formula = column.Formula.Replace("[Ticker]", "$A15");
            }

        }

        private Dictionary<string, WatchListItem> GetWatchList(Excel.Application app, List<PortfolioDTO> data)
        {
            app.StatusBar = "Reading watch list...";
            const string sheetName = "Watch List";
            const int headerRow = 5;

            Excel.Worksheet sheet;
            try
            {
                // Get exisiting sheet
                sheet = app.Sheets[sheetName];

                //// Make sure the header is in the right format
                //EnsureText(sheet, headerRow, WatchListColumns["Ticker"].Index, tickerColumnName);
                //EnsureText(sheet, headerRow, managerColumn, managerColumnName);
                //EnsureText(sheet, headerRow, convictionColumn, convictionColName);
                //EnsureText(sheet, headerRow, upsideColumn, upsideColName);
            }
            catch
            {
                // Create sheet
                sheet = app.Sheets.Add(Before: app.Sheets[1]);
                sheet.Name = sheetName;

                //// Set header
                //sheet.Cells[headerRow, WatchListColumns["Ticker"].Index] = tickerColumnName;
                //sheet.Cells[headerRow, managerColumn] = managerColumnName;
                //sheet.Cells[headerRow, convictionColumn] = convictionColName;
                //sheet.Cells[headerRow, upsideColumn] = upsideColName;
            }

            // Read existing tickers
            var watchList = new Dictionary<string, WatchListItem>();
            var row = headerRow + 1;
            var ticker = sheet.Cells[row, WatchListColumns.Ticker.Index.Value].Value2 as string;
            while (ticker != null)
            {
                watchList.Add(ticker, new WatchListItem
                {
                    RowIndex = row,
                    Ticker = ticker,
                    Quality = sheet.Cells[row, WatchListColumns.Quality.Index.Value].Value2 as string,
                    JHManagerOverride = sheet.Cells[row, WatchListColumns.Manager.Index.Value].Value2 as string,
                    Conviction = sheet.Cells[row, WatchListColumns.Conviction.Index.Value].Value2 as string,
                    Upside = sheet.Cells[row, WatchListColumns.Upside.Index.Value].Value2 as double?,
                });
                ticker = sheet.Cells[++row, WatchListColumns.Ticker.Index.Value].Value2 as string;
            }
            
            // Add new tickers
            var newTickers = data.Select(p => p.BloombergTicker).Distinct().Where(t => t != null).Except(watchList.Keys).OrderBy(t => t).ToList();
            foreach (var newTicker in newTickers)
            {
                watchList.Add(newTicker, new WatchListItem
                {
                    RowIndex = row,
                    Ticker = newTicker,
                });
                sheet.Cells[row, WatchListColumns.Ticker.Index.Value] = newTicker;
                ++row;
            }

            return watchList;
        }

        private void EnsureText(Excel.Worksheet sheet, int row, int column, string text)
        {
            Excel.Range c = sheet.Cells[row, column];
            if (c.Text != text)
            {
                throw new Exception($"Unexpected value '{c.Text}'. Expected '{text}'");
            }
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

    }
}
