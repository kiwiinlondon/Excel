using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using Odey.Framework.Keeley.Entities.Enums;
using Odey.PortfolioCache.Entities;
using System;
using System.Drawing;

namespace Odey.ExcelAddin
{
    class ExposureSheet
    {
        private static Dictionary<string, int> TargetItemCountByManager = new Dictionary<string, int>
        {
            { "JH", 25 },
            { "AC", 8 },
            { "JG", 10 },
        };

        public static void Write(Excel.Application app, FundIds fundId, List<PortfolioDTO> weightings, Dictionary<string, WatchListItem> watchList)
        {
            var fundName = Ribbon1.GetFundName(fundId, weightings);
            app.StatusBar = $"Writing {fundName} exposure sheet...";

            var rows = weightings
                .Where(p => p.ExposureTypeId == ExposureTypeIds.Primary && p.BloombergTicker != null && p.FundId == (int)fundId)
                .ToLookup(p => new { p.BloombergTicker, p.ManagerName, p.StrategyName, p.InstrumentClassId })
                .Select(g => new
                {
                    Ticker = g.Key.BloombergTicker,
                    Manager = Ribbon1.GetManagerInitials(g.Key.ManagerName),
                    PercentNAV = g.Sum(p => p.Exposure) / g.Select(p => p.FundNAV).Distinct().Single(),
                    NetPosition = g.Sum(p => p.NetPosition),
                    Strategy = g.Key.StrategyName,
                    InstrumentClassId = g.Key.InstrumentClassId,
                })
                .ToList();

            // Get the worksheet
            var isNewSheet = false;
            var sheetName = $"Exposure {fundName}";
            Excel.Worksheet sheet;
            try
            {
                sheet = app.Sheets[sheetName];
            }
            catch
            {
                isNewSheet = true;
                sheet = app.Sheets.Add();
                sheet.Name = sheetName;
            }

            var headers = new[] {
                new ColumnDef { Name = "#", Width = 3.4 },
                new ColumnDef { Name = "Ticker", Width = 22 },
                new ColumnDef { Name = "% NAV", Width = 7 },
                new ColumnDef { Name = "Net Position", Width = 0 },
                new ColumnDef { Name = "Daily Volume", Width = 0 },
                new ColumnDef { Name = "Upside", Width = 7 },
                new ColumnDef { Name = "Conviction", Width = 9.7 }
            };

            sheet.Cells.ClearContents();
            sheet.Cells.ClearFormats();

            var headerStyle = app.ActiveWorkbook.GetHeaderStyle();
            var rowStyle = app.ActiveWorkbook.GetNormalRowStyle();
            var excessRowStyle = app.ActiveWorkbook.GetExcessRowStyle();

            var row = 1;
            var column = 1;
            foreach (var manager in TargetItemCountByManager.Keys)
            {
                var fund = rows.Where(x => x.Manager == manager);
                var excessBelow = TargetItemCountByManager[manager];

                // Write PM initials
                sheet.Cells[row, column] = manager;
                row += 2;

                // Write column headers
                var y = 0;
                foreach (var col in headers)
                {
                    Excel.Range cell = sheet.Cells[row, column + y];
                    cell.Value = col.Name;
                    cell.ColumnWidth = col.Width;
                    cell.Style = headerStyle;
                    ++y;
                }
                row++;

                // Write longs
                var longs = fund.Where(x => x.PercentNAV > 0).OrderBy(x => (x.InstrumentClassId == (int)InstrumentClassIds.EquityIndexFuture || x.InstrumentClassId == (int)InstrumentClassIds.EquityIndexOption ? 1 : 0)).ThenByDescending(x => x.PercentNAV).ToList();
                sheet.WriteIndexColumn(row, column++, longs.Count(), excessBelow, rowStyle, excessRowStyle);
                sheet.WriteFieldColumn(row, column++, null, longs, "Ticker", excessBelow, rowStyle, excessRowStyle);
                sheet.WriteFieldColumn(row, column++, "0.0%", longs, "PercentNAV", excessBelow, rowStyle, excessRowStyle);
                sheet.WriteFieldColumn(row, column++, "#,##0", longs, "NetPosition", excessBelow, rowStyle, excessRowStyle);
                sheet.WriteWatchListColumn(row, column++, null, longs, excessBelow, rowStyle, excessRowStyle, watchList, WatchListSheet.AverageVolume);
                sheet.WriteWatchListColumn(row, column++, "0%", longs, excessBelow, rowStyle, excessRowStyle, watchList, WatchListSheet.Upside, "=IF(ISNUMBER([Address]), [Address], \"\")");
                sheet.WriteWatchListColumn(row, column++, null, longs, excessBelow, rowStyle, excessRowStyle, watchList, WatchListSheet.Conviction, "=[Address] & \"\"", Excel.XlHAlign.xlHAlignCenter);

                column += 5;

                // Write column headers
                row--;
                y = 0;
                foreach (var col in headers)
                {
                    Excel.Range cell = sheet.Cells[row, column + y];
                    cell.Value = col.Name;
                    cell.ColumnWidth = col.Width;
                    cell.Style = headerStyle;
                    ++y;
                    if (manager == "AC" && y > 2)
                    {
                        break;
                    }
                }
                row++;

                // Write shorts
                var shortQuery = fund.Where(x => x.PercentNAV < 0);
                if (manager == "AC")
                {
                    shortQuery = shortQuery.GroupBy(p => p.Strategy).Select(g => new {
                        Ticker = g.Key,
                        Manager = (string)null,
                        PercentNAV = g.Sum(p => p.PercentNAV),
                        NetPosition = (decimal?)null,
                        Strategy = (string)null,
                        InstrumentClassId = 0,
                    });
                }
                shortQuery = shortQuery.OrderBy(x => (x.InstrumentClassId == (int)InstrumentClassIds.EquityIndexFuture || x.InstrumentClassId == (int)InstrumentClassIds.EquityIndexOption ? 1 : 0)).ThenBy(x => x.PercentNAV);
                var shorts = shortQuery.ToList();
                sheet.WriteIndexColumn(row, column++, shorts.Count(), excessBelow, rowStyle, excessRowStyle);
                sheet.WriteFieldColumn(row, column++, null, shorts, "Ticker", excessBelow, rowStyle, excessRowStyle);
                sheet.WriteFieldColumn(row, column++, "0.0%", shorts, "PercentNAV", excessBelow, rowStyle, excessRowStyle);
                if (manager == "AC")
                {
                    column += 4;
                }
                else
                {
                    sheet.WriteFieldColumn(row, column++, "#,##0", shorts, "NetPosition", excessBelow, rowStyle, excessRowStyle);
                    sheet.WriteWatchListColumn(row, column++, null, shorts, excessBelow, rowStyle, excessRowStyle, watchList, WatchListSheet.AverageVolume);
                    sheet.WriteWatchListColumn(row, column++, "0%", shorts, excessBelow, rowStyle, excessRowStyle, watchList, WatchListSheet.Upside, "=IF(ISNUMBER([Address]), [Address], \"\")");
                    sheet.WriteWatchListColumn(row, column++, null, shorts, excessBelow, rowStyle, excessRowStyle, watchList, WatchListSheet.Conviction, "=[Address] & \"\"", Excel.XlHAlign.xlHAlignCenter);
                }
                column += 5;

                if (manager == "JH")
                {
                    row = 1;
                }
                else
                {
                    row += 1 + Math.Max(shorts.Count, longs.Count) + 1;
                    column -= (headers.Length + 5) * 2;
                }
            }
        }
    }
}
