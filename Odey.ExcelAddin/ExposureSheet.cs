using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using Odey.Framework.Keeley.Entities.Enums;
using Odey.PortfolioCache.Entities;

namespace Odey.ExcelAddin
{
    class ExposureSheet
    {
        private static Dictionary<string, int> ItemsPerManager = new Dictionary<string, int>
        {
            { "JH", 34 },
            { "AC", 10 },
            { "JG", 12 },
        };

        public static void Write(Excel.Application app, FundIds fundId, List<PortfolioDTO> weightings, Dictionary<string, WatchListItem> watchList)
        {
            var fundName = Ribbon1.GetFundName(fundId, weightings);
            app.StatusBar = $"Writing {fundName} exposure sheet...";

            var rows = weightings
                .Where(p => p.ExposureTypeId == ExposureTypeIds.Primary && p.BloombergTicker != null && p.FundId == (int)fundId)
                .ToLookup(p => new { p.BloombergTicker, p.ManagerName, p.StrategyName })
                .Select(g => new
                {
                    Ticker = g.Key.BloombergTicker,
                    Manager = Ribbon1.GetManagerInitials(g.Key.ManagerName),
                    PercentNAV = g.Sum(p => p.Exposure) / g.Select(p => p.FundNAV).Distinct().Single(),
                    NetPosition = g.Sum(p => p.NetPosition),
                    Strategy = g.Key.StrategyName
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
                new ColumnDef { Name = "% NAV", Width = 7, NumberFormat = "0.00%" },
                new ColumnDef { Name = "Net Position", Width = 0, NumberFormat = "#,###" },
                new ColumnDef { Name = "Daily Volume", Width = 0 },
                new ColumnDef { Name = "Upside", Width = 7 },
                new ColumnDef { Name = "Conviction", Width = 9.7 }
            };

            var row = 1;
            var column = 1;
            foreach (var manager in ItemsPerManager.Keys)
            {
                var fund = rows.Where(x => x.Manager == manager);
                var numItems = ItemsPerManager[manager];

                // Write PM initials
                if (isNewSheet)
                {
                    sheet.Cells[row, column] = manager;
                }
                row += 2;

                // Write column headers
                if (isNewSheet)
                {
                    sheet.WriteColumnHeader(row, column + 0, headers[0]);
                    sheet.WriteColumnHeader(row, column + 1, headers[1]);
                    sheet.WriteColumnHeader(row, column + 2, headers[2]);
                    sheet.WriteColumnHeader(row, column + 3, headers[3]);
                    sheet.WriteColumnHeader(row, column + 4, headers[4]);
                    sheet.WriteColumnHeader(row, column + 5, headers[5]);
                    sheet.WriteColumnHeader(row, column + 6, headers[6]);
                }
                row += 1;

                // Clear longs
                Excel.Range range = sheet.Range[sheet.Cells[row, column], sheet.Cells[row + numItems - 1, column + headers.Length - 1]];
                range.ClearContents();

                // Write longs
                var longs = fund.Where(x => x.PercentNAV > 0).OrderByDescending(x => x.PercentNAV).Take(numItems).ToList();
                sheet.WriteIndexColumn(row, column++, headers[0], longs.Count());
                sheet.WriteFieldColumn(row, column++, headers[1], longs, "Ticker");
                sheet.WriteFieldColumn(row, column++, headers[2], longs, "PercentNAV");
                sheet.WriteFieldColumn(row, column++, headers[3], longs, "NetPosition");
                sheet.WriteWatchListColumn(row, column++, headers[4], longs, watchList, WatchListSheet.AverageVolume);
                sheet.WriteWatchListColumn(row, column++, headers[5], longs, watchList, WatchListSheet.Upside);
                sheet.WriteWatchListColumn(row, column++, headers[6], longs, watchList, WatchListSheet.Conviction);
                
                column += 5;

                // Write column headers
                if (isNewSheet)
                {
                    row--;
                    sheet.WriteColumnHeader(row, column + 0, headers[0]);
                    sheet.WriteColumnHeader(row, column + 1, headers[1]);
                    sheet.WriteColumnHeader(row, column + 2, headers[2]);
                    sheet.WriteColumnHeader(row, column + 3, headers[3]);
                    sheet.WriteColumnHeader(row, column + 4, headers[4]);
                    sheet.WriteColumnHeader(row, column + 5, headers[5]);
                    sheet.WriteColumnHeader(row, column + 6, headers[6]);
                    row++;
                }

                // Clear shorts
                range = sheet.Range[sheet.Cells[row, column], sheet.Cells[row + numItems - 1, column + headers.Length - 1]];
                range.ClearContents();

                // Write shorts
                var shortQuery = fund.Where(x => x.PercentNAV < 0);
                if (manager == "AC")
                {
                    shortQuery = shortQuery.GroupBy(p => p.Strategy).Select(g => new {
                        Ticker = g.Key,
                        Manager = (string)null,
                        PercentNAV = g.Sum(p => p.PercentNAV),
                        NetPosition = (decimal?)null,
                        Strategy = (string)null
                    });
                }
                shortQuery = shortQuery.OrderBy(x => x.PercentNAV).Take(numItems);
                var shorts = shortQuery.ToList();
                sheet.WriteIndexColumn(row, column++, headers[0], shorts.Count());
                sheet.WriteFieldColumn(row, column++, headers[1], shorts, "Ticker");
                sheet.WriteFieldColumn(row, column++, headers[2], shorts, "PercentNAV");
                if (manager == "AC")
                {
                    column += 4;
                }
                else
                {
                    sheet.WriteFieldColumn(row, column++, headers[3], shorts, "NetPosition");
                    sheet.WriteWatchListColumn(row, column++, headers[4], shorts, watchList, WatchListSheet.AverageVolume);
                    sheet.WriteWatchListColumn(row, column++, headers[5], shorts, watchList, WatchListSheet.Upside);
                    sheet.WriteWatchListColumn(row, column++, headers[6], shorts, watchList, WatchListSheet.Conviction);
                }
                column += 5;

                if (manager == "JH")
                {
                    row = 1;
                }
                else
                {
                    row += 1 + numItems + 2;
                    column -= (headers.Length + 5) * 2;
                }
            }
        }
    }
}
