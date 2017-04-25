using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using Odey.Framework.Keeley.Entities.Enums;
using Odey.PortfolioCache.Entities;
using System;

namespace Odey.ExcelAddin
{
    class ExposureItem
    {
        public string Ticker { get; set; }
        public string Manager { get; set; }
        public decimal PercentNAV { get; set; }
        public decimal NetPosition { get; set; }
        public string Strategy { get; set; }
        public List<int> InstrumentClassIds { get; set; }
    }

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
                .ToLookup(p => new { p.BloombergTicker, p.ManagerName, p.StrategyName })
                .Select(g => new ExposureItem
                {
                    Ticker = g.Key.BloombergTicker,
                    Manager = Ribbon1.GetManagerInitials(g.Key.ManagerName),
                    PercentNAV = (g.Sum(p => p.Exposure) / g.Select(p => p.FundNAV).Distinct().Single()) ?? 0,
                    NetPosition = g.Sum(p => p.NetPosition) ?? 0,
                    Strategy = g.Key.StrategyName,
                    InstrumentClassIds = g.Select(p => p.InstrumentClassId).Distinct().ToList(),
                })
                .ToList();

            // Get the worksheet
            var sheetName = $"Exposure {fundName}";
            Excel.Worksheet sheet;
            try
            {
                sheet = app.Sheets[sheetName];
            }
            catch
            {
                sheet = app.Sheets.Add();
                sheet.Name = sheetName;
            }

            sheet.Cells.ClearContents();
            sheet.Cells.ClearFormats();

            var row = 1;

            // Total Gross Exposure
            sheet.Cells[row, 1] = "Total Gross Exposure";
            Excel.Range totalGrossExposureCell = sheet.Cells[row, 3];
            totalGrossExposureCell.Value2 = rows.Sum(p => Math.Abs(p.PercentNAV));
            totalGrossExposureCell.NumberFormat = "0.0%";

            // Date
            sheet.Cells[row, 7] = weightings.Select(p => p.ReferenceDate).Distinct().Single().ToString("dd/MM/yyyy");

            ++row;

            // Total Net Exposure
            sheet.Cells[row, 1] = "Total Net Exposure";
            Excel.Range totalNetExposureCell = sheet.Cells[row, 3];
            totalNetExposureCell.Value2 = rows.Sum(p => p.PercentNAV);
            totalNetExposureCell.NumberFormat = "0.0%";
            
            ++row;
            ++row;

            var column = 1;
            foreach (var manager in TargetItemCountByManager.Keys)
            {
                var managerPositions = rows.Where(x => x.Manager == manager);
                var excessBelow = TargetItemCountByManager[manager];

                // Manager initials
                sheet.Cells[row, column] = manager;

                // Manager Gross Exposure
                sheet.Cells[row + 1, column + 1] = "Gross Exposure";
                Excel.Range managerExposureCell = sheet.Cells[row + 1, column + 2];
                managerExposureCell.Value2 = managerPositions.Sum(p => Math.Abs(p.PercentNAV));
                managerExposureCell.NumberFormat = "0.0%";

                // Percent of Total Exposure
                sheet.Cells[row, column + 1] = "Percent of Total Exposure";
                Excel.Range fundPercentageCell = sheet.Cells[row, column + 2];
                fundPercentageCell.Formula = $"={managerExposureCell.Address}/{totalGrossExposureCell.Address}";
                fundPercentageCell.NumberFormat = "0.0%";

                // Write longs
                var longs = managerPositions.Where(x => x.PercentNAV > 0).OrderBy(x => (x.InstrumentClassIds.Contains((int)InstrumentClassIds.EquityIndexFuture) || x.InstrumentClassIds.Contains((int)InstrumentClassIds.EquityIndexOption) ? 1 : 0)).ThenByDescending(x => x.PercentNAV);
                var longHeight = WriteExposureTable(sheet, row + 3, column, longs.ToList(), watchList, excessBelow, "Long", "=BDP(\"[Ticker]\",\"SHORT_NAME\")");
                column += 7 + 5;

                // Write shorts
                var shorts = managerPositions.Where(x => x.PercentNAV < 0);
                if (manager == "AC")
                {
                    shorts = shorts.GroupBy(p => p.Strategy).Select(g => new ExposureItem { Ticker = g.Key, PercentNAV = g.Sum(p => p.PercentNAV) }).OrderBy(x => x.PercentNAV);
                }
                else
                {
                    shorts = shorts.OrderBy(x => (x.InstrumentClassIds.Contains((int)InstrumentClassIds.EquityIndexFuture) || x.InstrumentClassIds.Contains((int)InstrumentClassIds.EquityIndexOption) ? 1 : 0)).ThenBy(x => x.PercentNAV);
                }
                var shortHeight = WriteExposureTable(sheet, row + 3, column, shorts.ToList(), watchList, excessBelow, "Short", (manager != "AC" ? "=BDP(\"[Ticker]\",\"SHORT_NAME\")" : null));
                column += 7 + 5;

                if (manager == "JH")
                {
                    row = 3;
                }
                else
                {
                    row += 3 + Math.Max(shortHeight, longHeight) + 2;
                    column -= (7 + 5) * 2;
                }
            }
        }

        private static int WriteExposureTable(Excel.Worksheet sheet, int row, int column, List<ExposureItem> items, Dictionary<string, WatchListItem> watchList, int excessBelow, string nameHeader = "Name", string nameFormula = null)
        {
            var headers = new[] {
                new ColumnDef { Name = "#", Width = 3.4 },
                new ColumnDef { Name = nameHeader, Width = 23 },
                new ColumnDef { Name = "% NAV", Width = 7 },
                new ColumnDef { Name = "Net Position", Width = 0 },
                new ColumnDef { Name = "Daily Volume", Width = 0 },
                new ColumnDef { Name = "Upside", Width = 7 },
                new ColumnDef { Name = "Conviction", Width = 9.7 }
            };
            var wb = sheet.Application.ActiveWorkbook;
            var headerStyle = wb.GetHeaderStyle();
            var rowStyle = wb.GetNormalRowStyle();
            var excessRowStyle = wb.GetExcessRowStyle();

            // Write headers
            var y = 0;
            Excel.Range cell = null;
            foreach (var col in headers)
            {
                cell = sheet.Cells[row, column + y];
                cell.Value = col.Name;
                cell.ColumnWidth = col.Width;
                cell.Style = headerStyle;
                ++y;
            }
            ++row;

            // Write items
            var len = Math.Max(items.Count, excessBelow + 2);
            for (y = 0; y < len; ++y)
            {
                var item = items.ElementAtOrDefault(y);
                var watchListItem = (item != null && watchList.ContainsKey(item.Ticker) ? watchList[item.Ticker] : null);

                // Rank
                cell = sheet.Cells[row + y, column + 0];
                cell.Value2 = y + 1;
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);

                // Name
                cell = sheet.Cells[row + y, column + 1];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                if (item != null)
                {
                    if (nameFormula != null)
                    {
                        cell.Formula = nameFormula.Replace("[Ticker]", item.Ticker);
                    }
                    else
                    {
                        cell.Value2 = item.Ticker;
                    }
                    
                }

                // PercentNAV
                cell = sheet.Cells[row + y, column + 2];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                if (item != null)
                {
                    cell.NumberFormat = "0.0%";
                    cell.Value2 = item.PercentNAV;
                }

                // NetPosition
                cell = sheet.Cells[row + y, column + 3];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                if (item != null)
                {
                    cell.NumberFormat = "#,##0";
                    cell.Value2 = item.NetPosition;
                }

                // AverageVolume
                cell = sheet.Cells[row + y, column + 4];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                if (watchListItem != null)
                {
                    var address = VstoExtensions.GetAddress(WatchListSheet.Name, WatchListSheet.AverageVolume.AlphabeticalIndex, watchListItem.RowIndex);
                    cell.Formula = "=" + address;
                }

                // Upside
                cell = sheet.Cells[row + y, column + 5];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                if (watchListItem != null)
                {
                    var address = VstoExtensions.GetAddress(WatchListSheet.Name, WatchListSheet.Upside.AlphabeticalIndex, watchListItem.RowIndex);
                    cell.NumberFormat = "0%";
                    cell.Formula = $"=IF(ISNUMBER({address}), {address}, \"\")";
                }

                // Conviction
                cell = sheet.Cells[row + y, column + 6];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                if (watchListItem != null)
                {
                    var address = VstoExtensions.GetAddress(WatchListSheet.Name, WatchListSheet.Conviction.AlphabeticalIndex, watchListItem.RowIndex);
                    cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    cell.Formula = $"={address} & \"\"";
                }
            }

            return 1 + len;
        }
    }
}
