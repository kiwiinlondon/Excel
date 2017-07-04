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
        public List<string> Tickers { get; set; }
        public string Manager { get; set; }
        public decimal PercentNAV { get; set; }
        public decimal NetPosition { get; set; }
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
                .ToLookup(p => new { p.IssuerId, p.ManagerName })
                .Select(g => new ExposureItem
                {
                    Tickers = g.Select(p => p.BloombergTicker).Distinct().ToList(),
                    Manager = Ribbon1.GetManagerInitials(g.Key.ManagerName),
                    PercentNAV = (g.Sum(p => p.Exposure) / g.Select(p => p.FundNAV).Distinct().Single()) ?? 0,
                    NetPosition = g.Sum(p => p.NetPosition) ?? 0,
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

                // Manager Net Exposure
                sheet.Cells[row + 2, column + 1] = "Net Exposure";
                Excel.Range managerNetExposureCell = sheet.Cells[row + 2, column + 2];
                managerNetExposureCell.Value2 = managerPositions.Sum(p => p.PercentNAV);
                managerNetExposureCell.NumberFormat = "0.0%";

                // Percent of Total Exposure
                sheet.Cells[row, column + 1] = "Percent of Total Exposure";
                Excel.Range fundPercentageCell = sheet.Cells[row, column + 2];
                fundPercentageCell.Formula = $"={managerExposureCell.Address}/{totalGrossExposureCell.Address}";
                fundPercentageCell.NumberFormat = "0.0%";

                // Write longs
                var longs = managerPositions.Where(x => x.PercentNAV > 0).OrderBy(x => (x.InstrumentClassIds.Contains((int)InstrumentClassIds.EquityIndexFuture) || x.InstrumentClassIds.Contains((int)InstrumentClassIds.EquityIndexOption) ? 1 : 0)).ThenByDescending(x => x.PercentNAV);
                var longHeight = WriteExposureTable(sheet, row + 4, column, longs.ToList(), watchList, excessBelow, "Long", "=BDP(\"[Ticker]\",\"SHORT_NAME\")");
                column += headers.Length + 5;

                // Write shorts
                var shorts = managerPositions.Where(x => x.PercentNAV < 0).OrderBy(x => (x.InstrumentClassIds.Contains((int)InstrumentClassIds.EquityIndexFuture) || x.InstrumentClassIds.Contains((int)InstrumentClassIds.EquityIndexOption) ? 1 : 0)).ThenBy(x => x.PercentNAV);
                var shortHeight = WriteExposureTable(sheet, row + 4, column, shorts.ToList(), watchList, excessBelow, "Short", "=BDP(\"[Ticker]\",\"SHORT_NAME\")");
                column += headers.Length + 5;

                if (manager == "JH")
                {
                    row = 4;
                }
                else
                {
                    row += 4 + Math.Max(shortHeight, longHeight) + 2;
                    column -= (headers.Length + 5) * 2;
                }
            }
        }

        private static ColumnDef[] headers = new[] {
            new ColumnDef { Name = "#", Width = 3.4 },
            new ColumnDef { Name = "Name", Width = 23 },
            new ColumnDef { Name = "% NAV", Width = 7 },
            //new ColumnDef { Name = "Merged From", Width = 0 },
            new ColumnDef { Name = "Net Position", Width = 0 },
            new ColumnDef { Name = "Daily Volume", Width = 0 },
            new ColumnDef { Name = "Upside", Width = 7 },
            new ColumnDef { Name = "Conviction", Width = 9.7 },
            new ColumnDef { Name = "% Annual Volume", Width = 15 },
        };

        private static int WriteExposureTable(Excel.Worksheet sheet, int row, int column, List<ExposureItem> items, Dictionary<string, WatchListItem> watchList, int excessBelow, string nameHeader = "Name", string nameFormula = null)
        {

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
                cell.Value = (col == headers[1] ? nameHeader : col.Name);
                cell.ColumnWidth = col.Width;
                cell.Style = headerStyle;
                ++y;
            }
            ++row;

            // Write items
            var len = Math.Max(items.Count, excessBelow + 2);
            for (y = 0; y < len; ++y)
            {
                var x = 0;
                var item = items.ElementAtOrDefault(y);
                var ticker = item?.Tickers.FirstOrDefault();
                var watchListItem = (item != null && watchList.ContainsKey(ticker) ? watchList[ticker] : null);

                // Rank
                cell = sheet.Cells[row + y, column + x];
                cell.Value2 = y + 1;
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                ++x;

                // Name
                cell = sheet.Cells[row + y, column + x];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                ++x;
                if (item != null)
                {
                    if (nameFormula != null)
                    {
                        cell.Formula = nameFormula.Replace("[Ticker]", ticker);
                    }
                    else
                    {
                        cell.Value2 = ticker;
                    }
                }

                // PercentNAV
                cell = sheet.Cells[row + y, column + x];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                ++x;
                if (item != null)
                {
                    cell.NumberFormat = "0.0%";
                    cell.Value2 = item.PercentNAV;
                }

                //// Merged From
                //cell = sheet.Cells[row + y, column + x];
                //cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                //++x;
                //if (item != null && item.Tickers.Count > 1)
                //{
                //    cell.Value2 = string.Join(", ", item.Tickers);
                //}

                // NetPosition
                cell = sheet.Cells[row + y, column + x];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                ++x;
                if (item != null)
                {
                    cell.NumberFormat = "#,##0";
                    cell.Value2 = item.NetPosition;
                }

                // AverageVolume
                cell = sheet.Cells[row + y, column + x];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                ++x;
                if (watchListItem != null)
                {
                    var address = VstoExtensions.GetAddress(WatchListSheet.Name, WatchListSheet.AverageVolume.AlphabeticalIndex, watchListItem.RowIndex);
                    cell.Formula = "=" + address;
                }

                // Upside
                cell = sheet.Cells[row + y, column + x];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                ++x;
                if (watchListItem != null)
                {
                    var address = VstoExtensions.GetAddress(WatchListSheet.Name, WatchListSheet.Upside.AlphabeticalIndex, watchListItem.RowIndex);
                    cell.NumberFormat = "0%";
                    cell.Formula = $"=IF(ISNUMBER({address}), {address}, \"\")";
                }

                // Conviction
                cell = sheet.Cells[row + y, column + x];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                ++x;
                if (watchListItem != null)
                {
                    var address = VstoExtensions.GetAddress(WatchListSheet.Name, WatchListSheet.Conviction.AlphabeticalIndex, watchListItem.RowIndex);
                    cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    cell.Formula = $"={address} & \"\"";
                }

                // % Annual Volume
                cell = sheet.Cells[row + y, column + x];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                ++x;
                if (watchListItem != null)
                {
                    var address = VstoExtensions.GetAddress(WatchListSheet.Name, WatchListSheet.AverageVolume.AlphabeticalIndex, watchListItem.RowIndex);
                    cell.NumberFormat = "0.0%";
                    cell.Formula = $"={Math.Abs(item.NetPosition)}/{address}/250";
                }
            }

            return 1 + len;
        }
    }
}
