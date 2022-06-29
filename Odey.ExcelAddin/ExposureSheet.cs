using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using Odey.Framework.Keeley.Entities.Enums;
using System;

namespace Odey.ExcelAddin
{
    class ExposureItem
    {
        public List<string> Tickers { get; set; }
        public string ManagerInitials { get; set; }
        public ApplicationUserIds ManagerId { get; set; }
        public decimal PercentNAV { get; set; }
        public decimal NetPosition { get; set; }
        public List<InstrumentClassIds> InstrumentClassIds { get; set; }
        public bool IsShort { get; set; }
    }

    class ExposureSheet
    {
        private static Dictionary<ApplicationUserIds, int> TargetItemCountByManager = new Dictionary<ApplicationUserIds, int>
        {
            { ApplicationUserIds.JamesHanbury, 25 },
            { ApplicationUserIds.AdrianCourtenay, 8 },
            { ApplicationUserIds.JamieGrimston, 10 },
        };

        public static void WriteCombined(Excel.Application app, DateTime date, KeyValuePair<FundIds, string> fund, IEnumerable<PortfolioItem> items, Dictionary<string, WatchListItem> watchList)
        {
            app.StatusBar = $"Writing {fund.Value} exposure sheet...";

            var rows = items
                .Where(p => p.Ticker != null)
                .ToLookup(p => new { p.IssuerId, p.IsShort })
                .Select(g => new ExposureItem
                {
                    IsShort = g.Key.IsShort,
                    PercentNAV = g.Sum(p => p.Exposure),
                    NetPosition = g.Sum(p => p.NetPosition),
                    Tickers = g.Select(p => p.Ticker).Distinct().ToList(),
                    InstrumentClassIds = g.Select(p => p.InstrumentClassId).Distinct().ToList(),
                })
                .ToList();

            // Get the worksheet
            var sheetName = $"Exposure {fund.Value}";
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
            sheet.Cells[row, 1] = $"Gross {fund.Value} Exposure";
            Excel.Range totalGrossExposureCell = sheet.Cells[row, 3];
            totalGrossExposureCell.Value2 = rows.Sum(p => Math.Abs(p.PercentNAV));
            totalGrossExposureCell.NumberFormat = "0.0%";

            // Date
            sheet.Cells[row, 7] = date.ToString("dd/MM/yyyy");

            ++row;

            // Total Net Exposure
            sheet.Cells[row, 1] = $"Net {fund.Value} Exposure";
            Excel.Range totalNetExposureCell = sheet.Cells[row, 3];
            totalNetExposureCell.Value2 = rows.Sum(p => p.PercentNAV);
            totalNetExposureCell.NumberFormat = "0.0%";

            ++row;
            ++row;

            //var nameFormula = (Ribbon1.IsDebug ? "[Ticker]" : "=BDP(\"[Ticker]\",\"SHORT_NAME\")");
            var nameFormula = "=BDP(\"[Ticker]\",\"SHORT_NAME\")";
            var column = 1;
            var excessBelow = 25;

            // Write longs (They want index options to show at the end)
            var longs = rows.Where(x => !x.IsShort).OrderBy(x => (x.InstrumentClassIds.Contains(InstrumentClassIds.EquityIndexFuture) || x.InstrumentClassIds.Contains(InstrumentClassIds.EquityIndexOption) ? 1 : 0)).ThenByDescending(x => x.PercentNAV);
            var longHeight = WriteExposureTable(sheet, row, column, longs.ToList(), watchList, excessBelow, "Long", nameFormula);
            column += headers.Length + 5;

            // Write shorts (They want index options to show at the end)
            var shorts = rows.Where(x => x.IsShort).OrderBy(x => (x.InstrumentClassIds.Contains(InstrumentClassIds.EquityIndexFuture) || x.InstrumentClassIds.Contains(InstrumentClassIds.EquityIndexOption) ? 1 : 0)).ThenBy(x => x.PercentNAV);
            var shortHeight = WriteExposureTable(sheet, row, column, shorts.ToList(), watchList, excessBelow, "Short", nameFormula);
            column += headers.Length + 5;

            AddConditionalFormatting(sheet, "J");
            AddConditionalFormatting(sheet, "K");
            AddConditionalFormatting(sheet, "L");
            AddConditionalFormatting(sheet, "AA");
            AddConditionalFormatting(sheet, "AB");
            AddConditionalFormatting(sheet, "AC");

        }

        private static void AddConditionalFormatting(Excel.Worksheet worksheet, string column)
        {
            Excel.ColorScale cfColorScale = (Excel.ColorScale)(worksheet.get_Range($"{column}:{column}", $"{column}:{column}").FormatConditions.AddColorScale(3));
            cfColorScale.ColorScaleCriteria[1].Type = Excel.XlConditionValueTypes.xlConditionValueLowestValue;
            cfColorScale.ColorScaleCriteria[1].FormatColor.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(248, 105, 107)); //System.Drawing..ColorTranslator.FromHtml("#F8696B");  // Red

            cfColorScale.ColorScaleCriteria[2].Type = Excel.XlConditionValueTypes.xlConditionValuePercentile;
            cfColorScale.ColorScaleCriteria[2].Value = 50;
            cfColorScale.ColorScaleCriteria[2].FormatColor.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 235, 132)); //System.Drawing.ColorTranslator.FromHtml("#FFEB84");  // yellow
            cfColorScale.ColorScaleCriteria[3].Type = Excel.XlConditionValueTypes.xlConditionValueHighestValue;
            cfColorScale.ColorScaleCriteria[3].FormatColor.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(99, 190, 123));// System.Drawing.ColorTranslator.FromHtml("#63BE7B");  // green 
        }

        public static void Write(Excel.Application app, DateTime date, KeyValuePair<FundIds, string> fund, IEnumerable<PortfolioItem> items, Dictionary<string, WatchListItem> watchList)
        {
            app.StatusBar = $"Writing {fund.Value} exposure sheet...";

            var rows = items
                .Where(p => p.Ticker != null)
                .ToLookup(p => new { p.IssuerId, p.ManagerId, p.ManagerInitials, p.IsShort })
                .Select(g => new ExposureItem
                {
                    ManagerInitials = g.Key.ManagerInitials,
                    ManagerId = g.Key.ManagerId,
                    IsShort = g.Key.IsShort,
                    PercentNAV = g.Sum(p => p.Exposure),
                    NetPosition = g.Sum(p => p.NetPosition),
                    Tickers = g.Select(p => p.Ticker).Distinct().ToList(),
                    InstrumentClassIds = g.Select(p => p.InstrumentClassId).Distinct().ToList(),
                })
                .ToList();

            // Get the worksheet
            var sheetName = $"Exposure {fund.Value}";
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
            sheet.Cells[row, 1] = $"Total Gross {fund.Value} Exposure";
            Excel.Range totalGrossExposureCell = sheet.Cells[row, 3];
            totalGrossExposureCell.Value2 = rows.Sum(p => Math.Abs(p.PercentNAV));
            totalGrossExposureCell.NumberFormat = "0.0%";

            // Date
            sheet.Cells[row, 7] = date.ToString("dd/MM/yyyy");

            ++row;

            // Total Net Exposure
            sheet.Cells[row, 1] = "Total Net Exposure";
            Excel.Range totalNetExposureCell = sheet.Cells[row, 3];
            totalNetExposureCell.Value2 = rows.Sum(p => p.PercentNAV);
            totalNetExposureCell.NumberFormat = "0.0%";
            
            ++row;
            ++row;

            var nameFormula = (Ribbon1.IsDebug ? "[Ticker]" : "=BDP(\"[Ticker]\",\"SHORT_NAME\")");
            var column = 1;
            foreach (var manager in TargetItemCountByManager.Keys)
            {
                var managerPositions = rows.Where(x => x.ManagerId == manager);
                var excessBelow = TargetItemCountByManager[manager];

                // Manager initials
                sheet.Cells[row, column] = Ribbon1.ManagerInitials[manager];

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

                // Write longs (They want index options to show at the end)
                var longs = managerPositions.Where(x => !x.IsShort).OrderBy(x => (x.InstrumentClassIds.Contains(InstrumentClassIds.EquityIndexFuture) || x.InstrumentClassIds.Contains(InstrumentClassIds.EquityIndexOption) ? 1 : 0)).ThenByDescending(x => x.PercentNAV);
                var longHeight = WriteExposureTable(sheet, row + 4, column, longs.ToList(), watchList, excessBelow, "Long", nameFormula);
                column += headers.Length + 5;

                // Write shorts (They want index options to show at the end)
                var shorts = managerPositions.Where(x => x.IsShort).OrderBy(x => (x.InstrumentClassIds.Contains(InstrumentClassIds.EquityIndexFuture) || x.InstrumentClassIds.Contains(InstrumentClassIds.EquityIndexOption) ? 1 : 0)).ThenBy(x => x.PercentNAV);
                var shortHeight = WriteExposureTable(sheet, row + 4, column, shorts.ToList(), watchList, excessBelow, "Short", nameFormula);
                column += headers.Length + 5;

                if (manager == ApplicationUserIds.JamesHanbury)
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
            new ColumnDef { Name = "Target Price", Width = 7 },
            new ColumnDef { Name = "Basis For Target Price", Width = 130 },
            new ColumnDef { Name = "Upside", Width = 7 },
            new ColumnDef { Name = "Conviction", Width = 9.7 },
            new ColumnDef { Name = "% Annual Volume", Width = 15 },
            new ColumnDef { Name = "High (H) or Low (L) Liquidity", Width = 20 },
            new ColumnDef { Name = "Total Score", Width = 12 },
            new ColumnDef { Name = "Valuation+ Liquidity+ Brook Edge", Width = 12 },
            new ColumnDef { Name = "Company Score (Quality, Risk, ESG, other)", Width = 15 },
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
                    var address = VstoExtensions.GetAddress(WatchListSheet.Name, WatchListSheet.TargetPrice.AlphabeticalIndex, watchListItem.RowIndex);
                    cell.Formula = "=" + address;
                    cell.NumberFormat = "#,##0.00";                    
                }

                // AverageVolume
                cell = sheet.Cells[row + y, column + x];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                ++x;
                if (watchListItem != null)
                {
                    var address = VstoExtensions.GetAddress(WatchListSheet.Name, WatchListSheet.BasisForTargetPrice.AlphabeticalIndex, watchListItem.RowIndex);
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

                //High or LOw LIquidity
                cell = sheet.Cells[row + y, column + x];
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                ++x;
                if (watchListItem != null)
                {
                    var address = VstoExtensions.GetAddress(WatchListSheet.Name, WatchListSheet.LiquidityHL.AlphabeticalIndex, watchListItem.RowIndex);
                    //cell.NumberFormat = "0.0%";
                    cell.Formula = $"={address}";
                }

                //Total Score
                cell = sheet.Cells[row + y, column + x];
           //     cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                ++x;
                if (watchListItem != null)
                {
                    cell.NumberFormat = "0.00";
                    cell.Formula = $"= IF(ISNUMBER(K{row+y} + L{row + y}), (K{row + y} + L{row + y}),\"\")";
                }
                //Valuation Liquidity 
                cell = sheet.Cells[row + y, column + x];
               // cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                ++x;
                var sheetName = "Summary";
                var lookUpColumn = "B";
                if (nameHeader == "Short")
                {
                    sheetName = "ShortSummary";
                    lookUpColumn = "S";
                } 
                if (watchListItem != null)
                {
                    cell.NumberFormat = "0.00";                    
                    cell.Formula = $"=XLOOKUP({lookUpColumn}{row +y},{sheetName}!$C:$C,{sheetName}!$P:$P)";
                }
                //Company Score
                cell = sheet.Cells[row + y, column + x];
               // cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                ++x;
                if (watchListItem != null)
                {
                    cell.NumberFormat = "0.00";
                    cell.Formula = $"=XLOOKUP({lookUpColumn}{row + y},{sheetName}!$C:$C,{sheetName}!$Q:$Q)";
                }
            }

            return 1 + len;
        }
    }
}