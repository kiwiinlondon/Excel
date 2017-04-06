using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using Odey.Framework.Keeley.Entities.Enums;
using Odey.PortfolioCache.Entities;

namespace Odey.ExcelAddin
{
    class ScenarioSheet
    {
        private static int HeaderRow = 14;

        private static string[] ScenarioInputColumns = new[] { "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL" };

        public static void Write(Excel.Application app, FundIds fundId, List<PortfolioDTO> weightings, Dictionary<string, WatchListItem> watchList)
        {
            var fundName = Ribbon1.GetFundName(fundId, weightings);
            app.StatusBar = $"Writing {fundName} scenario sheet...";

            var rows = weightings
                .Where(p => p.ExposureTypeId == ExposureTypeIds.Primary && p.BloombergTicker != null && p.FundId == (int)fundId)
                .ToLookup(p => new { p.EquivalentInstrumentMarketId, p.BloombergTicker, p.ManagerName })
                .Select(g => new
                {
                    Ticker = g.Key.BloombergTicker,
                    Manager = Ribbon1.GetManagerInitials(g.Key.ManagerName),
                    PercentNAV = g.Sum(p => p.Exposure) / g.Select(p => p.FundNAV).Distinct().Single(),
                })
                .ToList();

            var sheet = app.GetOrCreateVstoWorksheet($"Scenarios {fundName}");

            var tName = $"Scenarios_{fundName}";
            var table = sheet.GetListObject(tName);
            if (table == null)
            {
                table = sheet.CreateListObject(tName, HeaderRow, 1);
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

            app.AutoCorrect.AutoFillFormulasInLists = false;
            var headerColumn = 4;
            foreach (var columnLetter in ScenarioInputColumns)
            {
                Excel.Range topHeaderCell = sheet.Cells[HeaderRow - 1, headerColumn];
                topHeaderCell.Formula = $"='{WatchListSheet.Name}'!{columnLetter}{WatchListSheet.HeaderRow}";
                topHeaderCell.Resize[1, 2].Merge();
                topHeaderCell.RowHeight = 75;
                headerColumn += 2;

                var col = table.ListColumns.Add();
                col.Name = $"{columnLetter} Factor";
                Excel.Range r = col.DataBodyRange;

                var y = 1;
                foreach (var row in rows)
                {
                    Excel.Range cell = r.Rows[y];
                    var wlItem = watchList[row.Ticker];
                    cell.Formula = $"='{WatchListSheet.Name}'!{columnLetter}{wlItem.RowIndex}";
                    ++y;
                }

                var col2 = table.ListColumns.Add();
                col2.Name = $"{columnLetter} Result";
                col2.DataBodyRange.Formula = $"=[{col.Name}]*[PercentNAV]";
            }
            app.AutoCorrect.AutoFillFormulasInLists = true;
        }

    }
}
