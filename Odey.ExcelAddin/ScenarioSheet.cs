using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using Odey.Framework.Keeley.Entities.Enums;
using Odey.Query.Reporting.Contracts;
using System.Diagnostics;

namespace Odey.ExcelAddin
{
    class ScenarioSheet
    {
        private static int HeaderRow = 14;

        private static string[] ScenarioInputColumns = new[] { "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL" };

        public static void Write(Excel.Application app, KeyValuePair<FundIds, string> fund, List<PortfolioItem> items, Dictionary<string, WatchListItem> watchList)
        {
            app.StatusBar = $"Writing {fund.Value} scenario sheet...";

            var rows = items
                .Where(p => p.Field == PortfolioFields.Instrument && p.Ticker != null && p.FundId == fund.Key)
                .ToLookup(p => new { p.Ticker, p.ManagerInitials })
                .Select(g => new
                {
                    // These property names will be used as column names
                    g.Key.Ticker,
                    Manager = g.Key.ManagerInitials,
                    PercentNAV = g.Sum(p => p.Exposure),
                })
                .OrderBy(x => x.Ticker)
                .ToArray();

            var sheet = app.GetOrCreateVstoWorksheet($"Scenarios {fund.Value}");

            var tName = $"Scenarios_{fund.Value}";
            var table = sheet.GetListObject(tName);
            if (table == null)
            {
                // Create table
                Debug.WriteLine($"Creating table {tName}");
                table = sheet.CreateListObject(tName, HeaderRow, 1);
                table.ShowTableStyleRowStripes = false;
                table.ShowTableStyleFirstColumn = true;
                table.AutoSetDataBoundColumnHeaders = true;
                table.HeaderRowRange.WrapText = true;
                table.HeaderRowRange.RowHeight = 75;
                table.HeaderRowRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                table.HeaderRowRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            else
            {
                Debug.WriteLine($"Using existing table {tName}");
            }

            // Set table data
            table.SetDataBinding(rows);

            // Update column styles
            table.ListColumns["Ticker"].DataBodyRange.ColumnWidth = 22;
            table.ListColumns["PercentNAV"].DataBodyRange.ColumnWidth = 14;
            table.ListColumns["PercentNAV"].DataBodyRange.NumberFormat = "0.00%";

            // Disconnect data binding
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
