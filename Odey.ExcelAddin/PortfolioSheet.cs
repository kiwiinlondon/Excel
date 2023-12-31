﻿using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using Odey.Framework.Keeley.Entities.Enums;
using System;
using System.Diagnostics;

namespace Odey.ExcelAddin
{
    class PortfolioSheet
    {
        private static int HeaderRow = 14;

        private static Dictionary<string, string> ColumnLocations = new Dictionary<string, string>
            {
                { "[Ticker]", "A" },
                { "[Upside]", "G" },
                { "[PercentNAV]", "C" },
            };

        public static List<ColumnDef> Columns = new List<ColumnDef>
        {
            new ColumnDef
            {
                AlphabeticalIndex = "B",
                CopyFormula = true,
            },
            new ColumnDef
            {
                AlphabeticalIndex = "C",
                CopyFormula = true,
            },
            new ColumnDef
            {
                AlphabeticalIndex = "D",
                CopyFormula = true,
            },
            new ColumnDef
            {
                AlphabeticalIndex = "T",
                RefAsNumber = true,
            },
            new ColumnDef
            {
                AlphabeticalIndex = "F",
                RefAsString = true,
            },
            new ColumnDef
            {
                AlphabeticalIndex = "E",
                RefAsNumber = true,
            },
            new ColumnDef
            {
                // Price
                AlphabeticalIndex = "S",
                CopyFormula = true,
            },
            new ColumnDef
            {
                Name = "Upside / %NAV",
                Formula = "=IFERROR([Upside]/[PercentNAV], \"\")",
                NumberFormat = "#,##0.00",
                Width = 12.29,
            },
            new ColumnDef
            {
                AlphabeticalIndex = "U",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "NET DEBT Inc PENSIONS (Mln)",
                AlphabeticalIndex = "V",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "ENTERPRISE VALUE",
                AlphabeticalIndex = "W",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "ENTERPRISE VALUE EST",
                AlphabeticalIndex = "X",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "DIVIDEND YIELD",
                AlphabeticalIndex = "Y",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "DIVIDEND YIELD EST",
                AlphabeticalIndex = "Z",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "EBIT (Mln)",
                AlphabeticalIndex = "AA",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "EBIT EST",
                AlphabeticalIndex = "AB",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "EV/EBIT",
                AlphabeticalIndex = "AC",
                RefAsNumber = true,
            },
            new ColumnDef
            {
                //Name = "EV/EBIT EST",
                AlphabeticalIndex = "AD",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "Sales",
                AlphabeticalIndex = "AF",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "Sales EST",
                AlphabeticalIndex = "AG",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "EV/Sales",
                AlphabeticalIndex = "AH",
                RefAsNumber = true,
            },
            new ColumnDef
            {
                //Name = "EV/Sales EST",
                AlphabeticalIndex = "AI",
                RefAsNumber = true,
            },
            new ColumnDef
            {
                //Name = "TRAIL 12M EPS",
                AlphabeticalIndex = "AJ",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "EPS EST",
                AlphabeticalIndex = "AK",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "P/E Ratio",
                AlphabeticalIndex = "AL",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "P/E Ratio EST",
                AlphabeticalIndex = "AM",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "Book Value Per SH",
                AlphabeticalIndex = "AN",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "Book Value Per SH EST",
                AlphabeticalIndex = "AO",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "P/NAV",
                AlphabeticalIndex = "AP",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "P/NAV EST",
                AlphabeticalIndex = "AQ",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "Tang Book Value Per SH",
                AlphabeticalIndex = "AR",
                CopyFormula = true,
            },
            new ColumnDef
            {
                //Name = "P/TNAV",
                AlphabeticalIndex = "AS",
                CopyFormula = true,
            },
        };

        public static void Write(Excel.Application app, KeyValuePair<FundIds, string> fund, List<PortfolioItem> items, Dictionary<string, WatchListItem> watchList)
        {
            app.StatusBar = $"Writing {fund.Value} portfolio sheet...";

            var rows = items
                .Where(p => p.Ticker != null && p.FundId == fund.Key)
                .ToLookup(p => new { p.Ticker, p.ManagerId, p.ManagerInitials })
                .Select(g => new
                {
                    // These will convert into columns of the same name
                    g.Key.Ticker,
                    Manager = g.Key.ManagerInitials,
                    PercentNAV = g.Sum(p => p.Exposure),
                })
                .OrderBy(x => x.Ticker)
                .ToArray();

            app.AutoCorrect.AutoFillFormulasInLists = false;
            var sheet = app.GetOrCreateVstoWorksheet($"Portfolio {fund.Value}");

            var tName = $"Portfolio_{fund.Value}";
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

            // Sort by whatever we set this column to
            Excel.ListColumn sortyByColumn = null;

            // Add additional columns to the table (static Columns list from above)
            foreach (var column in Columns)
            {
                var col = table.ListColumns.Add();
                col.Name = column.Name;
                col.Range.ColumnWidth = column.Width;
                if (column.NumberFormat != null)
                {
                    col.DataBodyRange.NumberFormat = column.NumberFormat;
                }
                if (column.Formula != null)
                {
                    col.DataBodyRange.Formula = PrepareFormula(column.Formula, watchList.Count);
                }
                else if (column.RefAsNumber)
                {
                    var y = 1;
                    foreach (var row in rows)
                    {
                        var item = watchList[row.Ticker];
                        var address = $"'{WatchListSheet.Name}'!{column.AlphabeticalIndex}{item.RowIndex}";
                        col.DataBodyRange[y, 1].Formula = $"=IF(ISNUMBER({address}), {address}, \"\")";
                        ++y;
                    }
                }
                else if (column.RefAsString)
                {
                    var y = 1;
                    foreach (var row in rows)
                    {
                        var item = watchList[row.Ticker];
                        var address = $"'{WatchListSheet.Name}'!{column.AlphabeticalIndex}{item.RowIndex}";
                        col.DataBodyRange[y, 1].Formula = $"={address} & \"\"";
                        ++y;
                    }
                }
                else
                {
                    throw new Exception($"Invalid ColumnDef '{column.Name}'");
                }

                if (column.Name == "Upside / %NAV")
                {
                    sortyByColumn = col;
                }
            }
            
            // Sort
            //table.Sort.SortFields.Add(sortyByColumn.Range, Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlAscending);
            //table.Sort.Header = Excel.XlYesNoGuess.xlYes;
            //table.Sort.Apply();

            app.AutoCorrect.AutoFillFormulasInLists = true;
        }

        private static string PrepareFormula(string formula, int watchListCount)
        {
            var row = (HeaderRow + 1);
            foreach (var placeholder in ColumnLocations)
            {
                formula = formula.Replace(placeholder.Key, "$" + placeholder.Value + row);
            }
            return formula;
        }
    }
}
