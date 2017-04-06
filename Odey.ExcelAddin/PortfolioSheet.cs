using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using Odey.Framework.Keeley.Entities.Enums;
using Odey.PortfolioCache.Entities;
using System.Diagnostics;

namespace Odey.ExcelAddin
{
    class PortfolioSheet
    {
        private static int HeaderRow = 14;

        private static Dictionary<string, string> ColumnLocations = new Dictionary<string, string>
            {
                { "[Ticker]", "A" },
                { "[Target Price]", "I" },
                { "[Price]", "J" },
                { "[Upside]", "G" },
                { "[PercentNAV]", "C" },
                { "[Enterprise Value]", "N" },
                { "[EBIT]", "R" },
                { "[Sales]", "V" },
                { "[Sales EST]", "W" },
            };

        private static List<ColumnDef> Columns = new List<ColumnDef>
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
            new ColumnDef
            {
                Name = "UPSIDE",
                Formula = "=IFERROR(([Target Price]-[Price])/[Price], \"\")",
                Width = 12.29,
                NumberFormat = "0%",
            },
            new ColumnDef
            {
                Name = "BASIS for TARGET PRICE",
                Formula = "=VLOOKUP([Ticker], Watch_List_Table, 6, FALSE) & \"\"",
                Width = 15.43,
            },
            new ColumnDef
            {
                Name = "TARGET PRICE",
                Formula = "=VLOOKUP([Ticker], Watch_List_Table, 5, FALSE) & \"\"",
                Width = 12.29,
                NumberFormat = "#,##0.00",
            },
            new ColumnDef
            {
                Name = "PRICE",
                Formula = "=BDP([Ticker],\"PX_LAST\",\"Fill=B\")",
                Width = 12.29,
                NumberFormat = "#,##0.00",
            },
            new ColumnDef
            {
                Name = "UPSIDE/Weight(%)",
                Formula = "=IFERROR([Upside]/[PercentNAV], \"\")",
                Width = 12.29,
                NumberFormat = "0.00",
            },
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
            new ColumnDef
            {
                Name = "EV/EBIT",
                Formula = "=IFERROR([Enterprise Value]/[EBIT],\"\")",
                Width = 12.29,
                NumberFormat = "#,##0.00",
            },
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
            new ColumnDef
            {
                Name = "EV/Sales",
                Formula = "=IFERROR([Enterprise Value]/[Sales],\"\")",
                Width = 12.29,
                NumberFormat = "#,##0.00",
            },
            new ColumnDef
            {
                Name = "EV/Sales EST",
                Formula = "=IFERROR([Enterprise Value]/[Sales EST],\"\")",
                Width = 12.29,
                NumberFormat = "#,##0.00",
            },
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
            new ColumnDef
            {
                Name = "EV/EBITDA",
                //Formula = "=IFERROR([Enterprise Value]/[],\"\")",
                Width = 9.43,
                NumberFormat = "#,##0.00",
            },
            new ColumnDef
            {
                Name = "60-day beta (MSCI world TR relevant currency for fund)",
                //Formula = "=BDP([Ticker],\"BETA_ADJ_OVERRIDABLE\",\"BETA_OVERRIDE_REL_INDEX=gdduwi index\",\"BETA_OVERRIDE_START_DT\",TEXT('Watch List'!BO6,\"YYYYMMDD\"),\"BETA_OVERRIDE_PERIOD=d\",\"Fill=B\")",
                Width = 10.71,
                NumberFormat = "#,##0",
            },
        };

        public static void Write(Excel.Application app, FundIds fundId, List<PortfolioDTO> weightings, Dictionary<string, WatchListItem> watchList)
        {
            var fundName = Ribbon1.GetFundName(fundId, weightings);
            app.StatusBar = $"Writing {fundName} portfolio sheet...";

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

            var sheet = app.GetOrCreateVstoWorksheet($"Portfolio {fundName}");

            var tName = $"Portfolio_{fundName}";
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

            // Add additional columns to the table
            var allColumns = new Dictionary<string, Excel.ListColumn>();
            foreach (var column in Columns)
            {
                var col = table.ListColumns.Add();
                col.Name = column.Name;
                col.Range.ColumnWidth = column.Width;
                if (column.NumberFormat != null)
                {
                    col.Range.NumberFormat = column.NumberFormat;
                }
                allColumns.Add(col.Name, col);
            }

            // Write formulas when the table is done
            foreach (var column in Columns)
            {
                if (column.Formula != null)
                {
                    var col = allColumns[column.Name];
                    col.DataBodyRange.Formula = GetFormula(column.Formula, watchList.Count);
                    Debug.WriteLine(GetFormula(column.Formula, watchList.Count));
                }
            }
        }

        private static string GetFormula(string formula, int watchListCount)
        {
            var row = (HeaderRow + 1).ToString();
            foreach (var placeholder in ColumnLocations)
            {
                formula = formula.Replace(placeholder.Key, "$" + placeholder.Value + row);
            }
            return formula.Replace("Watch_List_Table", $"'{WatchListSheet.Name}'!$A${WatchListSheet.HeaderRow + 1}:$BO${WatchListSheet.HeaderRow + watchListCount}");
        }
    }
}
