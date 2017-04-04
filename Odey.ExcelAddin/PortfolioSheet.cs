using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using Odey.Framework.Keeley.Entities.Enums;
using Odey.PortfolioCache.Entities;

namespace Odey.ExcelAddin
{
    class PortfolioSheet
    {
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

        public static void Write(Excel.Application app, FundIds fundId, List<PortfolioDTO> weightings, Dictionary<string, WatchListItem> watchList)
        {
            app.StatusBar = $"Writing {fundId} portfolio sheet...";

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

            foreach (var column in Columns)
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

    }
}
