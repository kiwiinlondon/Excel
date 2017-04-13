using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Diagnostics;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace Odey.ExcelAddin
{
    public static class VstoExtensions
    {
        public static Worksheet GetOrCreateVstoWorksheet(this Excel.Application app, string sheetName)
        {
            Worksheet sheet;
            try
            {
                sheet = Globals.Factory.GetVstoObject(app.Sheets[sheetName]);
            }
            catch
            {
                sheet = Globals.Factory.GetVstoObject(app.Sheets.Add());
                sheet.Name = sheetName;
            }
            return sheet;
        }

        public static ListObject GetListObject(this Worksheet sheet, string name)
        {
            ListObject lov = null;
            foreach (Excel.ListObject lo in sheet.ListObjects)
            {
                if (lo.Name == name)
                {
                    lov = Globals.Factory.GetVstoObject(lo);
                }
            }
            return lov;
        }

        public static ListObject CreateListObject(this Worksheet sheet, string name, int row = 1, int column = 1)
        {
            Excel.Range position = sheet.Cells[row, column];
            try
            {
                return sheet.Controls.AddListObject(position, name);
            }
            catch (Exception e)
            {
                throw new Exception($"Could not create table '{name}' at {position.AddressLocal[false, false]} in sheet '{sheet.Name}'", e);
            }
        }

        public static void WriteIndexColumn(this Excel.Worksheet sheet, int row, int column, int num, int excessBelow, Excel.Style rowStyle, Excel.Style excessRowStyle)
        {
            for (var y = 0; y < num; ++y)
            {
                Excel.Range cell = sheet.Cells[row + y, column];
                cell.Value2 = y + 1;
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
            }
        }

        public static void WriteFieldColumn<T>(this Excel.Worksheet sheet, int row, int column, string numberFormat, IEnumerable<T> data, string field, int excessBelow, Excel.Style rowStyle, Excel.Style excessRowStyle, string formula = null)
        {
            var y = 0;
            var pi = typeof(T).GetProperty(field);
            foreach (var item in data)
            {
                Excel.Range cell = sheet.Cells[row + y, column];
                if (formula != null)
                {
                    cell.Formula = formula.Replace("[StringValue]", (string)pi.GetValue(item));
                }
                else
                {
                    cell.Value2 = pi.GetValue(item);
                }
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                if (numberFormat != null)
                {
                    cell.NumberFormat = numberFormat;
                }
                ++y;
            }
        }

        public static void WriteWatchListColumn(this Excel.Worksheet sheet, int row, int column, string numberFormat, IEnumerable<dynamic> data, int excessBelow, Excel.Style rowStyle, Excel.Style excessRowStyle, Dictionary<string, WatchListItem> watchList, ColumnDef sourceColumn, string formula = "=[Address]", Excel.XlHAlign align = Excel.XlHAlign.xlHAlignGeneral)
        {
            var y = 0;
            foreach (var item in data)
            {
                var address = GetAddress(item.Ticker, sourceColumn.AlphabeticalIndex, watchList);
                Excel.Range cell = sheet.Cells[row + y, column];
                cell.Formula = formula.Replace("[Address]", address);
                cell.Style = (y < excessBelow ? rowStyle : excessRowStyle);
                cell.HorizontalAlignment = align;
                if (numberFormat != null)
                {
                    cell.NumberFormat = numberFormat;
                }
                ++y;
            }
        }

        public static string GetAddress(string ticker, string columnLetter, Dictionary<string, WatchListItem> watchList)
        {
            return $"'{WatchListSheet.Name}'!{columnLetter}{watchList[ticker].RowIndex}";
        }

        public static Excel.Style GetHeaderStyle(this Excel.Workbook wb)
        {
            foreach (Excel.Style style in wb.Styles)
            {
                if (style.Name == "Header")
                {
                    return style;
                }
            }

            var headerStyle = wb.Styles.Add("Header");
            headerStyle.WrapText = true;
            headerStyle.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerStyle.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerStyle.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;
            headerStyle.Font.Bold = true;
            headerStyle.Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;

            foreach (var index in new[] { Excel.Constants.xlTop, Excel.Constants.xlLeft, Excel.Constants.xlBottom, Excel.Constants.xlRight })
            {
                var border = headerStyle.Borders[(Excel.XlBordersIndex)index];
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = Excel.XlBorderWeight.xlThin;
                border.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
            }

            headerStyle.IncludeAlignment = true;
            headerStyle.IncludeBorder = true;
            headerStyle.IncludeFont = true;
            headerStyle.IncludeNumber = false;
            headerStyle.IncludePatterns = true;
            headerStyle.IncludeProtection = false;

            return headerStyle;
        }

        public static Excel.Style GetNormalRowStyle(this Excel.Workbook wb)
        {
            foreach (Excel.Style style in wb.Styles)
            {
                if (style.Name == "Normal Row")
                {
                    return style;
                }
            }

            var rowStyle = wb.Styles.Add("Normal Row");
            rowStyle.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0xDDD9C4));

            foreach (var index in new[] { Excel.Constants.xlTop, Excel.Constants.xlLeft, Excel.Constants.xlBottom, Excel.Constants.xlRight })
            {
                var border = rowStyle.Borders[(Excel.XlBordersIndex)index];
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = Excel.XlBorderWeight.xlThin;
                border.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
            }

            rowStyle.IncludeAlignment = false;
            rowStyle.IncludeBorder = true;
            rowStyle.IncludeFont = false;
            rowStyle.IncludeNumber = false;
            rowStyle.IncludePatterns = true;
            rowStyle.IncludeProtection = false;

            return rowStyle;
        }

        public static Excel.Style GetExcessRowStyle(this Excel.Workbook wb)
        {
            foreach (Excel.Style style in wb.Styles)
            {
                if (style.Name == "Excess Row")
                {
                    return style;
                }
            }

            var rowStyle = wb.Styles.Add("Excess Row");
            rowStyle.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0xFCAC80));

            foreach (var index in new[] { Excel.Constants.xlTop, Excel.Constants.xlLeft, Excel.Constants.xlBottom, Excel.Constants.xlRight })
            {
                var border = rowStyle.Borders[(Excel.XlBordersIndex)index];
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = Excel.XlBorderWeight.xlThin;
                border.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
            }

            rowStyle.IncludeAlignment = false;
            rowStyle.IncludeBorder = true;
            rowStyle.IncludeFont = false;
            rowStyle.IncludeNumber = false;
            rowStyle.IncludePatterns = true;
            rowStyle.IncludeProtection = false;

            return rowStyle;
        }
    }
}
