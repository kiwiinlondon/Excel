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

        public static void WriteColumnHeader(this Excel.Worksheet sheet, int row, int column, ColumnDef col, Excel.Style style)
        {
            Excel.Range header = sheet.Cells[row, column];
            header.Value = col.Name;
            header.ColumnWidth = col.Width;
            if (style != null)
            {
                header.Style = style;
            }
        }

        public static void WriteIndexColumn(this Excel.Worksheet sheet, int row, int column, ColumnDef col, int max, Excel.Style rowStyle)
        {
            for (var y = 0; y < max; ++y)
            {
                Excel.Range cell = sheet.Cells[row + y, column];
                cell.Value2 = y + 1;
                cell.Style = rowStyle;
            }
        }

        public static void WriteFieldColumn<T>(this Excel.Worksheet sheet, int row, int column, ColumnDef col, IEnumerable<T> data, string field, Excel.Style rowStyle)
        {
            var y = 0;
            var pi = typeof(T).GetProperty(field);
            foreach (var item in data)
            {
                Excel.Range cell = sheet.Cells[row + y, column];
                cell.Value2 = pi.GetValue(item);
                cell.Style = rowStyle;
                if (col.NumberFormat != null)
                {
                    cell.NumberFormat = col.NumberFormat;
                }
                ++y;
            }
        }

        public static void WriteWatchListColumn(this Excel.Worksheet sheet, int row, int column, ColumnDef col, IEnumerable<dynamic> data, Excel.Style rowStyle, Dictionary<string, WatchListItem> watchList, ColumnDef sourceColumn, string formula = "=[Address]")
        {
            var y = 0;
            foreach (var item in data)
            {
                var address = GetAddress(item.Ticker, sourceColumn.AlphabeticalIndex, watchList);
                Excel.Range cell = sheet.Cells[row + y, column];
                cell.Formula = formula.Replace("[Address]", address);
                cell.Style = rowStyle;
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
            rowStyle.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(0xDDD9C4));   //0xEEECE1

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
