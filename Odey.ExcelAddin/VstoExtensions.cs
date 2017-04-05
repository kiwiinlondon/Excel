using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Diagnostics;
using System;
using System.Collections.Generic;

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

        public static void WriteColumnHeader(this Excel.Worksheet sheet, int row, int column, ColumnDef col)
        {
            Excel.Range header = sheet.Cells[row, column];
            header.Value = col.Name;
            header.ColumnWidth = col.Width;
        }

        public static void WriteIndexColumn(this Excel.Worksheet sheet, int row, int column, ColumnDef col, int max)
        {
            for (var y = 0; y < max; ++y)
            {
                sheet.Cells[row + y, column] = y + 1;
            }
        }

        public static void WriteFieldColumn<T>(this Excel.Worksheet sheet, int row, int column, ColumnDef col, IEnumerable<T> data, string field)
        {
            var y = 0;
            var pi = typeof(T).GetProperty(field);
            foreach (var item in data)
            {
                Excel.Range cell = sheet.Cells[row + y, column];
                cell.Value = pi.GetValue(item);
                if (col.NumberFormat != null)
                {
                    cell.NumberFormat = col.NumberFormat;
                }
                ++y;
            }
        }

        public static void WriteWatchListColumn(this Excel.Worksheet sheet, int row, int column, ColumnDef col, IEnumerable<dynamic> data, Dictionary<string, WatchListItem> watchList, ColumnDef sourceColumn)
        {
            var y = 0;
            foreach (var item in data)
            {
                Excel.Range cell = sheet.Cells[row + y, column];
                var watchListItem = watchList[item.Ticker];
                cell.Formula = $"='{WatchListSheet.Name}'!{sourceColumn.AlphabeticalIndex}{watchListItem.RowIndex}";
                ++y;
            }
        }
    }
}
