﻿using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

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
            return sheet.Controls.AddListObject(sheet.Cells[row, column], name);
        }

        public static void SetColumnWidth(this Excel.Worksheet sheet, int column, int width)
        {
            Excel.Range cell = sheet.Cells[1, column];
            cell.ColumnWidth = width;
        }

        public static Excel.Range GetCell(this Worksheet sheet, int row, int column)
        {
            return sheet.Cells[row, column];
        }
    }
}
