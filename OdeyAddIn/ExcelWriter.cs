using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace OdeyAddIn
{
    public static class ExcelWriter
    {

        private static string GetColumnName(int columnNumber)
        {

            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;

        }

        public static void WritePercentage(Excel.Worksheet worksheet, int row, int? column, int? nominatorColumn, int? denominatorColumn, decimal nominator, decimal denominator, string format)
        {
            if (column.HasValue)
            {
                if (nominatorColumn.HasValue)
                {
                    string nominatorColumnLabel = GetColumnName(nominatorColumn.Value);
                    if (denominatorColumn.HasValue)
                    {
                        string denominatorColumnLabel = GetColumnName(denominatorColumn.Value);
                        worksheet.Cells[row, column.Value].Formula = String.Format("={0}{1}/{2}{1}", nominatorColumnLabel, row, denominatorColumnLabel);
                    }
                    else
                    {
                        worksheet.Cells[row, column.Value].Formula = String.Format("={0}{1}/{2}", nominatorColumnLabel, row, denominator);
                    }
                }
                else //no nominator column
                {
                    if (denominatorColumn.HasValue)
                    {
                        string denominatorColumnLabel = GetColumnName(denominatorColumn.Value);
                        worksheet.Cells[row, column.Value].Formula = String.Format("={0}/{2}{1}", nominator, row, denominatorColumnLabel);
                    }
                    else
                    {
                        worksheet.Cells[row, column.Value].Formula = String.Format("={0}/{1}", nominator, denominator);
                    }

                }
                if (!string.IsNullOrWhiteSpace(format))
                {
                    worksheet.Cells[row, column].NumberFormat = format;
                }
            }            
        }

        public static void WriteCellApendCurrencyName(Excel.Worksheet worksheet, int row, int? column, object value, string currency, string format)
        {
            if (!string.IsNullOrWhiteSpace(currency))
            {
                value = String.Format("{0}({1})", value, currency);
            }
            WriteCell(worksheet, row, column, value, format);
        }

        public static void WriteCell(Excel.Worksheet worksheet, int? row, int? column, decimal value, decimal? fxRate, string format)
        {
            if (fxRate.HasValue)
            {
                value = value * fxRate.Value;
            }
            WriteCell(worksheet, row, column, value, format);
        }

        public static void WriteCell(Excel.Worksheet worksheet, int? row, int? column, object value)
        {
            WriteCell(worksheet, row, column, value, null);
        }

        public static void WriteCell(Excel.Worksheet worksheet, int? row, int? column, object value, string format)
        {
            if (column.HasValue && row.HasValue)
            {
                worksheet.Cells[row, column.Value] = value;
                if (!string.IsNullOrWhiteSpace(format))
                {
                    worksheet.Cells[row, column].NumberFormat = format;
                }
            }
        }
    }
}
