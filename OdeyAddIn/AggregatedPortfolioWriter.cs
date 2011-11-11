using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Odey.Reporting.Entities;
using Excel = Microsoft.Office.Interop.Excel;
using Odey.Framework.Keeley.Entities.Enums;
using System.Reflection;

namespace OdeyAddIn
{
    public static class AggregatedPortfolioWriter
    {
        private static string GetExcelColumnName(int columnNumber)
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


        public static void Write(List<AggregatedPortfolio> aggregatedPortfolio, Excel.Worksheet worksheet, int row, int column, EntityTypeIds entityTypeId)
        {
            int referenceDateColumn = column;
            int fundColumn = column+1;
            int entityNameColumn = column + 2;
            int shortColumn = column + 3;
            int shortPercentOfNavColumn = column + 4;
            int longColumn = column + 5;
            int longPercentOfNavColumn = column + 6;
            int fundNavColumn = column + 7;
            
            string longColumnLabel = GetExcelColumnName(longColumn);
            string shortColumnLabel = GetExcelColumnName(shortColumn);
            string fundNavColumnLabel = GetExcelColumnName(fundNavColumn);
            worksheet.Cells[row, referenceDateColumn] = "Reference Date";
            worksheet.Cells[row, fundColumn] = "Fund";
            worksheet.Cells[row, entityNameColumn] = String.Format("{0} Name", entityTypeId.ToString());
            worksheet.Cells[row, shortColumn] = String.Format("Short");
            worksheet.Columns[shortColumn].NumberFormat = "#,###";
            worksheet.Cells[row, shortPercentOfNavColumn] = String.Format("% of NAV");
            worksheet.Columns[shortPercentOfNavColumn].NumberFormat = "0.00%";
            
            worksheet.Cells[row, longColumn] = String.Format("Long");
            worksheet.Columns[longColumn].NumberFormat = "#,###";
            worksheet.Cells[row, longPercentOfNavColumn] = String.Format("% of NAV");
            worksheet.Columns[longPercentOfNavColumn].NumberFormat = "0.00%";

            worksheet.Cells[row, fundNavColumn] = String.Format("Fund NAV");
            worksheet.Columns[fundNavColumn].NumberFormat = "#,###";
            row++;
           
            foreach (AggregatedPortfolio aggregatedPortfolioItem in aggregatedPortfolio)
            {
                worksheet.Cells[row, referenceDateColumn] = aggregatedPortfolioItem.ReferenceDate;
                worksheet.Cells[row, fundColumn] = aggregatedPortfolioItem.Fund;
                worksheet.Cells[row, entityNameColumn] = aggregatedPortfolioItem.EntityName;
                worksheet.Cells[row, shortColumn] = aggregatedPortfolioItem.Short;
                worksheet.Cells[row, longColumn] = aggregatedPortfolioItem.Long;
                worksheet.Cells[row, fundNavColumn] = aggregatedPortfolioItem.FundMarketValue;
              //  Excel.Range r = worksheet.Cells[row, marketValueColumn];
              //  string a = GetExcelColumnName(r.Column);
                worksheet.Cells[row, shortPercentOfNavColumn].Formula = String.Format("={0}{1}/{2}{1}", shortColumnLabel, row, fundNavColumnLabel);
                worksheet.Cells[row, longPercentOfNavColumn].Formula = String.Format("={0}{1}/{2}{1}", longColumnLabel, row, fundNavColumnLabel);
                row++;
            }
            worksheet.Columns.AutoFit();
        }
    }
}
