using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Odey.Reporting.Entities;
using Excel = Microsoft.Office.Interop.Excel;
using Odey.Framework.Keeley.Entities.Enums;

namespace OdeyAddIn
{
    public static class AggregatedPortfolioWriter
    {
        public static void Write(List<AggregatedPortfolio> aggregatedPortfolio, Excel.Worksheet worksheet, int row, int column, EntityTypeIds entityTypeId)
        {
            int referenceDateColumn = column;
            int fundColumn = column+1;
            int entityNameColumn = column + 2;
            int longShortColumn = column + 3;
            int marketValueColumn = column + 4;
            int fundNavColumn = column + 5;
            int percentOfNav = column + 6;

            worksheet.Cells[row, referenceDateColumn] = "Reference Date";
            worksheet.Cells[row, fundColumn] = "Fund";
            worksheet.Cells[row, entityNameColumn] = String.Format("{0} Name", entityTypeId.ToString());
            worksheet.Cells[row, longShortColumn] = String.Format("Long/Short");
            worksheet.Cells[row, marketValueColumn] = String.Format("Delta Market Value");
            worksheet.Columns[marketValueColumn].NumberFormat = "#,###";
            worksheet.Cells[row, fundNavColumn] = String.Format("Fund NAV");
            worksheet.Columns[fundNavColumn].NumberFormat = "#,###";
            worksheet.Cells[row++, percentOfNav] = String.Format("% of NAV");
            worksheet.Columns[percentOfNav].NumberFormat = "0.00%";
           
            foreach (AggregatedPortfolio aggregatedPortfolioItem in aggregatedPortfolio)
            {
                worksheet.Cells[row, referenceDateColumn] = aggregatedPortfolioItem.ReferenceDate;
                worksheet.Cells[row, fundColumn] = aggregatedPortfolioItem.Fund;
                worksheet.Cells[row, entityNameColumn] = aggregatedPortfolioItem.EntityName;
                worksheet.Cells[row, longShortColumn] = aggregatedPortfolioItem.IsLong ? 1 : -1;
                worksheet.Cells[row, marketValueColumn] = aggregatedPortfolioItem.DeltaMarketValue;
                worksheet.Cells[row, fundNavColumn] = aggregatedPortfolioItem.FundMarketValue;
                worksheet.Cells[row, percentOfNav].Formula = String.Format("={0}/{1}",worksheet.Cells[row, marketValueColumn].Address, worksheet.Cells[row, fundNavColumn].Address);
                row++;
            }
            worksheet.Columns.AutoFit();
        }
    }
}
