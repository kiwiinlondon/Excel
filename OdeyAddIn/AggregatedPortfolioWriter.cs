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
       

        private static int? GetNumericFieldColumn(Dictionary<Tuple<AggregatedPortfolioFields, string, DateTime>, int> detailColumnIds, AggregatedPortfolio portfolio, AggregatedPortfolioFields fieldId)
        {
            Tuple<AggregatedPortfolioFields, string, DateTime> key = new Tuple<AggregatedPortfolioFields, string, DateTime>(fieldId, portfolio.Fund, portfolio.ReferenceDate);
            int columnId;
            if (detailColumnIds.TryGetValue(key, out columnId))
            {
                return columnId;
            }
            return null;
        }



        public static void Write(List<AggregatedPortfolio> aggregatedPortfolio, Excel.Worksheet worksheet, int row, int column, EntityTypeIds entityTypeId, AggregatedPortfolioFields[] fieldsToReturn)
        {

            Dictionary<Tuple<AggregatedPortfolioFields, string, DateTime>, int> detailColumnIds = new Dictionary<Tuple<AggregatedPortfolioFields, string, DateTime>, int>();
                       
            DateTime[] referenceDates = aggregatedPortfolio.Select(p => p.ReferenceDate).Distinct().OrderBy(a => a).ToArray();
            string[] fundNames = aggregatedPortfolio.Select(p => p.Fund).Distinct().OrderBy(a => a).ToArray();

            
            int entityNameColumn = column++ ;
            List<AggregatedPortfolioFields> detailColumns = new List<AggregatedPortfolioFields>();

            if (fieldsToReturn.Contains(AggregatedPortfolioFields.Short))
            {
                detailColumns.Add(AggregatedPortfolioFields.Short);
            }
            if (fieldsToReturn.Contains(AggregatedPortfolioFields.ShortPercentNav))
            {
                detailColumns.Add(AggregatedPortfolioFields.ShortPercentNav);
            }
            if (fieldsToReturn.Contains(AggregatedPortfolioFields.Long))
            {
                detailColumns.Add(AggregatedPortfolioFields.Long);
            }
            if (fieldsToReturn.Contains(AggregatedPortfolioFields.LongPercentNav))
            {
                detailColumns.Add(AggregatedPortfolioFields.LongPercentNav);
            }
            
            if (fieldsToReturn.Contains(AggregatedPortfolioFields.Gross))
            {
                detailColumns.Add(AggregatedPortfolioFields.Gross);
            }
            if (fieldsToReturn.Contains(AggregatedPortfolioFields.GrossPercentNav))
            {
                detailColumns.Add(AggregatedPortfolioFields.GrossPercentNav);
            }            
            if (fieldsToReturn.Contains(AggregatedPortfolioFields.Net))
            {
                detailColumns.Add(AggregatedPortfolioFields.Net);
            }
            if (fieldsToReturn.Contains(AggregatedPortfolioFields.NetPercentNav))
            {
                detailColumns.Add(AggregatedPortfolioFields.NetPercentNav);
            }
            if (fieldsToReturn.Contains(AggregatedPortfolioFields.FundNav))
            {
                detailColumns.Add(AggregatedPortfolioFields.FundNav);
            }
            int parameterTitleRow = row++;
            int? fundTitleRow = null;
            if (fundNames.Length > 1)
            {
                fundTitleRow = row++;
            }
            int? referenceDateTitleRow = null;
            if (referenceDates.Length > 1)
            {
                referenceDateTitleRow = row++;
            }

            foreach (AggregatedPortfolioFields detailColumn in detailColumns)
            {
                ExcelWriter.WriteCell(worksheet, parameterTitleRow, column, detailColumn.ToString());
                foreach (string fundName in fundNames)
                {
                    ExcelWriter.WriteCell(worksheet, fundTitleRow, column, fundName);                    
                    foreach (DateTime referenceDate in referenceDates)
                    {
                        ExcelWriter.WriteCell(worksheet, referenceDateTitleRow, column, referenceDate,"dd-MMM-yyyy");
                        detailColumnIds.Add(new Tuple<AggregatedPortfolioFields, string, DateTime>(detailColumn, fundName, referenceDate), column++);
                    }
                }
            }
           
            worksheet.Cells[row-1, entityNameColumn] = String.Format("{0} Name", entityTypeId.ToString());            
            
            Dictionary<string, int> entityRowIds = new Dictionary<string, int>();
            foreach (AggregatedPortfolio aggregatedPortfolioItem in aggregatedPortfolio)
            {
                int entityRow;
                if (!entityRowIds.TryGetValue(aggregatedPortfolioItem.EntityName, out entityRow))
                {
                    entityRow = row;
                    ExcelWriter.WriteCell(worksheet, row, entityNameColumn, aggregatedPortfolioItem.EntityName);
                    entityRowIds.Add(aggregatedPortfolioItem.EntityName, entityRow);
                    row++;
                }
                int? grossColumn = GetNumericFieldColumn(detailColumnIds, aggregatedPortfolioItem, AggregatedPortfolioFields.Gross);
                int? grossPercentNavColumn = GetNumericFieldColumn(detailColumnIds, aggregatedPortfolioItem, AggregatedPortfolioFields.GrossPercentNav);

                int? netColumn = GetNumericFieldColumn(detailColumnIds, aggregatedPortfolioItem, AggregatedPortfolioFields.Net);
                int? netPercentNavColumn = GetNumericFieldColumn(detailColumnIds, aggregatedPortfolioItem, AggregatedPortfolioFields.NetPercentNav);

                int? shortColumn = GetNumericFieldColumn(detailColumnIds, aggregatedPortfolioItem, AggregatedPortfolioFields.Short);
                int? shortPercentNavColumn = GetNumericFieldColumn(detailColumnIds, aggregatedPortfolioItem, AggregatedPortfolioFields.ShortPercentNav);
                int? longColumn = GetNumericFieldColumn(detailColumnIds, aggregatedPortfolioItem, AggregatedPortfolioFields.Long);
                int? longPercentNavColumn = GetNumericFieldColumn(detailColumnIds, aggregatedPortfolioItem, AggregatedPortfolioFields.LongPercentNav);
                int? fundNavColumn = GetNumericFieldColumn(detailColumnIds, aggregatedPortfolioItem, AggregatedPortfolioFields.FundNav);

                decimal gross = Math.Abs(aggregatedPortfolioItem.Short) + aggregatedPortfolioItem.Long;
                decimal net = aggregatedPortfolioItem.Short + aggregatedPortfolioItem.Long;

                ExcelWriter.WriteCell(worksheet, entityRow, grossColumn, gross, "#,###");
                ExcelWriter.WriteCell(worksheet, entityRow, netColumn, net, "#,###"); 
                ExcelWriter.WriteCell(worksheet, entityRow, shortColumn, aggregatedPortfolioItem.Short, "#,###");                
                ExcelWriter.WriteCell(worksheet, entityRow, longColumn, aggregatedPortfolioItem.Long, "#,###");
                ExcelWriter.WriteCell(worksheet, entityRow, fundNavColumn, aggregatedPortfolioItem.FundMarketValue, "#,###");

                ExcelWriter.WritePercentage(worksheet, entityRow, shortPercentNavColumn, shortColumn, fundNavColumn,aggregatedPortfolioItem.Short, aggregatedPortfolioItem.FundMarketValue, "0.00%");
                ExcelWriter.WritePercentage(worksheet, entityRow, longPercentNavColumn, longColumn, fundNavColumn, aggregatedPortfolioItem.Long, aggregatedPortfolioItem.FundMarketValue, "0.00%");
                ExcelWriter.WritePercentage(worksheet, entityRow, grossPercentNavColumn, grossColumn, fundNavColumn, gross, aggregatedPortfolioItem.FundMarketValue, "0.00%");
                ExcelWriter.WritePercentage(worksheet, entityRow, netPercentNavColumn, netColumn, fundNavColumn, net, aggregatedPortfolioItem.FundMarketValue, "0.00%");
            }

            worksheet.Columns.AutoFit();
        }
    }
}
