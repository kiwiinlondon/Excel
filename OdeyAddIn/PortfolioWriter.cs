using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Odey.Reporting.Entities;
using Excel = Microsoft.Office.Interop.Excel;

namespace OdeyAddIn
{
    public static class PortfolioWriter
    {
        

        private static int? GetNumericFieldColumn(Dictionary<Tuple<PortfolioFields, string, DateTime>, int> detailColumnIds, CompletePortfolio portfolio, PortfolioFields fieldId)
        {
            Tuple<PortfolioFields, string, DateTime> key = new Tuple<PortfolioFields, string, DateTime>(fieldId, portfolio.FundName, portfolio.ReferenceDate);
            int columnId;
            if (detailColumnIds.TryGetValue(key, out columnId))
            {
                return columnId;
            }
            return null;
        }

        internal static void Write(List<CompletePortfolio> portfolio, Excel.Worksheet worksheet, int row, int column, PortfolioFields[] fieldsToReturn,string currency)
        {
            Dictionary<int, int> instrumentMarketRowIds = new Dictionary<int, int>();
            Dictionary<Tuple<PortfolioFields,string, DateTime>, int> detailColumnIds = new Dictionary<Tuple<PortfolioFields,string,DateTime>,int>();
            DateTime[] referenceDates = portfolio.Select(p => p.ReferenceDate).Distinct().OrderBy(a => a).ToArray();
            string[] fundNames = portfolio.Select(p => p.FundName).Distinct().OrderBy(a => a).ToArray();

            int? instrumentNameColumn = null;
            int? underlyingInstrumentNameColumn = null;
            int? bbExchangeCodeColumn = null;
            int? instrumentClassColumn = null;
            int? parentInstrumentClassColumn = null;
            int? underlyerInstrumentClassColumn = null;
            int? underlyerParentInstrumentClassColumn = null;
            int? countryColumn = null;
            int? underlyerCountryColumn = null;
            int? industryColumn = null;
            int? sectorColumn = null;
            int? underlyerIndustryColumn = null;
            int? underlyerSectorColumn = null;
            int? bloombergTickerColumn = null;
            int? underlyingBloombergTickerColumn = null;

            
            if (fieldsToReturn.Contains(PortfolioFields.InstrumentName))
            {
                instrumentNameColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.UnderlyingInstrumentName))
            {
                underlyingInstrumentNameColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.BBExchangeCode))
            {
                bbExchangeCodeColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.InstrumentClass))
            {
                instrumentClassColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.ParentInstrumentClass))
            {
                parentInstrumentClassColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.UnderlyerInstrumentClass))
            {
                underlyerInstrumentClassColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.UnderlyerParentInstrumentClass))
            {
                underlyerParentInstrumentClassColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.Country))
            {
                countryColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.UnderlyerCountry))
            {
                underlyerCountryColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.Industry))
            {
                industryColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.Sector))
            {
                sectorColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.UnderlyerIndustry))
            {
                underlyerIndustryColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.UnderlyerSector))
            {
                underlyerSectorColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.BloombergTicker))
            {
                bloombergTickerColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.UnderlyingBloombergTicker))
            {
                underlyingBloombergTickerColumn = column++;
            }            

            List<PortfolioFields> detailColumns = new List<PortfolioFields>();
            if (fieldsToReturn.Contains(PortfolioFields.NetPosition))
            {
                detailColumns.Add(PortfolioFields.NetPosition);
            }

            if (fieldsToReturn.Contains(PortfolioFields.MarketValue))
            {
                detailColumns.Add(PortfolioFields.MarketValue);
            }

            if (fieldsToReturn.Contains(PortfolioFields.DeltaMarketValue))
            {
                detailColumns.Add(PortfolioFields.DeltaMarketValue);
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

            foreach (PortfolioFields detailColumn in detailColumns)
            {
                ExcelWriter.WriteCell(worksheet, parameterTitleRow, column, detailColumn.ToString());
                foreach (string fundName in fundNames)
                {
                    ExcelWriter.WriteCell(worksheet, fundTitleRow, column, fundName);
                    foreach (DateTime referenceDate in referenceDates)
                    {
                        ExcelWriter.WriteCell(worksheet, referenceDateTitleRow, column, referenceDate,"dd-MMM-yyyy");
                        detailColumnIds.Add(new Tuple<PortfolioFields, string, DateTime>(detailColumn, fundName, referenceDate), column++);
                    }
                }
            }
                   
            int titleRow = row-1;
            ExcelWriter.WriteCell(worksheet, titleRow, instrumentNameColumn, "Instrument Name");
            ExcelWriter.WriteCell(worksheet, titleRow, underlyingInstrumentNameColumn, "Underlying Instrument Name");
            ExcelWriter.WriteCell(worksheet, titleRow, bbExchangeCodeColumn, "Exchange Code");
            ExcelWriter.WriteCell(worksheet, titleRow, instrumentClassColumn, "Instrument Class");
            ExcelWriter.WriteCell(worksheet, titleRow, parentInstrumentClassColumn, "Parent Instrument Class");
            ExcelWriter.WriteCell(worksheet, titleRow, underlyerInstrumentClassColumn, "Underlyer's Instrument Class");
            ExcelWriter.WriteCell(worksheet, titleRow, underlyerParentInstrumentClassColumn, "Underlyer's Parent Instrument Class");
            ExcelWriter.WriteCell(worksheet, titleRow, countryColumn, "Country");
            ExcelWriter.WriteCell(worksheet, titleRow, underlyerCountryColumn, "Underlyer's Country");

            ExcelWriter.WriteCell(worksheet, titleRow, industryColumn, "Industry");
            ExcelWriter.WriteCell(worksheet, titleRow, sectorColumn, "Sector");
            ExcelWriter.WriteCell(worksheet, titleRow, underlyerIndustryColumn, "Underlyer's Industry");
            ExcelWriter.WriteCell(worksheet, titleRow, underlyerSectorColumn, "Underlyer's Sector");

            ExcelWriter.WriteCell(worksheet, titleRow, bloombergTickerColumn, "Bloomberg Ticker");
            ExcelWriter.WriteCell(worksheet, titleRow, underlyingBloombergTickerColumn, "Underlyer's Bloomberg Ticker");

            foreach (CompletePortfolio portfolioItem in portfolio)
            {
                int instrumentMarketRow;
                if (!instrumentMarketRowIds.TryGetValue(portfolioItem.InstrumentMarketId, out instrumentMarketRow))
                {
                    instrumentMarketRow = row;
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, instrumentNameColumn, portfolioItem.InstrumentName);
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, underlyingInstrumentNameColumn, portfolioItem.UnderlyingInstrumentName);
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, bbExchangeCodeColumn, portfolioItem.BBExchangeCode);
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, instrumentClassColumn, portfolioItem.InstrumentClass);
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, parentInstrumentClassColumn, portfolioItem.ParentInstrumentClass);
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, underlyerInstrumentClassColumn, portfolioItem.UnderlyerInstrumentClass);
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, underlyerParentInstrumentClassColumn, portfolioItem.UnderlyerParentInstrumentClass);
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, countryColumn, portfolioItem.Country);
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, underlyerCountryColumn, portfolioItem.UnderlyerCountry);
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, industryColumn, portfolioItem.Industry);
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, sectorColumn, portfolioItem.Sector);
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, underlyerIndustryColumn, portfolioItem.UnderlyerIndustry);
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, underlyerSectorColumn, portfolioItem.UnderlyerSector);
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, bloombergTickerColumn, portfolioItem.BloombergTicker);
                    ExcelWriter.WriteCell(worksheet, instrumentMarketRow, underlyingBloombergTickerColumn, portfolioItem.UnderlyingBloombergTicker);
                    instrumentMarketRowIds.Add(portfolioItem.InstrumentMarketId, instrumentMarketRow);
                    row++;
                }

                int? netPositionColumn = GetNumericFieldColumn(detailColumnIds, portfolioItem,PortfolioFields.NetPosition);
                int? marketValueColumn = GetNumericFieldColumn(detailColumnIds, portfolioItem, PortfolioFields.MarketValue);
                int? deltaMarketValueColumn = GetNumericFieldColumn(detailColumnIds, portfolioItem, PortfolioFields.DeltaMarketValue);


                ExcelWriter.WriteCell(worksheet, instrumentMarketRow, netPositionColumn, portfolioItem.NetPosition, "#,###");
                ExcelWriter.WriteCell(worksheet, instrumentMarketRow, marketValueColumn, portfolioItem.MarketValue, portfolioItem.FXRateToReportCurrency, "#,###");
                ExcelWriter.WriteCell(worksheet, instrumentMarketRow, deltaMarketValueColumn, portfolioItem.DeltaMarketValue, portfolioItem.FXRateToReportCurrency, "#,###");

            }
            worksheet.Columns.AutoFit();
        }
    }


}

