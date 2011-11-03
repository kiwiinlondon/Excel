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
        internal static void WriteCell(Excel.Worksheet worksheet,int row, int? column, object value)
        {
            WriteCell(worksheet, row, column, value, null);
        }

        private static void WriteCell(Excel.Worksheet worksheet, int row, int? column, object value, string format)
        {
            if (column.HasValue)
            {
                worksheet.Cells[row, column.Value] = value;
                if (!string.IsNullOrWhiteSpace(format))
                {
                    worksheet.Columns[column.Value].NumberFormat = format;
                }
            }
        }

        private static int? GetNumericFieldColumn(int firstRepeatingColumn, int fundOffset, int numberOfRepeatingColumns, int? columnOffSet)
        {
            if (columnOffSet.HasValue)
            {
                return firstRepeatingColumn + fundOffset * numberOfRepeatingColumns + columnOffSet.Value;
            }
            else
            {
                return null;
            }
        }

        internal static void Write(List<PortfolioWithUnderlyer> portfolio, Excel.Worksheet worksheet, int row, int column, PortfolioFields[] fieldsToReturn)
        {
            Dictionary<int, int> instrumentMarketRowIds = new Dictionary<int, int>();
            Dictionary<int, int> fundOffsets = new Dictionary<int, int>();


            int? referenceDateColumn = null;
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

            int columnCount = 0;
            if (fieldsToReturn.Contains(PortfolioFields.ReferenceDate))
            {
                referenceDateColumn = column++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.InstrumentName))
            {
                instrumentNameColumn = column + columnCount++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.UnderlyingInstrumentName))
            {
                underlyingInstrumentNameColumn = column + columnCount++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.BBExchangeCode))
            {
                bbExchangeCodeColumn = column + columnCount++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.InstrumentClass))
            {
                instrumentClassColumn = column + columnCount++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.ParentInstrumentClass))
            {
                parentInstrumentClassColumn = column + columnCount++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.UnderlyerInstrumentClass))
            {
                underlyerInstrumentClassColumn = column + columnCount++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.UnderlyerParentInstrumentClass))
            {
                underlyerParentInstrumentClassColumn = column + columnCount++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.Country))
            {
                countryColumn = column + columnCount++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.UnderlyerCountry))
            {
                underlyerCountryColumn = column + columnCount++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.Industry))
            {
                industryColumn = column + columnCount++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.Sector))
            {
                sectorColumn = column + columnCount++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.UnderlyerIndustry))
            {
                underlyerIndustryColumn = column + columnCount++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.UnderlyerSector))
            {
                underlyerSectorColumn = column + columnCount++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.BloombergTicker))
            {
                bloombergTickerColumn = column + columnCount++;
            }
            if (fieldsToReturn.Contains(PortfolioFields.UnderlyingBloombergTicker))
            {
                underlyingBloombergTickerColumn = column + columnCount++;
            }

            int numberOfRepeatingColumns = 0;
            int? fundNameOffSet = null;
            int? netPositionColumnOffset = null;
            int? marketValueColumnOffset = null;
            int? deltaMarketValueColumnOffset = null;
            if (fieldsToReturn.Contains(PortfolioFields.NetPosition))
            {          
                fundNameOffSet = 0;
                netPositionColumnOffset = numberOfRepeatingColumns++;
            }

            if (fieldsToReturn.Contains(PortfolioFields.MarketValue))
            {
                fundNameOffSet = 0;
                marketValueColumnOffset = numberOfRepeatingColumns++;
            }

            if (fieldsToReturn.Contains(PortfolioFields.DeltaMarketValue))
            {
                fundNameOffSet = 0;
                deltaMarketValueColumnOffset = numberOfRepeatingColumns++;
            }

            int firstRepeatingColumn = column + columnCount;
            
            int titleRow = row;
            if (numberOfRepeatingColumns > 0)
            {
                titleRow = ++row; 
            }


            WriteCell(worksheet, titleRow, referenceDateColumn, "Reference Date");
            WriteCell(worksheet,titleRow, instrumentNameColumn, "Instrument Name");
            WriteCell(worksheet,titleRow, underlyingInstrumentNameColumn, "Underlying Instrument Name");
            WriteCell(worksheet,titleRow, bbExchangeCodeColumn, "Exchange Code");
            WriteCell(worksheet,titleRow, instrumentClassColumn, "Instrument Class");
            WriteCell(worksheet,titleRow, parentInstrumentClassColumn, "Parent Instrument Class");
            WriteCell(worksheet,titleRow, underlyerInstrumentClassColumn, "Underlyer's Instrument Class");
            WriteCell(worksheet,titleRow, underlyerParentInstrumentClassColumn, "Underlyer's Parent Instrument Class");
            WriteCell(worksheet,titleRow, countryColumn, "Country");
            WriteCell(worksheet,titleRow, underlyerCountryColumn, "Underlyer's Country");

            WriteCell(worksheet,titleRow, industryColumn, "Industry");
            WriteCell(worksheet,titleRow, sectorColumn, "Sector");
            WriteCell(worksheet,titleRow, underlyerIndustryColumn, "Underlyer's Industry");
            WriteCell(worksheet,titleRow, underlyerSectorColumn, "Underlyer's Sector");

            WriteCell(worksheet,titleRow, bloombergTickerColumn, "Bloomberg Ticker");
            WriteCell(worksheet,titleRow, underlyingBloombergTickerColumn, "Underlyer's Bloomberg Ticker");
            row++;

            int maxFundOffset = 0;
            foreach (PortfolioWithUnderlyer portfolioItem in portfolio)
            {
                int instrumentMarketRow;
                if (!instrumentMarketRowIds.TryGetValue(portfolioItem.InstrumentMarketId, out instrumentMarketRow))
                {
                    instrumentMarketRow = row;
                    WriteCell(worksheet,instrumentMarketRow, referenceDateColumn,portfolioItem.ReferenceDate);
                    WriteCell(worksheet,instrumentMarketRow, instrumentNameColumn,portfolioItem.InstrumentName);
                    WriteCell(worksheet,instrumentMarketRow, underlyingInstrumentNameColumn, portfolioItem.UnderlyingInstrumentName);
                    WriteCell(worksheet,instrumentMarketRow, bbExchangeCodeColumn, portfolioItem.BBExchangeCode);
                    WriteCell(worksheet,instrumentMarketRow, instrumentClassColumn, portfolioItem.InstrumentClass);
                    WriteCell(worksheet,instrumentMarketRow, parentInstrumentClassColumn, portfolioItem.ParentInstrumentClass);
                    WriteCell(worksheet,instrumentMarketRow, underlyerInstrumentClassColumn, portfolioItem.UnderlyerInstrumentClass);
                    WriteCell(worksheet,instrumentMarketRow, underlyerParentInstrumentClassColumn, portfolioItem.UnderlyerParentInstrumentClass);
                    WriteCell(worksheet,instrumentMarketRow, countryColumn, portfolioItem.Country);
                    WriteCell(worksheet,instrumentMarketRow, underlyerCountryColumn, portfolioItem.UnderlyerCountry);
                    WriteCell(worksheet,instrumentMarketRow, industryColumn, portfolioItem.Industry);
                    WriteCell(worksheet,instrumentMarketRow, sectorColumn, portfolioItem.Sector);
                    WriteCell(worksheet,instrumentMarketRow, underlyerIndustryColumn, portfolioItem.UnderlyerIndustry);
                    WriteCell(worksheet,instrumentMarketRow, underlyerSectorColumn, portfolioItem.UnderlyerSector);
                    WriteCell(worksheet,instrumentMarketRow, bloombergTickerColumn, portfolioItem.BloombergTicker);
                    WriteCell(worksheet,instrumentMarketRow, underlyingBloombergTickerColumn, portfolioItem.UnderlyingBloombergTicker);
                    instrumentMarketRowIds.Add(portfolioItem.InstrumentMarketId, instrumentMarketRow);
                    row++;
                }

                bool newFund = false;
                int fundOffset;
                if (!fundOffsets.TryGetValue(portfolioItem.FundId, out fundOffset))
                {
                    newFund = true;
                    fundOffset = maxFundOffset++;
                    fundOffsets.Add(portfolioItem.FundId, fundOffset);

                }
                int? netPositionColumn = GetNumericFieldColumn(firstRepeatingColumn,fundOffset, numberOfRepeatingColumns,netPositionColumnOffset);
                int? marketValueColumn =GetNumericFieldColumn(firstRepeatingColumn,fundOffset, numberOfRepeatingColumns,marketValueColumnOffset);
                int? deltaMarketValueColumn = GetNumericFieldColumn(firstRepeatingColumn,fundOffset, numberOfRepeatingColumns,deltaMarketValueColumnOffset);
                if (newFund)
                {
                    int? titleColumn = GetNumericFieldColumn(firstRepeatingColumn, fundOffset, numberOfRepeatingColumns, fundNameOffSet);
                    WriteCell(worksheet, titleRow - 1, titleColumn, portfolioItem.FundName);
                    WriteCell(worksheet,titleRow, netPositionColumn, "Net Position","#,###");
                    WriteCell(worksheet,titleRow, marketValueColumn, "Market Value","#,###");
                    WriteCell(worksheet,titleRow, deltaMarketValueColumn,"Delta Market Value","#,###");

                }

                WriteCell(worksheet,instrumentMarketRow, netPositionColumn,portfolioItem.NetPosition);
                WriteCell(worksheet,instrumentMarketRow, marketValueColumn,portfolioItem.MarketValue);
                WriteCell(worksheet,instrumentMarketRow, deltaMarketValueColumn,portfolioItem.DeltaMarketValue);

            }
            worksheet.Columns.AutoFit();
        }
    }


}

