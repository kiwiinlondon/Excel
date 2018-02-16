using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XL=Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class SheetAccess
    {
        public SheetAccess(ThisWorkbook workBook)
        {
            _worksheet = workBook.Sheets["Portfolio"];
        }
        private XL.Worksheet _worksheet;

        private static readonly string _controlColumn = "A";
        private static readonly int _controlColumnNumber = GetColumnNumber(_controlColumn);

        private static readonly string _tickerColumn = "B";
        private static readonly int _tickerColumnNumber = GetColumnNumber(_tickerColumn);

        private static readonly string _currencyColumn = "C";
        private static readonly int _currencyColumnNumber = GetColumnNumber(_currencyColumn);

        private static readonly string _nameColumn = "D";
        private static readonly int _nameColumnNumber = GetColumnNumber(_nameColumn);

        private static readonly string _closePriceColumn = "E";
        private static readonly int _closePriceColumnNumber = GetColumnNumber(_closePriceColumn);

        private static readonly string _currentPriceColumn = "F";
        private static readonly int _currentPriceColumnNumber = GetColumnNumber(_currentPriceColumn);

        private static readonly string _priceChangeColumn = "G";
        private static readonly int _priceChangeColumnNumber = GetColumnNumber(_priceChangeColumn);

        private static readonly string _pricePercentageChangeColumn = "H";
        private static readonly int _pricePercentageChangeColumnNumber = GetColumnNumber(_pricePercentageChangeColumn);

        private static readonly string _netPositionColumn = "I";
        private static readonly int _netPositionColumnNumber = GetColumnNumber(_netPositionColumn);

        private static readonly string _currencyTickerColumn = "J";
        private static readonly int _currencyTickerColumnNumber = GetColumnNumber(_currencyTickerColumn);

        private static readonly string _quoteFactorColumn = "K";
        private static readonly int _quoteFactorColumnNumber = GetColumnNumber(_quoteFactorColumn);

        private static readonly string _fxRateColumn = "L";
        private static readonly int _fxRateColumnNumber = GetColumnNumber(_fxRateColumn);

        private static readonly string _pnlColumn = "M";
        private static readonly int _pnlColumnNumber = GetColumnNumber(_pnlColumn);

        private static readonly string _contributionColumn = "N";
        private static readonly int _contributionColumnNumber = GetColumnNumber(_contributionColumn);

        private static readonly string _exposureColumn = "O";
        private static readonly int _exposureColumnNumber = GetColumnNumber(_exposureColumn);

        private static readonly string _exposurePercentageColumn = "P";
        private static readonly int _exposurePercentageColumnNumber = GetColumnNumber(_exposurePercentageColumn);

        private static readonly string _shortColumn = "Q";
        private static readonly int _shortColumnNumber = GetColumnNumber(_shortColumn);

        private static readonly string _longColumn = "R";
        private static readonly int _longColumnNumber = GetColumnNumber(_longColumn);

        private static readonly string _priceMultiplierColumn = "S";
        private static readonly int _priceMultiplierColumnNumber = GetColumnNumber(_priceMultiplierColumn);

        private static readonly string _tickerTypeColumn = "T";
        private static readonly int _tickerTypeColumnNumber = GetColumnNumber(_tickerTypeColumn);

        private static readonly string _priceDivisorColumn = "U";
        private static readonly int _priceDivisorColumnNumber = GetColumnNumber(_priceDivisorColumn);

        private static readonly string _lastColumn = _tickerTypeColumn;

        private static readonly string _firstColumn = _controlColumn;

        private int? FindRow(string toFind, string column)
        {
            XL.Range tickers = _worksheet.get_Range($"${column}:${column}");
            var currentFind = tickers.Find(toFind, System.Reflection.Missing.Value,
                XL.XlFindLookIn.xlFormulas, XL.XlLookAt.xlPart,//use xl forulas as column is hidden
                XL.XlSearchOrder.xlByRows, XL.XlSearchDirection.xlNext, false,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            if (currentFind == null)
            {
                return null;
            }
            else
            {
                return (int)currentFind.Row;
            }
        }

        private static int GetColumnNumber(string letter)
        {
            if (letter.Length != 1)
            {
                throw new ApplicationException($"Dont know how to convert strings that are not one char long ({letter})");
            }

            return char.ToUpper(letter[0]) - 'A' + 1;

        }

        public void UpdateTickerRow(bool forceRefresh, Location location)
        {
            if (forceRefresh)
            {
                WriteRow(location);
            }
            else
            {
                if (location.TickerTypeId.HasValue)
                {
                    WritePrivatePlacement(location);
                }
                else
                {
                    if (location.QuantityHasChanged)
                    {
                        WriteValue(location.Row.Value, _netPositionColumnNumber, location.NetPosition, null);
                    }
                }
            }
        }

        public void AddTickerRow(Location location, int rowNumber)
        {
            location.Row = rowNumber;
            AddRow(location.Row.Value);
            WriteRow(location);
        }

        private void WriteRow(Location location)
        {
            WriteValue(location.Row.Value, _tickerColumnNumber, location.Ticker, null);

            if (location.TickerTypeId.HasValue)
            {
                WritePrivatePlacement(location);
            }
            else
            {
                WriteNormal(location);
            }

            WriteFormula(location.Row.Value, _priceChangeColumnNumber, GetSubtractFormula(location.Row.Value, _currentPriceColumn, _closePriceColumn), null);
            WriteFormula(location.Row.Value, _pricePercentageChangeColumnNumber, GetDivideFormula(location.Row.Value, _priceChangeColumn, _closePriceColumn, false), null);
            WriteValue(location.Row.Value, _netPositionColumnNumber, location.NetPosition, false);
            WriteFormula(location.Row.Value, _currencyTickerColumnNumber, GetCurrencyTickerFormula(location.Row.Value), null);
            WriteFormula(location.Row.Value, _quoteFactorColumnNumber, GetQuoteFactorFormula(location.Row.Value), null);
            WriteFormula(location.Row.Value, _fxRateColumnNumber, GetFXRateFormula(location.Row.Value), null);
            WriteFormula(location.Row.Value, _pnlColumnNumber, GetMultiplyFormula(location.Row.Value, new string[] { _priceChangeColumn, _netPositionColumn, _priceMultiplierColumn }, new string[] { _fxRateColumn }), null);
            WriteFormula(location.Row.Value, _contributionColumnNumber, GetDivideByNavFormula(location.Row.Value, _pnlColumn, true), null);

            WriteFormula(location.Row.Value, _exposureColumnNumber, GetMultiplyFormula(location.Row.Value, new string[] { _currentPriceColumn, _netPositionColumn, _priceMultiplierColumn },new string[] { _fxRateColumn}), null);
            WriteFormula(location.Row.Value, _exposurePercentageColumnNumber, GetDivideByNavFormula(location.Row.Value, _exposureColumn, false), null);

            WriteFormula(location.Row.Value, _shortColumnNumber, GetWriteIfCorrectSignColumn(location.Row.Value, false, _exposurePercentageColumn), null);
            WriteFormula(location.Row.Value, _longColumnNumber, GetWriteIfCorrectSignColumn(location.Row.Value, true, _exposurePercentageColumn), null);
            WriteFormula(location.Row.Value, _priceMultiplierColumnNumber, GetPriceMultiplierFormula(location.Row.Value), null);
            WriteValue(location.Row.Value, _tickerTypeColumnNumber, location.TickerTypeId, false);
            WriteValue(location.Row.Value, _priceDivisorColumnNumber, location.PriceDivisor, false);
        }



        private void WriteNormal(Location location)
        {
            WriteFormula(location.Row.Value, _currencyColumnNumber, GetBloombergMnemonicFormula(location.Row.Value, _currencyColumn), null);
            WriteFormula(location.Row.Value, _nameColumnNumber, GetBloombergMnemonicFormula(location.Row.Value, _nameColumn), null);
            WriteFormula(location.Row.Value, _closePriceColumnNumber, GetBloombergMnemonicFormula(location.Row.Value, _closePriceColumn), null);
            WriteFormula(location.Row.Value, _currentPriceColumnNumber, GetBloombergMnemonicFormula(location.Row.Value, _currentPriceColumn), null);
        }

        private void WritePrivatePlacement(Location location)
        {
            WriteValue(location.Row.Value, _currencyColumnNumber, location.Currency, null);
            WriteValue(location.Row.Value, _nameColumnNumber, location.Name, null);
            WriteValue(location.Row.Value, _closePriceColumnNumber, location.OdeyPrice, null);
            WriteValue(location.Row.Value, _currentPriceColumnNumber, location.OdeyPrice, null);
        }

        public void UpdateSums(CountryLocation country)
        {
            WriteFormula(country.TotalRow.Value, _pnlColumnNumber, GetSumFormula(country.FirstRow.Value,country.TotalRow.Value-1,_pnlColumn),true);
            WriteFormula(country.TotalRow.Value, _contributionColumnNumber, GetSumFormula(country.FirstRow.Value, country.TotalRow.Value - 1, _contributionColumn),false);
            WriteFormula(country.TotalRow.Value, _exposureColumnNumber, GetSumFormula(country.FirstRow.Value, country.TotalRow.Value - 1, _exposureColumn),true);
            WriteFormula(country.TotalRow.Value, _exposurePercentageColumnNumber, GetSumFormula(country.FirstRow.Value, country.TotalRow.Value - 1, _exposurePercentageColumn), true);
            WriteFormula(country.TotalRow.Value, _shortColumnNumber, GetSumFormula(country.FirstRow.Value, country.TotalRow.Value - 1, _shortColumn), true);
            WriteFormula(country.TotalRow.Value, _longColumnNumber, GetSumFormula(country.FirstRow.Value, country.TotalRow.Value - 1, _longColumn), true);

        }
              
        private void AddRow(int row)
        {
            LastRow++;
            _worksheet.Rows[row].Insert(XL.XlDirection.xlUp, XL.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            _worksheet.Rows[row].Font.Bold = false;
        }

        private string GetCountryTotalString(string isoCode)
        {
            return $"{isoCode}{_countryTotalSuffix}";
        }


        private void UpdateTotalsOnTotalRow(int rowNumberToAdd)
        {
            UpdateTotalOnTotalRow(_pnlColumnNumber, _pnlColumn, rowNumberToAdd);
            UpdateTotalOnTotalRow(_contributionColumnNumber, _contributionColumn, rowNumberToAdd);
            UpdateTotalOnTotalRow(_exposureColumnNumber, _exposureColumn, rowNumberToAdd);
            UpdateTotalOnTotalRow(_exposurePercentageColumnNumber, _exposurePercentageColumn, rowNumberToAdd);
            UpdateTotalOnTotalRow(_shortColumnNumber, _shortColumn, rowNumberToAdd);
            UpdateTotalOnTotalRow(_longColumnNumber, _longColumn, rowNumberToAdd);
        }

        private void UpdateTotalOnTotalRow(int columnNumber,string column,int rowNumberToAdd)
        {
            var cell = _worksheet.Cells[LastRow, columnNumber];
            cell.Formula = $"{cell.Formula}+{column}{rowNumberToAdd}";
        }

        public void AddCountryTotalRow(int lastTotalRow, CountryLocation country)
        {
            AddRow(lastTotalRow);
            AddRow(lastTotalRow);
            WriteValue(lastTotalRow, _controlColumnNumber, GetCountryTotalString(country.IsoCode),false);
            WriteValue(lastTotalRow, _tickerColumnNumber, country.Name,true);
            country.TotalRow = lastTotalRow;
            country.FirstRow = lastTotalRow;
            var range = _worksheet.Range[_worksheet.Cells[lastTotalRow, _firstColumn], _worksheet.Cells[lastTotalRow, _lastColumn]];
            XL.Borders borders = range.Borders;
            borders[XL.XlBordersIndex.xlEdgeBottom].LineStyle = XL.XlLineStyle.xlContinuous;
            borders[XL.XlBordersIndex.xlEdgeTop].LineStyle = XL.XlLineStyle.xlContinuous;
            UpdateTotalsOnTotalRow(lastTotalRow);
        }

        public int LastRow { get; private set; }

        public Dictionary<string,CountryLocation> GetCountries()
        {
            
            int? lastRow = FindRow(_finalTotalName, _controlColumn);

            if (!lastRow.HasValue)
            {
                throw new ApplicationException("No Total Row exists");
            }
            LastRow = lastRow.Value;

            Dictionary<string, CountryLocation> countryLocations = new Dictionary<string, CountryLocation>();
            CountryLocation countryLocation = null;
            for (int i = _firstRowOfData;i< lastRow;i++)
            {               
                XL.Range row = _worksheet.get_Range($"{_firstColumn}{i}:{_lastColumn}{i}");

                string valueInTickerColumn = GetStringValue(row, _tickerColumnNumber);
                string valueInControlColumn = GetStringValue(row, _controlColumnNumber);
                
                if (string.IsNullOrWhiteSpace(valueInTickerColumn) && string.IsNullOrWhiteSpace(valueInControlColumn))
                {
                    continue;
                }

                if (countryLocation == null)
                {
                    countryLocation = new CountryLocation();
                    countryLocation.FirstRow = i;
                }

                string isoCode;
                if (RowIsCountryTotal(valueInControlColumn, out isoCode))
                {
                    countryLocation.IsoCode = isoCode;
                    countryLocation.Name = valueInTickerColumn;
                    countryLocation.TotalRow = i;
                    countryLocations.Add(isoCode, countryLocation);
                    countryLocation = null;
                }
                else
                {
                    var location = BuildLocation(row, valueInTickerColumn);
                    if (location!= null)
                    {
                        countryLocation.TickerRows.Add(location.Ticker,location);
                    }
                }                                            
            }
            return countryLocations;
        }


        private bool RowIsCountryTotal(string valueInControlColumn, out string isoCode)
        {
            isoCode = null;

            if (valueInControlColumn != null && valueInControlColumn.EndsWith(_countryTotalSuffix))
            {
                isoCode = valueInControlColumn.Replace(_countryTotalSuffix, "");
                return true;
            }

            return false;
        }

        private string GetStringValue(XL.Range row,int columnNumber)
        {
            object value = row.Cells[1, columnNumber].Value;
            if (value is string)
            {
                return (string)value;
            }
            return null;
        }

        private decimal? GetDecimalValue(XL.Range row, int columnNumber)
        {
            object value = row.Cells[1, columnNumber].Value;

            if (value is double)
            {
                return Convert.ToDecimal(value);
            }
            return null;
        }

        private int? GetIntValue(XL.Range row, int columnNumber)
        {
            object value = row.Cells[1, columnNumber].Value;

            if (value is int)
            {
                return (int)value;
            }
            return null;
        }

        private Location BuildLocation(XL.Range row, string ticker)
        {
            decimal? units = GetDecimalValue(row, _netPositionColumnNumber);
            string name = GetStringValue(row, _nameColumnNumber);
            int? tickerTypeId = GetIntValue(row, _tickerTypeColumnNumber);
            decimal? priceDivisor = GetDecimalValue(row, _priceDivisorColumnNumber);
            return new Location(row.Row, ticker, name, units ?? 0, tickerTypeId, null, null, priceDivisor ?? 1);            
        }

        public void WriteNAV(decimal nav)
        {
            _worksheet.Range[_fundNavLabel].Cells.Value = nav;
        }

        private void WriteValue(int rowNumber, int columnNumber, object value,bool? isBold)
        {
            var cell = _worksheet.Cells[rowNumber, columnNumber];
            cell.Value = value;
            if (isBold.HasValue)
            {
                cell.Font.Bold = isBold.Value;
            }
        }

        #region Formulas

        private static readonly int _firstRowOfData = 6;
        private static readonly int _bloombergMnemonicRow = 4;
        private static readonly string _fundCurrencyLabel = "FundCurrency";
        private static readonly string _fundNavLabel = "NAV";
        private static readonly string _finalTotalName = "Total_Total";
        private static readonly string _countryTotalSuffix = "_Total";



        private void WriteFormula(int rowNumber,int columnNumber, string formula, bool? isBold)
        {
            var cell = _worksheet.Cells[rowNumber, columnNumber];
            cell.Formula = formula;
            if (isBold.HasValue)
            {
                cell.Font.Bold = isBold.Value;
            }
        }

        private static readonly string _bloombergError = "\"#N/A N/A\"";

        private string GetSubtractFormula(int rowNumber, string column1, string column2)
        {
            string column1AC = $"{ column1 }{ rowNumber}";
            string column2AC = $"{ column2 }{ rowNumber}";
            return $"=if(or({column1AC}={_bloombergError},{column2AC}={_bloombergError}),0,  {column1AC} - {column2AC})";
        }


        private string GetMultiplyFormula(int rowNumber, string[] columns, string[] divideColumn)
        {
            string divideColumns = "";
            if (divideColumn != null && divideColumn.Length>0)
            {
                divideColumns = "/"+string.Join("/", divideColumn.Select(a => a + rowNumber));
            }
            return "="+string.Join("*",columns.Select(a=>a+rowNumber))+divideColumns;
        }

        private string GetDivideFormula(int rowNumber, string dividendColumn, string divisorColumn, bool displayedAsPercentage)
        {
            string multiplyBy100 = null;
            if (!displayedAsPercentage)
            {
                multiplyBy100 = "*100";
            }
            string divisor = $"{ divisorColumn }{ rowNumber}";
            return $"=if(or({divisor}=0,{divisor}={_bloombergError}),0,{dividendColumn}{rowNumber} / {divisor}{multiplyBy100})";
        }

        private string GetDivideByNavFormula(int rowNumber, string column, bool displayedAsPercentage)
        {
            string multiplyBy100 = null;
            if (!displayedAsPercentage)
            {
                multiplyBy100 = "*100";
            }
            return $"={column}{rowNumber} / {_fundNavLabel}{multiplyBy100}";
        }

        private string GetBloombergMnemonicFormula(int rowNumber,string column)
        {
            return GetBloombergMnemonicFormula(rowNumber, column, _tickerColumn);
        }

        private string GetQuoteFactorFormula(int rowNumber)
        {
            return $"=IF({_currencyColumn}{rowNumber} = {_fundCurrencyLabel},1,{GetBloombergMnemonicFormula(rowNumber, _quoteFactorColumn, _currencyTickerColumn).Replace("=", "")})";
        }

        private string GetFXRateFormula(int rowNumber)
        {
            return $"=IF({_currencyColumn}{rowNumber} = {_fundCurrencyLabel},1,{GetBloombergMnemonicFormula(rowNumber, _fxRateColumn, _currencyTickerColumn).Replace("=","")}*{_quoteFactorColumn}{rowNumber})";
        }

        private string GetBloombergMnemonicFormula(int rowNumber, string mnemonicColumn,string tickerColumn)
        {
            return $"=BDP({tickerColumn}{rowNumber},${mnemonicColumn}${_bloombergMnemonicRow})";
        }

        private string GetCurrencyTickerFormula(int rowNumber)
        {
            return $"=CONCATENATE({_fundCurrencyLabel},{_currencyColumn}{rowNumber}, \" Curncy\")";
        }

        private string GetPriceMultiplierFormula(int rowNumber)
        {
            return $"=IF(EXACT({_currencyColumn}{rowNumber},UPPER({_currencyColumn}{rowNumber})),1,0.01)/{_priceDivisorColumn}{rowNumber}";
        }

        private string GetWriteIfCorrectSignColumn(int rowNumber,bool isPositive, string columnToCheck)
        {
            string greaterThanLessThan = isPositive ? ">" : "<";
            return $"=IF({columnToCheck}{rowNumber}{greaterThanLessThan}0,{columnToCheck}{rowNumber},0)";
        }

        private string GetSumFormula(int firstRow, int lastRow, string column)
        {
            return $"= SUM({column}{firstRow}:{column}{lastRow})";
        }

        #endregion
    }
}
