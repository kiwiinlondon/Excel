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

        private static readonly string _shortWinnersColumn = "V";
        private static readonly int _shortWinnersColumnNumber = GetColumnNumber(_shortWinnersColumn);

        private static readonly string _longWinnersColumn = "W";
        private static readonly int _longWinnersColumnNumber = GetColumnNumber(_longWinnersColumn);

        private static readonly string _previousClosePriceColumn = "X";
        private static readonly int _previousClosePriceColumnNumber = GetColumnNumber(_previousClosePriceColumn);

        private static readonly string _previousPriceChangeColumn = "Y";
        private static readonly int _previousPriceChangeColumnNumber = GetColumnNumber(_previousPriceChangeColumn);


        private static readonly string _previousPricePercentageChangeColumn = "Z";
        private static readonly int _previousPricePercentageChangeColumnNumber = GetColumnNumber(_previousPricePercentageChangeColumn);
        private static readonly string _previousNetPositionColumn = "AA";
        private static readonly int _previousNetPositionColumnNumber = GetColumnNumber(_previousNetPositionColumn);

        private static readonly string _previousFXRateColumn = "AB";
        private static readonly int _previousFXRateColumnNumber = GetColumnNumber(_previousFXRateColumn);

        private static readonly string _previousContributionColumn = "AC";
        private static readonly int _previousContributionColumnNumber = GetColumnNumber(_previousContributionColumn);

        private static readonly string _lastColumn = _previousContributionColumn;

        private static readonly string _firstColumn = _controlColumn;

        private int? FindRow(string toFind, string column)
        {
            XL.Range tickers = _worksheet.get_Range($"${column}:${column}");
            var currentFind = tickers.Find(toFind, System.Reflection.Missing.Value,
                XL.XlFindLookIn.xlFormulas, XL.XlLookAt.xlPart,//use xl forulas as column is hidden
                XL.XlSearchOrder.xlByRows, XL.XlSearchDirection.xlNext, false,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            tickers[1, 1] = "Geoff";
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
            if (letter.Length == 1)
            {
                return GetColumnNumber(letter[0]);
            }
            else if (letter.Length == 2)
            {
                return 26 + GetColumnNumber(letter[1]);
            }

            throw new ApplicationException($"Dont know how to convert strings that are not two char long ({letter})");

        }
        private static int GetColumnNumber(char letter)
        {
            return char.ToUpper(letter) - 'A' + 1;
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
                        WriteValue(location.Row, _netPositionColumnNumber, location.NetPosition, null);
                    }
                    WriteValue(location.Row, _previousNetPositionColumnNumber, location.PreviousNetPosition, null);
                }
            }
        }

        public void AddTickerRow(Location location, int rowNumber)
        {
            location.RowNumber = rowNumber;
            location.Row = AddRow(location.RowNumber.Value);
            WriteRow(location);
        }

        private void WriteRow(Location location)
        {
            WriteValue(location.Row, _tickerColumnNumber, location.Ticker, null);

            if (location.TickerTypeId.HasValue)
            {
                WritePrivatePlacement(location);
            }
            else
            {
                WriteNormal(location);
            }

            WriteFormula(location.Row, _priceChangeColumnNumber, GetSubtractFormula(location.RowNumber.Value, _currentPriceColumn, _closePriceColumn), null);
            WriteFormula(location.Row, _pricePercentageChangeColumnNumber, GetDivideFormula(location.RowNumber.Value, _priceChangeColumn, _closePriceColumn, false), null);
            WriteValue(location.Row, _netPositionColumnNumber, location.NetPosition, false);
            WriteFormula(location.Row, _currencyTickerColumnNumber, GetCurrencyTickerFormula(location.RowNumber.Value), null);
            WriteFormula(location.Row, _quoteFactorColumnNumber, GetQuoteFactorFormula(location.RowNumber.Value), null);
            WriteFormula(location.Row, _fxRateColumnNumber, GetFXRateFormula(location.RowNumber.Value, _fxRateColumn), null);
            WriteFormula(location.Row, _pnlColumnNumber, GetMultiplyFormula(location.RowNumber.Value, new string[] { _priceChangeColumn, _netPositionColumn, _priceMultiplierColumn }, new string[] { _fxRateColumn }), null);
            WriteFormula(location.Row, _contributionColumnNumber, GetDivideByNavFormula(location.RowNumber.Value, _pnlColumn, true), null);

            WriteFormula(location.Row, _exposureColumnNumber, GetMultiplyFormula(location.RowNumber.Value, new string[] { _currentPriceColumn, _netPositionColumn, _priceMultiplierColumn },new string[] { _fxRateColumn}), null);
            WriteFormula(location.Row, _exposurePercentageColumnNumber, GetDivideByNavFormula(location.RowNumber.Value, _exposureColumn, false), null);

            WriteFormula(location.Row, _shortColumnNumber, GetWriteIfIsLongCorrectColumn(location.RowNumber.Value, false), null);
            WriteFormula(location.Row, _longColumnNumber, GetWriteIfIsLongCorrectColumn(location.RowNumber.Value, true), null);
            WriteFormula(location.Row, _priceMultiplierColumnNumber, GetPriceMultiplierFormula(location.RowNumber.Value), null);
            WriteValue(location.Row, _tickerTypeColumnNumber, location.TickerTypeId, false);
            WriteValue(location.Row, _priceDivisorColumnNumber, location.PriceDivisor, false);
            WriteFormula(location.Row, _shortWinnersColumnNumber, GetWinnerColumn(location.RowNumber.Value, false), null);
            WriteFormula(location.Row, _longWinnersColumnNumber, GetWinnerColumn(location.RowNumber.Value, true), null);

            WriteFormula(location.Row, _previousPriceChangeColumnNumber, GetSubtractFormula(location.RowNumber.Value, _closePriceColumn, _previousClosePriceColumn), null);
            WriteFormula(location.Row, _previousPricePercentageChangeColumnNumber, GetDivideFormula(location.RowNumber.Value, _previousPriceChangeColumn, _previousClosePriceColumn, false), null);

            WriteValue(location.Row, _previousNetPositionColumnNumber, location.PreviousNetPosition, false);
            WriteFormula(location.Row, _previousFXRateColumnNumber, GetFXRateFormula(location.RowNumber.Value, _previousFXRateColumn), null);
            WriteFormula(location.Row, _previousContributionColumnNumber, GetPreviousContribution(location.RowNumber.Value), null);

        }



        private void WriteNormal(Location location)
        {
            WriteFormula(location.Row, _currencyColumnNumber, GetBloombergMnemonicFormula(location.RowNumber.Value, _currencyColumn), null);
            WriteFormula(location.Row, _nameColumnNumber, GetBloombergMnemonicFormula(location.RowNumber.Value, _nameColumn), null);
            WriteFormula(location.Row, _closePriceColumnNumber, GetBloombergMnemonicFormula(location.RowNumber.Value, _closePriceColumn), null);
            WriteFormula(location.Row, _currentPriceColumnNumber, GetBloombergMnemonicFormula(location.RowNumber.Value, _currentPriceColumn), null);
            WriteFormula(location.Row, _previousClosePriceColumnNumber, GetBloombergMnemonicHistoryFormula(location.RowNumber.Value, _tickerColumn, _previousClosePriceColumn),null);            
        }

        private void WritePrivatePlacement(Location location)
        {
            WriteValue(location.Row, _currencyColumnNumber, location.Currency, null);
            WriteValue(location.Row, _nameColumnNumber, location.Name, null);
            WriteValue(location.Row, _closePriceColumnNumber, location.OdeyPreviousPrice, null);
            WriteValue(location.Row, _currentPriceColumnNumber, location.OdeyCurrentPrice, null);
            WriteValue(location.Row, _previousClosePriceColumnNumber, location.OdeyPreviousPreviousPrice, null);
  
        }

        public void UpdateSums(CountryLocation country)
        {
            WriteFormula(country.TotalRow, _pnlColumnNumber, GetSumFormula(country.FirstRowNumber.Value,country.TotalRowNumber.Value-1,_pnlColumn),true);
            WriteFormula(country.TotalRow, _contributionColumnNumber, GetSumFormula(country.FirstRowNumber.Value, country.TotalRowNumber.Value - 1, _contributionColumn),false);
            WriteFormula(country.TotalRow, _exposureColumnNumber, GetSumFormula(country.FirstRowNumber.Value, country.TotalRowNumber.Value - 1, _exposureColumn),true);
            WriteFormula(country.TotalRow, _exposurePercentageColumnNumber, GetSumFormula(country.FirstRowNumber.Value, country.TotalRowNumber.Value - 1, _exposurePercentageColumn), true);
            WriteFormula(country.TotalRow, _shortColumnNumber, GetSumFormula(country.FirstRowNumber.Value, country.TotalRowNumber.Value - 1, _shortColumn), true);
            WriteFormula(country.TotalRow, _longColumnNumber, GetSumFormula(country.FirstRowNumber.Value, country.TotalRowNumber.Value - 1, _longColumn), true);
            WriteFormula(country.TotalRow, _shortWinnersColumnNumber, GetSumFormula(country.FirstRowNumber.Value, country.TotalRowNumber.Value - 1, _shortWinnersColumn), true);
            WriteFormula(country.TotalRow, _longWinnersColumnNumber, GetSumFormula(country.FirstRowNumber.Value, country.TotalRowNumber.Value - 1, _longWinnersColumn), true);
            WriteFormula(country.TotalRow, _previousContributionColumnNumber, GetSumFormula(country.FirstRowNumber.Value, country.TotalRowNumber.Value - 1, _previousContributionColumn), true);

        }
              
        private XL.Range AddRow(int row)
        {
            LastRow++;
            _worksheet.Rows[row].Insert(XL.XlDirection.xlUp, XL.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            _worksheet.Rows[row].Font.Bold = false;
            return _worksheet.get_Range($"{_firstColumn}{row}:{_lastColumn}{row}");
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
            UpdateTotalOnTotalRow(_shortWinnersColumnNumber, _shortWinnersColumn, rowNumberToAdd);
            UpdateTotalOnTotalRow(_longWinnersColumnNumber, _longWinnersColumn, rowNumberToAdd);
            UpdateTotalOnTotalRow(_previousContributionColumnNumber, _previousContributionColumn, rowNumberToAdd);
        }

        private void UpdateTotalOnTotalRow(int columnNumber,string column,int rowNumberToAdd)
        {
            var cell = _worksheet.Cells[LastRow, columnNumber];
            cell.Formula = $"{cell.Formula}+{column}{rowNumberToAdd}";
        }

        public void AddCountryTotalRow(int lastTotalRow, CountryLocation country)
        {
            AddRow(lastTotalRow);
            country.TotalRow = AddRow(lastTotalRow);
            WriteValue(country.TotalRow, _controlColumnNumber, GetCountryTotalString(country.IsoCode),false);
            WriteValue(country.TotalRow, _nameColumnNumber, country.Name,true);
            country.TotalRowNumber = lastTotalRow;
            country.FirstRowNumber = lastTotalRow;
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
                string valueInNameColumn = GetStringValue(row, _nameColumnNumber);
                if (string.IsNullOrWhiteSpace(valueInTickerColumn) && string.IsNullOrWhiteSpace(valueInControlColumn) && string.IsNullOrWhiteSpace(valueInNameColumn))
                {
                    continue;
                }

                if (countryLocation == null)
                {
                    countryLocation = new CountryLocation();
                    countryLocation.FirstRowNumber = i;
                }

                string isoCode;
                if (RowIsCountryTotal(valueInControlColumn, out isoCode))
                {
                    countryLocation.IsoCode = isoCode;
                    countryLocation.Name = valueInNameColumn; 
                    countryLocation.TotalRowNumber = i;
                    countryLocation.TotalRow = row;
                    countryLocations.Add(isoCode, countryLocation);
                    countryLocation = null;
                }
                else
                {
                    var location = BuildLocation(row, valueInTickerColumn,valueInNameColumn);
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
            if (value != null)
            {
                return value.ToString();
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
            else if (value is double)
            {
                return Convert.ToInt32(value);
            }
            return null;
        }

        private Location BuildLocation(XL.Range row, string ticker, string name)
        {
            decimal? previousNetPosition = GetDecimalValue(row, _previousNetPositionColumnNumber);
            decimal? netPosition = GetDecimalValue(row, _netPositionColumnNumber);
            int? tickerTypeId = GetIntValue(row, _tickerTypeColumnNumber);
            decimal? priceDivisor = GetDecimalValue(row, _priceDivisorColumnNumber);
            return new Location(row.Row, ticker, name, previousNetPosition??0, netPosition ?? 0, tickerTypeId,null,null, null, null, priceDivisor ?? 1,row);            
        }

        public void WriteNAVs(decimal previousNav,decimal nav)
        {
            _worksheet.Range[_previousFundNavLabel].Cells.Value = previousNav;
            _worksheet.Range[_fundNavLabel].Cells.Value = nav;            
        }

        public void DisableCalculations()
        {
            _worksheet.Application.Calculation = XL.XlCalculation.xlCalculationManual;
            _worksheet.Application.ScreenUpdating = false;
            _worksheet.Application.EnableEvents = false;
        }

        public void EnableCalculations()
        {

            _worksheet.Application.Calculation = XL.XlCalculation.xlCalculationAutomatic;
            _worksheet.Application.ScreenUpdating = true;
            _worksheet.Application.EnableEvents = true;
        }

        private void WriteValue(XL.Range row, int columnNumber, object value,bool? isBold)
        {
            var cell = row.Cells[1, columnNumber];
            cell.Value = value;
            if (isBold.HasValue)
            {
                cell.Font.Bold = isBold.Value;
            }
        }

        #region Formulas

        private static readonly int _firstRowOfData = 10;
        private static readonly int _bloombergMnemonicRow = 7;
        private static readonly string _fundCurrencyLabel = "FundCurrency";
        private static readonly string _fundNavLabel = "NAV";
        private static readonly string _previousFundNavLabel = "PreviousNAV";
        private static readonly string _previousReferenceDateLabel = "$C$1";
        private static readonly string _finalTotalName = "Total_Total";
        private static readonly string _countryTotalSuffix = "_Total";



        private void WriteFormula(XL.Range row,int columnNumber, string formula, bool? isBold)
        {
            var cell = row.Cells[1, columnNumber];

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

        private string GetPreviousContribution(int rowNumber)
        {            
            string pnlFormula = GetMultiplyFormula(rowNumber, new string[] { _previousPriceChangeColumn, _previousNetPositionColumn, _priceMultiplierColumn },new string[] { _previousFXRateColumn }).Replace("=", "");
            return $"={pnlFormula} / {_previousFundNavLabel}";
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

        private string GetBloombergMnemonicHistoryFormula(int rowNumber,string tickerColumn, string column)
        {
            return $"=BDH({tickerColumn}{rowNumber},${column}${_bloombergMnemonicRow},{_previousReferenceDateLabel},{_previousReferenceDateLabel})";
        }

        private string GetQuoteFactorFormula(int rowNumber)
        {
            return $"=IF({_currencyColumn}{rowNumber} = {_fundCurrencyLabel},1,{GetBloombergMnemonicFormula(rowNumber, _quoteFactorColumn, _currencyTickerColumn).Replace("=", "")})";
        }

        private string GetFXRateFormula(int rowNumber,string fxRateColumn)
        {
            return $"=IF({_currencyColumn}{rowNumber} = {_fundCurrencyLabel},1,{GetBloombergMnemonicFormula(rowNumber, fxRateColumn, _currencyTickerColumn).Replace("=","")}*{_quoteFactorColumn}{rowNumber})";
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

        private string GetWriteIfIsLongCorrectColumn(int rowNumber,bool isLong)
        {
            return GetWriteIfStatement(rowNumber, GetExposureIsLongTest(rowNumber, isLong), _exposurePercentageColumn);
        }

        private string GetWinnerColumn(int rowNumber, bool isLong)
        {
            string exposureTest = GetExposureIsLongTest(rowNumber, isLong);
            string winnerTest = GetIsGreaterThanZeroTest(rowNumber, true, _contributionColumn);
            string test = $"AND({exposureTest},{winnerTest})";
            return GetWriteIfStatement(rowNumber, test, _contributionColumn);
        }

        private string GetExposureIsLongTest(int rowNumber, bool isLong)
        {
            return GetIsGreaterThanZeroTest(rowNumber,isLong, _exposurePercentageColumn);
        }

        private string GetIsGreaterThanZeroTest(int rowNumber, bool isPositive, string columnToCheck)
        {
            string greaterThanLessThan = isPositive ? ">" : "<";
            return $"{columnToCheck}{rowNumber}{greaterThanLessThan}0";
        }
        private string GetWriteIfStatement(int rowNumber, string test, string columnToReturn)
        {
            return $"=IF({test},{columnToReturn}{rowNumber},0)";
        }

        private string GetSumFormula(int firstRow, int lastRow, string column)
        {
            return $"= SUM({column}{firstRow}:{column}{lastRow})";
        }

        #endregion
    }
}
