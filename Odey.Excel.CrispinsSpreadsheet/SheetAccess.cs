﻿using Excel = Microsoft.Office.Interop.Excel;
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
            XL.Range controls = _worksheet.get_Range($"${column}:${column}");
            var found =  controls.Find(toFind, System.Reflection.Missing.Value,
                XL.XlFindLookIn.xlFormulas, XL.XlLookAt.xlPart,//use xl forulas as column is hidden
                XL.XlSearchOrder.xlByRows, XL.XlSearchDirection.xlNext, false,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            if (found!=null)
            {
                return found.Row;
            }
            return null;
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


        public void UpdatePosition(bool forceRefresh, Position position)
        {
            if (forceRefresh)
            {
                WritePosition(position);
            }
            else
            {
                if (position.TickerTypeId.HasValue)
                {
                    WritePrivatePlacementPosition(position);
                }
                else
                {

                    WriteValue(position.Row, _netPositionColumnNumber, position.NetPosition, null);
                    WriteValue(position.Row, _previousNetPositionColumnNumber, position.PreviousNetPosition, null);
                }
            }
        }

        public void AddPosition(Position previousPosition, Position position, GroupingEntity parent)
        {
            XL.Range previous;
            if (previousPosition == null)
            {
                previous = parent.TotalRow;
            }
            else
            {
                previous = previousPosition.Row;
            }
            position.Row = AddRow(previous);

            WritePosition(position);
        }

        private void WritePosition(Position position)
        {
            WriteValue(position.Row, _tickerColumnNumber, position.Ticker, null);

            if (position.TickerTypeId.HasValue)
            {
                WritePrivatePlacementPosition(position);
            }
            else
            {
                WriteNormalPosition(position);
            }

            WriteFormula(position.Row, _priceChangeColumnNumber, GetSubtractFormula(position.RowNumber, _currentPriceColumn, _closePriceColumn), null);
            WriteFormula(position.Row, _pricePercentageChangeColumnNumber, GetDivideFormula(position.RowNumber, _priceChangeColumn, _closePriceColumn, false), null);
            WriteValue(position.Row, _netPositionColumnNumber, position.NetPosition, false);
            WriteFormula(position.Row, _currencyTickerColumnNumber, GetCurrencyTickerFormula(position.RowNumber), null);
            WriteFormula(position.Row, _quoteFactorColumnNumber, GetQuoteFactorFormula(position.RowNumber), null);
            WriteFormula(position.Row, _fxRateColumnNumber, GetFXRateFormula(position.RowNumber, _fxRateColumn), null);
            WriteFormula(position.Row, _pnlColumnNumber, GetMultiplyFormula(position.RowNumber, new string[] { _priceChangeColumn, _netPositionColumn, _priceMultiplierColumn }, new string[] { _fxRateColumn }), null);
            WriteFormula(position.Row, _contributionColumnNumber, GetDivideByNavFormula(position.RowNumber, _pnlColumn, true), null);

            WriteFormula(position.Row, _exposureColumnNumber, GetMultiplyFormula(position.RowNumber, new string[] { _currentPriceColumn, _netPositionColumn, _priceMultiplierColumn },new string[] { _fxRateColumn}), null);
            WriteFormula(position.Row, _exposurePercentageColumnNumber, GetDivideByNavFormula(position.RowNumber, _exposureColumn, false), null);

            WriteFormula(position.Row, _shortColumnNumber, GetWriteIfIsLongCorrectColumn(position.RowNumber, false), null);
            WriteFormula(position.Row, _longColumnNumber, GetWriteIfIsLongCorrectColumn(position.RowNumber, true), null);
            WriteFormula(position.Row, _priceMultiplierColumnNumber, GetPriceMultiplierFormula(position.RowNumber), null);
            WriteValue(position.Row, _tickerTypeColumnNumber, position.TickerTypeId, false);
            WriteValue(position.Row, _priceDivisorColumnNumber, position.PriceDivisor, false);
            WriteFormula(position.Row, _shortWinnersColumnNumber, GetWinnerColumn(position.RowNumber, false), null);
            WriteFormula(position.Row, _longWinnersColumnNumber, GetWinnerColumn(position.RowNumber, true), null);

            WriteFormula(position.Row, _previousPriceChangeColumnNumber, GetSubtractFormula(position.RowNumber, _closePriceColumn, _previousClosePriceColumn), null);
            WriteFormula(position.Row, _previousPricePercentageChangeColumnNumber, GetDivideFormula(position.RowNumber, _previousPriceChangeColumn, _previousClosePriceColumn, false), null);

            WriteValue(position.Row, _previousNetPositionColumnNumber, position.PreviousNetPosition, false);
            WriteFormula(position.Row, _previousFXRateColumnNumber, GetFXRateFormula(position.RowNumber, _previousFXRateColumn), null);
            WriteFormula(position.Row, _previousContributionColumnNumber, GetPreviousContribution(position.RowNumber), null);

        }



        private void WriteNormalPosition(Position position)
        {
            WriteFormula(position.Row, _currencyColumnNumber, GetBloombergMnemonicFormula(position.RowNumber, _currencyColumn), null);
            WriteFormula(position.Row, _nameColumnNumber, GetBloombergMnemonicFormula(position.RowNumber, _nameColumn), null);
            WriteFormula(position.Row, _closePriceColumnNumber, GetBloombergMnemonicFormula(position.RowNumber, _closePriceColumn), null);
            WriteFormula(position.Row, _currentPriceColumnNumber, GetBloombergMnemonicFormula(position.RowNumber, _currentPriceColumn), null);
            WriteFormula(position.Row, _previousClosePriceColumnNumber, GetBloombergMnemonicHistoryFormula(position.RowNumber, _tickerColumn, _previousClosePriceColumn),null);            
        }

        private void WritePrivatePlacementPosition(Position position)
        {
            WriteValue(position.Row, _currencyColumnNumber, position.Currency, null);
            WriteValue(position.Row, _nameColumnNumber, position.Name, null);
            WriteValue(position.Row, _closePriceColumnNumber, position.OdeyPreviousPrice, null);
            WriteValue(position.Row, _currentPriceColumnNumber, position.OdeyCurrentPrice, null);
            WriteValue(position.Row, _previousClosePriceColumnNumber, position.OdeyPreviousPreviousPrice, null);
  
        }

        public void UpdateSums(GroupingEntity country,Position first)
        {
            country.FirstRow = first.Row;
            WriteFormula(country.TotalRow, _pnlColumnNumber, GetSumFormula(country.FirstRow.Row,country.TotalRow.Row-1,_pnlColumn),true);
            WriteFormula(country.TotalRow, _contributionColumnNumber, GetSumFormula(country.FirstRow.Row, country.TotalRow.Row - 1, _contributionColumn),false);
            WriteFormula(country.TotalRow, _exposureColumnNumber, GetSumFormula(country.FirstRow.Row, country.TotalRow.Row - 1, _exposureColumn),true);
            WriteFormula(country.TotalRow, _exposurePercentageColumnNumber, GetSumFormula(country.FirstRow.Row, country.TotalRow.Row - 1, _exposurePercentageColumn), true);
            WriteFormula(country.TotalRow, _shortColumnNumber, GetSumFormula(country.FirstRow.Row, country.TotalRow.Row - 1, _shortColumn), true);
            WriteFormula(country.TotalRow, _longColumnNumber, GetSumFormula(country.FirstRow.Row, country.TotalRow.Row - 1, _longColumn), true);
            WriteFormula(country.TotalRow, _shortWinnersColumnNumber, GetSumFormula(country.FirstRow.Row, country.TotalRow.Row - 1, _shortWinnersColumn), true);
            WriteFormula(country.TotalRow, _longWinnersColumnNumber, GetSumFormula(country.FirstRow.Row, country.TotalRow.Row - 1, _longWinnersColumn), true);
            WriteFormula(country.TotalRow, _previousContributionColumnNumber, GetSumFormula(country.FirstRow.Row, country.TotalRow.Row - 1, _previousContributionColumn), true);

        }
              
        private XL.Range AddRow(XL.Range row)
        {
            int rowNumber = row.Row;
            _worksheet.Rows[rowNumber].Insert(XL.XlDirection.xlUp, XL.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            XL.Range insertedRow = GetRow(rowNumber);
            insertedRow.Font.Bold = false;
            return insertedRow;
        }

        private XL.Range GetRow(int rowNumber)
        {
            return _worksheet.get_Range($"{_firstColumn}{rowNumber}:{_lastColumn}{rowNumber}");
        }





        public void UpdateTotalsOnTotalRow(GroupingEntity groupingEntity)
        {
            int[] rowNumbers = groupingEntity.Children.Select(a => a.Value.RowNumber).ToArray();
            UpdateTotalOnTotalRow(groupingEntity.TotalRow, _pnlColumnNumber, _pnlColumn, rowNumbers);
            UpdateTotalOnTotalRow(groupingEntity.TotalRow, _contributionColumnNumber, _contributionColumn, rowNumbers);
            UpdateTotalOnTotalRow(groupingEntity.TotalRow, _exposureColumnNumber, _exposureColumn, rowNumbers);
            UpdateTotalOnTotalRow(groupingEntity.TotalRow, _exposurePercentageColumnNumber, _exposurePercentageColumn, rowNumbers);
            UpdateTotalOnTotalRow(groupingEntity.TotalRow, _shortColumnNumber, _shortColumn, rowNumbers);
            UpdateTotalOnTotalRow(groupingEntity.TotalRow, _longColumnNumber, _longColumn, rowNumbers);
            UpdateTotalOnTotalRow(groupingEntity.TotalRow, _shortWinnersColumnNumber, _shortWinnersColumn, rowNumbers);
            UpdateTotalOnTotalRow(groupingEntity.TotalRow, _longWinnersColumnNumber, _longWinnersColumn, rowNumbers);
            UpdateTotalOnTotalRow(groupingEntity.TotalRow, _previousContributionColumnNumber, _previousContributionColumn, rowNumbers);
        }

        private void UpdateTotalOnTotalRow(XL.Range totalRow, int columnNumber,string column,int[] rowNumbers)
        {            
            var cell = totalRow.Cells[1, columnNumber];
            cell.Formula = "="+string.Join("+", rowNumbers.Select(a=>column+a));
        }

        private string GetControlString(string parentControlString, string codeToAdd)
        {
            var values = parentControlString.Split('#');
            for (int i = 0; i < values.Length; i++)
            {
                string value = values[i];
                if (string.IsNullOrEmpty(value))
                {
                    values[i] = codeToAdd;
                    break;
                }
            }
            return string.Join("#", values);
        }

        public void AddTotalRow(GroupingEntity previousGroup, GroupingEntity group, GroupingEntity parentGroup)
        {
            XL.Range rowToAddPriorTo;
            if (previousGroup == null)
            {
                rowToAddPriorTo = parentGroup.TotalRow;
            }
            else
            {
                rowToAddPriorTo = previousGroup.FirstRow;
            }

            rowToAddPriorTo = AddRow(rowToAddPriorTo);
            group.TotalRow = AddRow(rowToAddPriorTo);





            group.ControlString = GetControlString(parentGroup.ControlString, group.Code);
            WriteValue(group.TotalRow, _controlColumnNumber, group.ControlString, false);
            WriteValue(group.TotalRow, _nameColumnNumber, group.Name,true);
            group.TotalRow.Borders[XL.XlBordersIndex.xlEdgeBottom].LineStyle = XL.XlLineStyle.xlContinuous;
            group.TotalRow.Borders[XL.XlBordersIndex.xlEdgeTop].LineStyle = XL.XlLineStyle.xlContinuous;
            
        }



        private string CreateTotalLabel(string fund, string book, string assetClass, string country)
        {
            return string.Join("#", new[] { fund, book, assetClass, country }) + _totalSuffix;                
        }

        public Fund GetFund(string fundName)
        {
            string fundTotalLabel = CreateTotalLabel(fundName, null, null, null);
            int? lastRow = FindRow(fundTotalLabel, _controlColumn);

            if (!lastRow.HasValue)
            {
                throw new ApplicationException("No Total Row exists");
            }

            XL.Range fundRange = _worksheet.get_Range($"{_firstColumn}{_firstRowOfData}:{_lastColumn}{lastRow.Value}");
            
            Dictionary<string, Country> countrypositions = new Dictionary<string, Country>();
            Fund fund = new Fund(fundName);
            List<Position> positions = null;
            foreach (XL.Range row in fundRange.Rows)
            {
                string valueInTickerColumn = GetStringValue(row, _tickerColumnNumber);
                string controlString = GetStringValue(row, _controlColumnNumber);
                string valueInNameColumn = GetStringValue(row, _nameColumnNumber);
                if (string.IsNullOrWhiteSpace(valueInTickerColumn) && string.IsNullOrWhiteSpace(controlString) && string.IsNullOrWhiteSpace(valueInNameColumn))
                {
                    continue;
                }

                if (RowIsTotal(controlString))
                {
                    AddToParent(controlString, positions,valueInNameColumn, row, fund);
                    positions = null;
                }
                else
                {
                    var position = BuildPosition(row, valueInTickerColumn, valueInNameColumn);
                    if (positions == null)
                    {
                        positions = new List<Position>();
                    }
                    positions.Add(position);
                }                                         
            }
            return fund;
        }


        private void AddToParent(string controlString, List<Position> positions, string name, XL.Range row, Fund fund)
        {
            var values = controlString.Split('#');
            string fundCode = values[0];
            string bookCode = values[1];
            string assetClassCode = values[2];
            string countryCode = values[3];
                      
            GroupingEntity entity = fund;

            if (!string.IsNullOrWhiteSpace(bookCode))
            {
                entity = GetEntity(entity, bookCode, GroupingEntityTypes.Book);
            }

            if (!string.IsNullOrWhiteSpace(assetClassCode))
            {
                entity = GetEntity(entity, assetClassCode, GroupingEntityTypes.AssetClass);
            }

            if (!string.IsNullOrWhiteSpace(countryCode))
            {
                entity = GetEntity(entity, countryCode, GroupingEntityTypes.Country);
            }

            entity.TotalRow = row;
            entity.Name = name;
            entity.ControlString = controlString;
            if (positions != null && positions.Count > 0)
            {
                entity.FirstRow = positions[0].Row;
                entity.Children = positions.ToDictionary(a => a.Ticker, a => (IChildEntity)a);
            }
        }

        private GroupingEntity GetEntity(GroupingEntity parent, string code,GroupingEntityTypes entityType)
        {
            GroupingEntity entity;
            if (parent.Children.ContainsKey(code))
            {
                entity = (GroupingEntity)parent.Children[code];
            }
            else
            {
                switch (entityType)
                {
                    case GroupingEntityTypes.Book:
                        entity = new Book(code);
                        break;
                    case GroupingEntityTypes.AssetClass:
                        entity = new AssetClass(code);
                        break;
                    case GroupingEntityTypes.Country:
                        entity = new Country(code);
                        break;
                    default:
                        throw new ApplicationException($"Unknown Entity Type {entityType}");
                }
                parent.Children.Add(code,entity);

            }
            return entity;
        }

        



        private bool RowIsTotal(string valueInControlColumn)
        {
            return valueInControlColumn != null && valueInControlColumn.EndsWith(_totalSuffix);
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

        private Position BuildPosition(XL.Range row, string ticker, string name)
        {
            int? tickerTypeId = GetIntValue(row, _tickerTypeColumnNumber);
            decimal? priceDivisor = GetDecimalValue(row, _priceDivisorColumnNumber);
            return new Position( ticker, name, priceDivisor ?? 1, tickerTypeId, row);            
        }

        public void WriteNAVs(decimal previousNav,decimal nav)
        {
            _worksheet.Range[_previousFundNavLabel].Cells.Value = previousNav;
            _worksheet.Range[_fundNavLabel].Cells.Value = nav;            
        }

        public void WriteDates(DateTime previousReferenceDate, DateTime referenceDate)
        {
            _worksheet.Range[_previousReferenceDateLabel].Cells.Value = previousReferenceDate;
            _worksheet.Range[_referenceDateLabel].Cells.Value = referenceDate;
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
        private static readonly string _referenceDateLabel = "$D$1";
        private static readonly string _totalSuffix = "#Total";



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
