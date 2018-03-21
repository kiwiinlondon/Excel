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
    public class WorksheetAccess
    {
       

        public WorksheetAccess(XL.Worksheet worksheet)
        {
         //   _workbook = workBook;
            _worksheet = worksheet;
        }
        private XL.Worksheet _worksheet;

        private static readonly string _controlColumn = "A";
        private static readonly int _controlColumnNumber = GetColumnNumber(_controlColumn);

        private static readonly string _instrumentMarketIdColumn = "B";
        private static readonly int _instrumentMarketIdColumnNumber = GetColumnNumber(_instrumentMarketIdColumn);

        private static readonly string _tickerColumn = "C";
        private static readonly int _tickerColumnNumber = GetColumnNumber(_tickerColumn);

        private static readonly string _currencyColumn = "D";
        private static readonly int _currencyColumnNumber = GetColumnNumber(_currencyColumn);

        private static readonly string _nameColumn = "E";
        private static readonly int _nameColumnNumber = GetColumnNumber(_nameColumn);

        private static readonly string _closePriceColumn = "F";
        private static readonly int _closePriceColumnNumber = GetColumnNumber(_closePriceColumn);

        private static readonly string _currentPriceColumn = "G";
        private static readonly int _currentPriceColumnNumber = GetColumnNumber(_currentPriceColumn);

        private static readonly string _priceChangeColumn = "H";
        private static readonly int _priceChangeColumnNumber = GetColumnNumber(_priceChangeColumn);

        private static readonly string _pricePercentageChangeColumn = "I";
        private static readonly int _pricePercentageChangeColumnNumber = GetColumnNumber(_pricePercentageChangeColumn);

        private static readonly string _netPositionColumn = "J";
        private static readonly int _netPositionColumnNumber = GetColumnNumber(_netPositionColumn);

        private static readonly string _currencyTickerColumn = "K";
        private static readonly int _currencyTickerColumnNumber = GetColumnNumber(_currencyTickerColumn);

        private static readonly string _quoteFactorColumn = "L";
        private static readonly int _quoteFactorColumnNumber = GetColumnNumber(_quoteFactorColumn);

        private static readonly string _fxRateColumn = "M";
        private static readonly int _fxRateColumnNumber = GetColumnNumber(_fxRateColumn);

        private static readonly string _pnlColumn = "N";
        private static readonly int _pnlColumnNumber = GetColumnNumber(_pnlColumn);

        private static readonly string _contributionBookColumn = "O";
        private static readonly int _contributionBookColumnNumber = GetColumnNumber(_contributionBookColumn);

        private static readonly string _contributionFundColumn = "P";
        private static readonly int _contributionFundColumnNumber = GetColumnNumber(_contributionFundColumn);

        private static readonly string _exposureColumn = "Q";
        private static readonly int _exposureColumnNumber = GetColumnNumber(_exposureColumn);

        private static readonly string _exposurePercentageBookColumn = "R";
        private static readonly int _exposurePercentageBookColumnNumber = GetColumnNumber(_exposurePercentageBookColumn);

        private static readonly string _exposurePercentageFundColumn = "S";
        private static readonly int _exposurePercentageFundColumnNumber = GetColumnNumber(_exposurePercentageFundColumn);

        private static readonly string _shortBookColumn = "T";
        private static readonly int _shortBookColumnNumber = GetColumnNumber(_shortBookColumn);

        private static readonly string _longBookColumn = "U";
        private static readonly int _longBookColumnNumber = GetColumnNumber(_longBookColumn);

        private static readonly string _priceMultiplierColumn = "V";
        private static readonly int _priceMultiplierColumnNumber = GetColumnNumber(_priceMultiplierColumn);

        private static readonly string _instrumentTypeColumn = "W";
        private static readonly int _instrumentTypeColumnNumber = GetColumnNumber(_instrumentTypeColumn);

        private static readonly string _priceDivisorColumn = "X";
        private static readonly int _priceDivisorColumnNumber = GetColumnNumber(_priceDivisorColumn);

        private static readonly string _shortBookWinnersColumn = "Y";
        private static readonly int _shortBookWinnersColumnNumber = GetColumnNumber(_shortBookWinnersColumn);

        private static readonly string _longBookWinnersColumn = "Z";
        private static readonly int _longBookWinnersColumnNumber = GetColumnNumber(_longBookWinnersColumn);

        private static readonly string _navColumn = "AA";
        private static readonly int _navColumnNumber = GetColumnNumber(_navColumn);

        private static readonly string _previousClosePriceColumn = "AB";
        private static readonly int _previousClosePriceColumnNumber = GetColumnNumber(_previousClosePriceColumn);

        private static readonly string _previousPriceChangeColumn = "AC";
        private static readonly int _previousPriceChangeColumnNumber = GetColumnNumber(_previousPriceChangeColumn);

        private static readonly string _previousPricePercentageChangeColumn = "AD";
        private static readonly int _previousPricePercentageChangeColumnNumber = GetColumnNumber(_previousPricePercentageChangeColumn);

        private static readonly string _previousNetPositionColumn = "AE";
        private static readonly int _previousNetPositionColumnNumber = GetColumnNumber(_previousNetPositionColumn);

        private static readonly string _previousFXRateColumn = "AF";
        private static readonly int _previousFXRateColumnNumber = GetColumnNumber(_previousFXRateColumn);

        private static readonly string _previousContributionBookColumn = "AG";
        private static readonly int _previousContributionBookColumnNumber = GetColumnNumber(_previousContributionBookColumn);


        private static readonly string _previousContributionFundColumn = "AH";
        private static readonly int _previousContributionFundColumnNumber = GetColumnNumber(_previousContributionFundColumn);

        private static readonly string _previousNavColumn = "AI";
        private static readonly int _previousNavColumnNumber = GetColumnNumber(_previousNavColumn);

        private static readonly string _lastColumn = _previousNavColumn;

        private static readonly string _firstColumn = _controlColumn;

        private static readonly int _firstRowOfData = 11;
        private static readonly int _bloombergMnemonicRow = 8;
        private static readonly string _previousReferenceDateLabel = $"${_currencyColumn}$1";
        private static readonly string _referenceDateLabel = $"${_nameColumn}$1";
        private static readonly string _totalSuffix = "#Total";
        private static readonly string _ignoreLabel = "#IGNORE#";

        
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


        


        public void AddPosition(Position previousPosition, Position position, GroupingEntity parent, Book book, Fund fund)
        {
            int rowToAddAt;
            if (previousPosition == null)
            {
                if (parent.Previous == null)
                {
                    rowToAddAt = _firstRowOfData;
                }
                else
                {
                    rowToAddAt = parent.Previous.TotalRow.Row+2;
                }
            }
            else
            {
                rowToAddAt = previousPosition.RowNumber+1;
            }
            position.Row = AddRow(rowToAddAt);

            WritePosition(position, book, fund, true);
        }

        public void WritePosition(Position position, Book book, Fund fund, bool updateFormulas)
        {
            WriteValue(position.Row, _instrumentMarketIdColumnNumber, position.Identifier.Id, null);
            WriteValue(position.Row, _tickerColumnNumber, position.Identifier.Code, null);
            
            WriteName(position, updateFormulas);
            WriteCurrency(position, updateFormulas);
            WriteClosePrice(position, updateFormulas);
            WriteCurrentPrice(position, updateFormulas);
            WritePreviousClosePrice(position, updateFormulas);

            WriteFormula(position.Row, _priceChangeColumnNumber, GetSubtractFormula(position.RowNumber, _currentPriceColumn, _closePriceColumn), null, updateFormulas);
            WriteFormula(position.Row, _pricePercentageChangeColumnNumber, GetDivideFormula(position.RowNumber, _priceChangeColumn, _closePriceColumn, false), null, updateFormulas);
            WriteValue(position.Row, _netPositionColumnNumber, position.NetPosition, false);
            WriteFormula(position.Row, _currencyTickerColumnNumber, GetCurrencyTickerFormula(position.RowNumber,fund.TotalRow), null, updateFormulas);
            WriteFormula(position.Row, _quoteFactorColumnNumber, GetQuoteFactorFormula(position.RowNumber, fund.TotalRow), null, updateFormulas);
            WriteFormula(position.Row, _fxRateColumnNumber, GetFXRateFormula(position.RowNumber, _fxRateColumn, fund.TotalRow), null, updateFormulas);
            WriteFormula(position.Row, _pnlColumnNumber, GetPNLFormula(position), null, updateFormulas);
            WriteFormula(position.Row, _contributionBookColumnNumber, GetDivideByNavFormula(position.RowNumber, _pnlColumn, true,book), null, updateFormulas);
            WriteFormula(position.Row, _contributionFundColumnNumber, GetDivideByNavFormula(position.RowNumber, _pnlColumn, true, fund), null, updateFormulas);

            WriteFormula(position.Row, _exposureColumnNumber, GetExposureFormula(position.InstrumentTypeId, position.RowNumber), null, updateFormulas);
            WriteFormula(position.Row, _exposurePercentageBookColumnNumber, GetDivideByNavFormula(position.RowNumber, _exposureColumn, false, book), null, updateFormulas);
            WriteFormula(position.Row, _exposurePercentageFundColumnNumber, GetDivideByNavFormula(position.RowNumber, _exposureColumn, false, fund), null, updateFormulas);

            WriteFormula(position.Row, _shortBookColumnNumber, GetWriteIfIsLongCorrectColumn(position.InstrumentTypeId, position.RowNumber, false), null, updateFormulas);
            WriteFormula(position.Row, _longBookColumnNumber, GetWriteIfIsLongCorrectColumn(position.InstrumentTypeId, position.RowNumber, true), null, updateFormulas);
            WriteFormula(position.Row, _priceMultiplierColumnNumber, GetPriceMultiplierFormula(position.RowNumber), null, updateFormulas);
            WriteValue(position.Row, _instrumentTypeColumnNumber, position.InstrumentTypeId, false);
            WriteValue(position.Row, _priceDivisorColumnNumber, position.PriceDivisor, false);
            WriteFormula(position.Row, _shortBookWinnersColumnNumber, GetWinnerColumn(position.RowNumber, false), null, updateFormulas);
            WriteFormula(position.Row, _longBookWinnersColumnNumber, GetWinnerColumn(position.RowNumber, true), null, updateFormulas);

            WriteFormula(position.Row, _previousPriceChangeColumnNumber, GetSubtractFormula(position.RowNumber, _closePriceColumn, _previousClosePriceColumn), null, updateFormulas);
            WriteFormula(position.Row, _previousPricePercentageChangeColumnNumber, GetDivideFormula(position.RowNumber, _previousPriceChangeColumn, _previousClosePriceColumn, false), null, updateFormulas);

            WriteValue(position.Row, _previousNetPositionColumnNumber, position.PreviousNetPosition, false);
            WriteFormula(position.Row, _previousFXRateColumnNumber, GetFXRateFormula(position.RowNumber, _previousFXRateColumn, fund.TotalRow), null, updateFormulas);
            WriteFormula(position.Row, _previousContributionBookColumnNumber, GetPreviousContribution(position, position.RowNumber,book), null, updateFormulas);
            WriteFormula(position.Row, _previousContributionFundColumnNumber, GetPreviousContribution(position, position.RowNumber, fund), null, updateFormulas);
        }

        private void WriteName(Position position, bool updateFormula)
        {
            if (position.InstrumentTypeId == InstrumentTypeIds.DoNotDelete)
            {
                WriteFormula(position.Row, _nameColumnNumber, GetBloombergMnemonicFormula(position.Row.Row, _nameColumn,_tickerColumn), null, updateFormula);
            }       
            else
            {
                WriteValue(position.Row, _nameColumnNumber, position.Name, null);
            }
        }

        private void WriteCurrency(Position position, bool updateFormula)
        {
            if (position.InstrumentTypeId == InstrumentTypeIds.FX || position.InstrumentTypeId == InstrumentTypeIds.PrivatePlacement)
            {
                WriteValue(position.Row, _currencyColumnNumber, position.Currency, null);
            }
            else
            {
                WriteFormula(position.Row, _currencyColumnNumber, GetBloombergMnemonicFormula(position.RowNumber, _currencyColumn), null, updateFormula);
            }
        }

        private void WriteClosePrice(Position position, bool updateFormula)
        {
            if (position.InstrumentTypeId == InstrumentTypeIds.FX || position.InstrumentTypeId == InstrumentTypeIds.PrivatePlacement)
            {
                WriteValue(position.Row, _closePriceColumnNumber, position.OdeyPreviousPrice, null);           
            }
            else
            {
                WriteFormula(position.Row, _closePriceColumnNumber, GetBloombergMnemonicFormula(position.RowNumber, _closePriceColumn), null, updateFormula);
            }
        }

        private void WriteCurrentPrice(Position position, bool updateFormula)
        {
            if (position.InstrumentTypeId == InstrumentTypeIds.PrivatePlacement)
            {
                WriteValue(position.Row, _currentPriceColumnNumber, position.OdeyCurrentPrice, null);
            }
            else
            {
                WriteFormula(position.Row, _currentPriceColumnNumber, GetBloombergMnemonicFormula(position.RowNumber, _currentPriceColumn), null, updateFormula);
            }
        }

        private void WritePreviousClosePrice(Position position, bool updateFormula)
        {
            if (position.InstrumentTypeId == InstrumentTypeIds.FX || position.InstrumentTypeId == InstrumentTypeIds.PrivatePlacement)
            {
                WriteValue(position.Row, _previousClosePriceColumnNumber, position.OdeyPreviousPreviousPrice, null);        
            }
            else
            {
                WriteFormula(position.Row, _previousClosePriceColumnNumber, GetBloombergMnemonicHistoryFormula(position.RowNumber, _tickerColumn, _previousClosePriceColumn), null, updateFormula);
            }
        }

        public void UpdateSums(GroupingEntity entity)
        {

            int firstRowNumber = _firstRowOfData;
            if (entity.Previous!=null)
            {
                firstRowNumber = entity.Previous.TotalRow.Row + 1;
            }
            int lastRowNumber = entity.TotalRow.Row - 1;
            WriteFormula(entity.TotalRow, _pnlColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _pnlColumn),true, true);
            WriteFormula(entity.TotalRow, _contributionBookColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _contributionBookColumn),false, true);
            WriteFormula(entity.TotalRow, _contributionFundColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _contributionFundColumn), false, true);
            WriteFormula(entity.TotalRow, _exposureColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _exposureColumn),true, true);
            WriteFormula(entity.TotalRow, _exposurePercentageBookColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _exposurePercentageBookColumn), true, true);
            WriteFormula(entity.TotalRow, _exposurePercentageFundColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _exposurePercentageFundColumn), true, true);
            WriteFormula(entity.TotalRow, _shortBookColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _shortBookColumn), true, true);
            WriteFormula(entity.TotalRow, _longBookColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _longBookColumn), true, true);
            WriteFormula(entity.TotalRow, _shortBookWinnersColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _shortBookWinnersColumn), true, true);
            WriteFormula(entity.TotalRow, _longBookWinnersColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _longBookWinnersColumn), true, true);
            WriteFormula(entity.TotalRow, _previousContributionBookColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _previousContributionBookColumn), true, true);
            WriteFormula(entity.TotalRow, _previousContributionFundColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _previousContributionFundColumn), true, true);

        }



        private XL.Range AddRow(int rowNumber)
        {
            _worksheet.Rows[rowNumber].Insert(XL.XlDirection.xlUp, XL.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            XL.Range insertedRow = GetRow(rowNumber);
            insertedRow.Font.Bold = false;
            insertedRow.RowHeight = 12;
            return insertedRow;
        }

        public void DeleteRows(int startRow,int endRow)
        {
            var range = _worksheet.get_Range($"{_firstColumn}{startRow}:{_lastColumn}{endRow}");
            DeleteRange(range);
        }

        public void DeleteRange(XL.Range range)
        {
            range.Delete();            
        }
        public void ChangeRowVisibilty(int firstRow, int lastRow,bool hidden)
        {
            _worksheet.get_Range($"{_firstColumn}{firstRow}:{_lastColumn}{lastRow}").EntireRow.Hidden = hidden;
        }

        

        private XL.Range GetRow(int rowNumber)
        {
            return _worksheet.get_Range($"{_firstColumn}{rowNumber}:{_lastColumn}{rowNumber}");
        }

        public void UpdateTotalsOnTotalRow(GroupingEntity groupingEntity)
        {
            int[] rowNumbers = groupingEntity.Children.Select(a => a.Value.RowNumber).ToArray();
            UpdateTotalOnTotalRow(groupingEntity, _pnlColumnNumber, _pnlColumn, rowNumbers,false, true);
            UpdateTotalOnTotalRow(groupingEntity, _contributionBookColumnNumber, _contributionBookColumn, rowNumbers,true, true);
            UpdateTotalOnTotalRow(groupingEntity, _contributionFundColumnNumber, _contributionFundColumn, rowNumbers, false, true);
            UpdateTotalOnTotalRow(groupingEntity, _exposureColumnNumber, _exposureColumn, rowNumbers, false, true);
            UpdateTotalOnTotalRow(groupingEntity, _exposurePercentageBookColumnNumber, _exposurePercentageBookColumn, rowNumbers,true, true);
            UpdateTotalOnTotalRow(groupingEntity, _exposurePercentageFundColumnNumber, _exposurePercentageFundColumn, rowNumbers, false, true);
            UpdateTotalOnTotalRow(groupingEntity, _shortBookColumnNumber, _shortBookColumn, rowNumbers,true, true);
            UpdateTotalOnTotalRow(groupingEntity, _longBookColumnNumber, _longBookColumn, rowNumbers, true, true);
            UpdateTotalOnTotalRow(groupingEntity, _shortBookWinnersColumnNumber, _shortBookWinnersColumn, rowNumbers, true, true);
            UpdateTotalOnTotalRow(groupingEntity, _longBookWinnersColumnNumber, _longBookWinnersColumn, rowNumbers, true, true);
            UpdateTotalOnTotalRow(groupingEntity, _previousContributionBookColumnNumber, _previousContributionBookColumn, rowNumbers, true, true);
            UpdateTotalOnTotalRow(groupingEntity, _previousContributionFundColumnNumber, _previousContributionFundColumn, rowNumbers,false, true);
        }

        public void UpdateNavs(GroupingEntity groupingEntity)
        {
            WriteValue(groupingEntity.TotalRow, _navColumnNumber, groupingEntity.Nav,false);
            WriteValue(groupingEntity.TotalRow, _previousNavColumnNumber, groupingEntity.PreviousNav, false);
            if (groupingEntity is Fund)
            {
                WriteValue(groupingEntity.TotalRow, _currencyColumnNumber, ((Fund)groupingEntity).Currency, false);
            }
        }

        private void UpdateTotalOnTotalRow(GroupingEntity groupingEntity, int columnNumber,string column,int[] rowNumbers, bool isPercentageOfBookNavColumn, bool updateTotal)
        {
            if (updateTotal)
            {
                string formula;
                if (groupingEntity is Fund && isPercentageOfBookNavColumn)
                {
                    formula = null;
                }
                else
                {
                    formula = "=" + string.Join("+", rowNumbers.Select(a => column + a));
                }
                var cell = groupingEntity.TotalRow.Cells[1, columnNumber];
                cell.Formula = formula;
            }
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

        public void AddTotalRow(GroupingEntity group)
        {
            int addAtRowNumber;
            if (group.Previous == null)
            {
                addAtRowNumber = _firstRowOfData+1;
            }
            else
            {
                addAtRowNumber = group.Previous.TotalRow.Row+1;
            }

            group.TotalRow = AddRow(addAtRowNumber);

            AddRow(group.TotalRow.Row);//Gap Between sections

            group.ControlString = GetControlString(group.Parent.ControlString, group.Identifier.Code);
            WriteValue(group.TotalRow, _controlColumnNumber, group.ControlString, false);
            WriteValue(group.TotalRow, _tickerColumnNumber, group.Name,true);
            group.TotalRow.Borders[XL.XlBordersIndex.xlEdgeBottom].LineStyle = XL.XlLineStyle.xlContinuous;
            group.TotalRow.Borders[XL.XlBordersIndex.xlEdgeTop].LineStyle = XL.XlLineStyle.xlContinuous;
            
        }



        private string CreateTotalLabel(string fund, string book, string assetClass, string country)
        {
            return string.Join("#", new[] { fund, book, assetClass, country }) + _totalSuffix;                
        }

        public void AddFundRange(Fund fund)
        {
            int firstRowOfData = _firstRowOfData;
            if (fund.Previous != null)
            {
                var previousRange = ((Fund)fund.Previous).Range;
                firstRowOfData = previousRange.Row + previousRange.Rows.Count;
            }

            string fundTotalLabel = CreateTotalLabel(fund.Name, null, null, null);
            int? lastRow = FindRow(fundTotalLabel, _controlColumn);

            if (!lastRow.HasValue)
            {
                throw new ApplicationException($"No Total Row exists for fund {fund.Name}");
            }
            fund.Range = _worksheet.get_Range($"{_firstColumn}{firstRowOfData}:{_lastColumn}{lastRow.Value}");
        }

        public List<ExistingGroupDTO> GetExisting(Fund fund)
        {
            List<ExistingGroupDTO> existingGroups = new List<ExistingGroupDTO>();
            List<ExistingPositionDTO> positions = new List<ExistingPositionDTO>();
            foreach (XL.Range row in fund.Range.Rows)
            {
                string valueInTickerColumn = GetStringValue(row, _tickerColumnNumber);
                int? valueInInstrumentMarketColumn = GetIntValue(row, _instrumentMarketIdColumnNumber);

                string controlString = GetStringValue(row, _controlColumnNumber);

                if ((!valueInInstrumentMarketColumn.HasValue && string.IsNullOrWhiteSpace(valueInTickerColumn) && string.IsNullOrWhiteSpace(controlString)) || controlString == _ignoreLabel)
                {
                    continue;
                }
                if (RowIsTotal(controlString))
                {
                    string name = GetStringValue(row, _nameColumnNumber);
                    existingGroups.Add(new ExistingGroupDTO(controlString, name, row, positions));
                    positions = new List<ExistingPositionDTO>();
                }
                else
                {

                    var position = new ExistingPositionDTO(valueInInstrumentMarketColumn, valueInTickerColumn, row);
                    positions.Add(position);
                }
            }
            return existingGroups;
        }


        public string GetNameFromRow(XL.Range row)
        {
            return GetStringValue(row, _nameColumnNumber);
        }

        public Position BuildPosition(ExistingPositionDTO existingPosition)
        {            
            var row = existingPosition.Row;
            int? instrumentTypeIdAsInt = GetIntValue(row, _instrumentTypeColumnNumber);

            string name = GetNameFromRow(row);

            InstrumentTypeIds instrumentTypeId = (InstrumentTypeIds)(instrumentTypeIdAsInt ?? 0);
            string currency = null;
            bool invertPNL = false;
            if (instrumentTypeId == InstrumentTypeIds.FX)
            {
                currency = GetStringValue(row, _currencyColumnNumber);
                invertPNL = !existingPosition.Identifier.Code.StartsWith(currency);
            }
            decimal? priceDivisor = GetDecimalValue(row, _priceDivisorColumnNumber);
            return new Position(existingPosition.Identifier, name, priceDivisor ?? 1, instrumentTypeId, invertPNL) { Currency = currency};
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

        public void WriteDates(DateTime previousReferenceDate, DateTime referenceDate)
        {
            _worksheet.Range[_previousReferenceDateLabel].Cells.Value = previousReferenceDate;
            _worksheet.Range[_referenceDateLabel].Cells.Value = referenceDate;
        }

        

        private void WriteValue(XL.Range row, int columnNumber, object value, bool? isBold)
        {
            var cell = row.Cells[1, columnNumber];
            cell.Value = value;
            if (isBold.HasValue)
            {
                cell.Font.Bold = isBold.Value;
            }

        }

        #region Formulas







        private void WriteFormula(XL.Range row,int columnNumber, string formula, bool? isBold, bool updateFormula)
        {
            if (updateFormula)
            {
                var cell = row.Cells[1, columnNumber];

                cell.Formula = formula;
                if (isBold.HasValue)
                {
                    cell.Font.Bold = isBold.Value;
                }
            }
        }

        private static readonly string _bloombergError = "\"#N/A N/A\"";

        private string GetSubtractFormula(int rowNumber, string column1, string column2)
        {
            string column1AC = $"{ column1 }{ rowNumber}";
            string column2AC = $"{ column2 }{ rowNumber}";
            return $"=if(or({column1AC}={_bloombergError},{column2AC}={_bloombergError}),0,  {column1AC} - {column2AC})";
        }

        

        private string GetExposureFormula(InstrumentTypeIds instrumentTypeId, int rowNumber)
        {
            string[] columns;
            string[] divideColumn = new string[] { _fxRateColumn };
            if (instrumentTypeId == InstrumentTypeIds.FX)
            {
                columns = new string[] { _netPositionColumn };
            }
            else
            {
                columns = new string[] { _currentPriceColumn, _netPositionColumn, _priceMultiplierColumn };
            }
            string formula = GetMultiplyFormula(rowNumber, columns, divideColumn,false);
            if (instrumentTypeId == InstrumentTypeIds.FX)
            {
                formula = formula.Replace("=", "=Abs(") + ")";
            }
            return formula;
        }

        private string GetPNLFormula(Position position)
        {
            
            if (position.InstrumentTypeId == InstrumentTypeIds.FX)
            {
                return GetMultiplyFormula(position.RowNumber, new string[] { _priceChangeColumn, _netPositionColumn }, new string[] { _fxRateColumn,_currentPriceColumn }, position.InvertPNL);
            }
            else
            {
                return GetMultiplyFormula(position.RowNumber, new string[] { _priceChangeColumn, _netPositionColumn, _priceMultiplierColumn }, new string[] { _fxRateColumn },false);
            }
            
        }


        private string GetMultiplyFormula(int rowNumber, string[] columns, string[] divideColumn, bool invert)
        {
            string divideColumns = "";
            if (divideColumn != null && divideColumn.Length>0)
            {
                divideColumns = "/"+string.Join("/", divideColumn.Select(a => a + rowNumber));
            }
            if (invert)
            {
                int i = 0;
            }

            return "=" +string.Join("*",columns.Select(a=>a+rowNumber))+ divideColumns + (invert ? "*-1" : "");
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

        private string GetPreviousContribution(Position position,int rowNumber,GroupingEntity groupingEntity)
        {
            if (groupingEntity == null)
            {
                return null;
            }
            else
            {
                string pnlFormula;
                if (position.InstrumentTypeId == InstrumentTypeIds.FX)
                {
                    pnlFormula = GetMultiplyFormula(rowNumber, new string[] { _previousPriceChangeColumn, _previousNetPositionColumn }, new string[] { _previousFXRateColumn, _previousClosePriceColumn },position.InvertPNL);
                }
                else
                {
                    pnlFormula = GetMultiplyFormula(rowNumber, new string[] { _previousPriceChangeColumn, _previousNetPositionColumn, _priceMultiplierColumn }, new string[] { _previousFXRateColumn },false);
                }
                pnlFormula = pnlFormula.Replace("=", "");
                return $"={pnlFormula} / {_previousNavColumn}{groupingEntity.TotalRow.Row}";
            }
        }

        private string GetDivideByNavFormula(int rowNumber, string column, bool displayedAsPercentage,GroupingEntity groupingEntity)
        {
            if (groupingEntity== null)
            {
                return null;
            }
            string multiplyBy100 = null;
            if (!displayedAsPercentage)
            {
                multiplyBy100 = "*100";
            }
            return $"={column}{rowNumber} / {_navColumn}{groupingEntity.TotalRow.Row}{multiplyBy100}";
        }

        private string GetBloombergMnemonicFormula(int rowNumber,string column)
        {
            return GetBloombergMnemonicFormula(rowNumber, column, _tickerColumn);
        }

        private string GetBloombergMnemonicHistoryFormula(int rowNumber,string tickerColumn, string column)
        {
            return $"=BDH({tickerColumn}{rowNumber},${column}${_bloombergMnemonicRow},{_previousReferenceDateLabel},{_previousReferenceDateLabel})";
        }

        private string GetQuoteFactorFormula(int rowNumber,XL.Range fundTotalRow)
        {
            return $"=IF({_currencyColumn}{rowNumber} = {_currencyColumn}{fundTotalRow.Row},1,{GetBloombergMnemonicFormula(rowNumber, _quoteFactorColumn, _currencyTickerColumn).Replace("=", "")})";
        }

        private string GetFXRateFormula(int rowNumber, string fxRateColumn, XL.Range fundTotalRow)
        {
            return $"=IF({_currencyColumn}{rowNumber} = {_currencyColumn}{fundTotalRow.Row},1,{GetBloombergMnemonicFormula(rowNumber, fxRateColumn, _currencyTickerColumn).Replace("=","")}*{_quoteFactorColumn}{rowNumber})";
        }

        private string GetBloombergMnemonicFormula(int rowNumber, string mnemonicColumn,string tickerColumn)
        {
            return $"=BDP({tickerColumn}{rowNumber},${mnemonicColumn}${_bloombergMnemonicRow})";
        }

        private string GetCurrencyTickerFormula(int rowNumber,XL.Range fundTotalRow)
        {
            return $"=CONCATENATE({_currencyColumn}{fundTotalRow.Row},{_currencyColumn}{rowNumber}, \" Curncy\")";
        }

        private string GetPriceMultiplierFormula(int rowNumber)
        {
            return $"=IF(EXACT({_currencyColumn}{rowNumber},UPPER({_currencyColumn}{rowNumber})),1,0.01)/{_priceDivisorColumn}{rowNumber}";
        }

        private string GetWriteIfIsLongCorrectColumn(InstrumentTypeIds instrumentTypeId,int rowNumber, bool isLong)
        {
            if (instrumentTypeId == InstrumentTypeIds.FX)
            {
                return null;
            }
            return GetWriteIfStatement(rowNumber, GetExposureIsLongTest(rowNumber, isLong), _exposurePercentageBookColumn);
        }

        private string GetWinnerColumn(int rowNumber, bool isLong)
        {
            string exposureTest = GetExposureIsLongTest(rowNumber, isLong);
            string winnerTest = GetIsGreaterThanZeroTest(rowNumber, true, _contributionBookColumn);
            string test = $"AND({exposureTest},{winnerTest})";
            return GetWriteIfStatement(rowNumber, test, _contributionBookColumn);
        }

        private string GetExposureIsLongTest(int rowNumber, bool isLong)
        {
            return GetIsGreaterThanZeroTest(rowNumber,isLong, _exposurePercentageBookColumn);
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

        public List<string> GetBulkTickers()
        {
            List<string> tickers = new List<string>();
            XL.Range usedRange = _worksheet.UsedRange;
            foreach (XL.Range row in usedRange.Rows)
            {
                var ticker = GetStringValue(row, 1);
                if (ticker != null)
                {
                    tickers.Add(ticker);
                }
            }
            return tickers;
        }
    }
}
