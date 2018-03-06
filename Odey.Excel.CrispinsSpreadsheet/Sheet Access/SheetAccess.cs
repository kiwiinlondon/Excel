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

        private static readonly string _contributionColumn = "O";
        private static readonly int _contributionColumnNumber = GetColumnNumber(_contributionColumn);

        private static readonly string _exposureColumn = "P";
        private static readonly int _exposureColumnNumber = GetColumnNumber(_exposureColumn);

        private static readonly string _exposurePercentageColumn = "Q";
        private static readonly int _exposurePercentageColumnNumber = GetColumnNumber(_exposurePercentageColumn);

        private static readonly string _shortColumn = "R";
        private static readonly int _shortColumnNumber = GetColumnNumber(_shortColumn);

        private static readonly string _longColumn = "S";
        private static readonly int _longColumnNumber = GetColumnNumber(_longColumn);

        private static readonly string _priceMultiplierColumn = "T";
        private static readonly int _priceMultiplierColumnNumber = GetColumnNumber(_priceMultiplierColumn);

        private static readonly string _instrumentTypeColumn = "U";
        private static readonly int _instrumentTypeColumnNumber = GetColumnNumber(_instrumentTypeColumn);

        private static readonly string _priceDivisorColumn = "V";
        private static readonly int _priceDivisorColumnNumber = GetColumnNumber(_priceDivisorColumn);

        private static readonly string _shortWinnersColumn = "W";
        private static readonly int _shortWinnersColumnNumber = GetColumnNumber(_shortWinnersColumn);

        private static readonly string _longWinnersColumn = "X";
        private static readonly int _longWinnersColumnNumber = GetColumnNumber(_longWinnersColumn);

        private static readonly string _navColumn = "Y";
        private static readonly int _navColumnNumber = GetColumnNumber(_navColumn);

        private static readonly string _previousClosePriceColumn = "Z";
        private static readonly int _previousClosePriceColumnNumber = GetColumnNumber(_previousClosePriceColumn);

        private static readonly string _previousPriceChangeColumn = "AA";
        private static readonly int _previousPriceChangeColumnNumber = GetColumnNumber(_previousPriceChangeColumn);

        private static readonly string _previousPricePercentageChangeColumn = "AB";
        private static readonly int _previousPricePercentageChangeColumnNumber = GetColumnNumber(_previousPricePercentageChangeColumn);

        private static readonly string _previousNetPositionColumn = "AC";
        private static readonly int _previousNetPositionColumnNumber = GetColumnNumber(_previousNetPositionColumn);

        private static readonly string _previousFXRateColumn = "AD";
        private static readonly int _previousFXRateColumnNumber = GetColumnNumber(_previousFXRateColumn);

        private static readonly string _previousContributionColumn = "AE";
        private static readonly int _previousContributionColumnNumber = GetColumnNumber(_previousContributionColumn);

        private static readonly string _previousNavColumn = "AF";
        private static readonly int _previousNavColumnNumber = GetColumnNumber(_previousNavColumn);

        private static readonly string _lastColumn = _previousNavColumn;

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


        public void UpdatePosition(bool forceRefresh, Position position, Book book, Fund fund)
        {
            if (forceRefresh)
            {
                WritePosition(position, book,fund);
            }
            else
            {                
                if (position.InstrumentTypeId == InstrumentTypeIds.PrivatePlacement)
                {
                    WritePrivatePlacementPosition(position);
                }
                else if (position.InstrumentTypeId == InstrumentTypeIds.FX)
                {
                    WriteFXPosition(position);
                }
                else
                {
                    WriteValue(position.Row, _instrumentMarketIdColumnNumber, position.Identifier.Code, null);
                    WriteValue(position.Row, _nameColumnNumber, position.Name, null);
                    WriteValue(position.Row, _netPositionColumnNumber, position.NetPosition, null);
                    WriteValue(position.Row, _previousNetPositionColumnNumber, position.PreviousNetPosition, null);
                }
            }
        }

        public Position AddInstrument(InstrumentDTO instrument)
        {
            string fundTotalLabel = CreateTotalLabel("OEI", null, null, null);
            int? fundTotal = FindRow(fundTotalLabel, _controlColumn);

            string bookTotalLabel = CreateTotalLabel("OEI", "BK-OEI", null, null);
            int? bookTotal = FindRow(bookTotalLabel, _controlColumn);




            return new Position(instrument.Identifier, instrument.Name, instrument.PriceDivisor,instrument.InstrumentTypeId, null);
        }

        public void AddPosition(Position previousPosition, Position position, GroupingEntity parent, Book book, Fund fund)
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
            position.Row = AddRow(previous.Row,1);

            WritePosition(position, book, fund);
        }

        private void WritePosition(Position position, Book book, Fund fund)
        {
            WriteValue(position.Row, _instrumentMarketIdColumnNumber, position.Identifier.Id, null);
            WriteValue(position.Row, _tickerColumnNumber, position.Identifier.Code, null);
            WriteValue(position.Row, _nameColumnNumber, position.Name, null);
            if (position.InstrumentTypeId == InstrumentTypeIds.FX)
            {
                WriteFXPosition(position);
            }
            else if (position.InstrumentTypeId == InstrumentTypeIds.PrivatePlacement)
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
            WriteFormula(position.Row, _currencyTickerColumnNumber, GetCurrencyTickerFormula(position.RowNumber,fund.TotalRow), null);
            WriteFormula(position.Row, _quoteFactorColumnNumber, GetQuoteFactorFormula(position.RowNumber, fund.TotalRow), null);
            WriteFormula(position.Row, _fxRateColumnNumber, GetFXRateFormula(position.RowNumber, _fxRateColumn, fund.TotalRow), null);
            WriteFormula(position.Row, _pnlColumnNumber, GetPNLFormula(position.InstrumentTypeId, position.RowNumber), null);
            WriteFormula(position.Row, _contributionColumnNumber, GetDivideByNavFormula(position.RowNumber, _pnlColumn, true,fund.TotalRow), null);

            WriteFormula(position.Row, _exposureColumnNumber, GetExposureFormula(position.InstrumentTypeId, position.RowNumber), null);
            WriteFormula(position.Row, _exposurePercentageColumnNumber, GetDivideByNavFormula(position.RowNumber, _exposureColumn, false, fund.TotalRow), null);

            WriteFormula(position.Row, _shortColumnNumber, GetWriteIfIsLongCorrectColumn(position.RowNumber, false), null);
            WriteFormula(position.Row, _longColumnNumber, GetWriteIfIsLongCorrectColumn(position.RowNumber, true), null);
            WriteFormula(position.Row, _priceMultiplierColumnNumber, GetPriceMultiplierFormula(position.RowNumber), null);
            WriteValue(position.Row, _instrumentTypeColumnNumber, position.InstrumentTypeId, false);
            WriteValue(position.Row, _priceDivisorColumnNumber, position.PriceDivisor, false);
            WriteFormula(position.Row, _shortWinnersColumnNumber, GetWinnerColumn(position.RowNumber, false), null);
            WriteFormula(position.Row, _longWinnersColumnNumber, GetWinnerColumn(position.RowNumber, true), null);

            WriteFormula(position.Row, _previousPriceChangeColumnNumber, GetSubtractFormula(position.RowNumber, _closePriceColumn, _previousClosePriceColumn), null);
            WriteFormula(position.Row, _previousPricePercentageChangeColumnNumber, GetDivideFormula(position.RowNumber, _previousPriceChangeColumn, _previousClosePriceColumn, false), null);

            WriteValue(position.Row, _previousNetPositionColumnNumber, position.PreviousNetPosition, false);
            WriteFormula(position.Row, _previousFXRateColumnNumber, GetFXRateFormula(position.RowNumber, _previousFXRateColumn, fund.TotalRow), null);
            WriteFormula(position.Row, _previousContributionColumnNumber, GetPreviousContribution(position.InstrumentTypeId, position.RowNumber,fund.TotalRow), null);
        }

        private void WriteFXPosition(Position position)
        {
            WriteValue(position.Row, _currencyColumnNumber, position.Currency, null);
            WriteValue(position.Row, _closePriceColumnNumber, position.OdeyPreviousPrice, null);
            WriteFormula(position.Row, _currentPriceColumnNumber, GetBloombergMnemonicFormula(position.RowNumber, _currentPriceColumn), null);
            WriteValue(position.Row, _previousClosePriceColumnNumber, position.OdeyPreviousPreviousPrice, null);
        }

        private void WriteNormalPosition(Position position)
        {
            WriteFormula(position.Row, _currencyColumnNumber, GetBloombergMnemonicFormula(position.RowNumber, _currencyColumn), null);
            //WriteFormula(position.Row, _nameColumnNumber, GetBloombergMnemonicFormula(position.RowNumber, _nameColumn), null);
            WriteFormula(position.Row, _closePriceColumnNumber, GetBloombergMnemonicFormula(position.RowNumber, _closePriceColumn), null);
            WriteFormula(position.Row, _currentPriceColumnNumber, GetBloombergMnemonicFormula(position.RowNumber, _currentPriceColumn), null);
            WriteFormula(position.Row, _previousClosePriceColumnNumber, GetBloombergMnemonicHistoryFormula(position.RowNumber, _tickerColumn, _previousClosePriceColumn),null);            
        }

        private void WritePrivatePlacementPosition(Position position)
        {
            WriteValue(position.Row, _currencyColumnNumber, position.Currency, null);
          //  WriteValue(position.Row, _nameColumnNumber, position.Name, null);
            WriteValue(position.Row, _closePriceColumnNumber, position.OdeyPreviousPrice, null);
            WriteValue(position.Row, _currentPriceColumnNumber, position.OdeyCurrentPrice, null);
            WriteValue(position.Row, _previousClosePriceColumnNumber, position.OdeyPreviousPreviousPrice, null);  
        }

        public void UpdateSums(GroupingEntity entity)
        {

            int firstRowNumber = _firstRowOfData;
            if (entity.Previous!=null)
            {
                firstRowNumber = entity.Previous.TotalRow.Row + 1;
            }
            int lastRowNumber = entity.TotalRow.Row - 1;
            WriteFormula(entity.TotalRow, _pnlColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _pnlColumn),true);
            WriteFormula(entity.TotalRow, _contributionColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _contributionColumn),false);
            WriteFormula(entity.TotalRow, _exposureColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _exposureColumn),true);
            WriteFormula(entity.TotalRow, _exposurePercentageColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _exposurePercentageColumn), true);
            WriteFormula(entity.TotalRow, _shortColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _shortColumn), true);
            WriteFormula(entity.TotalRow, _longColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _longColumn), true);
            WriteFormula(entity.TotalRow, _shortWinnersColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _shortWinnersColumn), true);
            WriteFormula(entity.TotalRow, _longWinnersColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _longWinnersColumn), true);
            WriteFormula(entity.TotalRow, _previousContributionColumnNumber, GetSumFormula(firstRowNumber, lastRowNumber, _previousContributionColumn), true);

        }



        private XL.Range AddRow(int rowNumber,int numberOfRowsToAdd)
        {
            int lastRowNumberAdded = rowNumber-1;
            for (int i = 0; i < numberOfRowsToAdd; i++)
            {
                lastRowNumberAdded++;
                _worksheet.Rows[lastRowNumberAdded].Insert(XL.XlDirection.xlUp, XL.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                
            }
            XL.Range insertedRow = GetRow(lastRowNumberAdded);
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

        public void UpdateNavs(GroupingEntity groupingEntity)
        {

            WriteValue(groupingEntity.TotalRow, _navColumnNumber, groupingEntity.Nav,false);
            WriteValue(groupingEntity.TotalRow, _previousNavColumnNumber, groupingEntity.PreviousNav, false);
            if (groupingEntity is Fund)
            {
                WriteValue(groupingEntity.TotalRow, _tickerColumnNumber, ((Fund)groupingEntity).Currency, false);
            }
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

            group.TotalRow = AddRow(addAtRowNumber,2);

            group.ControlString = GetControlString(group.Parent.ControlString, group.Identifier.Code);
            WriteValue(group.TotalRow, _controlColumnNumber, group.ControlString, false);
            WriteValue(group.TotalRow, _nameColumnNumber, group.Name,true);
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
                firstRowOfData = fund.Previous.RowNumber + 1;
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
            //GroupingEntity previousGroupingEntity = fund.Previous;
            foreach (XL.Range row in fund.Range.Rows)
            {
                string valueInTickerColumn = GetStringValue(row, _tickerColumnNumber);
                int? valueInInstrumentMarketColumn = GetIntValue(row, _instrumentMarketIdColumnNumber);

                string controlString = GetStringValue(row, _controlColumnNumber);

                if ((!valueInInstrumentMarketColumn.HasValue && string.IsNullOrWhiteSpace(valueInTickerColumn) && string.IsNullOrWhiteSpace(controlString)) || controlString == _ignoreLabel)
                {
                    continue;
                }
               
                string valueInNameColumn = GetStringValue(row, _nameColumnNumber);

                if (RowIsTotal(controlString))
                {
                    existingGroups.Add(new ExistingGroupDTO(controlString, valueInNameColumn, row, positions));
                    //var entity = AddToParent(controlString, positions,valueInNameColumn, row, fund);
                    //entity.Previous = previousGroupingEntity;
                    //previousGroupingEntity = entity;
                    positions = new List<ExistingPositionDTO>();
                }
                else
                {
                    var position = new ExistingPositionDTO(valueInInstrumentMarketColumn, valueInTickerColumn, valueInNameColumn, row);
                    //if (positions == null)
                   // {
                   //     positions = new List<ExistingPositionDTO>();
                   // }
                    positions.Add(position);
                }
            }
            return existingGroups;
        }


        //private GroupingEntity AddToParent(string controlString, List<ExistingPositionDTO> existingPositions, string name, XL.Range row, Fund fund)
        //{
        //    var values = controlString.Split('#');
        //    string fundCode = values[0];
        //    string bookCode = values[1];
        //    string assetClassCode = values[2];
        //    string countryCode = values[3];
                      
        //    GroupingEntity entity = fund;

        //    if (!string.IsNullOrWhiteSpace(bookCode))
        //    {
        //        entity = GetEntity(entity, bookCode, GroupingEntityTypes.Book);
        //    }

        //    if (!string.IsNullOrWhiteSpace(assetClassCode))
        //    {
        //        entity = GetEntity(entity, assetClassCode, GroupingEntityTypes.AssetClass);
        //    }

        //    if (!string.IsNullOrWhiteSpace(countryCode))
        //    {
        //        entity = GetEntity(entity, countryCode, GroupingEntityTypes.Country);
        //    }

        //    entity.TotalRow = row;
        //    entity.Name = name;
        //    entity.ControlString = controlString;
        //    if (existingPositions != null && existingPositions.Count > 0)
        //    {
        //        AddPositions(entity, existingPositions);
        //    }
        //    return entity;
        //}

        //private void AddPositions(GroupingEntity entity,List<ExistingPositionDTO> existingPositions)
        //{
        //    foreach (var existingPosition in existingPositions)
        //    {
        //        Position position;
        //        if (entity.Children.ContainsKey(existingPosition.Identifier))
        //        {
        //            position = (Position)entity.Children[existingPosition.Identifier];
        //            if (position.InstrumentTypeId == InstrumentTypeIds.FX)
        //            {
        //                position.Name = existingPosition.Name;
        //            }
        //            position.Row = existingPosition.Row;
        //        }
        //        else
        //        {
        //            position = BuildPosition(existingPosition);
        //            entity.Children.Add(existingPosition.Identifier, position);
        //        }
        //    }
        //}        

        public Position BuildPosition(ExistingPositionDTO existingPosition)
        {
            
            var row = existingPosition.Row;
            int? instrumentTypeIdAsInt = GetIntValue(row, _instrumentTypeColumnNumber);

            InstrumentTypeIds instrumentTypeId = (InstrumentTypeIds)(instrumentTypeIdAsInt ?? 0);
            string currency = null;
            if (instrumentTypeId == InstrumentTypeIds.FX)
            {
                currency = GetStringValue(row, _currencyColumnNumber);
            }
            decimal? priceDivisor = GetDecimalValue(row, _priceDivisorColumnNumber);
            return new Position(existingPosition.Identifier, existingPosition.Name, priceDivisor ?? 1, instrumentTypeId, row) { Currency = currency };
        }


        //private GroupingEntity GetEntity(GroupingEntity parent, string code,GroupingEntityTypes entityType)
        //{
        //    Identifier identifier = new Identifier(null,code);
        //    GroupingEntity entity;
        //    if (parent.Children.ContainsKey(identifier))
        //    {
        //        entity = (GroupingEntity)parent.Children[identifier];
        //    }
        //    else
        //    {
        //        switch (entityType)
        //        {
        //            case GroupingEntityTypes.Book:
        //                throw new ApplicationException($"Book should have already been setup");
        //            case GroupingEntityTypes.AssetClass:
        //                throw new ApplicationException($"Asset Class should have already been setup");
        //            case GroupingEntityTypes.Country:
        //                entity = new Country(code);
        //                break;
        //            default:
        //                throw new ApplicationException($"Unknown Entity Type {entityType}");
        //        }
        //        parent.Children.Add(identifier, entity);

        //    }
        //    return entity;
        //}

        



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
        private static readonly string _previousReferenceDateLabel = $"${_currencyColumn}$1";
        private static readonly string _referenceDateLabel = $"${_nameColumn}$1";
        private static readonly string _totalSuffix = "#Total";
        private static readonly string _ignoreLabel = "#IGNORE#";





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
            string formula = GetMultiplyFormula(rowNumber, columns, divideColumn);
            if (instrumentTypeId == InstrumentTypeIds.FX)
            {
                formula = formula.Replace("=", "=Abs(") + ")";
            }
            return formula;
        }

        private string GetPNLFormula(InstrumentTypeIds instrumentTypeId, int rowNumber)
        {
            
            if (instrumentTypeId == InstrumentTypeIds.FX)
            {
                return GetMultiplyFormula(rowNumber, new string[] { _priceChangeColumn, _netPositionColumn }, new string[] { _fxRateColumn,_currentPriceColumn });
            }
            else
            {
                return GetMultiplyFormula(rowNumber, new string[] { _priceChangeColumn, _netPositionColumn, _priceMultiplierColumn }, new string[] { _fxRateColumn });
            }
            
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

        private string GetPreviousContribution(InstrumentTypeIds tickerTypeId,int rowNumber, XL.Range navRow)
        {
            string pnlFormula;
            if (tickerTypeId == InstrumentTypeIds.FX)
            {
                pnlFormula = GetMultiplyFormula(rowNumber, new string[] { _previousPriceChangeColumn, _previousNetPositionColumn}, new string[] { _previousFXRateColumn, _previousClosePriceColumn });
            }
            else
            {
                pnlFormula = GetMultiplyFormula(rowNumber, new string[] { _previousPriceChangeColumn, _previousNetPositionColumn, _priceMultiplierColumn }, new string[] { _previousFXRateColumn });
            }
            pnlFormula = pnlFormula.Replace("=", "");
            return $"={pnlFormula} / {_previousNavColumn}{navRow.Row}";
        }

        private string GetDivideByNavFormula(int rowNumber, string column, bool displayedAsPercentage,XL.Range navRow)
        {
            string multiplyBy100 = null;
            if (!displayedAsPercentage)
            {
                multiplyBy100 = "*100";
            }
            return $"={column}{rowNumber} / {_navColumn}{navRow.Row}{multiplyBy100}";
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
            return $"=IF({_currencyColumn}{rowNumber} = {_tickerColumn}{fundTotalRow.Row},1,{GetBloombergMnemonicFormula(rowNumber, _quoteFactorColumn, _currencyTickerColumn).Replace("=", "")})";
        }

        private string GetFXRateFormula(int rowNumber, string fxRateColumn, XL.Range fundTotalRow)
        {
            return $"=IF({_currencyColumn}{rowNumber} = {_tickerColumn}{fundTotalRow.Row},1,{GetBloombergMnemonicFormula(rowNumber, fxRateColumn, _currencyTickerColumn).Replace("=","")}*{_quoteFactorColumn}{rowNumber})";
        }

        private string GetBloombergMnemonicFormula(int rowNumber, string mnemonicColumn,string tickerColumn)
        {
            return $"=BDP({tickerColumn}{rowNumber},${mnemonicColumn}${_bloombergMnemonicRow})";
        }

        private string GetCurrencyTickerFormula(int rowNumber,XL.Range fundTotalRow)
        {
            return $"=CONCATENATE({_tickerColumn}{fundTotalRow.Row},{_currencyColumn}{rowNumber}, \" Curncy\")";
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
