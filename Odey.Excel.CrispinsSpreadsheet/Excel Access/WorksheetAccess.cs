using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XL=Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Drawing;
using Odey.Framework.Keeley.Entities.Enums;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public abstract class WorksheetAccess
    {


        public WorksheetAccess(XL.Worksheet worksheet)
        {
            _worksheet = worksheet;
        }
        private XL.Worksheet _worksheet;

        public string Name => _worksheet.Name;

        private static readonly string _controlColumn = "A";
        private static readonly int _controlColumnNumber = GetColumnNumber(_controlColumn).Value;

        private static readonly string _instrumentMarketIdColumn = "B";
        private static readonly int _instrumentMarketIdColumnNumber = GetColumnNumber(_instrumentMarketIdColumn).Value;

        private static readonly string _tickerColumn = "C";
        private static readonly int _tickerColumnNumber = GetColumnNumber(_tickerColumn).Value;

        private static readonly string _currencyColumn = "D";
        private static readonly int _currencyColumnNumber = GetColumnNumber(_currencyColumn).Value;

        private static readonly string _nameColumn = "E";
        private static readonly int _nameColumnNumber = GetColumnNumber(_nameColumn).Value;

        private static readonly string _closePriceColumn = "F";
        private static readonly int _closePriceColumnNumber = GetColumnNumber(_closePriceColumn).Value;

        private static readonly string _currentPriceColumn = "G";
        private static readonly int _currentPriceColumnNumber = GetColumnNumber(_currentPriceColumn).Value;

        private static readonly string _priceChangeColumn = "H";
        private static readonly int _priceChangeColumnNumber = GetColumnNumber(_priceChangeColumn).Value;

        private static readonly string _pricePercentageChangeColumn = "I";
        private static readonly int _pricePercentageChangeColumnNumber = GetColumnNumber(_pricePercentageChangeColumn).Value;

        private static readonly string _netPositionColumn = "J";
        private static readonly int _netPositionColumnNumber = GetColumnNumber(_netPositionColumn).Value;

        private static readonly string _currencyTickerColumn = "K";
        private static readonly int _currencyTickerColumnNumber = GetColumnNumber(_currencyTickerColumn).Value;

        private static readonly string _quoteFactorColumn = "L";
        private static readonly int _quoteFactorColumnNumber = GetColumnNumber(_quoteFactorColumn).Value;

        private static readonly string _fxRateColumn = "M";
        private static readonly int _fxRateColumnNumber = GetColumnNumber(_fxRateColumn).Value;

        private static readonly string _pnlColumn = "N";
        private static readonly int _pnlColumnNumber = GetColumnNumber(_pnlColumn).Value;


        protected abstract string ContributionFundColumn { get; }
        private int _contributionFundColumnNumber => GetColumnNumber(ContributionFundColumn).Value;

        protected abstract string ExposureColumn { get; }
        private int _exposureColumnNumber => GetColumnNumber(ExposureColumn).Value;


        protected abstract string ExposurePercentageFundColumn { get; }
        private int _exposurePercentageFundColumnNumber => GetColumnNumber(ExposurePercentageFundColumn).Value;


        protected abstract string ShortFundColumn { get; }
        private int? _shortFundColumnNumber => GetColumnNumber(ShortFundColumn);

        protected abstract string LongFundColumn { get; }
        private int? _longFundColumnNumber => GetColumnNumber(LongFundColumn);

        protected abstract string PriceMultiplierColumn { get; }
        private int _priceMultiplierColumnNumber => GetColumnNumber(PriceMultiplierColumn).Value;
        protected abstract string InstrumentTypeColumn { get; }
        private int _instrumentTypeColumnNumber => GetColumnNumber(InstrumentTypeColumn).Value;

        protected abstract string PriceDivisorColumn { get; }
        private int _priceDivisorColumnNumber => GetColumnNumber(PriceDivisorColumn).Value;


        protected abstract string ShortFundWinnersColumn { get; }
        private int? _shortFundWinnersColumnNumber => GetColumnNumber(ShortFundWinnersColumn);

        protected abstract string LongFundWinnersColumn { get; }
        private int? _longFundWinnersColumnNumber => GetColumnNumber(LongFundWinnersColumn);

        protected abstract string NavColumn { get; }
        private int _navColumnNumber => GetColumnNumber(NavColumn).Value;

        protected abstract string PreviousClosePriceColumn { get; }
        private int _previousClosePriceColumnNumber => GetColumnNumber(PreviousClosePriceColumn).Value;

        protected abstract string PreviousPriceChangeColumn { get; }
        private int _previousPriceChangeColumnNumber => GetColumnNumber(PreviousPriceChangeColumn).Value;

        protected abstract string PreviousPricePercentageChangeColumn { get; }
        private int _previousPricePercentageChangeColumnNumber => GetColumnNumber(PreviousPricePercentageChangeColumn).Value;

        protected abstract string PreviousNetPositionColumn { get; }
        private int _previousNetPositionColumnNumber => GetColumnNumber(PreviousNetPositionColumn).Value;

        protected abstract string PreviousFXRateColumn { get; }
        private int _previousFXRateColumnNumber => GetColumnNumber(PreviousFXRateColumn).Value;



        protected abstract string PreviousContributionFundColumn { get; }
        private int _previousContributionFundColumnNumber => GetColumnNumber(PreviousContributionFundColumn).Value;

        protected abstract string PreviousNavColumn { get; }
        private int _previousNavColumnNumber => GetColumnNumber(PreviousNavColumn).Value;

        Dictionary<int, ColumnDefinition> ColumnDefinitions = new Dictionary<int, ColumnDefinition>();



        public void SetColumnDefinitions()
        {
            
            ColumnDefinitions.Add(_controlColumnNumber, new ColumnDefinition(_controlColumnNumber, _controlColumn, "Control",CellStyler.StyleNormal,true,25.14m, null, null,null,XL.XlHAlign.xlHAlignLeft, true, false,false));
            ColumnDefinitions.Add(_instrumentMarketIdColumnNumber, new ColumnDefinition(_instrumentMarketIdColumnNumber, _instrumentMarketIdColumn, "Instrument Market Id", CellStyler.StyleNormal, true,18.57m, null, null, null, XL.XlHAlign.xlHAlignLeft, true, false, false));
            ColumnDefinitions.Add(_tickerColumnNumber, new ColumnDefinition(_tickerColumnNumber, _tickerColumn,"Ticker", CellStyler.StyleNormal, true,21.29m,null, null, null, XL.XlHAlign.xlHAlignLeft, true, false, false));
            ColumnDefinitions.Add(_currencyColumnNumber, new ColumnDefinition(_currencyColumnNumber, _currencyColumn,"Currency", CellStyler.StyleNormal, true, 11.86m,"CRNCY", null, null, XL.XlHAlign.xlHAlignLeft, true, false, false));
            ColumnDefinitions.Add(_nameColumnNumber, new ColumnDefinition(_nameColumnNumber, _nameColumn,"Name", CellStyler.StyleNormal, false,51.57m,"NAME", null, null,  XL.XlHAlign.xlHAlignLeft, true, false, false));
            ColumnDefinitions.Add(_closePriceColumnNumber, new ColumnDefinition(_closePriceColumnNumber, _closePriceColumn,"Close", CellStyler.StylePrice, false,12m,"PX_YEST_CLOSE", null, CellStyler.StyleFXRate, XL.XlHAlign.xlHAlignCenter, true, false, false));
            ColumnDefinitions.Add(_currentPriceColumnNumber, new ColumnDefinition(_currentPriceColumnNumber, _currentPriceColumn,"Current", CellStyler.StylePrice, false, 12m,"LAST_PRICE", null, CellStyler.StyleFXRate, XL.XlHAlign.xlHAlignCenter, true, false, false));
            ColumnDefinitions.Add(_priceChangeColumnNumber, new ColumnDefinition(_priceChangeColumnNumber, _priceChangeColumn,"Change", CellStyler.StylePriceChange, false, 12m, null, null, CellStyler.StyleFXRate, XL.XlHAlign.xlHAlignCenter, true, false, false));
            ColumnDefinitions.Add(_pricePercentageChangeColumnNumber, new ColumnDefinition(_pricePercentageChangeColumnNumber, _pricePercentageChangeColumn,"% Change", CellStyler.StylePercentageChange, false, 12m, null, null,null, XL.XlHAlign.xlHAlignCenter, true, false, false));
            ColumnDefinitions.Add(_netPositionColumnNumber, new ColumnDefinition(_netPositionColumnNumber, _netPositionColumn,"Units", CellStyler.StyleUnits, false,13.14m,null, null, null, XL.XlHAlign.xlHAlignCenter, true, false, false));
            ColumnDefinitions.Add(_currencyTickerColumnNumber, new ColumnDefinition(_currencyTickerColumnNumber, _currencyTickerColumn,"Currency Ticker", CellStyler.StyleNormal, true, 21.29m,null, null, null, XL.XlHAlign.xlHAlignCenter, true, false, false));
            ColumnDefinitions.Add(_quoteFactorColumnNumber, new ColumnDefinition(_quoteFactorColumnNumber, _quoteFactorColumn,"Quote Factor", CellStyler.StyleNormal, true,14.57m,"QUOTE_FACTOR",  null, null, XL.XlHAlign.xlHAlignCenter, true, false, false));
            ColumnDefinitions.Add(_fxRateColumnNumber, new ColumnDefinition(_fxRateColumnNumber, _fxRateColumn,"FX Rate", CellStyler.StyleFXRate, false,9m,"LAST_PRICE", null, null, XL.XlHAlign.xlHAlignCenter, true, false, true));           
            ColumnDefinitions.Add(_pnlColumnNumber, new ColumnDefinition(_pnlColumnNumber, _pnlColumn,"PNL", CellStyler.StylePNL, false, 12m,null, null, null, XL.XlHAlign.xlHAlignCenter, true, true, false));
            
            ColumnDefinitions.Add(_contributionFundColumnNumber, new ColumnDefinition(_contributionFundColumnNumber, ContributionFundColumn,"% Fund", CellStyler.StyleContribution, false, 9m, null, null, null, XL.XlHAlign.xlHAlignCenter, false,true,true));
            ColumnDefinitions.Add(_exposureColumnNumber, new ColumnDefinition(_exposureColumnNumber, ExposureColumn, "Exposure", CellStyler.StyleExposure, false,12m, null, null, null, XL.XlHAlign.xlHAlignCenter, true,true, false));

            ColumnDefinitions.Add(_exposurePercentageFundColumnNumber, new ColumnDefinition(_exposurePercentageFundColumnNumber, ExposurePercentageFundColumn,"% Fund", CellStyler.StyleExposurePercentage, false, 9m, null, null, null, XL.XlHAlign.xlHAlignCenter, true,true, true));

            if (_shortFundColumnNumber.HasValue)
            {
                ColumnDefinitions.Add(_shortFundColumnNumber.Value, new ColumnDefinition(_shortFundColumnNumber.Value, ShortFundColumn, "Short", CellStyler.StyleExposurePercentage, false, 9m, null, null, null, XL.XlHAlign.xlHAlignCenter, true, true, false));
            }
            if (_longFundColumnNumber.HasValue)
            {
                ColumnDefinitions.Add(_longFundColumnNumber.Value, new ColumnDefinition(_longFundColumnNumber.Value, LongFundColumn, "Long", CellStyler.StyleExposurePercentage, false, 9m, null, null, null, XL.XlHAlign.xlHAlignCenter, true, true, true));
            }

            ColumnDefinitions.Add(_priceMultiplierColumnNumber, new ColumnDefinition(_priceMultiplierColumnNumber, PriceMultiplierColumn,"Price Multiplier", CellStyler.StyleNormal,true,13.43m, null, null, null, XL.XlHAlign.xlHAlignCenter, true,false, false));
            ColumnDefinitions.Add(_instrumentTypeColumnNumber, new ColumnDefinition(_instrumentTypeColumnNumber, InstrumentTypeColumn,"Instrument Type", CellStyler.StyleNormal, true,14.29m, null, null, null, XL.XlHAlign.xlHAlignCenter, true, false, false));
            ColumnDefinitions.Add(_priceDivisorColumnNumber, new ColumnDefinition(_priceDivisorColumnNumber, PriceDivisorColumn,"Price Divisor", CellStyler.StyleNormal, true,11.29m, null, null, null, XL.XlHAlign.xlHAlignCenter, true, false, false));
           

            if (_shortFundWinnersColumnNumber.HasValue)
            {
                ColumnDefinitions.Add(_shortFundWinnersColumnNumber.Value, new ColumnDefinition(_shortFundWinnersColumnNumber.Value, ShortFundWinnersColumn, "Short Winners", CellStyler.StyleContribution, true, 12.57m, null, null, null, XL.XlHAlign.xlHAlignCenter, true, true, false));
            }
            if (_longFundWinnersColumnNumber.HasValue)
            {
                ColumnDefinitions.Add(_longFundWinnersColumnNumber.Value, new ColumnDefinition(_longFundWinnersColumnNumber.Value, LongFundWinnersColumn, "Long Winners", CellStyler.StyleContribution, true, 12.57m, null, null, null, XL.XlHAlign.xlHAlignCenter, true, true, false));
            }

            ColumnDefinitions.Add(_navColumnNumber, new ColumnDefinition(_navColumnNumber, NavColumn,"Nav", CellStyler.StyleNormal, true,11.29m, null, null, null, XL.XlHAlign.xlHAlignCenter, true, false, false));
            
            ColumnDefinitions.Add(_previousClosePriceColumnNumber, new ColumnDefinition(_previousClosePriceColumnNumber, PreviousClosePriceColumn,"Close", CellStyler.StylePrice, true,12m,"PX_CLOSE_1D", CellStyler.PreviousSectionGrey, null, XL.XlHAlign.xlHAlignCenter, true,false, false));
            ColumnDefinitions.Add(_previousPriceChangeColumnNumber, new ColumnDefinition(_previousPriceChangeColumnNumber, PreviousPriceChangeColumn,"Change", CellStyler.StylePrice, true,9m, null, CellStyler.PreviousSectionGrey, null, XL.XlHAlign.xlHAlignCenter, true, false, false));


            ColumnDefinitions.Add(_previousPricePercentageChangeColumnNumber, new ColumnDefinition(_previousPricePercentageChangeColumnNumber, PreviousPricePercentageChangeColumn, "% Change", CellStyler.StylePercentageChange, false,9m, null, CellStyler.PreviousSectionGrey, null, XL.XlHAlign.xlHAlignCenter, true, false, false));
            ColumnDefinitions.Add(_previousNetPositionColumnNumber, new ColumnDefinition(_previousNetPositionColumnNumber, PreviousNetPositionColumn,"Units", CellStyler.StyleUnits, true, 13.14m, null, CellStyler.PreviousSectionGrey, null, XL.XlHAlign.xlHAlignCenter, true, false, false));
            ColumnDefinitions.Add(_previousFXRateColumnNumber, new ColumnDefinition(_previousFXRateColumnNumber, PreviousFXRateColumn, "FX Rate", CellStyler.StyleFXRate, true,9m, "PX_YEST_CLOSE", CellStyler.PreviousSectionGrey, null, XL.XlHAlign.xlHAlignCenter, true, false, false));

            ColumnDefinitions.Add(_previousContributionFundColumnNumber, new ColumnDefinition(_previousContributionFundColumnNumber, PreviousContributionFundColumn, "% Fund", CellStyler.StyleContribution, false, 9m, null, CellStyler.PreviousSectionGrey, null, XL.XlHAlign.xlHAlignCenter, false, true, true));


            ColumnDefinitions.Add(_previousNavColumnNumber, new ColumnDefinition(_previousNavColumnNumber, PreviousNavColumn, "Nav", CellStyler.StyleNormal, true, 11.29m, null, CellStyler.PreviousSectionGrey, null, XL.XlHAlign.xlHAlignCenter, true, false,false));
        }

        

        private string _lastColumn => PreviousNavColumn;

        private readonly string _firstColumn = _controlColumn;

      //  private static int _firstRowOfData;
        private int _bloombergMnemonicRow;
        private static readonly string _previousReferenceDateLabel = $"${_currencyColumn}$1";
        private static readonly string _referenceDateLabel = $"${_nameColumn}$1";
        private static readonly string _totalSuffix = "#Total";
        private static readonly string _ignoreLabel = "#IGNORE#";
        private static readonly string _mnemonicLabel = "#MNEMONICS#";


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

        public static int? GetColumnNumber(string letter)
        {
            if (letter == null)
            {
                return null;
            }

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

        public void FinaliseFormatting(GroupingEntity lastFund)
        {
            _worksheet.Select();
            _worksheet.Activate();
            foreach (var style in ColumnDefinitions)
            {
                _worksheet.Columns[style.Key].ColumnWidth = style.Value.Width;
                _worksheet.Columns[style.Key].EntireColumn.Hidden = style.Value.IsHidden;
            }

            _worksheet.Application.ActiveWindow.Zoom = 115;
            _worksheet.Application.ActiveWindow.DisplayZeros = false;
            
            foreach (ColumnDefinition column in ColumnDefinitions.Values.Where(a => a.HasRightHandBorder))
            {
                var columnToAddBorderTo = _worksheet.get_Range($"{column.ColumnLabel}{_bloombergMnemonicRow}:{column.ColumnLabel}{lastFund.TotalRow.RowNumber}");
                columnToAddBorderTo.Borders[XL.XlBordersIndex.xlEdgeRight].LineStyle = XL.XlLineStyle.xlContinuous;
            }
            _worksheet.Rows[_bloombergMnemonicRow].EntireRow.Hidden = true;
            _worksheet.Application.ActiveWindow.FreezePanes = false;
            var a1 = _worksheet.get_Range("A1");
            a1.Select();
            a1.Activate();
            var toFreeze = _worksheet.get_Range($"{_closePriceColumn}{_bloombergMnemonicRow + 2}");
            toFreeze.Select();
            toFreeze.Activate();
            _worksheet.Application.ActiveWindow.FreezePanes = true;
        }
        


        public void AddPosition(Position previousPosition, Position position, GroupingEntity parent, Fund fund)
        {
            int rowToAddAt;
            if (previousPosition == null)
            {
                if (parent.Previous == null)
                {
                    rowToAddAt = _bloombergMnemonicRow+3;
                }
                else
                {
                    rowToAddAt = parent.Previous.TotalRow.RowNumber+2;
                }
            }
            else
            {
                rowToAddAt = previousPosition.RowNumber+1;
            }
            position.Row = AddRow(position.RowType, rowToAddAt);

            WritePosition(position, fund, true);
        }

        public void WritePosition(Position position, Fund fund, bool updateFormulas)
        {
            if (position.Identifier.Code != null && position.Identifier.Code.StartsWith("GB00BDX8CX86"))
            {
                int i = 0;
            }
            bool isArgentina = position.Currency == "ARS";
            bool isArgentinaCash = isArgentina && position.InstrumentTypeId == InstrumentTypeIds.FX;
            if (isArgentina)
            {
                int i = 0;
            }
            WriteValue(position.Row, ColumnDefinitions[_instrumentMarketIdColumnNumber], position.Identifier.Id, updateFormulas);
            WriteValue(position.Row, ColumnDefinitions[_tickerColumnNumber], isArgentinaCash ? _argentinaCurrencyTicker : position.Identifier.Code, updateFormulas);
            
            WriteName(position.Row, position.InstrumentTypeId, position.Name, updateFormulas);
            WriteCurrency(position.Row, position.InstrumentTypeId, position.Currency, updateFormulas);
            WriteClosePrice(position.Row, position.InstrumentTypeId, position.OdeyPreviousPrice, updateFormulas,position.OdeyPreviousPriceIsManual,position.PreviousInflationRatio);
            WriteCurrentPrice(position.Row, position.InstrumentTypeId, position.OdeyCurrentPrice, updateFormulas, position.OdeyCurrentPriceIsManual, isArgentinaCash,position.IsInflationAdjusted);
            WritePreviousClosePrice(position.Row, position.InstrumentTypeId, position.OdeyPreviousPreviousPrice, updateFormulas, position.OdeyPreviousPreviousPriceIsManual);

            WriteFormula(position.Row, ColumnDefinitions[ _priceChangeColumnNumber], GetSubtractFormula(position.RowNumber, _currentPriceColumn, _closePriceColumn), updateFormulas);
            WriteFormula(position.Row, ColumnDefinitions[_pricePercentageChangeColumnNumber], GetDivideFormula(position.RowNumber, _priceChangeColumn, _closePriceColumn, false), updateFormulas);
            WriteValue(position.Row, ColumnDefinitions[_netPositionColumnNumber], position.NetPosition, updateFormulas);
            
            WriteFormula(position.Row, ColumnDefinitions[_currencyTickerColumnNumber], GetCurrencyTickerFormula(position.RowNumber,fund.TotalRow.RowNumber, isArgentina), updateFormulas);
            WriteFormula(position.Row, ColumnDefinitions[_quoteFactorColumnNumber], GetQuoteFactorFormula(position.RowNumber, fund.TotalRow.RowNumber, isArgentina,fund.CurrencyId), updateFormulas);
            WriteFormula(position.Row, ColumnDefinitions[_fxRateColumnNumber], GetFXRateFormula(position.RowNumber, _fxRateColumn, fund.TotalRow.RowNumber), updateFormulas);
            WriteFormula(position.Row, ColumnDefinitions[_pnlColumnNumber], GetPNLFormula(position), updateFormulas);
            
            WriteFormula(position.Row, ColumnDefinitions[_contributionFundColumnNumber], GetDivideByNavFormula(position.RowNumber, _pnlColumn, true, fund), updateFormulas);

            WriteFormula(position.Row, ColumnDefinitions[_exposureColumnNumber], GetExposureFormula(position.InstrumentTypeId, position.RowNumber), updateFormulas);

            WriteFormula(position.Row, ColumnDefinitions[_exposurePercentageFundColumnNumber], GetDivideByNavFormula(position.RowNumber, ExposureColumn, false, fund), updateFormulas);

            
            if (_shortFundColumnNumber.HasValue)
            {
                WriteFormula(position.Row, ColumnDefinitions[_shortFundColumnNumber.Value], GetWriteIfIsLongCorrectColumn(position.InstrumentTypeId, position.RowNumber, false,ExposurePercentageFundColumn), updateFormulas);
            }
            if (_longFundColumnNumber.HasValue)
            {
                WriteFormula(position.Row, ColumnDefinitions[_longFundColumnNumber.Value], GetWriteIfIsLongCorrectColumn(position.InstrumentTypeId, position.RowNumber, true, ExposurePercentageFundColumn), updateFormulas);
            }

            WriteFormula(position.Row, ColumnDefinitions[_priceMultiplierColumnNumber], GetPriceMultiplierFormula(position.RowNumber), updateFormulas);
            WriteValue(position.Row, ColumnDefinitions[_instrumentTypeColumnNumber], position.InstrumentTypeId, updateFormulas);
            WriteValue(position.Row, ColumnDefinitions[_priceDivisorColumnNumber], position.PriceDivisor, updateFormulas);
           
            if (_shortFundWinnersColumnNumber.HasValue)
            {
                WriteFormula(position.Row, ColumnDefinitions[_shortFundWinnersColumnNumber.Value], GetWinnerColumn(position.RowNumber, false, ContributionFundColumn), updateFormulas);
            }
            if (_longFundWinnersColumnNumber.HasValue)
            {
                WriteFormula(position.Row, ColumnDefinitions[_longFundWinnersColumnNumber.Value], GetWinnerColumn(position.RowNumber, true, ContributionFundColumn), updateFormulas);
            }


            WriteFormula(position.Row, ColumnDefinitions[_previousPriceChangeColumnNumber], GetSubtractFormula(position.RowNumber, _closePriceColumn, PreviousClosePriceColumn), updateFormulas);
            WriteFormula(position.Row, ColumnDefinitions[_previousPricePercentageChangeColumnNumber], GetDivideFormula(position.RowNumber, PreviousPriceChangeColumn, PreviousClosePriceColumn, false), updateFormulas);

            WriteValue(position.Row, ColumnDefinitions[_previousNetPositionColumnNumber], position.PreviousNetPosition, updateFormulas);
            WriteFormula(position.Row, ColumnDefinitions[_previousFXRateColumnNumber], GetFXRateFormula(position.RowNumber, PreviousFXRateColumn, fund.TotalRow.RowNumber), updateFormulas);
            WriteFormula(position.Row, ColumnDefinitions[_previousContributionFundColumnNumber], GetPreviousContribution(position, position.RowNumber, fund), updateFormulas);
        }

        private void WriteName(Row row, InstrumentTypeIds instrumentTypeId, string odeyName,bool updateFormulas)
        {
            var columnDefinition = ColumnDefinitions[_nameColumnNumber];
            if (instrumentTypeId == InstrumentTypeIds.DoNotDelete)
            {
                WriteFormula(row, columnDefinition, GetBloombergMnemonicFormula(row.RowNumber, _nameColumn,_tickerColumn), updateFormulas);
            }       
            else
            {
                WriteValue(row, columnDefinition, odeyName,updateFormulas);
            }
        }

        private void WriteCurrency(Row row, InstrumentTypeIds instrumentTypeId, string odeyCurrency, bool updateFormulas)
        {
            var columnDefinition = ColumnDefinitions[_currencyColumnNumber];
            if (instrumentTypeId == InstrumentTypeIds.FX || instrumentTypeId == InstrumentTypeIds.PrivatePlacement)
            {
                WriteValue(row, columnDefinition, odeyCurrency, updateFormulas);
            }
            else
            {
                WriteFormula(row, columnDefinition, GetBloombergMnemonicFormula(row.RowNumber, _currencyColumn), updateFormulas);
            }
        }

        private void WriteClosePrice(Row row, InstrumentTypeIds instrumentTypeId, decimal? odeyPreviousPrice, bool updateFormulas,bool isManual,decimal? previousIndexRatio)
        {
            var columnDefinition = ColumnDefinitions[_closePriceColumnNumber];
            if (instrumentTypeId == InstrumentTypeIds.FX || instrumentTypeId == InstrumentTypeIds.PrivatePlacement || isManual)
            {
                decimal? value = odeyPreviousPrice;
                if (odeyPreviousPrice.HasValue && previousIndexRatio.HasValue)
                {
                    value = odeyPreviousPrice.Value * previousIndexRatio.Value;
                }
                WriteValue(row,  columnDefinition, value, updateFormulas);           
            }
            else
            {
                var formula = GetBloombergMnemonicFormula(row.RowNumber, _closePriceColumn);
                if (previousIndexRatio.HasValue)
                {
                    formula = $"{formula}*{previousIndexRatio.Value}";
                }
                WriteFormula(row, columnDefinition,formula , updateFormulas || previousIndexRatio.HasValue);
            }
        }

        private void WriteCurrentPrice(Row row, InstrumentTypeIds instrumentTypeId, decimal? odeyCurrentPrice, bool updateFormulas, bool isManual,bool isArgentinaCash, bool isInflationAdjusted)
        {
            var columnDefinition = ColumnDefinitions[_currentPriceColumnNumber];
            if (instrumentTypeId == InstrumentTypeIds.PrivatePlacement || isManual)
            {
                WriteValue(row, columnDefinition, odeyCurrentPrice, updateFormulas);
            }
            else
            {
                var formula = GetBloombergMnemonicFormula(row.RowNumber, _currentPriceColumn);
                if (isArgentinaCash)
                {
                    formula = $"{formula} * {_quoteFactorColumn}{row.RowNumber}";
                }
                if (isInflationAdjusted)
                {
                    formula = $@"{formula} * BDP({_tickerColumn}{row.RowNumber},""MOST_RECENT_REPORTED_FACTOR"")";
                }
                WriteFormula(row, columnDefinition, formula, updateFormulas);
            }
        }

        private void WritePreviousClosePrice(Row row, InstrumentTypeIds instrumentTypeId, decimal? odeyPreviousPreviousPrice, bool updateFormulas, bool isManual)
        {
            var columnDefinition = ColumnDefinitions[_previousClosePriceColumnNumber];
            if (instrumentTypeId == InstrumentTypeIds.FX || instrumentTypeId == InstrumentTypeIds.PrivatePlacement || isManual)
            {

                WriteValue(row, columnDefinition, odeyPreviousPreviousPrice, updateFormulas);        
            }
            else
            {
                WriteFormula(row, columnDefinition, GetBloombergMnemonicHistoryFormula(row.RowNumber, _tickerColumn, PreviousClosePriceColumn), updateFormulas);
            }
        }

        public void UpdateSums(GroupingEntity entity)
        {
            int firstRowNumber = _bloombergMnemonicRow+2;
            if (entity.Previous!=null)
            {
                firstRowNumber = entity.Previous.TotalRow.RowNumber + 1;
            }
            int lastRowNumber = entity.TotalRow.RowNumber - 1;

            foreach (ColumnDefinition definition in ColumnDefinitions.Values)
            {
                if (definition.IsSummable)
                {
                    WriteFormula(entity.TotalRow, definition, GetSumFormula(firstRowNumber, lastRowNumber, definition.ColumnLabel),true);                    
                }
                else
                {
                    CellStyler.Instance.ApplyStyle(entity.TotalRow, definition);
                }
            }
        }



        private Row AddRow(RowType rowType,int rowNumber)
        {
            _worksheet.Rows[rowNumber].Insert(XL.XlDirection.xlUp, XL.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            var insertedRow = GetRow(rowType, rowNumber);          
            insertedRow.Range.Interior.Color = XL.XlColorIndex.xlColorIndexNone;
            insertedRow.Range.RowHeight = 12;
            foreach (ColumnDefinition column in ColumnDefinitions.Values)
            {
                CellStyler.Instance.ApplyStyle(insertedRow, column);
            }
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

        

        private Row GetRow(RowType rowType,int rowNumber)
        {
            return new Row(rowType, _worksheet.get_Range($"{_firstColumn}{rowNumber}:{_lastColumn}{rowNumber}"));
        }

        public void UpdateTotalsOnTotalRow(GroupingEntity groupingEntity)
        {
            int[] rowNumbers = groupingEntity.Children.Select(a => a.Value.RowNumber).ToArray();


            foreach (ColumnDefinition column in ColumnDefinitions.Values)
            {
                if (column.IsSummable)
                {
                    UpdateTotalOnTotalRow(groupingEntity.TotalRow, column, rowNumbers);
                }
                else
                {
                    CellStyler.Instance.ApplyStyle(groupingEntity.TotalRow, column);
                }
            }

        }

        public void UpdateNavs(GroupingEntity groupingEntity)
        {           
            WriteValue(groupingEntity.TotalRow, ColumnDefinitions[_navColumnNumber], groupingEntity.Nav,false);
            WriteValue(groupingEntity.TotalRow, ColumnDefinitions[_previousNavColumnNumber], groupingEntity.PreviousNav, false);
            if (groupingEntity is Fund)
            {
                WriteValue(groupingEntity.TotalRow, ColumnDefinitions[_currencyColumnNumber], ((Fund)groupingEntity).Currency, false);
            }
        }

        private void UpdateTotalOnTotalRow(Row row, ColumnDefinition column, int[] rowNumbers)
        {

            string formula;

            formula = "=" + string.Join("+", rowNumbers.Select(a => column.ColumnLabel + a));

            WriteFormula(row, column, formula, true);

        }

        private string GetControlString(GroupingEntity parent, string codeToAdd)
        {
            EntityTypes entityType;
            string parentControlString;           
            if (parent == null)
            {
                entityType = EntityTypes.Fund;
                parentControlString = "###Total";
            }
            else
            {
                entityType = parent.ChildEntityType;
                parentControlString = parent.ControlString;
            }
            string[] values = parentControlString.Split('#');

            values[(int)entityType] = codeToAdd;

            return string.Join("#", values);
        }

        public void AddTotalRow(GroupingEntity group)
        {
            int addAtRowNumber;
            if (group.Previous == null)
            {              
                addAtRowNumber = _bloombergMnemonicRow+2;
            }
            else
            {
                addAtRowNumber = group.Previous.TotalRow.RowNumber+1;
            }

            group.TotalRow = AddRow(group.RowType, addAtRowNumber);

            AddRow(RowType.Blank, group.TotalRow.RowNumber);//Gap Between sections

            group.ControlString = GetControlString(group.Parent, group.Identifier.Code);
            WriteValue(group.TotalRow, ColumnDefinitions[_controlColumnNumber], group.ControlString,false);
            WriteValue(group.TotalRow, ColumnDefinitions[_nameColumnNumber], group.Name, false);                     
        }

        public void SetupSheet()
        {
            SetColumnDefinitions();
            int? mnemonicRow = FindRow(_mnemonicLabel, _controlColumn);
            if (!mnemonicRow.HasValue)
            {
                mnemonicRow = CreateMneumonicRow();                
            }
            _bloombergMnemonicRow = mnemonicRow.Value;   
        }

        private int CreateMneumonicRow()
        {
            var mnemonicRow = 3;
            var bloombergRow=  GetRow(RowType.Mnemonic,mnemonicRow);            

            var headerRow = GetRow(RowType.Header, mnemonicRow + 1);
            foreach (var column in ColumnDefinitions.Values)
            {
                if (!string.IsNullOrWhiteSpace(column.BloombergMneumonic))
                {
                    WriteValue(bloombergRow, column, column.BloombergMneumonic,true);
                }
                WriteValue(headerRow, column, column.HeaderLabel, true);
            }
            return mnemonicRow;
        }

        private string CreateTotalLabel(string fund,  string assetClass, string country)
        {
            return string.Join("#", new[] { fund, assetClass, country }) + _totalSuffix;                
        }

        public void AddFundRange(Fund fund)
        {
            int firstRowOfData = _bloombergMnemonicRow+2;
            if (fund.Previous != null)
            {
                var previousRange = ((Fund)fund.Previous).Range;
                firstRowOfData = previousRange.Row + previousRange.Rows.Count;
            }

            string fundTotalLabel = CreateTotalLabel(fund.Name, null, null);
            int? lastRow = FindRow(fundTotalLabel, _controlColumn);

            if (!lastRow.HasValue)
            {
                if (fund.Previous != null)
                {
                    throw new ApplicationException($"No Total Row exists for fund {fund.Name}");
                }
                else
                {
                    lastRow = _bloombergMnemonicRow + 3;
                }
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
            return new Position(existingPosition.Identifier, name, priceDivisor ?? 1, instrumentTypeId, invertPNL,false) { Currency = currency};
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

            var cell = _worksheet.Range[_referenceDateLabel].Cells;
            cell.Style = CellStyler.StyleNormal;
            cell.Value = referenceDate;
            cell.HorizontalAlignment = XL.XlHAlign.xlHAlignLeft;
            cell.Font.Bold = true;

        }

 

        private void WriteValue(Row row, ColumnDefinition columnDefinition, object value,bool updateStyle)
        {
            var cell = row.Range.Cells[1, columnDefinition.ColumnNumber];
            cell.Value = value;

            if (updateStyle)
            {
                CellStyler.Instance.ApplyStyle(cell, row.RowType, columnDefinition);
            }
        }

        #region Formulas







        private void WriteFormula(Row row,ColumnDefinition column, string formula, bool updateFormulas)
        {
            if (updateFormulas)
            {
                var cell = row.Range.Cells[1, column.ColumnNumber];

                cell.Formula = formula;

                CellStyler.Instance.ApplyStyle(cell, row.RowType, column);
            }
        }


        


        private static readonly string _bloombergError = "\"#N/A N/A\"";
        private static readonly string _noRealTimePrice = "\"#N/A Real Time\"";


        private string GetSubtractFormula(int rowNumber, string column1, string column2)
        {
            string column1AC = $"{ column1 }{ rowNumber}";
            string column2AC = $"{ column2 }{ rowNumber}";
            return $"=if(or(or({column1AC}={_bloombergError},{column1AC}={_noRealTimePrice}),or({column2AC}={_bloombergError},{column2AC}={_noRealTimePrice})),0,  {column1AC} - {column2AC})";
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
                columns = new string[] { _currentPriceColumn, _netPositionColumn, PriceMultiplierColumn };
            }
            string formula = GetMultiplyFormula(rowNumber, columns, divideColumn,false, _netPositionColumn, _currentPriceColumn, instrumentTypeId == InstrumentTypeIds.FX);

            return formula;
        }

        private string GetPNLFormula(Position position)
        {
            
            if (position.InstrumentTypeId == InstrumentTypeIds.FX)
            {
                return GetMultiplyFormula(position.RowNumber, new string[] { _priceChangeColumn, _netPositionColumn }, new string[] { _fxRateColumn,_currentPriceColumn }, position.InvertPNL,null,null,false);
            }
            else
            {
                return GetMultiplyFormula(position.RowNumber, new string[] { _priceChangeColumn, _netPositionColumn, PriceMultiplierColumn }, new string[] { _fxRateColumn },false, null, null, false);
            }
            
        }


        private string GetMultiplyFormula(int rowNumber, string[] columns, string[] divideColumn, bool invert, string columnToTestForZero, string columnToTestForErrors, bool absolute)
        {
            string divideColumns = "";
            if (divideColumn != null && divideColumn.Length>0)
            {
                divideColumns = "/"+string.Join("/", divideColumn.Select(a => a + rowNumber));
            }

            string formula = string.Join("*", columns.Select(a => a + rowNumber)) + divideColumns + (invert ? "*-1" : "");

            bool columnToTestForZeroExists = string.IsNullOrWhiteSpace(columnToTestForZero);
            bool columnToTestForErrorExists = string.IsNullOrWhiteSpace(columnToTestForErrors);

            if (!columnToTestForZeroExists && !columnToTestForZeroExists)
            {
                string qrc = columnToTestForZero + rowNumber;
                string prc = columnToTestForErrors + rowNumber;
                formula = $"if(OR(OR({qrc}=0,{prc} = {_bloombergError}),{prc}={_noRealTimePrice}),0,{formula})";
            }
            else if (!columnToTestForZeroExists || !columnToTestForZeroExists)
            {
                throw new ApplicationException("Not expecting only 1 of errors and zero.");
            }

            if (absolute)
            {
                formula = $"Abs({formula})";
            }

            return "=" + formula;
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
                    pnlFormula = GetMultiplyFormula(rowNumber, new string[] { PreviousPriceChangeColumn, PreviousNetPositionColumn }, new string[] { PreviousFXRateColumn, PreviousClosePriceColumn },position.InvertPNL,null, null, false);
                }
                else
                {
                    pnlFormula = GetMultiplyFormula(rowNumber, new string[] { PreviousPriceChangeColumn, PreviousNetPositionColumn, PriceMultiplierColumn }, new string[] { PreviousFXRateColumn },false, null, null, false);
                }
                pnlFormula = pnlFormula.Replace("=", "");
                return $"={pnlFormula} / {PreviousNavColumn}{groupingEntity.TotalRow.RowNumber}";
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
            return $"={column}{rowNumber} / {NavColumn}{groupingEntity.TotalRow.RowNumber}{multiplyBy100}";
        }

        private string GetBloombergMnemonicFormula(int rowNumber,string column)
        {
            return GetBloombergMnemonicFormula(rowNumber, column, _tickerColumn);
        }

        private string GetBloombergMnemonicHistoryFormula(int rowNumber,string tickerColumn, string column)
        {
            return $"=BDH({tickerColumn}{rowNumber},${column}${_bloombergMnemonicRow},{_previousReferenceDateLabel},{_previousReferenceDateLabel})";
        }

        private string GetQuoteFactorFormula(int rowNumber,int fundTotalRowNumber,bool isArgentina, int fundCurrencyId)
        {
            if (isArgentina)
            {
                if (fundCurrencyId == (int)CurrencyIds.USD)
                {
                    return "=1";
                }
                else if (fundCurrencyId!= (int)CurrencyIds.EUR)
                {
                    throw new ApplicationException("Need to map fx currency thats not eur"); 
                }
                return  $"=BDP(\"EURUSD Curncy\", \"PX_LAST\")";
            }
            return $"=IF({_currencyColumn}{rowNumber} = {_currencyColumn}{fundTotalRowNumber},1,{GetBloombergMnemonicFormula(rowNumber, _quoteFactorColumn, _currencyTickerColumn).Replace("=", "")})";
        }

        private string GetFXRateFormula(int rowNumber, string fxRateColumn, int fundTotalRowNumber)
        {
            return $"=IF({_currencyColumn}{rowNumber} = {_currencyColumn}{fundTotalRowNumber},1,{GetBloombergMnemonicFormula(rowNumber, fxRateColumn, _currencyTickerColumn).Replace("=","")}*{_quoteFactorColumn}{rowNumber})";
        }

        private string GetBloombergMnemonicFormula(int rowNumber, string mnemonicColumn,string tickerColumn)
        {
            return $"=BDP({tickerColumn}{rowNumber},${mnemonicColumn}${_bloombergMnemonicRow})";
        }

        private static readonly string _argentinaCurrencyTicker = ".AREQIMP G Index";

        private string GetCurrencyTickerFormula(int rowNumber,int fundTotalRowNumber, bool isArgentina)
        {
            if (isArgentina)
            {
                return _argentinaCurrencyTicker;
            }
            return $"=CONCATENATE({_currencyColumn}{fundTotalRowNumber},{_currencyColumn}{rowNumber}, \" Curncy\")";
        }

        private string GetPriceMultiplierFormula(int rowNumber)
        {
            return $"=IF(EXACT({_currencyColumn}{rowNumber},UPPER({_currencyColumn}{rowNumber})),1,0.01)/{PriceDivisorColumn}{rowNumber}";
        }

        private string GetWriteIfIsLongCorrectColumn(InstrumentTypeIds instrumentTypeId,int rowNumber, bool isLong,string exposurePercentageColumn)
        {
            if (instrumentTypeId == InstrumentTypeIds.FX)
            {
                return null;
            }
            return GetWriteIfStatement(rowNumber, GetExposureIsLongTest(rowNumber, isLong), exposurePercentageColumn);
        }

        private string GetWinnerColumn(int rowNumber, bool isLong,string contributionColumn)
        {
            string exposureTest = GetExposureIsLongTest(rowNumber, isLong);
            string winnerTest = GetIsGreaterThanZeroTest(rowNumber, true, contributionColumn);
            string test = $"AND({exposureTest},{winnerTest})";
            return GetWriteIfStatement(rowNumber, test, contributionColumn);
        }

        private string GetExposureIsLongTest(int rowNumber, bool isLong)
        {
            return GetIsGreaterThanZeroTest(rowNumber,isLong, ExposurePercentageFundColumn);
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
