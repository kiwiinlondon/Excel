using Odey.Framework.Keeley.Entities.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class Matcher
    {
        private EntityBuilder _entityBuilder;
        private DataAccess _dataAccess;
        private SheetAccess  _sheetAccess;
        private InstrumentRetriever _instrumentRetriever;
        public Matcher (EntityBuilder entityBuilder, DataAccess dataAccess,SheetAccess sheetAccess, InstrumentRetriever instrumentRetriever)
        {
            _entityBuilder = entityBuilder; 
            _dataAccess = dataAccess;
            _sheetAccess = sheetAccess;
            _instrumentRetriever = instrumentRetriever;
        }



        public void Match(bool resetExisting)
        {
            _sheetAccess.DisableCalculations();

            var rates = _dataAccess.GetFXRates();
            _sheetAccess.WriteDates(_dataAccess.PreviousReferenceDate, _dataAccess.ReferenceDate);
            Fund previousFund = null;
            foreach (Fund fund in _funds.Values.OrderBy(a=>a.Ordering))
            {
                fund.Previous = previousFund;
                MatchFund(fund, resetExisting, rates);
                previousFund = fund;
            }
            _sheetAccess.EnableCalculations();
        }


        public string AddTicker(string ticker)
        {
            if (_instrumentRetriever == null)
            {
                _instrumentRetriever = new InstrumentRetriever(new BloombergSecuritySetup(), _dataAccess);
            }
            ticker = _instrumentRetriever.FixTicker(ticker);
            string message;
            if (_instrumentRetriever.ValidateTicker(ticker, out message))
            {
                var instrument = _instrumentRetriever.Get(ticker, out message);
                if (instrument != null)
                {
                    var position = _sheetAccess.AddInstrument(instrument);
                    message = $"Success. {ticker} added to Country {instrument.ExchangeCountryName} at row ";// {position.RowNumber}";
                }
            }
            return message;
        }

        private Dictionary<FundIds, Fund> _funds;

        public void BuildFunds()
        {
            _funds = new Dictionary<FundIds, Fund>();
            
            BuildAndAddFund( FundIds.OEI);
            BuildAndAddFund( FundIds.OEIMAC);            
            BuildAndAddFund( FundIds.OEIMACGBPBSHARECLASS);
            BuildAndAddFund( FundIds.OEIMACGBPBMSHARECLASS);
        }

        private Fund BuildAndAddFund(FundIds fundId)
        {
            var fund = BuildFund(fundId);
            _funds.Add(fundId, fund);
            return fund;
        }

        private Fund BuildFund(FundIds fundId)
        {
            Fund fund = _entityBuilder.GetFund(fundId);
            return fund;
        }

       

        private void MatchFund(Fund fund, bool resetExisting, List<FXRateDTO> rates)
        {
            if (resetExisting)
            {
                _entityBuilder.RemovePositions(fund);
            }
            _entityBuilder.AddPortfolio(fund);
            _sheetAccess.AddFundRange(fund);
            _entityBuilder.AddExistingPortfolio(fund);
            WriteGroupingEntity(fund, rates, null, fund);
        }

        

        private void WriteGroupingEntity(GroupingEntity entity, List<FXRateDTO> rates, Book book, Fund fund)
        {
            if (entity.TotalRow == null)
            {               
                _sheetAccess.AddTotalRow(entity); 
            }
            if (entity.ChildrenArePositions)
            {
                WritePositions(entity, rates, book, fund);
                _sheetAccess.UpdateSums(entity);
            }
            else
            {
                GroupingEntity previous = null;
                foreach (IChildEntity childEntity in entity.Children.Values.OrderBy(a => a.Ordering))
                {       

                    if (childEntity is Book)
                    {
                        book = (Book)childEntity;
                    }
                    GroupingEntity groupingEntity = (GroupingEntity)childEntity;
                    groupingEntity.Previous = previous;
                    WriteGroupingEntity(groupingEntity, rates, book, fund);
                    previous = groupingEntity;
                }
                entity.Previous = previous;
                _sheetAccess.UpdateTotalsOnTotalRow(entity);
                _sheetAccess.UpdateNavs(entity);

            }  
            
        }

        private void WritePositions(GroupingEntity entity, List<FXRateDTO> rates, Book book, Fund fund)
        {
            Position previous = null;
            foreach (Position position in entity.Children.Values.OrderBy(a => a.Ordering))
            {
                WritePosition(previous, position, entity, rates, book, fund);
                previous = position;
            }
        }

       
       

        private void WritePosition(Position previousPosition,Position position,GroupingEntity parent,List<FXRateDTO> rates,
            Book book,Fund fund)
        {
            if (position.Row != null)
            {
                if (position.InstrumentTypeId == InstrumentTypeIds.FX)
                {
                    EnhanceFXPosition(position, rates);
                }
                _sheetAccess.UpdatePosition(true, position,book,fund);
            }
            else
            {
                _sheetAccess.AddPosition(previousPosition, position, parent, book, fund);
            }
        }


        private void EnhanceFXPosition(Position position, List<FXRateDTO> rates)
        {
            string ticker = position.Identifier.Code;
            string currency1 = ticker.Substring(0, 3);
            string currency2 = ticker.Substring(3, 3);

            var matchingRates = rates.Where(a => (a.FromCurrency == currency1 || a.FromCurrency == currency2) && (a.ToCurrency == currency1 || a.ToCurrency == currency2));
            if (matchingRates.Count()!=1)
            {
                throw new ApplicationException($"Cannot find rate for ticker {ticker}. Count = {matchingRates.Count()}");
            }
            var rate = matchingRates.First();
            if (rate.FromCurrency == currency1)
            {
                position.OdeyPreviousPreviousPrice = rate.PreviousPreviousValue;
                position.OdeyPreviousPrice = rate.PreviousValue;
            }
            else
            {
                position.OdeyPreviousPreviousPrice = 1 / rate.PreviousPreviousValue;
                position.OdeyPreviousPrice = 1 / rate.PreviousValue;
            }
        }
     
        //private void MatchDTOToTickerRow(GroupingEntity parent, PortfolioDTO dto)
        //{

        //    Position position;
        //    if (!parent.Children.ContainsKey(dto.Ticker))
        //    {
        //        position = new Position(dto.Ticker, dto.Name, dto.PriceDivisor, dto.TickerTypeId,null);
        //        parent.Children.Add(position.Ticker, position);
        //    }
        //    else
        //    {
        //        position = (Position)parent.Children[dto.Ticker];
        //    }
        //    position.PreviousNetPosition = dto.PreviousNetPosition;
        //    position.NetPosition = dto.CurrentNetPosition;
        //    position.Currency = dto.Currency;
        //    if (position.TickerTypeId != TickerTypeIds.FX)
        //    {
        //        position.Name = dto.Name;
        //    }
        //    else 
        //    {
        //        int i = 0;
        //    }
        //    position.PriceDivisor = dto.PriceDivisor;
        //    position.TickerTypeId = dto.TickerTypeId;
        //    position.OdeyPreviousPreviousPrice = dto.PreviousPreviousPrice;
        //    position.OdeyPreviousPrice = dto.PreviousPrice;
        //    position.OdeyCurrentPrice = dto.CurrentPrice;

        //}
    }
}
