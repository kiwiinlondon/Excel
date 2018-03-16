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



        public void Match(bool refreshFormulas)
        {
            _sheetAccess.DisableCalculations();

            var rates = _dataAccess.GetFXRates();
            _sheetAccess.WriteDates(_dataAccess.PreviousReferenceDate, _dataAccess.ReferenceDate);
            var funds = BuildFunds(new FundIds[] { FundIds.OEI, FundIds.OEIMAC, FundIds.OEIMACGBPBSHARECLASS, FundIds.OEIMACGBPBMSHARECLASS });


            foreach (Fund fund in funds.Values.OrderBy(a => a.Ordering))
            {                
                MatchFund(fund, rates, refreshFormulas);
            }
            _sheetAccess.EnableCalculations();
            _sheetAccess.Save();
        }


        public string AddTicker(string ticker)
        {
            var m = AddTicker(ticker, null);
            _sheetAccess.Save();
            return m;
        }

        private string AddTicker(string ticker,Fund fund)
        {
            if (_instrumentRetriever == null)
            {
                _instrumentRetriever = new InstrumentRetriever(new BloombergSecuritySetup(), _dataAccess);
            }
            ticker = _instrumentRetriever.FixTicker(ticker);
            string message;
            if (_instrumentRetriever.ValidateTicker(ticker, out message))
            {
                if (fund == null)
                {
                    fund = BuildOEIFromSheet();
                }
                Book book = (Book)fund.Children.Values.FirstOrDefault(a => (((Book)a).BookId == (int)BookIds.OEI));
                if (TickerAlreadyExists(ticker, book))
                {
                    message = "Ticker Already Exists";
                }
                else
                {
                    var instrument = _instrumentRetriever.Get(ticker, out message);
                    if (instrument != null)
                    {
                        var country = _entityBuilder.AddInstrument(book, instrument);

                        WriteGroupingEntity(fund, null, book, fund, false, false);
                        Position position = (Position)country.Children[instrument.Identifier];
                        message = $"Success. {ticker} added to Country {instrument.ExchangeCountryName} at row {position.RowNumber}";
                    }
                }
            }

            return message;
        }

        private bool TickerAlreadyExists(string ticker, GroupingEntity groupingEntity)
        {
            if (groupingEntity.ChildrenArePositions)
            {
                var identitifer = new Identifier(null, ticker);
                return groupingEntity.Children.ContainsKey(identitifer);
            }
            else
            {
                foreach (var child in groupingEntity.Children.Values)
                {
                    var found = TickerAlreadyExists(ticker, (GroupingEntity)child);
                    if (found)
                    {
                        return true;
                    }
                }
                return false;
            }
            
        }

        public string AddBulk()
        {
            _sheetAccess.DisableCalculations();
            var tickers = _sheetAccess.GetBulkTickers();
            Fund fund = BuildOEIFromSheet();
            foreach(var ticker in tickers)
            {
                AddTicker(ticker, fund);
            }
            _sheetAccess.EnableCalculations();
            _sheetAccess.Save();
            return "Success";
        }



        private Dictionary<FundIds, Fund> BuildFunds(FundIds[] fundIds)
        {
            var funds = new Dictionary<FundIds, Fund>();
            Fund previous = null;
            foreach(FundIds fundId in fundIds)
            {                
                var fund = BuildFund(fundId, previous);                
                funds.Add(fundId, fund);
                previous = fund;
            }
            return funds;
        }
        
        private Fund BuildFund(FundIds fundId,Fund previous)
        {
            Fund fund = _entityBuilder.GetFund(fundId);
            fund.Previous = previous;
            _sheetAccess.AddFundRange(fund);
            return fund;
        }

        private Fund BuildOEIFromSheet()
        {
            Fund fund = BuildFund(FundIds.OEI, null);
            
            List<Position> positionsToBeUpdatedFromDatabase = new List<Position>();
            _entityBuilder.AddExistingPortfolio(fund, positionsToBeUpdatedFromDatabase);
            return fund;
        }

        private void MatchFund(Fund fund, List<FXRateDTO> rates, bool refreshFormulas)
        {
            _entityBuilder.AddPortfolio(fund);

            List<Position> positionsToBeUpdatedFromDatabase = new List<Position>();
            _entityBuilder.AddExistingPortfolio(fund, positionsToBeUpdatedFromDatabase);
            UpdatePositionsFromDatabase(positionsToBeUpdatedFromDatabase);

            WriteGroupingEntity(fund, rates, null, fund, true, refreshFormulas);
        }

        private void UpdatePositionsFromDatabase(List<Position> positions)
        {
            var instruments = _dataAccess.GetInstruments(positions.Select(a => a.Identifier).ToList());
            foreach(Position position in positions)
            {
                var instrument = instruments.FirstOrDefault(a => a.Identifier == position.Identifier);
                if (instrument!=null)
                {
                    position.Identifier.Id = instrument.Identifier.Id;
                    position.Identifier.Code = instrument.Identifier.Code;
                    position.Name = instrument.Name;
                    position.PriceDivisor = instrument.PriceDivisor;
                    if (position.InstrumentTypeId != InstrumentTypeIds.DoNotDelete)
                    {
                        position.InstrumentTypeId = instrument.InstrumentTypeId;
                    }
                }
            }
        }

        private void WriteGroupingEntity(GroupingEntity entity, List<FXRateDTO> rates, Book book, Fund fund, bool updateExistingPositions, bool forceRefresh)
        {
            if (entity.TotalRow == null)
            {               
                _sheetAccess.AddTotalRow(entity); 
            }
            if (entity.ChildrenArePositions)
            {
                WritePositions(entity, rates, book, fund, updateExistingPositions, forceRefresh);               
            }
            else 
            {
                var updateTotal = false;
                GroupingEntity previous = null;
                foreach (IChildEntity childEntity in entity.Children.Values.OrderBy(a => a.Ordering))
                {       
                    if (childEntity is Book)
                    {
                        book = (Book)childEntity;
                    }
                    GroupingEntity groupingEntity = (GroupingEntity)childEntity;
                    if (groupingEntity.TotalRow==null)
                    {
                        updateTotal = true;
                    }
                    groupingEntity.Previous = previous;
                    WriteGroupingEntity(groupingEntity, rates, book, fund, updateExistingPositions, forceRefresh);
                    if (groupingEntity.Children.Count == 0 && entity.ChildrenAreDeleteable)
                    {
                        updateTotal = true;
                        entity.ChildrenToDelete.Add(groupingEntity);
                    }
                    else
                    {
                        previous = groupingEntity;
                    }
                }

                RemoveChildrenToBeDeleted(entity, updateExistingPositions);
                entity.Previous = previous;
                if (updateTotal || forceRefresh)
                {
                    _sheetAccess.UpdateTotalsOnTotalRow(entity);
                }
                if (updateExistingPositions)
                {
                    _sheetAccess.UpdateNavs(entity);
                }
            }
            HideRows(entity);
        }

        private void HideRows(GroupingEntity entity)
        {
            if (entity.ChildrenAreHidden)
            {
                _sheetAccess.HideRows(entity.Previous.TotalRow.Row + 1, entity.TotalRow.Row - 1);
            }
        }

        private void RemoveChildrenToBeDeleted(GroupingEntity entity, bool updateExistingPositions)
        {
            if (updateExistingPositions)
            {
                foreach (var child in entity.ChildrenToDelete)
                {
                    if (entity.Children.ContainsKey(child.Identifier))
                    {
                        entity.Children.Remove(child.Identifier);
                    }
                    if (entity.ChildrenArePositions)
                    {
                        _sheetAccess.DeleteRange(((Position)child).Row);
                    }
                    else
                    {
                        int rowNumber = child.RowNumber;
                        _sheetAccess.DeleteRows(rowNumber - 1, rowNumber);
                    }
                }
                entity.ChildrenToDelete = new List<IChildEntity>();
            }
        }

        private void WritePositions(GroupingEntity entity, List<FXRateDTO> rates, Book book, Fund fund, bool updateExisting, bool forceRefresh)
        {
            Position previous = null;
            var orderedPositions = entity.Children.Values.OrderBy(a => a.Ordering).ToList();
            var updateSums = false;
            foreach (Position position in orderedPositions)
            {
                WritePosition(previous, position, entity, rates, book, fund, updateExisting, forceRefresh, ref updateSums);
                previous = position;
            }
            RemoveChildrenToBeDeleted(entity, updateExisting);
            if (updateSums || forceRefresh)
            {
                _sheetAccess.UpdateSums(entity);
            }
        }

       
       

        private void WritePosition(Position previousPosition,Position position,GroupingEntity parent,List<FXRateDTO> rates,
            Book book,Fund fund, bool updateExisting,bool forceRefresh, ref bool updateSums)
        {
            if (position.Row != null)
            {
                if (updateExisting || forceRefresh)
                {
                    if (position.InstrumentTypeId == InstrumentTypeIds.FX)
                    {
                        EnhanceFXPosition(position, rates);
                    }
                    _sheetAccess.WritePosition(position, book, fund, forceRefresh);
                }
            }
            else
            {
                updateSums = true;
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
