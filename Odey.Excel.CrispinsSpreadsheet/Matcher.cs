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
        private WorkbookAccess  _workbookAccess;
        private InstrumentRetriever _instrumentRetriever;
        public Matcher (EntityBuilder entityBuilder, DataAccess dataAccess, WorkbookAccess workbookAccess, InstrumentRetriever instrumentRetriever)
        {
            _entityBuilder = entityBuilder; 
            _dataAccess = dataAccess;
            _workbookAccess = workbookAccess;
            _instrumentRetriever = instrumentRetriever;
        }



        List<Fund> _funds = new List<Fund>();
        public void Match(bool refreshFormulas)
        {
            _workbookAccess.DisableCalculations();

            var rates = _dataAccess.GetFXRates();

            var funds = new List<Fund>();
            funds.AddRange(MatchFundSet(FundIds.OEI, new FundIds[] { FundIds.OEIMAC, FundIds.OEIMACGBPBSHARECLASS, FundIds.OEIMACGBPBMSHARECLASS }, rates, refreshFormulas));
            //MatchFundSet(FundIds.ODIF, null, rates, refreshFormulas);
            funds.AddRange(MatchFundSet(FundIds.SWAN, null, rates, refreshFormulas));
            funds.AddRange(MatchFundSet(FundIds.GILT, null, rates, refreshFormulas));
            funds.AddRange(MatchFundSet(FundIds.OPUS, null, rates, refreshFormulas));
            funds.AddRange(MatchFundSet(FundIds.OPE, null, rates, refreshFormulas));
            funds.AddRange(MatchFundSet(FundIds.FDXC, null, rates, refreshFormulas));

            var fxWorksheet = _workbookAccess.GetFXWorksheetAccess();
            fxWorksheet.Write(funds, new FundIds[] { FundIds.OEI, FundIds.OEIMAC, FundIds.SWAN });

            _workbookAccess.EnableCalculations();

            for (int i = _funds.Count - 1; i >= 0; i--)
            {
                var fund = _funds[i];

                fund.WorksheetAccess.FinaliseFormatting(fund.LastFund);
            }

            _workbookAccess.Save();

        }

        private List<Fund> MatchFundSet(FundIds primaryFundId,FundIds[] additionalFundIds,List<FXRateDTO> rates, bool refreshFormulas)
        {
            List<Fund> fundsToReturn = new List<Fund>();


            var primaryFund = BuildFund(primaryFundId, null,null);
            _funds.Add(primaryFund);
            fundsToReturn.Add(primaryFund);
            primaryFund.AdditionalFunds = BuildAdditionalFunds(primaryFund,additionalFundIds);

            primaryFund.WorksheetAccess.WriteDates(_dataAccess.PreviousReferenceDate, _dataAccess.ReferenceDate);

            MatchFund(primaryFund, rates, refreshFormulas);
            primaryFund.LastFund = primaryFund;
            foreach (Fund fund in primaryFund.AdditionalFunds.OrderBy(a => a.Ordering))
            {
                fundsToReturn.Add(fund);
                MatchFund(fund, rates, refreshFormulas);
                primaryFund.LastFund = fund;
            }

            return fundsToReturn;
        }


        public string AddTicker(string ticker)
        {
            var m = AddTicker(ticker, null);
            _workbookAccess.Save();
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
              //  Book book = (Book)fund.Children.Values.FirstOrDefault(a => (((Book)a).BookId == (int)BookIds.OEI));
                if (TickerAlreadyExists(ticker, fund))
                {
                    message = "Ticker Already Exists";
                }
                else
                {
                    var instrument = _instrumentRetriever.Get(ticker, out message);
                    if (instrument != null)
                    {
                        var rates = _dataAccess.GetFXRates();
                        var country = _entityBuilder.AddInstrument(fund, instrument);
                        WriteGroupingEntity(fund, rates, fund, false, false);
                        Position position = (Position)country.Children[instrument.Identifier];
                        message = $"Success. {ticker} added to Country {instrument.ExchangeCountryName} at row {position.RowNumber}";
                    }
                }
            }

            return message;
        }

        private bool TickerAlreadyExists(string ticker, GroupingEntity groupingEntity)
        {
            if (groupingEntity.ChildEntityType == EntityTypes.Position)
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
            _workbookAccess.DisableCalculations();
            var sheetAccess = _workbookAccess.GetBulkLoadTickerWorksheetAccess();
            var tickers = sheetAccess.GetBulkTickers();
            Fund fund = BuildOEIFromSheet();
            foreach(var ticker in tickers)
            {
                AddTicker(ticker, fund);
            }
            _workbookAccess.EnableCalculations();
            _workbookAccess.Save();
            return "Success";
        }



        private List<Fund> BuildAdditionalFunds(Fund primaryFund, FundIds[] fundIds)
        {
            var funds = new List<Fund>();
            if (fundIds != null)
            {
                Fund previous = primaryFund;
                foreach (FundIds fundId in fundIds)
                {
                    var fund = BuildFund(fundId, previous, primaryFund.WorksheetAccess);
                    funds.Add( fund);
                    previous = fund;
                }
            }
            return funds;
        }
        
        private Fund BuildFund(FundIds fundId,Fund previous, WorksheetAccess worksheetAccess)
        {
            Fund fund = _entityBuilder.GetFund(fundId,previous == null);
            fund.Previous = previous;
            if (worksheetAccess == null)
            {
                fund.WorksheetAccess = _workbookAccess.GetWorksheetAccess(fund);
            }
            else
            {
                fund.WorksheetAccess = worksheetAccess;
            }
            fund.WorksheetAccess.AddFundRange(fund);
            return fund;
        }

        private Fund BuildOEIFromSheet()
        {
            Fund fund = BuildFund(FundIds.OEI,null,null);
            
            List<Position> positionsToBeUpdatedFromDatabase = new List<Position>();
            _entityBuilder.AddExistingPortfolio(fund, positionsToBeUpdatedFromDatabase);
            return fund;
        }

        private void MatchFund(Fund fund, List<FXRateDTO> rates, bool refreshFormulas)
        {
            if (fund.FundId == 5513)
            {
                int i = 0;
            }
            _entityBuilder.AddPortfolio(fund);

            List<Position> positionsToBeUpdatedFromDatabase = new List<Position>();
            _entityBuilder.AddExistingPortfolio(fund, positionsToBeUpdatedFromDatabase);
            UpdatePositionsFromDatabase(positionsToBeUpdatedFromDatabase);

            WriteGroupingEntity(fund, rates, fund, true, refreshFormulas);
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

        private void WriteGroupingEntity(GroupingEntity entity, List<FXRateDTO> rates, Fund fund, bool updateExistingPositions, bool forceRefresh)
        {
            ChangeGroupVisiblity(entity, fund.WorksheetAccess, false);
            if (entity.TotalRow == null)
            {
                fund.WorksheetAccess.AddTotalRow(entity);
            }
            if (entity.ChildEntityType == EntityTypes.Position)
            {
                WritePositions(entity, rates, fund, updateExistingPositions, forceRefresh);               
            }
            else 
            {

                var updateTotal = false;
                GroupingEntity previous = null;
                foreach (IChildEntity childEntity in entity.Children.Values.OrderBy(a => a.Ordering))
                {       
                    GroupingEntity groupingEntity = (GroupingEntity)childEntity;
                    if (groupingEntity.TotalRow==null)
                    {
                        updateTotal = true;
                    }
                    groupingEntity.Previous = previous;
                    WriteGroupingEntity(groupingEntity, rates, fund, updateExistingPositions, forceRefresh);
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

                RemoveChildrenToBeDeleted(entity, updateExistingPositions,fund.WorksheetAccess);
                entity.Previous = previous;
                if (updateTotal || forceRefresh)
                {
                    fund.WorksheetAccess.UpdateTotalsOnTotalRow(entity);
                }

            }
            
            if (updateExistingPositions)
            {
                fund.WorksheetAccess.UpdateNavs(entity);
            }
            ChangeGroupVisiblity(entity,fund.WorksheetAccess,true);
        }

        private void ChangeGroupVisiblity(GroupingEntity entity,WorksheetAccess worksheetAccess,bool hidden)
        {
            if (entity.ChildrenAreHidden)
            {
                worksheetAccess.ChangeRowVisibilty(entity.Previous.TotalRow.RowNumber + 1, entity.TotalRow.RowNumber - 1, hidden);
            }
        }

        private void RemoveChildrenToBeDeleted(GroupingEntity entity, bool updateExistingPositions,WorksheetAccess worksheetAccess)
        {
            if (updateExistingPositions)
            {
                foreach (var child in entity.ChildrenToDelete)
                {
                    if (entity.Children.ContainsKey(child.Identifier))
                    {
                        entity.Children.Remove(child.Identifier);
                    }
                    if (entity.ChildEntityType == EntityTypes.Position)
                    {
                        worksheetAccess.DeleteRange(((Position)child).Row.Range);
                    }
                    else
                    {
                        int rowNumber = child.RowNumber;
                        worksheetAccess.DeleteRows(rowNumber - 1, rowNumber);
                    }
                }
                entity.ChildrenToDelete = new List<IChildEntity>();
            }
        }

        private void WritePositions( GroupingEntity entity, List<FXRateDTO> rates, Fund fund, bool updateExisting, bool forceRefresh)
        {
            Position previous = null;
            var orderedPositions = entity.Children.Values.OrderBy(a => a.Ordering).ToList();
            var updateSums = false;
            foreach (Position position in orderedPositions)
            {
                WritePosition(previous, position, entity, rates, fund, updateExisting, forceRefresh, ref updateSums);
                previous = position;
            }
            RemoveChildrenToBeDeleted(entity, updateExisting, fund.WorksheetAccess);
            if (updateSums || forceRefresh)
            {
                fund.WorksheetAccess.UpdateSums(entity);
            }
        }

       
       

        private void WritePosition(Position previousPosition,Position position,GroupingEntity parent,List<FXRateDTO> rates,Fund fund, bool updateExisting,bool forceRefresh, ref bool updateSums)
        {
            if (position.InstrumentTypeId == InstrumentTypeIds.FX)
            {
                EnhanceFXPosition(position, rates);
            }
            if (position.Row != null)
            {
                if (updateExisting || forceRefresh)
                {                    
                    fund.WorksheetAccess.WritePosition(position, fund, forceRefresh);
                }
            }
            else
            {
                updateSums = true;
                fund.WorksheetAccess.AddPosition(previousPosition, position, parent, fund);
            }
        }


        private void EnhanceFXPosition(Position position, List<FXRateDTO> rates)
        {
            string ticker = position.Identifier.Code;
            if (ticker == ".AREQIMP G Index")
            {
                ticker = "EURARS";
            }
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
