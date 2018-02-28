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
        private DataAccess _dataAccess;
        private SheetAccess  _sheetAccess;
        public Matcher (DataAccess dataAccess,SheetAccess sheetAccess)
        {
            _dataAccess = dataAccess;
            _sheetAccess = sheetAccess;
        }



        public void Match(int fundId)
        {
            _sheetAccess.DisableCalculations();


            var rates = _dataAccess.GetFXRates();
            _sheetAccess.WriteDates(_dataAccess.PreviousReferenceDate, _dataAccess.ReferenceDate);
            Fund previousFund = MatchFund(null, FundIds.OEI,  rates);
            previousFund = MatchFund(previousFund, FundIds.OEIMAC, rates);
            previousFund = MatchFund(previousFund, FundIds.OEIMACGBPBSHARECLASS, rates);
            previousFund = MatchFund(previousFund, FundIds.OEIMACGBPBMSHARECLASS, rates);

            _sheetAccess.EnableCalculations();
        }

        private Fund MatchFund(Fund previousFund,FundIds fundId,List<FXRateDTO> rates)
        {
            FundDTO fundStructure = _dataAccess.GetFund(fundId);
            var dTOs = _dataAccess.Get(fundStructure);
            Fund fund = _sheetAccess.GetFund(previousFund, fundStructure);            
            AddDTOsToFund(dTOs, fund);
            AddNavsToGroupings(fundStructure, fund);
            WriteGroupingEntity(null, fund, null,  rates, null, fund);
            return fund;
        }

        private void AddNavsToGroupings(FundDTO fundStructure, Fund fund)
        {
            fund.Currency = fundStructure.Currency;
            fund.Nav = fundStructure.CurrentNav;
            fund.PreviousNav = fundStructure.PreviousNav;
            foreach (var child in fund.Children.Values)
            {
                if (child is Book)
                {
                    Book book = (Book)child;
                    var bookStructure = fundStructure.Books[book.Name];
                    book.Nav = bookStructure.Nav;
                    book.PreviousNav = bookStructure.PreviousNav;
                }
            }
        }

        private GroupingEntity GetEntity(GroupingEntity parentEntity,string code, string name, GroupingEntityTypes groupingEntityType)
        {
            if (string.IsNullOrWhiteSpace(code))
            {
                return parentEntity;
            }
            else if (parentEntity.Children.ContainsKey(code))
            {
                return (GroupingEntity)parentEntity.Children[code];
            }
            else
            {
                GroupingEntity entity;
                switch (groupingEntityType)
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
                        throw new ApplicationException($"Unknown grouping type od {groupingEntityType}");
                }
                entity.Name = name;
                parentEntity.Children.Add(entity.Code, entity);
                return entity;
            }
        }


        private void AddDTOsToFund(List<PortfolioDTO> dTOs, Fund fund)
        {
            foreach (PortfolioDTO dTO in dTOs)
            {
                GroupingEntity parentEntity = fund;
                parentEntity = GetEntity(parentEntity, dTO.Book, dTO.Book, GroupingEntityTypes.Book);
                parentEntity = GetEntity(parentEntity, dTO.AssetClass, dTO.AssetClass, GroupingEntityTypes.AssetClass);
                parentEntity = GetEntity(parentEntity, dTO.CountryIsoCode, dTO.CountryName, GroupingEntityTypes.Country);
                MatchDTOToTickerRow(parentEntity, dTO);
            }
        }

        private void WriteGroupingEntity(GroupingEntity previous, GroupingEntity entity, GroupingEntity parentEntity,List<FXRateDTO> rates, Book book, Fund fund)
        {
            if (entity.TotalRow == null)
            {
                _sheetAccess.AddTotalRow(previous, entity, parentEntity);
            }
            IChildEntity previousChildEntity = null;
            foreach (IChildEntity childEntity in entity.Children.Values.OrderByDescending(a => a.Name))
            {
                if (childEntity is GroupingEntity)
                {
                    if (childEntity is Book)
                    {
                        book = (Book)childEntity;
                    }
                    WriteGroupingEntity((GroupingEntity)previousChildEntity, (GroupingEntity)childEntity, entity, rates, book,fund);
                }
                else
                {
                    WritePosition((Position)previousChildEntity, (Position)childEntity, entity, rates,book, fund);
                }
                previousChildEntity = childEntity;
            }
            UpdateTotals(entity, previousChildEntity);
        }

        private void UpdateTotals(GroupingEntity entity, IChildEntity previousChildEntity)
        {
            if (previousChildEntity is Position)
            {
                _sheetAccess.UpdateSums(entity);
            }
            else
            {
                _sheetAccess.UpdateTotalsOnTotalRow(entity);
            }
            _sheetAccess.UpdateNavs(entity);
        }

        private void WritePosition(Position previousPosition,Position position,GroupingEntity parent,List<FXRateDTO> rates,
            Book book,Fund fund)
        {
            if (position.Row != null)
            {
                if (position.TickerTypeId == TickerTypeIds.FX)
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
            string currency1 = position.Ticker.Substring(0, 3);
            string currency2 = position.Ticker.Substring(3, 3);

            var matchingRates = rates.Where(a => (a.FromCurrency == currency1 || a.FromCurrency == currency2) && (a.ToCurrency == currency1 || a.ToCurrency == currency2));
            if (matchingRates.Count()!=1)
            {
                throw new ApplicationException($"Cannot find rate for ticker {position.Ticker}. Count = {matchingRates.Count()}");
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
     
        private void MatchDTOToTickerRow(GroupingEntity parent, PortfolioDTO dto)
        {

            Position position;
            if (!parent.Children.ContainsKey(dto.Ticker))
            {
                position = new Position(dto.Ticker, dto.Name, dto.PriceDivisor, dto.TickerTypeId,null);
                parent.Children.Add(position.Ticker, position);
            }
            else
            {
                position = (Position)parent.Children[dto.Ticker];
            }
            position.PreviousNetPosition = dto.PreviousNetPosition;
            position.NetPosition = dto.CurrentNetPosition;
            position.Currency = dto.Currency;
            if (position.TickerTypeId != TickerTypeIds.FX)
            {
                position.Name = dto.Name;
            }
            else 
            {
                int i = 0;
            }
            position.PriceDivisor = dto.PriceDivisor;
            position.TickerTypeId = dto.TickerTypeId;
            position.OdeyPreviousPreviousPrice = dto.PreviousPreviousPrice;
            position.OdeyPreviousPrice = dto.PreviousPrice;
            position.OdeyCurrentPrice = dto.CurrentPrice;

        }
    }
}
