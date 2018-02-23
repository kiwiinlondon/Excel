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


            
        public void Match(int fundId, DateTime referenceDate)
        {
            _sheetAccess.DisableCalculations();
            decimal nav;
            decimal previousNav;
            DateTime previousReferenceDate;
            var dTOs = _dataAccess.Get(fundId, referenceDate, out nav, out previousNav, out previousReferenceDate);

            _sheetAccess.WriteDates(previousReferenceDate, referenceDate);
            _sheetAccess.WriteNAVs(previousNav,nav);
            Fund fund = _sheetAccess.GetFund("OEI");
            AddDTOsToFund(dTOs, fund);
            WriteGroupingEntity(null, fund, null);
            _sheetAccess.EnableCalculations();
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

        private void WriteGroupingEntity(GroupingEntity previous, GroupingEntity entity, GroupingEntity parentEntity)
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
                    WriteGroupingEntity((GroupingEntity)previousChildEntity, (GroupingEntity)childEntity, entity);
                }
                else
                {
                    WritePosition((Position)previousChildEntity, (Position)childEntity, entity);
                }
                previousChildEntity = childEntity;
            }
            UpdateTotals(entity, previousChildEntity);
        }

        private void UpdateTotals(GroupingEntity entity, IChildEntity previousChildEntity)
        {
            if (previousChildEntity is Position)
            {
                
                _sheetAccess.UpdateSums(entity, (Position)previousChildEntity);
            }
            else
            {
                _sheetAccess.UpdateTotalsOnTotalRow(entity);
            }
        }

        private void WritePosition(Position previousPosition,Position position,GroupingEntity parent)
        {
            if (position.Row != null)
            {
                _sheetAccess.UpdatePosition(true, position);
            }
            else
            {
                _sheetAccess.AddPosition(previousPosition, position, parent);
            }
        }


        //private void WriteBook(Book previousBook, Book book, Fund fund)
        //{
        //    if (book.TotalRow == null)
        //    {
        //        _sheetAccess.AddTotalRow(previousBook, book, fund);
        //    }
        //    AssetClass previousAssetClass = null;
        //    foreach (AssetClass assetClass in book.Children.Values.OrderByDescending(a => a.Name))
        //    {
        //        WriteAssetClass(previousAssetClass,assetClass, book);
        //        previousAssetClass = assetClass;
        //    }
        //    _sheetAccess.UpdateTotalsOnTotalRow(book);
        //}



        //private void WriteAssetClass(AssetClass previousAssetClass, AssetClass assetClass,Book book)
        //{
        //    if (assetClass.TotalRow == null)
        //    {
        //        _sheetAccess.AddTotalRow(previousAssetClass, assetClass, book);
        //    }

        //    Country previousCountry = null;
        //    foreach (var child in assetClass.Children.Values.OrderByDescending(a=>a.Name))
        //    {
        //        var country = (Country)child;
        //        WriteCountry(previousCountry, country, assetClass);
        //        previousCountry = country;
        //    }
        //    _sheetAccess.UpdateTotalsOnTotalRow(assetClass);

        //}

        //private void WriteCountry(Country previousCountry, Country country, AssetClass assetClass)
        //{
            
        //    if (country.TotalRow==null)
        //    {
        //        _sheetAccess.AddTotalRow(previousCountry, country, assetClass);
        //    }
        //    Position previous = null;
        //    foreach (var child in country.Children.Values.OrderByDescending(a=>a.Name))
        //    {
        //        Position position = (Position)child;
        //        if (position.Row!=null)
        //        {
        //            _sheetAccess.UpdatePosition(true, position);
        //        }
        //        else
        //        {
        //            _sheetAccess.AddPosition(previous, position, country);
        //        }
        //        previous = position;
        //    }
            
        //    _sheetAccess.UpdateSums(country, previous);
        //}

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
            position.Name = dto.Name;
            position.PriceDivisor = dto.PriceDivisor;
            position.TickerTypeId = dto.TickerTypeId;
            position.OdeyPreviousPreviousPrice = dto.PreviousPreviousPrice;
            position.OdeyPreviousPrice = dto.PreviousPrice;
            position.OdeyCurrentPrice = dto.CurrentPrice;

        }
    }
}
