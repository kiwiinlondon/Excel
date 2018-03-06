using Odey.Framework.Keeley.Entities.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class EntityBuilder
    {

        private DataAccess _dataAccess;
        private SheetAccess _sheetAccess;
        public static readonly string EquityLabel = "Equity";
        public static readonly string MacroLabel = "Macro";
        public static readonly string FXLabel = "FX";

        public EntityBuilder(DataAccess dataAccess,SheetAccess sheetAccess)
        {
            _dataAccess = dataAccess;
            _sheetAccess = sheetAccess;
        }

        public Fund GetFund(FundIds fundId)
        {
            Fund fund = _dataAccess.GetFund(fundId);
            fund.ChildrenArePositions = fundId != FundIds.OEI;
            if (!fund.ChildrenArePositions)
            {
                AddBooks(fund);
            }           
            return fund;
        }

        

        private void AddBooks(Fund fund)
        {
            var books = _dataAccess.GetBooks(fund);
            foreach (var book in books)
            {
                AddBook(fund, book);
            }

        }

        private void AddBook(Fund fund, Book book)
        {
            book.ChildrenArePositions = book.BookId != (int)BookIds.OEI;
            if (!book.ChildrenArePositions)
            {
                AddAssetClasses(book);
            }
            fund.Children.Add(book.Identifier, book);
        }

        private void AddAssetClasses(Book book)
        {
            if (!book.ChildrenArePositions)
            {
                AddAssetClass(book, EquityLabel,1);
                AddAssetClass(book, MacroLabel,2);
                AddAssetClass(book, FXLabel,3);
            }
        }

        private void AddAssetClass(Book book, string assetClassLabel,int ordering)
        {
            var assetClass = new AssetClass(book,assetClassLabel, assetClassLabel != EquityLabel, ordering);
            book.Children.Add(assetClass.Identifier, assetClass);
        }


        

        public void RemovePositions(GroupingEntity groupingEntity)
        {           
            if (groupingEntity.ChildrenArePositions)
            {
                groupingEntity.Children = new Dictionary<Identifier, IChildEntity>();
            }
            else
            {
                foreach (var child in groupingEntity.Children)
                {
                    RemovePositions((GroupingEntity)child.Value);
                }
            }
        }

        public void AddPortfolio(Fund fund)
        {
            bool includeHedging = fund.FundId == (int)FundIds.OEIMACGBPBSHARECLASS || fund.FundId == (int)FundIds.OEIMACGBPBMSHARECLASS;
            bool includeOnlyFX = fund.FundId != (int)FundIds.OEI;
            List<PortfolioDTO> portfolio = _dataAccess.GetPortfolio(fund, includeHedging, includeOnlyFX);
            foreach (PortfolioDTO position in portfolio)
            {
                GroupingEntity parentEntity = fund;
                if (!fund.ChildrenArePositions)
                {
                    parentEntity = (GroupingEntity)parentEntity.Children[new Identifier(null, position.Book)];
                    if (!parentEntity.ChildrenArePositions)
                    {
                        parentEntity = (GroupingEntity)parentEntity.Children[new Identifier(null,position.Instrument.AssetClass)];
                        if (!parentEntity.ChildrenArePositions)
                        {
                            parentEntity = GetCountry((AssetClass)parentEntity, position.Instrument.ExchangeCountryIsoCode, position.Instrument.ExchangeCountryName);
                        }
                    }
                }
                AddDTOToParent(parentEntity, position);
            }
        }

        

        public void AddExistingPortfolio(Fund fund)
        {
            List<ExistingGroupDTO> existingGroups = _sheetAccess.GetExisting(fund);
            foreach(var existingGroup in existingGroups)
            {
                GroupingEntity entity = fund;
                if (!string.IsNullOrWhiteSpace(existingGroup.BookCode))
                {
                    entity = GetEntity(entity, existingGroup.BookCode);
                }

                if (!string.IsNullOrWhiteSpace(existingGroup.AssetClassCode))
                {
                    entity = GetEntity(entity, existingGroup.AssetClassCode);
                }

                if (!string.IsNullOrWhiteSpace(existingGroup.CountryCode))
                {
                    entity = GetCountry((AssetClass)entity, existingGroup.CountryCode, existingGroup.Name);
                }
                entity.TotalRow = existingGroup.TotalRow;
                entity.ControlString = existingGroup.ControlString;
                AddExistingPositionsToParent(entity, existingGroup);
            }
        }

        private GroupingEntity GetEntity(GroupingEntity parent, string code)
        {
            Identifier identifier = new Identifier(null, code);
            if (parent.Children.ContainsKey(identifier))
            {
                return (GroupingEntity)parent.Children[identifier];
            }
            return null;
        }


        private GroupingEntity GetCountry(AssetClass parentEntity, string code, string name)
        {
            Identifier identifier = new Identifier(null, code);
            if (parentEntity.Children.ContainsKey(identifier))
            {
                return (GroupingEntity)parentEntity.Children[identifier];
            }
            else
            {
                if (code == "GB")
                {
                    name = "United Kingdom";
                }
                GroupingEntity entity = new Country(parentEntity, code, name);                        
                parentEntity.Children.Add(identifier, entity);
                return entity;
            }
        }

        private void AddExistingPositionsToParent(GroupingEntity parent,ExistingGroupDTO existingGroup)
        {
            if (existingGroup.Positions != null && existingGroup.Positions.Count>0)
            {
                if (!parent.ChildrenArePositions)
                {
                    throw new ApplicationException($"Not expecting positions on parent entity {parent}");
                }
                foreach (var existingPosition in existingGroup.Positions)
                {
                    AddExistingPositionToParent(parent, existingPosition);
                }
            }
            

        }

        private void AddExistingPositionToParent(GroupingEntity parent,ExistingPositionDTO existingPosition)
        {
            Position position;
            if (parent.Children.ContainsKey(existingPosition.Identifier))
            {
                position = (Position)parent.Children[existingPosition.Identifier];
                if (position.InstrumentTypeId == InstrumentTypeIds.FX)
                {
                    position.Name = existingPosition.Name;
                }
                position.Row = existingPosition.Row;
            }
            else
            {
                position = _sheetAccess.BuildPosition(existingPosition);
                parent.Children.Add(existingPosition.Identifier, position);
            }
        }

        private void AddDTOToParent(GroupingEntity parent, PortfolioDTO dto)
        {

            Position position = new Position(dto.Instrument.Identifier, dto.Instrument.Name, dto.Instrument.PriceDivisor, dto.Instrument.InstrumentTypeId, null);
            parent.Children.Add(dto.Instrument.Identifier, position);

            position.PreviousNetPosition = dto.PreviousNetPosition;
            position.NetPosition = dto.CurrentNetPosition;
            position.Currency = dto.Instrument.Currency;            
            position.PriceDivisor = dto.Instrument.PriceDivisor;
            position.InstrumentTypeId = dto.Instrument.InstrumentTypeId;
            position.OdeyPreviousPreviousPrice = dto.PreviousPreviousPrice;
            position.OdeyPreviousPrice = dto.PreviousPrice;
            position.OdeyCurrentPrice = dto.CurrentPrice;

        }
    }        
}
