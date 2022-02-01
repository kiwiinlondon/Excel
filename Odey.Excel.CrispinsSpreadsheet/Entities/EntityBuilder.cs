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
        private WorkbookAccess _workBookAccess;
        public static readonly string EquityLabel = "Equity";
        public static readonly string MacroLabel = "Macro";
        public static readonly string FXLabel = "FX";
        public static readonly string HedgeLabel = "Unhedged";

        public EntityBuilder(DataAccess dataAccess,WorkbookAccess sheetAccess)
        {
            _dataAccess = dataAccess;
            _workBookAccess = sheetAccess;
        }

        private EntityTypes GetFundChildType(FundIds fundId)
        {
            switch (fundId)
            {
                
                case FundIds.OEIMAC:
                case FundIds.OEIMACGBPBSHARECLASS:
                case FundIds.OEIMACGBPBMSHARECLASS:
                case FundIds.BEST:
                case FundIds.OBID:   
                //case FundIds.ODIF:
                case FundIds.GILT:
                    return EntityTypes.Position;

                case FundIds.FDXC:
                case FundIds.OPUS:
                case FundIds.OPE:                
                    return EntityTypes.Country;
                
                case FundIds.SWAN:
                case FundIds.OEI:
                    return EntityTypes.AssetClass;
                default:
                    throw new ApplicationException($"Unknown Fund {fundId}");
            }
        }




        public Fund GetFund(FundIds fundId,bool isPrimary)
        {
            Fund fund = _dataAccess.GetFund(fundId, GetFundChildType(fundId), isPrimary);
            
            fund.ChildrenAreDeleteable = isPrimary && (fundId != FundIds.OEI);
            switch (fund.ChildEntityType)
            {
                //case EntityTypes.Book:
                //    AddBooks(fund);
                //    break;
                case EntityTypes.AssetClass:
                    AddAssetClasses(fund);
                    break;
                    
            }           
            return fund;
        }

        





        private void AddAssetClasses(GroupingEntity parent)
        {
            AddAssetClass(parent, EquityLabel, 1, true);
            AddAssetClass(parent, MacroLabel, 2, true);
            AddAssetClass(parent, FXLabel, 3, false);
            AddAssetClass(parent, HedgeLabel, 4, true);

        }

        

        private void AddAssetClass(GroupingEntity parent, string assetClassLabel, int ordering, bool childrenAreDeletable)
        {
            EntityTypes childEntityType = EntityTypes.Position;
            if (assetClassLabel == EquityLabel)
            {
                childEntityType = EntityTypes.Country;
            }
            var assetClass = new AssetClass(parent, assetClassLabel, childEntityType, ordering);
            assetClass.ChildrenAreDeleteable = childrenAreDeletable;
            parent.Children.Add(assetClass.Identifier, assetClass);
        }


        

        

        public void AddPortfolio(Fund fund)
        {

            List<PortfolioDTO> portfolio = _dataAccess.GetPortfolio(fund);
     
            foreach (PortfolioDTO position in portfolio)
            {
                if (position.Instrument.InstrumentTypeId == InstrumentTypeIds.FX && position.Instrument.AssetClass == HedgeLabel)
                {
                    if (position.Instrument.Identifier.Code.Contains("ARS"))
                    {
                        int ii = 0;
                    }
                    int i = 0;
                }
                GroupingEntity parentForPositions = GetChildEntityWithPositions(fund, position,fund);                
                AddDTOToParent(parentForPositions, position);
            }
        }

        private GroupingEntity GetChildEntityWithPositions(GroupingEntity parentEntity, PortfolioDTO position,Fund fund)
        {
            GroupingEntity child = null;
            switch (parentEntity.ChildEntityType)
            {
                case EntityTypes.Position:
                    return parentEntity;
                //case EntityTypes.Book:
                //    child = (GroupingEntity)parentEntity.Children[new Identifier(null, position.Book)];
                //    break;
                case EntityTypes.AssetClass:
                    child = (GroupingEntity)parentEntity.Children[new Identifier(null, position.Instrument.AssetClass)];
                    break;
                case EntityTypes.Country:
                    child = GetCountry((GroupingEntity)parentEntity, position.Instrument.ExchangeCountryIsoCode, position.Instrument.ExchangeCountryName,fund);
                    break;
                default:
                    throw new ApplicationException($"Unknown Child Enity Type {parentEntity.ChildEntityType}");
            }
            return GetChildEntityWithPositions(child, position,fund);
        }

        public void AddExistingPortfolio(Fund fund,List<Position> toBeUpdatedFromDatabase)
        {
            if (fund.Name == "OEIMAC")
            {
                int i = 0;
            }
            List<ExistingGroupDTO> existingGroups = fund.WorksheetAccess.GetExisting(fund);
            foreach(var existingGroup in existingGroups)
            {
                GroupingEntity entity = fund;

                if (!string.IsNullOrWhiteSpace(existingGroup.AssetClassCode))
                {
                    entity = GetEntity(entity, existingGroup.AssetClassCode);
                }

                if (!string.IsNullOrWhiteSpace(existingGroup.CountryCode))
                {
                    entity = GetCountry(entity, existingGroup.CountryCode, existingGroup.Name,fund);
                }
                entity.TotalRow = new Row(entity.RowType, existingGroup.TotalRow);
                entity.ControlString = existingGroup.ControlString;
                AddExistingPositionsToParent(fund.WorksheetAccess, entity, existingGroup, toBeUpdatedFromDatabase);
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


        private Country GetCountry(GroupingEntity parentEntity, string code, string name,Fund fund)
        {
            Identifier identifier = new Identifier(null, code);
            if (parentEntity.Children.ContainsKey(identifier))
            {
                return (Country)parentEntity.Children[identifier];
            }
            else
            {
                if (code == "GB")
                {
                    name = "United Kingdom";
                }
                Country entity = new Country(parentEntity, code, name);
                entity.ChildrenAreDeleteable = fund.ChildrenAreDeleteable;
                parentEntity.Children.Add(identifier, entity);
                return entity;
            }
        }

        private void AddExistingPositionsToParent(WorksheetAccess worksheetAccess, GroupingEntity parent,ExistingGroupDTO existingGroup, List<Position> toBeUpdatedFromDatabase)
        {
            if (existingGroup.Positions != null && existingGroup.Positions.Count>0)
            {
                if (parent.ChildEntityType != EntityTypes.Position)
                {
                    throw new ApplicationException($"Not expecting positions on parent entity {parent}");
                }
                foreach (var existingPosition in existingGroup.Positions)
                {
                    AddExistingPositionToParent(worksheetAccess, parent, existingPosition, toBeUpdatedFromDatabase);
                }
            }
            

        }

        private void AddExistingPositionToParent(WorksheetAccess worksheetAccess, GroupingEntity parent,ExistingPositionDTO existingPosition,List<Position> toBeUpdatedFromDatabase)
        {            
            Position position;
            if (parent.Children.ContainsKey(existingPosition.Identifier))
            {
                position = (Position)parent.Children[existingPosition.Identifier];
                if (position.InstrumentTypeId == InstrumentTypeIds.FX)
                {
                    position.Name = worksheetAccess.GetNameFromRow(existingPosition.Row);
                }
            }
            else
            {
                position = worksheetAccess.BuildPosition(existingPosition);
                parent.Children.Add(existingPosition.Identifier, position);
                if (position.InstrumentTypeId == InstrumentTypeIds.DeleteableDerivative || position.InstrumentTypeId == InstrumentTypeIds.PrivatePlacement || (parent.ChildrenAreDeleteable && position.InstrumentTypeId == InstrumentTypeIds.Normal))
                {
                    parent.ChildrenToDelete.Add(position);
                }
                else if (position.InstrumentTypeId != InstrumentTypeIds.FX)
                {
                    toBeUpdatedFromDatabase.Add(position);
                }

            }
            position.Row = new Row(position.RowType, existingPosition.Row);
        }

        private void AddDTOToParent(GroupingEntity parent, PortfolioDTO dto)
        {
            if (dto.Instrument.Name == "VAL 6 7/8 08/15/20")
            {
                int i = 0;
            }
            Position position = new Position(dto.Instrument.Identifier, dto.Instrument.Name, dto.Instrument.PriceDivisor, dto.Instrument.InstrumentTypeId, dto.Instrument.InvertPNL);
            parent.Children.Add(dto.Instrument.Identifier, position);

            position.PreviousNetPosition = dto.PreviousNetPosition;
            position.NetPosition = dto.CurrentNetPosition;
            position.Currency = dto.Instrument.Currency;            
            position.PriceDivisor = dto.Instrument.PriceDivisor;
            position.InstrumentTypeId = dto.Instrument.InstrumentTypeId;
            position.OdeyPreviousPreviousPrice = dto.PreviousPreviousPrice;
            position.OdeyPreviousPrice = dto.PreviousPrice;
            position.OdeyCurrentPrice = dto.CurrentPrice;
            position.OdeyCurrentPriceIsManual = dto.CurrentPriceIsManual;
            position.OdeyPreviousPriceIsManual = dto.PreviousPriceIsManual;
            position.OdeyPreviousPreviousPriceIsManual = dto.PreviousPreviousPriceIsManual;

        }

        private GroupingEntity GetWhereChildrenAreCountry(GroupingEntity groupingEntity)
        {
            if (groupingEntity.ChildEntityType == EntityTypes.Country)
            {
                return groupingEntity;
            }
            else if (groupingEntity.ChildEntityType == EntityTypes.AssetClass)
            {
                return GetWhereChildrenAreCountry((AssetClass)groupingEntity.Children[new Identifier(null, EntityBuilder.EquityLabel)]);
            }
            //else if (groupingEntity.ChildEntityType == EntityTypes.Book)
            //{
            //    return GetWhereChildrenAreCountry((Book)groupingEntity.Children.Values.FirstOrDefault(a => ((Book)a).IsPrimary));
            //}
            else
            { 
                throw new ApplicationException("Unknown way to find country");
            }
        }

        public GroupingEntity AddInstrument(Fund fund, InstrumentDTO instrument)
        {

            GroupingEntity countryParent = GetWhereChildrenAreCountry(fund);

            Country country = GetCountry(countryParent, instrument.ExchangeCountryIsoCode, instrument.ExchangeCountryName,fund);
            var position = new Position(instrument.Identifier, instrument.Name, instrument.PriceDivisor, instrument.InstrumentTypeId,instrument.InvertPNL);
            country.Children.Add(position.Identifier, position);
            
            return country;
        }
    }        
}
