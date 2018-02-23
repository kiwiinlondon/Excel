using Odey.Framework.Keeley.Entities;
using Odey.Framework.Keeley.Entities.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class DTOGroupBuilder
    {
        private static readonly DTOGroupBuilder instance = new DTOGroupBuilder();

        private DTOGroupBuilder()
        {

        }

        public static DTOGroupBuilder Instance
        {
            get
            {
                return instance;
            }
        }

        private static readonly int[] EquityAssetClassIds = { (int)DerivedAssetClassIds.Equity, (int)DerivedAssetClassIds.Bond };

        private static readonly string _equityLabel = "Equity";
        private static readonly string _macroLabel = "Macro";
        private static readonly string _fxLabel = "FX";

        private string GetAssetClass(Framework.Keeley.Entities.Position position)
        {
            if (!position.Book.IsPrimary)
            {
                return null;
            }
            if (EquityAssetClassIds.Contains(position.InstrumentMarket.Instrument.DerivedAssetClassId))
            {
                return _equityLabel;
            }
            else
            {
                return _macroLabel;
            }
        }

        private string GetCountryCode(Framework.Keeley.Entities.Position position,string assetClass)
        {
            if (assetClass == null || assetClass != _equityLabel)
            {
                return null;
            }
            return position.InstrumentMarket.Market.LegalEntity.Country.IsoCode;
        }

        private string GetCountryName(Framework.Keeley.Entities.Position position, string countryCode)
        {
            if (countryCode == null)
            {
                return null;
            }
            return position.InstrumentMarket.Market.LegalEntity.Country.Name;
        }

        public DTOGroup Get(Framework.Keeley.Entities.Position position)
        {
            int? tickerTypeId = GetTickerType(position.InstrumentMarket);
            string assetClass = GetAssetClass(position);
            string countryIsoCode = GetCountryCode(position, assetClass);
            return new DTOGroup(
                    position.Book.Name,
                    assetClass,
                    countryIsoCode,
                    GetCountryName(position, countryIsoCode),
                    position.InstrumentMarket.Name,
                    tickerTypeId,
                    GetTicker(position.InstrumentMarket, tickerTypeId),
                    position.InstrumentMarket.PriceCurrency.IsoCode,
                    position.InstrumentMarket.PriceDivisor);
        }

        private string GetTicker(InstrumentMarket instrumentMarket,int? tickerTypeId)
        {
            if (tickerTypeId.HasValue)
            {
                return instrumentMarket.InstrumentMarketID.ToString();
            }
            return instrumentMarket.BloombergTicker;
        }

        private static readonly int[] PrivateListingStatusIds = { (int)ListingStatusIds.Delisted, (int)ListingStatusIds.PrivatePlacement };

        private int? GetTickerType(InstrumentMarket instrumentMarket)
        {
            if (string.IsNullOrWhiteSpace(instrumentMarket.BloombergTicker) || instrumentMarket.BloombergTicker.StartsWith(".") || PrivateListingStatusIds.Contains(instrumentMarket.ListingStatusId))
            {
                return 1;
            }
            return null;
        }

    }
}
