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

        private string GetAssetClass(Framework.Keeley.Entities.Position position, string book)
        {
            if (!position.Book.IsPrimary || book == null)
            {
                return null;
            }
            if (position.InstrumentMarket.Instrument.DerivedAssetClassId== (int)DerivedAssetClassIds.ForeignExchange)
            {
                return _fxLabel;
            }
            else if (EquityAssetClassIds.Contains(position.InstrumentMarket.Instrument.DerivedAssetClassId))
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

        private string GetBookName(Framework.Keeley.Entities.Position position, FundDTO fund)
        {
            if (fund.FundFXTreatmentId!= FundFXTreatmentIds.Normal)
            {
                return null;
            }
            return position.Book.Name;
        }

        public DTOGroup Get(Framework.Keeley.Entities.Position position, FundDTO fund)
        {
            var tickerTypeId = GetTickerType(position.InstrumentMarket);
            string book = GetBookName(position, fund);
            string assetClass = GetAssetClass(position,book);
            string countryIsoCode = GetCountryCode(position, assetClass);
            return new DTOGroup(
                    book,
                    assetClass,
                    countryIsoCode,
                    GetCountryName(position, countryIsoCode),
                    tickerTypeId,
                    GetTicker(position.InstrumentMarket, tickerTypeId),
                    position.InstrumentMarket.PriceCurrency.IsoCode,
                    position.InstrumentMarket.PriceDivisor);
        }

        private string GetTicker(InstrumentMarket instrumentMarket, TickerTypeIds tickerTypeId)
        {
            if (tickerTypeId == TickerTypeIds.PrivatePlacement)
            {
                return instrumentMarket.InstrumentMarketID.ToString();
            }
            return instrumentMarket.BloombergTicker;
        }

        private static readonly int[] PrivateListingStatusIds = { (int)ListingStatusIds.Delisted, (int)ListingStatusIds.PrivatePlacement };

        private static readonly int[] PrivateInstrumentMarketIds = { 24106 };//Cadiz Bond       

        private TickerTypeIds GetTickerType(InstrumentMarket instrumentMarket)
        {
            if (string.IsNullOrWhiteSpace(instrumentMarket.BloombergTicker) 
                || instrumentMarket.BloombergTicker.StartsWith(".") 
                || PrivateListingStatusIds.Contains(instrumentMarket.ListingStatusId)
                || PrivateInstrumentMarketIds.Contains(instrumentMarket.InstrumentMarketID))
            {
                return TickerTypeIds.PrivatePlacement;
            }
            return TickerTypeIds.Normal;
        }

        

        public DTOGroup GetFX(Framework.Keeley.Entities.Position position, FundDTO fund)
        {
            string currency1;
            string currency2;

            GetCurrencyPair(position.InstrumentMarket, out currency1, out currency2);
            string book = GetBookName(position, fund);
            return new DTOGroup(
                    book,
                    GetAssetClass(position, book),
                    null,
                    null,
                    TickerTypeIds.FX,
                    GetFXTicker(currency1, currency2),
                    GetFXCurrency(currency1, currency2),
                    1);
        }


        private void GetCurrencyPair(InstrumentMarket instrumentMarket, out string currency1, out string currency2)
        {
            currency1 = instrumentMarket.Name.Substring(0, 3);
            currency2 = instrumentMarket.Name.Substring(4, 3);
        }

        private string GetFXTicker(string currency1, string currency2)
        {
            return $"{currency1}{currency2} Curncy";
        }

        private string GetFXCurrency(string currency1, string currency2)
        {

            if (currency1 == "GBP" && currency2 == "USD")
            {
                return "GBP";
            }

            if (currency1 == "USD" || currency2 == "USD")
            {
                return "USD";
            }
            else if (currency1 == "EUR")
            {
                return currency2;
            }
            else if (currency2 == "EUR")
            {
                return currency1;
            }
            if (currency1 == "GBP" || currency2 == "GBP")
            {
                return "GBP";
            }
            else return currency1;
        }

        
    }
}
