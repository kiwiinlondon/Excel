using Odey.Framework.Keeley.Entities;
using Odey.Framework.Keeley.Entities.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class InstrumentBuilder
    {
        private static readonly InstrumentBuilder instance = new InstrumentBuilder();

        private InstrumentBuilder()
        {

        }

        public static InstrumentBuilder Instance
        {
            get
            {
                return instance;
            }
        }

        private static readonly int[] EquityAssetClassIds = { (int)DerivedAssetClassIds.Equity, (int)DerivedAssetClassIds.Bond };

        private string GetAssetClass(InstrumentMarket instrumentMarket)
        {
           
            if (instrumentMarket.Instrument.DerivedAssetClassId == (int)DerivedAssetClassIds.ForeignExchange)
            {
                return EntityBuilder.FXLabel;
            }
            else if (EquityAssetClassIds.Contains(instrumentMarket.Instrument.DerivedAssetClassId))
            {
                return EntityBuilder.EquityLabel;
            }
            else
            {
                return EntityBuilder.MacroLabel;
            }
        }


        private string GetCountryCode(InstrumentMarket instrumentMarket)
        {
            return instrumentMarket.Market.LegalEntity.Country.IsoCode;
        }


        private string GetCountryName(InstrumentMarket instrumentMarket)
        {
            return instrumentMarket.Market.LegalEntity.Country.Name;
        }

        private int GetInstrumentMarketId(InstrumentMarket instrumentMarket)
        {
            if (instrumentMarket.InstrumentClassIdAsEnum == InstrumentClassIds.ContractForDifference)
            {
                return instrumentMarket.UnderlyingInstrumentMarketId;
            }
            return instrumentMarket.InstrumentMarketID;
        }

        private string GetName(InstrumentMarket instrumentMarket)
        {
            if (instrumentMarket.InstrumentClassIdAsEnum == InstrumentClassIds.ContractForDifference)
            {
                return instrumentMarket.Name.Replace(" -CFD","");
            }
            return instrumentMarket.Name;
        }

        public InstrumentDTO Get(InstrumentMarket instrumentMarket)
        {
            var instrumentTypeId = GetInstrumentType(instrumentMarket);
            string assetClass = GetAssetClass(instrumentMarket);
            string countryIsoCode = GetCountryCode(instrumentMarket);
            return new InstrumentDTO(
                GetInstrumentMarketId(instrumentMarket),
                GetTicker(instrumentMarket, instrumentTypeId),
                GetName(instrumentMarket),
                    assetClass,
                    countryIsoCode,
                    GetCountryName(instrumentMarket),
                    instrumentMarket.PriceDivisor,
                    instrumentTypeId,                    
                    instrumentMarket.PriceCurrency.IsoCode,
                    false);
        }

        private string GetTicker(InstrumentMarket instrumentMarket, InstrumentTypeIds instrumentTypeId)
        {
            if (instrumentTypeId == InstrumentTypeIds.PrivatePlacement)
            {
                return null;
            }
            return instrumentMarket.BloombergTicker;
        }

        private static readonly int[] PrivateListingStatusIds = { (int)ListingStatusIds.Delisted, (int)ListingStatusIds.PrivatePlacement };

        private static readonly int[] PrivateInstrumentMarketIds = { 24106 };//Cadiz Bond       

        public InstrumentTypeIds GetInstrumentType(InstrumentMarket instrumentMarket)
        {
            if (string.IsNullOrWhiteSpace(instrumentMarket.BloombergTicker) 
                || instrumentMarket.BloombergTicker.StartsWith(".") 
                || PrivateListingStatusIds.Contains(instrumentMarket.ListingStatusId)
                || PrivateInstrumentMarketIds.Contains(instrumentMarket.InstrumentMarketID))
            {
                return InstrumentTypeIds.PrivatePlacement;
            }
            if ((instrumentMarket.ParentInstrumentClassIdAsEnum == ParentInstrumentClassIds.Funds && instrumentMarket.InstrumentClassIdAsEnum != InstrumentClassIds.ExchangeTradedFunds ) ||
                instrumentMarket.ParentInstrumentClassIdAsEnum == ParentInstrumentClassIds.FixedIncome || 
                instrumentMarket.ParentInstrumentClassIdAsEnum == ParentInstrumentClassIds.Option || 
                instrumentMarket.ParentInstrumentClassIdAsEnum == ParentInstrumentClassIds.Future)
            { 
                return InstrumentTypeIds.DeleteableDerivative;
            }
            return InstrumentTypeIds.Normal;
        }



        public InstrumentDTO GetFX(InstrumentMarket instrumentMarket)
        {
            string currency1;
            string currency2;

            GetCurrencyPair(instrumentMarket, out currency1, out currency2);
            string currency = GetFXCurrency(currency1, currency2);
            string ticker = GetFXTicker(currency1, currency2);

            bool invertPNL = currency != currency1;
            return new InstrumentDTO(
                    null,
                    ticker,
                    GetFXName(currency1, currency2),
                    GetAssetClass(instrumentMarket),
                    null,
                    null,
                    1,
                    InstrumentTypeIds.FX,
                    currency,
                    invertPNL
                    );
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

        private string GetFXName(string currency1, string currency2)
        {
            return $"{currency1}/{currency2}";
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
