using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class InstrumentDTO 
    {
        public InstrumentDTO()
        {
        }


        public InstrumentDTO (int? instrumentMarketId, string ticker, string name, string assetClass, string exchangeCountryIsoCode, string exchangeCountryName, decimal priceDivisor, InstrumentTypeIds instrumentTypeId, string currency,bool invertPNL,bool isInflationAdjusted)
        {
            InstrumentTypeId = instrumentTypeId;
            Identifier = new Identifier(instrumentMarketId, ticker);            
            Name = name;
            AssetClass = assetClass;
            ExchangeCountryIsoCode = exchangeCountryIsoCode;
            ExchangeCountryName = exchangeCountryName;
            PriceDivisor = priceDivisor;
            Currency = currency;
            InvertPNL = invertPNL;
            IsInflationAdjusted = isInflationAdjusted;
        }


        public Identifier Identifier { get; set; }


        public bool InvertPNL { get; set; }
        public bool IsInflationAdjusted { get; set; }

        public string Name { get; set; }

        public string AssetClass { get; set; }

        public string ExchangeCountryIsoCode { get; set; }
        public string ExchangeCountryName { get; set; }

        public decimal PriceDivisor { get; set; }

        public InstrumentTypeIds InstrumentTypeId { get; set; }

        public string Currency { get; set; }

        protected bool Equals(InstrumentDTO other)
        {
            return Identifier.Equals(other.Identifier) && AssetClass.Equals(other.AssetClass);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((InstrumentDTO)obj);
        }

        public static bool operator ==(InstrumentDTO left, InstrumentDTO right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(InstrumentDTO left, InstrumentDTO right)
        {
            return !Equals(left, right);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return Identifier.GetHashCode();
            }
        }

    }
}
