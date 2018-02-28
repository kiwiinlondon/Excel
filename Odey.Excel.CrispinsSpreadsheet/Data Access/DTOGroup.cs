using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class DTOGroup
    {
        public DTOGroup (string book, string assetClass, string countryIso,string countryName, TickerTypeIds tickerTypeId, string ticker,string currency, decimal priceDivisor)
        {
            Book = book;
            AssetClass = assetClass;
            CountryIso = countryIso;
            CountryName = countryName;
            TickerTypeId = tickerTypeId;
            Ticker = ticker;            
            Currency = currency;
            PriceDivisor = priceDivisor;
        }
        public string Book { get; set; }
        public string AssetClass { get; set; }
        public string CountryIso { get; set; }
        public string CountryName { get; set; }
        public string Ticker { get; set; }
        public string Currency { get; set; }
        public decimal PriceDivisor { get; set; }
        public TickerTypeIds TickerTypeId { get; set; }
        protected bool Equals(DTOGroup other)
        {
            return string.Equals(Book , other.Book) &&
                string.Equals(AssetClass, other.AssetClass) &&
                string.Equals(CountryIso, other.CountryIso) &&
                string.Equals(CountryName, other.CountryName) &&
                string.Equals(Ticker, other.Ticker) &&
                TickerTypeId == other.TickerTypeId &&
                string.Equals(Currency, other.Currency) &&
                PriceDivisor == other.PriceDivisor;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((DTOGroup)obj);
        }

        public static bool operator ==(DTOGroup left, DTOGroup right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(DTOGroup left, DTOGroup right)
        {
            return !Equals(left, right);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return Ticker.GetHashCode() ^ (Book==null ? 1 : Book.GetHashCode());
            }
        }

    }
}
