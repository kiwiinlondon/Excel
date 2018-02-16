using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class PortfolioDTO
    {

        public PortfolioDTO (string name, string ticker, string currency,string countryIsoCode, string countryName, decimal netPosition, int? tickerTypeId, decimal price, decimal priceDivisor)
        {
            Name = name;
            Ticker = ticker;
            CountryIsoCode = countryIsoCode;
            CountryName = countryName;
            NetPosition = netPosition;
            TickerTypeId = tickerTypeId;
            Price = price;
            Currency = currency;
            PriceDivisor = priceDivisor;
        }

        public string Name { get; set; }

        public string Ticker { get; set; }

        public string CountryIsoCode { get; set; }

        public string CountryName { get; set; }

        public string Currency { get; set; }

        public decimal NetPosition { get; set; }

        public int? TickerTypeId { get; set; }

        public decimal Price { get; set; }

        public decimal PriceDivisor { get; set; }
    }
}
