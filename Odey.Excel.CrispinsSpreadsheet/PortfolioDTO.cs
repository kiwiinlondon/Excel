using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class PortfolioDTO
    {

        public PortfolioDTO (string name, string ticker, string currency,string countryIsoCode, string countryName,
             decimal previousNetPosition,decimal currentNetPosition, int? tickerTypeId, decimal? previousPreviousPrice, decimal? previousPrice, decimal? currentPrice, decimal priceDivisor)
        {
            Name = name;
            Ticker = ticker;
            CountryIsoCode = countryIsoCode;
            CountryName = countryName;
            PreviousNetPosition = previousNetPosition;
            CurrentNetPosition = currentNetPosition;
            TickerTypeId = tickerTypeId;
            PreviousPrice = 
            CurrentPrice = currentPrice;
            Currency = currency;
            PriceDivisor = priceDivisor;
        }

        public string Name { get; set; }

        public string Ticker { get; set; }

        public string CountryIsoCode { get; set; }

        public string CountryName { get; set; }

        public string Currency { get; set; }

        public decimal CurrentNetPosition { get; set; }
        public decimal PreviousNetPosition { get; set; }

        public int? TickerTypeId { get; set; }

        public decimal? CurrentPrice { get; set; }

        public decimal? PreviousPrice { get; set; }

        public decimal? PreviousPreviousPrice { get; set; }

        public decimal PriceDivisor { get; set; }
    }
}
