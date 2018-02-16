using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class Location
    {
        public Location(int? row, string ticker, string name, decimal netPosition, int? tickerTypeId, decimal? odeyPrice, string currency, decimal priceDivisor)
        {
            Row = row;
            Ticker = ticker;
            Name = name;
            _originalNetPosition = netPosition;
            NetPosition = netPosition;     
            OdeyPrice = odeyPrice;
            TickerTypeId = tickerTypeId;
            Currency = currency;
            PriceDivisor = priceDivisor;
        }

        private decimal _originalNetPosition;


        public int? Row { get; set; }

        public string Ticker { get; set; }

        public string Name { get; set; }

        public string Currency { get; set; }

        public decimal NetPosition { get; set; }

        public int? TickerTypeId { get; set; }

        public decimal? OdeyPrice { get; set; }

        public bool QuantityHasChanged { get { return _originalNetPosition != NetPosition; } }

        public decimal PriceDivisor { get; set; }

    }
}
