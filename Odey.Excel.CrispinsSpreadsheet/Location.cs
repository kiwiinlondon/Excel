using XL = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class Location
    {
        public Location(int? rowNumber, string ticker, string name, decimal previousNetPosition, decimal netPosition, int? tickerTypeId, 
            decimal? odeyPreviousPreviousPrice, decimal? odeyPreviousPrice, decimal? odeyCurrentPrice, string currency, decimal priceDivisor,XL.Range row)
        {
            Row = row;
            RowNumber = rowNumber;
            Ticker = ticker;
            Name = name;
            _originalNetPosition = netPosition;
            NetPosition = netPosition;     
            OdeyPreviousPreviousPrice = odeyPreviousPreviousPrice;
            OdeyPreviousPrice = odeyPreviousPrice;
            OdeyCurrentPrice = odeyCurrentPrice;
            TickerTypeId = tickerTypeId;
            Currency = currency;
            PriceDivisor = priceDivisor;
            PreviousNetPosition = previousNetPosition;
        }

        private decimal _originalNetPosition;

        public XL.Range Row { get; set; }
        public int? RowNumber { get; set; }

        public string Ticker { get; set; }

        public string Name { get; set; }

        public string Currency { get; set; }

        public decimal NetPosition { get; set; }

        public int? TickerTypeId { get; set; }

        public decimal? OdeyCurrentPrice { get; set; }

        public decimal? OdeyPreviousPrice { get; set; }

        public decimal? OdeyPreviousPreviousPrice { get; set; }

        public bool QuantityHasChanged { get { return _originalNetPosition != NetPosition; } }

        public decimal PriceDivisor { get; set; }

        public decimal PreviousNetPosition { get; set; }

    }
}
