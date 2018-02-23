using XL = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class Position : IChildEntity
    {
        public Position(string ticker, string name, decimal priceDivisor,int? tickerTypeId, XL.Range row)
        {
            Row = row;
            Ticker = ticker;
            Name = name;
            TickerTypeId = tickerTypeId;
            PriceDivisor = priceDivisor;
        }


        public XL.Range Row { get; set; }

        public int RowNumber => Row.Row;

        public string Ticker { get; set; }

        public string Name { get; set; }

        public string Currency { get; set; }

        public decimal NetPosition { get; set; }

        public int? TickerTypeId { get; set; }

        public decimal? OdeyCurrentPrice { get; set; }

        public decimal? OdeyPreviousPrice { get; set; }

        public decimal? OdeyPreviousPreviousPrice { get; set; }

        public decimal PriceDivisor { get; set; }

        public decimal PreviousNetPosition { get; set; }

    }
}
