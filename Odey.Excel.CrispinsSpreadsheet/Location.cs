using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class Location
    {
        public Location(int? row, string ticker, string name, decimal netPosition)
        {
            Row = row;
            Ticker = ticker;
            Name = name;
            _originalNetPosition = netPosition;
            NetPosition = netPosition;
        }

        private decimal _originalNetPosition;


        public int? Row { get; set; }

        public string Ticker { get; set; }

        public string Name { get; set; }

        private decimal _netPosition;
        public decimal NetPosition { get; set; }

        public bool QuantityHasChanged { get { return _originalNetPosition != NetPosition; } }

    }
}
