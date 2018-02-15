using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class PortfolioDTO
    {

        public string Name { get; set; }

        public string Ticker { get; set; }

        public string CountryIso { get; set; }

        public string CountryName { get; set; }

        public decimal NetPosition { get; set; }
    }
}
