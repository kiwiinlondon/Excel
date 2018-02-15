using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class CountryLocation
    {
        public string IsoCode { get; set; }

        public string Name { get; set; }

        public int? FirstRow { get; set; }

        public int? TotalRow { get; set; }

        public Dictionary<string,Location> TickerRows { get; set; } = new Dictionary<string, Location>();
        
        public List<PortfolioDTO> PortfolioDtos { get; set; } = new List<PortfolioDTO>();

        public override string ToString()
        {
            return $"{IsoCode}-{Name}";
        }

    }
}
