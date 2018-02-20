using XL=Microsoft.Office.Interop.Excel;
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

        public int? FirstRowNumber { get; set; }

        public int? TotalRowNumber { get; set; }

        public XL.Range TotalRow { get; set; }

        public Dictionary<string,Location> TickerRows { get; set; } = new Dictionary<string, Location>();
        
        public List<PortfolioDTO> PortfolioDtos { get; set; } = new List<PortfolioDTO>();

        public override string ToString()
        {
            return $"{IsoCode}-{Name}";
        }

    }
}
