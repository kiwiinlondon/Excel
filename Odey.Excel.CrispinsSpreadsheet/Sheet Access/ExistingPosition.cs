using XL=Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class ExistingPositionDTO
    {
        public ExistingPositionDTO(int? instrumentMarketId,string ticker,string name, XL.Range row )
        {
            Identifier = new Identifier(instrumentMarketId, ticker);           
            Name = name;
            Row = row;
        }


        public Identifier Identifier { get; private set; }
        public string Name { get; set; }

        public XL.Range Row { get; set; }
    }
}
