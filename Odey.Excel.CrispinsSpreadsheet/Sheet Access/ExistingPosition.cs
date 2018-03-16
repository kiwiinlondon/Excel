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
        public ExistingPositionDTO(int? instrumentMarketId,string ticker, XL.Range row )
        {
            Identifier = new Identifier(instrumentMarketId, ticker);           
            Row = row;
        }


        public Identifier Identifier { get; private set; }

        public XL.Range Row { get; set; }
    }
}
