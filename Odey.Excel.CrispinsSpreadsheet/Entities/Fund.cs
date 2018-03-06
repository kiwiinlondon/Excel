using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XL = Microsoft.Office.Interop.Excel;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class Fund : GroupingEntity
    {
        public int FundId { get; private set; }

        public Fund(int fundId, string name, string currency,bool childrenArePositions) : base(null,name,name, childrenArePositions,fundId)
        {
            FundId = fundId;
            Currency = currency;
        }
        public string Currency { get; private set; }        

        public XL.Range Range { get; set; }

    }
}
