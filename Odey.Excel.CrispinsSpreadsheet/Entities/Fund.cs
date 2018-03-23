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

        public Fund(int fundId, string name, string currency,bool childrenArePositions, bool isLongOnly, EntityTypes childEntityType, bool includeHedging, bool includeOnlyFX) : base(null,name,name, childEntityType, fundId)
        {
            FundId = fundId;
            Currency = currency;
            IsLongOnly = isLongOnly;
            IncludeHedging = includeHedging;
            IncludeOnlyFX = includeOnlyFX;
        }
        public string Currency { get; private set; }        

        public XL.Range Range { get; set; }

        public WorksheetAccess WorksheetAccess { get; set; }

        public bool IsLongOnly { get; set; }

        public bool IncludeHedging { get; set; }
        
        public bool IncludeOnlyFX { get; set; }
            


    }
}
