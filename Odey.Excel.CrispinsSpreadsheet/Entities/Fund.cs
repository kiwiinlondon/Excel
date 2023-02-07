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

        public Fund(int fundId, string name, string currency, int currencyId, bool childrenArePositions, bool isLongOnly, EntityTypes childEntityType, bool includeHedging, bool includeOnlyFX,bool isPrimary) : base(null,name,name, childEntityType, fundId)
        {
            FundId = fundId;
            Currency = currency;
            IsLongOnly = isLongOnly;
            IncludeHedging = includeHedging;
            IncludeOnlyFX = includeOnlyFX;
            if (isPrimary)
            {
                _rowType = RowType.FundTotal;
                
            }
            else
            {
                _rowType = RowType.AdditionalFundTotal;
            }
            CurrencyId = currencyId;
            
        }

        public FXExposureManager FXExposureManager { get; set; }

        public List<Fund> AdditionalFunds { get; set; }

        public Fund LastFund { get; set; }

        public int CurrencyId { get; private set; }

        public string Currency { get; private set; }        

        public XL.Range Range { get; set; }

        public WorksheetAccess WorksheetAccess { get; set; }

        public bool IsLongOnly { get; set; }

        public bool IncludeHedging { get; set; }
        
        public bool IncludeOnlyFX { get; set; }

        private RowType _rowType;
        protected override RowType RowTypeForNewRow => _rowType;

    }
}
