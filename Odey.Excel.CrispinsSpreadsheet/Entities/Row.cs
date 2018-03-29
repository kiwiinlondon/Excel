using XL=Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class Row
    {
        public Row(RowType rowType, XL.Range range)
        {
            RowType = rowType;
            Range = range;
        }
        public RowType RowType { get; private set; }

        public XL.Range Range { get; private set; }

        public int RowNumber { get { return Range.Row; } }
    }
}
