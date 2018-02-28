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
        public Fund(string code,int firstRowOffset) : base(code, firstRowOffset)
        {

        }
        public string Currency { get; set; }

    }
}
