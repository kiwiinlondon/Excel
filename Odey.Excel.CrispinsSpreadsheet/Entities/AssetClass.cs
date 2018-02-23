using XL=Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class AssetClass : GroupingEntity, IChildEntity
    {
        public AssetClass(string code) : base(code)
        {

        }

    }
}
