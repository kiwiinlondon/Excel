using XL=Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class Country : GroupingEntity, IChildEntity
    {
        public Country(GroupingEntity assetClass, string code,string name) : base(assetClass,code, name, EntityTypes.Position, name)
        {

        }

        public string BloombergCode { get; set; }


        protected override RowType RowTypeForNewRow => RowType.Total;

    }
}
