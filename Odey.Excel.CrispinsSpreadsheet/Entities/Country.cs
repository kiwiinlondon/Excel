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
        public Country(AssetClass assetClass, string code,string name) : base(assetClass,code, name,true,name)
        {

        }

        public string BloombergCode { get; set; }

 


    }
}
