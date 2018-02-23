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
        public Country(string code) : base(code)
        {

        }

        public string BloombergCode { get; set; }

 


    }
}
