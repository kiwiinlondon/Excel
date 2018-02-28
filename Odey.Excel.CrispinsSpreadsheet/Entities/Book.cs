using XL = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class Book : GroupingEntity, IChildEntity
    {
        public Book(string code) : base(code,0)
        {

        }

        public bool IsPrimary { get; set; }

    }
}
