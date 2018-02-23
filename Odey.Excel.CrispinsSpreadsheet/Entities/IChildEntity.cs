using XL=Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public interface IChildEntity
    {
        string Name { get; set; }

        int RowNumber { get; }
    }
}
