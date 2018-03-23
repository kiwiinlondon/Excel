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
        public int BookId { get; private set; }
        public Book(Fund fund,int bookId,string code, EntityTypes childEntityType) : base(fund,code, code, childEntityType, code)
        {
            BookId = bookId;
        }

        public bool IsPrimary { get; set; }

    }
}
