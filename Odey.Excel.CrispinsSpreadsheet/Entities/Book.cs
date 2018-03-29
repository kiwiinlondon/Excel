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
        public Book(Fund fund,int bookId,string code, EntityTypes childEntityType, bool isPrimaryBook) : base(fund,code, code, childEntityType, code)
        {
            BookId = bookId;
            IsPrimary = isPrimaryBook;
        }

        protected override RowType RowTypeForNewRow
        {
            get
            {
                return IsPrimary ? RowType.MainBookOrAssetClassTotal : RowType.SecondaryBookTotal;
            }
        }

        public bool IsPrimary { get; set; }

    }
}
