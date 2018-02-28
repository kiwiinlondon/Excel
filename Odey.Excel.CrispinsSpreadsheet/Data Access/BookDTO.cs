using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class BookDTO
    {
        public int BookId { get; set; }

        public string Name { get; set; }

        public decimal Nav { get; set; }

        public decimal PreviousNav { get; set; }
    }
}
