using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class FXRateDTO
    {
        public string FromCurrency { get; set; }

        public string ToCurrency { get; set; }

        public decimal PreviousValue { get; set; }

        public decimal PreviousPreviousValue { get; set; }
    }
}
