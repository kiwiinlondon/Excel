using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class FundDTO
    {
        public int FundId { get; set; }

        public string Name { get; set; }

        public string Currency { get; set; }

        public Dictionary<string,BookDTO> Books { get; set; }

        public decimal CurrentNav { get; set; }

        public decimal PreviousNav { get; set; }

        public FundFXTreatmentIds FundFXTreatmentId { get; set; }

        
    }
}
