using Entities=Odey.Framework.Keeley.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class DTOGrouping
    {
        public Entities.Position Position { get; set; }
        public Entities.Portfolio PreviousPrevious { get; set; }
        public Entities.Portfolio Previous { get; set; }
        public Entities.Portfolio Current { get; set; }
    }
}
