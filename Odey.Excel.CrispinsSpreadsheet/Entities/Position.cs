using XL = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class Position : IChildEntity
    {
        public Position(Identifier identifier, string name, decimal priceDivisor, InstrumentTypeIds instrumentTypeId, bool invertPNL)
        {
            Identifier = identifier;
            Name = name;
            InstrumentTypeId = instrumentTypeId;
            PriceDivisor = priceDivisor;
            InvertPNL = invertPNL;
        }

        public Identifier Identifier { get; private set; }

        public XL.Range Row { get; set; }

        public int RowNumber => Row.Row;

        public string Name { get; set; }

        public object Ordering => InstrumentTypeId == InstrumentTypeIds.DoNotDelete ? "_" : "" + Name?.ToUpper();

        public string Currency { get; set; }

        public bool InvertPNL { get; set; }

        public decimal NetPosition { get; set; }

        public InstrumentTypeIds InstrumentTypeId { get; set; }

        public decimal? OdeyCurrentPrice { get; set; }

        public decimal? OdeyPreviousPrice { get; set; }

        public decimal? OdeyPreviousPreviousPrice { get; set; }

        public decimal PriceDivisor { get; set; }

        public decimal PreviousNetPosition { get; set; }

        public override string ToString()
        {
            return $"{Identifier}: {Name}";
        }

    }
}
