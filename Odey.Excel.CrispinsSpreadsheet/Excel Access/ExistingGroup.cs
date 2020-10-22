using XL=Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class ExistingGroupDTO
    {

        public ExistingGroupDTO(string controlString,string name, XL.Range totalRow, List<ExistingPositionDTO> positions)
        {
            ControlString = controlString;
            var values = ControlString.Split('#');
            FundCode = values[0];
            AssetClassCode = values[1];
            CountryCode = values[2];
            Name = name;
            TotalRow = totalRow;
            Positions = positions;
        }


        public string FundCode { get; private set; }
        public string AssetClassCode { get; private set; }
        public string CountryCode { get; private set; }

        public XL.Range TotalRow { get; set; }
        public string Name { get; set; }
        public string ControlString { get; set; }

        public List<ExistingPositionDTO> Positions { get; set; }
    }
}
