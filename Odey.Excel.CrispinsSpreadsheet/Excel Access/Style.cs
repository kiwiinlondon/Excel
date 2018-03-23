using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class Style
    {
        public Style(string name, bool isHidden, decimal width)
        {
            Name = name;
            IsHidden = isHidden;
            Width = width;
        }
        public string Name { get; set; }
        public bool IsHidden { get; set; }

        public decimal Width { get; set; }
    }
}
