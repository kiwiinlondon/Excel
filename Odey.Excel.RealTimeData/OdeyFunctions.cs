using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.RealTimeData
{
    public class OdeyFunctions
    {

        [ExcelFunction(Description = "My first .NET function")]
        public static string HelloDna(string name)
        {
            return "Hello " + name;
        }
    }
}
