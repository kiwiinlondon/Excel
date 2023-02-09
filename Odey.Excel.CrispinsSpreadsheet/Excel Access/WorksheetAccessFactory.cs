using XL=Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class WorksheetAccessFactory
    {

        private static readonly WorksheetAccessFactory instance = new WorksheetAccessFactory();

        private WorksheetAccessFactory()
        {

        }

        public static WorksheetAccessFactory Instance
        {
            get
            {
                return instance;
            }
        }

        public WorksheetAccess Get(XL.Worksheet worksheet, bool isLong)
        {
            if (!isLong)
            {

                return new LongShortWithoutBooksWorksheetAccess(worksheet);

            }
            else
            {

                return new LongOnlyWorksheetAccess(worksheet);

            }
        }
    }
}
