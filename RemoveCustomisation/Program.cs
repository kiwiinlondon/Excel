using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Tools.Applications;

namespace RemoveCustomisation
{
    class Program
    {
        static void Main(string[] args)
        {
            string documentPath = @"C:\git\odey\code\Odey.Excel\Odey.Excel.CrispinsSpreadsheet\Crispin Spreadsheet.xlsx";
            int runtimeVersion = 0;

            runtimeVersion = ServerDocument.GetCustomizationVersion(documentPath);

            if (runtimeVersion == 3)
            {
                ServerDocument.RemoveCustomization(documentPath);
            }


        }    
    }
}
