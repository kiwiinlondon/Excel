using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.Tools.Applications;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            string documentPath = @"C:\SVN\Odey\Odey.Excel\OUAR Valuation Matrix\OUAR valuation matrix 38.xlsx";
            int runtimeVersion = 0;

            
                runtimeVersion = ServerDocument.GetCustomizationVersion(documentPath);

               // if (runtimeVersion == 3)
              //  {
                    ServerDocument.RemoveCustomization(documentPath);
                  //  System.Windows.Forms.MessageBox.Show("The customization has been removed.");
             //   }
            


        }
    }
}
