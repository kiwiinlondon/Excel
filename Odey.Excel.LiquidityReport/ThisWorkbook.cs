using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Exc = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Odey.LiquidityCalculator.Clients;
using Odey.LiquidityCalculator.Contracts;

namespace Odey.Excel.LiquidityReport
{
    public partial class ThisWorkbook
    {
        

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            LiquidityCalculatorClient client = new LiquidityCalculatorClient();
            Dictionary<string,LiquidityCalculatorOutput> outputs = client.Calculate(new LiquidityCalculatorInput(20, 20, 2, 50, 5), 2);
            Exc.Worksheet mainWorkSheet = (Exc.Worksheet)Globals.ThisWorkbook.Worksheets[1];

            mainWorkSheet.Cells[1, 1] = "Fund";
            mainWorkSheet.Cells[1, 2] = "GrossNav";
            mainWorkSheet.Cells[1, 3] = "Nav";
            mainWorkSheet.Cells[1, 4] = "Number Nonliquidated Positions";
            mainWorkSheet.Cells[1, 5] = "Number Of Positions";
            mainWorkSheet.Cells[1, 6] = "Total Fine";
            mainWorkSheet.Cells[1, 7] = "Unsold Value";
            mainWorkSheet.Cells[1, 8] = "Number Of Exceptions";
            mainWorkSheet.Cells[1, 9] = "Weighted Number of Days to 100% Liquidated ";
            
            int row = 2;
            foreach (KeyValuePair<string, LiquidityCalculatorOutput> output in outputs)
            {
                mainWorkSheet.Cells[row, 1] = output.Key;
                mainWorkSheet.Cells[row, 2] = output.Value.GrossNav;
                mainWorkSheet.Cells[row, 3] = output.Value.Nav;
                mainWorkSheet.Cells[row, 4] = output.Value.NonLiquidatedPortfolio.Count;
                mainWorkSheet.Cells[row, 5] = output.Value.NumberOfPositions;
                mainWorkSheet.Cells[row, 6] = output.Value.TotalFine;
                mainWorkSheet.Cells[row, 7] = output.Value.UnsoldValue;
                mainWorkSheet.Cells[row, 8] = output.Value.Exceptions.Count;
                mainWorkSheet.Cells[row, 9] = output.Value.TotalWeightedDaysToLiquidatePortfolio;
                row++;
            }
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
