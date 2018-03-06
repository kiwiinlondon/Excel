using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Odey.Framework.Keeley.Entities.Enums;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public partial class CrispinRibbon
    {


        private Matcher _matcher = null;
        

        private void CrispinRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            var dataAccess = new DataAccess(DateTime.Today);
            var sheetAccess = new SheetAccess(Globals.ThisWorkbook);
            _matcher = new Matcher(new EntityBuilder(dataAccess, sheetAccess), dataAccess, sheetAccess, new InstrumentRetriever(new BloombergSecuritySetup(), dataAccess));
            _matcher.BuildFunds();
            _matcher.Match(false);
        }



        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            _matcher.Match(true);
            DisplayMessage("Success");
        }

        private void DisplayMessage(string message)
        {
            this.label1.Label = message;
            this.label1.ShowLabel = true;
        }
        private void button2_Click_1(object sender, RibbonControlEventArgs e)
        {
            string ticker = this.editBox1.Text;
            string message = _matcher.AddTicker(ticker);
            DisplayMessage(message);
        }
    }
}
