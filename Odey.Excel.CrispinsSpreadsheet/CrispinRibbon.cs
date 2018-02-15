using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public partial class CrispinRibbon
    {

        Matcher _matcher = null;
        private void CrispinRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            _matcher = new Matcher(new DataAccess(),new SheetAccess(Globals.ThisWorkbook));
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            _matcher.Match(741, DateTime.Today);
        }        

    }
}
