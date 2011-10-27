using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace OdeyAddIn
{
    public partial class OdeyRibbonTab
    {
        private void OdeyRibbonTab_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void industryButton_Click(object sender, RibbonControlEventArgs e)
        {


           //industryControlPane.DockPosition.Control.Location = new System.Drawing.Point(15, 15);
         //   Globals.ThisAddIn.Application.CommandBars["MyCustomTaskPane"].Left = 500;

            Globals.ThisAddIn.IndustryPane.Visible = ((RibbonToggleButton)sender).Checked;
        }
    }
}
