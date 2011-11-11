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
            Globals.ThisAddIn.IndustryPane.Visible = ((RibbonToggleButton)sender).Checked;
        }

        private void countryButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CountryPane.Visible = ((RibbonToggleButton)sender).Checked;
        }

        private void portfolioButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.PortfolioPane.Visible = ((RibbonToggleButton)sender).Checked;
        }

        private void TopHoldings_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TopHoldingsPane.Visible = ((RibbonToggleButton)sender).Checked;
        }

        private void CurrencyButton_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CurrencyPane.Visible = ((RibbonToggleButton)sender).Checked;
        }
    }
}
