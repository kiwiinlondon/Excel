using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Exc = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Odey.Reporting.Entities;
using Odey.Reporting.Clients;
using System.ComponentModel;
namespace OdeyAddIn
{
    public partial class ThisAddIn
    {

        private BindingList<Fund> _fundsWithPositions = new BindingList<Fund>();
        bool fundsLoaded = false;

        public BindingList<Fund> FundsWithPositions
        {
            get
            {                
                return _fundsWithPositions;
            }            
        }

        public void LoadFunds()
        {
            if (!fundsLoaded)
            {
                FundClient client = new FundClient();
                foreach(Fund fund in client.GetFundsWithPositions())
                {
                    _fundsWithPositions.Add(fund);
                }
                fundsLoaded = true;
            }
        }

        private Microsoft.Office.Tools.CustomTaskPane industryControlPane;
        private Microsoft.Office.Tools.CustomTaskPane countryControlPane;
        private Microsoft.Office.Tools.CustomTaskPane portfolioControlPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            IndustryControlPane industryControlPaneToAdd = new IndustryControlPane();
            industryControlPane = this.CustomTaskPanes.Add(
                industryControlPaneToAdd, "Industry Parameters");
            industryControlPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            industryControlPane.VisibleChanged +=
                new EventHandler(industryControlPanelValue_VisibleChanged);
            
            CountryControlPane countryControlPaneToAdd = new CountryControlPane();
            countryControlPane = this.CustomTaskPanes.Add(
                countryControlPaneToAdd, "Country Parameters");
            countryControlPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            countryControlPane.VisibleChanged +=
                new EventHandler(countryControlPanelValue_VisibleChanged);

            PortfolioControlPane portfolioControlPaneToAdd = new PortfolioControlPane();
            portfolioControlPane = this.CustomTaskPanes.Add(
                portfolioControlPaneToAdd, "Portfolio Parameters");
            portfolioControlPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            portfolioControlPane.VisibleChanged +=
                new EventHandler(portfolioControlPanelValue_VisibleChanged);
        }

        private void industryControlPanelValue_VisibleChanged(object sender, System.EventArgs e)
        {
            LoadFunds();
            Globals.Ribbons.OdeyRibbonTab.industryButton.Checked =
                industryControlPane.Visible;
            
        }

        private void countryControlPanelValue_VisibleChanged(object sender, System.EventArgs e)
        {
            LoadFunds();
            Globals.Ribbons.OdeyRibbonTab.countryButton.Checked =
                countryControlPane.Visible;

        }

        private void portfolioControlPanelValue_VisibleChanged(object sender, System.EventArgs e)
        {
            LoadFunds();
            Globals.Ribbons.OdeyRibbonTab.portfolioButton.Checked =
                portfolioControlPane.Visible;

        }

        public Microsoft.Office.Tools.CustomTaskPane IndustryPane
        {
            get
            {
                return industryControlPane;
            }
        }

        public Microsoft.Office.Tools.CustomTaskPane CountryPane
        {
            get
            {
                return countryControlPane;
            }
        }

        public Microsoft.Office.Tools.CustomTaskPane PortfolioPane
        {
            get
            {
                return portfolioControlPane;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
