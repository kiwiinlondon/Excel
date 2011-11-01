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

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            IndustryControlPane industryControlPaneToAdd = new IndustryControlPane();
            industryControlPane = this.CustomTaskPanes.Add(
                industryControlPaneToAdd, "Industry Parameters");
            industryControlPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            industryControlPane.VisibleChanged +=
                new EventHandler(industryControlPanelValue_VisibleChanged);
        }

        private void industryControlPanelValue_VisibleChanged(object sender, System.EventArgs e)
        {
            LoadFunds();
            Globals.Ribbons.OdeyRibbonTab.industryButton.Checked =
                industryControlPane.Visible;
            
        }

        public Microsoft.Office.Tools.CustomTaskPane IndustryPane
        {
            get
            {
                return industryControlPane;
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
