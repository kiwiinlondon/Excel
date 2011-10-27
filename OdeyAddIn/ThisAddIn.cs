using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Exc = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace OdeyAddIn
{
    public partial class ThisAddIn
    {

        private Microsoft.Office.Tools.CustomTaskPane industryControlPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            IndustryControlPane industryControlPaneToAdd = new IndustryControlPane();
            industryControlPane = this.CustomTaskPanes.Add(
                industryControlPaneToAdd, "Liquidity Parameters");
            industryControlPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            industryControlPane.VisibleChanged +=
                new EventHandler(industryControlPanelValue_VisibleChanged);
        }

        private void industryControlPanelValue_VisibleChanged(object sender, System.EventArgs e)
        {
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
