﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace OUAR_Valuation_Matrix
{
    public partial class Sheet18
    {
        private void Sheet18_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet18_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet18_Startup);
            this.Shutdown += new System.EventHandler(Sheet18_Shutdown);
        }

        #endregion

    }
}
