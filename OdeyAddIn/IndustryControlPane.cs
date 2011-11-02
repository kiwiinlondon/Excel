using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Odey.Reporting.Clients;
using Odey.Framework.Keeley.Entities.Enums;

namespace OdeyAddIn
{
    public partial class IndustryControlPane : UserControl
    {
        public IndustryControlPane()
        {
            InitializeComponent();                  
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PortfolioWebClient client = new PortfolioWebClient();

            AggregatedPortfolioWriter.Write(client.GetAggregatedByIndustry(fundAndReferenceDatePicker1.FundId, fundAndReferenceDatePicker1.DaysBeforeToday).OrderBy(a => a.EntityName).ToList(),
                Globals.ThisAddIn.Application.ActiveSheet, Globals.ThisAddIn.Application.ActiveCell.Row, Globals.ThisAddIn.Application.ActiveCell.Column,EntityTypeIds.Industry);
        }
    }
}
