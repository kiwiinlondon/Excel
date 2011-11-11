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
    public partial class CountryControlPane : UserControl
    {
        public CountryControlPane()
        {
            InitializeComponent();
            fundAndReferenceDatePicker1.CurrentDate = DateTime.Now.Date;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            PortfolioWebClient client = new PortfolioWebClient();
            bool? equitiesOnly = equityPicker1.Selected;
           
            bool includeCash = true;
            if (!checkBox1.Checked)
            {
                includeCash = false;
            }
            AggregatedPortfolioWriter.Write(client.GetAggregatedByCountry(fundAndReferenceDatePicker1.FundId, fundAndReferenceDatePicker1.DaysBeforeToday, equitiesOnly, includeCash).OrderBy(a => a.Long).ToList(),
                Globals.ThisAddIn.Application.ActiveSheet, Globals.ThisAddIn.Application.ActiveCell.Row, Globals.ThisAddIn.Application.ActiveCell.Column,EntityTypeIds.Country);        
        }
    }
}
