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
    public partial class FundAndDateControlPane : UserControl
    {
        public FundAndDateControlPane()
        {
            InitializeComponent();
            referenceDatePicker.MaxDate = DateTime.Now.Date;
            referenceDatePicker.Value = DateTime.Now.Date;
            comboBox1.DataSource = Globals.ThisAddIn.FundsWithPositions;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "FundId";            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PortfolioWebClient client = new PortfolioWebClient();

            int daysBeforeToDays = DateTime.Now.Date.Subtract(referenceDatePicker.Value.Date).Days;
            int fundId = (int)comboBox1.SelectedValue;
            AggregatedPortfolioWriter.Write(client.GetAggregatedByIndustry(fundId, daysBeforeToDays).OrderBy(a=>a.EntityName).ToList(),
                Globals.ThisAddIn.Application.ActiveSheet, Globals.ThisAddIn.Application.ActiveCell.Row, Globals.ThisAddIn.Application.ActiveCell.Column,EntityTypeIds.Industry);
        }

        


    

       
    }
}
