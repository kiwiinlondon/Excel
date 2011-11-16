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
using Odey.Reporting.Entities;

namespace OdeyAddIn
{
    public partial class TopHoldingsControlPane : UserControl
    {
        public TopHoldingsControlPane()
        {
            InitializeComponent();
            fundAndReferenceDatePicker1.CurrentDate = DateTime.Now.Date;
        }

        private void buttn1_Click(object sender, EventArgs e)
        {
            
            PortfolioWebClient client = new PortfolioWebClient();
            bool? equitiesOnly = equityPicker1.Selected;

            AggregatedPortfolioFields[] fieldsToReturn = AggregatedPortfolioFieldsHelper.Get(this.grossNetPicker1.IncludeRawData, this.grossNetPicker1.OutputOption);

            int numberOfResults = 0;
            if (int.TryParse(textBox1.Text, out numberOfResults))
            {
                List<AggregatedPortfolio> portfolio = null;
                if (fundAndReferenceDatePicker1.UsePeriodicity)
                {
                    portfolio = client.GetTopHoldingsMultipleOverTime(fundAndReferenceDatePicker1.FundIds, fundAndReferenceDatePicker1.PeriodicityId, fundAndReferenceDatePicker1.FromDaysPriorToToday, fundAndReferenceDatePicker1.ToDaysPriorToToday, equitiesOnly, numberOfResults).OrderByDescending(a => a.Long + Math.Abs(a.Short)).ToList();
                }
                else
                {
                    portfolio = client.GetTopHoldingsMultiple(fundAndReferenceDatePicker1.FundIds, fundAndReferenceDatePicker1.SelectedDates, equitiesOnly, numberOfResults).ToList();
                }

                AggregatedPortfolioWriter.Write(portfolio,Globals.ThisAddIn.Application.ActiveSheet, Globals.ThisAddIn.Application.ActiveCell.Row, Globals.ThisAddIn.Application.ActiveCell.Column, EntityTypeIds.InstrumentMarket, fieldsToReturn);        
            }
        
        }

       
    }
}
