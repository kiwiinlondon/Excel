﻿using System;
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
    public partial class IndustryControlPane : UserControl
    {
        public IndustryControlPane()
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
            int[] fundIds = fundAndReferenceDatePicker1.FundIds;

            AggregatedPortfolioFields[] fieldsToReturn = AggregatedPortfolioFieldsHelper.Get(this.grossNetPicker1.IncludeRawData, this.grossNetPicker1.OutputOption);

            List<AggregatedPortfolio> portfolio = null;
            if (fundAndReferenceDatePicker1.UsePeriodicity)
            {                
                portfolio = client.GetAggregatedByIndustryMultipleOverTime(fundIds, fundAndReferenceDatePicker1.PeriodicityId, fundAndReferenceDatePicker1.FromDaysPriorToToday, fundAndReferenceDatePicker1.ToDaysPriorToToday, equitiesOnly, includeCash).OrderBy(a => a.Long).ToList();
            }
            else
            {
                portfolio = client.GetAggregatedByIndustryMultiple(fundIds, fundAndReferenceDatePicker1.SelectedDates, equitiesOnly, includeCash).OrderBy(a => a.Long).ToList();
            }

            AggregatedPortfolioWriter.Write(portfolio,Globals.ThisAddIn.Application.ActiveSheet, Globals.ThisAddIn.Application.ActiveCell.Row, Globals.ThisAddIn.Application.ActiveCell.Column, EntityTypeIds.Industry, fieldsToReturn);
        }
        
    }
}
