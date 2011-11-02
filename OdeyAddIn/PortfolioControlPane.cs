using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Odey.Reporting.Clients;
using Odey.Reporting.Entities;

namespace OdeyAddIn
{
    public partial class PortfolioControlPane : UserControl
    {
        public PortfolioControlPane()
        {
            InitializeComponent();
        }

        private List<PortfolioWithUnderlyer> GetPortfolio()
        {
            PortfolioWebClient client = new PortfolioWebClient();
            if (checkBox1.Checked)
            {
                return client.GetPortfolioWithUnderlyerAllFunds(fundAndReferenceDatePicker1.DaysBeforeToday).OrderBy(a => a.UnderlyerParentInstrumentClass).ToList();
            }
            else
            {
                return client.GetPorfolioWithUnderlyer(fundAndReferenceDatePicker1.FundId, fundAndReferenceDatePicker1.DaysBeforeToday).OrderBy(a => a.UnderlyerParentInstrumentClass).ToList();
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            List<PortfolioWithUnderlyer> portfolio = GetPortfolio();
            PortfolioFields[] fieldsToReturn = GetFieldsToReturn();
            PortfolioWriter.Write(portfolio, Globals.ThisAddIn.Application.ActiveSheet, Globals.ThisAddIn.Application.ActiveCell.Row, Globals.ThisAddIn.Application.ActiveCell.Column, fieldsToReturn);
        }

        private PortfolioFields[] GetFieldsToReturn()
        {
            List<PortfolioFields> fieldsToReturn = new List<PortfolioFields>();
            if (ReferenceDate.Checked)fieldsToReturn.Add(PortfolioFields.ReferenceDate);
            if (InstrumentName.Checked) fieldsToReturn.Add(PortfolioFields.InstrumentName);
            if (ExchangeCode.Checked) fieldsToReturn.Add(PortfolioFields.BBExchangeCode);
            if (InstrumentClass.Checked)  fieldsToReturn.Add(PortfolioFields.InstrumentClass);
            if (Ticker.Checked)  fieldsToReturn.Add(PortfolioFields.BloombergTicker);
            if (ParentInstrumentClass.Checked) fieldsToReturn.Add(PortfolioFields.ParentInstrumentClass);
            if (Country.Checked) fieldsToReturn.Add(PortfolioFields.Country);
            if (Industry.Checked) fieldsToReturn.Add(PortfolioFields.Industry);
            if (Sector.Checked) fieldsToReturn.Add(PortfolioFields.Sector);
            if (UnderlyingInstrumentClass.Checked) fieldsToReturn.Add(PortfolioFields.UnderlyerInstrumentClass);
            if (UnderlyingParentInstrumentClass.Checked) fieldsToReturn.Add(PortfolioFields.UnderlyerParentInstrumentClass);
            if (UnderlyingInstrumentName.Checked) fieldsToReturn.Add(PortfolioFields.UnderlyingInstrumentName);
            if (UnderlyingCountry.Checked) fieldsToReturn.Add(PortfolioFields.UnderlyerCountry);
            if (UnderlyingTicker.Checked) fieldsToReturn.Add(PortfolioFields.UnderlyingBloombergTicker);
            if (UnderlyingIndustry.Checked) fieldsToReturn.Add(PortfolioFields.UnderlyerIndustry);
            if (UnderlyingSector.Checked) fieldsToReturn.Add(PortfolioFields.UnderlyerSector);
            if (NetPosition.Checked) fieldsToReturn.Add(PortfolioFields.NetPosition);
            if (MarketValue.Checked) fieldsToReturn.Add(PortfolioFields.MarketValue);
            if (DeltaMarketValue.Checked) fieldsToReturn.Add(PortfolioFields.DeltaMarketValue);
            return fieldsToReturn.ToArray();
        }


       
    }
}
