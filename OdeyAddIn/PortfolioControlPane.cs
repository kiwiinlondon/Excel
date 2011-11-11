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
            fundAndReferenceDatePicker1.CurrentDate = DateTime.Now.Date;
        }

        private List<CompletePortfolio> GetPortfolio()
        {
            PortfolioWebClient client = new PortfolioWebClient();
           
            int[] fundIds = null;

            if (!checkBox1.Checked)
            {
                fundIds = new int[] {fundAndReferenceDatePicker1.FundId};
            }

            int? reportCurrencyId = null;
            if (ReportFXRate.Checked)
            {
                reportCurrencyId = (int)currencyPicker1.SelectedValue;
            }
            int[] daysBeforeToday = new int[] {fundAndReferenceDatePicker1.DaysBeforeToday};
            bool includeShortPositions = true;
            if (ExcludeShortPositions.Checked)
            {
                includeShortPositions = false;
            }
            return client.GetCompletePortfolio(fundIds, daysBeforeToday, includeShortPositions, reportCurrencyId, null, null, null, null).OrderBy(a => a.InstrumentClass).ToList();            
        }


        private void button1_Click(object sender, EventArgs e)
        {
            List<CompletePortfolio> portfolio = GetPortfolio();
            PortfolioFields[] fieldsToReturn = GetFieldsToReturn();
            string reportCurrency = null;
            if (ReportFXRate.Checked)
            {
                reportCurrency = currencyPicker1.Text;
            }

            PortfolioWriter.Write(portfolio, Globals.ThisAddIn.Application.ActiveSheet, Globals.ThisAddIn.Application.ActiveCell.Row, Globals.ThisAddIn.Application.ActiveCell.Column, fieldsToReturn, reportCurrency);
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
