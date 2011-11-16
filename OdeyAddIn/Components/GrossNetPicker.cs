using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OdeyAddIn.Components
{
    public partial class GrossNetPicker : UserControl
    {
        public GrossNetPicker()
        {
            InitializeComponent();
        }

        public bool IncludeRawData
        {
            get
            {
                return this.checkBox1.Checked;
            }
        }

        public AggregatedPortfolioOutputOptions OutputOption
        {
            get
            {
                string selectedItem = comboBox1.SelectedItem.ToString();
                switch (selectedItem)
                {
                    case "Long/Short":
                        return AggregatedPortfolioOutputOptions.LongShort;
                    case "Net":
                        return AggregatedPortfolioOutputOptions.Net;
                    default:
                        return AggregatedPortfolioOutputOptions.Gross;
                }
            }
        }
    }
}
