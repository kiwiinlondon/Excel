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
    public partial class EquityPicker : UserControl
    {
        public EquityPicker()
        {
            InitializeComponent();
        }

        public bool? Selected
        {
            get
            {
                string selectedItem = comboBox1.SelectedItem.ToString();
                switch (selectedItem)
                {
                    case "Equity":
                        return true;
                    case "Non-Equity":
                        return false;
                    default:
                        return null;
                }                
            }
        }
    }
}
