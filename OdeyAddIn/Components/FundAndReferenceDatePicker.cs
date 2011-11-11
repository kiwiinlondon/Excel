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
    public partial class FundAndReferenceDatePicker : UserControl
    {
        public FundAndReferenceDatePicker()
        {
            InitializeComponent();
        }

        public int FundId
        {
            get
            {
                return (int)fundPicker1.SelectedValue;
            }
        }

        public int DaysBeforeToday
        {
            get
            {
                return DateTime.Now.Date.Subtract(referenceDatePicker1.Value.Date).Days;
            }
        }

        public DateTime CurrentDate
        {
            set
            {
                referenceDatePicker1.CurrentDate = value;
            }
        }
    }
}
