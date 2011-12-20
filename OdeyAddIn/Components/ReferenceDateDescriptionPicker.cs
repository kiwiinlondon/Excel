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
    public partial class ReferenceDateDescriptionPicker : UserControl
    {
        public ReferenceDateDescriptionPicker()
        {
            InitializeComponent();
        }

        public int PeriodicityId
        {
            get
            {
                return (int)this.periodicityPicker1.SelectedValue;
            }
        }

        public string PeriodicityText
        {
            get
            {
                return this.periodicityPicker1.Text;
            }
        }

        public DateTime CurrentDate
        {
            set
            {
                this.referenceDatePicker1.MaxDate = value;
                this.referenceDatePicker1.Value = value.AddMonths(-1);
                this.referenceDatePicker2.CurrentDate = value;
                
            }
        }        

        public int? FromDaysBeforeToday
        {
            get
            {
                if (checkBox1.Checked)
                {
                    return null;
                }
                else
                {
                    return DateTime.Now.Date.Subtract(referenceDatePicker1.Value.Date).Days;
                }
            }
        }

        public int ToDaysBeforeToday
        {
            get
            {
                return DateTime.Now.Date.Subtract(referenceDatePicker2.Value.Date).Days;
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                this.referenceDatePicker1.Enabled = false;
            }
            else
            {
                this.referenceDatePicker1.Enabled = true;
            }
        }
    }
}
