using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OdeyAddIn.Components
{
    public partial class ReferenceDateDescriptorForm : Form
    {
        public ReferenceDateDescriptorForm()
        {
            InitializeComponent();
            DisableGroup2();
        }

        public DateTime CurrentDate
        {            
            set
            {
                this.multipleReferenceDatePicker1.CurrentDate = value;
                this.referenceDateDescriptionPicker1.CurrentDate = value; 
            }
        }        

        public int PeriodicityId
        {
            get
            {
                return this.referenceDateDescriptionPicker1.PeriodicityId;
            }            
        }

        public string PeriodicityText
        {
            get
            {
                return this.referenceDateDescriptionPicker1.PeriodicityText;
            }
        }

        public int? FromDaysBeforeToday
        {
            get
            {
                return this.referenceDateDescriptionPicker1.FromDaysBeforeToday;
            }
        }
        public int ToDaysBeforeToday
        {
            get
            {
                return this.referenceDateDescriptionPicker1.ToDaysBeforeToday;
            }
        }
        
        private void ReferenceDateDescriptorForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Hide();
            e.Cancel = true;
        }

        public DateTime[] SelectedDates
        {
            get
            {
                return this.multipleReferenceDatePicker1.SelectedDates;
            }
        }

        private void DisableGroup1()
        {
            multipleReferenceDatePicker1.IsEnabled = false;
            referenceDateDescriptionPicker1.Enabled = true;
        }

        private void DisableGroup2()
        {
            referenceDateDescriptionPicker1.Enabled = false;
            multipleReferenceDatePicker1.IsEnabled = true;
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                DisableGroup2();
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                DisableGroup1();
            }

        }

        public bool IsPeriodicityUsed
        {
            get
            {
                return radioButton2.Checked;
            }
        }
    }
}
