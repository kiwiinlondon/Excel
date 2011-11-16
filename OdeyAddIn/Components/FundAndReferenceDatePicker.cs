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
        BindingList<string> dates = null;
        public FundAndReferenceDatePicker()
        {
            InitializeComponent();
            RefDescriptionPicker.FormClosing +=
                new FormClosingEventHandler(refDescriptionPicker_FormClosing);
            dates = new BindingList<string>() { DateTime.Now.Date.ToLongDateString() };
            //referenceDatePicker1.DataSource = dates;
            referenceDatePicker1.Text = DateTime.Now.Date.ToLongDateString();
        }

        private void refDescriptionPicker_FormClosing(object sender, System.EventArgs e)
        {
          //  if (!dates.Contains(RefDescriptionPicker.CurrentDate.ToLongDateString()))
          //  {
         //       dates.Add(RefDescriptionPicker.CurrentDate.ToLongDateString());
         //   }
            if (RefDescriptionPicker.IsPeriodicityUsed)
            {
                referenceDatePicker1.Text = RefDescriptionPicker.PeriodicityText;                
            }
            else
            {
                if (RefDescriptionPicker.SelectedDates.Length == 0)
                {
                    referenceDatePicker1.Text = "NO DATE SELECTED";
                }
                else if (RefDescriptionPicker.SelectedDates.Length > 1)
                {
                    referenceDatePicker1.Text = "Multiple";
                }
                else
                {
                    referenceDatePicker1.Text = RefDescriptionPicker.SelectedDates[0].ToLongDateString();
                }
            }
        }

        public int[] FundIds
        {
            get
            {
                return fundPicker1.SelectedFundIds;
            }
        }

        public int[] SelectedDates
        {
            get
            {
                return RefDescriptionPicker.SelectedDates.Select(a => DateTime.Now.Date.Subtract(a).Days).ToArray();
            }
        }

        public bool UsePeriodicity
        {
            get
            {
                return RefDescriptionPicker.IsPeriodicityUsed;
            }
        }

        public int PeriodicityId
        {
            get
            {
                return RefDescriptionPicker.PeriodicityId;
            }
        }
        public int? FromDaysPriorToToday
        {
            get
            {
                return RefDescriptionPicker.FromDaysBeforeToday;
            }
        }

        public int? ToDaysPriorToToday
        {
            get
            {
                return RefDescriptionPicker.ToDaysBeforeToday;
            }
        }


        public DateTime CurrentDate
        {          
            set
            {
                RefDescriptionPicker.CurrentDate = value;
            }
            
        }

        

        private ReferenceDateDescriptorForm _myFrm = new ReferenceDateDescriptorForm();
        private ReferenceDateDescriptorForm RefDescriptionPicker
        {
            get
            {               
                return _myFrm;
            }
        }
        


        private void referenceDatePicker1_MouseClick(object sender, MouseEventArgs e)
        {
            RefDescriptionPicker.StartPosition = FormStartPosition.CenterParent;

            RefDescriptionPicker.ShowDialog();
        }

        
    }
}
