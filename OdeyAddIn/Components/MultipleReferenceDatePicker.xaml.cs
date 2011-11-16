using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace OdeyAddIn.Components
{
    /// <summary>
    /// Interaction logic for MultipleReferenceDatePicker.xaml
    /// </summary>
    public partial class MultipleReferenceDatePicker : UserControl
    {
        public MultipleReferenceDatePicker()
        {
            InitializeComponent();
        }

        public DateTime[] SelectedDates
        {
            get
            {
                SelectedDatesCollection collection = this.calendar1.SelectedDates;
                return collection.ToArray<DateTime>();
            }
        }

        public DateTime CurrentDate
        {
            set
            {
                this.calendar1.DisplayDateStart = new DateTime(1999, 7, 30);
                this.calendar1.DisplayDateEnd = value;
                this.calendar1.SelectedDate = value;
            }
        }
    }
}
