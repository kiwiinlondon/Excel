using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
namespace OdeyAddIn
{
    public class ReferenceDatePicker : DateTimePicker
    {
        public ReferenceDatePicker()
        {
            MinDate = new DateTime(1999, 7, 30);
        }

        public DateTime CurrentDate
        {
            set
            {
                MaxDate = value;
                Value = value;
                
            }
        }
    }
}
