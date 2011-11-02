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
            MaxDate = DateTime.Now.Date;
            Value = DateTime.Now.Date;
        }
    }
}
