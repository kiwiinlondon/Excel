using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OdeyAddIn.Components
{
    public class PeriodicityPicker : ComboBox
    {
        public PeriodicityPicker()
        {
            if (Globals.ThisAddIn != null)
            {
                DataSource = Globals.ThisAddIn.Periodicities;
            }
            DisplayMember = "Name";
            ValueMember = "PeriodicityId";
        }
    }
}
