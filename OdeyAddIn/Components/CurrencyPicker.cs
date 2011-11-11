using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OdeyAddIn.Components
{
    class CurrencyPicker : ComboBox
    {
        public CurrencyPicker()
        {
            if (Globals.ThisAddIn != null)
            {
                DataSource = Globals.ThisAddIn.Currencies;
            }
            DisplayMember = "CcyIsoCode";
            ValueMember = "CurrencyId";
        }
    }
}
