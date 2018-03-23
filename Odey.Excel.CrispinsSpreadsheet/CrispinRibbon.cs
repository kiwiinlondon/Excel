using System;
using System.Collections.Generic;
using System.Deployment.Application;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using Odey.Framework.Keeley.Entities.Enums;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public partial class CrispinRibbon
    {


        private static readonly log4net.ILog Logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private void CrispinRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            
            log4net.Config.XmlConfigurator.Configure();
            Logger.Info("Loaded Ribbon");
            DoMatch(false);
            DisplayMessage(GetVersion());
            Logger.Info("Finished Loading Ribbon");
        }

        public string GetVersion()
        {
            if (System.Diagnostics.Debugger.IsAttached)
            {
                return "Debug Mode";
            }
            else
            {
                return $"Version { System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion}";
            }
        }


        private void DoMatch(bool refreshFormulas)
        {
            Logger.Info("Starting Match");

          //  try
         //   {

                var dataAccess = new DataAccess(DateTime.Today);
                var workbookAccess = new WorkbookAccess(Globals.ThisWorkbook);
                var matcher = new Matcher(new EntityBuilder(dataAccess, workbookAccess), dataAccess, workbookAccess, new InstrumentRetriever(new BloombergSecuritySetup(), dataAccess));
                matcher.Match(refreshFormulas);


         //   }
          //  catch (Exception ex)
         //   {
              //  Logger.Info(ex);
              //  DisplayMessage($"Error: {ex.Message}");
          //  }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            DoMatch(checkBox1.Checked);
            
        }

        private void DisplayMessage(string message)
        {
            this.label1.Label = message;
            this.label1.ShowLabel = true;
        }
        private void button2_Click_1(object sender, RibbonControlEventArgs e)
        {
            string ticker = this.editBox1.Text;
            var dataAccess = new DataAccess(DateTime.Today);
            var workbookAccess = new WorkbookAccess(Globals.ThisWorkbook);
            var matcher = new Matcher(new EntityBuilder(dataAccess, workbookAccess), dataAccess, workbookAccess, new InstrumentRetriever(new BloombergSecuritySetup(), dataAccess));
            string message = matcher.AddTicker(ticker);
            DisplayMessage(message);
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            var dataAccess = new DataAccess(DateTime.Today);
            var workbookAccess = new WorkbookAccess(Globals.ThisWorkbook);
            var matcher = new Matcher( new EntityBuilder(dataAccess, workbookAccess), dataAccess, workbookAccess, new InstrumentRetriever(new BloombergSecuritySetup(), dataAccess));
            string message = matcher.AddBulk();
            DisplayMessage(message);
        }

        
    }
}
