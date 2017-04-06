using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Odey.Framework.Keeley.Entities.Enums;
using Odey.PortfolioCache.Clients;
using Odey.PortfolioCache.Entities;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Odey.ExcelAddin
{

    public class ColumnDef
    {
        public int? Index { get; set; }
        public string AlphabeticalIndex { get; set; }

        public string Name { get; set; }
        public string Formula { get; set; }
        public string NumberFormat { get; set; }
        public double Width { get; set; }
    }

    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private static FundIds[] funds = new[] { FundIds.ARFF, FundIds.BVFF, FundIds.DEVM, FundIds.FDXH, FundIds.OUAR };

        private static Dictionary<string, string> managerInitials = new Dictionary<string, string>
        {
            { "Adrian Courtenay", "AC" },
            { "Jamie Grimston", "JG" },
            { "James Hanbury", "JH" },
        };

        public static string GetFundName(FundIds fund, List<PortfolioDTO> data)
        {
            foreach (var item in data)
            {
                if (item.FundId == (int)fund)
                {
                    return item.FundName;
                }
            }

            // Not found
            throw new Exception($"Received no data for '{fund}'");
        }

        public static string GetManagerInitials(string fullName)
        {
            if (managerInitials.ContainsKey(fullName))
            {
                return managerInitials[fullName];
            }
            else
            {
                return fullName;
            }
        }

        public Ribbon1()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Odey.ExcelAddin.Ribbon1.xml");
        }

        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnActionCallback(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;

            var prevScreenUpdating = app.ScreenUpdating;
            var prevEvents = app.EnableEvents;
            var prevCalculation = app.Calculation;
            app.ScreenUpdating = false;
            app.EnableEvents = false;
            app.Calculation = Excel.XlCalculation.xlCalculationManual;

            try
            {
                app.StatusBar = "Loading portfolio weightings...";
                var data = new PortfolioCacheClient().GetPortfolioExposures(new PortfolioRequestObject
                {
                    FundIds = funds.Cast<int>().ToArray(),
                    ReferenceDates = new[] { DateTime.Today },
                });
                //var Funds = new StaticDataClient().GetAllFunds().ToDictionary(f => f.EntityId);

                var watchList = WatchListSheet.GetWatchList(app, data);
                ApplyManagerOverrides(data, watchList);
                WatchListSheet.Write(app, watchList, "Watch List Top", true);
                WatchListSheet.Write(app, watchList, "Watch List Bottom", false);
                WatchListSheet.Write(app, watchList, "Watch List High Quality", true, "H");
                WatchListSheet.Write(app, watchList, "Watch List Low Quality", false, "L");
                foreach (var fund in funds)
                {
                    ExposureSheet.Write(app, fund, data, watchList);
                }
                foreach (var fund in funds)
                {
                    PortfolioSheet.Write(app, fund, data, watchList);
                }
                foreach (var fund in funds)
                {
                    ScenarioSheet.Write(app, fund, data, watchList);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, e.GetType().Name);
            }
            finally
            {
                app.StatusBar = null;
                app.EnableEvents = prevEvents;
                app.ScreenUpdating = prevScreenUpdating;
                app.Calculation = prevCalculation;
            }
        }

        private void ApplyManagerOverrides(List<PortfolioDTO> data, Dictionary<string, WatchListItem> watchList)
        {
            foreach (var row in data)
            {
                if (row.StrategyName == "None")
                {
                    row.StrategyName = null;
                }

                if (row.BloombergTicker != null && row.ManagerId == (int)ApplicationUserIds.JamesHanbury)
                {
                    // Automatic DEVM & FDXH manager override
                    if (row.FundId == (int)FundIds.DEVM || row.FundId == (int)FundIds.FDXH)
                    {
                        var others = data.Where(p => p.BloombergTicker == row.BloombergTicker && p.ManagerId != (int)ApplicationUserIds.JamesHanbury && p.FundId != (int)FundIds.DEVM && p.FundId != (int)FundIds.FDXH).ToList();
                        var ids = others.Select(p => p.ManagerId).Distinct();
                        if (ids.Count() == 1)
                        {
                            row.ManagerId = ids.Single();
                            row.ManagerName = others.Select(p => p.ManagerName).First();
                        }
                    }

                    // Manual manager override
                    if (watchList.ContainsKey(row.BloombergTicker))
                    {
                        var item = watchList[row.BloombergTicker];
                        if (item.JHManagerOverride != null)
                        {
                            row.ManagerName = item.JHManagerOverride;
                            row.ManagerId = -1;
                        }
                    }
                }
            }
        }
        
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

    }
}
