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
using Odey.Reporting.Clients;

namespace Odey.ExcelAddin
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private static FundIds[] funds = new[] { FundIds.ARFF, FundIds.BVFF, FundIds.DEVM, FundIds.FDXH, FundIds.OUAR };

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
           
            return GetResourceText("Odey.ExcelAddin.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnActionCallback(Office.IRibbonControl control)
        {
            var Funds = new StaticDataClient().GetAllFunds().ToDictionary(f => f.EntityId);
            var data = new PortfolioCacheClient().GetPortfolioExposures(new PortfolioRequestObject
            {
                FundIds = funds.Cast<int>().ToArray(),
                ReferenceDates = new[] { DateTime.Today },
            });


            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Worksheet sheet;
            try
            {
                sheet = app.Sheets["Weightings"];
            }
            catch
            {
                sheet = app.Sheets.Add(After: app.Sheets[app.Sheets.Count]);
                sheet.Name = "Weightings";
            }

            // Clear content
            sheet.UsedRange.ClearContents();

            // Headers
            sheet.Cells[1, 1] = "Ticker";
            sheet.Cells[1, 2] = "Name";
            var startColumn = 3;
            foreach (var fund in funds)
            {
                sheet.Cells[1, startColumn + Array.IndexOf(funds, fund)] = fund.ToString();
            }

            // Rows
            var row = 2;

            var instruments = data.Where(p => p.ExposureTypeId == ExposureTypeIds.Primary).OrderBy(p => p.BloombergTicker).ToLookup(p => p.EquivalentInstrumentMarketId);
            foreach (var instrument in instruments)
            {
                var ticker = instrument.Select(p => p.BloombergTicker).Distinct().Single();
                sheet.Cells[row, 1] = ticker;

                var nonCFD = instrument.FirstOrDefault(p => p.InstrumentClassId != (int)InstrumentClassIds.ContractForDifference);
                sheet.Cells[row, 2] = (nonCFD ?? instrument.First()).InstrumentName;

                foreach (var fund in instrument.ToLookup(p => p.FundId)) {
                    var fundNAV = fund.Select(p => p.FundNAV).Distinct().Single();
                    var column = (startColumn + Array.IndexOf(funds, (FundIds)fund.Key));
                    sheet.Cells[row, column] = fund.Sum(p => p.Exposure) / fundNAV;
                }
                ++row;
            }

            ++row;
            var currencies = data.Where(p => p.ExposureTypeId == ExposureTypeIds.Currency).OrderBy(p => p.PositionCurrency).ToLookup(p => p.PositionCurrencyId);
            foreach (var currency in currencies)
            {
                sheet.Cells[row, 2] = currency.Select(p => p.PositionCurrency).Distinct().Single();
                foreach (var fund in currency.ToLookup(p => p.FundId))
                {
                    var fundNAV = fund.Select(p => p.FundNAV).Distinct().Single();
                    var column = (startColumn + Array.IndexOf(funds, (FundIds)fund.Key));
                    var exposure = fund.Sum(p => p.Exposure);
                    if (Funds[fund.Key].CurrencyId == currency.Key)
                    {
                        exposure -= fundNAV;
                        Excel.Range cell = sheet.Cells[row, column];
                        cell.Font.Bold = true;
                    }
                    sheet.Cells[row, column] = exposure / fundNAV;
                }
                ++row;
            }

            // Set column width
            Excel.Range firstColumn = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 1]];
            firstColumn.ColumnWidth = 22;

            Excel.Range secondColumn = sheet.Range[sheet.Cells[1, 2], sheet.Cells[1, 2]];
            secondColumn.ColumnWidth = 35;

            // Set number format
            Excel.Range numbers = sheet.Range[sheet.Cells[2, startColumn], sheet.Cells[row - 1, startColumn + funds.Count()]];
            numbers.NumberFormat = "0.00%";

            // Trigger refresh all
            Globals.ThisAddIn.Application.ActiveWorkbook.RefreshAll();
        }

        #endregion

        #region Helpers

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

        #endregion
    }
}
