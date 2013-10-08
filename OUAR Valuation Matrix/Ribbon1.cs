using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Odey.Reporting.Clients;
using Odey.Reporting.Entities;

namespace OUAR_Valuation_Matrix
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ReplaceFormulasWithTheirValues(string worksheetName)
        {
            Excel.Worksheet worksheet = Globals.ThisWorkbook.Sheets[worksheetName];
            ReplaceFormulasWithTheirValues(worksheet);
        }

        private void ReplaceFormulasWithTheirValues(Excel.Worksheet worksheet)
        {
            worksheet.Cells.Copy();
            worksheet.Cells.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
        }

        private void SortSheet(string worksheetName, int firstRow, string lastColumn, string sortColumn, Tuple<string, string>[] stringsToStopAtWithColumn, bool ascending)
        {
            Excel.Worksheet worksheet = Globals.ThisWorkbook.Sheets[worksheetName];
            SortSheet(worksheet, firstRow, lastColumn, sortColumn, stringsToStopAtWithColumn, ascending);
        }

        private void SortSheet(Excel.Worksheet worksheet, int firstRow, string lastColumn, string sortColumn, Tuple<string,string>[] stringsToStopAtWithColumn, bool ascending)
        {
            Excel.Range startOfRange = worksheet.Range[String.Format("A{0}", firstRow)];
            Excel.Range endOfRange = worksheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            if (!string.IsNullOrWhiteSpace(lastColumn))
            {
                endOfRange = worksheet.Cells[endOfRange.Row, lastColumn];
            }
            if (stringsToStopAtWithColumn != null)
            {
                foreach (Tuple<string, string> stringToStopAtWithColumn in stringsToStopAtWithColumn)
                {            
                    Excel.Range startOfSearchRange = startOfRange;
                    Excel.Range endOfSearchRange = endOfRange;
                    string stringToStopAtColumn = stringToStopAtWithColumn.Item2;
                    string stringToStopAt = stringToStopAtWithColumn.Item1;
                    if (!string.IsNullOrWhiteSpace(stringToStopAtColumn))
                    {
                        startOfSearchRange = worksheet.Range[String.Format("{0}{1}", stringToStopAtColumn, firstRow)];
                        endOfSearchRange = worksheet.Range[String.Format("{0}{1}", stringToStopAtColumn, endOfRange.Row)];
                    }
                    Excel.Range searchRange = worksheet.Range[startOfSearchRange,endOfSearchRange];
                    Excel.Range firstRecordFoundInSearchRange = searchRange.Find(stringToStopAt);
                    if (firstRecordFoundInSearchRange != null)
                    {
                        endOfRange = worksheet.Cells[firstRecordFoundInSearchRange.Row-1, endOfRange.Column];
                        break;
                    }
                }                
            }
            Excel.Range range = worksheet.Range[startOfRange, endOfRange];

            string firstCellSortRange = String.Format("{0}{1}", sortColumn,firstRow);                       
            Excel.Range sortRange = worksheet.Range[firstCellSortRange, firstCellSortRange];

            Excel.XlSortOrder sortOrder = Excel.XlSortOrder.xlAscending;

            if (!ascending)
            {
                sortOrder = Excel.XlSortOrder.xlDescending;
            }

            range.Sort(sortRange, sortOrder, Orientation: Excel.XlSortOrientation.xlSortColumns);
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
           // Globals.ThisWorkbook.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
            string fileNameWithoutExtension = Save();
            ReplaceFormulasWithTheirValues(MainSheetName);
            ReplaceFormulasWithTheirValues(LongOnlySheetName);
            ReplaceFormulasWithTheirValues(TotalInventorySheetName);
            ReplaceFormulasWithTheirValues(HighQualtityInventorySheetName);
            ReplaceFormulasWithTheirValues(LowQualtityInventorySheetName);

            SortSheet(MainSheetName, 10, null, "BX", new Tuple<string,string>[]{ new Tuple<string,string>("Median valuation of Longs",null)},true);
            SortSheet(LongOnlySheetName, 10, null, "BY", null,true);
            SortSheet(TotalInventorySheetName, 11, null, "BU", new Tuple<string,string>[]{ new Tuple<string,string>("Investment Trusts",null)},true);
            SortSheet(TotalInventorySheetName, 11, null, "BU", new Tuple<string,string>[]{ new Tuple<string,string>("NA","BU"), new Tuple<string,string>("Investment Trusts",null)},false);
            SortSheet(HighQualtityInventorySheetName, 11, null, "BU", null, true);
            SortSheet(HighQualtityInventorySheetName, 11, null, "BU", new Tuple<string,string>[]{ new Tuple<string,string>("NA","BU")}, false);
            SortSheet(LowQualtityInventorySheetName, 11, null, "BU", null, true);
            ReplaceFormulasWithTheirValues(LongOnlySummarySheetName);
            SortSheet(LongOnlySummarySheetName, 5,"D", "D", null, true);
            SortSheet(LongOnlySummarySheetName, 5, "D", "D", new Tuple<string, string>[] { new Tuple<string, string>("NA", "D") }, false);
            ExportToPDF(fileNameWithoutExtension);
            Excel.Worksheet worksheet = Globals.ThisWorkbook.Worksheets[MainSheetName];
            worksheet.Activate();
           // Globals.ThisWorkbook.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            FinalSave();
        }

        private readonly static string MainSheetName = "OAR";
        private readonly static string LongOnlySheetName = "Longonly";
        private readonly static string TotalInventorySheetName = "Total Inventory";
        private readonly static string HighQualtityInventorySheetName = "High quality inventory";
        private readonly static string LowQualtityInventorySheetName = "Low quality inventory";
        private readonly static string LongOnlySummarySheetName = "Long only summary";
        private readonly static string DirectoryForCopies = @"\\oam.odey.com\shared\Share\OUAR valuations\Old Sheets";

        private readonly string[] WorksheetsForPDF = new string[] { MainSheetName, LongOnlySheetName, TotalInventorySheetName, HighQualtityInventorySheetName, LowQualtityInventorySheetName, LongOnlySummarySheetName };

        private string Save()
        {
            string path = Path.Combine(DirectoryForCopies, DateTime.Today.Year.ToString());
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string fileName = Path.Combine(path, DateTime.Today.ToString("yyyy-MM-dd"));
            if (File.Exists(String.Format("{0}.xlsx",fileName)))
            {
                fileName = Path.Combine(path, DateTime.Now.ToString("yyyy-MM-dd-hhmmss"));
            }
            Globals.ThisWorkbook.SaveAs(fileName);
            return fileName;
        }

        private void FinalSave()
        {            
            
            Globals.ThisWorkbook.RemoveCustomization();
            Globals.ThisWorkbook.Save();
        }

        private void ExportToPDF(string fileName)
        {
            List<Excel.Worksheet> worksheetsToUnhide = new List<Excel.Worksheet>();
            foreach (Excel.Worksheet worksheet in Globals.ThisWorkbook.Worksheets)
            {
                if (worksheet.Visible == Excel.XlSheetVisibility.xlSheetVisible && !WorksheetsForPDF.Contains(worksheet.Name))
                {
                    worksheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                    worksheetsToUnhide.Add(worksheet);
                }
            }
            Globals.ThisWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF,fileName);
            foreach (Excel.Worksheet worksheet in worksheetsToUnhide)
            {
                worksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            PortfolioWebClient client = new PortfolioWebClient();
            List<SimplePortfolio> portfolios = client.GetEquityPortfolio(new int[] { 3609, 6253 }, DateTime.Now);
            WritePortfolio(1,"OAR Weightings", "OAR", portfolios);
            WritePortfolio(2,"Long only weightings", "DEVM", portfolios);
        }

        private void WritePortfolio(int rowNumber, string worksheetName, string fundName, List<SimplePortfolio> portfolios)
        {
            Excel.Worksheet worksheet = Globals.ThisWorkbook.Sheets[worksheetName];
            worksheet.Cells.ClearContents();
            foreach (SimplePortfolio portfolio in portfolios.Where(a => a.FundName == fundName))
            {
                worksheet.Cells[rowNumber, 1] = portfolio.BloombergTicker;
                worksheet.Cells[rowNumber, 2] = portfolio.DeltaMarketValuePercentNav/100;
                rowNumber++;
            }
        }
              
    }
}
