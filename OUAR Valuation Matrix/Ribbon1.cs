using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace OUAR_Valuation_Matrix
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ReplaceFormulasWithTheirValuesThenSort(string worksheetName, int firstRow, string lastColumn, string sortColumn, string nextStringPastEndOfRange)
        {
            Excel.Worksheet worksheet = Globals.ThisWorkbook.Sheets[worksheetName];
            ReplaceFormulasWithTheirValues(worksheet);
            SortSheet(worksheet, firstRow, lastColumn, sortColumn, nextStringPastEndOfRange);
        }

        private void ReplaceFormulasWithTheirValues(Excel.Worksheet worksheet)
        {
            worksheet.Cells.Copy();
            worksheet.Cells.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
        }

        private void SortSheet(string worksheetName, int firstRow, string lastColumn, string sortColumn, string nextStringPastEndOfRange)
        {
            Excel.Worksheet worksheet = Globals.ThisWorkbook.Sheets[worksheetName];
            SortSheet(worksheet, firstRow, lastColumn, sortColumn, nextStringPastEndOfRange);
        }

        private void SortSheet(Excel.Worksheet worksheet, int firstRow, string lastColumn, string sortColumn, string nextStringPastEndOfRange)
        {
            string firstCellAddress = String.Format("A{0}", firstRow);

            Excel.Range endOfRange = worksheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);

            Excel.Range range = worksheet.Range[firstCellAddress, endOfRange];

            if (!string.IsNullOrWhiteSpace(nextStringPastEndOfRange))
            {
                Excel.Range bottomOfRange = range.Find(nextStringPastEndOfRange);
                if (bottomOfRange != null)
                {
                    endOfRange = worksheet.Cells[bottomOfRange.Row, endOfRange.Column];
                }
                
            }
            if (!string.IsNullOrWhiteSpace(lastColumn))
            {
                endOfRange = worksheet.Cells[endOfRange.Row, lastColumn];
            }
            range = worksheet.Range[firstCellAddress, endOfRange];
            string firstCellSortRange = String.Format("{0}{1}", sortColumn,firstRow);
                       
            Excel.Range sortRange = worksheet.Range[firstCellSortRange, firstCellSortRange];
            range.Sort(sortRange,Orientation:Excel.XlSortOrientation.xlSortColumns);
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            string fileNameWithoutExtension = Save();
            ReplaceFormulasWithTheirValuesThenSort(MainSheetName, 10, null, "BU", "Median valuation of Longs");
            ReplaceFormulasWithTheirValuesThenSort(LongOnlySheetName, 10, null, "BU", null);
            ReplaceFormulasWithTheirValuesThenSort(TotalInventorySheetName, 11, null, "BU", "Investment Trusts");
            ReplaceFormulasWithTheirValuesThenSort(HighQualtityInventorySheetName, 11, null, "BU", null);
            ReplaceFormulasWithTheirValuesThenSort(LowQualtityInventorySheetName, 11, null, "BU", null);
            SortSheet(LongOnlySummarySheetName, 6,"D", "C",null);
            ExportToPDF(fileNameWithoutExtension);
            Excel.Worksheet worksheet = Globals.ThisWorkbook.Worksheets[MainSheetName];
            worksheet.Cells[1,1].Select();
            FinalSave();
        }

        private readonly static string MainSheetName = "OUAR";
        private readonly static string LongOnlySheetName = "Longonly";
        private readonly static string TotalInventorySheetName = "Total Inventory";
        private readonly static string HighQualtityInventorySheetName = "High quality inventory";
        private readonly static string LowQualtityInventorySheetName = "Low quality inventory";
        private readonly static string LongOnlySummarySheetName = "Long only summary";
        private readonly static string DirectoryForCopies = @"\\oam.odey.com\shared\Share\OUAR valuations2\Old Sheets";

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
              
    }
}
