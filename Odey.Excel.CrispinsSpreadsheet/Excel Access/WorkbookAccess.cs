using XL=Microsoft.Office.Interop.Excel;
using Odey.Framework.Keeley.Entities.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class WorkbookAccess
    {
        public WorkbookAccess(ThisWorkbook workbook)
        {
            _workbook = workbook;
        }

        private ThisWorkbook _workbook;

        public void Save()
        {
            _workbook.Save();
        }

        private Dictionary<string, WorksheetAccess> _worksheets = new Dictionary<string, WorksheetAccess>();

        private static readonly string _bulkLoadTickerWorksheetName = "Sheet1";

        public WorksheetAccess GetBulkLoadTickerWorksheetAccess()
        {
            return null;// GetWorksheetAccess(_bulkLoadTickerWorksheetName);
        }

        //public void WriteDates(DateTime previousReferenceDate, DateTime referenceDate)
        //{
        //    foreach (var worksheet in _worksheets.Where(a=>a.Key!=_bulkLoadTickerWorksheetName).Select(a=>a.Value))
        //    {
        //        worksheet.WriteDates( previousReferenceDate, referenceDate);
        //    }
        //}

       

        public WorksheetAccess GetWorksheetAccess(Fund fund)
        {
            string sheetName = fund.Name;
            WorksheetAccess worksheetAccess;
            if (!_worksheets.TryGetValue(fund.Name, out worksheetAccess))
            {                
                worksheetAccess = WorksheetAccessFactory.Instance.Get(_workbook.Sheets[sheetName],fund.IsLongOnly,fund.ChildEntityType== EntityTypes.Book);
                worksheetAccess.SetupSheet();
                _worksheets.Add(sheetName, worksheetAccess);
            }
            return worksheetAccess;
        }

        public void DisableCalculations()
        {
            _workbook.Application.Calculation = XL.XlCalculation.xlCalculationManual;
            _workbook.Application.ScreenUpdating = false;
            _workbook.Application.EnableEvents = false;
        }

        public void EnableCalculations()
        {

            _workbook.Application.Calculation = XL.XlCalculation.xlCalculationAutomatic;
            _workbook.Application.ScreenUpdating = true;
            _workbook.Application.EnableEvents = true;
        }

    }
}
