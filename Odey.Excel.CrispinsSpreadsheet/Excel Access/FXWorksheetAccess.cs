using Odey.Framework.Keeley.Entities.Enums;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XL = Microsoft.Office.Interop.Excel;

namespace Odey.Excel.CrispinsSpreadsheet
{
    public class FXWorksheetAccess
    {
        public FXWorksheetAccess(XL.Worksheet worksheet)
        {
            _worksheet = worksheet;
        }
        private XL.Worksheet _worksheet;


        private int GetCurrenyOrdering(int currencyId)
        {
            switch ((CurrencyIds)currencyId)
            {
                case CurrencyIds.EUR:
                    return 0;
                case CurrencyIds.GBP:
                    return 1;
                case CurrencyIds.USD:
                    return 2;
                default:
                    return currencyId;

            }
        }

        public void Write(List<Fund> funds, FundIds[] fundIdsToWrite)
        {

            var currencyAndLabels = new Dictionary<(int CurrencyId, string IsoCode), HashSet<(string Label, FXExposureTypeIds ExposureTypeId)>>();
            var exposuresByCurrencyLabelAndFund = new Dictionary<(int CurrencyId, string Label), Dictionary<int, FXExposure>>();
            var fundsToWrite = new List<(int FundId, string Name, int ColumnNumber)>();
            int columnNumber = 3;
            foreach (var fund in funds.OrderBy(a=>a.FundId))
            {
                if (fundIdsToWrite.Contains((FundIds)fund.FundId))
                {
                    fundsToWrite.Add((fund.FundId, fund.Name,columnNumber++));
                    var exposures = fund.FXExposureManager.GetFXExposureOnReferenceDate();

                    foreach (var exposure in exposures)
                    {

                        var currency = (exposure.CurrencyId, exposure.Currency.IsoCode);
                        if (!currencyAndLabels.TryGetValue(currency, out var labels))
                        {
                            labels = new HashSet<(string, FXExposureTypeIds)>();
                            currencyAndLabels.Add(currency, labels);
                        }
                        var label = (exposure.Label, exposure.FXExposureTypeId);
                        if (!labels.Contains(label))
                        {
                            labels.Add(label);
                        }

                        var key = (exposure.CurrencyId, exposure.Label);

                        if (!exposuresByCurrencyLabelAndFund.TryGetValue(key, out var exposuresByFund))
                        {
                            exposuresByFund = new Dictionary<int, FXExposure>();
                            exposuresByCurrencyLabelAndFund.Add(key,exposuresByFund);
                        }

                        exposuresByFund.Add(fund.FundId, exposure);
                    }
                }
            }
            _worksheet.Cells.ClearContents();
            _worksheet.Cells.ClearFormats();
            Dictionary<int, string> fundColumns = new Dictionary<int, string>();

            ColorConverter cc = new ColorConverter();
            var headingcolour = ColorTranslator.ToOle((Color)cc.ConvertFromString("#2F75B5"));
            var lastColumn = "";
            _worksheet.Columns[1].ColumnWidth = 2;
            _worksheet.Columns[2].ColumnWidth = 16;
            foreach (var fund in fundsToWrite)
            {
                var cell = _worksheet.Cells[1, fund.ColumnNumber];
                cell.Value = fund.Name;
                cell.Font.Bold = true;
                cell.Interior.Color = headingcolour;
                cell.Font.Color = XL.XlRgbColor.rgbWhite;
                cell.HorizontalAlignment = XL.XlHAlign.xlHAlignCenter;
                lastColumn = cell.Address[false, false, XL.XlReferenceStyle.xlA1].Replace("1", "");
                fundColumns.Add(fund.FundId, lastColumn);
                _worksheet.Columns[fund.ColumnNumber].ColumnWidth = 9;
            }
            int rowNumber = 2;
            foreach (var currency in currencyAndLabels.OrderBy(a => GetCurrenyOrdering(a.Key.CurrencyId)))
            {
                var currencyCell = _worksheet.Cells[rowNumber, 1];
                currencyCell.Value = currency.Key.IsoCode;
                currencyCell.Font.Bold = true;

                var currencyRow = _worksheet.get_Range($"A{rowNumber}", $"{lastColumn}{rowNumber}");
                currencyRow.Interior.Color = XL.XlRgbColor.rgbLightGray;
                var firstRowOfCurrency = rowNumber++;
                foreach (var label in currency.Value.OrderBy(a=>a.ExposureTypeId))
                {

                    _worksheet.Cells[rowNumber, 2].Value = label.Label;
                    var exposuresByFund = exposuresByCurrencyLabelAndFund[(currency.Key.CurrencyId, label.Label)];

                    foreach (var fund in fundsToWrite)
                    {
                        if (exposuresByFund.TryGetValue(fund.FundId, out var exposureForFund))
                        {
                            var exposureCell = _worksheet.Cells[rowNumber, fund.ColumnNumber];
                            exposureCell.Value = exposureForFund.MarketValue / exposureForFund.FundNav;
                            exposureCell.NumberFormat = "0%";
                        }
                    }
                    
                    rowNumber++;
                }
                foreach(var fund in fundsToWrite)
                {
                    var fundColumn = fundColumns[fund.FundId];
                    var totalCell = _worksheet.Cells[rowNumber,fund.ColumnNumber];
                    totalCell.Formula = $"=Sum({fundColumn}{firstRowOfCurrency}:{fundColumn}{rowNumber-1})";
                    totalCell.NumberFormat = "0%";
                    totalCell.Font.Bold = true;
                    totalCell.Borders(XL.XlBordersIndex.xlEdgeTop).LineStyle = XL.XlLineStyle.xlContinuous;
                }
                rowNumber++;
            }

        }
        
    }
}

