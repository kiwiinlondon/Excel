﻿using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using Odey.Framework.Keeley.Entities.Enums;
using Odey.PortfolioCache.Entities;
using System;
using System.Diagnostics;

namespace Odey.ExcelAddin
{
    public class WatchListItem
    {
        public int RowIndex { get; set; }
        public string Ticker { get; set; }
        public string Quality { get; set; }
        public string JHManagerOverride { get; set; }
        public double? Upside { get; set; }
    }

    class WatchListSheet
    {
        public static string Name = "Watch List";
        public static int HeaderRow = 5;
        public static int FirstColumn = 1;
        public static int NumItems = 30;

        public static ColumnDef Ticker = new ColumnDef { Index = 1, AlphabeticalIndex = "A", Name = "Ticker" };
        public static ColumnDef Upside = new ColumnDef { Index = 20, AlphabeticalIndex = "T", Name = "Upside" };
        public static ColumnDef AverageVolume = new ColumnDef { Index = 47, AlphabeticalIndex = "AU", Name = "Average volume all exchanges 3m" };
        public static ColumnDef Quality = new ColumnDef { Index = 49, AlphabeticalIndex = "AW", Name = "High (H) or Low (L) Quality?" };
        public static ColumnDef Manager = new ColumnDef { Index = 50, AlphabeticalIndex = "AX", Name = "Portfolio Manager" };
        public static ColumnDef Conviction = new ColumnDef { Index = 51, AlphabeticalIndex = "AY", Name = "Conviction Level" };

        public static Dictionary<string, WatchListItem> GetWatchList(Excel.Application app, List<PortfolioDTO> data)
        {
            app.StatusBar = "Reading watch list...";

            Excel.Worksheet sheet;
            try
            {
                sheet = app.Sheets[Name];
            }
            catch
            {
                sheet = app.Sheets.Add(Before: app.Sheets[1]);
                sheet.Name = Name;
            }

            // Read existing tickers
            var watchList = new Dictionary<string, WatchListItem>(StringComparer.OrdinalIgnoreCase);
            var row = HeaderRow + 1;
            var ticker = sheet.Cells[row, Ticker.Index.Value].Value2 as string;
            while (ticker != null)
            {
                if (watchList.ContainsKey(ticker))
                {
                    throw new Exception($"Duplicate watch list entry: \"{ticker}\".\n\nPlease remove all but one.");
                }
                watchList.Add(ticker, new WatchListItem
                {
                    RowIndex = row,
                    Ticker = ticker,
                    Quality = sheet.Cells[row, Quality.Index.Value].Value2 as string,
                    JHManagerOverride = sheet.Cells[row, Manager.Index.Value].Value2 as string,
                    Upside = sheet.Cells[row, Upside.Index.Value].Value2 as double?,
                });

                ++row;
                ticker = sheet.Cells[row, Ticker.Index.Value].Value2 as string;
            }

            // Add new tickers
            var newTickers = data.Select(p => p.BloombergTicker).Distinct().Where(t => t != null).Except(watchList.Keys, StringComparer.OrdinalIgnoreCase).OrderBy(t => t).ToList();
            foreach (var newTicker in newTickers)
            {
                watchList.Add(newTicker, new WatchListItem
                {
                    RowIndex = row,
                    Ticker = newTicker,
                });
                sheet.Cells[row, Ticker.Index.Value] = newTicker;
                ++row;
            }

            return watchList;
        }

        public static void Write(Excel.Application app, Dictionary<string, WatchListItem> watchList, string sheetName, bool descending, string onlyQuality = null)
        {
            var columnList = new[] { "B", "E", "F", "S", "T", "U", "W", "Z", "AD", "AI", "AM", "AQ", "AS", "AT", "AW", "BK" };

            // Query the data
            var rows = watchList.Values.Where(w => w.Upside.HasValue);
            if (onlyQuality != null)
            {
                rows = rows.Where(w => w.Quality == onlyQuality);
            }
            if (descending)
            {
                rows = rows.OrderByDescending(w => w.Upside);
            }
            else
            {
                rows = rows.OrderBy(w => w.Upside);
            }
            rows = rows.Take(NumItems);

            // Get the worksheet
            var isNewSheet = false;
            Excel.Worksheet sheet;
            try
            {
                sheet = app.Sheets[sheetName];
            }
            catch
            {
                isNewSheet = true;
                sheet = app.Sheets.Add(After: app.Sheets[Name]);
                sheet.Name = sheetName;
            }

            // Start
            var y = HeaderRow;
            if (isNewSheet)
            {
                // Format header
                Excel.Range headerRange = sheet.Range[sheet.Cells[y, 1], sheet.Cells[y, 1 + columnList.Length]];
                headerRange.WrapText = true;
                headerRange.RowHeight = 75;
                headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            else
            {
                // Clear data
                Excel.Range r = sheet.Range[sheet.Cells[y, 1], sheet.Cells[y + NumItems, 1 + columnList.Length]];
                r.ClearContents();
            }
            
            // Write header
            Excel.Range cell = sheet.Cells[y, 1];
            cell.Value = "Ticker";
            cell.ColumnWidth = 14;
            var x = 2;
            foreach (var columnIndex in columnList)
            {
                cell = sheet.Cells[y, x];
                cell.Formula = $"='{Name}'!{columnIndex}{5}";
                cell.ColumnWidth = 14;
                ++x;
            }
            ++y;

            // Write content
            foreach (var row in rows)
            {
                sheet.Cells[y, 1] = row.Ticker;
                x = 2;
                foreach (var columnIndex in columnList)
                {
                    sheet.Cells[y, x].Formula = $"='{Name}'!{columnIndex}{row.RowIndex}";
                    ++x;
                }
                ++y;
            }
        }

        public static void ReadColumns(Excel.Application app, List<ColumnDef> columns)
        {
            Excel.Worksheet sheet = app.Sheets[Name];
            Excel.Range tickerCell = sheet.Cells[HeaderRow + 1, 1];
            var tickerAddress = tickerCell.Address[false, true];

            foreach (var column in columns)
            {
                if (column.AlphabeticalIndex != null)
                {
                    column.Name = sheet.Cells[HeaderRow, column.AlphabeticalIndex].Value2;
                    Excel.Range data = sheet.Cells[HeaderRow + 1, column.AlphabeticalIndex];
                    column.NumberFormat = data.NumberFormat;
                    column.Width = data.ColumnWidth;
                }
                if (column.CopyFormula)
                {
                    Excel.Range data = sheet.Cells[HeaderRow + 1, column.AlphabeticalIndex];
                    column.Formula = (data.Formula as string).Replace(tickerAddress, "[Ticker]");
                }
            }
        }
    }
}