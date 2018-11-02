using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Odey.Framework.Keeley.Entities.Enums;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Windows.Forms;
using Odey.Query.Client;
using Odey.Query.Reporting.Contracts;
using Odey.Intranet.Entities.Grid;
using System.Diagnostics;
using Newtonsoft.Json;
using System.Deployment.Application;
using Odey.Query.Contracts;

namespace Odey.ExcelAddin
{
    public class PortfolioItem
    {
        public PortfolioItem(PortfolioItem parent = null)
        {
            if (parent == null)
            {
                return;
            }
            Parent = parent;
            ManagerId = parent.ManagerId;
            Manager = parent.Manager;
            ManagerInitials = parent.ManagerInitials;
            FundId = parent.FundId;
            Fund = parent.Fund;
            BookId = parent.BookId;
            Book = parent.Book;
            IssuerId = parent.IssuerId;
            Issuer = parent.Issuer;
            InstrumentId = parent.InstrumentId;
            Instrument = parent.Instrument;
        }

        public PortfolioItem Parent { get; set; }

        public PortfolioFields Field { get; set; }

        public ApplicationUserIds ManagerId { get; set; }
        public string Manager { get; set; }
        public string ManagerInitials { get; set; }

        public FundIds FundId { get; set; }
        public string Fund { get; set; }

        public BookIds BookId { get; set; }
        public string Book { get; set; }

        public int IssuerId { get; set; }
        public string Issuer { get; set; }

        public int InstrumentId { get; set; }
        public string Instrument { get; set; }

        public string Ticker { get; set; }
        public decimal Exposure { get; set; }
        public decimal NetPosition { get; internal set; }

        public Node Node { get; set; }
    }

    public class ColumnDef
    {
        public int? Index { get; set; }
        public string AlphabeticalIndex { get; set; }

        public string Name { get; set; }
        public string Formula { get; set; }
        public string NumberFormat { get; set; }
        public double Width { get; set; }
        public bool CopyFormula { get; set; }
        public bool RefAsNumber { get; set; }
        public bool RefAsString { get; set; }
    }

    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
#if DEBUG
        public const bool IsDebug = true;
#else
        public const bool IsDebug = false;
#endif

        private Office.IRibbonUI ribbon;

        private static Dictionary<ApplicationUserIds, string> managerInitials = new Dictionary<ApplicationUserIds, string>
        {
            { ApplicationUserIds.AdrianCourtenay, "AC" },
            { ApplicationUserIds.JamieGrimston, "JG" },
            { ApplicationUserIds.JamesHanbury, "JH" },
        };

        private readonly string AddonName;

        public Ribbon1()
        {
            var assemblyName = Assembly.GetExecutingAssembly().GetName().Name;
            var versionString = (ApplicationDeployment.IsNetworkDeployed ? "v" + ApplicationDeployment.CurrentDeployment.CurrentVersion : "(Local)");
            AddonName = $"{assemblyName} {versionString}";
            Debug.WriteLine($"Starting {AddonName}...");
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Odey.ExcelAddin.Ribbon1.xml");
        }

        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        public void OnActionCallback(Office.IRibbonControl control)
        {
            var app = Globals.ThisAddIn.Application;

            var prevScreenUpdating = app.ScreenUpdating;
            var prevEvents = app.EnableEvents;
            var prevCalculation = app.Calculation;
            try
            {
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
            }
            catch
            {
                MessageBox.Show("Please stop editing");
                return;
            }

            try
            {
                app.StatusBar = "Loading portfolio weightings...";

                // Generate request
                var tickerColReq = new ColumnRequest { ColumnRequestId = 2, PortfolioField = PortfolioFields.BloombergTicker };
                var netPosColReq = new ColumnRequest { ColumnRequestId = 3, PortfolioField = PortfolioFields.NetPosition };
                var exposureColReq = new ColumnRequest { ColumnRequestId = 5, PortfolioField = PortfolioFields.Exposure, GrossOrNet = GrossOrNet.Net, ApplyTenYearAdjustment = true };
                var request = new AdhocRequest
                {
                    Dates = new[] { DateTime.Today },
                    Funds = new[] { FundIds.ARFF, FundIds.BVFF, FundIds.DEVM, FundIds.FDXH, FundIds.OUAR }.Cast<int>(),
                    ColumnHierarchy = new[] { ColumnHierarchyTypes.Column, ColumnHierarchyTypes.Fund, ColumnHierarchyTypes.Date },
                    Columns = new[] { tickerColReq, netPosColReq, exposureColReq }.ToList(),
                    TotalFields = new[] { TotalField.Total }.ToList(),
                    PivotFundsAsColumns = true,
                    PropsHierarchy = PropsHierarchyType.Off,
                    IncludeOffsetCash = false,
                    MakeWeightsSumToOne = false,
                    IsTransactionBasedPerformance = false,
                    ShowColumnGroups = false,
                    ProvideEntityIds = true,
                    //CreateCurrencyRows = false,
                    Drilldown = new DrilldownNode
                    {
                        Field = PortfolioFields.Fund,
                        Default = new DrilldownNode
                        {
                            Field = PortfolioFields.Book,
                            Default = new DrilldownNode
                            {
                                Field = PortfolioFields.Manager,
                                Default = new DrilldownNode
                                {
                                    Field = PortfolioFields.Issuer,
                                    Default = new DrilldownNode
                                    {
                                        Field = PortfolioFields.Instrument,
                                    }
                                }
                            }
                        }
                    },
                    CurrencyDrilldown = new DrilldownNode { Field = PortfolioFields.Instrument },
                };

                // Log request
                Debug.WriteLine("Running ExecuteAdhocReport with request:\n" + JsonConvert.SerializeObject(request));

                // Run request
                var response = new ReportClient().ExecuteAdhocReport(request);

                Debug.WriteLine("Received GridResponse.");
                Debug.WriteLine($"ResponseId: {response.GridResponseId}");
                Debug.WriteLine($"Nodes: {response.Nodes?.Count() ?? 0}");
                Debug.WriteLine($"Columns: {response.Columns?.Count() ?? 0}");

                // Check for notifications in the response
                if (response.Notifications != null && response.Notifications.Any())
                {
                    var notification = response.Notifications.First();
                    throw new Exception($"Unexpected {notification.NotificationType}: {notification.Description}");
                }

                // Get ColumnIds
                var tickerColId = response.Columns.Single(c => c.ColumnRequestId == tickerColReq.ColumnRequestId).ColumnId;
                var netPosColId = response.Columns.Single(c => c.ColumnRequestId == netPosColReq.ColumnRequestId).ColumnId;
                var exposureColId = response.Columns.Single(c => c.ColumnRequestId == exposureColReq.ColumnRequestId).ColumnId;

                // Create a map from FundId to FundName
                var funds = response.Nodes.Where(n => n.NodeTypeId == (int)PortfolioFields.Fund && n.LogicalEntityId == (int)LogicalEntities.Fund).GroupBy(n => (FundIds)(n.EntityId ?? -1)).ToDictionary(g => g.Key, g => {
                    var node = g.First();
                    return node.Values[node.HierarchyColumn.Value].ToString();
                });
                if (!funds.Any())
                {
                    throw new Exception("Expected funds in GridResponse");
                }

                // Unroll hierarchy
                var rows = GetFlattenedRows(response.Nodes, tickerColId, exposureColId, netPosColId);

                // Separate array that only contains the leaves of the hierarchy (Instrument nodes)
                var instrumentRows = rows.Where(i => i.Field == PortfolioFields.Instrument).ToArray();
                Debug.WriteLine($"Instrument nodes: {instrumentRows.Length}");

                // Find current tickers
                var currentTickers = instrumentRows.Select(i => i.Ticker).Distinct().Where(t => t != null).ToArray();
                Debug.WriteLine($"Distinct tickers: {currentTickers.Length}");

                // Read watch list sheet into a dictionary (by ticker)
                // And append missing tickers
                var watchList = WatchListSheet.GetWatchList(app, currentTickers);

                // Reads watch list column metadata (formulas, number format, width, etc.)
                // and puts it to the PortfolioSheet class
                WatchListSheet.ReadColumns(app, PortfolioSheet.Columns);

                // Make sure that the AC books' manager is actually Adrian Courtenay
                // (Geoff changed the manager of those books to "James Hanbury", so that
                // the trades appear on his trade blotters to sign off)
                var acBooks = new[] { BookIds.ArffAC, BookIds.BvffAC, BookIds.DevmAC, BookIds.FdxhAC, BookIds.OuarAC }.ToList();
                foreach (var row in rows)
                {
                    if (acBooks.Contains(row.BookId))
                    {
                        row.Manager = "Adrian Courtenay";
                        row.ManagerId = ApplicationUserIds.AdrianCourtenay;
                        row.ManagerInitials = managerInitials[ApplicationUserIds.AdrianCourtenay];
                    }
                }

                // Apply manager overrides that were read from the Watch List
                ApplyManagerOverrides(rows, watchList);

                // Write Excel sheets (the order matters)
                //foreach (var fund in funds)
                //{
                //    ScenarioSheet.Write(app, fund, rows, watchList);
                //}
                foreach (var fund in funds)
                {
                    ExposureSheet.Write(app, request.Dates.First(), fund, rows, watchList);
                }
                //foreach (var fund in funds)
                //{
                //    PortfolioSheet.Write(app, fund, rows, watchList);
                //}
                //WatchListSheet.Write(app, watchList, "Watch List Top", true);
                //WatchListSheet.Write(app, watchList, "Watch List Bottom", false);
                //WatchListSheet.Write(app, watchList, "Watch List High Quality", true, "H");
                //WatchListSheet.Write(app, watchList, "Watch List Low Quality", false, "L");
                Debug.WriteLine("Done");
            }
#if !DEBUG
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), AddonName);
            }
#endif
            finally
            {
                app.StatusBar = null;
                app.EnableEvents = prevEvents;
                app.ScreenUpdating = prevScreenUpdating;
                app.Calculation = prevCalculation;
            }
        }

        private static List<PortfolioItem> GetFlattenedRows(IEnumerable<Node> nodes, uint tickerColumnId, uint exposureColumnId, uint netPositionColumnId, List<PortfolioItem> flattened = null, PortfolioItem parent = null)
        {
            flattened = flattened ?? new List<PortfolioItem>();

            foreach (var node in nodes)
            {
                var current = new PortfolioItem(parent);
                var name = node.Values[node.HierarchyColumn.Value].ToString();
                var field = (PortfolioFields)node.NodeTypeId.Value;
                var id = node.EntityId.Value;
                current.Field = field;
                current.Node = node;
                if (field == PortfolioFields.Book)
                {
                    current.Book = name;
                    current.BookId = (BookIds)id;
                }
                else if (field == PortfolioFields.Fund)
                {
                    current.Fund = name;
                    current.FundId = (FundIds)id;
                }
                else if (field == PortfolioFields.Manager)
                {
                    current.Manager = name;
                    current.ManagerId = (ApplicationUserIds)id;
                    current.ManagerInitials = managerInitials[(ApplicationUserIds)id];
                }
                else if (field == PortfolioFields.Issuer)
                {
                    current.Issuer = name;
                    current.IssuerId = id;
                }
                else if (field == PortfolioFields.Instrument)
                {
                    current.Instrument = name;
                    current.InstrumentId = id;

                    // Read column values as well
                    if (node.Values.ContainsKey(tickerColumnId))
                    {
                        current.Ticker = node.Values[tickerColumnId].ToString(); // Ticker
                    }
                    if (node.Values.ContainsKey(exposureColumnId))
                    {
                        current.Exposure = node.Values[exposureColumnId].NumericValue.Value; // Exposure as % NAV
                    }
                    if (node.Values.ContainsKey(netPositionColumnId))
                    {
                        current.NetPosition = node.Values[netPositionColumnId].NumericValue.Value; // Net Position
                    }
                }
                else
                {
                    throw new NotImplementedException($"The field {field} is not implemented");
                }
                flattened.Add(current);

                if (node.Children != null)
                {
                    GetFlattenedRows(node.Children, tickerColumnId, exposureColumnId, netPositionColumnId, flattened, current);
                }
            }

            return flattened;
        }

        private void ApplyManagerOverrides(IEnumerable<PortfolioItem> items, Dictionary<string, WatchListItem> watchList)
        {
            //var jhSecondaryTickers = items.Where(n => n.Ticker != null && n.ManagerId == ApplicationUserIds.JamesHanbury && (n.FundId == FundIds.DEVM || n.FundId == FundIds.FDXH));
            //var otherTickers = items.Where(n => n.Ticker != null && n.ManagerId != ApplicationUserIds.JamesHanbury && n.FundId != FundIds.DEVM && n.FundId != FundIds.FDXH).ToArray();
            //foreach (var item in jhSecondaryTickers)
            //{
            //    var others = otherTickers.Where(n => n.Ticker == item.Ticker).ToArray();
            //    if (others.Select(p => p.ManagerId).Distinct().Count() == 1)
            //    {
            //        item.ManagerId = others.First().ManagerId;
            //        item.Manager = others.First().Manager;
            //    }
            //}

            // Apply manager override column from the watch list
            foreach (var item in items)
            {
                if (item.Ticker == null || item.ManagerId != ApplicationUserIds.JamesHanbury)
                {
                    continue;
                }
                watchList.TryGetValue(item.Ticker, out var watchListItem);
                if (watchListItem == null || watchListItem.JHManagerOverride == null)
                {
                    continue;
                }
                item.Manager = watchListItem.JHManagerOverride;
                item.ManagerId = (ApplicationUserIds)(-1);
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
