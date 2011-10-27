namespace OdeyAddIn
{
    partial class OdeyRibbonTab : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public OdeyRibbonTab()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OdeyRibbonTab));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Odey = this.Factory.CreateRibbonTab();
            this.Portfolio = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.industryButton = this.Factory.CreateRibbonToggleButton();
            this.Country = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Odey.SuspendLayout();
            this.Portfolio.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // Odey
            // 
            this.Odey.Groups.Add(this.Portfolio);
            this.Odey.Label = "Odey";
            this.Odey.Name = "Odey";
            // 
            // Portfolio
            // 
            this.Portfolio.Items.Add(this.industryButton);
            this.Portfolio.Items.Add(this.separator1);
            this.Portfolio.Items.Add(this.Country);
            this.Portfolio.Label = "Portfolio";
            this.Portfolio.Name = "Portfolio";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // industryButton
            // 
            this.industryButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.industryButton.Image = ((System.Drawing.Image)(resources.GetObject("industryButton.Image")));
            this.industryButton.Label = "Industry";
            this.industryButton.Name = "industryButton";
            this.industryButton.ShowImage = true;
            this.industryButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.industryButton_Click);
            // 
            // Country
            // 
            this.Country.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Country.Image = ((System.Drawing.Image)(resources.GetObject("Country.Image")));
            this.Country.Label = "Country";
            this.Country.Name = "Country";
            this.Country.ShowImage = true;
            // 
            // OdeyRibbonTab
            // 
            this.Name = "OdeyRibbonTab";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.Odey);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.OdeyRibbonTab_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Odey.ResumeLayout(false);
            this.Odey.PerformLayout();
            this.Portfolio.ResumeLayout(false);
            this.Portfolio.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab Odey;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Portfolio;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Country;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton industryButton;
    }

    partial class ThisRibbonCollection
    {
        internal OdeyRibbonTab OdeyRibbonTab
        {
            get { return this.GetRibbon<OdeyRibbonTab>(); }
        }
    }
}
