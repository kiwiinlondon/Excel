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
            this.PortfolioGroup = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.industryButton = this.Factory.CreateRibbonToggleButton();
            this.countryButton = this.Factory.CreateRibbonToggleButton();
            this.portfolioButton = this.Factory.CreateRibbonToggleButton();
            this.tab1.SuspendLayout();
            this.Odey.SuspendLayout();
            this.PortfolioGroup.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // Odey
            // 
            this.Odey.Groups.Add(this.PortfolioGroup);
            this.Odey.Label = "Odey";
            this.Odey.Name = "Odey";
            // 
            // PortfolioGroup
            // 
            this.PortfolioGroup.Items.Add(this.industryButton);
            this.PortfolioGroup.Items.Add(this.separator1);
            this.PortfolioGroup.Items.Add(this.countryButton);
            this.PortfolioGroup.Items.Add(this.separator2);
            this.PortfolioGroup.Items.Add(this.portfolioButton);
            this.PortfolioGroup.Label = "Portfolio";
            this.PortfolioGroup.Name = "PortfolioGroup";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
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
            // countryButton
            // 
            this.countryButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.countryButton.Image = global::OdeyAddIn.Properties.Resources._033404_rounded_glossy_black_icon_culture_globe_black_sc48;
            this.countryButton.Label = "Country";
            this.countryButton.Name = "countryButton";
            this.countryButton.ShowImage = true;
            this.countryButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.countryButton_Click);
            // 
            // portfolioButton
            // 
            this.portfolioButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.portfolioButton.Image = global::OdeyAddIn.Properties.Resources._086211_rounded_glossy_black_icon_business_charts1_sc1;
            this.portfolioButton.Label = "Portfolio";
            this.portfolioButton.Name = "portfolioButton";
            this.portfolioButton.ShowImage = true;
            this.portfolioButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.portfolioButton_Click);
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
            this.PortfolioGroup.ResumeLayout(false);
            this.PortfolioGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab Odey;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup PortfolioGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton industryButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton countryButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton portfolioButton;
    }

    partial class ThisRibbonCollection
    {
        internal OdeyRibbonTab OdeyRibbonTab
        {
            get { return this.GetRibbon<OdeyRibbonTab>(); }
        }
    }
}
