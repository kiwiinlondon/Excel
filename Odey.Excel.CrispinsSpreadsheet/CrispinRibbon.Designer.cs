namespace Odey.Excel.CrispinsSpreadsheet
{
    partial class CrispinRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CrispinRibbon()
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
            this.CrispinTab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.CrispinTab.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // CrispinTab
            // 
            this.CrispinTab.Groups.Add(this.group1);
            this.CrispinTab.Label = "Crispin";
            this.CrispinTab.Name = "CrispinTab";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Label = "Crsipin";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.Label = "Update";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // CrispinRibbon
            // 
            this.Name = "CrispinRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.CrispinTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.CrispinRibbon_Load);
            this.CrispinTab.ResumeLayout(false);
            this.CrispinTab.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab CrispinTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal CrispinRibbon CrispinRibbon
        {
            get { return this.GetRibbon<CrispinRibbon>(); }
        }
    }
}
