namespace OdeyAddIn
{
    partial class PortfolioControlPane
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.fundAndReferenceDatePicker1 = new OdeyAddIn.Components.FundAndReferenceDatePicker();
            this.button1 = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.ReferenceDate = new System.Windows.Forms.CheckBox();
            this.InstrumentName = new System.Windows.Forms.CheckBox();
            this.UnderlyingInstrumentName = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.UnderlyingCountry = new System.Windows.Forms.CheckBox();
            this.Country = new System.Windows.Forms.CheckBox();
            this.UnderlyingSector = new System.Windows.Forms.CheckBox();
            this.UnderlyingParentInstrumentClass = new System.Windows.Forms.CheckBox();
            this.ParentInstrumentClass = new System.Windows.Forms.CheckBox();
            this.ExchangeCode = new System.Windows.Forms.CheckBox();
            this.InstrumentClass = new System.Windows.Forms.CheckBox();
            this.Industry = new System.Windows.Forms.CheckBox();
            this.UnderlyingInstrumentClass = new System.Windows.Forms.CheckBox();
            this.Ticker = new System.Windows.Forms.CheckBox();
            this.UnderlyingIndustry = new System.Windows.Forms.CheckBox();
            this.Sector = new System.Windows.Forms.CheckBox();
            this.UnderlyingTicker = new System.Windows.Forms.CheckBox();
            this.NetPosition = new System.Windows.Forms.CheckBox();
            this.MarketValue = new System.Windows.Forms.CheckBox();
            this.DeltaMarketValue = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // fundAndReferenceDatePicker1
            // 
            this.fundAndReferenceDatePicker1.Location = new System.Drawing.Point(4, 4);
            this.fundAndReferenceDatePicker1.Name = "fundAndReferenceDatePicker1";
            this.fundAndReferenceDatePicker1.Size = new System.Drawing.Size(120, 54);
            this.fundAndReferenceDatePicker1.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(7, 83);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(115, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Get";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(7, 60);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(69, 17);
            this.checkBox1.TabIndex = 2;
            this.checkBox1.Text = "All Funds";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // ReferenceDate
            // 
            this.ReferenceDate.AutoSize = true;
            this.ReferenceDate.Location = new System.Drawing.Point(7, 137);
            this.ReferenceDate.Name = "ReferenceDate";
            this.ReferenceDate.Size = new System.Drawing.Size(99, 17);
            this.ReferenceDate.TabIndex = 3;
            this.ReferenceDate.Text = "ReferenceDate";
            this.ReferenceDate.UseVisualStyleBackColor = true;
            // 
            // InstrumentName
            // 
            this.InstrumentName.AutoSize = true;
            this.InstrumentName.Location = new System.Drawing.Point(7, 160);
            this.InstrumentName.Name = "InstrumentName";
            this.InstrumentName.Size = new System.Drawing.Size(103, 17);
            this.InstrumentName.TabIndex = 4;
            this.InstrumentName.Text = "InstrumentName";
            this.InstrumentName.UseVisualStyleBackColor = true;
            // 
            // UnderlyingInstrumentName
            // 
            this.UnderlyingInstrumentName.AutoSize = true;
            this.UnderlyingInstrumentName.Location = new System.Drawing.Point(7, 343);
            this.UnderlyingInstrumentName.Name = "UnderlyingInstrumentName";
            this.UnderlyingInstrumentName.Size = new System.Drawing.Size(159, 17);
            this.UnderlyingInstrumentName.TabIndex = 5;
            this.UnderlyingInstrumentName.Text = "Underlying Instrument Name";
            this.UnderlyingInstrumentName.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(4, 115);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(157, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Fields to Include in Output";
            // 
            // UnderlyingCountry
            // 
            this.UnderlyingCountry.AutoSize = true;
            this.UnderlyingCountry.Location = new System.Drawing.Point(7, 435);
            this.UnderlyingCountry.Name = "UnderlyingCountry";
            this.UnderlyingCountry.Size = new System.Drawing.Size(115, 17);
            this.UnderlyingCountry.TabIndex = 7;
            this.UnderlyingCountry.Text = "Underlying Country";
            this.UnderlyingCountry.UseVisualStyleBackColor = true;
            // 
            // Country
            // 
            this.Country.AutoSize = true;
            this.Country.Location = new System.Drawing.Point(7, 274);
            this.Country.Name = "Country";
            this.Country.Size = new System.Drawing.Size(62, 17);
            this.Country.TabIndex = 8;
            this.Country.Text = "Country";
            this.Country.UseVisualStyleBackColor = true;
            // 
            // UnderlyingSector
            // 
            this.UnderlyingSector.AutoSize = true;
            this.UnderlyingSector.Location = new System.Drawing.Point(7, 481);
            this.UnderlyingSector.Name = "UnderlyingSector";
            this.UnderlyingSector.Size = new System.Drawing.Size(110, 17);
            this.UnderlyingSector.TabIndex = 9;
            this.UnderlyingSector.Text = "Underlying Sector";
            this.UnderlyingSector.UseVisualStyleBackColor = true;
            // 
            // UnderlyingParentInstrumentClass
            // 
            this.UnderlyingParentInstrumentClass.AutoSize = true;
            this.UnderlyingParentInstrumentClass.Location = new System.Drawing.Point(7, 412);
            this.UnderlyingParentInstrumentClass.Name = "UnderlyingParentInstrumentClass";
            this.UnderlyingParentInstrumentClass.Size = new System.Drawing.Size(190, 17);
            this.UnderlyingParentInstrumentClass.TabIndex = 10;
            this.UnderlyingParentInstrumentClass.Text = "Underlying Parent Instrument Class";
            this.UnderlyingParentInstrumentClass.UseVisualStyleBackColor = true;
            // 
            // ParentInstrumentClass
            // 
            this.ParentInstrumentClass.AutoSize = true;
            this.ParentInstrumentClass.Location = new System.Drawing.Point(7, 251);
            this.ParentInstrumentClass.Name = "ParentInstrumentClass";
            this.ParentInstrumentClass.Size = new System.Drawing.Size(137, 17);
            this.ParentInstrumentClass.TabIndex = 11;
            this.ParentInstrumentClass.Text = "Parent Instrument Class";
            this.ParentInstrumentClass.UseVisualStyleBackColor = true;
            // 
            // ExchangeCode
            // 
            this.ExchangeCode.AutoSize = true;
            this.ExchangeCode.Location = new System.Drawing.Point(7, 205);
            this.ExchangeCode.Name = "ExchangeCode";
            this.ExchangeCode.Size = new System.Drawing.Size(102, 17);
            this.ExchangeCode.TabIndex = 12;
            this.ExchangeCode.Text = "Exchange Code";
            this.ExchangeCode.UseVisualStyleBackColor = true;
            // 
            // InstrumentClass
            // 
            this.InstrumentClass.AutoSize = true;
            this.InstrumentClass.Location = new System.Drawing.Point(7, 228);
            this.InstrumentClass.Name = "InstrumentClass";
            this.InstrumentClass.Size = new System.Drawing.Size(103, 17);
            this.InstrumentClass.TabIndex = 13;
            this.InstrumentClass.Text = "Instrument Class";
            this.InstrumentClass.UseVisualStyleBackColor = true;
            // 
            // Industry
            // 
            this.Industry.AutoSize = true;
            this.Industry.Location = new System.Drawing.Point(7, 297);
            this.Industry.Name = "Industry";
            this.Industry.Size = new System.Drawing.Size(63, 17);
            this.Industry.TabIndex = 14;
            this.Industry.Text = "Industry";
            this.Industry.UseVisualStyleBackColor = true;
            // 
            // UnderlyingInstrumentClass
            // 
            this.UnderlyingInstrumentClass.AutoSize = true;
            this.UnderlyingInstrumentClass.Location = new System.Drawing.Point(7, 389);
            this.UnderlyingInstrumentClass.Name = "UnderlyingInstrumentClass";
            this.UnderlyingInstrumentClass.Size = new System.Drawing.Size(156, 17);
            this.UnderlyingInstrumentClass.TabIndex = 15;
            this.UnderlyingInstrumentClass.Text = "Underlying Instrument Class";
            this.UnderlyingInstrumentClass.UseVisualStyleBackColor = true;
            // 
            // Ticker
            // 
            this.Ticker.AutoSize = true;
            this.Ticker.Location = new System.Drawing.Point(7, 183);
            this.Ticker.Name = "Ticker";
            this.Ticker.Size = new System.Drawing.Size(56, 17);
            this.Ticker.TabIndex = 16;
            this.Ticker.Text = "Ticker";
            this.Ticker.UseVisualStyleBackColor = true;
            // 
            // UnderlyingIndustry
            // 
            this.UnderlyingIndustry.AutoSize = true;
            this.UnderlyingIndustry.Location = new System.Drawing.Point(7, 458);
            this.UnderlyingIndustry.Name = "UnderlyingIndustry";
            this.UnderlyingIndustry.Size = new System.Drawing.Size(116, 17);
            this.UnderlyingIndustry.TabIndex = 17;
            this.UnderlyingIndustry.Text = "Underlying Industry";
            this.UnderlyingIndustry.UseVisualStyleBackColor = true;
            // 
            // Sector
            // 
            this.Sector.AutoSize = true;
            this.Sector.Location = new System.Drawing.Point(7, 320);
            this.Sector.Name = "Sector";
            this.Sector.Size = new System.Drawing.Size(57, 17);
            this.Sector.TabIndex = 18;
            this.Sector.Text = "Sector";
            this.Sector.UseVisualStyleBackColor = true;
            // 
            // UnderlyingTicker
            // 
            this.UnderlyingTicker.AutoSize = true;
            this.UnderlyingTicker.Location = new System.Drawing.Point(7, 366);
            this.UnderlyingTicker.Name = "UnderlyingTicker";
            this.UnderlyingTicker.Size = new System.Drawing.Size(104, 17);
            this.UnderlyingTicker.TabIndex = 19;
            this.UnderlyingTicker.Text = "Underlyer Ticker";
            this.UnderlyingTicker.UseVisualStyleBackColor = true;
            // 
            // NetPosition
            // 
            this.NetPosition.AutoSize = true;
            this.NetPosition.Location = new System.Drawing.Point(7, 505);
            this.NetPosition.Name = "NetPosition";
            this.NetPosition.Size = new System.Drawing.Size(83, 17);
            this.NetPosition.TabIndex = 20;
            this.NetPosition.Text = "Net Position";
            this.NetPosition.UseVisualStyleBackColor = true;
            // 
            // MarketValue
            // 
            this.MarketValue.AutoSize = true;
            this.MarketValue.Location = new System.Drawing.Point(7, 529);
            this.MarketValue.Name = "MarketValue";
            this.MarketValue.Size = new System.Drawing.Size(89, 17);
            this.MarketValue.TabIndex = 21;
            this.MarketValue.Text = "Market Value";
            this.MarketValue.UseVisualStyleBackColor = true;
            // 
            // DeltaMarketValue
            // 
            this.DeltaMarketValue.AutoSize = true;
            this.DeltaMarketValue.Location = new System.Drawing.Point(7, 553);
            this.DeltaMarketValue.Name = "DeltaMarketValue";
            this.DeltaMarketValue.Size = new System.Drawing.Size(117, 17);
            this.DeltaMarketValue.TabIndex = 22;
            this.DeltaMarketValue.Text = "Delta Market Value";
            this.DeltaMarketValue.UseVisualStyleBackColor = true;
            // 
            // PortfolioControlPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.DeltaMarketValue);
            this.Controls.Add(this.MarketValue);
            this.Controls.Add(this.NetPosition);
            this.Controls.Add(this.UnderlyingTicker);
            this.Controls.Add(this.Sector);
            this.Controls.Add(this.UnderlyingIndustry);
            this.Controls.Add(this.Ticker);
            this.Controls.Add(this.UnderlyingInstrumentClass);
            this.Controls.Add(this.Industry);
            this.Controls.Add(this.InstrumentClass);
            this.Controls.Add(this.ExchangeCode);
            this.Controls.Add(this.ParentInstrumentClass);
            this.Controls.Add(this.UnderlyingParentInstrumentClass);
            this.Controls.Add(this.UnderlyingSector);
            this.Controls.Add(this.Country);
            this.Controls.Add(this.UnderlyingCountry);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.UnderlyingInstrumentName);
            this.Controls.Add(this.InstrumentName);
            this.Controls.Add(this.ReferenceDate);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.fundAndReferenceDatePicker1);
            this.Name = "PortfolioControlPane";
            this.Size = new System.Drawing.Size(197, 651);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Components.FundAndReferenceDatePicker fundAndReferenceDatePicker1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.CheckBox ReferenceDate;
        private System.Windows.Forms.CheckBox InstrumentName;
        private System.Windows.Forms.CheckBox UnderlyingInstrumentName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox UnderlyingCountry;
        private System.Windows.Forms.CheckBox Country;
        private System.Windows.Forms.CheckBox UnderlyingSector;
        private System.Windows.Forms.CheckBox UnderlyingParentInstrumentClass;
        private System.Windows.Forms.CheckBox ParentInstrumentClass;
        private System.Windows.Forms.CheckBox ExchangeCode;
        private System.Windows.Forms.CheckBox InstrumentClass;
        private System.Windows.Forms.CheckBox Industry;
        private System.Windows.Forms.CheckBox UnderlyingInstrumentClass;
        private System.Windows.Forms.CheckBox Ticker;
        private System.Windows.Forms.CheckBox UnderlyingIndustry;
        private System.Windows.Forms.CheckBox Sector;
        private System.Windows.Forms.CheckBox UnderlyingTicker;
        private System.Windows.Forms.CheckBox NetPosition;
        private System.Windows.Forms.CheckBox MarketValue;
        private System.Windows.Forms.CheckBox DeltaMarketValue;
    }
}
