namespace OdeyAddIn
{
    partial class CurrencyControlPane
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
            this.equityPicker1 = new OdeyAddIn.Components.EquityPicker();
            this.button1 = new System.Windows.Forms.Button();
            this.grossNetPicker1 = new OdeyAddIn.Components.GrossNetPicker();
            this.SuspendLayout();
            // 
            // fundAndReferenceDatePicker1
            // 
            this.fundAndReferenceDatePicker1.Location = new System.Drawing.Point(3, 4);
            this.fundAndReferenceDatePicker1.Name = "fundAndReferenceDatePicker1";
            this.fundAndReferenceDatePicker1.Size = new System.Drawing.Size(120, 134);
            this.fundAndReferenceDatePicker1.TabIndex = 0;
            // 
            // equityPicker1
            // 
            this.equityPicker1.Location = new System.Drawing.Point(6, 135);
            this.equityPicker1.Name = "equityPicker1";
            this.equityPicker1.Size = new System.Drawing.Size(140, 28);
            this.equityPicker1.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(30, 252);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "Get";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // grossNetPicker1
            // 
            this.grossNetPicker1.Location = new System.Drawing.Point(6, 169);
            this.grossNetPicker1.Name = "grossNetPicker1";
            this.grossNetPicker1.Size = new System.Drawing.Size(131, 67);
            this.grossNetPicker1.TabIndex = 3;
            // 
            // CurrencyControlPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.grossNetPicker1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.equityPicker1);
            this.Controls.Add(this.fundAndReferenceDatePicker1);
            this.Name = "CurrencyControlPane";
            this.Size = new System.Drawing.Size(140, 331);
            this.ResumeLayout(false);

        }

        #endregion

        private Components.FundAndReferenceDatePicker fundAndReferenceDatePicker1;
        private Components.EquityPicker equityPicker1;
        private System.Windows.Forms.Button button1;
        private Components.GrossNetPicker grossNetPicker1;
    }
}
