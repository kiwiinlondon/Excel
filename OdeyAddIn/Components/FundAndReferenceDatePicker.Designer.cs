namespace OdeyAddIn.Components
{
    partial class FundAndReferenceDatePicker
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
            this.referenceDatePicker1 = new OdeyAddIn.ReferenceDatePicker();
            this.fundPicker1 = new OdeyAddIn.FundPicker();
            this.SuspendLayout();
            // 
            // referenceDatePicker1
            // 
            this.referenceDatePicker1.Location = new System.Drawing.Point(3, 3);
            this.referenceDatePicker1.MaxDate = new System.DateTime(2011, 11, 2, 0, 0, 0, 0);
            this.referenceDatePicker1.Name = "referenceDatePicker1";
            this.referenceDatePicker1.Size = new System.Drawing.Size(114, 20);
            this.referenceDatePicker1.TabIndex = 1;
            this.referenceDatePicker1.Value = new System.DateTime(2011, 11, 2, 0, 0, 0, 0);
            // 
            // fundPicker1
            // 
            this.fundPicker1.DisplayMember = "Name";
            this.fundPicker1.FormattingEnabled = true;
            this.fundPicker1.Location = new System.Drawing.Point(3, 29);
            this.fundPicker1.Name = "fundPicker1";
            this.fundPicker1.Size = new System.Drawing.Size(114, 21);
            this.fundPicker1.TabIndex = 0;
            this.fundPicker1.ValueMember = "FundId";
            // 
            // FundAndReferenceDatePicker
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.referenceDatePicker1);
            this.Controls.Add(this.fundPicker1);
            this.Name = "FundAndReferenceDatePicker";
            this.Size = new System.Drawing.Size(120, 54);
            this.ResumeLayout(false);

        }

        #endregion

        private FundPicker fundPicker1;
        private ReferenceDatePicker referenceDatePicker1;
    }
}
