namespace OdeyAddIn
{
    partial class CountryControlPane
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
            this.equityPicker1 = new OdeyAddIn.Components.EquityPicker();
            this.grossNetPicker1 = new OdeyAddIn.Components.GrossNetPicker();
            this.SuspendLayout();
            // 
            // fundAndReferenceDatePicker1
            // 
            this.fundAndReferenceDatePicker1.Location = new System.Drawing.Point(4, 4);
            this.fundAndReferenceDatePicker1.Name = "fundAndReferenceDatePicker1";
            this.fundAndReferenceDatePicker1.Size = new System.Drawing.Size(120, 127);
            this.fundAndReferenceDatePicker1.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(20, 270);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "Get";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(7, 164);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(88, 17);
            this.checkBox1.TabIndex = 5;
            this.checkBox1.Text = "Include Cash";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // equityPicker1
            // 
            this.equityPicker1.Location = new System.Drawing.Point(7, 137);
            this.equityPicker1.Name = "equityPicker1";
            this.equityPicker1.Size = new System.Drawing.Size(134, 21);
            this.equityPicker1.TabIndex = 6;
            // 
            // grossNetPicker1
            // 
            this.grossNetPicker1.Location = new System.Drawing.Point(5, 187);
            this.grossNetPicker1.Name = "grossNetPicker1";
            this.grossNetPicker1.Size = new System.Drawing.Size(131, 67);
            this.grossNetPicker1.TabIndex = 7;
            // 
            // CountryControlPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.grossNetPicker1);
            this.Controls.Add(this.equityPicker1);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.fundAndReferenceDatePicker1);
            this.Name = "CountryControlPane";
            this.Size = new System.Drawing.Size(136, 334);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Components.FundAndReferenceDatePicker fundAndReferenceDatePicker1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox checkBox1;
        private Components.EquityPicker equityPicker1;
        private Components.GrossNetPicker grossNetPicker1;

    }
}
