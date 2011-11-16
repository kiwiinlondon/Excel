namespace OdeyAddIn
{
    partial class TopHoldingsControlPane
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.buttn1 = new System.Windows.Forms.Button();
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
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(5, 137);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(30, 20);
            this.textBox1.TabIndex = 1;
            this.textBox1.Text = "10";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(40, 140);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Number of Results";
            // 
            // buttn1
            // 
            this.buttn1.Location = new System.Drawing.Point(32, 271);
            this.buttn1.Name = "buttn1";
            this.buttn1.Size = new System.Drawing.Size(75, 23);
            this.buttn1.TabIndex = 3;
            this.buttn1.Text = "Get";
            this.buttn1.UseVisualStyleBackColor = true;
            this.buttn1.Click += new System.EventHandler(this.buttn1_Click);
            // 
            // equityPicker1
            // 
            this.equityPicker1.Location = new System.Drawing.Point(5, 164);
            this.equityPicker1.Name = "equityPicker1";
            this.equityPicker1.Size = new System.Drawing.Size(140, 28);
            this.equityPicker1.TabIndex = 4;
            // 
            // grossNetPicker1
            // 
            this.grossNetPicker1.Location = new System.Drawing.Point(3, 191);
            this.grossNetPicker1.Name = "grossNetPicker1";
            this.grossNetPicker1.Size = new System.Drawing.Size(131, 67);
            this.grossNetPicker1.TabIndex = 5;
            // 
            // TopHoldingsControlPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.grossNetPicker1);
            this.Controls.Add(this.equityPicker1);
            this.Controls.Add(this.buttn1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.fundAndReferenceDatePicker1);
            this.Name = "TopHoldingsControlPane";
            this.Size = new System.Drawing.Size(150, 309);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Components.FundAndReferenceDatePicker fundAndReferenceDatePicker1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttn1;
        private Components.EquityPicker equityPicker1;
        private Components.GrossNetPicker grossNetPicker1;
    }
}
