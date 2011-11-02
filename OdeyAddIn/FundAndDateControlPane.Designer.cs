namespace OdeyAddIn
{
    partial class FundAndDateControlPane
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
            this.referenceDatePicker = new System.Windows.Forms.DateTimePicker();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // referenceDatePicker
            // 
            this.referenceDatePicker.Location = new System.Drawing.Point(3, 3);
            this.referenceDatePicker.MaxDate = new System.DateTime(2011, 10, 25, 0, 0, 0, 0);
            this.referenceDatePicker.MinDate = new System.DateTime(1999, 1, 1, 0, 0, 0, 0);
            this.referenceDatePicker.Name = "referenceDatePicker";
            this.referenceDatePicker.Size = new System.Drawing.Size(121, 20);
            this.referenceDatePicker.TabIndex = 0;
            this.referenceDatePicker.Value = new System.DateTime(2011, 10, 25, 0, 0, 0, 0);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(3, 29);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 21);
            this.comboBox1.TabIndex = 1;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(24, 56);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "Get";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // IndustryControlPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.referenceDatePicker);
            this.Name = "IndustryControlPane";
            this.Size = new System.Drawing.Size(129, 136);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DateTimePicker referenceDatePicker;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Button button1;
    }
}
