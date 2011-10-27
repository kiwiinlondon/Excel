namespace OdeyAddIn
{
    partial class IndustryControlPane
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
            // IndustryControlPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.referenceDatePicker);
            this.Name = "IndustryControlPane";
            this.Size = new System.Drawing.Size(129, 136);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DateTimePicker referenceDatePicker;
        private System.Windows.Forms.ComboBox comboBox1;
    }
}
