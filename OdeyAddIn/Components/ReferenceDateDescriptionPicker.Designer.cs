namespace OdeyAddIn.Components
{
    partial class ReferenceDateDescriptionPicker
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.periodicityPicker1 = new OdeyAddIn.Components.PeriodicityPicker();
            this.referenceDatePicker2 = new OdeyAddIn.ReferenceDatePicker();
            this.referenceDatePicker1 = new OdeyAddIn.ReferenceDatePicker();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(122, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(30, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "From";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(122, 34);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(20, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "To";
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(154, 7);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(48, 17);
            this.checkBox1.TabIndex = 5;
            this.checkBox1.Text = "BOT";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // periodicityPicker1
            // 
            this.periodicityPicker1.DisplayMember = "Name";
            this.periodicityPicker1.FormattingEnabled = true;
            this.periodicityPicker1.Location = new System.Drawing.Point(4, 57);
            this.periodicityPicker1.Name = "periodicityPicker1";
            this.periodicityPicker1.Size = new System.Drawing.Size(112, 21);
            this.periodicityPicker1.TabIndex = 4;
            this.periodicityPicker1.ValueMember = "PeriodicityId";
            // 
            // referenceDatePicker2
            // 
            this.referenceDatePicker2.Location = new System.Drawing.Point(4, 30);
            this.referenceDatePicker2.MinDate = new System.DateTime(1999, 7, 30, 0, 0, 0, 0);
            this.referenceDatePicker2.Name = "referenceDatePicker2";
            this.referenceDatePicker2.Size = new System.Drawing.Size(112, 20);
            this.referenceDatePicker2.TabIndex = 2;
            // 
            // referenceDatePicker1
            // 
            this.referenceDatePicker1.Location = new System.Drawing.Point(4, 4);
            this.referenceDatePicker1.MinDate = new System.DateTime(1999, 7, 30, 0, 0, 0, 0);
            this.referenceDatePicker1.Name = "referenceDatePicker1";
            this.referenceDatePicker1.Size = new System.Drawing.Size(112, 20);
            this.referenceDatePicker1.TabIndex = 0;
            // 
            // ReferenceDateDescriptionPicker
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.periodicityPicker1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.referenceDatePicker2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.referenceDatePicker1);
            this.Name = "ReferenceDateDescriptionPicker";
            this.Size = new System.Drawing.Size(202, 82);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ReferenceDatePicker referenceDatePicker1;
        private System.Windows.Forms.Label label1;
        private ReferenceDatePicker referenceDatePicker2;
        private System.Windows.Forms.Label label2;
        private PeriodicityPicker periodicityPicker1;
        private System.Windows.Forms.CheckBox checkBox1;
    }
}
