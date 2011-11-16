namespace OdeyAddIn.Components
{
    partial class ReferenceDateDescriptorForm
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.shapeContainer1 = new Microsoft.VisualBasic.PowerPacks.ShapeContainer();
            this.lineShape1 = new Microsoft.VisualBasic.PowerPacks.LineShape();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.referenceDateDescriptionPicker1 = new OdeyAddIn.Components.ReferenceDateDescriptionPicker();
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.multipleReferenceDatePicker1 = new OdeyAddIn.Components.MultipleReferenceDatePicker();
            this.SuspendLayout();
            // 
            // shapeContainer1
            // 
            this.shapeContainer1.Location = new System.Drawing.Point(0, 0);
            this.shapeContainer1.Margin = new System.Windows.Forms.Padding(0);
            this.shapeContainer1.Name = "shapeContainer1";
            this.shapeContainer1.Shapes.AddRange(new Microsoft.VisualBasic.PowerPacks.Shape[] {
            this.lineShape1});
            this.shapeContainer1.Size = new System.Drawing.Size(252, 378);
            this.shapeContainer1.TabIndex = 1;
            this.shapeContainer1.TabStop = false;
            // 
            // lineShape1
            // 
            this.lineShape1.BorderColor = System.Drawing.SystemColors.ControlDarkDark;
            this.lineShape1.Name = "lineShape1";
            this.lineShape1.X1 = 21;
            this.lineShape1.X2 = 222;
            this.lineShape1.Y1 = 200;
            this.lineShape1.Y2 = 200;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Checked = true;
            this.radioButton1.Location = new System.Drawing.Point(13, 16);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(14, 13);
            this.radioButton1.TabIndex = 2;
            this.radioButton1.TabStop = true;
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(13, 231);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(14, 13);
            this.radioButton2.TabIndex = 3;
            this.radioButton2.UseVisualStyleBackColor = true;
            this.radioButton2.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
            // 
            // referenceDateDescriptionPicker1
            // 
            this.referenceDateDescriptionPicker1.Location = new System.Drawing.Point(33, 224);
            this.referenceDateDescriptionPicker1.Name = "referenceDateDescriptionPicker1";
            this.referenceDateDescriptionPicker1.Size = new System.Drawing.Size(181, 96);
            this.referenceDateDescriptionPicker1.TabIndex = 0;
            // 
            // elementHost1
            // 
            this.elementHost1.Location = new System.Drawing.Point(33, 12);
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Size = new System.Drawing.Size(181, 163);
            this.elementHost1.TabIndex = 5;
            this.elementHost1.Text = "elementHost1";
            this.elementHost1.Child = this.multipleReferenceDatePicker1;
            // 
            // ReferenceDateDescriptorForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(252, 378);
            this.Controls.Add(this.elementHost1);
            this.Controls.Add(this.radioButton2);
            this.Controls.Add(this.radioButton1);
            this.Controls.Add(this.referenceDateDescriptionPicker1);
            this.Controls.Add(this.shapeContainer1);
            this.Name = "ReferenceDateDescriptorForm";
            this.Text = "ReferenceDateDescriptorForm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ReferenceDateDescriptorForm_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ReferenceDateDescriptionPicker referenceDateDescriptionPicker1;
        private Microsoft.VisualBasic.PowerPacks.ShapeContainer shapeContainer1;
        private Microsoft.VisualBasic.PowerPacks.LineShape lineShape1;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.Integration.ElementHost elementHost1;
        private MultipleReferenceDatePicker multipleReferenceDatePicker1;
    }
}