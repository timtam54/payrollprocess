namespace PayrollProcess
{
    partial class frmPlantJobDate
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.rbMatchOnJobDate = new System.Windows.Forms.RadioButton();
            this.rbMathOnDateOnly = new System.Windows.Forms.RadioButton();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rbMathOnDateOnly);
            this.groupBox1.Controls.Add(this.rbMatchOnJobDate);
            this.groupBox1.Location = new System.Drawing.Point(26, 39);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(416, 192);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Add Plant to Time entries matching time entries as per;";
            // 
            // rbMatchOnJobDate
            // 
            this.rbMatchOnJobDate.AutoSize = true;
            this.rbMatchOnJobDate.Checked = true;
            this.rbMatchOnJobDate.Location = new System.Drawing.Point(36, 42);
            this.rbMatchOnJobDate.Name = "rbMatchOnJobDate";
            this.rbMatchOnJobDate.Size = new System.Drawing.Size(368, 17);
            this.rbMatchOnJobDate.TabIndex = 0;
            this.rbMatchOnJobDate.TabStop = true;
            this.rbMatchOnJobDate.Text = "Matching on both Date and Job Number (more chance of orphan entries)";
            this.rbMatchOnJobDate.UseVisualStyleBackColor = true;
            this.rbMatchOnJobDate.CheckedChanged += new System.EventHandler(this.rbMatchOnJobDate_CheckedChanged);
            // 
            // rbMathOnDateOnly
            // 
            this.rbMathOnDateOnly.AutoSize = true;
            this.rbMathOnDateOnly.Location = new System.Drawing.Point(37, 121);
            this.rbMathOnDateOnly.Name = "rbMathOnDateOnly";
            this.rbMathOnDateOnly.Size = new System.Drawing.Size(336, 17);
            this.rbMathOnDateOnly.TabIndex = 1;
            this.rbMathOnDateOnly.Text = "Matching On Date Only - ignore job no (minimise orphan plant entries)";
            this.rbMathOnDateOnly.UseVisualStyleBackColor = true;
            this.rbMathOnDateOnly.CheckedChanged += new System.EventHandler(this.rbMathOnDateOnly_CheckedChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(62, 253);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 1;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // frmPlantJobDate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(563, 301);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Name = "frmPlantJobDate";
            this.Text = "frmPlantJobDate";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rbMathOnDateOnly;
        private System.Windows.Forms.RadioButton rbMatchOnJobDate;
        private System.Windows.Forms.Button button1;
    }
}