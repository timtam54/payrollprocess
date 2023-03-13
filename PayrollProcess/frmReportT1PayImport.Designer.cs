namespace PayrollProcess
{
    partial class frmReportT1PayImport
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
            this.components = new System.ComponentModel.Container();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource1 = new Microsoft.Reporting.WinForms.ReportDataSource();
            this.DataTable1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.mainDS1 = new PayrollProcess.MainDS();
            this.JobCostHoursPivotDS = new PayrollProcess.JobCostHoursPivotDS();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnSummaryHours = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.btnStaffMIssing = new System.Windows.Forms.Button();
            this.btnT1PayCompPivot = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.btnT1Import = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.DataTable1BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.mainDS1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.JobCostHoursPivotDS)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // DataTable1BindingSource
            // 
            this.DataTable1BindingSource.DataMember = "DataTable1";
            this.DataTable1BindingSource.DataSource = this.mainDS1;
            // 
            // mainDS1
            // 
            this.mainDS1.DataSetName = "MainDS";
            this.mainDS1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // JobCostHoursPivotDS
            // 
            this.JobCostHoursPivotDS.DataSetName = "JobCostHoursPivotDS";
            this.JobCostHoursPivotDS.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // reportViewer1
            // 
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            reportDataSource1.Name = "DataSet1";
            reportDataSource1.Value = this.DataTable1BindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "PayrollProcess.T!Import.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(0, 73);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.ServerReport.BearerToken = null;
            this.reportViewer1.Size = new System.Drawing.Size(1150, 370);
            this.reportViewer1.TabIndex = 0;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnSummaryHours);
            this.panel1.Controls.Add(this.button3);
            this.panel1.Controls.Add(this.btnStaffMIssing);
            this.panel1.Controls.Add(this.btnT1PayCompPivot);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.btnT1Import);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.comboBox1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1150, 73);
            this.panel1.TabIndex = 1;
            // 
            // btnSummaryHours
            // 
            this.btnSummaryHours.Location = new System.Drawing.Point(412, 20);
            this.btnSummaryHours.Name = "btnSummaryHours";
            this.btnSummaryHours.Size = new System.Drawing.Size(97, 23);
            this.btnSummaryHours.TabIndex = 12;
            this.btnSummaryHours.Text = "Summary Hours";
            this.btnSummaryHours.UseVisualStyleBackColor = true;
            this.btnSummaryHours.Click += new System.EventHandler(this.btnSummaryHours_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(850, 20);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(97, 23);
            this.button3.TabIndex = 11;
            this.button3.Text = "Matrix Emp/PC";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // btnStaffMIssing
            // 
            this.btnStaffMIssing.Location = new System.Drawing.Point(952, 20);
            this.btnStaffMIssing.Name = "btnStaffMIssing";
            this.btnStaffMIssing.Size = new System.Drawing.Size(88, 23);
            this.btnStaffMIssing.TabIndex = 10;
            this.btnStaffMIssing.Text = "Staff MIssing";
            this.btnStaffMIssing.UseVisualStyleBackColor = true;
            this.btnStaffMIssing.Click += new System.EventHandler(this.btnStaffMIssing_Click);
            // 
            // btnT1PayCompPivot
            // 
            this.btnT1PayCompPivot.Location = new System.Drawing.Point(748, 20);
            this.btnT1PayCompPivot.Name = "btnT1PayCompPivot";
            this.btnT1PayCompPivot.Size = new System.Drawing.Size(99, 23);
            this.btnT1PayCompPivot.TabIndex = 9;
            this.btnT1PayCompPivot.Text = "T1PayCompPivot";
            this.btnT1PayCompPivot.UseVisualStyleBackColor = true;
            this.btnT1PayCompPivot.Click += new System.EventHandler(this.btnT1PayCompPivot_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(648, 20);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(96, 23);
            this.button1.TabIndex = 8;
            this.button1.Text = "T1Import - Leave";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnT1Import
            // 
            this.btnT1Import.Location = new System.Drawing.Point(512, 20);
            this.btnT1Import.Name = "btnT1Import";
            this.btnT1Import.Size = new System.Drawing.Size(130, 23);
            this.btnT1Import.TabIndex = 7;
            this.btnT1Import.Text = "T1Import - Work + Plant";
            this.btnT1Import.UseVisualStyleBackColor = true;
            this.btnT1Import.Click += new System.EventHandler(this.btnT1Import_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(358, 20);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(52, 23);
            this.button2.TabIndex = 6;
            this.button2.Text = "Excel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(97, 22);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(264, 21);
            this.comboBox1.TabIndex = 1;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(26, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Employee";
            // 
            // frmReportT1PayImport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1150, 443);
            this.Controls.Add(this.reportViewer1);
            this.Controls.Add(this.panel1);
            this.Name = "frmReportT1PayImport";
            this.Text = "frmReport";
            this.Load += new System.EventHandler(this.frmReport_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DataTable1BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.mainDS1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.JobCostHoursPivotDS)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.BindingSource DataTable1BindingSource;
        private JobCostHoursPivotDS JobCostHoursPivotDS;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button btnT1Import;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnT1PayCompPivot;
        private System.Windows.Forms.Button btnStaffMIssing;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button btnSummaryHours;
        private MainDS mainDS1;
    }
}