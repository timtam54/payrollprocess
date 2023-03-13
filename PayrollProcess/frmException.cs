using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static PayrollProcess.frmReportT1PayImport;

namespace PayrollProcess
{
    public partial class frmException : Form
    {

        int PayPeriod;
        int StaffID;
        bool All;

        public frmException(int _PayPeriod, int _StaffID, bool _All)
        {
            PayPeriod = _PayPeriod;
            StaffID = _StaffID;
            All = _All;
            InitializeComponent();
            reportViewer1.Drillthrough += ReportViewer1_Drillthrough;
            reportViewer1.LocalReport.EnableHyperlinks = true;
        }
        string Filter="";

        private void ReportViewer1_Drillthrough(object sender, Microsoft.Reporting.WinForms.DrillthroughEventArgs e)
        {
            Microsoft.Reporting.WinForms.LocalReport rep = (Microsoft.Reporting.WinForms.LocalReport)e.Report;
            e.Cancel = true;
            if (e.ReportPath == "Filter")
            {
                object oo = rep.OriginalParametersToDrillthrough[0].Values[0];
                frmFilter filt = new frmFilter(errorsDS1.TSException, oo.ToString(), Cursor.Position);
                filt.ShowDialog();
                string val = filt.FilterVal;
                if (val == null)
                    return;
                if ((val != "-Remove Filter-") && (val != ""))
                {
                    if (Filter == "" || Filter == null)
                        Filter = "";
                    else
                        Filter = Filter + " and ";

                    Filter = Filter + oo.ToString() + " in (" + val.ToString() + ")";
                }
                else
                    Filter = "";
                TSExceptionBindingSource.Filter = Filter;
                this.reportViewer1.RefreshReport();
                return;
            }
            string Filename = "";
            foreach (ReportParameter item in rep.OriginalParametersToDrillthrough)
            {
                if (item.Name == "Filename")
                    Filename = Convert.ToString(item.Values[0]);
            }
            System.Diagnostics.Process.Start(Filename);
        }

        DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

        All_Error_Warning aew = All_Error_Warning.All;
        public enum All_Error_Warning
        {
            All = 0,
            Error = 1,
            Warning = 2
        }

        void Populate()
        {
            errorsDS1.Clear();
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
            int TimesheetID = Convert.ToInt32(comboBox1.SelectedValue);
            List<TSException> exds;

            if (TimesheetID == -1)
                exds = (from ex in db.TSExceptions join ts in db.Timesheets on ex.TimesheetID equals ts.TimesheetID where ts.PayNoYear == PayPeriod select ex).ToList();
            else
                exds = (from ex in db.TSExceptions where ex.TimesheetID == TimesheetID select ex).ToList();

            if (exds == null)
            {
                MessageBox.Show("no data");
                return;
            }
            foreach (var item in exds)
            {
                if (aew == All_Error_Warning.All)
                    errorsDS1.TSException.AddTSExceptionRow(item.TimesheetID, item.Field, item.Exception, item.Filename, Convert.ToBoolean(item.Error_elseWarning), item.EmpIdent, item.Tab,Convert.ToInt32(item.EmpNo));
                else if (aew == All_Error_Warning.Error && IsError(item.Error_elseWarning))
                    errorsDS1.TSException.AddTSExceptionRow(item.TimesheetID, item.Field, item.Exception, item.Filename, Convert.ToBoolean(item.Error_elseWarning), item.EmpIdent, item.Tab, Convert.ToInt32(item.EmpNo));
                else if (aew == All_Error_Warning.Warning && !IsError(item.Error_elseWarning))
                    errorsDS1.TSException.AddTSExceptionRow(item.TimesheetID, item.Field, item.Exception, item.Filename, Convert.ToBoolean(item.Error_elseWarning), item.EmpIdent, item.Tab, Convert.ToInt32(item.EmpNo));
            }
            this.reportViewer1.RefreshReport();
        }

        bool IsError(bool? Err)
        {
            if (Err == null)
                return true;
            return Convert.ToBoolean(Err);
        }

        private void bindingNavigator1_RefreshItems(object sender, EventArgs e)
        {

        }
        bool Loaded = false;
        private void frmException_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'ErrorsDS.TSException' table. You can move, or remove it, as needed.
            frmT1ImpSummary.LoadAndSetCbo(comboBox1, All, PayPeriod, StaffID);
            Loaded = true;
            Populate();
            this.WindowState = FormWindowState.Maximized;
        }

        private void toolStripDropDownButton1_Click(object sender, EventArgs e)
        {
            aew = All_Error_Warning.All;
            Populate();
        }

        private void errorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            aew = All_Error_Warning.Error;
            Populate();
        }

        private void nonErrorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            aew = All_Error_Warning.Warning;
            Populate();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!Loaded)
                return;
            Populate();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void rbError_CheckedChanged(object sender, EventArgs e)
        {
            aew = All_Error_Warning.Error;
            Populate();
        }

        private void tbAll_CheckedChanged(object sender, EventArgs e)
        {
            aew = All_Error_Warning.All;
            Populate();
        }

        private void rbWarning_CheckedChanged(object sender, EventArgs e)
        {
            aew = All_Error_Warning.Warning;
            Populate();
        }

        public static void SaveToDisk(string SaveFileName, bool open, ReportViewer reportViewer1)
        {
            try
            {
                string _sPathFilePDF = String.Empty;
                String v_mimetype;
                String v_encoding;
                String v_filename_extension;
                String[] v_streamids;
                Microsoft.Reporting.WinForms.Warning[] warnings;
                byte[] bytes = reportViewer1.LocalReport.Render("Excel", null, out v_mimetype, out v_encoding, out v_filename_extension, out v_streamids, out warnings);
                using (FileStream fs = new FileStream(SaveFileName, FileMode.Create))
                {
                    fs.Write(bytes, 0, bytes.Length);
                }
            }
            catch (Exception ex)
            {
                ;// MessageBox.Show("An Exception occured save file +" + SaveFileName + ".  Details of exception:" + ex.Message);
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "*Excel files (*.xls)|*.xls";
            sfd.ShowDialog();
            if (sfd.FileName == "")
                return;
            SaveToDisk(sfd.FileName,true,reportViewer1);
            System.Diagnostics.Process.Start(sfd.FileName);

       }
    }
}