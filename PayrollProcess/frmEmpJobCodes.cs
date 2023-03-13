using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PayrollProcess
{
    public partial class frmEmpJobCodes : Form
    {
        string JobCodes;
        int PayNoYear;
        public frmEmpJobCodes(string _JobCodes,int _PayNoYear)
        {
            InitializeComponent();
           WindowState= FormWindowState.Maximized;
            JobCodes = _JobCodes;
            PayNoYear = _PayNoYear;
            this.Text = "JobCodes:" + JobCodes;
            reportViewer1.Drillthrough += ReportViewer1_Drillthrough;
        }

        private void ReportViewer1_Drillthrough(object sender, Microsoft.Reporting.WinForms.DrillthroughEventArgs e)
        {
            Microsoft.Reporting.WinForms.LocalReport rep = (Microsoft.Reporting.WinForms.LocalReport)e.Report;
            e.Cancel = true;
            if (e.ReportPath == "Filter")
            {
                string colname = rep.OriginalParametersToDrillthrough[0].Values[0].ToString();
                DataTable dtt = new DataTable();
                DataColumn dc = new DataColumn(colname);
                dtt.Columns.Add(dc);
                foreach (var item in emps)
                {
                    DataRow dr = dtt.NewRow();
                    var fieldValue = item.GetType().GetProperty(colname).GetValue(item, null);
                    dr[colname] = fieldValue;
                    dtt.Rows.Add(dr);
                }
                frmFilter filt = new frmFilter(dtt, colname, Cursor.Position);
                filt.ShowDialog();
                string val = filt.FilterVal;
                string Filter;
                if (val == null)
                    return;
                if ((val != "-Remove Filter-") && (val != ""))
                {
                    List<EmpNoFilename> filter = new List<EmpNoFilename>();
                    foreach (var item in emps)
                    {
                        string itemval = item.GetType().GetProperty(colname).GetValue(item, null).ToString();
                        if (val.Contains("'" + itemval + "'"))
                            filter.Add(item);
                    }
                    EmpNoFilenameBindingSource.DataSource = filter;
                    Filter = colname + "=" + val.ToString();
                }
                else
                {
                    EmpNoFilenameBindingSource.DataSource = emps;
                    Filter = "";

                }
                this.reportViewer1.RefreshReport();
                return;
            }
            string filename = "";
            foreach (ReportParameter item in rep.OriginalParametersToDrillthrough)
            {
                if (item.Name == "filename")
                    filename = item.Values[0].ToString();
            }
           
               System.Diagnostics.Process.Start(filename) ;
        }

        public class EmpNoFilename
        {
            public string filename { get; set; }
            public int EmpNo { get; set; }
            public string EmpName { get; set; }
            public decimal? TimeCode { get; set; }
            public decimal? Allowance { get; set; }
            public int? Plant { get; set; }
        }
        List<EmpNoFilename> emps;
        private void FrmEmpJobCodes_Load(object sender, EventArgs e)
        {
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
            emps = (from th in db.Timesheets join tds in db.TimesheetDatas on th.TimesheetID equals tds.TimesheetID join emp in db.Employees on th.StaffID equals emp.T1EmpNo where th.PayNoYear == PayNoYear && tds.job==JobCodes group tds by new EmpNoFilename {TimeCode=tds.TImeCode, Allowance= tds.AllowanceCode, Plant = tds.PlantNo, EmpNo = th.StaffID, filename = th.filename,EmpName=emp.FirstName +" "+ emp.Surname } into grp select new EmpNoFilename { EmpNo = grp.Key.EmpNo, filename = grp.Key.filename, EmpName = grp.Key.EmpName, TimeCode = grp.Key.TimeCode, Allowance = grp.Key.Allowance, Plant = grp.Key.Plant }  ).ToList();
            BindingSource bs = new BindingSource();
            bs.DataSource = emps;
            //dataGridView1.DataSource = bs;
            EmpNoFilenameBindingSource.DataSource = bs;
            this.reportViewer1.RefreshReport();
        }
    }
}
