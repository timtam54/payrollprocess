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
    public partial class frmLedgerJobNoMapping : Form
    {
        int PayPeriod;
        public frmLedgerJobNoMapping(int _PayPeriod)
        {
            InitializeComponent();
            WindowState= FormWindowState.Maximized;
            PayPeriod = _PayPeriod;
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
                foreach (var item in jls)
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
                    List<JobLedger> filter = new List<JobLedger>();
                    foreach (var item in jls)
                    {
                        string itemval = item.GetType().GetProperty(colname).GetValue(item, null).ToString();
                        if (val.Contains("'" + itemval + "'"))
                            filter.Add(item);
                    }
                    JobLedgerBindingSource.DataSource = filter;
                    Filter = colname + "=" + val.ToString();
                }
                else
                {
                    JobLedgerBindingSource.DataSource = jls;
                    Filter = "";

                }
                this.reportViewer1.RefreshReport();
                return;
            }

            string JobCodes = "";
            foreach (ReportParameter item in rep.OriginalParametersToDrillthrough)
            {
                if (item.Name == "JobCodes")
                    JobCodes = item.Values[0].ToString();
            }
            (new frmEmpJobCodes(JobCodes, PayPeriod)).ShowDialog();
        }

        public class JobEmp
        {
            public string JobCodes { get; set; }
            public int Emp { get; set; }
        }

        public class JobLedger
        {
            public string JobCodes { get; set; }
            public string Ledger { get; set; }
            public int Count { get; set; }
        }
        List<JobLedger> jls;
        private void FrmLedgerJobNoMapping_Load(object sender, EventArgs e)
        {
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
            List<JobEmp> jle = (from td in db.TimesheetDatas join th in db.Timesheets on td.TimesheetID equals th.TimesheetID where th.PayNoYear == PayPeriod group th by new JobEmp { JobCodes = td.job, Emp= th.StaffID } into newGroup select new JobEmp { JobCodes = newGroup.Key.JobCodes, Emp=newGroup.Key.Emp }).ToList();
            jls = (from je in jle group je by je.JobCodes into newgroup select new JobLedger{ JobCodes=newgroup.Key,Count=newgroup.Count()  }).ToList();
            foreach (JobLedger jl in jls)
                jl.Ledger= frmReportT1PayImport.GetLedger(jl.JobCodes);
            JobLedgerBindingSource.DataSource = jls;
            this.reportViewer1.RefreshReport();
        }
    }
}
