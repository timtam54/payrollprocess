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
    public partial class frmSummaryHours : Form
    {
        JobCostHoursPivotDS jobds;
        int PayPeriod;
        public frmSummaryHours(JobCostHoursPivotDS _jobds, int _PayPeriod)
        {
            InitializeComponent();
            jobds = _jobds;
            PayPeriod = _PayPeriod;
            reportViewer1.Drillthrough += ReportViewer_Drillthrough;
        }

        private void ReportViewer_Drillthrough(object sender, Microsoft.Reporting.WinForms.DrillthroughEventArgs e)
        {
            Microsoft.Reporting.WinForms.LocalReport rep = (Microsoft.Reporting.WinForms.LocalReport)e.Report;

            e.Cancel = true;
            if (e.ReportPath == "Filter")
            {
                string colname = rep.OriginalParametersToDrillthrough[0].Values[0].ToString();
                DataTable dtt = new DataTable();
                DataColumn dc = new DataColumn(colname);
                dtt.Columns.Add(dc);
                foreach (var item in sumhrs)
                {
                    DataRow dr= dtt.NewRow();
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
                    List<SumHour> filter = new List<SumHour>();
                    foreach (var item in sumhrs)
                    {
                        string itemval = item.GetType().GetProperty(colname).GetValue(item, null).ToString();
                        if (val.Contains("'"+itemval+"'"))
                            filter.Add(item);
                    }
                    //SumHourBindingSource.Filter = Filter;
                    SumHourBindingSource.DataSource = filter;
                    //SumHourBindingSource.Filter = Filter;

                    Filter = colname + "=" + val.ToString();
                }
                else
                {
                    SumHourBindingSource.DataSource = sumhrs;
                    Filter = "";

                }
                this.reportViewer1.RefreshReport();
                return;
            }
            int EmpNo=-1;
            string PayCompType = "";
            foreach (ReportParameter item in rep.OriginalParametersToDrillthrough)
            {
                if (item.Name == "EmpNo")
                    EmpNo = Convert.ToInt32(item.Values[0]);
                else if (item.Name == "PayCompType")
                    PayCompType = item.Values[0].ToString();
            }
            if (e.ReportPath == "Itemise")
            {
                (new frmItemise(jobds, EmpNo.ToString(), PayCompType)).ShowDialog();
                return;
            }

            var timesheet = db.Timesheets.Where(ts => ts.StaffID == EmpNo && ts.PayNoYear == PayPeriod).FirstOrDefault();
            System.Diagnostics.Process.Start(timesheet.filename);

        }

        public  class SumHour
        {
            public int EmpNo { get; set; }
            public string EmployeeName { get; set; }
            public decimal StandardHours { get; set; }
            public double OrdinaryHours { get; set; }
            public double LeaveHours { get; set; }
            public double OvertimeHours { get; set; }
            public double OrdPlusLeaveMinStd { get; set; }
            public string EmpType { get; set; }
        }

        DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
        List<SumHour> sumhrs = new List<SumHour>();

        private void frmSummaryHours_Load(object sender, EventArgs e)
        {
            WindowState= FormWindowState.Maximized;
            var emps = db.Employees.ToList();
            var empall = db.Emp_Allowances.Where(ea => ea.PayComponentCode == 50).ToList();

            foreach (var item in jobds.DataTable1)
            {
                SumHour sumhr = sumhrs.Where(sh => sh.EmpNo == item.StaffID).FirstOrDefault();
                if (sumhr==null)
                {
                    sumhr = new SumHour();
                    sumhr.EmpNo = item.StaffID;
                    sumhr.EmployeeName = item.EmpName;
                    sumhr.OrdinaryHours = 0;
                    sumhr.LeaveHours = 0;
                    var emp = emps.Where(em => em.T1EmpNo == item.StaffID).FirstOrDefault();
                    if (emp.Hours == null)
                        sumhr.StandardHours = 0;
                    else
                       sumhr.StandardHours = 2* Convert.ToDecimal( emp.Hours);

                    sumhr.EmpType = emp.Emp_Type.ToString();
                    var pc50 = empall.Where(ea => ea.T1_EmpID == item.StaffID).FirstOrDefault();
                    if (pc50 != null)
                        sumhr.StandardHours =Convert.ToDecimal( pc50.units);
                    sumhrs.Add(sumhr);
                }
                if (item.Leave)
                    sumhr.LeaveHours += item.LeaveHours;
                else if (item.Overtime)
                    sumhr.OvertimeHours += item.Hours;
                else

                    sumhr.OrdinaryHours += item.Hours;
            }
            foreach (SumHour sumhr in sumhrs)
            {

            sumhr.OrdPlusLeaveMinStd = sumhr.OrdinaryHours + sumhr.LeaveHours - Convert.ToDouble(sumhr.StandardHours);
            }
            SumHourBindingSource.DataSource = sumhrs;
            this.reportViewer1.RefreshReport();
           // this.reportViewer1.Height = 400;
        }

    }
}
