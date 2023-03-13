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
    public partial class frmT1PayCompPivotMatrix : Form
    {
        int PayPeriod;
        public frmT1PayCompPivotMatrix(JobCostHoursPivotDS _ds, int _PayPeriod)
        {
            InitializeComponent();
            JobCostHoursPivotDS.Merge(_ds);
            PayPeriod = _PayPeriod;
            reportViewer1.LocalReport.EnableHyperlinks = true;
//            reportViewer1.Drillthrough += ReportViewer1_Drillthrough1;
        }


        private void frmT1PayCompPivotMatrix_Load(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Maximized;

            //DialogResult dr = MessageBox.Show("aggregate staff totals", "agg", MessageBoxButtons.YesNo);
            //if (dr == DialogResult.Yes)
            //    foreach (var item in JobCostHoursPivotDS.DataTable1)
            //    {

            //        item.StaffID = 0;
            //        item.EmpName = "all staff";
            //        item.T1EmpNo = "0";
            //    }
            this.reportViewer1.RefreshReport();
            this.reportViewer1.LocalReport.EnableHyperlinks = true;
            this.reportViewer1.Drillthrough += ReportViewer1_Drillthrough;
        }

        private void ReportViewer1_Drillthrough(object sender, Microsoft.Reporting.WinForms.DrillthroughEventArgs e)
        {
            decimal PayComp = 1;
            int EmpNo = 1;

            try
            {
                e.Cancel = true;

                Microsoft.Reporting.WinForms.LocalReport rep = (Microsoft.Reporting.WinForms.LocalReport)e.Report;

                if (e.ReportPath == "Filter")
                {

                    object oo = rep.OriginalParametersToDrillthrough[0].Values[0];
                    frmFilter filt = new frmFilter(JobCostHoursPivotDS.DataTable1, oo.ToString(), Cursor.Position);
                    filt.ShowDialog();
                    string val = filt.FilterVal;
                    string Filter;
                    if (val == null)
                        return;
                    if ((val != "-Remove Filter-") && (val != ""))
                    {
                        //Filter = "";

                        //Filter = Filter + oo.ToString() + " in (" + val.ToString() + ")";
                        if (DataTable1BindingSource.Filter == "" || DataTable1BindingSource.Filter == null)
                            Filter = "";
                        else
                            Filter = DataTable1BindingSource.Filter + " and ";

                        Filter = Filter + oo.ToString() + " in (" + val.ToString() + ")";

                    }
                    else
                        Filter = "";
                    DataTable1BindingSource.Filter = Filter;
                    //reportViewer1.Refresh();
                    this.reportViewer1.RefreshReport();

                    return;
                }


                foreach (ReportParameter item in rep.OriginalParametersToDrillthrough)
                {
                    if (item.Name == "EmpNo")
                        EmpNo = Convert.ToInt32(item.Values[0]);
                    else if (item.Name == "PayComp")
                        PayComp = Convert.ToDecimal(item.Values[0]);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("reportViewer1_drillthrough Error " + ex.Message);
            }
            frmItemiseTran frmTrans = new frmItemiseTran(JobCostHoursPivotDS, PayComp, EmpNo, PayPeriod,"",-1, DateTime.MinValue);
            frmTrans.ShowDialog();
        }
    }
}
