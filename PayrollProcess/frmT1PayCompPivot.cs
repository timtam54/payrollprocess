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
    public partial class frmT1PayCompPivot : Form
    {
        public frmT1PayCompPivot(JobCostHoursPivotDS _ds)
        {
            InitializeComponent();
            JobCostHoursPivotDS.Merge(_ds);
            reportViewer1.LocalReport.EnableHyperlinks = true;
        }


        private void frmT1PayCompPivot_Load(object sender, EventArgs e)
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
            reportViewer1.Drillthrough += ReportViewer1_Drillthrough;
            this.reportViewer1.RefreshReport();
        }

        private void ReportViewer1_Drillthrough(object sender, Microsoft.Reporting.WinForms.DrillthroughEventArgs e)
        {
            Microsoft.Reporting.WinForms.LocalReport rep = (Microsoft.Reporting.WinForms.LocalReport)e.Report;
            e.Cancel = true;

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



            string T1EmpNo ="";
            decimal T1PayComp=0;
           // int RowID=-1;
            foreach (ReportParameter item in rep.OriginalParametersToDrillthrough)
            {
                //if (item.Name == "Dte")
                //    Dte=Convert.ToDateTime(item.Values[0]);
                //else
                //if (item.Name == "RowID")
                //    RowID = Convert.ToInt32(item.Values[0]);
                 if (item.Name == "T1PayComp")
                    T1PayComp = Convert.ToDecimal(item.Values[0]);
                else if (item.Name == "T1EmpNo")
                    T1EmpNo= Convert.ToString(item.Values[0]);
            }

            (new frmItemise(JobCostHoursPivotDS,  T1EmpNo, T1PayComp)).ShowDialog();
        }
    }
}
