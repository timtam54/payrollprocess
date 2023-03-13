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
    public partial class frmItemiseTran : Form
    {
        decimal PayComp = 1;
        int EmpNo = 1;
        int PayNo = 1;
        string JobNo;
        int RowID;
        JobCostHoursPivotDS jb;
        DateTime Dte;
        public frmItemiseTran(JobCostHoursPivotDS _jobds, decimal _PayComp, int _EmpNo,int _PayNo,string _JobNo,int _RowID,DateTime _Dte)
        {
            Dte = _Dte;
            InitializeComponent();
            EmpNo = _EmpNo;
            PayComp = _PayComp;
            PayNo = _PayNo;
            JobNo = _JobNo;
            RowID = _RowID;
            Load += FrmItemiseTran_Load;
            jb = _jobds;
        }
        //public class TD
        //{
        //    public decimal AllowanceCode { get; set; }
        //    public int? ClassNo { get; set; }
        //    public string Description { get; set; }
        //    public string Job { get; set; }
        //    public decimal? TimeCode { get; set; }
        //    public double hours { get; set; }
        //    public DateTime Dte { get; set; }
        //}

        private void FrmItemiseTran_Load(object sender, EventArgs e)
        {
            try
            {
                WindowState = FormWindowState.Maximized;
                if (JobNo == "")
                    JobCostHoursPivotDS.Merge(jb.DataTable1.Select("T1EmpNo = " + EmpNo.ToString() + " and T1PayComp =" + PayComp.ToString()));

                // jl = jb.DataTable1.Select(j => j.T1EmpNo == EmpNo.ToString() && j.T1PayComp == PayComp).ToList();
                else
                {
                    if (RowID==-1)
                    JobCostHoursPivotDS.Merge(jb.DataTable1.Select("T1EmpNo = " + EmpNo.ToString() + " and Code =" + PayComp.ToString() + " and JobNo='" + JobNo + "'"));
                    else
                        JobCostHoursPivotDS.Merge(jb.DataTable1.Select("T1EmpNo = " + EmpNo.ToString() + " and T1PayComp =" + PayComp.ToString() + " and RowID =" + RowID.ToString() + " and JobNo='" + JobNo + "'"));


                }
                this.reportViewer1.RefreshReport();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Eception in frmItemiseTran.Load " + ex.Message);
            }
        }

        private void frmItemiseTran_Load_1(object sender, EventArgs e)
        {

        }
    }
}
