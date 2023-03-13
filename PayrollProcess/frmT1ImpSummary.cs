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
    public partial class frmT1ImpSummary : Form
    {
        int PayPeriod;
        int StaffID;
        bool All;

        public frmT1ImpSummary(int _PayPeriod, int _StaffID,bool _All)
        {
            InitializeComponent();
            PayPeriod = _PayPeriod;
            StaffID = _StaffID;
            All = _All;
        }

        DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

        private void frmT1ImpSummary_Load(object sender, EventArgs e)
        {
            LoadAndSetCbo(comboBox1,All, PayPeriod, StaffID);
            Loaded = true;
            FillReport();
            this.reportViewer1.RefreshReport();
        }

        public static void LoadAndSetCbo(ComboBox comboBox1,bool All,int PayPeriod,int EmpNo)
        {
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
            frmReport.TSIDStaff allstaff = new frmReport.TSIDStaff();
            allstaff.TimesheetID = -1;
            allstaff.StaffName = "-All-";
            List<frmReport.TSIDStaff> mergedlist = new List<frmReport.TSIDStaff>();
            mergedlist.Add(allstaff);
            List<frmReport.TSIDStaff> staff = (from st in db.Employees join ts in db.Timesheets on st.T1EmpNo equals ts.StaffID where ts.PayNoYear == PayPeriod orderby st.T1EmpNo /*st.Surname + st.FirstName*/ select new frmReport.TSIDStaff { TimesheetID = ts.TimesheetID, StaffName = ((st.T1EmpNo == null) ? "" : st.T1EmpNo.ToString()) + " - " + ((st.Surname == null) ? "" : st.Surname) + ", " + ((st.FirstName == null) ? "" : st.FirstName) }).ToList();
            mergedlist.Concat(staff);

            //tsbCboEmp.Items.Add("-All-");
            //foreach (var item in db.Employees)
            //{
            //    tsbCboEmp.Items.Add(item.T1EmpNo + " - " + item.Surname + ", " + item.FirstName);

            //}


            foreach (frmReport.TSIDStaff s in staff)
            {
                mergedlist.Add(s);
            }
            comboBox1.DataSource = mergedlist;
            comboBox1.ValueMember = "TimesheetID";
            comboBox1.DisplayMember = "StaffName";
            if (!All && EmpNo!=0)
            {
                var staffsel = (from ts in db.Timesheets where ts.PayNoYear == PayPeriod && ts.StaffID == EmpNo select ts.TimesheetID).FirstOrDefault();
                comboBox1.SelectedValue = staffsel;
            }
        }

        bool Loaded = false;

        void FillReport()
        {
            int TimesheetID = Convert.ToInt32(comboBox1.SelectedValue);
            Populate(TimesheetID);

        }

        double CalcTax(double TaxableIncome)
        {
            double lreturn = 0;
            double Top;double a_Factor;double b_sub;//double Top;
            Top= 72;a_Factor = 0.19;b_sub = .19;
            if (TaxableIncome <= Top)
            {
                lreturn += TaxableIncome * a_Factor - b_sub;
                return lreturn;
            }
            lreturn += Top * a_Factor - b_sub;
            TaxableIncome -= Top;
            Top = 361-72; a_Factor = 0.2342; b_sub = 3.2130;
            if (TaxableIncome <= Top)
            {
                lreturn += TaxableIncome * a_Factor - b_sub;
                return lreturn;
            }
            lreturn += Top * a_Factor - b_sub;
            TaxableIncome -= Top;

            Top = 932-361 ; a_Factor = 0.3477; b_sub = 44.2476;
            if (TaxableIncome <= Top)
            {
                lreturn += TaxableIncome * a_Factor - b_sub;
                return lreturn;
            }
            lreturn += Top * a_Factor - b_sub;
            TaxableIncome -= Top;

            Top = 1380-932; a_Factor = 0.3450; b_sub = 41.7311;
            if (TaxableIncome <= Top)
            {
                lreturn += TaxableIncome * a_Factor - b_sub;
                return lreturn;
            }
            lreturn += Top * a_Factor - b_sub;
            TaxableIncome -= Top;

            Top = 3111- 1380 ; a_Factor = 0.3900; b_sub = 103.8657;
            if (TaxableIncome <= Top)
            {
                lreturn += TaxableIncome * a_Factor - b_sub;
                return lreturn;
            }
            lreturn += Top * a_Factor - b_sub;
            TaxableIncome -= Top;


            Top = 9999999999-3111 ; a_Factor = 0.4700; b_sub = 352.7888;
           // if (TaxableIncome <= Top)
            //{
                lreturn += TaxableIncome * a_Factor - b_sub;
                return lreturn;
            //}

//            throw new Exception("income not more than 999999999");
//            lreturn += Top * a_Factor - b_sub;
  //          TaxableIncome -= Top;

        }

        void Populate(int TimesheetID)
        {
            JobCostHoursPivotDS JobCostHoursPivotDS = new JobCostHoursPivotDS();
            frmReport.FillDS(TimesheetID, JobCostHoursPivotDS, db, PayPeriod);
            var PayPer= db.PayYears.Where(ii => ii.PayNoYear == PayPeriod).FirstOrDefault();
            double NoOfWeeks =  PayPer.EndDate.Subtract(PayPer.StartDate).TotalDays / (double)7;
            T1SumDS.Clear();
            foreach (var item in JobCostHoursPivotDS.DataTable1)
            {
                T1SumDS.DataTable1Row t1row = T1SumDS.DataTable1.Where(ii => ii.EmpNo == item.StaffID).FirstOrDefault();
                if (t1row == null)
                {
                    //decimal AutoAllow = 0;
                    //var alls = db.StaffAllowDeds.Where(i => i.EmpNo == item.StaffID && i.Item == "Allowances").Select(i => i.Val).ToList();
                    //if (alls.Count()>0)
                    //    AutoAllow= alls.Sum();
                    //decimal Deductions = 0;
                    //var deds = db.StaffAllowDeds.Where(i => i.EmpNo == item.StaffID && i.Item != "Allowances").Select(i => i.Val).ToList();
                    //if (deds.Count() > 0)
                    //    Deductions = deds.Sum();

                    //string EmpType = db.Employees.Where(i => i.EmpNo == item.StaffID).FirstOrDefault().Emp_Type.ToString();
                    //t1row = T1SumDS.DataTable1.AddDataTable1Row(item.Period, item.StaffID, item.Dept, item.EmpName, item.Hours, EmpType, item.Hours, item.Hours - item.LeaveHours, 0, item.LeaveHours, Convert.ToDouble(item.ClassRate) * item.Hours, 0,Convert.ToDecimal(NoOfWeeks), Convert.ToDecimal(NoOfWeeks/2),Convert.ToDouble( Deductions), Convert.ToDouble(item.AllowRate) * item.AllHours,Convert.ToDouble( AutoAllow));
                }
                else
                {
                    t1row.PAYHRS += item.Hours;
                    t1row.EQUIVALENT_HRS += item.Hours;
                    t1row.ORDHOURS+= (item.Hours - item.LeaveHours);
                    t1row.LVEHOURS += item.LeaveHours;
                    t1row.Gross +=  Convert.ToDouble(item.ClassRate) * item.Hours;
                    t1row.AllTotals += Convert.ToDouble(item.AllowRate) * item.AllHours;
                }
            }
            foreach (var item in T1SumDS.DataTable1)
            {
                item.Tax = CalcTax(item.Gross);
            }
            this.reportViewer1.RefreshReport();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!Loaded)
                return;
            FillReport();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string SaveFileName;
            SaveFileName = frmReportT1PayImport.GetSaveFileName();

            if (SaveFileName == "")
                return;
            frmReportT1PayImport.SaveToDisk(SaveFileName, true, reportViewer1);
            System.Diagnostics.Process.Start(SaveFileName);
        }
    }
}
