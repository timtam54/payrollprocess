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
    public partial class frmReport : Form
    {
       int PayPeriod;
        int StaffID;
        bool All;
        public frmReport(int _PayPeriod, int _StaffID,bool _All)
        {
            InitializeComponent();
            PayPeriod = _PayPeriod;
            StaffID = _StaffID;
            All = _All;
            reportViewer1.LocalReport.EnableHyperlinks = true;
            reportViewer1.Drillthrough += ReportViewer1_Drillthrough;
        }

        private void ReportViewer1_Drillthrough(object sender, DrillthroughEventArgs e)
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
//                    Filter = "";
  //                  Filter = Filter + oo.ToString() + " in (" + val.ToString() + ")";

                    if (DataTable1BindingSource.Filter == "" || DataTable1BindingSource.Filter == null)
                        Filter = "";
                    else
                        Filter = DataTable1BindingSource.Filter + " and ";

                    Filter = Filter + oo.ToString() + " in (" + val.ToString() + ")";

                }
                else
                    Filter = "";
                DataTable1BindingSource.Filter = Filter;
                this.reportViewer1.RefreshReport();

                return;
            }


            int EmpNo =-1;
            foreach (ReportParameter item in rep.OriginalParametersToDrillthrough)
            {
                if (item.Name == "EmpNo")
                    EmpNo = Convert.ToInt32(item.Values[0]);
            }
            var timesheet = db.Timesheets.Where(ts => ts.StaffID == EmpNo && ts.PayNoYear == PayPeriod).FirstOrDefault();
            System.Diagnostics.Process.Start(timesheet.filename);
        }

        DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

        public class JobDte
        {
            public string Source { get; set; }
            public bool Leave { get; set; }
            public int T1EmpNo { get; set; }
            public decimal  T1_PayComponentID { get; set; }
            public decimal T1_AllowCodeID { get; set; }
            public int RowID { get; set; }
            public decimal ClassRate { get; set; }
            public string JobCode { get; set; }
            public string JobDesc { get; set; }
            public string Time { get; set; }
            public decimal? Allowance { get; set; }
            public string AllowanceDesc { get; set; }
            public decimal? AllowanceUnits { get; set; }
            public DateTime Dte { get; set; }
            public DateTime StrtDte { get; set; }
            public DateTime EndDte { get; set; }
            public int EmpNo { get; set; }
            public string EmpName { get; set; }
            public int Dept { get; set; }
            public string T1Job { get; set; }
            public string ClassDesc { get; set; }
            public int ClassNo { get; set; }
            public decimal Code { get; set; }
            public string CodeDesc { get; set; }
            public decimal AllowRate { get; set; }
            
        }


        void Populate(int TimesheetID)
        {
            FillDS(TimesheetID, JobCostHoursPivotDS,db,PayPeriod);
            this.reportViewer1.RefreshReport();
        }

        public static void FillDS(int TimesheetID, JobCostHoursPivotDS JobCostHoursPivotDS, DataClasses1DataContext db, int PayPeriod)
        {
            List<JobDte> results = ExtractTimesheetDataFromDBForPeriodEmployees(TimesheetID, db, PayPeriod);
            try
            {
                //clear the data set and then fill it up with time data from the database
                JobCostHoursPivotDS.Clear();
                foreach (JobDte item in results)
                {
                    try { 
                    if (item.JobDesc == null)
                        item.JobDesc = "None";
                    if (item.ClassDesc == null)
                        item.ClassDesc = "None";
                    double Hours = item.EndDte.Subtract(item.StrtDte).TotalMinutes / (double)60;
                    double AllowanceUnits = 0;
                    string T1PayCompDesc;
                    if (item.Allowance != 0)
                    {
                        T1PayCompDesc = item.AllowanceDesc;
                        AllowanceUnits = Hours;
                        Hours = 0;
                    }
                    else
                        T1PayCompDesc = item.CodeDesc;
                    double Loading = 0;
                    double LeaveHours = 0;
                    if (item.Leave)
                        LeaveHours = Hours;
                    var clss = db.Classes.Where(cl => cl.PCSClassNo == item.ClassNo).FirstOrDefault();

                    string EmploymentConditionCode = "";
                    string EmploymentConditionGrade = "";
                    string EmploymentConditionLevel = "";
                    if (clss != null)
                    {
                        if (clss.Emp_Condition != "NA")
                            EmploymentConditionCode = clss.Emp_Condition;
                        if (clss.GradeCodeID != "NA")
                            EmploymentConditionGrade = clss.GradeCodeID;
                        if (clss.LevelCodeID != "NA")
                            EmploymentConditionLevel = clss.LevelCodeID;
                    }
                    var pcs = db.PayComponents.Where(tc => tc.PayCompCode == item.T1_PayComponentID).FirstOrDefault();
                    if (pcs == null)
                        throw new Exception("paycomp not found");
                    bool Overtime = IsOvertime(pcs);
                    JobCostHoursPivotDS.DataTable1.AddDataTable1Row(item.EmpNo, item.JobCode, Hours, item.JobDesc, item.Dte, item.Time, item.ClassDesc, (item.Allowance == null) ? 0 : (Convert.ToDecimal(item.Allowance)), item.ClassRate, item.ClassNo, item.EmpName, item.Code, T1PayCompDesc, PayPeriod, item.AllowanceDesc, item.Dept, AllowanceUnits, Loading, LeaveHours, item.AllowRate, item.RowID, item.T1_AllowCodeID, item.T1_PayComponentID, item.T1EmpNo.ToString(), AllowanceUnits + Hours, item.Leave, EmploymentConditionCode, EmploymentConditionGrade, EmploymentConditionLevel, item.Source, Overtime,"");
                    if (Overtime)
                        PayCompOT(JobCostHoursPivotDS, db, PayPeriod, item, Hours, T1PayCompDesc, Loading, pcs, EmploymentConditionCode, EmploymentConditionGrade, EmploymentConditionLevel);
                    if (CodesThatContribToToolAllowBool(pcs.PayCompCode))// tcs.PayCompCode.Equals(560) || tcs.PayCompCode.Equals(100))
                        PayCompTool(JobCostHoursPivotDS, db, PayPeriod, item, Hours, pcs, Loading, EmploymentConditionCode, EmploymentConditionGrade, EmploymentConditionLevel);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error in FillDSFirst : " + ex.Message);
                    }

                }
                AddFixedAllowanceTransactions(JobCostHoursPivotDS, db, PayPeriod, results);
                foreach (var item in JobCostHoursPivotDS.DataTable1)
                {
                    try
                    { 
                    item.Ledger =frmReportT1PayImport.GetLedger(item.JobNo).Ledger;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error in FillDSLegder: " + ex.Message);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in FillDS Typed Dataset and Add Trans " + ex.Message);
            }
        }

        private static bool IsOvertime(PayComponent pcs)
        {
            bool Overtime = false;
            if (pcs.PayCompTypeDesc != null)
            {
                if (pcs.PayCompTypeDesc.ToLower().Equals("overtime"))
                    Overtime = true;
            }

            return Overtime;
        }

        private static double GetFactorFromPayComp(PayComponent tcs)
        {
            double factor;
            if (tcs.PayPeriodUnit.ToLower().Equals("percentage"))
                factor = Convert.ToDouble(tcs.Units) / (double)100;
            else if (tcs.PayPeriodUnit.ToLower().Equals("hour"))
                factor = Convert.ToDouble(tcs.Units);
            else
                throw new Exception("cant determine OT factor");
            return factor;
        }

        private static void AddFixedAllowanceTransactions(JobCostHoursPivotDS JobCostHoursPivotDS, DataClasses1DataContext db, int PayPeriod, List<JobDte> results)
        {
            try
            {
                var emps = (from res in results group res by new { EmpNo = res.EmpNo, T1EmpNo = res.T1EmpNo, EmpName = res.EmpName } into g select new { EmpNo = g.Key.EmpNo, T1EmpNo = g.Key.T1EmpNo, EmpName = g.Key.EmpName }).ToList();
                foreach (var emp in emps)
                {
                    try { 
                    double totalhours = 0;
                    string PayComponentCodes = "";
                    DateTime Dte = DateTime.Now;
                    foreach (JobDte item in results.Where(res => res.EmpNo == emp.EmpNo))
                    {
                        Dte = item.Dte;
                        if (CodesThatContribToFixedAllowBool(item.Code))
                        {
                            if (PayComponentCodes == "")
                                PayComponentCodes = item.Code.ToString();
                            else
                                PayComponentCodes = PayComponentCodes + "," + item.Code.ToString();
                            double Hours = item.EndDte.Subtract(item.StrtDte).TotalMinutes / (double)60;
                            totalhours += Hours;
                        }
                    }
                    if (totalhours > 0)
                    {
                        string EmploymentConditionCode = "";
                        string EmploymentConditionGrade = "";
                        string EmploymentConditionLevel = "";
                        PayCompFixedAllow(JobCostHoursPivotDS, db, PayPeriod, totalhours, PayComponentCodes, emp.EmpNo, "0", Dte, emp.EmpName.ToString(), 0, EmploymentConditionCode, EmploymentConditionGrade, EmploymentConditionLevel);
                    }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error in AddFixedAllowanceTransactionsLoop: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in AddFixedAllowanceTransactions: " + ex.Message);
            }
        }

        private static List<JobDte> ExtractTimesheetDataFromDBForPeriodEmployees(int TimesheetID, DataClasses1DataContext db, int PayPeriod)
        {
            List<JobDte> results = new List<JobDte>();
            try
            {

                if (TimesheetID == -1)
                {
                    results = (from td in db.TimesheetDatas
                               join th in db.Timesheets on td.TimesheetID equals th.TimesheetID
                               join jb in db.Jobs on td.job equals jb.JobCode into gj
                               from jb in gj.DefaultIfEmpty()
                               join tc in db.PayComponents on td.TImeCode equals tc.PayCompCode
                               join em in db.Employees on th.StaffID equals em.T1EmpNo
                               join cl in db.Classes on td.ClassNo equals cl.PCSClassNo
                               join all in db.PayComponents on td.AllowanceCode equals all.PayCompCode
                               where th.PayNoYear == PayPeriod
                               select new JobDte
                               {
                                   Leave = (tc.PayCompTypeDesc == null) ? false : (tc.PayCompTypeDesc.ToLower() == "Absences".ToLower()),
                                   T1EmpNo = em.T1EmpNo,//.Replace("T1_", ""),
                                   T1_AllowCodeID = td.AllowanceCode, // (all.T1_Code == null) ? -all.PCSAllowCode : Convert.ToDecimal(all.T1_Code),
                                   T1_PayComponentID = (td.TImeCode == null || td.TImeCode == 0) ? td.AllowanceCode : Convert.ToDecimal(td.TImeCode),//                               (all.PCSAllowCode != 0) ? ((all.T1_Code == null) ? -all.PCSAllowCode : Convert.ToDecimal(all.T1_Code)) : tc.PayCompCode ,
                                   RowID = (td.Row == null) ? 0 : Convert.ToInt32(td.Row),
                                   AllowRate = 0,//all.Rate,
                                   AllowanceUnits = 0,
                                   Dept = 0,
                                   Code = tc.PayCompCode,
                                   CodeDesc = tc.PayCompDesc,
                                   ClassNo = cl.PCSClassNo,
                                   Allowance = td.AllowanceCode,
                                   AllowanceDesc = all.PayCompDesc,
                                   ClassDesc = cl.PYCLASS_CLASSDESC,
                                   EmpName = em.Surname + ", " + em.FirstName,
                                   EmpNo = th.StaffID,
                                   JobDesc = jb.JobDesc,
                                   Time = tc.PayCompDesc,
                                   Dte = td.start_date.Date,
                                   JobCode = td.job,
                                   StrtDte = td.start_date,
                                   EndDte = td.end_date,
                                   Source = td.Source,
                                   ClassRate = (cl.HoursPerFN == null) ? (decimal)0 : (cl.AmountPerFN == null) ? 0 : Convert.ToDecimal(cl.AmountPerFN / cl.HoursPerFN)
                               }).ToList();
                }
                else
                {

                    results = (from td in db.TimesheetDatas
                               join th in db.Timesheets on td.TimesheetID equals th.TimesheetID
                               join jb in db.Jobs on td.job equals jb.JobCode into gj
                               from jb in gj.DefaultIfEmpty()
                               join tc in db.PayComponents on td.TImeCode equals tc.PayCompCode
                               join em in db.Employees on th.StaffID equals em.T1EmpNo
                               join cl in db.Classes on td.ClassNo equals cl.PCSClassNo
                               join all in db.PayComponents on td.AllowanceCode equals all.PayCompCode
                               where th.TimesheetID == TimesheetID
                               select new JobDte
                               {
                                   Leave = (tc.PayCompTypeDesc == null) ? false : (tc.PayCompTypeDesc.ToLower() == "Absences".ToLower()),
                                   T1EmpNo = em.T1EmpNo,//.Replace("T1_", ""),
                                   T1_AllowCodeID = td.AllowanceCode,// (td.TImeCode == 0) ? ((all.T1_Code == null) ? 0 : Convert.ToDecimal(all.T1_Code)) : 0,
                                                                     //                               T1_PayComponentID = ((tc.T1_PayComponent == null) ? -1 : Convert.ToDecimal(tc.T1_PayComponent)),
                                   T1_PayComponentID = (td.TImeCode == null || td.TImeCode == 0) ? td.AllowanceCode : Convert.ToDecimal(td.TImeCode),//(all.PCSAllowCode != 0) ? ((all.T1_Code == null) ? -all.PCSAllowCode : Convert.ToDecimal(all.T1_Code)) : tc.PayCompCode,
                                   RowID = (td.Row == null) ? 0 : Convert.ToInt32(td.Row),
                                   AllowRate = 0,//all.Rate,
                                   AllowanceUnits = 0,
                                   Dept = 0,
                                   Code = tc.PayCompCode,
                                   CodeDesc = tc.PayCompDesc,
                                   ClassNo = cl.PCSClassNo,
                                   Allowance = td.AllowanceCode,
                                   AllowanceDesc = all.PayCompDesc,
                                   ClassDesc = cl.PYCLASS_CLASSDESC,
                                   EmpName = em.Surname + ", " + em.FirstName,
                                   EmpNo = th.StaffID,
                                   JobDesc = jb.JobDesc,
                                   Time = tc.PayCompDesc,
                                   Dte = td.start_date.Date,
                                   JobCode = td.job,
                                   StrtDte = td.start_date,
                                   EndDte = td.end_date,
                                   Source = td.Source,
                                   ClassRate = (cl.HoursPerFN == null) ? (decimal)0 : (cl.AmountPerFN == null) ? 0 : Convert.ToDecimal(cl.AmountPerFN / cl.HoursPerFN)
                               }).ToList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in FIll Linq " + ex.Message);
            }

            return results;
        }

        static string _CodesThatContribToToolAllow = "empty";
        public static string CodesThatContribToToolAllow
        {
            //100,560
            get
            {
                if (_CodesThatContribToToolAllow == "empty")
                {
                    DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
                    _CodesThatContribToToolAllow = db.Settings.Where(st => st.SettingCode == "CodesThatContribToToolAllow").FirstOrDefault().Vals;
                }
                return _CodesThatContribToToolAllow;
            }
        }

        static bool CodesThatContribToToolAllowBool(decimal Paycomp)
        {
            string[] codes = CodesThatContribToToolAllow.Split(new char[] { Convert.ToChar(",") });
            foreach (var cd in codes)
            {
                if (Convert.ToDecimal(cd) == Paycomp)
                    return true;
            }
            return false;
        }


        //
        static string _CreateOTTransForCodes = "empty";
        public static decimal[] CreateOTTransForCodes
        {
            //Ordinary,Statutory Holiday,Rostered Day Off,Time in Lieu - Banked
            get
            {
                if (_CreateOTTransForCodes == "empty")
                {
                    DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
                    _CreateOTTransForCodes = db.Settings.Where(st => st.SettingCode == "CreateOTTransForCodes").FirstOrDefault().Vals;
                }
                //string[] ss = _CreateOTTransForCodes.Split(new char[] { Convert.ToChar(",") });

                return Array.ConvertAll(_CreateOTTransForCodes.Split(','), decimal.Parse);

                //return _CreateOTTransForCodes;
            }
        }


        static string _CodesThatContribToFixedAllow = "empty";
        public static string CodesThatContribToFixedAllow
        {
            //Ordinary,Statutory Holiday,Rostered Day Off,Time in Lieu - Banked
            get
            {
                if (_CodesThatContribToFixedAllow == "empty")
                {
                    DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
                    _CodesThatContribToFixedAllow = db.Settings.Where(st => st.SettingCode == "CodesThatContribToFixedAllow").FirstOrDefault().Vals;
                }
                return _CodesThatContribToFixedAllow;
            }
        }

        static bool CodesThatContribToFixedAllowBool(decimal Paycomp)
        {
            string[] codes = CodesThatContribToFixedAllow.Split(new char[] { Convert.ToChar(",") });
            foreach (var cd in codes)
            {
                if (Convert.ToDecimal(cd) == Paycomp)
                    return true;
            }
            return false;
        }



        private static void PayCompOT(JobCostHoursPivotDS JobCostHoursPivotDS, DataClasses1DataContext db, int PayPeriod, JobDte item, double Hours, string T1PayCompDesc, double Loading, PayComponent pcs, string EmploymentConditionCode, string EmploymentConditionGrade, string EmploymentConditionLevel)
        {
            //1/400.00   CWA 401.00 
            //2/401.00   Construction Work -Formsetter
            //5/403.00   On Site
            //8/404.00   In Charge Water / Weir / Sew
            //9/405.00   Rubbish Allowance
            //10/406.00  Over Award Allowance(Lozo MZ)
            //21/413.00  District Allowance -Building
            //22/414.00  District Allowance
            //31/420.00  Leading Hand
            //151/475.00 Super Relief converted to salary
            double factor = GetFactorFromPayComp(pcs);
            if (factor == 0)
                return;
            var OTAllow = (from ep in db.Emp_Allowances join all in db.PayComponents on ep.PayComponentCode equals all.PayCompCode where CreateOTTransForCodes.Contains( all.PayCompCode) && ep.T1_EmpID == item.T1EmpNo select all).ToList();
            foreach (PayComponent iii in OTAllow)
            {
                decimal T1AllowCodeOT_PayComp;
                T1AllowCodeOT_PayComp = Convert.ToDecimal(iii.PayCompCode) + (decimal)0.5;
                string T1AllowCodeOT_PayComp_Desc = iii.PayCompDesc;
                double HrsFactr=Hours * factor;
                JobCostHoursPivotDS.DataTable1.AddDataTable1Row(item.EmpNo, item.JobCode, 0, item.JobDesc, item.Dte, item.Time, item.ClassDesc, iii.PayCompCode, item.ClassRate, item.ClassNo, item.EmpName, iii.PayCompCode, T1AllowCodeOT_PayComp_Desc, PayPeriod, T1AllowCodeOT_PayComp_Desc, item.Dept, HrsFactr, Loading, 0, 0 /*iii.Rate*/, -1, T1AllowCodeOT_PayComp, T1AllowCodeOT_PayComp, item.T1EmpNo.ToString(), HrsFactr, item.Leave, EmploymentConditionCode, EmploymentConditionGrade, EmploymentConditionLevel, "Trans Created by App. 'CreateOTTransForCodes' based on record in Excel Timesheet :" + T1PayCompDesc + " ("+ item.Code.ToString() + ") for amount ("+Hours.ToString()+") where Paycomp Factor ("+factor.ToString()+ ") is non zero which is OT paycomp type. Code " + iii.PayCompCode.ToString() + " is in Settings CreateOTTrans and also in Employees ("+item.EmpName+") Entitlements and OTFactor=Y (for " + iii.PayCompDesc + "-" + iii.PayCompCode.ToString() + " so record is added for this for " + T1AllowCodeOT_PayComp.ToString() + " x Hours (" + Hours.ToString() + ") x Factor ("+factor.ToString()+") for code " + item.Code.ToString(),false,  "");
            }
        }

        /*        private static void PayExecVehAllow(JobCostHoursPivotDS JobCostHoursPivotDS, DataClasses1DataContext db, int PayPeriod, string T1_EmpNo, int EmpNo, string JobCode, DateTime Dte, string EmpName, int Dept,string EmploymentConditionCode , string EmploymentConditionGrade, string EmploymentConditionLevel)
                {
                    int T1EmpNo = Convert.ToInt32(T1_EmpNo.ToLower().Replace("t1_", ""));
                    {
                        double TotalHours = 1;
                        int AllowanceCode = 101;//Executive Vehicle Allowance
                        var EVAllow = db.PayComponents.Where(all => all.PayCompCode == AllowanceCode).FirstOrDefault();
                        string Desc = EVAllow.PayCompDesc;

                        var allow = (from ep in db.Emp_Allowances where (ep.PayComponentCode == AllowanceCode) && ep.T1_EmpID == T1EmpNo select ep).ToList().FirstOrDefault(); 
                        if (allow != null)
                            JobCostHoursPivotDS.DataTable1.AddDataTable1Row(EmpNo, JobCode, 0, Desc, Dte, Desc, "NA", AllowanceCode, 0, 0, EmpName, AllowanceCode, Desc, PayPeriod, Desc, Dept, TotalHours, 0, 0, 0, -1, allow.PayComponentCode, allow.PayComponentCode, T1EmpNo.ToString(), TotalHours, false, EmploymentConditionCode, EmploymentConditionGrade, EmploymentConditionLevel, "Employee Allowance file has Executive Vehicle Allowance (use 1 unit)");
                    }
                }
        *
        */
        //private static void PayCompLocalityAllow(JobCostHoursPivotDS JobCostHoursPivotDS, DataClasses1DataContext db, int PayPeriod, double TotalHours, string T1_EmpNo, int EmpNo, string JobCode, DateTime Dte, string EmpName, int Dept,string EmploymentConditionCode, string EmploymentConditionGrade, string EmploymentConditionLevel)
        //{
        //    int T1EmpNo = Convert.ToInt32(T1_EmpNo.ToLower().Replace("t1_", ""));
        //    {
        //        {
        //            int PCSCode = 3;//Locality - Full Rate
        //            var LocalFullAllow = db.Allowances.Where(all => all.PCSAllowCode == PCSCode).FirstOrDefault();
        //            decimal T1_PayComp = Convert.ToDecimal(LocalFullAllow.T1_Code);
        //            string Desc = LocalFullAllow.Allowance_Code_Description;

        //            var allow = (from ep in db.Emp_Allowances where (ep.PayComponentCode == T1_PayComp) && ep.T1_EmpID == T1EmpNo select ep).ToList().FirstOrDefault(); ;
        //            if (allow != null)
        //                JobCostHoursPivotDS.DataTable1.AddDataTable1Row(EmpNo, JobCode, 0, Desc, Dte, Desc, "NA"/*item.ClassDesc*/, PCSCode, 0, 0, EmpName, PCSCode, Desc, PayPeriod, Desc, Dept, TotalHours, 0, 0, 0, -1, allow.PayComponentCode, allow.PayComponentCode, T1EmpNo.ToString(), TotalHours, false, EmploymentConditionCode,EmploymentConditionGrade, EmploymentConditionLevel, "Not used - If Emp Allowance eligible for Locality - Full Rate");
        //        }
        //        {
        //            int PCSCode = 4;//Locality - Half Rate
        //            var LocalFullAllow = db.Allowances.Where(all => all.PCSAllowCode == PCSCode).FirstOrDefault();
        //            decimal T1_PayComp = Convert.ToDecimal(LocalFullAllow.T1_Code);
        //            string Desc = LocalFullAllow.Allowance_Code_Description;

        //            var allow = (from ep in db.Emp_Allowances where (ep.PayComponentCode == T1_PayComp) && ep.T1_EmpID == T1EmpNo select ep).ToList().FirstOrDefault(); ;
        //            if (allow != null)
        //                JobCostHoursPivotDS.DataTable1.AddDataTable1Row(EmpNo, JobCode, 0, Desc, Dte, Desc, "NA"/*item.ClassDesc*/, PCSCode, 0, 0, EmpName, PCSCode, Desc, PayPeriod, Desc, Dept, TotalHours*0.5, 0, 0, 0, -1, allow.PayComponentCode, allow.PayComponentCode, T1EmpNo.ToString(), TotalHours*0.5, false, EmploymentConditionCode, EmploymentConditionGrade, EmploymentConditionLevel, "Not used - If Emp Allowance eligible for Locality - Full Half Rate");
        //        }
        //        {
        //            int PCSCode = 773;//WT_Broken Shift Allowance
        //            var LocalFullAllow = db.Allowances.Where(all => all.PCSAllowCode == PCSCode).FirstOrDefault();
        //            decimal T1_PayComp = Convert.ToDecimal(LocalFullAllow.T1_Code);
        //            string Desc = LocalFullAllow.Allowance_Code_Description;

        //            var allow = (from ep in db.Emp_Allowances where (ep.PayComponentCode == T1_PayComp) && ep.T1_EmpID == T1EmpNo select ep).ToList().FirstOrDefault(); ;
        //            if (allow != null)
        //                JobCostHoursPivotDS.DataTable1.AddDataTable1Row(EmpNo, JobCode, 0, Desc, Dte, Desc, "NA"/*item.ClassDesc*/, PCSCode, 0, 0, EmpName, PCSCode, Desc, PayPeriod, Desc, Dept, TotalHours, 0, 0, 0, -1, allow.PayComponentCode, allow.PayComponentCode, T1EmpNo.ToString(), TotalHours, false, EmploymentConditionCode, EmploymentConditionGrade, EmploymentConditionLevel, "Not used - If Emp Allowance eligible for WT_Broken Shift Allowance");
        //        }
        //        {
        //            int PCSCode = 776;//WT_Over Award Allowance
        //            var LocalFullAllow = db.Allowances.Where(all => all.PCSAllowCode == PCSCode).FirstOrDefault();
        //            decimal T1_PayComp = Convert.ToDecimal(LocalFullAllow.T1_Code);
        //            string Desc = LocalFullAllow.Allowance_Code_Description;

        //            var allow = (from ep in db.Emp_Allowances where (ep.PayComponentCode == T1_PayComp) && ep.T1_EmpID == T1EmpNo select ep).ToList().FirstOrDefault(); ;
        //            if (allow != null)
        //                JobCostHoursPivotDS.DataTable1.AddDataTable1Row(EmpNo, JobCode, 0, Desc, Dte, Desc, "NA"/*item.ClassDesc*/, PCSCode, 0, 0, EmpName, PCSCode, Desc, PayPeriod, Desc, Dept, TotalHours, 0, 0, 0, -1, allow.PayComponentCode, allow.PayComponentCode, T1EmpNo.ToString(), TotalHours, false, EmploymentConditionCode, EmploymentConditionGrade, EmploymentConditionLevel, "Not used - If Emp Allowance eligible for WT_Over Award Allowance");
        //        }
        //    }
        //}

        static string _CreateFixedAllowTransForCodes = "empty";
        public static string CreateFixedAllowTransForCodes
        {
            //428,412
            get
            {
                if (_CreateFixedAllowTransForCodes == "empty")
                {
                    DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
                    _CreateFixedAllowTransForCodes = db.Settings.Where(st => st.SettingCode == "CreateFixedAllowTransForCodes").FirstOrDefault().Vals;
                }
                return _CreateFixedAllowTransForCodes;
            }
        }


        static string _CreateToolAllowTransForCodes = "empty";
        public static string CreateToolAllowTransForCodes
        {
            //421,422,423,424
            get
            {
                if (_CreateToolAllowTransForCodes == "empty")
                {
                    DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
                    _CreateToolAllowTransForCodes = db.Settings.Where(st => st.SettingCode == "CreateToolAllowTransForCodes").FirstOrDefault().Vals;
                }
                return _CreateToolAllowTransForCodes;
            }
        }

        private static void PayCompFixedAllow(JobCostHoursPivotDS JobCostHoursPivotDS, DataClasses1DataContext db, int PayPeriod, double TotalHours, string PayComponentCodes, int EmpNo, string JobCode, DateTime Dte, string EmpName, int Dept, string EmploymentConditionCode, string EmploymentConditionGrade, string EmploymentConditionLevel)
        {
            string[] PayComps = CreateFixedAllowTransForCodes.Split(new char[] { Convert.ToChar(",") });
            foreach (var PayComp in PayComps)
            {
                try
                {
                    decimal T1_PayComp = Convert.ToDecimal(PayComp);// 428;
                    var FixedAllow = db.PayComponents.Where(all => all.PayCompCode == T1_PayComp).FirstOrDefault();
                    string Desc = FixedAllow.PayCompDesc;
                    var EmpFixedAllow = (from ep in db.Emp_Allowances where (ep.PayComponentCode == T1_PayComp) && ep.T1_EmpID == EmpNo select ep).ToList().FirstOrDefault(); ;
                    if (EmpFixedAllow != null)
                        JobCostHoursPivotDS.DataTable1.AddDataTable1Row(EmpNo, JobCode, 0, Desc, Dte, Desc, "NA", T1_PayComp, 0, 0, EmpName, T1_PayComp, Desc, PayPeriod, Desc, Dept, TotalHours, 0, 0, 0, -1, EmpFixedAllow.PayComponentCode, EmpFixedAllow.PayComponentCode, EmpNo.ToString(), TotalHours, false, EmploymentConditionCode, EmploymentConditionGrade, EmploymentConditionLevel, "Employee " + EmpName + " (" + EmpNo.ToString() + ") has records in timesheet with timecode " + PayComponentCodes + " corresponding to codes in Settings 'CodesThatContribToFixedAllow' which sum to " + TotalHours.ToString() + ".  In the settings 'CreateFixedAllowanceTransForCodes' is " + EmpFixedAllow.PayComponentCode.ToString() + " which is in the Employee (" + EmpNo.ToString() + "-" + EmpName + ") entitlements Therefore the app adds a transaction (" + EmpFixedAllow.PayComponentCode.ToString() + ") for Hours: " + TotalHours.ToString() + " - " + Desc, false, "");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("PayCompFixedAllow:"+ex.Message);
                }
            }
        }

        private static void PayCompTool(JobCostHoursPivotDS JobCostHoursPivotDS, DataClasses1DataContext db, int PayPeriod, JobDte item, double Hours,PayComponent pcs, double Loading, string EmploymentConditionCode, string EmploymentConditionGrade, string EmploymentConditionLevel)
        {
            int T1EmpNo = Convert.ToInt32(item.T1EmpNo);
            string[] PayComps = CreateToolAllowTransForCodes.Split(new char[] { Convert.ToChar(",") });
            foreach (var PayComp in PayComps)
            {
                decimal T1_PayComp = Convert.ToDecimal(PayComp);// //421 || ep.PayComponentCode == 422 || ep.PayComponentCode == 423 || ep.PayComponentCode == 424
                var OTAllow = (from ep in db.Emp_Allowances where (ep.PayComponentCode == T1_PayComp) && ep.T1_EmpID == T1EmpNo select ep).ToList();
                foreach (Emp_Allowance iii in OTAllow)
                {
                    decimal T1AllowCodeOT_PayComp = iii.PayComponentCode;
                    var paycomp = db.PayComponents.Where(all => all.PayCompCode == T1AllowCodeOT_PayComp).FirstOrDefault();
                    JobCostHoursPivotDS.DataTable1.AddDataTable1Row(item.EmpNo, item.JobCode, 0, item.JobDesc, item.Dte, item.Time, item.ClassDesc, paycomp.PayCompCode, item.ClassRate, item.ClassNo, item.EmpName, paycomp.PayCompCode, paycomp.PayCompDesc, PayPeriod, paycomp.PayCompDesc, item.Dept, Hours, Loading, 0, item.AllowRate, -1, item.T1_AllowCodeID, T1AllowCodeOT_PayComp, item.T1EmpNo.ToString(), (Hours), false, EmploymentConditionCode, EmploymentConditionGrade, EmploymentConditionLevel, "In Excel timesheet it finds the row with time code "+pcs.PayCompCode.ToString() + " which is in settings 'CodesThatContribToToolAllow' with hours " + Hours.ToString() + ". The following code "+  T1AllowCodeOT_PayComp.ToString()+ " is in CreateToolAllowTransForCodes and also in this Employees "+item.EmpName+" Entitlements therefore is added by the app as a transaction for the same hours "+Hours.ToString(), false,  "");
                }
            }
        }

        bool Loaded = false;

        private void frmReport_Load(object sender, EventArgs e)
        {
            frmT1ImpSummary.LoadAndSetCbo(comboBox1, All, PayPeriod, StaffID);
            WindowState = FormWindowState.Maximized;
            Loaded = true;
            FillReport();
            Cursor = Cursors.Default;

        }

        public class TSIDStaff
        {
            public int TimesheetID { get; set; }
            public string StaffName { get; set; }
        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!Loaded)
                return;
            FillReport();
        }

        void FillReport()
        {
            int TimesheetID = Convert.ToInt32(comboBox1.SelectedValue);
            Populate(TimesheetID);
        }
    }
}
