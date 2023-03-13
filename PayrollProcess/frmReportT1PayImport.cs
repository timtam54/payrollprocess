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

namespace PayrollProcess
{
    public partial class frmReportT1PayImport : Form
    {
        int PayPeriod;
        int StaffID;
        bool All;

        public frmReportT1PayImport(int _PayPeriod, int _StaffID, bool _All)
        {
            InitializeComponent();
            reportViewer1.LocalReport.EnableHyperlinks = true;
            PayPeriod = _PayPeriod;
            StaffID = _StaffID;
            All = _All;
            reportViewer1.Drillthrough += ReportViewer1_Drillthrough;
        }

        private void ReportViewer1_Drillthrough(object sender, DrillthroughEventArgs e)
        {
            decimal PayComp = -1;
            int EmpNo = 1;
            string JobNo = "";
            int RowID = -1;
            DateTime Dte = DateTime.MinValue;
            try
            {
                Microsoft.Reporting.WinForms.LocalReport rep = (Microsoft.Reporting.WinForms.LocalReport)e.Report;
                e.Cancel = true;
                if (e.ReportPath == "Filter")
                {
                    object oo = rep.OriginalParametersToDrillthrough[0].Values[0];
                    frmFilter filt = new frmFilter(mainDS1.DataTable1, oo.ToString(), Cursor.Position);
                    filt.ShowDialog();
                    string val = filt.FilterVal;
                    string Filter;
                    if (val == null)
                        return;
                    if ((val != "-Remove Filter-") && (val != ""))
                    {
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

                foreach (ReportParameter item in rep.OriginalParametersToDrillthrough)
                {
                    if (item.Name == "EmpNo")
                        EmpNo = Convert.ToInt32(item.Values[0]);
                    else if (item.Name == "PayComp")
                        PayComp = Convert.ToDecimal(item.Values[0]);
                    else if (item.Name == "JobNo")
                        JobNo = Convert.ToString(item.Values[0]);
                    else if (item.Name == "RowID")
                        RowID = Convert.ToInt32(item.Values[0]);
                    else if (item.Name == "Dte")
                        Dte = Convert.ToDateTime(item.Values[0]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("reportViewer1_drillthrough Error " + ex.Message);
            }
            if (PayComp == -1)
            {
                var timesheet = db.Timesheets.Where(ts => ts.StaffID == EmpNo && ts.PayNoYear == PayPeriod).FirstOrDefault();
                System.Diagnostics.Process.Start(timesheet.filename);

                return;
            }
            frmItemiseTran frmTrans = new frmItemiseTran(JobCostHoursPivotDS, PayComp, EmpNo, PayPeriod, JobNo, RowID, Dte);
            frmTrans.ShowDialog();
        }

        DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

        public class JobDte
        {
            public string JobCode { get; set; }
            public string JobDesc { get; set; }
            public string Time { get; set; }
            public string Allowance { get; set; }
            public DateTime Dte { get; set; }
            public DateTime StrtDte { get; set; }
            public DateTime EndDte { get; set; }
            public int StaffID { get; set; }
            public string T1Job { get; set; }
            public string ClassDesc { get; set; }
        }

        void Populate(int TimesheetID)
        {
            frmReport.FillDS(TimesheetID, JobCostHoursPivotDS, db, PayPeriod);
            FillMain();
            this.reportViewer1.RefreshReport();
        }

        void FillMain()

        {
            mainDS1.Clear();
            foreach (JobCostHoursPivotDS.DataTable1Row item in JobCostHoursPivotDS.DataTable1)
            {
                if (item.IsAllowanceDescNull())
                    item.AllowanceDesc = "";
                if (item.IsTimeNull())
                    item.Time = "";

                MainDS.DataTable1Row mrow = mainDS1.DataTable1.Where(i => i.TimeDesc == item.Time && i.EmpNo == item.StaffID && i.Code == item.Code && i.Period == item.Period && i.JobNo == item.JobNo && i.RowID == item.RowID && i.Allowance == item.Allowance).FirstOrDefault();
                if (mrow == null)
                    mainDS1.DataTable1.AddDataTable1Row(item.Period, item.StaffID, item.EmpName, item.ClassNo, item.ClassDesc, item.Code, item.CodeDesc, item.Time, item.Hours, item.JobNo, item.JobDesc, item.Allowance, item.AllowanceDesc, item.AllHours, item.RowID, item.Dte, GetLedger(item.JobNo));
                else
                {
                    mrow.Hours += item.Hours;
                    mrow.AllowUnits += item.AllHours;
                }
            }
            int cnt = mainDS1.DataTable1.Count();
        }
        private void frmReport_Load(object sender, EventArgs e)
        {
            frmT1ImpSummary.LoadAndSetCbo(comboBox1, All, PayPeriod, StaffID);
            Loaded = true;
            FillReport();
            this.WindowState = FormWindowState.Maximized;
            this.reportViewer1.RefreshReport();
            Cursor = Cursors.Default;
        }

        bool Loaded = false;
        void FillReport()
        {
            int TimesheetID = Convert.ToInt32(comboBox1.SelectedValue);
            Populate(TimesheetID);
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

        public static string GetSaveFileName()
        {
            string SaveFileName;
            string _sSuggestedName = "T1Detail.xls";
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "*Excel files (*.xls)|*.xls";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            saveFileDialog1.FileName = _sSuggestedName;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                SaveFileName = saveFileDialog1.FileName;
            else
                SaveFileName = "";
            return SaveFileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string SaveFileName;
            SaveFileName = GetSaveFileName();
            if (SaveFileName == "")
                return;
            SaveToDisk(SaveFileName, true, reportViewer1);
            System.Diagnostics.Process.Start(SaveFileName);
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

        public class DataColumnMap
        {
            public DataColumnMap(string _datacolumnname, string _output, string _NullOutputVal)
            {
                datacolumnname = _datacolumnname;
                output = _output;
                NullOutputVal = _NullOutputVal;
            }

            public DataColumnMap(string _output, string _NullOutputVal)//, string _output)
            {
                output = _output;
                NullOutputVal = _NullOutputVal;

            }
            public string NullOutputVal { get; set; }
            public string datacolumnname { get; set; }
            public string output { get; set; }

        }

        private void btnT1Import_Click(object sender, EventArgs e)
        {
            frmPlantJobDate plant_JobDate = new frmPlantJobDate();
            plant_JobDate.ShowDialog();

            DataColumnMap[] dcms = new DataColumnMap[] { new DataColumnMap("T1EmpNo", "EmployeeId","EmpNoNotMapped") , new DataColumnMap("Dte", "EntryDate","")
           , new DataColumnMap("Ledger","EntryType","Operations"), new DataColumnMap( "JobNo","Activity","")
            , new DataColumnMap("PlantAssetNumber","")//	
	
            , new DataColumnMap("AggUnits", "Units",""), new DataColumnMap( "ClockIn",""), new DataColumnMap("ClockOut",""), new DataColumnMap("Comments","")
            , new DataColumnMap("AuthorisedInd","Y")
            , new DataColumnMap("HistoricalInd",""), new DataColumnMap("PositionCode",""), new DataColumnMap("TimesheetWorkflowSystem","")
            , new DataColumnMap("TimesheetWorkflowName",""), new DataColumnMap("T1PayComp","PayComponentCode","")//"100"
            , new DataColumnMap("EmploymentConditionCode","EmploymentConditionCode",""), new DataColumnMap("EmploymentConditionGrade","EmploymentConditionGrade",""), new DataColumnMap("EmploymentConditionLevel","EmploymentConditionLevel",""), new DataColumnMap("PayrollRate","")
            ,new DataColumnMap("PlantEntry1PlantAssetNumber","NO"),new DataColumnMap("PlantEntry1Units","NO")
            ,new DataColumnMap("PlantEntry2PlantAssetNumber","NO"),new DataColumnMap("PlantEntry2Units","NO")
            ,new DataColumnMap("PlantEntry3PlantAssetNumber","NO"),new DataColumnMap("PlantEntry3Units","NO")
            ,new DataColumnMap("PlantEntry4PlantAssetNumber","NO"),new DataColumnMap("PlantEntry4Units","NO")
            ,new DataColumnMap("PlantEntry5PlantAssetNumber","NO"),new DataColumnMap("PlantEntry5Units","NO")

};

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "csv|.csv";
            sfd.ShowDialog();
            if (sfd.FileName == "")
                return;
            List<PayrollProcess.Models.T1_TS_ImportTmp> output;
            string ss;

            StringBuilder sb;

            sb = new StringBuilder();

            ss = "";
            sb.AppendLine("FORMAT TIMESHEET , STANDARD 1.0");
            foreach (DataColumnMap dcm in dcms)
            {
                if (ss != "")
                    ss += ",";
                ss += dcm.output.ToString();
            }
            sb.AppendLine(ss);

            List<frmPlant.PlantData> plantdata = GetPlant();//tds

            output = new List<Models.T1_TS_ImportTmp>();
            foreach (JobCostHoursPivotDS.DataTable1Row item in JobCostHoursPivotDS.DataTable1.OrderBy(oo => oo.StaffID))//.Where(ii=>!ii.CodeDesc.ToLower().Contains("leave") && !ii.CodeDesc.ToLower().Contains("sick")).OrderBy(oo=>oo.StaffID))
            {

                if (!item.Leave)
                {
                    item.Ledger = GetLedger(item.JobNo);
                    ss = "";
                    try
                    {
                        foreach (DataColumnMap dcm in dcms)
                        {
                            //if (dcm.datacolumnname != null)
                            //{
//                            if (dcm.datacolumnname.ToLower().Contains("pay"))
  //                              ss += "";

                            //}
                            if (dcm.datacolumnname == null)
                            {
                                if (dcm.NullOutputVal != "NO")
                                {
                                    if (ss != "")
                                        ss += ",";

                                    ss += dcm.NullOutputVal;
                                }
                            }
                            else
                            {
                                if (dcm.datacolumnname.ToLower().Contains("pay"))
                                    ss += "";

                                if (ss != "")
                                    ss += ",";

                                if (item[dcm.datacolumnname] == null)
                                    ss += dcm.NullOutputVal;
                                else
                                {
                                    string tt = item[dcm.datacolumnname].ToString();
                                    if (tt == "")
                                        ss += dcm.NullOutputVal;
                                    else
                                    {
                                        if (dcm.datacolumnname.ToLower().Contains("jobno"))
                                        {
                                            if (tt == "0")
                                                tt = "";
                                        }

                                        string tpe = item[dcm.datacolumnname].GetType().ToString();
                                        if ("System.DateTime" == tpe)
                                            ss += Convert.ToDateTime(item[dcm.datacolumnname]).ToString("dd/MM/yyyy");
                                        else if ("System.Double" == tpe)
                                            ss += (Convert.ToDouble(item[dcm.datacolumnname])).ToString("0.00");
                                        else if ("System.Decimal" == tpe)
                                            ss += (Convert.ToDecimal(item[dcm.datacolumnname])).ToString("0.00");
                                        else
                                        {
                                            ss += tt;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    try
                    {
                        if (item.Ledger.ToLower() != "normal")
                        {
                            List<frmPlant.PlantData> EmpDatePlantDatas;
                            if (plant_JobDate.MatchJobDate)
                                EmpDatePlantDatas = plantdata.Where(pd => pd.Dte == item.Dte.Date && pd.EmpNo == item.StaffID && pd.JobCode == item.JobNo).ToList();
                            else
                                EmpDatePlantDatas = plantdata.Where(pd => pd.Dte == item.Dte.Date && pd.EmpNo == item.StaffID).ToList();
                            //  List<frmPlant.PlantData> EmpDatePlantDatas = EmpDatePlantDatasNoJob;
                            for (int i = 0; i < 5; i++)
                            {

                                if (EmpDatePlantDatas.Count > i)
                                {
                                    frmPlant.PlantData EmpDatePlantData = EmpDatePlantDatas[i];
                                    if (EmpDatePlantData.PlantNo != 0)
                                   {
                                        ss += ",";
                                        ss += EmpDatePlantData.PlantNo.ToString();
                                        ss += ",";
                                        ss += EmpDatePlantData.Hours.ToString();
                                        plantdata.Remove(EmpDatePlantData);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                    sb.AppendLine(ss);
                }
            }
            if (plantdata.Count() > 0)
                (new frmPlantUnattached(plantdata)).ShowDialog();
            System.IO.File.WriteAllText(sfd.FileName.Replace(".csv", "_Work.csv"), sb.ToString());
            System.Diagnostics.Process.Start(sfd.FileName.Replace(".csv", "_Work.csv"));
        }
        List<frmPlant.PlantData> GetPlant()//JobCostHoursPivotDS tds
        {
            int TimesheetID = Convert.ToInt32(comboBox1.SelectedValue);
            List<frmPlant.PlantData> results;
            if (TimesheetID == -1)
            {
                results = (from td in db.TimesheetDatas
                           join th in db.Timesheets on td.TimesheetID equals th.TimesheetID
                           join jb in db.Jobs on td.job equals jb.JobCode into jbs
                           from jb in jbs.DefaultIfEmpty()
                           join em in db.Employees on th.StaffID equals em.T1EmpNo
                           where th.PayNoYear == PayPeriod
                           && td.PlantNo != null
                           select new frmPlant.PlantData
                           {
                               T1EmpNo = em.T1EmpNo,
                               RowID = (td.Row == null) ? 0 : Convert.ToInt32(td.Row),
                               PlantNo = Convert.ToInt32(td.PlantNo),
                               EmpName = em.FirstName + " " + em.Surname,
                               EmpNo = th.StaffID,
                               JobDesc = jb.JobDesc,
                               Dte = td.start_date.Date,
                               JobCode = td.job,
                               StrtDte = td.start_date,
                               EndDte = td.end_date,
                               Hours = 0
                           }
                   ).ToList();
            }
            else
            {
                results = (from td in db.TimesheetDatas
                           join th in db.Timesheets on td.TimesheetID equals th.TimesheetID
                           join jb in db.Jobs on td.job equals jb.JobCode into jbs
                           from jb in jbs.DefaultIfEmpty()
                           join em in db.Employees on th.StaffID equals em.T1EmpNo
                           where th.TimesheetID == TimesheetID
                           && td.PlantNo != null
                           select new frmPlant.PlantData
                           {
                               T1EmpNo = em.T1EmpNo,
                               RowID = (td.Row == null) ? 0 : Convert.ToInt32(td.Row),
                               PlantNo = Convert.ToInt32(td.PlantNo),
                               EmpName = em.FirstName + " " + em.Surname,
                               EmpNo = th.StaffID,
                               JobDesc = jb.JobDesc,
                               Dte = td.start_date.Date,
                               JobCode = td.job,
                               StrtDte = td.start_date,
                               EndDte = td.end_date,
                               Hours = 0
                           }
                   ).ToList();
            }
            foreach (var item in results)
            {
                item.Hours = (item.EndDte.Subtract(item.StrtDte).TotalMinutes / (double)60);
            }
            return results;

            ////    JobCostHoursPivotDS.DataTable1Row dr= tds.DataTable1.NewDataTable1Row();
            ////    dr.T1EmpNo = item.EmpNo.ToString();
            ////    dr.Dte = item.Dte;
            ////    dr.StaffID = item.EmpNo;
            ////    dr.Leave = false;
            ////    //dr.PlantAssetNumber=PlantAssetNumber
            ////    dr.PlantEntry1PlantAssetNumber = item.PlantNo;
            ////    dr.PlantEntry1Units = item.Hours;
            ////    tds.DataTable1.AddDataTable1Row(dr);

            ////}
            ////return tds;
        }

        static LedgerJob _ljdefault=null;
        static LedgerJob ljdefault
        {
            get
            {
                if (_ljdefault == null)
                {
                    foreach (LedgerJob lj in ledgerjob)
                    {
                        if (lj.JobNoFirstLetter == "*" && lj.Length == "*")
                        {
                            _ljdefault = lj;
                        }
                    }
                    if (_ljdefault == null)
                    {
                        _ljdefault = new LedgerJob();
                        _ljdefault.JobNoFirstLetter = "*";
                        _ljdefault.Length = "*";
                        _ljdefault.Ledger = "Not Mapped";
                    }
                }
                return _ljdefault;
            }
        }
        public static string GetLedger(string JobNo)
        {
                foreach (LedgerJob lj in ledgerjob)
            {
                if (lj.JobNoFirstLetter.ToLower() == "null")
                {
                    if (JobNo == null)
                        return lj.Ledger;
                    if (JobNo == "")
                        return lj.Ledger;
                }
                if (JobNo != null)
                {
                    if (JobNo != "")
                    {
                        if (JobNo.Substring(0, 1) == lj.JobNoFirstLetter)
                        {
                            if (lj.Length=="" || lj.Length == "*")
                                return lj.Ledger;

                            else if (JobNo.Length.ToString() == lj.Length)
                                return lj.Ledger;
                        }
                    }
                }
                

            }
            return ljdefault.Ledger;// "Not Mapped";

            //return "Operations";

            //if (JobNo == null)
            //    return "Normal";
            //if (JobNo == "")
            //    return "Normal";
            //if (JobNo.Substring(0, 1) == "1")
            //    return "Operations";
            //if (JobNo.Substring(0, 1) == "3")
            //    return "Capital";
            //if (JobNo.Substring(0, 1) == "4")
            //    return "Program and Events";
            //if (JobNo.Substring(0, 1) == "9")
            //    return "Fleet";
            //if (JobNo.Substring(0, 1) == "6")
            //    return "Water";
            //return "Normal";
        }

        public static List<LedgerJob> _ledgerjob;

        public static List<LedgerJob> ledgerjob
        {
            get
            {
                if (_ledgerjob == null)
                {
                    _ledgerjob = new List<LedgerJob>();
                    DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
                    string sss = db.Settings.Where(st => st.SettingCode == "Ledger_JobNoMap").FirstOrDefault().Vals;
                    string[] ss = sss.Split(new string[] { "~" }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var item in ss)
                    {
                        string[] ljitem = item.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                        LedgerJob lj = new LedgerJob();
                        if (ljitem.Length > 0)
                        {
                            lj.Ledger = ljitem[0];
                            if (ljitem.Length > 1)
                                lj.JobNoFirstLetter = ljitem[1];
                            if (ljitem.Length > 2)
                                lj.Length = ljitem[2];
                           _ledgerjob.Add(lj);
                        }
                    }
                }
                return _ledgerjob;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataColumnMap[] dcms = new DataColumnMap[] {new DataColumnMap( "LineType","L"),new DataColumnMap("T1EmpNo", "EmployeeId","EmpNoNotMapped")
               ,new DataColumnMap("LeaveRequestType","REQUEST"),  new DataColumnMap("Dte", "StartDate",""),  new DataColumnMap("Dte", "EndDate","")
           , new DataColumnMap("T1PayComp","PayComponentCode","100"), new DataColumnMap("PositionCode","")
                , new DataColumnMap("AggUnits", "Units","")
            , new DataColumnMap( "Reason",""), new DataColumnMap( "RequestPaymentInAdvance",""), new DataColumnMap( "Status","A")
            , new DataColumnMap( "HistoricalInd",""), new DataColumnMap( "WorkflowSystem",""), new DataColumnMap( "WorkflowName","")
                , new DataColumnMap( "OriginalPayComponentCode","")
                , new DataColumnMap( "OriginalStartDate","")
                , new DataColumnMap( "OriginalEndDate","")
                , new DataColumnMap( "OriginalPositionCode","")
                , new DataColumnMap( "EntryDate","")
                , new DataColumnMap( "EntryPayComponentCode","")
                , new DataColumnMap( "EntryUnits","")
                , new DataColumnMap( "Comments","")
                , new DataColumnMap( "PayrollRate","")
                , new DataColumnMap( "LeaveOverrideUnits","")
                , new DataColumnMap( "LeaveOverrideUnitType","")
};

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "csv|.csv";

            sfd.ShowDialog();
            if (sfd.FileName == "")
                return;
            List<PayrollProcess.Models.T1_TS_ImportTmp> output;
            string ss;

            StringBuilder sb;


            sb = new StringBuilder();
            sb.AppendLine("FORMAT LEAVE , STANDARD 1.0");
            ss = "";
            foreach (DataColumnMap dcm in dcms)
            {
                if (ss != "")
                    ss += ",";
                ss += dcm.output.ToString();
            }

            sb.AppendLine(ss);

            output = new List<Models.T1_TS_ImportTmp>();
            foreach (JobCostHoursPivotDS.DataTable1Row item in JobCostHoursPivotDS.DataTable1.OrderBy(oo => oo.StaffID))//.Where(ii => ii.CodeDesc.ToLower().Contains("leave") || ii.CodeDesc.ToLower().Contains("sick")).OrderBy(oo => oo.StaffID))
            {
                if (item.Leave)
                {
                    if ((item.AggUnits != 0) && (!item.IsT1EmpNoNull()) && (item.T1EmpNo != ""))
                    {
                        ss = "";
                        foreach (DataColumnMap dcm in dcms)
                        {
                            if (ss != "")
                                ss += ",";
                            if (dcm.datacolumnname == null)
                                ss += dcm.NullOutputVal;
                            else
                            {
                                string tt = item[dcm.datacolumnname].ToString();
                                string tpe = item[dcm.datacolumnname].GetType().ToString();
                                if ("System.DateTime" == tpe)
                                    ss += Convert.ToDateTime(tt).ToString("dd/MM/yyyy");
                                else if ("System.Decimal" == tpe)
                                    ss += Math.Abs(Convert.ToDecimal(tt)).ToString("0.00");
                                else if ("System.Double" == tpe)
                                    ss += Math.Abs(Convert.ToDecimal(tt)).ToString("0.00");
                                else
                                {
                                    if (tt.Contains("-"))
                                        tt = tt;

                                    ss += tt.ToString();
                                }
                            }
                        }                    //                PayrollProcess.Models.T1_TS_ImportTmp rec= new Models.T1_TS_ImportTmp();
                        sb.AppendLine(ss);
                    }

                }
            }
            System.IO.File.WriteAllText(sfd.FileName.Replace(".csv", "_Leave.csv"), sb.ToString());
            System.Diagnostics.Process.Start(sfd.FileName.Replace(".csv", "_Leave.csv"));

        }

        private void btnT1PayCompPivot_Click(object sender, EventArgs e)
        {
            frmT1PayCompPivot dl=new frmT1PayCompPivot(JobCostHoursPivotDS);
            dl.ShowDialog();
        }

        private void btnStaffMIssing_Click(object sender, EventArgs e)
        {
            (new frmStaffMIssing(PayPeriod)).ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            frmT1PayCompPivotMatrix dl = new frmT1PayCompPivotMatrix(JobCostHoursPivotDS, PayPeriod);
            dl.ShowDialog();

        }

        private void btnSummaryHours_Click(object sender, EventArgs e)
        {
            frmSummaryHours sh = new frmSummaryHours(JobCostHoursPivotDS,PayPeriod);
            sh.ShowDialog();
        }
    }
}
