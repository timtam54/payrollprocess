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
using static PayrollProcess.frmReportT1PayImport;

namespace PayrollProcess
{
    public partial class frmPlant : Form
    {
        int PayPeriod;
        public frmPlant(int _PayPeriod)
        {
            InitializeComponent();
            PayPeriod = _PayPeriod;
            reportViewer1.LocalReport.EnableHyperlinks = true;
            reportViewer1.Drillthrough += ReportViewer1_Drillthrough;
            this.WindowState=           FormWindowState.Maximized;
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
                foreach (var item in results)
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
                    List<PlantData> filter = new List<PlantData>();
                    foreach (var item in results)
                    {
                        string itemval = item.GetType().GetProperty(colname).GetValue(item, null).ToString();
                        if (val.Contains("'" + itemval + "'"))
                            filter.Add(item);
                    }
                    
                    //SumHourBindingSource.Filter = Filter;
                    PlantDataBindingSource.DataSource = filter;
                    //SumHourBindingSource.Filter = Filter;

                    Filter = colname + "=" + val.ToString();
                }
                else
                {
                    PlantDataBindingSource.DataSource = results;
                    Filter = "";

                }
                this.reportViewer1.RefreshReport();
                return;
            }
            int EmpNo = -1;
            //string PayCompType = "";
            foreach (ReportParameter item in rep.OriginalParametersToDrillthrough)
            {
                if (item.Name == "EmpNo")
                    EmpNo = Convert.ToInt32(item.Values[0]);
            }

            var timesheet = db.Timesheets.Where(ts => ts.StaffID == EmpNo && ts.PayNoYear == PayPeriod).FirstOrDefault();
            System.Diagnostics.Process.Start(timesheet.filename);

        }

        DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

        public class PlantData
        {
            public int T1EmpNo { get; set; }
            public int RowID { get; set; }
            public int PlantNo { get; set; }
            public string EmpName { get; set; }
            public int EmpNo { get; set; }
            public string JobDesc { get; set; }
            public DateTime Dte { get; set; }
            public DateTime StrtDte { get; set; }
            public DateTime EndDte { get; set; }
            public string JobCode { get; set; }
            public double Hours { get; set; }
        }
        List<PlantData> results;
        private void frmPlant_Load(object sender, EventArgs e)
        {
            results = (from td in db.TimesheetDatas
                       join th in db.Timesheets on td.TimesheetID equals th.TimesheetID
                       join jb in db.Jobs on td.job equals jb.JobCode into jbs
                       from jb in jbs.DefaultIfEmpty()
                       join em in db.Employees on th.StaffID equals em.T1EmpNo
                       where th.PayNoYear == PayPeriod
                       && td.PlantNo != null
                       select new PlantData
                       {
                           T1EmpNo = em.T1EmpNo,//.Replace("T1_", ""),
                                                //         T1_PayComponentID = (all.PCSAllowCode != 0) ? ((all.T1_Code == null) ? -all.PCSAllowCode : Convert.ToDecimal(all.T1_Code)) : ((tc.T1_PayComponent == null) ? -tc.TIMECODE1 : Convert.ToDecimal(tc.T1_PayComponent)),

                           RowID = (td.Row == null) ? 0 : Convert.ToInt32(td.Row),
                           //         AllowRate = all.Rate,
                           //       AllowanceUnits = 0,
                           //Dept = em.Dept,
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
            foreach (var item in results)
            {
                item.Hours = (item.EndDte.Subtract(item.StrtDte).TotalMinutes / (double)60);
            }
            PlantDataBindingSource.DataSource = results;

            this.reportViewer1.RefreshReport();


        }

        void Excedump()
        {

            DataColumnMap[] dcms = new DataColumnMap[] { new DataColumnMap("T1EmpNo","EmployeeId","EmpNoNotMapped") 
                , new DataColumnMap("Dte", "EntryDate","")
           , new DataColumnMap("EntryType","Operations"), new DataColumnMap( "Activity","0000011")
            , new DataColumnMap("PlantNo","PlantAssetNumber","")//	
            , new DataColumnMap("Hours", "Units",""), new DataColumnMap( "ClockIn",""), new DataColumnMap("ClockOut","")
            , new DataColumnMap("Comments","")
            , new DataColumnMap("AuthorisedInd","Y")
            , new DataColumnMap("HistoricalInd",""), new DataColumnMap("PositionCode",""), new DataColumnMap("TimesheetWorkflowSystem","")
            , new DataColumnMap("TimesheetWorkflowName",""),
                new DataColumnMap("PayComponentCode","")
           , new DataColumnMap("EmploymentConditionCode",""), new DataColumnMap("EmploymentConditionGrade",""), new DataColumnMap("EmploymentConditionLevel",""), new DataColumnMap("PayrollRate","")
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
            output = new List<Models.T1_TS_ImportTmp>();
            foreach (PlantData item in results.OrderBy(oo => oo.T1EmpNo))//.Where(ii=>!ii.CodeDesc.ToLower().Contains("leave") && !ii.CodeDesc.ToLower().Contains("sick")).OrderBy(oo=>oo.StaffID))
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
                        object val = typeof(PlantData).GetProperty(dcm.datacolumnname).GetValue(item, null);
                        if (val == null)
                            ss += dcm.NullOutputVal;
                        else
                        {
                            string tt = val.ToString();
                            if (tt == "")
                                ss += dcm.NullOutputVal;
                            else
                            {
                                string tpe = val.GetType().ToString();
                                if ("System.DateTime" == tpe)
                                    ss += Convert.ToDateTime(val).ToString("dd/MM/yyyy");
                                else if ("System.Double" == tpe)
                                    ss += Math.Abs(Convert.ToDouble(val)).ToString("0.00");
                                else if ("System.Decimal" == tpe)
                                    ss += Math.Abs(Convert.ToDecimal(val)).ToString("0.00");
                                else
                                {
                                    if (tt.Contains("-"))
                                        tt = tt;
                                    ss += tt;
                                }
                            }
                        }
                    }
                }
                sb.AppendLine(ss);

            }
            System.IO.File.WriteAllText(sfd.FileName.Replace(".csv", "_Plant.csv"), sb.ToString());
            System.Diagnostics.Process.Start(sfd.FileName.Replace(".csv", "_Plant.csv"));

        }

        private void TsbPrint_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "*Excel files (*.xls)|*.xls";
            sfd.ShowDialog();
            if (sfd.FileName == "")
                return;
            frmException.SaveToDisk(sfd.FileName, true, reportViewer1);
            System.Diagnostics.Process.Start(sfd.FileName);

        }

        private void TsbClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void TsbT1PlantImport_Click(object sender, EventArgs e)
        {
            Excedump();

        }
    }
}
