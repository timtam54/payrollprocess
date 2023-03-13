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
    public partial class frmPlantUnattached : Form
    {
        List<frmPlant.PlantData> pd;
        public frmPlantUnattached(List<frmPlant.PlantData> _pd)
        {
            InitializeComponent();
            pd = _pd;
        }

        private void FrmPlantUnattached_Load(object sender, EventArgs e)
        {

            BindingSource bs = new BindingSource();
            bs.DataSource = pd;
            dataGridView1.DataSource = bs;// pd;
            bindingNavigator1.BindingSource = bs;
        }

        private void tsbExportExcel_Click(object sender, EventArgs e)
        {

            DataColumnMap[] dcms = new DataColumnMap[] { 
                new DataColumnMap("T1EmpNo","EmployeeId","EmpNoNotMapped")
                , new DataColumnMap("RowID", "RowID","")
            , new DataColumnMap("PlantNo","PlantAssetNumber","")
            , new DataColumnMap("EmpName","Emp Name","")
            , new DataColumnMap("EmpNo","Emp No","")
            , new DataColumnMap("JobDesc","JobDesc","")
            , new DataColumnMap("Dte", "Date","")
            , new DataColumnMap("StrtDte", "Start Date","")
            , new DataColumnMap("EndDte", "End Date","")
            , new DataColumnMap("JobCode", "Job Code","")
            , new DataColumnMap("Hours", "Hours","")
//           , new DataColumnMap("EntryType","Operations"), new DataColumnMap( "Activity","0000011")
  //          , new DataColumnMap("Hours", "Units",""), new DataColumnMap( "ClockIn",""), 
    //            new DataColumnMap("ClockOut","")
      //      , new DataColumnMap("Comments","")
        //    , new DataColumnMap("AuthorisedInd","Y")
          //  , new DataColumnMap("HistoricalInd",""), new DataColumnMap("PositionCode",""), 
            //    new DataColumnMap("TimesheetWorkflowSystem","")
    //        , new DataColumnMap("TimesheetWorkflowName",""),
      //          new DataColumnMap("PayComponentCode","")
        //   , new DataColumnMap("EmploymentConditionCode",""), 
          //      new DataColumnMap("EmploymentConditionGrade",""), 
            //    new DataColumnMap("EmploymentConditionLevel",""),
              //  new DataColumnMap("PayrollRate","")
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
            string element = "\"FORMAT TIMESHEET , STANDARD 1.0\"";
            sb.AppendLine(element);
            foreach (DataColumnMap dcm in dcms)
            {
                if (ss != "")
                    ss += ",";
                ss += dcm.output.ToString();
            }
            sb.AppendLine(ss);
            output = new List<Models.T1_TS_ImportTmp>();
            foreach (frmPlant.PlantData item in pd)//.Where(ii=>!ii.CodeDesc.ToLower().Contains("leave") && !ii.CodeDesc.ToLower().Contains("sick")).OrderBy(oo=>oo.StaffID))
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
                        object val = typeof(frmPlant.PlantData).GetProperty(dcm.datacolumnname).GetValue(item, null);
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
            System.IO.File.WriteAllText(sfd.FileName.Replace(".csv", "_PlantOrphaned.csv"), sb.ToString());
            System.Diagnostics.Process.Start(sfd.FileName.Replace(".csv", "_PlantOrphaned.csv"));

        }

    }
}
