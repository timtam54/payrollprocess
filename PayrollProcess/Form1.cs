using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace PayrollProcess
{
    public partial class Form1 : Form
    {
        public Form1()
        {


            InitializeComponent();

            RunScripts();
            FormInitialise();
            tsbNormal_Event.TextChanged += TsbNormal_Event_TextChanged;
            backgroundWorker1.RunWorkerCompleted += BackgroundWorker1_RunWorkerCompleted;
            this.WindowState = FormWindowState.Maximized;
        }


        private void FormInitialise()
        {
            panel1.Visible = false;
            this.WindowState = FormWindowState.Maximized;
            lblExcelFile.MouseHover += LblExcelFile_MouseHover;
            lblExcelFile.MouseLeave += LblExcelFile_MouseLeave;
            try
            {
                openFileDialog1.InitialDirectory = Path.GetDirectoryName(ElectronicTimesheetFile);// @"\\filesrv02\electronic_timesheets$";
            }
            catch (Exception ex)
            {
                ;
            }
            dataGridView1.DataError += DataGridView1_DataError;
        }

        static void RunScripts()
        {
            ExecSQL("ALTER TABLE TSException ALTER COLUMN Exception nVARCHAR (1800)  NULL;");
        }

        public static void ExecSQL(string sql)
        {
            string constring="";
            try
            {
                constring = Form1.ConString;
                //constring = global::PayrollProcess.Properties.Settings.Default.CTRCTSPayrollDBConnectionString1;
                SqlConnection sqlcon = new SqlConnection(constring);
                sqlcon.Open();
                SqlCommand com = new SqlCommand(sql, sqlcon);
                com.ExecuteNonQuery();
                sqlcon.Close();
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Error running direct SQL scripts " + ex.Message);
                //    MessageBox.Show("Constring:"+constring);
                ex = ex;
            }
            finally
            {

            }
        }
        
        private void PromptPayPeriod()
        {
            frmPayPeriod pp = new frmPayPeriod(PayNo);
            pp.ShowDialog();
            for (int i = 0; i < cboPayPeriod.Items.Count; i++)
            {
                string ss = cboPayPeriod.Items[i].ToString();
                string[] sss = ss.Split(new string[] { "-" }, StringSplitOptions.RemoveEmptyEntries);
                string PayNoYear =Convert.ToInt32(sss[0]).ToString();
                if (PayNoYear == pp.PayPeriod.ToString())
                {
                    cboPayPeriod.SelectedIndex = i;
                    PayNo = pp.PayPeriod;
                    break;
                }
            }
        }

        //void finish(IAsyncResult ar)
        //{
        //    MessageBox.Show("Complete");
        //    this.toolStripProgressBar1.Value = this.toolStripProgressBar1.Maximum;
        //    PayPeriodExceptions();

        //}

        delegate void MigrateSQLDel();

        static string rootfolder;

        public int PayNo = -1;

        void ImportTBExcelFolder()
        {
            Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog dialog = new Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog();
            dialog.InitialDirectory = ElectronicTimesheetFolderPath;
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok)
            {
                MessageBox.Show("You selected: " + dialog.FileName);
            }
            else
                return;
            rootfolder = dialog.FileName;
            //folderBrowserDialog1.SelectedPath = ElectronicTimesheetFolderPath;
            //folderBrowserDialog1.ShowNewFolderButton = true; 
            //if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
            //{
            //    return;  //handle Cancel
            //}
            //rootfolder = folderBrowserDialog1.SelectedPath;
            
            ElectronicTimesheetFolderPath = rootfolder;
            var AllDirectoryFiles = DirSearch(rootfolder);
            MessageBox.Show("There are " + AllDirectoryFiles.Count().ToString() + " files in this folder (inc subfolders).");
            this.toolStripProgressBar1.Maximum = AllDirectoryFiles.Count();
            CleatAllTimedataForPeriod();
            dataGridView1.DataSource = null;
            backgroundWorker1.RunWorkerAsync();
    //        ExtractSubFolder(rootfolder);
      //      Finish();
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            ExtractSubFolder(rootfolder);
        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            Finish();
        }

        private void Finish()
        {
            this.toolStripStatusLabel1.Text = "Complete";
            this.toolStripProgressBar1.Value = this.toolStripProgressBar1.Maximum;
            MessageBox.Show("Complete");
            tssPerc.Text = ".";
            PayPeriodExceptions();
        }

        private List<String> DirSearch(string sDir)
        {
            List<String> files = new List<String>();
            try
            {
                int xlsxcount = 0;
                foreach (string f in Directory.GetFiles(sDir))
                {
                    if (f.ToLower().Contains(".xlsx"))
                    {
                        files.Add(f);
                        xlsxcount++;
                    }
                }
                int FolderCnt = 0;
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    files.AddRange(DirSearch(d));
                    FolderCnt++;
                }
                if (xlsxcount > 1)
                    MessageBox.Show("Multiple xlsx files in folder:" + sDir + " This occurs when either the Excel file is opened or there are duplicates of this file in this folder");
                if (xlsxcount + FolderCnt == 0)
                    MessageBox.Show("No files or subfolders in folder:" + sDir);

            }
            catch (System.Exception excpt)
            {
                MessageBox.Show("DirSearch:"+excpt.Message);
            }
            return files;
        }

        private void PayPeriodExceptions()
        {
            (new frmException(PayNo, CboEmpNo(), All)).ShowDialog();
        }

        private void ExtractSubFolder(string folder)
        {
            try
            {
                foreach (string fileloop in System.IO.Directory.GetFiles(folder))
                {
                    if (fileloop.ToLower().Contains(".xlsx"))
                    {
                        try
                        {
                            ImportXLTSHours(fileloop);
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                }
                foreach (string subfolder in System.IO.Directory.GetDirectories(folder))
                {
                    try
                    {
                        ExtractSubFolder(subfolder);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        void SetcboEmp(int EmpNo)
        {
            foreach (var item in tsbCboEmp.Items)
            {
                char cc = Convert.ToChar("-");
                string[] ss = item.ToString().Split(new char[] { cc });
                if (ss.Count() > 0)
                {
                    if (ss[0].ToString().Replace(" ", "") == EmpNo.ToString())
                    {
                        tsbCboEmp.SelectedItem = item;
                        return;
                    }
                }
            }
        }

        private void ImportSingleExcelFile()
        {
            try
            {

                openFileDialog1.Filter = "XLSX Files(*.xlsx)|*.xlsx";//|Excel Files(.xlsx)|*.xlsx|Excel Files(.xls)|*.xls| Excel Files(*.xlsm)|*.xlsm
                try
                {
                    if (System.IO.File.Exists(ElectronicTimesheetFile))
                    {
                        openFileDialog1.InitialDirectory = Path.GetDirectoryName(ElectronicTimesheetFile);
                        openFileDialog1.FileName = ElectronicTimesheetFile;// LegacyFile;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ImportSingleExcelFile1"+ex.Message);
                    EventLog("ImportSingleExcelFile1" + ex.Message, ElectronicTimesheetFile, "");

                }
                //MessageBox.Show("ImportSingleExcelFile_2");
                //MessageBox.Show(openFileDialog1.FileName);

                if (openFileDialog1.ShowDialog()== DialogResult.Cancel)
                {
                    EventLog("opening file cancelled","", "");

                    return;

                }

                try
                { 
                lblExcelFile.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    string err = "Error ImportSingleExcelFile2" + ex.Message;
                    MessageBox.Show(err);
                    EventLog(err, "", "");

                }

                try
                { 
                ElectronicTimesheetFile= openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ImportSingleExcelFile3" + ex.Message);
                    EventLog("ImportSingleExcelFile3" + ex.Message, openFileDialog1.FileName, "");

                    ;
                }

                if (lblExcelFile.Text == "")
                {
                    string message = "No file was selected.. action has been cancelled.";
                    MessageBox.Show(message);
                    EventLog(message, openFileDialog1.FileName, "");
                    return;
                }
                ImportXLTSHours(lblExcelFile.Text);
                if (ts == null)
                {
                    string message = "Timesheet could not be resolved - Employee No. ("+ EmpNo.ToString() + ") was not found or Period ("+ PayNo.ToString() + ")not found";
                    MessageBox.Show(message);
                    EventLog(message, openFileDialog1.FileName, "");
                }
                else
                {
                    EventLog("SetEmpNo Start", openFileDialog1.FileName, "");

                    SetcboEmp(EmpNo);
                    (new frmException(PayNo, EmpNo, All)).ShowDialog();
                    EventLog("Show Exception End", openFileDialog1.FileName, "");

                }
            }
            catch (Exception ex)
            {
                string message = "ImportSingleExcelFile-Error on Step 1, Import Excel.  Details: " + ex.Message;
                MessageBox.Show(message);
                EventLog(message, openFileDialog1.FileName, "");

            }
        }

        /// <summary>
        /// connect to the Excel file and read the data from specified tabs 
        /// </summary>
        /// <param name="filename"></param>
        private void ImportXLTSHours(string filename)
        {
            EventLog("Opening file", filename, "");


            XLWorkbook workbook;
            try
            {
                workbook = new XLWorkbook(filename);
            }
            catch (Exception ex)
            {
                string message = "Unable to connect to Excel file.  Please check office data driver is installed: " + ex.Message;
            //     MessageBox.Show(message);
                 Exception(-1, filename, "XLFile", message, true, filename, frmExcel.Plant_TS123.NA);
                 EventLog(message, filename, "");
                //workbook.Dispose();
                ii++;
                //ShowProgressDel del = new ShowProgressDel(ShowProgress);
                //del(ii);

                return;

                //throw ex;
            }
            //var ws1 = workbook.Worksheet(1);

            //string excelConnectionString;
            //if (System.IO.Path.GetExtension(lblExcelFile.Text).ToLower().Equals(".xls"))
             //   excelConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" + filename + @""";Extended Properties=""Excel 8.0;HDR=NO;IMEX=1;""";
            //else
            //    excelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""" + filename + @""";Extended Properties=""Excel 12.0;HDR=NO;IMEX=1;""";
            //using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
            {
                //try
                //{
                //    ;//      connection.Open();
                //}
                //catch (Exception ex)
                //{
                //    string message = "Unable to connect to Excel file.  Please check office data driver is installed: " + ex.Message;
                //    MessageBox.Show(message);
                //    Exception(-1, filename, "XLFile", message, true, filename, frmExcel.Plant_TS123.NA);
                //    EventLog(message, filename, "");
                //    return;
                //}
                TotalHours = 0;
                EventLog("Begin ExtractTS1", filename, "TS1");
                TabInfo GetTSHoursT1 = ExtractTSPlantDataFromExcel(filename, workbook, frmExcel.Plant_TS123.TS1);

                //                TabInfo GetTSHoursT1 = ExtractTSPlantDataFromExcel(filename, connection, frmExcel.Plant_TS123.TS1);
                EventLog("End ExtractTS1", filename, "TS1");

                if (GetTSHoursT1.HasData)
                {
                    EventLog("Begin ExtractTS1", filename, "TS2");
                    ExtractTSPlantDataFromExcel(filename, workbook, frmExcel.Plant_TS123.TS2);
                    EventLog("End ExtractTS1 / Begin ExtractTS2", filename, "TS23");

                    ExtractTSPlantDataFromExcel(filename, workbook, frmExcel.Plant_TS123.TS3);
                    EventLog("End ExtractTS3/Start SumDBHoursExcOT", filename, "TS3");
                    double hrsExcOT = SumDBHoursExcOT();
                    EventLog("End SumDBHoursExcOT", filename, "TS3");

                    var emp = db.Employees.Where(ep => ep.T1EmpNo == ts.StaffID).FirstOrDefault();
                    double EquivHrs = Convert.ToDouble(emp.Hours * 2);
                    //check the hours in the timesheet equals the employee standard hours - all tabs
                    EventLog("End GetEquiv Emp Hours", filename, "");

                    double hrs = SumDBHours();
                    EventLog("End SumDBHours", filename, "");
                    if (Math.Abs(hrsExcOT - EquivHrs) >= 0.1)
                    {
                        string mess;

                        mess = "EquivHrs " + EquivHrs.ToString() + " != Hours Worked:" + hrsExcOT;
                        Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Hours ", mess, false, filename, frmExcel.Plant_TS123.NA);
                        EventLog(mess, filename, "");

                    }
                    if (hrs == 0)
                    {
                        string mess = "Sheet Hours Total is 0 yet standard hours is " + EquivHrs.ToString() + ", ie empty sheet found";
                        Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Total Hours Sheet", mess, true, filename, frmExcel.Plant_TS123.NA);
                        EventLog(mess, filename, "");
                    }
                    //check the hours total at the bottom of T1 equals the sume of all hours in the timesheet - all tabs
                    else if (Math.Abs(hrs - Convert.ToDouble(TotalHours)) >= 0.1)
                    {
                        string mess = "Sheet Hours Total Label is " + Math.Round(Convert.ToDouble(TotalHours), 1).ToString() + " which is different to aggregate hours extracted " + (Math.Round(hrs, 1)).ToString();
                        Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Total Hours Sheet", mess, true, filename, frmExcel.Plant_TS123.NA);
                        EventLog(mess, filename, "");

                    }
                    string PlantSheetName = GetTab(workbook, frmExcel.Plant_TS123.Plant);
                    EventLog("Begin ExtractPlant", filename, PlantSheetName);
                    GetPlant(filename, workbook, frmExcel.Plant_TS123.Plant, PlantSheetName, false);
                    EventLog("End ExtractPlant", filename, PlantSheetName);
                }
            }
            workbook.Dispose();
            ii++;
            ShowProgressDel del = new ShowProgressDel(ShowProgress);
            del(ii);
        }

        private double SumDBHours()
        {
            //aggregate all time entries for a timesheet (employee / period)
            double hrs = 0;
            var tsd = db.TimesheetDatas.Where(td => td.TimesheetID == ts.TimesheetID && td.AllowanceCode == 0).ToList();
            foreach (var item in tsd)
            {
                if (item.TImeCode!=null)
                    hrs += item.end_date.Subtract(item.start_date).TotalMinutes;
            }
            hrs = hrs / (double)60;
            return hrs;
        }

        private double SumDBHoursExcOT()
        {
            try
            {
                //aggregate all time entries excluding overtime for a timesheet (employee / period)
                double hrs = 0;
                var tsd = db.TimesheetDatas.Where(td => td.TimesheetID == ts.TimesheetID && td.AllowanceCode == 0).ToList();
                foreach (var item in tsd)
                {
                    try
                    {
                        if (item.TImeCode != null)
                        {
                            if (db.PayComponents.Where(pc => pc.PayCompCode == item.TImeCode).FirstOrDefault().PayCompTypeDesc != "Overtime")
                                hrs += item.end_date.Subtract(item.start_date).TotalMinutes;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }

                }
                hrs = hrs / (double)60;
                return hrs;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public class TabInfo
        {
            public bool ContainsPlant = false;
            public bool HasData;

            public TabInfo(bool _HasData)
            {
                HasData = _HasData;
            }
        }

        private TabInfo ExtractTSPlantDataFromExcel(string filename, XLWorkbook /*OleDbConnection*/ connection, frmExcel.Plant_TS123 tsno)
        {
            try
            {
                EventLog("Clear excelImportDS1", filename, tsno.ToString());
                excelImportDS1.Excel.Clear();
                excelImportDS1.AcceptChanges();
                string SheetName = GetTab(connection, tsno);
                if (SheetName == null)
                {
                    EventLog("SheetName not found", filename, tsno.ToString());
                    return new TabInfo(false);
                }
                EventLog("SheetName " + SheetName.ToString(), filename, tsno.ToString());
                HeaderData hd= ExtractDataFromExcelTab(connection, SheetName,filename);
                bool StaffPeriodFound;
                if (tsno == frmExcel.Plant_TS123.TS1)
                {
                    EventLog("GetStaff_PayPeriod", filename, tsbTabNameText);
                    StaffPeriodFound = GetStaff_PayPeriod(tsbTabNameText, filename,hd);
                }
                else
                    StaffPeriodFound = true;
                if (StaffPeriodFound)
                {
                    TabInfo lret = new TabInfo(true);

                    if (!All)
                    {

                        if (tsno == frmExcel.Plant_TS123.TS1)
                        {
                            EventLog("ClearTimesheetInDB", filename, tsbTabNameText);

                            ClearTimesheetInDB(ts.TimesheetID, ts.StaffID);
                        }
                        Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Tabs", "Reading tab" + SheetName, false, filename, tsno);

                    }
                    EventLog("InitHI", filename, tsbTabNameText);

                    InitHI();
                    EventLog("ClearTopRedunantRows", filename, tsbTabNameText);

                    ClearTopRedunantRows(tsbTabNameText);
                    if (tsno != frmExcel.Plant_TS123.Plant)
                    {
                        EventLog("FindCellValues plant", filename, tsbTabNameText);

                        lret.ContainsPlant = hd.ContainsPlant;// FindCellValues("Plant", "Number");
                    }
                    EventLog("RemoveColumnsHeaderEmpty", filename, tsbTabNameText);

                    RemoveColumnsHeaderEmpty();
                    ResizeGrid();
                    EventLog("AutoMap_Import", filename, tsbTabNameText);

                    AutoMap_Import(tsbTabNameText, tsno, filename);
                    ResizeGrid2();
                    EventLog("ClearHeaderFieldMaps", filename, tsbTabNameText);

                    ClearHeaderFieldMaps();
                    if (lret.ContainsPlant)// > 0)
                    {
                        EventLog("GetPlant", filename, tsbTabNameText);

                        GetPlant(filename, connection, tsno, SheetName, lret.ContainsPlant);
                    }
                    return lret;
                }
                return new TabInfo(false);
            }
            catch (Exception ex)
            {
                EventLog("Error ExtractTSPlantDataFromExcel:"+ex.Message, filename, "");

                throw ex;

            }
        }

        int ColumnsUsed = -1;

        public class ExcelField
        {
            public ExcelField(string _SearchKey)
            {
                SearchKey = _SearchKey;
            }
            public string SearchKey { get; set; }
            public object Val { get; set; }
            public bool FoundSearchKey { get; set; }
            public bool FoundVal
            {
                get
                {
                    if (Val == null)
                        return false;
                    return true;
                }
            }

        }
        /// <summary>
        /// Read data out of Excel sheet / tab map to fields and import into database
        /// </summary>
        /// <param name="connection"></param>
        /// <param name="SheetName"></param>
        /// 
        private HeaderData ExtractDataFromExcelTab(XLWorkbook /*OleDbConnection*/ connection, string SheetName,string filename)
        {
            try
            {
                EventLog("Select * FROM [" + SheetName + "]", filename, tsbTabNameText);
                IXLWorksheet dr = connection.Worksheet(SheetName);
                IXLRange ru = dr.RangeUsed();
                var rrows = ru.RowsUsed();//.Skip(1);
                ColumnsUsed = ru.ColumnCount();//.ColumnsUsed().Count();//.ColumnCount();
                List<int> Popcells = new List<int>();
                int rowstart = 99;// 13;
                ExcelField EmpNoEF = new ExcelField("Employee Number");
                ExcelField EmpNameEF = new ExcelField("Employee Name");
                ExcelField PayNoEF = new ExcelField("Pay No");
                ExcelField PayEndEF = new ExcelField("Pay End");
                bool containsplant = false;
                // bool containsNumber = false;
                ExcelField[] EFS;
                if (All)
                    EFS = new ExcelField[] { EmpNoEF, EmpNameEF, PayNoEF };
                else
                    EFS = new ExcelField[] { EmpNoEF , EmpNameEF, PayNoEF, PayEndEF };
                foreach (var EF in EFS)
                {
                    EF.FoundSearchKey = false;
                }
                foreach (var rrow in rrows)
                {
                    try { 
                    int rn = rrow.RowNumber();
                        if (rn < rowstart)
                        {
                            foreach (var ccell in rrow.Cells())
                            {
                                try
                                {
                                    if (!containsplant)
                                    {
                                        if (ccell.CachedValue.ToString().ToLower().Contains("plant") && ccell.CachedValue.ToString().ToLower().Contains("number"))
                                            containsplant = true;
                                    }

                                    foreach (var EF in EFS)
                                    {
                                        if (!EF.FoundVal)
                                        {
                                            try
                                            {
                                                if (EF.FoundSearchKey)
                                                {
                                                    if (!ccell.CachedValue.Equals(""))
                                                    {
                                                        EF.FoundSearchKey = false;
                                                        EF.Val = ccell.CachedValue.ToString();
                                                        EF.Val = EF.Val.ToString().ToLower().Replace("t1_", "");
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {

                                            }
                                            try
                                            {
                                                if (ccell.CachedValue.ToString().ToLower().Contains(EF.SearchKey.ToLower()))
                                                    EF.FoundSearchKey = true;
                                            }
                                            catch (Exception ex)
                                            {

                                            }
                                        }
                                    }
                                    try
                                    {
                                        if (rrow.RowNumber() > 4)
                                        {
                                            if (/*ccell.CachedValue.ToString().ToLower().Contains("Plant".ToLower()) ||todotim */ ccell.CachedValue.ToString().ToLower().Contains("Allow".ToLower()) || ccell.CachedValue.ToString().ToLower().Contains("Time Code".ToLower()) || ccell.CachedValue.ToString().ToLower().Contains("Leave".ToLower()))
                                            {
                                                rowstart = rn;
                                                break;
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }
                                catch (Exception ex)
                                {

                                }
                            }
                        }
//                    else
                    if (rrow.RowNumber() >= rowstart)
                    {
                        EventLog("read row" + rrow.RowNumber().ToString(), filename, tsbTabNameText);
                        bool Populated = false;
                        ExcelImportDS.ExcelRow exrow = excelImportDS1.Excel.NewExcelRow();
                        int iordhours = -10;
                        int i = 0;
                        int popi = 0;
                        //bool LastCellEmpNo = false;
                        foreach (var ccell in rrow.Cells())
                        {
                            if (!containsplant)
                            {
                                if (ccell.CachedValue.ToString().ToLower().Contains("plant") && ccell.CachedValue.ToString().ToLower().Contains("number"))
                                    containsplant = true;
                            }
                            if (rrow.RowNumber() == rowstart || Popcells.Contains(i))
                            {
                                object oo = "";
                                try
                                {
                                    
                                    oo = ccell.CachedValue;
                                }
                                catch (Exception rex)
                                {
                                    try
                                    {
                                        oo = ccell.Value;
                                    }
                                    catch (Exception ex)
                                    {
                                        object ll = ex;
                                    }
                                }

                                try
                                {
                                    if (oo.ToString() != "")
                                    {
                                        Populated = true;
                                        if (rrow.RowNumber() == rowstart)
                                            Popcells.Add(i);
                                    }
                                    try
                                    {
                                        //pingpong    

                                        if (rrow.RowNumber() == rowstart)
                                        {
                                            if (oo.ToString() != "")
                                            {
                                                //if (oo.Equals("Total Hours"))
                                                //  oo = "Total Hours";
                                                exrow[popi] = oo;
                                                if (oo.Equals("Ord          Hours"))
                                                    iordhours = i;
                                                else if (iordhours + 1 == i)
                                                    exrow[popi] = "Overtime";
                                                else if (iordhours + 2 == i)
                                                    exrow[popi] = "LeaveHours";
                                            }
                                        }
                                        else
                                        {
                                            exrow[popi] = oo;
                                            if (oo.Equals("Ord          Hours"))
                                                iordhours = i;
                                            else if (iordhours + 1 == i)
                                                exrow[popi] = "Overtime";
                                            else if (iordhours + 2 == i)
                                                exrow[popi] = "LeaveHours";

                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        throw ex;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    throw ex;
                                }
                                if (rrow.RowNumber() == rowstart)
                                {
                                    if (oo.ToString() != "")
                                        popi++;
                                }
                                else
                                    popi++;
                            }
                            i++;
                        }
                        if (Populated)
                            excelImportDS1.Excel.AddExcelRow(exrow);
                    }
                    }
                    catch (Exception ex)
                    {

                    }
                }
                ShowExcelDel del = new ShowExcelDel(ShowExcel);
                
                del((ExcelImportDS)excelImportDS1.Copy());

                HeaderData hd = new HeaderData();
                hd.EmpName = EmpNameEF.Val.ToString();

                hd.EmpNo =0;
                try
                {
                    hd.EmpNo=Convert.ToInt32(EmpNoEF.Val);
                }
                catch (Exception ex)
                {

                }
                hd.PayNo = PayNoEF.Val.ToString();
                hd.ContainsPlant = containsplant ;
                if (!All)
                {
                    DateTime pedt = new DateTime(1899, 12, 31).AddDays(Convert.ToInt32(PayEndEF.Val));
                    if (pedt.Month < 7)
                        hd.FinYear = pedt.Year ;
                    else
                        hd.FinYear = pedt.Year + 1;

                }
                return hd;
                //Popcells = Popcells;
            }
            catch (Exception ex)
            {
                EventLog("Extract data from Excel failed:" + ex.Message, filename, tsbTabNameText);

                throw ex;
            }

        }
        public class HeaderData
        {
            public int? FinYear { get; set; }
            public int EmpNo { get; set; }
            public string EmpName { get; set; }
            public string PayNo { get; set; }

            public bool ContainsPlant { get; set; }

        }
        /// <summary>
        /// reset field mapping to Excel Header 
        /// </summary>
        void ClearHeaderFieldMaps()
        {
            //clears all mapping elements / Header Item 
            foreach (HeaderIndex item in HeaderColumnMaps)
            {
                item.SelectedColumn = -1;
                item.dayindex = -1;

            }
        }
        private void GetPlant(string filename, XLWorkbook /*OleDbConnection*/ connection,frmExcel.Plant_TS123 Plant_TS, string SheetName, bool T1containsPlant)
        {
            try
            {
                excelImportDS1.Excel.Clear();
                excelImportDS1.AcceptChanges();

                List<int> Popcells = new List<int>();
                int rowstart = 99;// 13;

                //                if ((T1containsPlant != -1) || ((T1containsPlant == -1) && SheetName.ToLower().Contains("plant")))
                if (T1containsPlant || SheetName.ToLower().Contains("plant"))
                {
                    IXLWorksheet dr = connection.Worksheet(SheetName);

                    IXLRange ru = dr.RangeUsed();
                    var rrows = ru.RowsUsed();//.Skip(1);
                    ColumnsUsed = ru.ColumnCount();//.ColumnsUsed().Count();//.ColumnCount();


                    int containsplant = -1;
                    foreach (var rrow in rrows)
                    {

                        int rn = rrow.RowNumber();
                        if (rn < rowstart)
                        {
                            int cc = 0;
                            foreach (var ccell in rrow.Cells())
                            {

                                if (containsplant == -1)
                                {
                                    if (ccell.CachedValue.ToString().ToLower().Contains("plant") && ccell.CachedValue.ToString().ToLower().Contains("number"))
                                        containsplant = cc;
                                }

                                if (rrow.RowNumber() > 4)
                                {
                                    if (/*ccell.CachedValue.ToString().ToLower().Contains("Plant".ToLower()) ||todotim */ ccell.CachedValue.ToString().ToLower().Contains("Allow".ToLower()) || ccell.CachedValue.ToString().ToLower().Contains("Time Code".ToLower()) || ccell.CachedValue.ToString().ToLower().Contains("Leave".ToLower()))
                                    {
                                        rowstart = rn;
                                        break;
                                    }
                                }
                                cc++;
                            }
                        }
                        if (rrow.RowNumber() >= rowstart)
                        {
                            bool Populated = false;
                            ExcelImportDS.ExcelRow exrow = excelImportDS1.Excel.NewExcelRow();
                            int i = 0;
                            int popi = 0;

                            foreach (var ccell in rrow.Cells())
                            {
                                if (containsplant == -1)
                                {
                                    if (ccell.CachedValue.ToString().ToLower().Contains("plant") && ccell.CachedValue.ToString().ToLower().Contains("number"))
                                        containsplant = i;
                                }

                                if (rrow.RowNumber() == rowstart || Popcells.Contains(i))
                                {
                                    object oo = "";
                                    try
                                    {


                                        oo = ccell.CachedValue;
                                    }
                                    catch (Exception rex)
                                    {
                                        try
                                        {
                                            oo = ccell.Value;
                                        }
                                        catch (Exception ex)
                                        {
                                            object ll = ex;
                                        }
                                    }
                                    if (oo.ToString() != "")
                                    {
                                        Populated = true;
                                        if (rrow.RowNumber() == rowstart)
                                        {
                                            //int num = 0;
                                            if (Weekdays.Contains(oo.ToString()) && containsplant == -1)

                                                //                                           if (int.TryParse(oo.ToString(), out num) && i < containsplant)
                                                ;//ignore this is day for hours worked
                                            else
                                            {
                                                Popcells.Add(i);
                                                exrow[popi] = oo;
                                                popi++;
                                            }
                                        }
                                    }
                                    try
                                    {
                                        if (rrow.RowNumber() != rowstart)
                                            //{
                                            //    if (oo.ToString() != "")
                                            //    {
                                            //        exrow[popi] = oo;
                                            //    }
                                            //}
                                            //else
                                            exrow[popi] = oo;
                                    }
                                    catch (Exception ex)
                                    {
                                        throw ex;
                                    }
                                    if (rrow.RowNumber() != rowstart)
                                        popi++;
                                }
                                i++;

                            }
                            if (Populated)
                                excelImportDS1.Excel.AddExcelRow(exrow);
                        }
                    }
                    //         return;
                    Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Tabs", "Reading tab" + SheetName, false, filename, Plant_TS);

                    if (excelImportDS1.Excel.Count > 0)
                    {
                        ShowExcelDel del = new ShowExcelDel(ShowExcel);
                        del((ExcelImportDS)excelImportDS1.Copy());

                        InitHIPlant();
                        ClearTopRedunantRows(tsbTabNameText);
                        //todotim RemoveColumnsHeaderEmpty();
                        ResizeGrid();
                        if (T1containsPlant || SheetName.ToLower().Contains("plant"))

//                            if (T1containsPlant)// != -1)
                        {
                           // MessageBox.Show("Plant data loaded");
                            ;//todotim ResetHeaderMapping();
                             //todotim MapHeader();
                             //todotim RemoveColumns(HIJobCodes.SelectedColumn + 1, HIPlantNo.SelectedColumn - HIJobCodes.SelectedColumn - 2);

                            AutoMap_Import(tsbTabNameText, frmExcel.Plant_TS123.Plant, filename);
                            ResizeGrid2();
                            ClearHeaderFieldMaps();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("GetPlant:" + ex.Message);
            }
        }

        List<string> _Weekdays;


        List<string> Weekdays
        {
            get
            {
                if (_Weekdays==null)
                {
                    _Weekdays = new List<string>();
                    _Weekdays.Add("M");
                    _Weekdays.Add("T");
                    _Weekdays.Add("W");
                    _Weekdays.Add("TH");
                    _Weekdays.Add("F");
                    _Weekdays.Add("S");
                    _Weekdays.Add("Class");
                    _Weekdays.Add("Time Code");
                }
                return _Weekdays;
            }
        }

        private string GetTab(XLWorkbook/*OleDbConnection*/ connection, frmExcel.Plant_TS123 Plant_ElseTSHours)
        {
            IXLWorksheets dt = connection.Worksheets;//.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            frmExcel excSheet = new frmExcel(dt, Plant_ElseTSHours);
            excSheet.ShowDialog();
            SetTitleTabDel del = new SetTitleTabDel(SetTitleTab);
            tsbTabNameText = excSheet.SheetName;
            del(null, tsbTabNameText );

//            tsbTabName.Text = excSheet.SheetName;
            if (Plant_ElseTSHours == frmExcel.Plant_TS123.TS2)
            {
                if (!tsbTabNameText.Contains("2"))
                    return null;
            }
            if (Plant_ElseTSHours == frmExcel.Plant_TS123.TS3)
            {
                if (!tsbTabNameText.Contains("3"))
                    return null;
            }
            return excSheet.SheetName;
        }



        static int ii = 0;
        void ShowProgress(decimal Perc)
        {
            if (!All)
                return;
            try
            {
                Invoke(new MethodInvoker(() =>
                {
                    decimal pc = ((decimal)Perc) / ((decimal)toolStripProgressBar1.Maximum);

                    if (Convert.ToInt32(Perc) <= this.toolStripProgressBar1.Maximum)
                    {
                        this.toolStripProgressBar1.Value = Convert.ToInt32(Perc);
                        tssPerc.Text = Perc.ToString() + " out of " + toolStripProgressBar1.Maximum.ToString() + "  " + pc.ToString("P") ;
                    }
                }));

            }
            catch (Exception ex)
            {
                ;
            }
        }

        void ShowExcel(ExcelImportDS excel)
        {
            if (!All)
                return;
            try
            {
                Invoke(new MethodInvoker(() =>
                {
                    bindingSource1.DataSource = excel;
                    dataGridView1.DataSource = bindingSource1;
                    this.bindingNavigator1.BindingSource = bindingSource1;
//                    dataGridView1.DataSource=excel.Copy
                    //if (Convert.ToInt32(Perc) <= this.toolStripProgressBar1.Maximum)
                    //{
                    //    this.toolStripProgressBar1.Value = Convert.ToInt32(Perc);
                    //    tssPerc.Text = Perc.ToString() + " out of " + toolStripProgressBar1.Maximum.ToString();
                    //}
                }));

            }
            catch (Exception ex)
            {
                ;
            }
        }

        delegate void ShowExcelDel(ExcelImportDS excelds);

        delegate void ShowProgressDel(decimal Perc);

        static string _ConString;
        public static string ConString
        {
            get
            {
                //var xx = System.Configuration.ConfigurationManager.ConnectionStrings["CTRCTSPayrollDBConnectionString1"].ConnectionString;

                if (_ConString == null)
                    _ConString = Properties.Settings.Default.CTRCTSPayrollDBConnectionString1;
                //{
                //    Assembly service = Assembly.GetExecutingAssembly();
                //    ConnectionStringsSection css = ConfigurationManager.OpenExeConfiguration(service.Location).ConnectionStrings;
                //    string cs = css.ConnectionStrings["PayrollProcess.Properties.Settings.CTRCTSPayrollDBConnectionString1"].ConnectionString;
                //}
                //    _ConString = "Data Source=.;Initial Catalog=CTRCTSPayrollDB;Integrated Security=true;";

              //  _ConString = Properties.Settings.Default.CTRCTSPayrollDBConnectionString1;
                return _ConString;
            }
        }

        static DataClasses1DataContext db = new DataClasses1DataContext(ConString);

        int EmpNo = 0;

        bool FindCellValue(string Match)
        {
            for (int j = 0; j < 20; j++)
            {
                try
                {
                    string oo = excelImportDS1.Excel[0][j].ToString().ToLower();
                    if (oo.Contains(Match.ToLower()))
                    {
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            return false;
        }

        int FindCellValues(string Match, string Match2)
        {
            int Max = Math.Min(excelImportDS1.Excel.Count, 45);
            for (int i = 0; i < Max; i++)
            {
                for (int j = 0; j < ColumnsUsed; j++)
                {
                    string oo = excelImportDS1.Excel[i][j].ToString().ToLower();
                    if (oo.Contains(Match.ToLower()))
                    {
                        if (oo.Contains(Match2.ToLower()))
                        {
                            return j;
                        }
                    }
                }
            }
            return 0;
        }

        public class cell
        {
            public object val { get; set; }
            public int row { get; set; }
            public int col { get; set; }

        }
        cell FindCellValueBoth(string Match)
        {
            cell retn = new cell();
            int Max = Math.Min(excelImportDS1.Excel.Count, 45);
            for (int i = 0; i < Max; i++)
            {
                for (int j = 0; j < 30; j++)
                {
                    string oo = excelImportDS1.Excel[i][j].ToString().ToLower();
                    if (oo.Contains(Match.ToLower()))
                    {
                        int k = 1;
                        while (true)
                        {
                            object ret = excelImportDS1.Excel[i][j + k];
                            if (!ret.Equals(""))
                            {
                                retn.val = ret;
                                retn.col = j + k;
                                retn.row = i;
                                return retn;
                            }
                            k++;
                        }
                    }
                }
            }
            return null;
        }

        bool GetStaff_PayPeriod(string tabname, string filename,HeaderData hd)
        {
            EventLog("GetStaff_PayPeriod", filename, tabname);
            EmpNo = 0;// object FinYear;
            object EmpNoObj;// object PayNoObj;
            if (tabname.ToLower().Contains("_t1"))
            {
                //EmpNoObj = FindCellValueBoth("Employee Number").val;
                //EmpNoObj = EmpNoObj.ToString().ToLower().Replace("t1_", "");
                EmpNoObj = hd.EmpNo;
                if (EmpNoObj.Equals("") || EmpNoObj.Equals("0"))
                {
                    string EmpNam = hd.EmpName;// FindCellValueBoth("Employee Name").val.ToString();
                    var emp2 = db.Employees.Where(emp => emp.Surname + ", " + emp.FirstName == EmpNam).FirstOrDefault();
                    if (emp2 != null)
                        EmpNoObj = emp2.T1EmpNo;
                    ;
                }
                EventLog("Find PayNo", filename, tabname);
                //cell PayNoObjcell =FindCellValueBoth("Pay No");

                //PayNoObj = hd.PayNo;// FindCellValueBoth("Pay No").val;
             //todotim   if (!All)
                //    FinYear = excelImportDS1.Excel[PayNoObjcell.row][PayNoObjcell.col + 1];

               // else
                    //FinYear = -1;
                    EventLog("PayNoObj:" + hd.PayNo.ToString(), filename, tabname);
                //should paynumber be 0 if 0 as per message
            }
            else
            {
                Exception(-1, filename, "TABS", "Timesheets is wrong format - missing expected tabs / data", true, filename, frmExcel.Plant_TS123.TS1);
                return false;
            }
            if (!hd.PayNo.All(c => Char.IsDigit(c)))
            {
                Exception(-1, filename, "HeaderPayNo", hd.PayNo + " is not a PayNo - Timesheet no processed", true, filename, frmExcel.Plant_TS123.TS1);
                return false;
            }
            else
            {

                EventLog("SetPeriod:", filename, tabname);

                if (SetPeriod)
                {
                    for (int i = cboPayPeriod.Items.Count; i >= 1; i--)
                    //for (int i = 1; i <= cboPayPeriod.Items.Count; i++)
                    {

                        string ss = cboPayPeriod.Items[cboPayPeriod.Items.Count - i].ToString();
                        string[] sss = ss.Split(new string[] { "-" }, StringSplitOptions.RemoveEmptyEntries);
                        string PayNoYear = Convert.ToInt32(sss[0]).ToString();

                        if (PayNoYear.Length >= 2)
                        {
//                            string PayNo = ;
                            if (PayNoYear.Substring(PayNoYear.Length - 2, 2) == Convert.ToInt32(hd.PayNo).ToString("00"))
                            {
                                cboPayPeriod.SelectedIndex = cboPayPeriod.Items.Count - i;
                                //check yea match
                                if (!All)
                                {
                                    if (PayNoYear.Length >= 4)
                                    {
                                        int cboYear = Convert.ToInt32(PayNoYear.Substring(0, 4));
                                        //string[] yy = hd.FinYear.ToString().Split(new char[] { char.Parse("/") });
                                        //if (Convert.ToInt32(yy[1].Replace("-", "")) != cboYear)
                                        if (hd.FinYear!=cboYear)
                                        {
                                            MessageBox.Show("error matched to wrong year");
                                        }
                                        //FinYear
//                                        if (!All)
                                            PayNo =hd.FinYear.Value *100 + Convert.ToInt32(hd.PayNo);

                                    }
                                }
                                MessageBox.Show("This single timesheet will be imported into payrun period " + cboPayPeriod.SelectedItem.ToString());

                                break;
                            }
                        }
                    }
                }
//                else
  //                   PayNo = Convert.ToInt32(PayNoObj);

                if (EmpNoObj == null)
                {
                    EventLog(EmpNoObj + " is null", filename, tabname);

                    Exception(-1, filename, "HeaderEmpNo", EmpNoObj + " is null", true, filename,  frmExcel.Plant_TS123.TS1);
                    return false;
                }
                else if (!EmpNoObj.ToString().All(c => Char.IsDigit(c)))
                {
                    Exception(-1, filename, "HeaderEmpNo", EmpNoObj + " is not a EmpNo - contain letters", true, filename, frmExcel.Plant_TS123.TS1);
                    return false;
                }
                else if (EmpNoObj.Equals(""))
                {
                    Exception(-1, filename, "HeaderEmpNo", "Blank is not a EmpNo - contain letters", true, filename, frmExcel.Plant_TS123.TS1);
                    return false;
                }
                else
                {
                    EmpNo = Convert.ToInt32(EmpNoObj);
                    EventLog("EmpNo:" + EmpNo.ToString(), filename, tabname);
                    bool empnofound = true;

                    if (tabname.ToLower().Contains("_t1"))
                    {
                        staff = db.Employees.Where(i => i.T1EmpNo == EmpNo).FirstOrDefault();
                        if (staff == null)
                        {
                            Exception(-1, filename, "EmpNo", "T1 Emp No not found " + EmpNo.ToString(), true, filename, frmExcel.Plant_TS123.TS1);
                            EventLog("EmpNo not found", filename, tabname);

                            return false;
                        }
                        else
                            EmpNo = staff.T1EmpNo;
                    }
                    else
                        staff = db.Employees.Where(i => i.T1EmpNo == EmpNo).FirstOrDefault();
                    if (staff == null)
                    {
                        EventLog("Add Employee if not found:" + EmpNo.ToString(), filename, tabname);

                        Employee ee = new Employee();
                        ee.T1EmpNo = EmpNo;
                        string empname = filename.Substring(filename.LastIndexOf('\\') + 1);
                        empname = empname.Replace(".xlsx", "");
                        ee.FirstName = empname;
                        ee.Surname = empname;
                        db.Employees.InsertOnSubmit(ee);
                        db.SubmitChanges();
                        empnofound = false;
                        staff = db.Employees.Where(i => i.T1EmpNo == EmpNo).FirstOrDefault();
                    }
                    EventLog("find ts no for :" + PayNo.ToString() + " and " + EmpNo.ToString(), filename, tabname);
                    ts = db.Timesheets.Where(i => i.PayNoYear == PayNo && i.StaffID == EmpNo).FirstOrDefault();
                    if (ts == null)
                    {
                        ts = new Timesheet();
                        ts.PayNoYear = PayNo;
                        ts.StaffID = EmpNo;
                        db.Timesheets.InsertOnSubmit(ts);
                        db.SubmitChanges();
                        ts = db.Timesheets.Where(i => i.PayNoYear == PayNo && i.StaffID == EmpNo).FirstOrDefault();
                        if (!empnofound)
                        {
                            Exception(ts.TimesheetID, filename, "HeaderEmpNo", EmpNoObj + " not found", true, filename, frmExcel.Plant_TS123.TS1);
                            return false;
                        }
                    }

                    //staff.Payrun = PayNo;
                    ts.filename = filename;
                    db.SubmitChanges();

                    SetTitleTabDel del = new SetTitleTabDel(SetTitleTab);
                    del(staff.Surname + " - " + ts.PayNoYear, ".");
                    //tsTitle.Text = staff.Surname + " - " + ts.PayNoYear;
                    //MessageBox.Show("1. Timesheet is " + staff.GivenName + " " + staff.Surname + " Pay Period " + payyear.PayNoYear.ToString() + " from " + payyear.StartDate.ToString("dd-MMM-yyyy") + " to " + payyear.EndDate.ToString("dd-MMM-yyyy"));
                }
            }
            return true;
        }

        delegate void SetTitleTabDel(string Title, string Tab);

        string tsbTabNameText = ".";
        void SetTitleTab(string Title, string Tab)
        {

            try
            {
                Invoke(new MethodInvoker(() =>
                {
                    if (Title != null)
                        tsTitle.Text = Title;
                    if (Tab != null)
                        tsbTabName.Text = Tab;
                }));

            }
            catch (Exception ex)
            {
                ;
            }
        }

        Guid EventGroup;
        private void EventLog(string Error,  string filename, string TabSheet)
        {
            if (All)
                return;//todotim
            if (tsbNormal_Event.SelectedIndex==1)
            {
                try
                {
                    EventLog el = new EventLog();
                    el.Error = Error;
                    el.Filename = filename;
                    el.EventDT = DateTime.Now;
                    el.TabSheet = TabSheet;
                    el.EventGroup = EventGroup;
                    db.EventLogs.InsertOnSubmit(el);
                    db.SubmitChanges();
                }
                catch (Exception ex)
                {
                    ;
                }
            }
        }


        private void Exception(int timesheetid, string EmpIdent, string field, string error, bool Error_elseWarning, string filename, frmExcel.Plant_TS123 tsno)
        {
            try
            {
                if (timesheetid == -1)
                {
                    TSException x = new TSException();
                    Timesheet tt;
                    var staffid = db.Timesheets.Where(ts => ts.PayNoYear == PayNo).Select(ts => ts.StaffID).Min();//.FirstOrDefault();
                    staffid -= 1;
                    {
                        tt = new Timesheet();
                        tt.StaffID = staffid;
                        tt.PayNoYear = PayNo;
                        db.Timesheets.InsertOnSubmit(tt);
                        db.SubmitChanges();
                        tt = db.Timesheets.Where(ts => ts.StaffID == staffid && ts.PayNoYear == PayNo).FirstOrDefault();
                    }
                    timesheetid = tt.TimesheetID;
                }
                //     else
                {
                    var excep = db.TSExceptions.Where(tse => tse.TimesheetID == timesheetid && tse.Field == field).FirstOrDefault();
                    if (excep != null)
                    {
                        if (!excep.Exception.ToLower().Contains(error.ToLower()))
                        {
                            string result = excep.Exception + ", " + error;
                            if (result.Length < 1800)
                            {
                                excep.Exception = result;
                                db.SubmitChanges();
                            }
                        }
                    }
                    else
                    {
                        TSException x = new TSException();
                        x.TimesheetID = timesheetid;
                        x.Field = field;
                        x.Exception = error;
                        x.EmpNo = EmpNo;
                        if (tsno == null)
                            x.Tab = tsbTabNameText.Replace("'", "");
                        else
                            x.Tab = tsbTabNameText.Replace("'", "") + "/" + tsno.ToString();
                        if (EmpIdent.Contains("\\") || EmpIdent.Contains(".xlsx"))
                        {
                            EmpIdent = EmpIdent.Substring(EmpIdent.LastIndexOf('\\') + 1);
                            EmpIdent = EmpIdent.Replace(".xlsx", "");
                        }
                        x.EmpIdent = EmpIdent;

                        x.Filename = filename;
                        x.Error_elseWarning = Error_elseWarning;
                        db.TSExceptions.InsertOnSubmit(x);
                        db.SubmitChanges();
                        //                exds.TSException.AddTSExceptionRow(ts.TimesheetID, field, error, filename);
                    }
                }
            }
            catch (Exception ex)
            {
                ex = ex;
            }
        }

        Employee staff;
        Timesheet ts;

        void RemoveColumnsHeaderEmpty()
        {
            int removed = 0;
            int i = 0;
            while ((removed + i) < ColumnsUsed)
            {
                DataColumn dc = excelImportDS1.Excel.Columns[i];
                if (excelImportDS1.Excel.Rows[0][dc].Equals(""))
                {
                    RemoveColumn(i,(ColumnsUsed-removed));
                    removed++;
                    //                        i--;
                }
                else
                    i++;

        //        i++;
            }
        }

        void RemoveColumns(int From, int Count)
        {
            for (int i = 0; i < Count; i++)
            {
                DataColumn dc = excelImportDS1.Excel.Columns[From];
                {
                    string Header = excelImportDS1.Excel[0][dc.ColumnName].ToString();
                    //if (excelImportDS1.Excel.Rows[0][dc].Equals(""))
                    {
                        RemoveColumn(From,ColumnsUsed);
                        //  i--;
                    }
                }
            }
        }





        void RemoveColumn(int Tag,int max)
        {
            foreach (var row in excelImportDS1.Excel)
            {
                for (int i = Tag; i < max-1/*excelImportDS1.Excel.Columns.Count - 1*/; i++)
                {
                    row[i] = row[i + 1];
                }
            }
        }

        private bool ColumnAllBlank(DataColumn dc)
        {
            for (int rowindex = 0; rowindex < excelImportDS1.Excel.Rows.Count; rowindex++)
            {
                if (!excelImportDS1.Excel.Rows[rowindex][dc].Equals(""))
                    return false;
            }
            return true;
        }

        public class HeaderIndex
        {
            public int Index;
            public int SelectedColumn = -1;
            public string Header; public string Header2 = "";public string Header3="";
            public int dayindex = -1;
            public HeaderIndex(string _Header)
            {
                Header = _Header;
            }
            public HeaderIndex(string _Header, string _Header2)
            {
                Header = _Header;
                Header2 = _Header2;
            }

            public HeaderIndex(string _Header, string _Header2, string _Header3)
            {
                Header = _Header;
                Header2 = _Header2;
                Header3 = _Header3;
            }


            public HeaderIndex(string _Header, int _dayindex)
            {
                Header = _Header;
                dayindex = _dayindex;
            }
        }
        public class MapCtrl
        {
            public ComboBox ctrl;
            public int ctrlSelectedIndex;
            public List<string> ctrlItems;
            //public string ctrlSelectedItem;
            public int ColIndex;
            public MapCtrl(ComboBox _ctrl, int _ColIndex)
            {
                ctrl = _ctrl;
                ColIndex = _ColIndex;
            }
        }

        HeaderIndex[] HeaderColumnMaps;
        MapCtrl[] mpcs;

        /// <summary>
        /// keep deleting rows at top until header row 
        /// </summary>
        /// <param name="tabname"></param>
        void ClearTopRedunantRows(string tabname)
        {
 //           int cnt;
            if ((tabname.ToLower().Contains("_t2")) || (tabname.ToLower().Contains("_t3")))
                tabname = tabname;
            if ((tabname.ToLower().Contains("_t1")) || (tabname.ToLower().Contains("plant")) || (tabname.ToLower().Contains("_t2")) || (tabname.ToLower().Contains("_t3")))
            {
                while (true)
                {
                    try
                    {
                        if (excelImportDS1.Excel.Count == 0)//added by thams 28/05/21
                            return;
                        if (FindCellValue("Time Code"))
                            return;
                        if (FindCellValue("Leave"))
                            return;
                        if (FindCellValue("Allow"))
                            return;
                        if (FindCellValue("Plant"))
                            return;
                        ExcelImportDS.ExcelRow delexrow = excelImportDS1.Excel[0];
                        delexrow.Delete();
                        excelImportDS1.AcceptChanges();

                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }
            tabname = tabname;
        }

        /// <summary>
        /// Map column Headers and then import table data into the timesheet database
        /// </summary>
        /// <param name="tabname"></param>
        /// <param name="Plant_ElseTSHours"></param>
        /// <param name="filename"></param>
        /// <returns></returns>
        bool AutoMap_Import(string tabname,frmExcel.Plant_TS123 Plant_ElseTSHours,string filename)
        {
            try
            {
                ResetHeaderMapping();
                MapHeader();
                RemoveHeaders();
                if (Plant_ElseTSHours == frmExcel.Plant_TS123.TS1)
                    TotalHours = FindCellValueBoth("Total Hours").val;
                RemoveFooter();

                if (Validation(tabname, Plant_ElseTSHours, filename))
                {
               
                    ImportMigrate(Plant_ElseTSHours, Convert.ToDouble(TotalHours), filename);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                
                MessageBox.Show("AutoMap_Import filename " +filename+". "+ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Map timesheet data fields to Excel TS column headers
        /// </summary>
        private void MapHeader()
        {
            List<int> selected = new List<int>();
            if (!All)
                panel1.Visible = true;
            ExcelImportDS.ExcelRow exrow = excelImportDS1.Excel[0];
            ExcelImportDS.ExcelRow exrow2 = excelImportDS1.Excel[1];

            for (int i = 0; i < ColumnsUsed /*excelImportDS1.Excel.Columns.Count*/; i++)
            {
                try
                {
                    string colname = exrow[i].ToString();
                    string colname2 = exrow2[i].ToString();
                    colname = colname.Replace("\n", " ");
                    int selectedindex = -1;

                    foreach (HeaderIndex item in HeaderColumnMaps)
                    {

                        string Match;
                        if (item.dayindex == -1)
                            Match = colname.ToLower();
                        else
                            Match = colname2.ToLower();
                        //string itemHeader = item.Header;
                        if (Match == "code")
                            Match = "time code";
                        Match = Match.Replace("desscription", "description");
                        while (Match.Contains("  "))
                            Match = Match.Replace("  ", " ");
                        if (Match == item.Header.ToLower() || (item.Header2 != "" && Match == item.Header2.ToLower()))
                        {
                            object oo;
                            if (item == HIOrdHours)
                                oo = HIOrdHours;
                            try
                            {
                                if (item.SelectedColumn == -1)
                                {
                                    item.SelectedColumn = i;
                                    selectedindex = item.Index;
                                }
                                break;
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                    }
                    if (selectedindex > -1)
                    {
                        if (!selected.Contains(selectedindex))
                        {
                            selected.Add(selectedindex);
                            foreach (var mpcsitem in mpcs)
                            {
                                if (i == mpcsitem.ColIndex)
                                {
                                    try
                                    {
                                        if (All)
                                            mpcsitem.ctrlSelectedIndex = selectedindex;
                                        else
                                            mpcsitem.ctrl.SelectedIndex = selectedindex;
                                        break;
                                    }
                                    catch (Exception ex)
                                    {
                                        throw ex;
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            for (int i = 0; i < ColumnsUsed /*excelImportDS1.Excel.Columns.Count*/; i++)
            {
                try
                {

                    string colname = exrow[i].ToString();
                    string colname2 = exrow2[i].ToString();
                    colname = colname.Replace("\n", " ");
                    int selectedindex = -1;
                    foreach (HeaderIndex item in HeaderColumnMaps)
                    {
                        string g;
                        string Match;
                        if (item.dayindex == -1)
                            Match = colname.ToLower();
                        else
                            Match = colname2.ToLower();
                        //string itemHeader = item.Header;
                        Match = Match.Replace("desscription", "description");
                        if (Match.Contains(item.Header.ToLower()) || ((item.Header2 != "") && (Match.Contains(item.Header2.ToLower()))))
                        {
                            object oo;
                            if (item == HIOrdHours)
                                oo = HIOrdHours;

                            try
                            {
                                if (item.SelectedColumn == -1)
                                {
                                    if (item == HITimeCode)
                                        g = "we";
                                    item.SelectedColumn = i;
                                    selectedindex = item.Index;
                                }
                                break;
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                    }
                    if (selectedindex > -1)
                    {
                        if (!selected.Contains(selectedindex))
                        {
                            selected.Add(selectedindex);
                            foreach (var mpcsitem in mpcs)
                            {
                                if (i == mpcsitem.ColIndex)
                                {
                                    try
                                    {
                                        if (All)
                                            mpcsitem.ctrlSelectedIndex = selectedindex;
                                        else

                                            mpcsitem.ctrl.SelectedIndex = selectedindex;
                                        break;
                                    }
                                    catch (Exception ex)
                                    {
                                        throw ex;
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }

            }
            for (int i = 0; i < mpcs.Count(); i++)
            {
                try
                {
                    int thisselindex;
                    if (All)
                        thisselindex = mpcs[i].ctrlSelectedIndex;
                    else
                        thisselindex = mpcs[i].ctrl.SelectedIndex;
                    if (thisselindex > 0)
                    {
                        for (int j = i + 1; j < mpcs.Count(); j++)
                        {
                            int otherselindex;
                            if (All)
                                otherselindex = mpcs[j].ctrlSelectedIndex;
                            else
                                otherselindex = mpcs[j].ctrl.SelectedIndex;
                            if (thisselindex == otherselindex)
                                throw new Exception("duplicatese index");
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        private void ResetHeaderMapping()
        {
            foreach (var item in HeaderColumnMaps)
            {
                item.SelectedColumn = -1;
            }
            foreach (var mpcsitem in mpcs)
            {
                if (All)
                    mpcsitem.ctrlSelectedIndex = -1;
                else
                {
                    ComboBox cbo = (ComboBox)mpcsitem.ctrl;
                    var iii = cbo.Items;
                    mpcsitem.ctrl.SelectedItem = -1;
                    mpcsitem.ctrl.SelectedItem = null;
                }
            }
        }

        object TotalHours;
        private void RemoveHeaders()
        {
            if (excelImportDS1.Excel.Count == 0)
                return;
            ExcelImportDS.ExcelRow exrowdel = excelImportDS1.Excel[0];
            exrowdel.Delete();
            excelImportDS1.AcceptChanges();
            exrowdel = excelImportDS1.Excel[0];
            exrowdel.Delete();
            excelImportDS1.AcceptChanges();
        }

        void RemoveFooter()
        {
            int ii = 0;
            while (ii < excelImportDS1.Excel.Count())
            {
                try
                {
                    ExcelImportDS.ExcelRow exrow = excelImportDS1.Excel[ii];

                    string jobcode = exrow[HIJobCodes.SelectedColumn].ToString();
                    if (jobcode == "")
                    {
                        bool PlantFound = false;
                        if (HIPlantNo.SelectedColumn!=-1)
                        {
                            string plantno = exrow[HIPlantNo.SelectedColumn].ToString();
                            if (plantno == "")
                            {
                                exrow.Delete();
                                excelImportDS1.AcceptChanges();
                            }
                            else
                            {
                                PlantFound = true;
                                ii++;
                            }
                        }
                        if (!PlantFound)
                        {
                            if (HITimeCode.SelectedColumn == -1)
                            {
                                exrow.Delete();
                                excelImportDS1.AcceptChanges();
                            }
                            else
                            {
                                string timecode = exrow[HITimeCode.SelectedColumn].ToString();
                                if (timecode == "")
                                {
                                    exrow.Delete();
                                    excelImportDS1.AcceptChanges();
                                }
                                else
                                    ii++;
                            }
                        }
                    }
                    else
                        ii++;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        /// <summary>
        /// validate electronic timeheet
        /// </summary>
        /// <param name="tabname"></param>
        /// <param name="Plant_ElseTSHours"></param>
        /// <param name="filename"></param>
        /// <returns></returns>
        private bool Validation(string tabname,frmExcel.Plant_TS123 Plant_ElseTSHours,string filename)
        {
            if (HIJobCodes.SelectedColumn == -1)
            {
                Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Job Codes", "Job code column not found", true, filename, Plant_ElseTSHours);
                return false;
            }

            if (Plant_ElseTSHours == frmExcel.Plant_TS123.TS1)
            {
                var tsh = db.Timesheets.Where(ts => ts.StaffID == EmpNo && ts.PayNoYear == PayNo).FirstOrDefault();
                if (tsh != null)
                {
                    var cnt = db.TimesheetDatas.Count(ts => ts.TimesheetID == tsh.TimesheetID);
                    if (cnt > 0)
                    {
                        Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Timesheet", "Duplicate timesheet for " + staff.FirstName + " " + staff.Surname, true, filename, Plant_ElseTSHours);
                        return false;
                    }
                }
            }
            int Bse;
            if (Plant_ElseTSHours== frmExcel.Plant_TS123.Plant)
                Bse = 3;
            else
                Bse = 10;
            bool valid = true;
            foreach (ExcelImportDS.ExcelRow exrow in excelImportDS1.Excel)
            {
                try
                {
                    string jobcode = exrow[HIJobCodes.SelectedColumn].ToString();
                    if (db.Jobs.Count(i => i.JobCode == jobcode) == 0)
                    {
                        if (db.Jobs.Count() != 0)
                        {
                            if ((jobcode == "0.00") || (jobcode == "0000-0000-0000"))
                                exrow[HIJobCodes.SelectedColumn] = "";
                            else
                            {
                                exrow[HIJobCodes.SelectedColumn] = jobcode;// "xxxx-xxxx-xxxx";//modified thams Mar 2023
                                if (jobcode!="")
                                   Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Job Codes", jobcode + " is not a valid code", false, filename, Plant_ElseTSHours);
                            }
                        }
                    }
                    int TC_ALL_Pop = 0;
                    if (Plant_ElseTSHours != frmExcel.Plant_TS123.Plant)
                    {
                        if (exrow[HITimeCode.SelectedColumn].ToString() != "")
                        {
                            if (exrow[HITimeCode.SelectedColumn].Equals("-"))
                                exrow[HITimeCode.SelectedColumn] = "";//dean default
                            else if (exrow[HITimeCode.SelectedColumn].ToString().Contains("N/A"))
                                exrow[HITimeCode.SelectedColumn] = "";

                            else
                            {
                                if (tabname.ToLower().Contains("_t1"))
                                {
                                    try
                                    {
                                        object T1_code_obj = exrow[HITimeCode.SelectedColumn];
                                        decimal T1_code = Convert.ToDecimal(T1_code_obj);
                                        var tcd = db.PayComponents.Where(tc => tc.PayCompCode == T1_code).FirstOrDefault();
                                        if (tcd == null)
                                        {
                                            Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "T!_PayCompID", T1_code + " T1 Paycomp not found in Timecode table", true, filename, Plant_ElseTSHours);
                                            return false;
                                        }
                                        TC_ALL_Pop++;
                                    }
                                    catch (Exception ex)
                                    {
                                        throw ex;
                                    }
                                }
                            }
                        }
                        if (exrow[HIAllowance.SelectedColumn].ToString() != "")
                        {
                            try
                            {
                                string allstringx = exrow[HIAllowance.SelectedColumn].ToString();
                                allstringx = allstringx.Replace(".00", "");
                                exrow[HIAllowance.SelectedColumn] = allstringx;
                                int allcode = Convert.ToInt32(allstringx);
                                if (db.PayComponents.Count(i => i.PayCompCode == allcode /*|| i.T1_Code == allcode*/) == 0)
                                {
                                    valid = false;
                                    Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "AllowCode ", allcode + " is not a valid option", true, filename, Plant_ElseTSHours);
                                }
                                else if (allstringx != "0")
                                    TC_ALL_Pop++;
                            }
                            catch (Exception ex)
                            {
                                string err = "Allowance Code is not an int " + exrow[HIAllowance.SelectedColumn].ToString() + " - " + ex.Message;
                                Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "AllowCode", err, true, filename, Plant_ElseTSHours);
                                exrow[HIAllowance.SelectedColumn] = "";
                                valid = false;
                            }

                        }

                        if (TC_ALL_Pop != 1)
                        {
                            string iss;
                            if (TC_ALL_Pop == 2)
                                iss = "Entry Has both Allowance and timecode - which T1Paycomponent";
                            else
                                iss = "Entry does not have Allowance nor timecode - no T1Paycomponent";
                            Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "TimeCode_AllowCode", iss, true, filename, Plant_ElseTSHours);
                        }
                    }
                    else
                    {
                        ;//could check if plant no and no job no then error
                    }
                    for (int dayind = 0; dayind < 14; dayind++)
                    {
                        try
                        {
                            int selcol = HeaderColumnMaps[Bse + dayind].SelectedColumn;
                            if (selcol != -1)
                            {
                                string hours = exrow[selcol].ToString();
                                if (hours != "")
                                {
                                    decimal value;
                                    if (Decimal.TryParse(hours, out value))

                                        ;//ok
                                    else
                                    {
                                        valid = false;
                                        Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Hours", hours + " is not a decimal - data ignored",true, filename, Plant_ElseTSHours);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            ex = ex;
                        }
                    }

                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            return valid;
        }


        void ClearTimesheetInDB(int TimesheetID,int EmpID)
        {

            var delts = db.TimesheetDatas.Where(i => i.TimesheetID == TimesheetID).ToList();
            db.TimesheetDatas.DeleteAllOnSubmit(delts);
            db.SubmitChanges();

            var deltse = db.TSExceptions.Where(i => i.TimesheetID == TimesheetID);
            db.TSExceptions.DeleteAllOnSubmit(deltse);
            db.SubmitChanges();

            if (EmpID < 0)
            {
                var delt = db.Timesheets.Where(i => i.TimesheetID == TimesheetID);
                db.Timesheets.DeleteAllOnSubmit(delt);
                db.SubmitChanges();
            }
        }

        /// <summary>
        ///this method could reference a table which checks which employeetype F,C,P can book Paycomponents - currrently doesn't check
        /// </summary>
        /// <param name="EmpType"></param>
        /// <param name="TimeCode"></param>
        /// <returns></returns>
        bool CheckEmpTimeCode(string EmpType,decimal TimeCode)
        {
             return true;
            if (db.EmpTypeTimeCodes.Any(et=>et.EmpType==EmpType && et.TimeCode==TimeCode))
                return true;
            return false;
        }

        static bool DataRead;
        public int PlantSuffix = 9000000;
        private void ImportMigrate(frmExcel.Plant_TS123 Plant_TS, double TotalHours, string filename)
        {
            if (excelImportDS1.Excel.Count == 0)
            {
                if (Plant_TS== frmExcel.Plant_TS123.TS1)
                    Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Data", EmpNo + " timesheet is empty", true, filename, Plant_TS);
                return;
            }
            try
            {
                DataRead = false;
                DateTime StartDate = db.PayYears.Where(ii => ii.PayNoYear == PayNo).FirstOrDefault().StartDate;
                var emp = db.Employees.Where(ee => ee.T1EmpNo == EmpNo).FirstOrDefault();
                if (emp == null)
                {
                    Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "EmpNo", EmpNo + " is not found", true, filename,Plant_TS);
                    return;
                }
                if (emp.T1EmpNo == 0)
                {
                    Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "EmpNo", EmpNo + " is blank in T1", true, filename, Plant_TS);
                    return;
                }
                int rowid = 0;
                foreach (ExcelImportDS.ExcelRow exrow in excelImportDS1.Excel)
                {
                    try
                    {
                        int Bse;
                        if (Plant_TS == frmExcel.Plant_TS123.Plant)
                            Bse = 3;
                        else
                            Bse = 10;
                        string Desc;
                        decimal TimeCode = 0;
                        int ClassNo = 0;
                        int AllowCode = 0;
                        int PlantNo = 0;
                        string jobcode = exrow[HIJobCodes.SelectedColumn].ToString();
                        if (jobcode == "9012145")
                            jobcode = "9012145";
                        if (jobcode != "")
                        {
                            if ((db.Jobs.Count(i => i.JobCode == jobcode) == 0) && (db.Jobs.Count() > 0))
                            {
                                Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "JobCodes", "JobCodes/Workorders table does not contain an entry for " + jobcode, true, filename, Plant_TS);
                            }
                        }
                        if (Plant_TS == frmExcel.Plant_TS123.Plant)
                        {
                            Desc = "Plant";
                            object pn = exrow[HIPlantNo.SelectedColumn];
                            if (pn.ToString()!="")
                            PlantNo = Convert.ToInt32(pn);

                            if (db.Plants.Count() > 0)
                            {
                                if (PlantNo != 0)
                                {
                                    if ((db.Plants.Count(i => i.PlantSource == (PlantSuffix + PlantNo).ToString()) == 0))// && (db.Plants.Count() > 0))
                                    {
                                        //todotim
                                        Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "plantno", "plantno table does not contain code " + PlantNo.ToString(), true, filename, Plant_TS);
                                    }
                                }
                                Plant pl = db.Plants.Where(p => p.PlantSource == "PL" + PlantNo.ToString()).FirstOrDefault();
                                if (pl == null)
                                    pl = db.Plants.Where(p => p.PlantSource.ToLower().Contains(PlantNo.ToString().ToLower())).FirstOrDefault();
                                if (pl == null)
                                {
                                    Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Plant No", "Plant No " + PlantNo.ToString() + " not found", false, filename, Plant_TS);
                                    //Valid = false;
                                }
                                else
                                    PlantNo = pl.PlantTarget;
                            }
                            else
                            {
                                if (PlantNo > 0)
                                {
                                    PlantNo = PlantNo + AddValueToPlantForT1Map;
                                    //plant no exists

                                }
                            }
                            if (PlantNo > 0)
                            {
                                if (jobcode == "")
                                    Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "PlantJobNo ", "Plant No Exists:" + PlantNo.ToString() + " yet JobNo is blank", true, filename, frmExcel.Plant_TS123.Plant);

                            }

                        }
                        else
                        {
                            if (HIDesc.SelectedColumn == -1)
                                Desc = "DescColumnNotfound";
                            else
                                Desc = exrow[HIDesc.SelectedColumn].ToString();

                            object ClassObj = exrow[HIClass.SelectedColumn];
                            ClassObj = ClassObj.ToString().Replace(" ", "");
                            if (!ClassObj.Equals(""))
                            {
                                ClassObj = ClassObj.ToString().Replace(".00", "");
                                try
                                {
                                    ClassNo = Convert.ToInt32(ClassObj);
                                }
                                catch (Exception ex)
                                {
                                    ClassNo = 0;
                                    string error = "Classification value in row " + (rowid+1).ToString() + " is value " + ClassObj.ToString() + " which is not valid. Classification value must be an integer ie number with no decimal places";
                                    Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Invalid Class no", error, true, filename, Plant_TS);

                                }
                            }
                            else
                            {
                                throw new Exception("No Emp ClassNo");
                            }
                            object AllowObj = exrow[HIAllowance.SelectedColumn];
                            if (!AllowObj.Equals(""))
                            {
                                AllowCode = Convert.ToInt32(AllowObj);
                            }
                            if (db.PayComponents.Count(all => all.PayCompCode == AllowCode) == 0)
                            {
                                Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "AllowCode", "AllowCode " + AllowCode + " not found in list", true, filename, Plant_TS);
                            }
                            object tc = exrow[HITimeCode.SelectedColumn];
                            if (tc.ToString() != "")
                            {
                                TimeCode = Convert.ToDecimal(exrow[HITimeCode.SelectedColumn]);
                                if (TimeCode == 1230)
                                    TimeCode = 137;
                                //negate hours
                            }
                            //if (!CheckEmpTimeCode(emp.Emp_Type.ToString(), TimeCode))
                            //{
                            //    Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Employee Type", "Employee Type " + emp.Emp_Type + " TImeCode" + TimeCode.ToString() + " is not a valid timecode for this emp type", true, filename);
                            //}
                        }
                        double daysum = ReadDayColumns(Plant_TS, StartDate, rowid, exrow, Bse, Desc, TimeCode, ClassNo, AllowCode, PlantNo, jobcode);
                        try
                        {
                            if (Plant_TS == frmExcel.Plant_TS123.Plant)
                                CompareRowTotalToDayAggregateAndCorrectIfDifferent(TimeType.Plant, HITotalPlantHours, filename, StartDate, rowid, exrow, Desc, PlantNo, ClassNo, AllowCode, jobcode, daysum, Plant_TS);
                            else
                            {

                                if (AllowCode != 0)//if this is an allowance entry
                                    CompareRowTotalToDayAggregateAndCorrectIfDifferent(TimeType.Allowance, HIAllowUnits, filename, StartDate, rowid, exrow, Desc, TimeCode, ClassNo, AllowCode, jobcode, daysum,Plant_TS);
                                else
                                {
                                    TimeType timetype = GetTimeType(TimeCode);
                                    if (timetype == TimeType.Overtime)
                                        CompareRowTotalToDayAggregateAndCorrectIfDifferent(timetype, HIOvertime, filename, StartDate, rowid, exrow, Desc, TimeCode, ClassNo, AllowCode, jobcode, daysum,Plant_TS);
                                    else if (timetype == TimeType.Ordinary)//.CODEDESC.ToLower().Contains("leave")
                                        CompareOrdHoursRowTotalToDayAggregateAndCorrectIfDifferent(timetype, filename, StartDate, rowid, exrow, Desc, TimeCode, ClassNo, AllowCode, jobcode, daysum, Plant_TS);
                                    else//leave
                                        CompareRowTotalToDayAggregateAndCorrectIfDifferent(timetype, HILvHours, filename, StartDate, rowid, exrow, Desc, TimeCode, ClassNo, AllowCode, jobcode, daysum, Plant_TS);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error ImportMigrate Loop. filename:" +filename + ":"+ ex.Message);
                        }
                        rowid++;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error ImportMigrate Main.  Filenamme:" +filename+": "+ ex.Message);
                    }
                }
                if (Plant_TS == frmExcel.Plant_TS123.TS1)
                {
                    if ((!DataRead) && (ts != null))
                        Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Data", "No Data Read", true, filename, Plant_TS);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error ImportMigrate Outer" + ex.Message);
            }
        }

        private double ReadDayColumns(frmExcel.Plant_TS123 Plant_TS, DateTime StartDate, int rowid, ExcelImportDS.ExcelRow exrow, int Bse, string Desc, decimal TimeCode, int ClassNo, int AllowCode, int PlantNo, string jobcode)
        {
            double daysum = 0;
            for (int dayind = 0; dayind < 14; dayind++)
            {
                try
                {
                    int selcol = HeaderColumnMaps[Bse + dayind].SelectedColumn;
                    if (selcol != -1)
                    {
                        string hours = exrow[selcol].ToString();
                        if (hours != "")
                        {
                            double hour = Convert.ToDouble(hours);
                            daysum += hour;
                            if (Plant_TS == frmExcel.Plant_TS123.Plant)
                                AddPlant(StartDate, rowid, Desc, jobcode, PlantNo, dayind, hour, ts.TimesheetID, "Extracted from Day column in Excel timesheet");
                            else
                                AddTD(StartDate, rowid, Desc, jobcode, TimeCode, ClassNo, AllowCode, dayind, hour, ts.TimesheetID, "Extracted from Day column in Excel timesheet");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error ImportMigrate day sum Loop" + ex.Message);
                }
            }
            return daysum;
        }

        /*
                                        CompareRowTotalToDayAggregateAndCorrectIfDifferent(timetype, HILvHours, filename, StartDate, rowid, exrow, Desc, TimeCode, ClassNo, AllowCode, jobcode, daysum);

                               //    double otherAgg = 0;
                               //    object oh = exrow[HILvHours.SelectedColumn];
                               //    if (!exrow[HIOrdHours.SelectedColumn].Equals(""))
                               //    {
                               //        otherAgg = Convert.ToDouble(exrow[HIOrdHours.SelectedColumn]);
                               //        Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Absense Total", "Timecode:" + TimeCode.ToString() + " is marked as Absence yet Leave aggregate is empty and normal hours aggregate is " + otherAgg.ToString(), false, filename);

                               //    }
                               //    if (!oh.Equals(""))
                               //    {
                               //        double ordhour = Convert.ToDouble(oh) + otherAgg;
                               //        ordhour = Math.Round(ordhour, 2);
                               //        daysum = Math.Round(daysum, 2);

                               //        if (Math.Abs(ordhour - daysum) > 0.1)
                               //        {
                               //            if (Math.Abs(Math.Abs(ordhour) - Math.Abs(daysum)) > 0.1)
                               //            {
                               //                if (CorrectDaySumToTotalColumnForPayCompCodeBool(TimeCode))
                               //                {
                               //                    AddTD(StartDate.Date, rowid, Desc, jobcode, TimeCode, ClassNo, AllowCode, 13, ordhour, ts.TimesheetID, "Correction to make sum of day sum amount " + daysum.ToString() + " equal total column (" + ordhour.ToString() + ")");
                               //                    //Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Leave Hours " + TimeCode.ToString() + " Correction", "Timecode:" + TimeCode.ToString() + "Allowance Code:" + AllowCode.ToString() + " Leave Hours Total column " + ordhour.ToString() + " does not match day aggregate " + hoursum.ToString(), false, filename);
                               //                }
                               //                else
                               //                {
                               //                    //Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Leave Hours " + TimeCode.ToString() + " Correction", "Timecode:" + TimeCode.ToString() + "Allowance Code:" + AllowCode.ToString() + " Leave Hours Total column " + ordhour.ToString() + " does not match day aggregate " + hoursum.ToString(), false, filename);
                               //                }
                               //            }
                               //        }
                               //    }

* 
                                        CompareRowTotalToDayAggregateAndCorrectIfDifferent(timetype, HIOvertime, filename, StartDate, rowid, exrow, Desc, TimeCode, ClassNo, AllowCode, jobcode, daysum);

                               //object oh = exrow[HIOvertime.SelectedColumn];
                               //if (!oh.Equals(""))
                               //{
                               //    double ordhour = Convert.ToDouble(oh);
                               //    ordhour = Math.Round(ordhour, 2);
                               //    daysum = Math.Round(daysum, 2);
                               //    if (Math.Abs(ordhour - daysum) > 0.1)
                               //    {
                               ////        object ord = exrow[HIOrdHours.SelectedColumn];
                               //        if (!oh.Equals(""))
                               //        {
                               //            double ordagg = Convert.ToDouble(ord);
                               //            double diff = ordagg + ordhour - daysum;
                               //            if (Math.Abs(diff) > (double)0.1)
                               //            {
                               //                if (CorrectDaySumToTotalColumnForPayCompCodeBool(TimeCode))

                               //                    AddTD(StartDate.Date, rowid, Desc, jobcode, TimeCode, ClassNo, AllowCode, 13, ordhour + ordagg - daysum, ts.TimesheetID, "Correction to make sum of day overtime equal overtime total");

                               //                Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "OT Correction", "Timecode:" + TimeCode.ToString() + "Allowance Code:" + AllowCode.ToString() + " OT totals (" + ordagg.ToString() + "," + ordhour.ToString() + ") column does not match day aggregate " + daysum.ToString(), false, filename);

                               //            }
                               //        }
                               //        else
                               //        {
                               //            if (CorrectDaySumToTotalColumnForPayCompCodeBool(TimeCode))
                               //                AddTD(StartDate.Date, rowid, Desc, jobcode, TimeCode, ClassNo, AllowCode, 13, ordhour - daysum, ts.TimesheetID, "Correction to make sum of day overtime equal overtime total");
                               //            Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "OT Correction", "Timecode:" + TimeCode.ToString() + "Allowance Code:" + AllowCode.ToString() + " OT totals column does not match day aggregate", false, filename);
                               //        }
                               //    }
                               //}

* 
                                CompareRowTotalToDayAggregateAndCorrectIfDifferent(TimeType.Plant, HITotalPlantHours, filename, StartDate, rowid, exrow, Desc, PlantNo, ClassNo, AllowCode, jobcode, daysum);
                       //object PlantTotal = exrow[HITotalPlantHours.SelectedColumn];
                       //if (!PlantTotal.Equals(""))
                       //{
                       //    double PlantTotalVal = Convert.ToDouble(PlantTotal);
                       //    if (!AmountsEqualOrWithinPointOne(PlantTotalVal, daysum))
                       //    {
                       //        //AddPlant(StartDate, rowid, Desc, jobcode, PlantNo, 0, PlantTotalVal - daysum, ts.TimesheetID);
                       //        Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Plant Total", "Plant:" + PlantNo.ToString() + " Totals " + PlantTotalVal.ToString() + " total column does not match day aggregate " + daysum.ToString(), false, filename);
                       //    }
                       //}

*/
        TimeType GetTimeType(decimal TimeCode)
        {
            var tcs = db.PayComponents.Where(tt => tt.PayCompCode == TimeCode).FirstOrDefault();
            if (tcs.PayCompTypeDesc != null)
            {
                if (tcs.PayCompTypeDesc.ToLower() == "Overtime".ToLower())
                    return TimeType.Overtime;
                if (tcs.PayCompTypeDesc.ToLower() == "Accruals".ToLower())
                    return TimeType.Overtime;

                if (tcs.PayCompTypeDesc != "Absences")
                    return TimeType.Ordinary;
                return TimeType.Leave;
            }
            if (tcs.PayCompDesc.ToLower().Contains("overtime"))
                return TimeType.Overtime;
            if (tcs.PayCompTypeDesc.ToLower().Contains(" ot "))
                return TimeType.Overtime;
            return TimeType.Ordinary;
        }

        public enum TimeType
        {
            Ordinary=0,
            Leave=1,
            Overtime=2,
            Allowance=3,
            Plant=4
        }


        private void CompareOrdHoursRowTotalToDayAggregateAndCorrectIfDifferent(TimeType timetype, string filename, DateTime StartDate, int rowid, ExcelImportDS.ExcelRow exrow, string Desc, decimal TimeCode, int ClassNo, int AllowCode, string jobcode, double daysum,frmExcel.Plant_TS123 Plant_TS)
        {
            /*
                                                        object oh = exrow[HIOrdHours.SelectedColumn];
                                            {
                                                double ordhour = 0;
                                                if (!oh.Equals(""))
                                                    ordhour = Convert.ToDouble(oh);
                                                ordhour = Math.Round(ordhour, 2);
                                                daysum = Math.Round(daysum, 2);
                                                if (Math.Abs(ordhour - daysum) > 0.1)
                                                {
                                                    double hrsdiff = ordhour - daysum;

                                                    if (CorrectDaySumToTotalColumnForPayCompCodeBool(TimeCode))
                                                    {
                                                        DateTime start_date = StartDate.Date.AddDays(13);
                                                        DateTime end_date = start_date.AddHours(-hrsdiff);

                                                        var tds = db.TimesheetDatas.Where(td => td.start_date == start_date && td.end_date == end_date && td.TImeCode == TimeCode && td.TimesheetID == ts.TimesheetID).ToList();
                                                        if (tds.Count() > 0)
                                                        {
                                                            var td = tds.FirstOrDefault();
                                                            db.TimesheetDatas.DeleteOnSubmit(td);
                                                            db.SubmitChanges();
                                                        }
                                                        else
                                                        {
                                                            AddTD(StartDate.Date, rowid, Desc, jobcode, TimeCode, ClassNo, AllowCode, 13, hrsdiff, ts.TimesheetID, "Correction to make sum of day hours " + daysum.ToString() + " equal ord hours total column (" + ordhour.ToString() + ")");
                                                        }
                                                        if (Math.Abs(ordhour) != Math.Abs(daysum))
                                                            Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Ordinary Hours Correction", "Timecode:" + TimeCode.ToString() + "Allowance Code:" + AllowCode.ToString() + " Ord Hours Total (" + ordhour.ToString() + ") column does not match day aggregate (" + daysum.ToString() + ")", false, filename);
                                                    }
                                                    else
                                                    {
                                                        //    AddTD(StartDate.Date, rowid, Desc, jobcode, TimeCode, ClassNo, AllowCode, 13, hrsdiff, ts.TimesheetID, "Correction to make sum of day ordinary hours " + hoursum.ToString() + " equal ord hours total column (" + ordhour.ToString() + ")");
                                                        Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Ordinary Hours Correction", "Timecode:" + TimeCode.ToString() + "Allowance Code:" + AllowCode.ToString() + " Ord Hours Total (" + ordhour.ToString() + ") column does not match day aggregate (" + daysum.ToString() + ")", false, filename);
                                                    }
                                                }
                                            }
            */
            object OrdRowTotal = exrow[HIOrdHours.SelectedColumn];
            double OrdRowTotalVal = 0;
            if (!OrdRowTotal.Equals(""))
                OrdRowTotalVal = Convert.ToDouble(OrdRowTotal);
            if (AmountsEqualOrWithinPointOne(OrdRowTotalVal, daysum))
                return;
            if (CorrectDaySumToTotalColumnForPayCompCodeBool(TimeCode))
            {
                DateTime start_date = StartDate.Date.AddDays(13);
                DateTime end_date = start_date.AddHours(daysum - OrdRowTotalVal);
                var tds = db.TimesheetDatas.Where(td => td.start_date == start_date && td.end_date == end_date && td.TImeCode == TimeCode && td.TimesheetID == ts.TimesheetID).ToList();
                if (tds.Count() > 0)
                {
                    var td = tds.FirstOrDefault();
                    db.TimesheetDatas.DeleteOnSubmit(td);
                    db.SubmitChanges();
                }
                else
                    AddTD(StartDate.Date, rowid, Desc, jobcode, TimeCode, ClassNo, AllowCode, 13, OrdRowTotalVal - daysum, ts.TimesheetID, "Correction to make sum of day allowance equal to total " + HIOrdHours.Header);
                if (Math.Abs(OrdRowTotalVal) != Math.Abs(daysum))
                    Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, HIOrdHours.Header + "Total Correction", "Timecode:" + TimeCode.ToString() + "Allowance Code:" + AllowCode.ToString() + " " + HIOrdHours.Header + " Total column (" + OrdRowTotalVal.ToString() + ") does not match day aggregate " + daysum.ToString(), false, filename, Plant_TS);
            }
            else
                //for ord hours if CorrectDaySumToTotalColumnForPayCompCodeBool and (Math.Abs(RowTotal) == Math.Abs(daysum)) then no exception
                Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, HIOrdHours.Header + "Total Correction", "Timecode:" + TimeCode.ToString() + "Allowance Code:" + AllowCode.ToString() + " " + HIOrdHours.Header + " Total column (" + OrdRowTotalVal.ToString() + ") does not match day aggregate " + daysum.ToString(), false, filename, Plant_TS);
        }

        private void CompareRowTotalToDayAggregateAndCorrectIfDifferent(TimeType timetype, HeaderIndex TotalColumnMap, string filename, DateTime StartDate, int rowid, ExcelImportDS.ExcelRow exrow, string Desc, object TimeCode_PlantNo, int ClassNo, int AllowCode, string jobcode, double daysum,frmExcel.Plant_TS123 Plant_TS)
        {
            /*
              object PlantTotal = exrow[HITotalPlantHours.SelectedColumn];
                                    if (!PlantTotal.Equals(""))
                                    {
                                        double PlantTotalVal = Convert.ToDouble(PlantTotal);
                                        if (!AmountsEqualOrWithinPointOne(PlantTotalVal, daysum))
                                        {
                                            AddPlant(StartDate, rowid, Desc, jobcode, PlantNo, 0, PlantTotalVal - daysum, ts.TimesheetID);
                                            Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, "Plant Total", "Plant:" + PlantNo.ToString() + " Totals " + PlantTotalVal.ToString() + " total column does not match day aggregate " + daysum.ToString(), false, filename);
                                        }
                                    }

    */
            double OrdRowTotalVal = 0;
            if (HIOrdHours.SelectedColumn != -1)
            {
                object OrdRowTotal = exrow[HIOrdHours.SelectedColumn];

                if (!OrdRowTotal.Equals(""))
                {
                    OrdRowTotalVal = Convert.ToDouble(OrdRowTotal);
                    if (OrdRowTotalVal != 0)
                    {
                        if (TotalColumnMap != HIOrdHours)
                            Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, TotalColumnMap.Header + " Total", "Timecode:" + TimeCode_PlantNo.ToString() + " is marked as " + timetype + " yet ordinary hours total is populated " + exrow[HIOrdHours.SelectedColumn].ToString(), false, filename, Plant_TS);
                    }
                }
            }
            double RowTotalVal = 0;
            object RowTotal = exrow[TotalColumnMap.SelectedColumn];

            if (!RowTotal.Equals(""))

                RowTotalVal = Convert.ToDouble(RowTotal);
            //if (AmountsEqualOrWithinPointOne(RowTotalVal, daysum))
            //    return;
            ////lets check total wasn't put in ordinary hours by mistake
            //if (OrdRowTotalVal!=0)
            //{
            if (AmountsEqualOrWithinPointOne(OrdRowTotalVal + RowTotalVal, daysum))
                return;
            // }
            if (timetype == TimeType.Plant)
                AddPlant(StartDate, rowid, Desc, jobcode, (int)TimeCode_PlantNo, 0, RowTotalVal - daysum, ts.TimesheetID, "Correction to make plant total add to day sum");
            else
            {
                if (CorrectDaySumToTotalColumnForPayCompCodeBool((decimal)TimeCode_PlantNo))
                    AddTD(StartDate.Date, rowid, Desc, jobcode, (decimal)TimeCode_PlantNo, ClassNo, AllowCode, 13, RowTotalVal - daysum, ts.TimesheetID, "Correction to make sum of day allowance equal to total " + TotalColumnMap.Header);
            }
            //for ord hours if CorrectDaySumToTotalColumnForPayCompCodeBool and (Math.Abs(RowTotal) == Math.Abs(daysum)) then no exception
            Exception(ts.TimesheetID, staff.FirstName + " " + staff.Surname, TotalColumnMap.Header + "Total Correction", "Timecode:" + TimeCode_PlantNo.ToString() + "Allowance Code:" + AllowCode.ToString() + " " + TotalColumnMap.Header + " Total column, Ord Total (" + RowTotalVal.ToString() + "," + OrdRowTotalVal.ToString() + ") does not match day aggregate " + daysum.ToString(), false, filename, Plant_TS);
        }

        bool AmountsEqualOrWithinPointOne(double TotalColumn, double DaySum)
        {
            double difference = TotalColumn - DaySum;
            if (Math.Abs(difference) > 0.1)
                return false;
            //allowanceunitsval = Math.Round(allowanceunitsval, 2);
            //daysum = Math.Round(daysum, 2);
            //if (Math.Abs(allowanceunitsval - daysum) > 0.1)
            return true;
        }

        /// <summary>
        /// add plant time to database
        /// </summary>
        /// <param name="StartDate"></param>
        /// <param name="rowid"></param>
        /// <param name="Desc"></param>
        /// <param name="jobcode"></param>
        /// <param name="PlantNo"></param>
        /// <param name="dayind"></param>
        /// <param name="hour"></param>
        /// <param name="TimesheetID"></param>
        static void AddPlant(DateTime StartDate, int rowid, string Desc, string jobcode, int PlantNo, int dayind, double hour, int TimesheetID,string Source)
        {
            if (hour != 0)
            {
                DataRead = true;
                TimesheetData tsd = new TimesheetData();
                tsd.TimesheetID = TimesheetID;
                tsd.job = jobcode;
                tsd.PlantNo=PlantNo;
                tsd.Row = rowid;
                tsd.Source = Source;
                tsd.start_date = StartDate.Date.AddDays(dayind);
                tsd.end_date = tsd.start_date.AddHours(hour);
                tsd.Description = Desc;
                db.TimesheetDatas.InsertOnSubmit(tsd);
                db.SubmitChanges();
            }
        }

        /// <summary>
        /// Add Employee Worked time to Database
        /// </summary>
        /// <param name="StartDate"></param>
        /// <param name="rowid"></param>
        /// <param name="Desc"></param>
        /// <param name="jobcode"></param>
        /// <param name="TimeCode"></param>
        /// <param name="ClassNo"></param>
        /// <param name="AllowCode"></param>
        /// <param name="dayind"></param>
        /// <param name="hour"></param>
        /// <param name="TimesheetID"></param>
        /// <param name="source"></param>
        static void AddTD(DateTime StartDate, int rowid, string Desc, string jobcode, decimal TimeCode, int ClassNo, int AllowCode, int dayind, double hour, int TimesheetID,string source)
        {
            if (hour != 0)
            {
                try
                {
                    DataRead = true;
                    TimesheetData tsd = new TimesheetData();
                    tsd.TimesheetID = TimesheetID;
                    tsd.Source = source;
                    tsd.job = jobcode;
                    tsd.ClassNo = ClassNo;
                    tsd.AllowanceCode = AllowCode;
                    tsd.Row = rowid;

                    tsd.start_date = StartDate.Date.AddDays(dayind);
                    tsd.end_date = tsd.start_date.AddHours(hour);
                    tsd.TImeCode = TimeCode;
                    tsd.Description = Desc;
                    db.TimesheetDatas.InsertOnSubmit(tsd);
                    db.SubmitChanges();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        void ResizeGrid()
        {
            if (All)
                return;
            int mywidth = 80;
            int start = 40;
            for (int i = 0; i < mpcs.Count(); i++)
            {
                MapCtrl item = mpcs[i];
                item.ctrl.Left = start + i * mywidth;
                item.ctrl.Width = mywidth;
            }
            foreach (DataGridViewColumn item in dataGridView1.Columns)
            {
                item.Width = mywidth;
            }
        }

        HeaderIndex GetHeaderColumnMapItem(int i)
        {
            foreach (HeaderIndex itemHI in HeaderColumnMaps)
            {
                if (itemHI.SelectedColumn == i)
                    return itemHI;
            }
            return null;
        }

        void ResizeGrid2()
        {
            try
            {
                if (All) return;
                int mywidth = 80;
                int left = 40;
                for (int i = 0; i < mpcs.Count(); i++)
                {
                    MapCtrl item = mpcs[i];
                    HeaderIndex xhi = GetHeaderColumnMapItem(i);
                    if (xhi == null)
                        mywidth = 40;
                    else
                    {
                        if (xhi.dayindex == -1)
                            mywidth = 80;
                        else
                            mywidth = 40;
                    }
                    item.ctrl.Left = left;// start + i * mywidth;
                    item.ctrl.Width = mywidth;
                    left += mywidth;
                    DataGridViewColumn citem = dataGridView1.Columns[i];
                    citem.Width = mywidth;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        HeaderIndex HIPlantNo = new HeaderIndex("Plant Number");
        HeaderIndex HITotalPlantHours = new HeaderIndex("Total Plant Hours");

        HeaderIndex HIDesc = new HeaderIndex("Description / Location of Work", "Description Classification / Allowance");
        HeaderIndex HITimeCode = new HeaderIndex("Time Code");
        HeaderIndex HIJobCodes = new HeaderIndex("Job / Account Number","job");
        HeaderIndex HIClass = new HeaderIndex("Class");
        HeaderIndex HIAllowance = new HeaderIndex("Allow Code");
        HeaderIndex HIOrdHours = new HeaderIndex("Ord");
        HeaderIndex HILvHours = new HeaderIndex("LeaveHours", "Leave");
        HeaderIndex HIAllowUnits = new HeaderIndex("Allow Units", "Allowance");// Units);
        HeaderIndex HIOvertime = new HeaderIndex("Overtime", "O'time");// Units);

        void InitHIPlant()
        {
            try
            {

                HeaderColumnMaps = new HeaderIndex[3 + 14];
                HeaderColumnMaps[0] = HIPlantNo;
                HeaderColumnMaps[1] = HITotalPlantHours;
                HeaderColumnMaps[2] = HIJobCodes;
                DateTime StartDate = db.PayYears.Where(ii => ii.PayNoYear == PayNo).FirstOrDefault().StartDate;
                for (int dayind = 0; dayind < 14; dayind++)
                {
                    HeaderColumnMaps[3 + dayind] = new HeaderIndex(StartDate.AddDays(dayind).Day.ToString(), dayind);
                }
                int i = 0;
                foreach (HeaderIndex item in HeaderColumnMaps)
                {
                    item.Index = i;
                    i++;
                }
                if (All)
                {
                    foreach (var mpcsitem in mpcs)
                    {
                        if (mpcsitem.ctrlItems == null)
                            mpcsitem.ctrlItems = new List<string>();
                        else
                            mpcsitem.ctrlItems.Clear();
                        foreach (HeaderIndex Hiitem in HeaderColumnMaps)
                        {
                            mpcsitem.ctrlItems.Add(Hiitem.Header);
                        }
                    }

                }
                else
                {
                    foreach (var mpcsitem in mpcs)
                    {
                        mpcsitem.ctrl.Items.Clear();
                        foreach (HeaderIndex Hiitem in HeaderColumnMaps)
                        {
                            mpcsitem.ctrl.Items.Add(Hiitem.Header);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void InitHI()
        {
            try
            {

                HeaderColumnMaps = new HeaderIndex[10 + 14];
                HeaderColumnMaps[0] = HIDesc;
                HeaderColumnMaps[1] = HITimeCode;
                HeaderColumnMaps[2] = new HeaderIndex("Day");
                HeaderColumnMaps[3] = HIJobCodes;
                HeaderColumnMaps[4] = HIClass;
                HeaderColumnMaps[5] = HIAllowance;
                HeaderColumnMaps[6] = HIOrdHours;
                HeaderColumnMaps[7] = HILvHours;
                HeaderColumnMaps[8] = HIAllowUnits;
                HeaderColumnMaps[9] = HIOvertime;
                
                HIDesc.SelectedColumn = -1;
                HIDesc.dayindex = -1;

                HITimeCode.SelectedColumn = -1;
                HITimeCode.dayindex = -1;

                HIJobCodes.SelectedColumn = -1;
                HIJobCodes.dayindex = -1;

                HIClass.SelectedColumn = -1;
                HIClass.dayindex = -1;

                HIAllowance.SelectedColumn = -1;
                HIAllowance.dayindex = -1;

                DateTime StartDate = db.PayYears.Where(ii => ii.PayNoYear == PayNo).FirstOrDefault().StartDate;
                for (int dayind = 0; dayind < 14; dayind++)
                {
                    try
                    {
                        HeaderColumnMaps[10 + dayind] = new HeaderIndex(StartDate.AddDays(dayind).Day.ToString(), dayind);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }

                int i = 0;

                foreach (HeaderIndex item in HeaderColumnMaps)
                {
                    item.SelectedColumn = -1;
                    //item.dayindex = -1;

                    item.Index = i;
                    i++;
                }
                if (All)
                {
                    foreach (var mpcsitem in mpcs)
                    {
                        if (mpcsitem.ctrlItems == null)
                            mpcsitem.ctrlItems = new List<string>();
                        else
                            mpcsitem.ctrlItems.Clear();
                        foreach (HeaderIndex Hiitem in HeaderColumnMaps)
                        {
                            mpcsitem.ctrlItems.Add(Hiitem.Header);
                        }
                    }

                }
                else
                {
                    foreach (var mpcsitem in mpcs)
                    {
                        mpcsitem.ctrl.Items.Clear();
                        foreach (HeaderIndex Hiitem in HeaderColumnMaps)
                        {
                            mpcsitem.ctrl.Items.Add(Hiitem.Header);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        void FillEmployeeCombo()
        {
            tsbCboEmp.Items.Add("-All-");
            foreach (var item in db.Employees)
            {
                tsbCboEmp.Items.Add(item.T1EmpNo + " - " + item.Surname + ", " + item.FirstName);

            }
        }

        int _AddValueToPlantForT1Map = -1;
        int AddValueToPlantForT1Map
        {
            get
            {
                if (_AddValueToPlantForT1Map == -1)
                {
                    _AddValueToPlantForT1Map = Convert.ToInt32( db.Settings.Where(st => st.SettingCode == "AddValueToPlantForT1Map").FirstOrDefault().Vals);
                }
                return _AddValueToPlantForT1Map;

            }
        }
        string _ElectronicTimesheetFolderPath = "empty";
        public string ElectronicTimesheetFolderPath
        {
            get
            {
                if (_ElectronicTimesheetFolderPath == "empty")
                {
                    _ElectronicTimesheetFolderPath = db.Settings.Where(st => st.SettingCode == "ElectronicTimesheetFolderPath").FirstOrDefault().Vals;
                }
                return _ElectronicTimesheetFolderPath;
            }
            set
            {
                var sets = db.Settings.Where(st => st.SettingCode == "ElectronicTimesheetFolderPath").FirstOrDefault();
                sets.Vals = value;
                db.SubmitChanges();
                _ElectronicTimesheetFolderPath = value;
            }
        }

        string _ElectronicTimesheetFile = "empty";
        public string ElectronicTimesheetFile
        {
            get
            {
                if (_ElectronicTimesheetFile == "empty")
                {
                    _ElectronicTimesheetFile = db.Settings.Where(st => st.SettingCode == "ElectronicTimesheetFile").FirstOrDefault().Vals;
                }
                return _ElectronicTimesheetFile;
            }
            set
            {
                var sets = db.Settings.Where(st => st.SettingCode == "ElectronicTimesheetFile").FirstOrDefault();
                int xx=value.Length;
                sets.Vals = value;
                db.SubmitChanges();
                _ElectronicTimesheetFile = value;
            }
        }


        /// <summary>
        /// paycomponents where total column overrides day sum correction applied to day sum
        /// </summary>
        string _DoNotValidateTotalForPayCompCode = "empty";
        public string CorrectDaySumToTotalColumnForPayCompCode
        {
            get
            {
                if (_DoNotValidateTotalForPayCompCode == "empty")
                {
                    _DoNotValidateTotalForPayCompCode = db.Settings.Where(st => st.SettingCode == "CorrectDaySumToTotalColumnForPayCompCode").FirstOrDefault().Vals;
                }
                return _DoNotValidateTotalForPayCompCode;
            }
        }

        /// <summary>
        /// paycomponents where total column overrides day sum correction applied to day sum
        /// </summary>
        /// <param name="Paycomp"></param>
        /// <returns></returns>
        bool CorrectDaySumToTotalColumnForPayCompCodeBool(decimal Paycomp)
        {
            string[] codes = CorrectDaySumToTotalColumnForPayCompCode.Split(new char[] { Convert.ToChar(",") });
            foreach (var cd in codes)
            {
                if (Convert.ToDecimal(cd) == Paycomp)
                    return true;
            }
            return false;
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            tsbNormal_Event.SelectedIndex = 0;
            FillPayPeriodCombo();
            FillEmployeeCombo();
            ConfigureMappingControls();
            toolStripProgressBar1.Width = this.Width;
        }

        private void ConfigureMappingControls()
        {
            mpcs = new MapCtrl[] { new MapCtrl(cboA, 0), new MapCtrl(cboB, 1), new MapCtrl(cboC, 2), new MapCtrl(cboD, 3), new MapCtrl(cboE, 4), new MapCtrl(cboF, 5), new MapCtrl(cboGG, 6), new MapCtrl(cboHH, 7), new MapCtrl(cboI, 8), new MapCtrl(cboJJ, 9), new MapCtrl(cboK, 10), new MapCtrl(cboL, 11), new MapCtrl(cboM, 12), new MapCtrl(cboN, 13), new MapCtrl(cboO, 14), new MapCtrl(cboP, 15), new MapCtrl(cboQ, 16), new MapCtrl(cboR, 17), new MapCtrl(cboS, 18), new MapCtrl(cboT, 19), new MapCtrl(cboU, 20), new MapCtrl(cboV, 21), new MapCtrl(cboW, 22), new MapCtrl(cboX, 23), new MapCtrl(cboY, 24), new MapCtrl(cboZ, 25) };
        }

        private void FillPayPeriodCombo()
        {
            foreach (var item in db.PayYears.OrderByDescending(i=>i.PayNoYear).ToList())
            {
                cboPayPeriod.Items.Add(item.PayNoYear + "-" + item.Comment);
            }
        }

        private void LblExcelFile_MouseLeave(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
        }

        private void LblExcelFile_MouseHover(object sender, EventArgs e)
        {
            Cursor = Cursors.Hand;
        }

        private void DataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            ;
//            int cnt=excelImportDS1.Excel.Rows.Count;
          //   MessageBox.Show("DataGridView1_DataError:" + e.Exception.Message);
        }

        bool All;

        private void tsbReport_Click(object sender, EventArgs e)
        {
        }

        private void tsbT1Export_Click(object sender, EventArgs e)
        {
            (new frmT1ImpSummary(PayNo, EmpNo,All)).ShowDialog();
        }

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
//            public string T1Job { get; set; }
        }

        void ExportCSV(ExcelImportDS T1csv)
        {
            StringBuilder sb = new StringBuilder();
            int cnt = 0;
            foreach (ExcelImportDS.ExcelRow row in T1csv.Excel)//.Where(i => i.Status.ToLower().Contains("suc")))
            {
                try
                {
                    string line = "";
                    for (int i = 0; i < 10; i++)
                    {
                            string val = "";

                            if (!row.IsNull(i))
                                val = row[i].ToString();

                            val = val.Replace("\r", "");
                            val = val.Replace("\n", "");
                            val = "\"" + val + "\"";
                            if (line == "")
                                line = val;

                            else
                                line = line + "," + val;

                    }
                    sb.AppendLine(line);
                }
                catch (Exception ex)
                {
                    ex = ex;
                }
                cnt++;
            }
            saveFileDialog1.FileName = "T1_PayTestLoad.csv";
            saveFileDialog1.ShowDialog();
            string fn;
            if (saveFileDialog1.FileName.Contains(".csv"))
                fn = saveFileDialog1.FileName;
            else
                fn = saveFileDialog1.FileName + ".csv";
            System.IO.File.WriteAllText(fn, sb.ToString());
            System.Diagnostics.Process.Start(fn);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }
        bool SetPeriod;
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
        }

        private void tsbAutomapColumns_Click(object sender, EventArgs e)
        {
            //ClearTS(ts.TimesheetID);

            //AutoMap("");
        }

        private void tsbMigrateToImportTable_Click(object sender, EventArgs e)
        {
            //ClearTS(ts.TimesheetID);

            //Migrate(false);
        }

        private void cboY_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
        }


        private void tsbEmpAllow_Click(object sender, EventArgs e)
        {
        }

        private void tsTitle_Click(object sender, EventArgs e)
        {

        }

        private void employeesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            (new frmEmp()).ShowDialog();

        }



        private void toolStripButton4_Click(object sender, EventArgs e)
        {

        }

        private void classToolStripMenuItem_Click(object sender, EventArgs e)
        {
            (new frmClass()).ShowDialog();
        }

        private void employeeAllowancesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            (new frmEmpAllow()).ShowDialog();
        }

        private void timesheetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PromptPayPeriod();
            Cursor = Cursors.WaitCursor;
            (new frmReport(PayNo, CboEmpNo(), All)).ShowDialog();
            Cursor = Cursors.Default;
        }


        /// <summary>
        /// get employee selected into combo box / dropdown 
        /// </summary>
        /// <returns></returns>
        int CboEmpNo()
        {
            if (tsbCboEmp.SelectedItem==null)
                return 0;

            string item = tsbCboEmp.SelectedItem.ToString();
            if (item == "-All-")
                return 0;

            char cc = Convert.ToChar("-");
            string[] ss = item.ToString().Split(new char[] { cc });
            if (ss.Count() > 0)
            {
                return Convert.ToInt32(ss[0].ToString().Replace(" ", ""));
            }
            return 0;
        }

        private void t1ImportWorkLeaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PromptPayPeriod();
            Cursor = Cursors.WaitCursor;
            (new frmReportT1PayImport(PayNo, CboEmpNo(), All)).ShowDialog();
            Cursor = Cursors.Default;
        }

        private void t1ImportPlantToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PromptPayPeriod();

            (new frmPlant(PayNo)).ShowDialog();

        }

        private void errorReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PromptPayPeriod();

            PayPeriodExceptions();

        }

        private void jobCodesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            (new frmJobCode()).ShowDialog();

        }

        private void payComponentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            (new frmPayComp()).ShowDialog();

        }

        private void lblExcelFile_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(lblExcelFile.Text);
        }

        private void toolStripDropDownButton3_Click(object sender, EventArgs e)
        {

        }

        private void xLTimesheetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EventGroup = Guid.NewGuid();
            panel1.Visible = true;
            if (PayNo == 0)
            {
                SetPeriod = false;
                string message = "This single timesheet will be imported into test payrun period";
                MessageBox.Show(message);
                EventLog(message,  "", "");
            }
            else
            {
                SetPeriod = true;
            }
            All = false;
            ImportSingleExcelFile();
            panel1.Visible = false;
        }

        private void entirePayPeriodFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            EventGroup = Guid.NewGuid();
            panel1.Visible = true;
            SetPeriod = false;
            All = true;
            // ths = this;
            toolStripStatusLabel1.Text = "Progress";
            tsbCboEmp.SelectedIndex = 0;
            PromptPayPeriod();
            ImportTBExcelFolder();
            panel1.Visible = false;
         //   ths = null;
        }

       // static Form1 ths;
        private void CleatAllTimedataForPeriod()
        {
            var timsheets = db.Timesheets.Where(ts => ts.PayNoYear == PayNo).ToList();
            foreach (var item in timsheets)
            {
                ClearTimesheetInDB(item.TimesheetID, item.StaffID);
            }
        }

        private void tsbPlant_Click(object sender, EventArgs e)
        {
            (new frmPlantList()).ShowDialog();

        }

        private void payPeriodToolStripMenuItem_Click(object sender, EventArgs e)
        {
            (new frmPeriod()).ShowDialog();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            (new frmSettings()).ShowDialog();
        }

        private void CboPayPeriod_SelectedIndexChanged(object sender, System.EventArgs e)

        {
//            {
  //              get
    //        {
                    if (cboPayPeriod.SelectedItem == null)
          //              return -1;
          PayNo=-1;
                    string ss = cboPayPeriod.SelectedItem.ToString();
                    string[] sss = ss.Split(new string[] { "-" }, StringSplitOptions.RemoveEmptyEntries);
            //     return Convert.ToInt32(sss[0]);
            PayNo = Convert.ToInt32(sss[0]);
      //          }
      //    }

        }

        private void LedgerJobCodeMapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PromptPayPeriod();

            (new frmLedgerJobNoMapping(PayNo)).ShowDialog();
        }

        private void tsbEventLog_Click(object sender, EventArgs e)
        {
            (new frmEvents()).ShowDialog();
        }

        private void tsbNormal_Event_Click(object sender, EventArgs e)
        {
            var del = db.EventLogs.Where(el => el.EventDT < DateTime.Now.AddDays(-2));
            db.EventLogs.DeleteAllOnSubmit(del);
            db.SubmitChanges();
        }
        private void TsbNormal_Event_TextChanged(object sender, System.EventArgs e)
        {
            if (tsbNormal_Event.SelectedIndex == 0)
                tsbEventLog.Visible = false;
            else
                tsbEventLog.Visible = true;

        }

        private void tsbHelp_Click(object sender, EventArgs e)
        {
            (new frmHelp()).ShowDialog();
        }

      
    }
}

