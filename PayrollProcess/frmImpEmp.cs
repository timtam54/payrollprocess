using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static PayrollProcess.Form1;

namespace PayrollProcess
{
    public partial class frmImpEmp : Form
    {
        public frmImpEmp()
        {
            InitializeComponent();
        }

        private void frmImpEmp_Load(object sender, EventArgs e)
        {
            mpcs = new MapCtrl[] { new MapCtrl(cboA, 0), new MapCtrl(cboB, 1), new MapCtrl(cboC, 2), new MapCtrl(cboD, 3), new MapCtrl(cboE, 4), new MapCtrl(cboF, 5), new MapCtrl(cboGG, 6), new MapCtrl(cboHH, 7), new MapCtrl(cboI, 8), new MapCtrl(cboJJ, 9), new MapCtrl(cboK, 10), new MapCtrl(cboL, 11), new MapCtrl(cboM, 12), new MapCtrl(cboN, 13), new MapCtrl(cboO, 14), new MapCtrl(cboP, 15), new MapCtrl(cboQ, 16), new MapCtrl(cboR, 17), new MapCtrl(cboS, 18), new MapCtrl(cboT, 19), new MapCtrl(cboU, 20), new MapCtrl(cboV, 21), new MapCtrl(cboW, 22), new MapCtrl(cboX, 23), new MapCtrl(cboY, 24), new MapCtrl(cboZ, 25) };
           this.WindowState= FormWindowState.Maximized;
            Import();
        }

        void Import()
        {
            ImportTBExcel();
//            InitHI();

        }

        HeaderIndex[] HI;
        MapCtrl[] mpcs;
        HeaderIndex HIEmpNo;
        HeaderIndex HISurname;
        HeaderIndex HIFirstName;
        //HeaderIndex HIDept;
        HeaderIndex HIEmpType;
        HeaderIndex HIHours;
        private void InitHI()
        {
            try
            {
                HI = new HeaderIndex[5];
                HIEmpNo = new HeaderIndex("Employee ID", "EMPID", "id Number");
                HI[0] = HIEmpNo;
                HISurname = new HeaderIndex("Family Name", "SURNAME");
                HI[1] = HISurname;
                HIFirstName = new HeaderIndex("Given Name", "FIRST_NAME");
                HI[2] = HIFirstName;
//                HIDept = new HeaderIndex("DEPT");
  //              HI[3] = HIDept;
                HIEmpType = new HeaderIndex("EMP_TYPE", "Employee Type");
                HI[3] = HIEmpType;


                HIHours = new HeaderIndex("WEEKLY_HOURS_BASE","HOURS","Units");//, "NORMAL_WEEKLY_WORK_HOURS");
                HI[4] = HIHours;


                int i = 0;

                foreach (HeaderIndex item in HI)
                {
                    item.Index = i;
                    i++;
                }

                foreach (var mpcsitem in mpcs)
                {
                    mpcsitem.ctrl.Items.Clear();
                    foreach (HeaderIndex Hiitem in HI)
                    {
                        mpcsitem.ctrl.Items.Add(Hiitem.Header);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void ImportTBExcel()
        {
            try
            {
                MessageBox.Show("Please select the Excel T1 Employee file that you wish to import");
                openFileDialog1.Filter = "XLSX Files(*.xlsx)|*.xlsx";//|Excel Files(.xlsx)|*.xlsx|Excel Files(.xls)|*.xls| Excel Files(*.xlsm)|*.xlsm
                openFileDialog1.ShowDialog();
                if (openFileDialog1.FileName == "")
                {
                    MessageBox.Show("No file was selected.. action has been cancelled.");
                    return;
                }
                ImportXL(openFileDialog1.FileName);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void ImportXL(string filename)
        {
            excelImportDS1.Excel.Clear();
            excelImportDS1.AcceptChanges();

            XLWorkbook connection = new XLWorkbook(filename);
            IXLWorksheets dt = connection.Worksheets;//.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            frmExcel excSheet = new frmExcel(dt, frmExcel.Plant_TS123.NA);
            excSheet.ShowDialog();
            tsbTabName.Text = excSheet.SheetName;

            IXLWorksheet dr = connection.Worksheet(excSheet.SheetName);

            var rrows = dr.RangeUsed().RowsUsed();//.Skip(1);
            foreach (var rrow in rrows)
            {

                bool Populated = false;
                ExcelImportDS.ExcelRow exrow = excelImportDS1.Excel.NewExcelRow();
                int cc = 0;
                foreach (var ccell in rrow.Cells())
                {
                    try
                    {
                        object oo = "";
                        try
                        {
                            oo = ccell.Value;
                        }
                        catch (Exception rex)
                        {
                            try
                            {
                                oo = ccell.CachedValue;
                            }
                            catch (Exception ex)
                            {
                                object ll = ex;
                            }
                        }
                        exrow[cc] = oo;
                        if (!oo.Equals(""))
                            Populated = true;
                    }
                    catch (Exception ex)
                    {
                        ex = ex;
                    }
                    cc++;
                }
                if (Populated)
                    excelImportDS1.Excel.AddExcelRow(exrow);
            }
            //string excelConnectionString;
            //if (System.IO.Path.GetExtension(filename).ToLower().Equals(".xls"))
            //    excelConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=""" + filename + @""";Extended Properties=""Excel 8.0;HDR=NO;IMEX=1;""";
            //else
            //    excelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""" + filename + @""";Extended Properties=""Excel 12.0;HDR=NO;IMEX=1;""";
            //using (OleDbConnection connection = new OleDbConnection(excelConnectionString))
            //{
            //    try
            //    {
            //        connection.Open();
            //    }
            //    catch (Exception ex)
            //    {
            //        throw ex;
            //    }
            //    DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            //    frmExcel excSheet = new frmExcel(dt, frmExcel.Plant_TS123.NA);
            //    excSheet.ShowDialog();
            //    tsbTabName.Text = excSheet.SheetName;
            //    OleDbCommand command = new OleDbCommand("Select * FROM [" + excSheet.SheetName + "]", connection);
            //    int row = 0;
            //    using (OleDbDataReader dr = command.ExecuteReader())
            //    {
            //        while (dr.Read())
            //        {
            //            row++;
            //            bool Populated = false;
            //            ExcelImportDS.ExcelRow exrow = excelImportDS1.Excel.NewExcelRow();

            //            for (int i = 0; i < dr.FieldCount; i++)
            //            {
            //                try
            //                {
            //                    object oo = dr[i].ToString();
            //                    exrow[i] = oo;
            //                    if (oo != "")
            //                        Populated = true;
            //                }
            //                catch (Exception ex)
            //                {
            //                    ex = ex;
            //                }
            //            }
            //            if (Populated)
            //                excelImportDS1.Excel.AddExcelRow(exrow);
            //        }
            //    }
            //}
            InitHI();
            ResizeGrid();
            frmImpEmpAllow.DeleteRowsAboveHeader(excelImportDS1);

            ExcelImportDS.ExcelRow exrowdel = excelImportDS1.Excel[0];
            while ((exrowdel[2].ToString() == ""))
            {
                exrowdel.Delete();
                excelImportDS1.AcceptChanges();
                exrowdel = excelImportDS1.Excel[0];
            }
            //exrowdel.Delete();
            //excelImportDS1.AcceptChanges();



            AutoMap();
            //if (MessageBox.Show("Please confirm headers are mapped correctly?", "Mapped", MessageBoxButtons.YesNo) == DialogResult.Yes)
            //{
            //}
        }

        void ResizeGrid()
        {
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

        void Migrate()
        {
            foreach (MapCtrl mc in mpcs)
            {
                if (mc.ctrl.SelectedIndex != -1)
                {
                    foreach (HeaderIndex item in HI)
                    {
                        if (mc.ctrl.SelectedItem.ToString() == item.Header)
                            item.SelectedColumn = mc.ColIndex;
                    }
                }
            }
            foreach (ExcelImportDS.ExcelRow exrow in excelImportDS1.Excel)
            {
                try
                {

                    Employee EA;
                    object empno = exrow[HIEmpNo.SelectedColumn];
                    if (!empno.Equals(""))
                    {
                        int T1EmpNo = Convert.ToInt32(empno.ToString().ToLower().Replace("t1_", ""));
                        EA = db.Employees.Where(ea => ea.T1EmpNo == T1EmpNo).FirstOrDefault();
                        if (EA != null)
                            MessageBox.Show("Duplicate Class EmpNo:" + T1EmpNo.ToString());
                        else
                        {
                            EA = new Employee();
                            EA.T1EmpNo = T1EmpNo;
                            EA.FirstName = exrow[HIFirstName.SelectedColumn].ToString();
                            EA.Surname = exrow[HISurname.SelectedColumn].ToString();
                            EA.Emp_Type = Convert.ToChar(exrow[HIEmpType.SelectedColumn].ToString().Substring(0, 1));
                            if (HIHours.SelectedColumn != -1)
                                EA.Hours = Convert.ToDecimal(exrow[HIHours.SelectedColumn]);
                            else
                                EA.Hours = 0;

                            db.Employees.InsertOnSubmit(EA);
                            db.SubmitChanges();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            MessageBox.Show("Complete");
            this.Controls.Remove(dataGridView1);
            Close();
        }
        void AutoMap()
        {
            List<int> selected = new List<int>();

            panel1.Visible = true;

            ExcelImportDS.ExcelRow exrow = excelImportDS1.Excel[0];

            for (int i = 0; i < excelImportDS1.Excel.Columns.Count; i++)
            {
                try
                {
                    string colname = exrow[i].ToString();
                    colname = colname.Replace("\n", " ");
                    int selectedindex = -1;

                    foreach (HeaderIndex item in HI)
                    {
                        string g;
                        string Match;
                        Match = colname.ToLower();
                        if (Match != "")
                        {
                            if (Match == item.Header.ToLower())
                            {

                                item.SelectedColumn = i;
                                selectedindex = item.Index;
                                break;
                            }
                            else if (Match == item.Header2.ToLower())
                            {
                                if (item.Header2 != "")
                                {
                                    item.SelectedColumn = i;
                                    selectedindex = item.Index;
                                    break;
                                }
                            }
                            else if (Match == item.Header3.ToLower())
                            {

                                item.SelectedColumn = i;
                                selectedindex = item.Index;
                                break;
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
                                        //if (mpcsitem.ctrl.SelectedIndex==-1)
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
            foreach (HeaderIndex Hi in HI)
            {
                if (Hi.SelectedColumn == -1)
                {
                    MessageBox.Show(Hi.Header + " not mapped.  Either manually map this column or extract data file from Techone and reimport");
                    return;
                }
            }
        }

        void RemoveHeader()
        {
            ExcelImportDS.ExcelRow delexrow = excelImportDS1.Excel[0];
            delexrow.Delete();
            excelImportDS1.AcceptChanges();

        }
        DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
        void ClearEA()
        {
            var delts = db.Employees;
            db.Employees.DeleteAllOnSubmit(delts);
            db.SubmitChanges();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            RemoveHeader();
            ClearEA();

            Migrate();

        }
    }
}
