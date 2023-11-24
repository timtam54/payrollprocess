using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static PayrollProcess.Form1;

namespace PayrollProcess
{
    public partial class frmImpEmpAllow : Form
    {
        public frmImpEmpAllow()
        {
            InitializeComponent();
        }

        private void frmImpEmpAllow_Load(object sender, EventArgs e)
        {
            mpcs = new MapCtrl[] { new MapCtrl(cboA, 0), new MapCtrl(cboB, 1), new MapCtrl(cboC, 2), new MapCtrl(cboD, 3), new MapCtrl(cboE, 4), new MapCtrl(cboF, 5), new MapCtrl(cboGG, 6), new MapCtrl(cboHH, 7), new MapCtrl(cboI, 8), new MapCtrl(cboJJ, 9), new MapCtrl(cboK, 10), new MapCtrl(cboL, 11), new MapCtrl(cboM, 12), new MapCtrl(cboN, 13), new MapCtrl(cboO, 14), new MapCtrl(cboP, 15), new MapCtrl(cboQ, 16), new MapCtrl(cboR, 17), new MapCtrl(cboS, 18), new MapCtrl(cboT, 19), new MapCtrl(cboU, 20), new MapCtrl(cboV, 21), new MapCtrl(cboW, 22), new MapCtrl(cboX, 23), new MapCtrl(cboY, 24), new MapCtrl(cboZ, 25) };
            this.WindowState = FormWindowState.Maximized;
            ImportTBExcel();
          //  InitHI();
        }

        HeaderIndex[] HI;
        MapCtrl[] mpcs;
        HeaderIndex HIEmpNo;
        HeaderIndex HIPayComp;
        HeaderIndex HIUnits;
        private void InitHI()
        {
            try
            {
                HI = new HeaderIndex[3];
                HIEmpNo = new HeaderIndex("EmpNo", "T1_Emp","id Number");
                HI[0] = HIEmpNo;
                HIPayComp = new HeaderIndex("PayComponentCode", "PayCompCode","Pay Component Code");
                HI[1] = HIPayComp;
                HIUnits = new HeaderIndex("Units", "Unit");
                HI[2] = HIUnits;
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
                MessageBox.Show("Please select the Excel T1 'Employee Allowances / Entitlements' file that you wish to import");

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
            //                object oo = dr[i].ToString();
            //                exrow[i] = oo;
            //                if (oo != "")
            //                    Populated = true;

            //            }
            //            if (Populated)
            //                excelImportDS1.Excel.AddExcelRow(exrow);
            //        }
            //    }
            //}

            InitHI();
            ResizeGrid();
            frmImpEmpAllow.DeleteRowsAboveHeader(excelImportDS1);
            AutoMap();
        }

        public static void DeleteRowsAboveHeader(ExcelImportDS excelImportDS1)
        {
            while (true)
            {
                ExcelImportDS.ExcelRow exrow = excelImportDS1.Excel[0];

                if (PopColumnCount(exrow) < 2)
                    exrow.Delete();
                else
                    return;
            }
        }

        static int PopColumnCount(ExcelImportDS.ExcelRow exrow)
        {
            int cnt = 0;
            for (int i = 0; i < 10; i++)
            {
                if (exrow[i].ToString() != "")
                    cnt++;
            }
            return cnt;
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
                    Emp_Allowance EA;
                                        object empno = exrow[HIEmpNo.SelectedColumn];
                    if (!empno.Equals(""))
                    {
                        int EmpNo = Convert.ToInt32(empno);
                        object pcc = exrow[HIPayComp.SelectedColumn];
                        decimal PayComponentCode = Convert.ToDecimal(pcc);
                        EA=db.Emp_Allowances.Where(ea => ea.PayComponentCode == PayComponentCode && ea.T1_EmpID == EmpNo).FirstOrDefault();
                        if (EA != null)
                        {
                            MessageBox.Show("Duplicate Employee Allowances - EmpNo:" + EmpNo.ToString() + ",PayComponentCode:" + PayComponentCode.ToString());
                        }
                        else
                        {
                            EA = new Emp_Allowance();
                            EA.PayComponentCode = PayComponentCode;
                            EA.T1_EmpID = EmpNo;



                            object un = exrow[HIUnits.SelectedColumn];
                            if (un!=DBNull.Value)
                                EA.units = Convert.ToDecimal(un);
                            db.Emp_Allowances.InsertOnSubmit(EA);
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
            Close();
        }
        void AutoMap()
        {
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

                                item.SelectedColumn = i;
                                selectedindex = item.Index;
                                break;
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
                        foreach (var mpcsitem in mpcs)
                        {
                            if (i == mpcsitem.ColIndex)
                            {
                                try
                                {
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
            var delts = db.Emp_Allowances;
            db.Emp_Allowances.DeleteAllOnSubmit(delts);
            db.SubmitChanges();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            RemoveHeader();
            ClearEA();
            Migrate();
            this.Cursor = Cursors.Default;
        }
    }
}