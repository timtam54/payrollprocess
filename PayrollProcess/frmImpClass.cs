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
    public partial class frmImpClass : Form
    {
        public frmImpClass()
        {
            InitializeComponent();
        }

        private void frmImpClass_Load(object sender, EventArgs e)
        {
            mpcs = new MapCtrl[] { new MapCtrl(cboA, 0), new MapCtrl(cboB, 1), new MapCtrl(cboC, 2), new MapCtrl(cboD, 3), new MapCtrl(cboE, 4), new MapCtrl(cboF, 5), new MapCtrl(cboGG, 6), new MapCtrl(cboHH, 7), new MapCtrl(cboI, 8), new MapCtrl(cboJJ, 9), new MapCtrl(cboK, 10), new MapCtrl(cboL, 11), new MapCtrl(cboM, 12), new MapCtrl(cboN, 13), new MapCtrl(cboO, 14), new MapCtrl(cboP, 15), new MapCtrl(cboQ, 16), new MapCtrl(cboR, 17), new MapCtrl(cboS, 18), new MapCtrl(cboT, 19), new MapCtrl(cboU, 20), new MapCtrl(cboV, 21), new MapCtrl(cboW, 22), new MapCtrl(cboX, 23), new MapCtrl(cboY, 24), new MapCtrl(cboZ, 25) };
            this.WindowState = FormWindowState.Maximized;
            Import();
        }

        public void Import()
        { 
            ImportTBExcel();
            //InitHI();
        }

        HeaderIndex[] HI;
        MapCtrl[] mpcs;
        HeaderIndex HIClassNo;
        HeaderIndex HIEmpCondition;
        HeaderIndex HIGradeCodeID;
        HeaderIndex HILevelCodeID;
        HeaderIndex HIHoursPerFN;
        HeaderIndex HIAmountPerFN;
        HeaderIndex HIPYCLASS_CLASSCODE;
        HeaderIndex HIPYCLASS_CLASSDESC;

        private void InitHI()
        {
            try
            {
                HI = new HeaderIndex[8];

                HIClassNo = new HeaderIndex("PCS_Classification_ No");
                HI[0] = HIClassNo;

                HIEmpCondition = new HeaderIndex("Emp_Condition");
                HI[1] = HIEmpCondition;

                HIGradeCodeID = new HeaderIndex("Grade Code ID");
                HI[2] = HIGradeCodeID;

                HILevelCodeID = new HeaderIndex("Level Code ID");
                HI[3] = HILevelCodeID;

                HIHoursPerFN = new HeaderIndex("Hours per fortnight");
                HI[4] = HIHoursPerFN;

                HIAmountPerFN = new HeaderIndex("Amount per fortnight");
                HI[5] = HIAmountPerFN;

                HIPYCLASS_CLASSCODE = new HeaderIndex("PYCLASS_CLASSCODE");
                HI[6] = HIPYCLASS_CLASSCODE;

                HIPYCLASS_CLASSDESC = new HeaderIndex("PYCLASS_CLASSDESC");
                HI[7] = HIPYCLASS_CLASSDESC;

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

            //           ClearEA();
            InitHI();
            ResizeGrid();
            frmImpEmpAllow.DeleteRowsAboveHeader(excelImportDS1);
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
                    Class EA;

                    object ClassNo = exrow[HIClassNo.SelectedColumn];
                    if (!ClassNo.Equals(""))
                    {
                        int ClassNum = Convert.ToInt32(ClassNo.ToString());
                        EA = db.Classes.Where(ea => ea.PCSClassNo == ClassNum).FirstOrDefault();
                        if (EA != null)
                            MessageBox.Show("Duplicate Class ClassNo:" + ClassNum.ToString());
                        else
                        {
                            EA = new Class();

                            EA.PCSClassNo = ClassNum;

                            EA.Emp_Condition = exrow[HIEmpCondition.SelectedColumn].ToString();
                            EA.GradeCodeID = exrow[HIGradeCodeID.SelectedColumn].ToString();
                            EA.LevelCodeID = exrow[HILevelCodeID.SelectedColumn].ToString();
                            EA.HoursPerFN = Convert.ToDecimal(exrow[HIHoursPerFN.SelectedColumn]);
                            EA.AmountPerFN = Convert.ToDecimal(exrow[HIAmountPerFN.SelectedColumn]);
                            EA.PYCLASS_CLASSCODE = Convert.ToInt32(exrow[HIPYCLASS_CLASSCODE.SelectedColumn]);
                            EA.PYCLASS_CLASSDESC = Convert.ToString(exrow[HIPYCLASS_CLASSDESC.SelectedColumn]);


                            db.Classes.InsertOnSubmit(EA);
                            db.SubmitChanges();
                        }


                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            {
                Class EA = new Class();
                EA.PCSClassNo = 0;
                db.Classes.InsertOnSubmit(EA);
                db.SubmitChanges();

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
            var delts = db.Classes;
            db.Classes.DeleteAllOnSubmit(delts);
            db.SubmitChanges();
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            RemoveHeader();
            ClearEA();
            Migrate();
            this.Cursor = Cursors.Default;
        }
    }
}
