using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static PayrollProcess.Form1;

namespace PayrollProcess
{
    public partial class frmImpPlant : Form
    {
        public frmImpPlant()
        {
            InitializeComponent();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            RemoveHeader();
            ClearEA();
            Migrate();
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
            var delts = db.Plants;
            db.Plants.DeleteAllOnSubmit(delts);
            db.SubmitChanges();
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
                    Plant EA;

                    string WO = exrow[HiPlantNo.SelectedColumn].ToString();
                    if (!WO.Equals(""))
                    {
                        // int WONum = Convert.ToInt32(WO.ToString());
                       EA = db.Plants.Where(ea => ea.PlantSource == WO).FirstOrDefault();
                        if (EA != null)
                            MessageBox.Show("Duplicate plant PlsntNo:" + WO.ToString());
                        else
                        {
                            EA = new Plant();

                            EA.PlantSource = WO;
                            EA.PlantTarget =Convert.ToInt32( WO);
                            EA.PlantDesc = exrow[HIDesc.SelectedColumn].ToString();



                            db.Plants.InsertOnSubmit(EA);
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
                Plant EA = new Plant();
                EA.PlantSource = "0";
                EA.PlantTarget = 0;
                EA.PlantDesc = "";
                db.Plants.InsertOnSubmit(EA);
                db.SubmitChanges();

            }

            MessageBox.Show("Complete");
            this.Controls.Remove(dataGridView1);
            Close();
        }
        private void FrmImpPlant_Load(object sender, System.EventArgs e)
        {
            mpcs = new MapCtrl[] { new MapCtrl(cboA, 0), new MapCtrl(cboB, 1), new MapCtrl(cboC, 2), new MapCtrl(cboD, 3), new MapCtrl(cboE, 4), new MapCtrl(cboF, 5), new MapCtrl(cboGG, 6), new MapCtrl(cboHH, 7), new MapCtrl(cboI, 8), new MapCtrl(cboJJ, 9), new MapCtrl(cboK, 10), new MapCtrl(cboL, 11), new MapCtrl(cboM, 12), new MapCtrl(cboN, 13), new MapCtrl(cboO, 14), new MapCtrl(cboP, 15), new MapCtrl(cboQ, 16), new MapCtrl(cboR, 17), new MapCtrl(cboS, 18), new MapCtrl(cboT, 19), new MapCtrl(cboU, 20), new MapCtrl(cboV, 21), new MapCtrl(cboW, 22), new MapCtrl(cboX, 23), new MapCtrl(cboY, 24), new MapCtrl(cboZ, 25) };
            this.WindowState = FormWindowState.Maximized;
            Import();
        }
        HeaderIndex[] HI;
        MapCtrl[] mpcs;
        HeaderIndex HiPlantNo;
        HeaderIndex HIDesc;

        public void Import()
        {
            ImportTBExcel();
            //InitHI();
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

        private void InitHI()
        {
            try
            {
                HI = new HeaderIndex[2];

                HiPlantNo = new HeaderIndex("Asset");
                HI[0] = HiPlantNo;

                HIDesc = new HeaderIndex("Description");
                HI[1] = HIDesc;

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

    }
}
