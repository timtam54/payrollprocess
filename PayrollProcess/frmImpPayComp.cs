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
    public partial class frmImpPayComp : Form
    {
        public frmImpPayComp()
        {
            InitializeComponent();
            /* todotim
             * You're correct - Units column looks incorrect in the current Paycomponents.
Hopefully re-importing last week will fix
*/
        }

        private void frmImpPayComp_Load(object sender, EventArgs e)
        {

            mpcs = new MapCtrl[] { new MapCtrl(cboA, 0), new MapCtrl(cboB, 1), new MapCtrl(cboC, 2), new MapCtrl(cboD, 3), new MapCtrl(cboE, 4), new MapCtrl(cboF, 5), new MapCtrl(cboGG, 6), new MapCtrl(cboHH, 7), new MapCtrl(cboI, 8), new MapCtrl(cboJJ, 9), new MapCtrl(cboK, 10), new MapCtrl(cboL, 11), new MapCtrl(cboM, 12), new MapCtrl(cboN, 13), new MapCtrl(cboO, 14), new MapCtrl(cboP, 15), new MapCtrl(cboQ, 16), new MapCtrl(cboR, 17), new MapCtrl(cboS, 18), new MapCtrl(cboT, 19), new MapCtrl(cboU, 20), new MapCtrl(cboV, 21), new MapCtrl(cboW, 22), new MapCtrl(cboX, 23), new MapCtrl(cboY, 24), new MapCtrl(cboZ, 25) };

            this.WindowState = FormWindowState.Maximized;
            Import();


        }

         void Import()
        {
            ImportTBExcel();
            //InitHI();
        }

        HeaderIndex[] HeaderFieldMaps;
        MapCtrl[] mpcs;
        HeaderIndex HIPayCompCode;
        HeaderIndex HIPayCompDesc;
        HeaderIndex HIType;
        HeaderIndex HITypeDesc;
        HeaderIndex HIPayDed;
        HeaderIndex HIUnits;
        HeaderIndex HIPPUnits;
        private void InitHI()
        {
            try
            {
                HeaderFieldMaps = new HeaderIndex[7];
                HIPayCompCode = new HeaderIndex("PayCompCode", "Pay Component Code");
                HeaderFieldMaps[0] = HIPayCompCode;
                HIPayCompDesc = new HeaderIndex("Pay Component Description");
                HeaderFieldMaps[1] = HIPayCompDesc;

                HIType = new HeaderIndex("Pay Component Type");
                HeaderFieldMaps[2] = HIType;

                HITypeDesc = new HeaderIndex("Pay Component Type Description");
                HeaderFieldMaps[3] = HITypeDesc;

                HIPayDed = new HeaderIndex("Payment / Deduction");
                HeaderFieldMaps[4] = HIPayDed;

                HIUnits = new HeaderIndex("Units");
                HeaderFieldMaps[5] = HIUnits;
                HIPPUnits = new HeaderIndex("Pay Period Unit");
                HeaderFieldMaps[6] = HIPPUnits;

                int i = 0;

                foreach (HeaderIndex item in HeaderFieldMaps)
                {
                    item.Index = i;
                    i++;
                }

                foreach (var mpcsitem in mpcs)
                {
                    mpcsitem.ctrl.Items.Clear();
                    foreach (HeaderIndex Hiitem in HeaderFieldMaps)
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
                MessageBox.Show("Please select the Excel T1 Pay Component file that you wish to import");
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

            AutoMap();
            
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
                    foreach (HeaderIndex item in HeaderFieldMaps)
                    {
                        if (mc.ctrl.SelectedItem.ToString() == item.Header)
                            item.SelectedColumn = mc.ColIndex;
                    }
                    if (mc.ctrl.SelectedIndex == -1)
                    {
                        MessageBox.Show("");
                    }
                }
            }

            foreach (ExcelImportDS.ExcelRow exrow in excelImportDS1.Excel)
            {
                try
                {
                    PayComponent EA;
                    object code = exrow[HIPayCompCode.SelectedColumn];
                    if (!code.Equals(""))
                    {
                        decimal PayCompCode = Convert.ToDecimal(code);
                        EA = db.PayComponents.Where(ea => ea.PayCompCode == PayCompCode).FirstOrDefault();
                        if (EA != null)
                            MessageBox.Show("Duplicate Class PayCompCode:" + PayCompCode.ToString());
                        else
                        {

                            EA = new PayComponent();
                            EA.PayCompCode = PayCompCode;
                            EA.PayCompDesc = exrow[HIPayCompDesc.SelectedColumn].ToString();
                            EA.PayCompType = Convert.ToInt32(exrow[HIType.SelectedColumn]);
                            EA.PayCompTypeDesc = Convert.ToString(exrow[HITypeDesc.SelectedColumn]);
                            EA.Payment_deduct = Convert.ToString(exrow[HIPayDed.SelectedColumn]);
                            decimal Units = Convert.ToDecimal(exrow[HIUnits.SelectedColumn]);
                            EA.Units = Convert.ToInt32(Units);
                            EA.PayPeriodUnit = Convert.ToString(exrow[HIPPUnits.SelectedColumn]);

                            db.PayComponents.InsertOnSubmit(EA);
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
                PayComponent EA = new PayComponent();
                EA.PayCompCode = 0;
                db.PayComponents.InsertOnSubmit(EA);
                db.SubmitChanges();
            }
            MessageBox.Show("Complete");
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

                    foreach (HeaderIndex item in HeaderFieldMaps)
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
            foreach (HeaderIndex Hi in HeaderFieldMaps)
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
            var delts = db.PayComponents;
            db.PayComponents.DeleteAllOnSubmit(delts);
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
