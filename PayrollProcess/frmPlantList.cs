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
    public partial class frmPlantList : Form
    {
        public frmPlantList()
        {
            InitializeComponent();
        }

        private void frmPlantList_Load(object sender, EventArgs e)
        {

            WindowState = FormWindowState.Maximized;
            FillNavGrid();
        }

        private void FillNavGrid()
        {
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

            var emp = db.Plants.ToList();

            BindingSource bs = new BindingSource();
            bs.DataSource = emp;
            dataGridView1.DataSource = bs;
            bindingNavigator1.BindingSource = bs;

            dataGridView1.Columns[0].HeaderText = "Asset";
            dataGridView1.Columns[1].HeaderText = "Asset";
            dataGridView1.Columns[2].HeaderText = "Description";
        }

        private void tbimport_Click(object sender, EventArgs e)
        {
            frmImpPlant fie = new frmImpPlant();
            fie.ShowDialog();
            FillNavGrid();
            if (ValidateData())
                MessageBox.Show("Issues found - you may need to export the Pay Components file from Techone again and reimport");
        }
        DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);


        bool ValidateData()
        {
            List<string> vals = new List<string>();
            {
                int TimesheetID = db.Timesheets.OrderByDescending(i => i.PayNoYear).Select(i => i.TimesheetID).FirstOrDefault();

                var query =
            (from c in db.TimesheetDatas
             where c.TimesheetID==TimesheetID && c.PlantNo != null && !(from o in db.Plants
                                      select o.PlantSource).Contains(c.PlantNo.ToString())
             select c.PlantNo).ToList();

                string ss;
                if (query.Count() == 0)
                    ;// MessageBox.Show("Paycomponent code looks ok");
                else
                {
                    int cnt = 0;

                    ss = "The following plant codes are in Employee Timesheet but not in the plant list file: ";

                    foreach (var item in query)
                    {
                        if (item != null)
                        {
                            ss = ss + item.ToString() + ", ";
                            cnt++;
                        }
                    }
                    if (cnt>0)
                    vals.Add(ss);
                }
            }
        //    {
        //        var query =
        //(from c in db.Classes
        // where c.HoursPerFN < 20 || c.HoursPerFN > 80
        // select c.PCSClassNo).ToList();

        //        string ss;
        //        if (query.Count() == 0)
        //            ;// MessageBox.Show("Paycomponent code looks ok");
        //        else
        //        {
        //            //                    Issues = true;
        //            ss = "The following plant Nos have hours per fortnight outside 20-80: ";
        //            foreach (var item in query)
        //            {
        //                ss = ss + item.ToString() + ", ";
        //            }
        //            //                  MessageBox.Show(ss);
        //            vals.Add(ss);
        //        }
        //    }

            if (vals.Count > 0)
            {
                (new frmValidation(vals)).ShowDialog();
                return true;
            }
            MessageBox.Show("No validation issues found with class data");
            return false;
        }


        private void toolStripSeparator1_Click(object sender, EventArgs e)
        {

        }

        private void tsbValidate_Click(object sender, EventArgs e)
        {

            ValidateData();
        }
    }
}
