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
    public partial class frmClass : Form
    {
        public frmClass()
        {
            InitializeComponent();
        }

        DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

        private void frmClass_Load(object sender, EventArgs e)
        {
            this.Text = "Class";
            WindowState = FormWindowState.Maximized;
            FillNavGrid();
        }

        private void FillNavGrid()
        {
//            DataClasses1DataContext dbs = new DataClasses1DataContext(Form1.ConString);

            var emp = db.Classes.ToList();

            BindingSource bs = new BindingSource();
            bs.DataSource = emp;
            dataGridView1.DataSource = bs;
            bindingNavigator1.BindingSource = bs;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            frmImpClass fie=new frmImpClass();
            fie.ShowDialog();
            FillNavGrid();
            if (ValidateData())
                MessageBox.Show("Issues found - you may need to export the Pay Components file from Techone again and reimport");
        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            FillNavGrid();

        }
        bool ValidateData()
        {
          
            List<string> vals = new List<string>();

//            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
            {
                var query =
            (from c in db.TimesheetDatas
             where !(from o in db.Classes
                     select o.PCSClassNo)
                    .Contains(((c.ClassNo == null) ? 0 : c.ClassNo.Value))
             select c.ClassNo).ToList();

                string ss;
                if (query.Count() == 0)
                    ;// MessageBox.Show("Paycomponent code looks ok");
                else
                {
//                    Issues = true;
                    ss = "The following class codes are in Employee Timesheet but not in the class file: ";
                    foreach (var item in query)
                    {
                        ss = ss + item.ToString() + ", ";
                    }
                    vals.Add(ss);
                }
            }
            {
                var query =
        (from c in db.Classes
         where c.HoursPerFN < 20 || c.HoursPerFN > 80
         select c.PCSClassNo).ToList();

                string ss;
                if (query.Count() == 0)
                    ;// MessageBox.Show("Paycomponent code looks ok");
                else
                {
//                    Issues = true;
                    ss = "The following Classes have hours per fortnight outside 20-80: ";
                    foreach (var item in query)
                    {
                        ss = ss + item.ToString() + ", ";
                    }
                    //                  MessageBox.Show(ss);
                    vals.Add(ss);
                }
            }

            if (vals.Count > 0)
            {
                (new frmValidation(vals)).ShowDialog();
                return true;
            }
            MessageBox.Show("No validation issues found with class data");
            return false;
        }
        private void tsbValidate_Click(object sender, EventArgs e)
        {
            ValidateData();
        }
    }
}
