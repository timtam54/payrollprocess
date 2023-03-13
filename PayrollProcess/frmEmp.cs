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
    public partial class frmEmp : Form
    {
        public frmEmp()
        {
            InitializeComponent();
        }
        private void frmEmp_Load(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Maximized;
            FillNavGrid();
        }

        private void FillNavGrid()
        {
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

            var emp = db.Employees.ToList();

            BindingSource bs = new BindingSource();
            bs.DataSource = emp;
            dataGridView1.DataSource = bs;
            bindingNavigator1.BindingSource = bs;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            frmImpEmp fie = new frmImpEmp();
            fie.ShowDialog();
            FillNavGrid();
            try
            {
                if (ValidateData())
                    MessageBox.Show("Issues found - you may need to export the Employees file from Techone again and re-import");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void TsbHalveHours_Click(object sender, EventArgs e)
        {
            decimal Factor = (decimal)0.5;
            FactorHours(Factor);
        }

        private void FactorHours(decimal Factor)
        {
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

            var emp = db.Employees.ToList();
            foreach (var item in emp)
            {
                try
                {
                    if (item.Hours != null)
                    {
                        item.Hours = item.Hours * Factor;
                        db.SubmitChanges();

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            BindingSource bs = new BindingSource();
            bs.DataSource = emp;
            dataGridView1.DataSource = bs;
            bindingNavigator1.BindingSource = bs;
            MessageBox.Show("restart app for new values to take effect");
        }

        private void TsbDoubleHours_Click(object sender, EventArgs e)
        {
            decimal Factor = (decimal)2;
            FactorHours(Factor);

        }

        bool ValidateData()
        {
            //            bool Issues = false;
            List<string> vals = new List<string>();
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
            {
                var query =
(from c in db.Employees
 where c.FirstName=="" || c.Surname==""
 select c.T1EmpNo).ToList();

                string ss;
                if (query.Count() == 0)
                    ;// MessageBox.Show("Paycomponent code looks ok");
                else
                {
//                    Issues = true;
                    ss = "The following EmployeeIds are missing first name or surname in Employees: ";
                    foreach (var item in query)
                    {
                        ss = ss + item.ToString() + ", ";
                    }
                    vals.Add(ss);
                    //MessageBox.Show(ss);
                }
            }


            //todotim

            {
                var query =
(from c in db.Emp_Allowances
 where !(from o in db.Employees
         select o.T1EmpNo)
     .Contains(c.T1_EmpID)
 select c.T1_EmpID).ToList();

                string ss;
                if (query.Count() == 0)
                    ;// MessageBox.Show("Paycomponent code looks ok");
                else
                {
//                    Issues = true;
                    ss = "The following EmployeeIds are in Employee Allowances but not in the Employees file: ";
                    foreach (var item in query)
                    {
                        ss = ss + item.ToString() + ", ";
                    }
                    //                  MessageBox.Show(ss);
                    vals.Add(ss);

                }

            }


            {
                char[] emptypes = new char[] { Char.Parse("C"), Char.Parse("F"), Char.Parse("P") };
                var query =
            (from c in db.Employees
             where !(from o in emptypes.ToList()
                     select o)
                    .Contains(c.Emp_Type.Value)
             select c.Emp_Type).ToList();

                string ss;
                if (query.Count() == 0)
                    ;// MessageBox.Show("Paycomponent code looks ok");
                else
                {
//                    Issues = true;
                    ss = "The following Employee Types are in Employee file and are not valid";
                    foreach (var item in query)
                    {
                        ss =ss+ item.ToString() + ", ";
                    }
                    vals.Add(ss);
//                    MessageBox.Show(ss);
                }
            }
            {
                var query =
        (from c in db.Employees
         where c.Hours<1 || c.Hours>40
         select c.T1EmpNo + " - "+ c.FirstName +" " + c.Surname).ToList();

                string ss;
                if (query.Count() == 0)
                    ;// MessageBox.Show("Paycomponent code looks ok");
                else
                {
  //                  Issues = true;
                    ss = "The following Employees have hours outside 1-40: ";
                    foreach (var item in query)
                    {
                        ss = ss + item.ToString() + ", ";
                    }
                    vals.Add(ss);
//                    MessageBox.Show(ss);
                }
            }
            {
                var query =
        (from c in db.Employees
         where c.T1EmpNo < 10000 || c.T1EmpNo > 99999
         select c.T1EmpNo + " - " + c.FirstName + " " + c.Surname).ToList();

                string ss;
                if (query.Count() == 0)
                    ;// MessageBox.Show("Paycomponent code looks ok");
                else
                {
  //                  Issues = true;
                    ss = "The following Employees have an EmpNo outside 10000-99999: ";
                    foreach (var item in query)
                    {
                        ss = ss + item.ToString() + ", ";
                    }
                    vals.Add(ss);
//                    MessageBox.Show(ss);
                }
            }

            if (vals.Count > 0)
            {
                (new frmValidation(vals)).ShowDialog();
                return true;
            }
            MessageBox.Show("No validation issues found with employee data");
            return false;
        }
        private void tsbValidate_Click(object sender, EventArgs e)
        {
            ValidateData();
        }

    }
}
