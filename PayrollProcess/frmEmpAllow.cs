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
    public partial class frmEmpAllow : Form
    {
        public frmEmpAllow()
        {
            InitializeComponent();
        }


        private void frmEmpAllow_Load(object sender, EventArgs e)
        {
            this.Text = "Employee Allowances";
            WindowState = FormWindowState.Maximized;
            FillNavGrid();
        }

        private void FillNavGrid()
        {
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

            var emp = db.Emp_Allowances.ToList();
            BindingSource bs = new BindingSource();
            bs.DataSource = emp;
            dataGridView1.DataSource = bs;
            bindingNavigator1.BindingSource = bs;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            (new frmImpEmpAllow()).ShowDialog();
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

        private void tsbValidate_Click(object sender, EventArgs e)
        {
            ValidateData();
        }

        bool ValidateData()
        {
           // bool Issues = false;
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
            List<string> vals = new List<string>();
            //todotim
            {
                var query =
(from c in db.Employees
 where !(from o in db.Emp_Allowances
         select o.T1_EmpID)
     .Contains(c.T1EmpNo)
 select c.T1EmpNo+" "+c.FirstName+" " + c.Surname).ToList();

                string ss;
                if (query.Count() == 0)
                    ;// MessageBox.Show("Paycomponent code looks ok");
                else
                {
//                    Issues = true;
                    ss = "The following EmployeeIds are in Employees but not in the Employees Allowances file: ";
                    foreach (var item in query)
                    {
                        ss = ss + item.ToString() + ", ";
                    }
                    //MessageBox.Show(ss);
                    vals.Add(ss);
                }

            }

            {
                var query =
        (from c in db.Emp_Allowances
         where c.units < 1 || c.units > 100
         select c.T1_EmpID + " " + c.PayComponentCode).ToList();

                string ss;
                if (query.Count() == 0)
                    ;// MessageBox.Show("Paycomponent code looks ok");
                else
                {
  //                  Issues = true;
                    ss = "The following Employees/PayCompCodes have units outside 1-100: ";
                    foreach (var item in query)
                    {
                        ss = ss + item.ToString() + ", ";
                    }
                    vals.Add(ss);
                }
            }
            {
                var query =
(from c in db.Emp_Allowances
where !(from o in db.PayComponents
     select o.PayCompCode)
    .Contains(c.PayComponentCode)
select c.PayComponentCode).ToList();

                string ss;
                if (query.Count() == 0)
                    ;// MessageBox.Show("Paycomponent code looks ok");
                else
                {
    //                Issues = true;
                    ss = "The following paycomponent codes are in Employee Allowances but not in the Paycomponent file: ";
                    foreach (var item in query)
                    {
                        ss = ss + item.ToString() + ", ";
                    }
                    vals.Add(ss);
                }

            }
            if (vals.Count > 0)
            {
                (new frmValidation(vals)).ShowDialog();
                return true;
            }
                MessageBox.Show("No validation issues found with Employee allowance data");
            return false;

        }

    }
}
