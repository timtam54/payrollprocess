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
    public partial class frmStaffMIssing : Form
    {
        int PayPeriod;

        public frmStaffMIssing(int _PayPeriod)
        {
            InitializeComponent();
            PayPeriod = _PayPeriod;
        }

        DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

        private void frmStaffMIssing_Load(object sender, EventArgs e)
        {
            var list = (from emp in db.Employees
                        where !db.Timesheets.Where(ts => ts.PayNoYear == PayPeriod).Any(f => f.StaffID == emp.T1EmpNo)
                        select emp).ToList();
            dataGridView1.DataSource = list;
        }
    }
}
