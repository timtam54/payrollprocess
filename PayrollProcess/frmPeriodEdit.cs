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
    public partial class frmPeriodEdit : Form
    {
        PayYear period;
        public frmPeriodEdit(PayYear _period)
        {
            InitializeComponent();
            period = _period;
        }

        private void FrmPeriodEdit_Load(object sender, EventArgs e)
        {
            if (period.StartDate.Year!=1)
                nudFrom.Value = period.StartDate;
            if (period.EndDate.Year !=1)
                nudTo.Value = period.EndDate;
            budPeriod.Value = period.PayNoYear;
            textBox1.Text = period.Comment;
            if (db.Timesheets.Where(ts => ts.PayNoYear == period.PayNoYear).Count() == 0)
            {
                btnDelete.Visible = true;
                lblDel.Text = "There are no timesheet headers associated with this Period so you can delete this period";
                lblDel.ForeColor = Color.Green;
            }
            else
            {
                btnDelete.Visible = false;
                lblDel.Text = "There are timesheet headers associated with this Period so you cannot delete this period";
                lblDel.ForeColor = Color.Red;

            }

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            Form1.ExecSQL("update payyear set [PayNoYear]="+ Convert.ToInt32(budPeriod.Value).ToString() + " from payyear where [PayNoYear]="+ period.PayNoYear);
            period = db.PayYears.Where(py => py.PayNoYear == Convert.ToInt32(budPeriod.Value)).FirstOrDefault();

            period.StartDate=nudFrom.Value.Date;
            period.EndDate=nudTo.Value.Date;
            period.Comment=textBox1.Text ;
            db.SubmitChanges();
            Close();
            MessageBox.Show("You will need to restart the app for New Periods to take effect");
        }
        DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (!budPeriod.Value.Equals(period.PayNoYear))
            {
                MessageBox.Show("Period has been modified - cannot delete");
                return;
            }
            var pny = db.PayYears.Where(py => py.PayNoYear == period.PayNoYear).FirstOrDefault();
            db.PayYears.DeleteOnSubmit(pny);
            db.SubmitChanges();
            MessageBox.Show("Deleted - please restart app for period to be cleared from list");
            Close();
        }

        private void BudPeriod_ValueChanged(object sender, EventArgs e)
        {
            btnDelete.Visible = false;
        }
    }
}
