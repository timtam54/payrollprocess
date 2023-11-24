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
    public partial class frmJobCode : Form
    {
        public frmJobCode()
        {
            InitializeComponent();
        }
        DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

        private void frmJobCode_Load(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Maximized;
            FillNavGrid();
           

        }
        private void FillNavGrid()
        {
            var emp = db.Jobs.ToList();
            dataGridView1.DataSource = emp;
            dataGridView1.Columns[0].HeaderText = "Work Order";
            dataGridView1.Columns[1].HeaderText = "Work Order Description";
        }
        private void tbimport_Click(object sender, EventArgs e)
        {
            frmImpJob fie = new frmImpJob();
            fie.ShowDialog();
            FillNavGrid();
            if (ValidateData())
                MessageBox.Show("Issues found - you may need to export the Pay Components file from Techone again and reimport");
        }

        bool ValidateData()
        { 
            
            List<string> vals = new List<string>();
            {

                int TimesheetID = db.Timesheets.OrderByDescending(i => i.PayNoYear).Select(i => i.TimesheetID).FirstOrDefault();
                var query =
            (from c in db.TimesheetDatas
             where c.TimesheetID==TimesheetID &&  c.job!=null  && !(from o in db.Jobs
                     select o.JobCode).Contains(c.job)
             select c.job).ToList();
                int cnt = 0;
                string ss;
                if (query.Count() == 0)
                    ;// MessageBox.Show("Paycomponent code looks ok");
                else
                {
                    //                    Issues = true;
                    ss = "The following job codes are in Employee Timesheet but not in the job file: ";
                    foreach (var item in query)
                    {
                        if (item != "")
                        {
                            ss = ss + item.ToString() + ", ";
                            cnt++;
                        }
                    }
                    if (cnt > 0)
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

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            FillNavGrid();
        }

        private void toolStripSeparator1_Click(object sender, EventArgs e)
        {

        }
    }
}
