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
    public partial class frmPayComp : Form
    {
        public frmPayComp()
        {
            InitializeComponent();
        }

        private void frmPayComp_Load(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Maximized;
            FillNavGrid();

            DataGridViewButtonColumn but = new DataGridViewButtonColumn();
            but.Text = "X";
            but.Width = 40;
            dataGridView1.Columns.Add(but);

        }
        //        DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

        private void FillNavGrid()
        {

            DataClasses1DataContext dbs = new DataClasses1DataContext(Form1.ConString);

            var emp = dbs.PayComponents.ToList();

            BindingSource bs = new BindingSource();
            bs.DataSource = emp;
            dataGridView1.DataSource = bs;
            bindingNavigator1.BindingSource = bs;


        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            frmImpPayComp fie = new frmImpPayComp();
            fie.ShowDialog();
            FillNavGrid();
            try
            {
                if (ValidateData())
                    MessageBox.Show("Issues found - you may need to export the Pay Components file from Techone again and reimport");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
            int rw = e.RowIndex;
            object oo= dataGridView1.Rows[rw].DataBoundItem;
            PayrollProcess.PayComponent pc = (PayrollProcess.PayComponent)oo;
            if (MessageBox.Show("Do you wish to delete paycomp " + pc.PayCompCode,"Delete", MessageBoxButtons.YesNo)== DialogResult.Yes)
            {
                PayComponent pcdel = db.PayComponents.Where(p => p.PayCompCode == pc.PayCompCode).FirstOrDefault();
                db.PayComponents.DeleteOnSubmit(pcdel);
                db.SubmitChanges();
                FillNavGrid();
            }
        }

        bool ValidateData()
        {
            //            bool Issues = false;
            List<string> vals = new List<string>();
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
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
                //Issues = true;
                ss = "The following paycomponent codes are in Employee Allowances but not in the Paycomponent file: ";
                foreach (var item in query)
                {
                    ss = ss+item.ToString() + ", ";
                }
               vals.Add(ss) ;// MessageBox.Show(ss);
            }
            {
                string[] codes = frmReport.CodesThatContribToToolAllow.Split(new char[] { Convert.ToChar(",") });
                ss = "";
                foreach (var cd in codes)
                {
                    if (db.PayComponents.Where(i => i.PayCompCode == Convert.ToDecimal(cd)).Count() == 0)
                    {
                        ss = ss+cd.ToString() + ", ";
                    }
                }
                if (ss != "")
                {
               //     Issues = true;
//                    MessageBox.Show();
                    vals.Add("The following paycomponent codes are in CodesThatContribToToolAllow but not in the Paycomponent file:" + ss);//
                }
            }
            {
                foreach (var cd in frmReport.CreateOTTransForCodes)
                {
                    if (db.PayComponents.Where(i => i.PayCompCode == Convert.ToDecimal(cd)).Count() == 0)
                    {
                        ss = ss+cd.ToString() + ", ";
                    }
                }
                if (ss != "")
                {
                    // Issues = true;
                    vals.Add("The following paycomponent codes are in CreateOTTransForCodes but not in the Paycomponent file:" + ss);
                }
            }
            if (db.PayComponents.Where(i => i.PayCompDesc.ToLower().Contains("overtime")).Count() == 0)
            {
                //Issues = true;
                vals.Add("Normally there would be one or more paycomponent with 'PayCompDesc' column containing the word 'Overtime'");
            }
            if (db.PayComponents.Where(i => i.PayCompTypeDesc.ToLower() == "overtime").Count() == 0)
            {
                //Issues = true;
                vals.Add("Normally there would be one or more paycomponent with 'PayCompTypeDesc' column containing the word 'Overtime'");
            }
            if (db.PayComponents.Where(i => i.PayCompTypeDesc.ToLower() == "accruals").Count() == 0)
            {
                //Issues = true;
                vals.Add("Normally there would be one or more paycomponent with 'PayCompTypeDesc' column containing the word 'Accruals'");
            }
            if (db.PayComponents.Where(i => i.PayCompTypeDesc.ToLower() == "absences").Count() == 0)
            {
                //Issues = true;
                vals.Add("Normally there would be one or more paycomponent with 'PayCompTypeDesc' column containing the word 'Absences'");
            }
            string[] dp = new string[] { "Neither", "Deduction", "Payment" };
            var query3 =
     (from c in db.PayComponents
      where c.Payment_deduct != null && !(from o in dp
                                          select o)
             .Contains(c.Payment_deduct)
      select c.PayCompCode.ToString() +"-"+c.PayCompDesc+" ("+  c.Payment_deduct+")").ToList();
            if (query3.Count() == 0)
                ;// MessageBox.Show("Paycomponent code looks ok");
            else
            {
                //Issues = true;
                ss = "The following paycomponent the Payment_deduct column is neither (Deduction,Payment, or Neither) - typically one of these options would be expected: ";
                foreach (var item in query3)
                {
                    ss = ss+item.ToString() + ", ";
                }
                vals.Add(ss);
            }


            string[] ppu = new string[] { "Hour", "Only", "Percentage", "Kilometers", "Week", "", "Day" };
            var query4 =
     (from c in db.PayComponents
      where c.PayPeriodUnit != null && !(from o in ppu
                                         select o)
             .Contains(c.PayPeriodUnit)
      select c.PayPeriodUnit).ToList();
            if (query4.Count() == 0)
                ;// MessageBox.Show("Paycomponent code looks ok");
            else
            {
//                Issues = true;
                ss = "The following paycomponent the Payment_deduct column is neither (Deduction,Payment, or Neither) - typically one of these options would be expected: ";
                foreach (var item in query4)
                {
                    ss = ss+item.ToString() + ", ";
                }
                vals.Add(ss);
            }
            if (db.PayComponents.Where(i => i.Units != 0).Count() == 0)
            {
                //Issues = true;
                vals.Add("All units in Paycomponets are all 0.  This is not what would be expected");
            }
            if (vals.Count > 0)
            {
                (new frmValidation(vals)).ShowDialog();
                return true;
            }
            MessageBox.Show("No validation issues found with Pay components data");
            return false;
        }

        private void tsbValidate_Click(object sender, EventArgs e)
        {
            try
            {
                ValidateData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}
