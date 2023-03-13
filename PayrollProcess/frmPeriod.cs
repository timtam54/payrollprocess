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
    public partial class frmPeriod : Form
    {
        public frmPeriod()
        {
            InitializeComponent();
        }

        private void frmPeriod_Load(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Maximized;
            FillNavGrid();

            DataGridViewButtonColumn but = new DataGridViewButtonColumn();
            but.Text = "Edit";
            but.Width = 40;
            dataGridView1.Columns.Add(but);

            //dataGridView1.Columns[0].Width = 200;
            //dataGridView1.Columns[1].Width = 200;
            //dataGridView1.Columns[2].Width = 200;
            dataGridView1.Columns[3].Width = 300;

        }

        private void FillNavGrid()
        {
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

            List<PayYear> emp = db.PayYears.ToList();
            if (emp.Count() == 0)
            {
                PayYear NewPayYear = new PayYear();
                NewPayYear.StartDate = new DateTime(2018, 11, 24);
                NewPayYear.EndDate = NewPayYear.StartDate.AddDays(13);
                NewPayYear.PayNoYear = 201812;
                db.PayYears.InsertOnSubmit(NewPayYear);
                db.SubmitChanges();
                //FillNavGrid();
                emp = new List<PayYear>() { NewPayYear };
            }

            BindingSource bs = new BindingSource();
            bs.DataSource = emp;
            dataGridView1.DataSource = bs;
            bindingNavigator1.BindingSource = bs;
        }

        private void tbbAddMore_Click(object sender, EventArgs e)
        {
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

            var LastPayYear = db.PayYears.Max(py => py.PayNoYear);
            var LastPayYearRecord = db.PayYears.Where(py => py.PayNoYear == LastPayYear).FirstOrDefault();
            int Year = Convert.ToInt32(LastPayYear.ToString().Substring(0, 4));
            int Period = Convert.ToInt32(LastPayYear.ToString().Substring(4, 2));
            DateTime LastEnd = LastPayYearRecord.EndDate;
            DateTime Start = LastEnd.AddDays(1);
            if ((LastEnd.Month == 6) && (Start.Month == 7))
            {
                Year++;
                Period = 0;
            }
            DateTime End = LastEnd.AddDays(14);
            if ((End.Month == 7) && (Start.Month == 6))
                End = new DateTime(End.Year, 6, 30);


            PayYear NewPayYear = new PayYear();
            NewPayYear.StartDate = Start;
            NewPayYear.EndDate = End;
            Period++;
            NewPayYear.PayNoYear = Year * 100 + Period;
            db.PayYears.InsertOnSubmit(NewPayYear);
            db.SubmitChanges();
            FillNavGrid();





        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            object dbi = dataGridView1.Rows[e.RowIndex].DataBoundItem;
            PayYear st = (PayYear)dbi;

            frmPeriodEdit seted = new frmPeriodEdit(st);
            seted.ShowDialog();

            FillNavGrid();
        }
    }
}