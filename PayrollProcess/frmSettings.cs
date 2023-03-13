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
    public partial class frmSettings : Form
    {
        public frmSettings()
        {
            InitializeComponent();
        }

        private void frmSettings_Load(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Maximized;
            FillNavGrid();

            DataGridViewButtonColumn but = new DataGridViewButtonColumn();
            but.Text = "Edit";
            but.Width = 200;
            dataGridView1.Columns.Add(but);

            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[1].Width = 200;
            dataGridView1.Columns[2].Width = 200;
            dataGridView1.Columns[3].Width = 200;

        }

        private void FillNavGrid()
        {
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);

            var emp = db.Settings.ToList();

            BindingSource bs = new BindingSource();
            bs.DataSource = emp;
            dataGridView1.DataSource = bs;
            bindingNavigator1.BindingSource = bs;
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            object dbi = dataGridView1.Rows[e.RowIndex].DataBoundItem;
            Setting st=(Setting)dbi;
            if (st.SettingCode == "Ledger_JobNoMap")
            {
                frmLedgerEdit seted = new frmLedgerEdit(st);
                seted.ShowDialog();

            }
            else
            {
                frmSettingsEdit seted = new frmSettingsEdit(st);
                seted.ShowDialog();
            }
            FillNavGrid();
        }
    }
}
