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
    public partial class frmValidation : Form
    {

        public class dta
        {
            public string Issue { get; set; }
        }
        public frmValidation(List<string> ss)
        {
            InitializeComponent();
            List<dta> iss = (from s in ss select new dta { Issue = s }).ToList();
            dataGridView1.DataSource = iss;//.ToList();
            dataGridView1.Columns[0].Width = 900;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmValidation_Load(object sender, EventArgs e)
        {

        }
    }
}
