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
    public partial class frmPlantJobDate : Form
    {
        public frmPlantJobDate()
        {
            InitializeComponent();
        }

        public bool MatchJobDate = true;

        private void rbMatchOnJobDate_CheckedChanged(object sender, EventArgs e)
        {
            MatchJobDate = true;
        }

        private void rbMathOnDateOnly_CheckedChanged(object sender, EventArgs e)
        {
            MatchJobDate = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
