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
    public partial class frmEvents : Form
    {
        public frmEvents()
        {
            InitializeComponent();
        }

        private void frmEvents_Load(object sender, EventArgs e)
        {
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
            dataGridView1.DataSource = db.EventLogs.OrderByDescending(el=>el.EventLogID).ToList();
        }
    }
}
