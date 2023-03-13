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
    public partial class frmPayPeriod : Form
    {
        public frmPayPeriod(int _PayPeriod)
        {
            InitializeComponent();
            PayPeriod = _PayPeriod;
        }

        public class ValDesc
        {
            public int Val { get; set; }
            public string Desc { get; set; }

        }
        DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
        private void frmPayPeriod_Load(object sender, EventArgs e)
        {
            cboPayPeriod.DataSource = (from pp in db.PayYears orderby -pp.PayNoYear  select new ValDesc {Val=pp.PayNoYear,Desc=pp.PayNoYear + "-"+((pp.Comment==null)?"":pp.Comment)} ).ToList();
            cboPayPeriod.ValueMember = "Val";
            cboPayPeriod.DisplayMember = "Desc";


            for (int i = 0; i < cboPayPeriod.Items.Count; i++)
            {
                ValDesc py = (ValDesc)cboPayPeriod.Items[i];
                if (py.Val == PayPeriod)
                {
                    cboPayPeriod.SelectedIndex = i;
                    break;
                }
            }

        }

        public int PayPeriod;
        private void button1_Click(object sender, EventArgs e)
        {
            PayPeriod =Convert.ToInt32( cboPayPeriod.SelectedValue);
            Close();
        }
    }
}
