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
    public partial class frmLedgerEdit : Form
    {
        Setting stng;
        public frmLedgerEdit(Setting _stng)
        {
            InitializeComponent();
            stng = _stng;
        }

   
        private void FrmLedgerEdit_Load(object sender, EventArgs e)
        {
            List<LedgerJob> ljs = new List<LedgerJob>();
            string[] ss = stng.Vals.Split(new string[] { "~" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var item in ss)
            {
                string[] ljitem = item.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                LedgerJob lj = new LedgerJob();
                if (ljitem.Length > 0)
                {
                    lj.Ledger = ljitem[0];
                    if (ljitem.Length>1)
                        lj.JobNoFirstLetter = ljitem[1];
                    if (ljitem.Length > 2)
                        lj.Length =  ljitem[2];
                    ljs.Add(lj);
                }
            }
            bs = new BindingSource();
            bs.DataSource = ljs.ToList();
            dataGridView1.DataSource = bs;
            bindingNavigator1.BindingSource = bs;
            dataGridView1.Columns[0].Width = 300;
        }
        BindingSource bs;
        private void TsbSaveClose_Click(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
            List<LedgerJob> ljs = (List<LedgerJob>)bs.DataSource;
            foreach (var item in ljs)
            {
                if (sb.Length > 0)
                    sb.Append("~");
                if (item.Length==null)
                    sb.Append(item.Ledger + "," + item.JobNoFirstLetter + ",");
                else
                    sb.Append(item.Ledger + "," + item.JobNoFirstLetter+","+item.Length);
            }
            var set = db.Settings.Where(st => st.SettingCode == stng.SettingCode).FirstOrDefault();
            set.Vals = sb.ToString();
            db.SubmitChanges();
            MessageBox.Show("Please restart application for settings to take effect");
            Close();

        }
    }
}
