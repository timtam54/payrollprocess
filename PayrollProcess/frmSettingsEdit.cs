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
    public partial class frmSettingsEdit : Form
    {
        Setting stng;
        public frmSettingsEdit(Setting _stng)
        {
            InitializeComponent();
            stng = _stng;
        }

        private void frmSettingsEdit_Load(object sender, EventArgs e)
        {
            label1.Text = stng.SettingDesc;
            textBox1.Text = stng.Vals;
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            DataClasses1DataContext db = new DataClasses1DataContext(Form1.ConString);
            var set= db.Settings.Where(st => st.SettingCode == stng.SettingCode).FirstOrDefault();
            set.Vals = textBox1.Text;
            db.SubmitChanges();
            MessageBox.Show("Please restart application in order for settings to take effect");

            Close();
        }
    }
}
