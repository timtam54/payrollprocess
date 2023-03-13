using ClosedXML.Excel;
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
    public partial class frmExcel : Form
    {
        public enum Plant_TS123
        {

            Plant=0,
            TS1=1,
            TS2=2,
            TS3=3,
            NA=4
        }

        public frmExcel(DataTable dt, Plant_TS123 _plant_ts)
        {
            ;//delete thi mthod
        }
        public class ExcelTab
        {
            public string TABLE_NAME { get; set; }
        }
        public string SheetName;
        Plant_TS123 plant_ts;
        public frmExcel(IXLWorksheets dt, Plant_TS123 _plant_ts)
        {
            InitializeComponent();
            {
                var xx = (from d in  dt select new ExcelTab { TABLE_NAME = d.Name }).ToList();
                comboBox1.DataSource = xx;
                comboBox1.DisplayMember = "TABLE_NAME";
                comboBox1.ValueMember = "TABLE_NAME";
                plant_ts = _plant_ts;
            }
        }

        private void frmExcel_Load(object sender, EventArgs e)
        {
            if (comboBox1.Items.Count == 1)
            {
                comboBox1.SelectedIndex = 0;
                SheetName = comboBox1.SelectedValue.ToString();
                Close();
                return;
            }
            bool found = false;
            for (int i = 0; i < comboBox1.Items.Count; i++)
            {
                comboBox1.SelectedIndex = i;
                ExcelTab dr = ((ExcelTab)comboBox1.SelectedItem);
                object oo = dr.TABLE_NAME;
                if (plant_ts == Plant_TS123.Plant)
                {
                    if (oo.ToString().ToLower().Contains("plant"))
                    {
                        found = true;
                        break;
                    }
                }
                else if (plant_ts == Plant_TS123.TS1 && oo.ToString().ToLower().Contains("_t1-1"))
                {
                    found = true;
                    break;
                }
                else if (plant_ts == Plant_TS123.TS2 && oo.ToString().ToLower().Contains("_t1-2"))
                {
                    found = true;
                    break;
                }
                else if (plant_ts == Plant_TS123.TS3 && oo.ToString().ToLower().Contains("_t1-3"))
                {
                    found = true;
                    break;
                }
                else if (plant_ts == Plant_TS123.NA)
                {
                    found = true;
                    break;

                }
            }
            if (!found)
            {
                for (int i = 0; i < comboBox1.Items.Count; i++)
                {
                    comboBox1.SelectedIndex = i;
                    ExcelTab dr = ((ExcelTab)comboBox1.SelectedItem);
                    object oo = dr.TABLE_NAME;
                    if (plant_ts == Plant_TS123.TS1 && oo.ToString().ToLower().Contains("_t1"))
                    {
                        found = true;
                        break;
                    }
                }
            }
            if (!found)
            {
                for (int i = 0; i < comboBox1.Items.Count; i++)
                {
                    comboBox1.SelectedIndex = i;
                    ExcelTab dr = ((ExcelTab)comboBox1.SelectedItem);
                    object oo = dr.TABLE_NAME;
                    if (oo.ToString().ToLower().Contains("timesheet"))
                        break;
                }
            }
            SheetName = comboBox1.SelectedValue.ToString();
            Close();
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            SheetName = comboBox1.SelectedValue.ToString();
            Close();
        }
    }
}
