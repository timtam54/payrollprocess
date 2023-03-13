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
    public partial class frmFilter : Form
    {
        Point myPos;
        DataTable dt;
        string FilterColumn;
        public frmFilter(DataTable _dt, string _FilterColumn, Point _myPos)
        {
            InitializeComponent();
            dt = _dt;
            FilterColumn = _FilterColumn;
            DataView dv = dt.DefaultView;
            dv.Sort = FilterColumn;
            DataTable sorted = dv.ToTable();
            foreach (DataRow item in sorted.Rows)
            {
                string val = item[FilterColumn].ToString();
                if (!cboFilterVal.Items.Contains(val))
                    cboFilterVal.Items.Add(val);
            }
            myPos = _myPos;
        }

        public string FilterVal;

      


        private void frmFilter_Load(object sender, EventArgs e)
        {
            this.Left = myPos.X;
            this.Top = myPos.Y;
            //cboFilterValOld.DroppedDown = true;
            this.Height = 40;// 
            cboFilterVal.Height = cboFilterVal.Items.Count * 20;
            this.Height= cboFilterVal.Height + 40;
        }

        private void btnFilter_Click(object sender, EventArgs e)
        {
            if (cboFilterVal.CheckedItems.Count==0)
            {
                MessageBox.Show("Nothing selected");
                return;

            }
            FilterVal = "";// -Remove Filter-";
            foreach (var item in cboFilterVal.CheckedItems)
            {
                if (FilterVal != "")
                    FilterVal = FilterVal + ",";

                string dtp = dt.Columns[FilterColumn.ToString()].DataType.Name;
                if (dtp == "String")
                    FilterVal = FilterVal + "'" + item.ToString() +"'";

                else


                    FilterVal = FilterVal+item.ToString();

            }
            Close();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            FilterVal = "-Remove Filter-";
            Close();
        }

        private void BtnCheckAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < cboFilterVal.Items.Count; i++)
            {
                cboFilterVal.SetItemChecked(i, true);
            }
        }

        private void uncheckall_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < cboFilterVal.Items.Count; i++)
            {
                cboFilterVal.SetItemChecked(i, false);
            }
        }

        private void CboFilterVal_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
