using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.SpecialCase
{
    public partial class frmUnChkWO : Form
    {
        public frmUnChkWO()
        {
            InitializeComponent();
        }
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.SpecialCase.CloseUnCheckWO Process = new DbLibrary.SpecialCase.CloseUnCheckWO();
        DataTable rs;

        private void reFreshData()
        {
            rs = Process.CloseWO_UnCheck();
            if(rs.Rows.Count > 0)
            {
                dgInfo.DataSource = rs;
            }
        }

        private void frmUnChkWO_Load(object sender, EventArgs e)
        {
            reFreshData();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (txtWO.Text.Trim() == "")
                if (string.IsNullOrEmpty(txtWO.Text.Trim()))
                {
                    MessageBox.Show("Please input WO !");
                    txtWO.Focus();
                    return;
                }
            rs = Process.GetCloseWO_UnCheck(txtWO.Text.Trim());
            if (rs.Rows.Count > 0)
            {
                MessageBox.Show("这个工单已经添加过，不需要再添加!");
                txtWO.Focus();
                return;
            }
            else
            {
                Process.Insert_data(txtWO.Text.Trim(), Parameter.g_userName);
                reFreshData();
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            //if (txtWO.Text.Trim() == "")
            if (string.IsNullOrEmpty(txtWO.Text.Trim()))
            {
                MessageBox.Show("Please input WO !");
                txtWO.Focus();
                return;
            }
            Process.Delete_data(txtWO.Text.Trim());
            reFreshData();
        }

        private void txtWO_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if (txtWO.Text.Trim() != "" && e.KeyChar == 13)
            if (!string.IsNullOrEmpty(txtWO.Text.Trim()) && e.KeyChar == 13)
            {
                btnAdd.Focus();
            }
        }

        private void frmUnChkWO_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmUnChkWO");
        }

    }
}
