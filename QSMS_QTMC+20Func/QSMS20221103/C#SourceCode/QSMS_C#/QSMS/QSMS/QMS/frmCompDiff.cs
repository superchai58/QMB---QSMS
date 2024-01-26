using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.QMS
{
    public partial class frmCompDiff : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.QMS.CompDiffProcess bll = new DbLibrary.QMS.CompDiffProcess();

        public frmCompDiff()
        {
            InitializeComponent();
        }

        private void frmCompDiff_Load(object sender, EventArgs e)
        {
            GetLine();
        }

        private void GetLine()
        {
            cboLine.Items.Clear();
            cboLine.Items.Add("");
            DataTable dtLine = bll.GetLine();
            for (int i = 0; i < dtLine.Rows.Count; i++)
            {
                cboLine.Items.Add(dtLine.Rows[i]["Line"].ToString());
            }
            cboLine.SelectedIndex = 0;
        }

        private void GetGroupID()
        {
            cboGroupID.Items.Clear();
            cboGroupID.Items.Add("");
            string strLine = cboLine.Text.Trim();
            //方法一
            //string BegDT=string.Format("{0:yyyyMMdd}",DateTime.Parse(dtpBDate.Text.Trim()));
            //string EndDT = string.Format("{0:yyyyMMdd}", DateTime.Parse(dtpEDate.Text.Trim()));
            //方法二
            string BegDT = dtpBDate.Value.ToString("yyyyMMdd");
            string EndDT = dtpEDate.Value.ToString("yyyyMMdd");

            DataTable dtGroupID = bll.GetGroupID(BegDT, EndDT, strLine);
            if (dtGroupID.Rows.Count > 0)
            {
                for (int j = 0; j < dtGroupID.Rows.Count; j++)
                {
                    cboGroupID.Items.Add(dtGroupID.Rows[j]["GroupID"].ToString());
                }
            }
            else
            {
                MessageBox.Show("NO Data");
            }
                cboGroupID.SelectedIndex = 0;
        }

        private void GetGroupWO(string GroupID)
        {
            cboOKWO.Items.Clear();
            cboOKWO.Items.Add("");

            DataTable dtWO = bll.GetGroupWO(GroupID);
            if (dtWO.Rows.Count > 0)
            {
                for (int k = 0; k < dtWO.Rows.Count; k++)
                {
                    cboOKWO.Items.Add(dtWO.Rows[k]["Work_Order"].ToString());
                }
            }
            cboOKWO.SelectedIndex = 0;
        }

        private void GetWoInfo(string WO)
        {
            DataTable dtWOinfo = bll.GetWoInfo(txtWO.Text.Trim());
            if (dtWOinfo.Rows.Count > 0)
            {
                txtMBPN.Text = dtWOinfo.Rows[0]["PN"].ToString();
                txtQty.Text = dtWOinfo.Rows[0]["Qty"].ToString();
                txtRev.Text = dtWOinfo.Rows[0]["MB_Rev"].ToString();
                txtModel.Text = dtWOinfo.Rows[0]["PN"].ToString() + "-" + dtWOinfo.Rows[0]["MB_Rev"].ToString();
            }
        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(cboLine.Text.Trim()))
            {
                MessageBox.Show("Please input Line!!!");
                return;
            }
            GetGroupID();
        }

        private void cboGroupID_SelectedValueChanged(object sender, EventArgs e)
        {
            GetGroupWO(cboGroupID.Text.Trim());
        }

        private void cboOKWO_SelectedValueChanged(object sender, EventArgs e)
        {
            txtWO.Text = cboOKWO.Text.Trim();
            GetWoInfo(txtWO.Text.Trim());
        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            DataTable dtChkWO = bll.ChkWO(cboOKWO.Text.Trim());
            if (dtChkWO.Rows.Count > 0)
            {
                dgvDispatch.DataSource = dtChkWO;
                dgvDispatch.Show();
            }
            else
            {
                MessageBox.Show("Work_Order:{0} check OK!",cboOKWO.Text.Trim());
            }
        }

        private void frmCompDiff_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmCompDiff");
        }
    }
}
