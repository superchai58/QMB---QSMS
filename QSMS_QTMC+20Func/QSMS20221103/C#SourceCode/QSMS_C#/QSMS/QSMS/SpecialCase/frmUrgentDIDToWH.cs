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
    public partial class frmUrgentDIDToWH : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.SpecialCase.UrgentDIDToWH UrgentDIDToWH = new DbLibrary.SpecialCase.UrgentDIDToWH();
        public frmUrgentDIDToWH()
        {
            InitializeComponent();
        }

        private void cmdDelete_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            //if (txtRefID.Text == "")
            if (string.IsNullOrEmpty(txtRefID.Text.Trim()))
            {
                MessageBox.Show("Please input ReferenceID !");
                txtRefID.Focus();
                return;
            }
            if (MessageBox.Show("Do you delete it really" + " ?\r\n", "确认", MessageBoxButtons.YesNo).ToString().ToUpper() == "YES")
            {
                dt = UrgentDIDToWH.QSMS_DID_ToWH(txtRefID.Text);
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("Can not find this ReferenceID in QSMS_DID_ToWH !");
                    txtRefID.Focus();
                    return;
                }
                else
                {
                    dt = UrgentDIDToWH.XL_UrgentToWH(txtRefID.Text, Parameter.UID, "Delete");
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["Result"].ToString() == "1")
                        {
                            dgInfo.DataSource = null;
                            MessageBox.Show("Delete OK");
                        }
                        else
                        {
                            MessageBox.Show("Delete Fail");
                        }
                    }
                }
            }
        }

        private void cmdQuery_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            //if (txtRefID.Text == "")
            if (string.IsNullOrEmpty(txtRefID.Text.Trim()))
            {
                MessageBox.Show("Please input ReferenceID !");
                txtRefID.Focus();
                return;
            }
            dt = UrgentDIDToWH.QSMS_DID_ToWH(txtRefID.Text);
            if (dt.Rows.Count == 0)
            {
                dgInfo.DataSource = null;
                MessageBox.Show("Can not find this ReferenceID in QSMS_DID_ToWH !");
                txtRefID.Focus();
                return;
            }
            else
            {
                dgInfo.DataSource = dt;

            }
        }

        private void cmdUpdate_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            //if (txtRefID.Text == "")
            if (string.IsNullOrEmpty(txtRefID.Text.Trim()))
            {
                MessageBox.Show("Please input ReferenceID !");
                txtRefID.Focus();
                return;
            }
            //string aa = MessageBox.Show("Do you update it really" + " ?\r\n", "确认", MessageBoxButtons.YesNo).ToString().ToUpper();
            if (MessageBox.Show("Do you update it really" + " ?\r\n", "确认", MessageBoxButtons.YesNo).ToString().ToUpper() == "YES")
            {
                dt = UrgentDIDToWH.QSMS_DID_ToWH(txtRefID.Text);
                if (dt.Rows.Count == 0)
                {
                    dgInfo.DataSource = null;
                    MessageBox.Show("Can not find this ReferenceID in QSMS_DID_ToWH !");
                    txtRefID.Focus();
                    return;
                }
                else
                {
                    dt = UrgentDIDToWH.XL_UrgentToWH(txtRefID.Text, Parameter.UID, "Update");
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["Result"].ToString() == "1")
                        {
                            dgInfo.DataSource = null;
                            MessageBox.Show("Update OK");
                        }
                        else
                        {
                            MessageBox.Show("Update Fail");
                        }
                    }
                    cmdQuery_Click(null, null);
                }
            }
        }

        private void txtRefID_KeyPress(object sender, KeyPressEventArgs e)
        {
            //if (e.KeyChar == 13 && txtRefID.Text != "")
            if (e.KeyChar == 13 && string.IsNullOrEmpty(txtRefID.Text.Trim()))
            {
                cmdQuery_Click(null, null);
            }
        }

        private void frmUrgentDIDToWH_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmUrgentDIDToWH");
        }
    }
}
