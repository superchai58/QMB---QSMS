using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.PD
{
    public partial class frmUnlinkFeederDID : Form
    {
        public frmUnlinkFeederDID()
        {
            InitializeComponent();
        }
        DbLibrary.PD.UnlinkFeederDID UnlinkFeederDID = new DbLibrary.PD.UnlinkFeederDID();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        string DID = string.Empty;
        string strSQL = string.Empty;
        DataTable dt = new DataTable();

        private void frmUnlinkFeederDID_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmUnlinkFeederDID");
        }

        private void txtFeederDID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && !string.IsNullOrEmpty(txtFeederDID.Text.ToString()))
            {
                if (txtFeederDID.Text.IndexOf("-") > -1)
                {
                    DID = txtFeederDID.Text.ToString().Trim();
                    txtDID.Text = DID;
                }
                else
                {
                    dt = UnlinkFeederDID.QueryFeederDID(txtFeederDID.Text.ToString().Trim());
                    if (dt.Rows.Count > 0)
                    {
                        DID = dt.Rows[0]["DID"].ToString();
                        txtDID.Text = DID;
                    }
                    else
                    {
                        MessageBox.Show("没有找到记录,请检查");
                        lblMsg.BackColor = Color.Green;
                        lblMsg.Text = "Unlink DID OK";
                        pubFunction.Sound("OK");
                        return;
                    }
                }
            }
            else
            {
                lblMsg.Text = "请输入FeederDID!按回车确认";
                txtFeederDID.Focus();
                return;
            }
        }

        private void frmUnlinkFeederDID_Load(object sender, EventArgs e)
        {
            cboStatus.Items.Add("Finished");
            cboStatus.Items.Add("NotFinished");
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (pubFunction.ConfigListGetValue("UnlinkFeederDID") == "Y")
            {
                if (!string.IsNullOrEmpty(txtFeederDID.Text.ToString()))
                {
                    dt = UnlinkFeederDID.UnlinkFeeder(txtFeederDID.Text.ToString().Trim(), Parameter.g_userName, "");
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["Result"].ToString() == "1")
                        {
                            lblMsg.Text = dt.Rows[0]["Msg"].ToString();
                            lblMsg.BackColor = Color.Red;
                            pubFunction.Sound("ERROR");
                            return;
                        }
                        else
                        {
                            lblMsg.Text = dt.Rows[0]["Msg"].ToString();
                            lblMsg.BackColor = Color.Green;
                            pubFunction.Sound("OK");
                            return;
                        }
                    }
                }
                else
                {
                    lblMsg.Text = "请输入FeederDID!";
                    txtFeederDID.Focus();
                    return;
                }

            }
            else
            {
                if (!string.IsNullOrEmpty(txtFeederDID.Text.ToString()))
                {
                    if (txtFeederDID.Text.IndexOf("-") > -1)
                    {
                        DID = txtFeederDID.Text.ToString().Trim();
                        txtDID.Text = DID;
                    }
                    else
                    {
                        dt = UnlinkFeederDID.QueryFeederDID(txtFeederDID.Text.ToString().Trim());
                        if (dt.Rows.Count > 0)
                        {
                            DID = dt.Rows[0]["DID"].ToString();
                            txtDID.Text = DID;
                        }
                        else
                        {
                            MessageBox.Show("没有找到记录,请检查");
                            lblMsg.BackColor = Color.Green;
                            lblMsg.Text = "Unlink DID OK";
                            pubFunction.Sound("OK");
                            return;
                        }
                    }
                }
                else
                {
                    lblMsg.Text = "请输入FeederDID!";
                    txtFeederDID.Focus();
                    return;
                }

                UnlinkFeederDID.DeleteFeederDID(Parameter.g_userName, DID);   //0001

                UnlinkFeederDID.DeleteFromFeederDID(DID);           //0001

                if (Parameter.BU == "ESBU")              ////20220628  增加删除QSMS_FeederDID_Current表数据的动作
                {
                    UnlinkFeederDID.DeleteFromFeederDID_Current(DID);          
                }

                lblMsg.BackColor = Color.Green;
                lblMsg.Text = "Unlink DID OK";
                pubFunction.Sound("OK");
            }

        }

    }
}
