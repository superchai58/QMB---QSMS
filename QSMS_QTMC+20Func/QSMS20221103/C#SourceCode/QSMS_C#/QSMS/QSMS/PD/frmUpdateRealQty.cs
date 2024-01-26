using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace QSMS.QSMS.PD
{
    public partial class frmUpdateRealQty : Form
    {
        string Line = string.Empty;
        string Side = string.Empty;
        DataTable dt = new DataTable();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.PD.PDProcess process = new DbLibrary.PD.PDProcess();
        public frmUpdateRealQty()
        {
            InitializeComponent();
        }
        private void btnReset_Click(object sender, EventArgs e)
        {
            txtDID.Text = "";
            txtRealQty.Text = "";
            txtTotalQty.Text = "";
            txtUpdateTo.Text = "";
            txtReason.Text = "";
            txtDID.Focus();
        }

        private void txtDID_TextChanged(object sender, EventArgs e)
        {
            txtMsg.Text = "";
            txtDID.Text = txtDID.Text.ToString().ToUpper();
            txtDID.Select(txtDID.Text.Length, 0);
            txtDID.ScrollToCaret();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (txtDID.Text != "" && txtRealQty.Text != "" && txtUpdateTo.Text != "" && txtReason.Text != "")
            {
                if(Regex.IsMatch(txtUpdateTo.Text.ToString(), "^[0-9]*$"))
                {
                    if (Int32.Parse(txtUpdateTo.Text.ToString()) > Int32.Parse(txtTotalQty.Text.ToString()))
                    {
                        txtMsg.Text = "Update Qty can not larger than Total Qty!";
                        txtUpdateTo.Text = "";
                        txtUpdateTo.Focus();
                        return;
                    }
                    process.UpdateRealQty(txtUpdateTo.Text.ToString(), txtDID.Text.ToString());
                    //process.InsertLog("Line=" + Line + ";Side=" + Side + ";TotalQty=" + txtTotalQty.Text.ToString() + ";RealQty=" + txtRealQty.Text.ToString() + ";UpdateTo=" + txtUpdateTo.Text.ToString());
                    process.InsertLog("Line=" + Line + ";Side=" + Side + ";TotalQty=" + txtTotalQty.Text.ToString() + ";RealQty=" + txtRealQty.Text.ToString() + ";UpdateTo=" + txtUpdateTo.Text.ToString(), txtDID.Text.ToString() ,";Reason="+txtReason .Text .ToString ());  //20230116 Ellen
                    btnReset_Click(sender, e);
                    txtMsg.Text = "OK!";
                }
                else
                {
                    MessageBox.Show("请数入数字");
                    txtUpdateTo.Text = "";
                }
            }
            else
            {
                MessageBox.Show("DID/RealQty/UpdateTo/Reason 不能为空");
            }
        }

        private void txtDID_KeyDown(object sender, KeyEventArgs e)
        {
            if (txtDID.Text != "" && e.KeyCode == Keys.Enter)
            {
                dt = process.GetBaseDIDInfo(txtDID.Text.ToString());
                if (dt.Rows.Count > 0)
                {
                    txtTotalQty.Text = dt.Rows[0]["Qty"].ToString();
                    txtRealQty.Text = dt.Rows[0]["RealQty"].ToString();
                    Line = dt.Rows[0]["Line"].ToString();
                    Side = dt.Rows[0]["Side"].ToString();
                    txtUpdateTo.Focus();
                }
                else
                {
                    txtMsg.Text = "Can not find this DID, check it please !!";
                }
            }
        }

        private void txtReason_KeyDown(object sender, KeyEventArgs e)
        {
            if (txtReason.Text != "" && e.KeyCode == Keys.Enter)
            {
                btnSave_Click(sender, e);
            }
        }

        private void txtUpdateTo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue == 13)
            {
                if (txtUpdateTo.Text != "")
                {
                    if(Regex.IsMatch(txtUpdateTo.Text.ToString(), "^[0-9]*$"))
                    {
                        txtReason.Focus();
                    }
                    else
                    {
                        MessageBox.Show("请数入数字");
                        txtUpdateTo.Text = "";
                    }
                }
                else
                {
                    MessageBox.Show("数量不能为空");
                }
            }
        }

        private void frmUpdateReelQty_Load(object sender, EventArgs e)
        {

        }

        private void frmUpdateRealQty_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmUpdateRealQty");
        }
    }
}
