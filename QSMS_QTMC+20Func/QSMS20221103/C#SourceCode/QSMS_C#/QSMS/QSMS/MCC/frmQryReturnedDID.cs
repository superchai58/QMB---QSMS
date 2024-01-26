using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.MCC
{
    public partial class frmQryReturnedDID : Form
    {
        public frmQryReturnedDID()
        {
            InitializeComponent();
        }
        DbLibrary.MCC.MCCProcess MCC = new DbLibrary.MCC.MCCProcess();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DataSet ds = new DataSet();

        private void frmQryReturnedDID_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmQryReturnedDID");
        }

        private void txtDID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && !string.IsNullOrEmpty(txtDID.Text.ToString()))
            {
                btnQuery_Click(sender, e);
            }
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtDID.Text.ToString()))
            {
                ds = MCC.QSMS_ReturnQry("ReturnDID", txtDID.Text.ToString());
            }
            else
            {
                MessageBox.Show("请输入ReturnDID");
                txtDID.Focus();
                return;
            }
            if (ds.Tables[0].Rows.Count > 0)
            {
                if (ds.Tables[0].Rows[0]["Result"] != "0")
                {
                    MessageBox.Show(ds.Tables[0].Rows[0]["Description"].ToString());
                    return;
                }
                else 
                {
                    dgReturnedDID.DataSource = ds.Tables[1];
                }

            }
            else
            {
                MessageBox.Show("没有找到该DID:"+txtDID.Text.ToString()+"数据");
                txtDID.Focus();
                return;
            }
        }

        
        private void txtNewDID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && !string.IsNullOrEmpty(txtNewDID.Text.ToString()))
            {
                btnQueryA_Click(sender, e);
            }
        }
        private void btnQueryA_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtNewDID.Text.ToString()))
            {
                ds = MCC.QSMS_ReturnQry("NewDID", txtNewDID.Text.ToString());
            }
            else 
            {
                MessageBox.Show("请输入NewDID");
                txtNewDID.Focus();
                return;
            }
            if (ds.Tables[0].Rows.Count > 0)
            {
                if (ds.Tables[0].Rows[0]["Result"].ToString()!="0")
                {
                    MessageBox.Show(ds.Tables[0].Rows[0]["Description"].ToString());
                    return;
                }
                else
                {
                    dgReturnedDID.DataSource = ds.Tables[1];
                }

            }
            else
            {
                MessageBox.Show("没有找到该DID:" + txtNewDID.Text.ToString() + "数据");
                txtNewDID.Focus();
                return;
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            pubFunction.CopyToExcel(dgReturnedDID,"ReturnDID", true);
        }

    }
}
