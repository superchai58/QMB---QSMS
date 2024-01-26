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
    public partial class frmTransferFujiNexim : Form
    {
        private DataTable dt;
        DbLibrary.MCC.MCCProcess MCC = new DbLibrary.MCC.MCCProcess();
        public frmTransferFujiNexim()
        {
            InitializeComponent();
        }
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            if (txtFactory.Text == "" || txtJobGr.Text == ""||txtBuidT.Text==""||txtLine.Text==""||txtRev.Text==""||txtSide.Text=="")
            {
                MessageBox.Show("Factory,Line,JobGroup,Version,BuildType,Side,都不可以为空，请确认！！", "提示");
                return;
            }
            BtnLoad.Enabled = false;
            labMsg.Text = "Uploading...,please wait a moment,thanks";
            dt = MCC.QSMS_InsertMEBom_Nexim(txtFactory.Text,txtLine.Text,txtJobGr.Text,txtRev.Text,txtBuidT.Text,txtSide.Text,Parameter.g_userName.Trim());
            labMsg.Text = dt.Rows[0]["Description"].ToString();
            BtnLoad.Enabled = true;
        }

        private void frmTransferFujiNexim_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmTransferFujiNexim");
        }
    }
}
