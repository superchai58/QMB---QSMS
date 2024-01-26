using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.IPQC
{
    public partial class frmInRelieve : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.IPQC.IPQCProcess IPQC = new DbLibrary.IPQC.IPQCProcess();
        
        public frmInRelieve()
        {
            InitializeComponent();
        }

        private void btnRelieve_Click(object sender, EventArgs e)
        {
            string strDID = txtDID.Text.Replace("\n", "").Replace("\r", "").Trim();
            if (string.IsNullOrEmpty(strDID))
            {
                MessageBox.Show("Please input DID!!!");
                txtDID.Focus();
                return;
            }
            DialogResult dr = MessageBox.Show("请核对DID信息是否正确？", "提示信息", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (dr == DialogResult.Yes)
            {
                IPQC.InRelieve(strDID, "", "", "", "", "", "", "", 0, "", "", "Relieve");
            }
            txtDID.Text = "";
            txtDID.Focus();
        }

        private void frmInRelieve_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmInRelieve");
        }
    }
}
