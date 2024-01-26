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
    public partial class frmInput : Form
    {
        public string strInPut = string.Empty;
        public frmInput(string title)
        {
            InitializeComponent();
            lblmsg.Text = "请输入:"+title;
        }

        private void txtInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && txtInput.Text!="")
            {
                strInPut = txtInput.Text.ToString();
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }


    }
}
