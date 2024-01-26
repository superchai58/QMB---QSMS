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
    public partial class frmInputBox : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        public frmInputBox()
        {
            InitializeComponent();
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                if (textBox1.Text != "")
                {
                    frmMCCPreMaterial.PN = textBox1.Text;
                }
            }
        }

        private void frmInputBox_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmInputBox");
        }
    }
}
