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
    public partial class frmStartSplitLineMC : Form
    {
        public frmStartSplitLineMC()
        {
            InitializeComponent();
        }
        DbLibrary.SpecialCase.SpecialCaseProcess SpecialCase = new DbLibrary.SpecialCase.SpecialCaseProcess();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        private void CmdStartSplitLineMC_Click(object sender, EventArgs e)
        {
            SpecialCase.QSMS_SplitLineMC(Parameter.g_userName);
            MessageBox.Show("已启动分仓!");
        }

        private void frmStartSplitLineMC_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmStartSplitLineMC");
        }
    }
}
