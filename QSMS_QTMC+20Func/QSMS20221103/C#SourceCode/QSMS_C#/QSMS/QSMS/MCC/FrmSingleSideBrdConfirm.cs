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
    public partial class FrmSingleSideBrdConfirm : Form
    {
        DbLibrary.MCC.MCCProcess mccProcess = new DbLibrary.MCC.MCCProcess();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();

        public FrmSingleSideBrdConfirm()
        {
            InitializeComponent();
        }

        private void FrmSingleSideBrdConfirm_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("FrmSingleSideBrdConfirm");
        }

        //private void label9_Click(object sender, EventArgs e)
        //{

        //}
    }
}
