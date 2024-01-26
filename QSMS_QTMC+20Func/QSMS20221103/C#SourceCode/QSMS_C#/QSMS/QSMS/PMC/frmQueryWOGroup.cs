using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.PMC
{
    public partial class frmQueryWOGroup : Form
    {
        DbLibrary.PMC.QueryWOGroupProcess process = new DbLibrary.PMC.QueryWOGroupProcess();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DataTable dt = null;
        DataTable dtUnClosed = null;
        DataTable dtClosed = null;

        public frmQueryWOGroup()
        {
            InitializeComponent();
        }

        private void cmdQuery_Click(object sender, EventArgs e)
        {
            //if (!(cboLine.Text!=null || cboLine.Text!=""))
            if (string.IsNullOrEmpty(cboLine.Text.Trim()))
            {
                MessageBox.Show("Please input Line");
                return;
            }
            string beginTime = dtpSDate.Value.ToString("yyyyMMdd");
            string endTime = dtpEDate.Value.ToString("yyyyMMdd");
            dtUnClosed = null;
            dtClosed = null;
            if (txtWO.Text.Length == 9)
            {
                dtUnClosed = process.GetDataByWO("UnClosedByWO", txtWO.Text.ToString(), cboLine.Text.ToString());
                DGNotFinished.DataSource = dtUnClosed;
                dtClosed = process.GetDataByWO("ClosedByWO", txtWO.Text.ToString(), cboLine.Text.ToString());
                DGFinish.DataSource = dtClosed;
                return;
            }
            dtUnClosed = process.GetData("UnClosedWO", beginTime, endTime, cboLine.Text.ToString());
            DGNotFinished.DataSource = dtUnClosed;
            dtClosed = process.GetData("", beginTime, endTime, cboLine.Text.ToString());
            DGFinish.DataSource = dtClosed;
        }

        private void frmQueryWOGroup_Load(object sender, EventArgs e)
        {
            dtpSDate.Value = DateTime.Now;
            dtpEDate.Value = DateTime.Now;

            dt=process.GetLine();
            foreach(DataRow dr in dt.Select())
            {
                cboLine.Items.Add(dr["line"]);
            }
        }
        
        private void cmdUnClosed_Click(object sender, EventArgs e)
        {
            pubFunction.doExport(dtUnClosed);
        }

        private void cmdClosed_Click(object sender, EventArgs e)
        {
            pubFunction.doExport(dtClosed);
        }

        private void frmQueryWOGroup_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmQueryWOGroup");
        }
    }
}
