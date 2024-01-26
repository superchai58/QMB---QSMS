using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.Report
{
    public partial class frmQueryCheckBOM : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.Report.QueryCheckBOMProcess QueryCheckBOM = new DbLibrary.Report.QueryCheckBOMProcess();
        public frmQueryCheckBOM()
        {
            InitializeComponent();
        }

        private void Command1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FrmQueryCheckBOM_Load(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dtpSDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
            dtpEDate.Text = DateTime.Now.AddDays(1).ToString("yyyy/MM/dd");
            CboLine.Items.Clear();
            dt = QueryCheckBOM.GetLine();
            if(dt.Rows.Count>0)
            {
                for(int i=0;i<dt.Rows.Count;i++)
                {
                    CboLine.Items.Add(dt.Rows[i]["Line"].ToString());
                }
               
            }
            dt = null; 
            dt = QueryCheckBOM.QSMS_QueryCheckBOM("","","","");
            DG1.DataSource = dt;
        }

        private void CmdQuery_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            string BeginDate="", EndDate="";
            BeginDate = dtpSDate.Value.ToString("yyyyMMdd");
            EndDate = dtpEDate.Value.ToString("yyyyMMdd");
            dt = QueryCheckBOM.QSMS_QueryCheckBOM(TxtWO.Text, CboLine.Text,BeginDate, EndDate);
            DG1.DataSource = dt;
        }

        private void frmQueryCheckBOM_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmQueryCheckBOM");
        }
    }
}
