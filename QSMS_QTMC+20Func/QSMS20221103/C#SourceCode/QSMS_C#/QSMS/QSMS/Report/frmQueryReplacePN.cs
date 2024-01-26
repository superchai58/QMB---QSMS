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
    public partial class frmQueryReplacePN : Form
    {
        public frmQueryReplacePN()
        {
            InitializeComponent();
        }

        DbLibrary.Report.QueryReplacePN QueryReplacePN = new DbLibrary.Report.QueryReplacePN();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DataSet DS = new DataSet();

        private void frmQueryReplacePN_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmQueryReplacePN");
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtCompPN.Text.ToString().Trim()) || !string.IsNullOrEmpty(txtModel.Text.ToString().Trim()))
            {

                DS = QueryReplacePN.QuerySAPBOM(txtCompPN.Text.ToString().Trim(), txtModel.Text.ToString().Trim());
                dgSAP_BOM.DataSource = DS.Tables[0];
            }
            else
            {
                MessageBox.Show("ComPN或Model不能全为空!");
                return;
            }

        }

        private void txtCompPN_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && txtCompPN.Text != "")
            {
                txtModel.Focus();
            }
        }

        private void txtModel_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && txtModel.Text != "")
            {
                btnQuery_Click(sender, e);
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            if (DS.Tables[0].Rows.Count > 0)
            {
                pubFunction.doExport(DS.Tables[0]);
            }
            else
            {
                MessageBox.Show("没有数据需要导出");
                return;
            }

        }


    }
}
