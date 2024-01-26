using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace QSMS.QSMS.Report
{
    public partial class frmBeforeHandCheckBom : Form
    {
        public frmBeforeHandCheckBom()
        {
            InitializeComponent();
        }
        DbLibrary.Report.BeforeHandCheckBom BeforeHandCheckBom = new DbLibrary.Report.BeforeHandCheckBom();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DataTable dt = new DataTable();

        private void frmBeforeHandCheckBom_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmBeforeHandCheckBom");
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {
                pubFunction.doExport(dt);
            }
            else
            {
                MessageBox.Show("没有数据需要导出");
                return;
            }
        }

        private void frmBeforeHandCheckBom_Load(object sender, EventArgs e)
        {
            dt = BeforeHandCheckBom.GetLine();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cboLine.Items.Add(dt.Rows[i]["Line"]);
                }
            }
            dt.Clear();
            dt = BeforeHandCheckBom.GetModel();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cboModel.Items.Add(dt.Rows[i]["Model"]);
                }
            }
            dt.Clear();
            cboFactory.Items.Add("F1");
            cboFactory.Items.Add("F2");
            cboFactory.Items.Add("F3");
            cboFactory.Items.Add("QB");
            cboFactory.Items.Add("QC");
        }

        private void btnCheckBom_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(cboLine.Text.ToString()))
            {
                MessageBox.Show("请选择线别!");
                return;
            }
            if (string.IsNullOrEmpty(cboModel.Text.ToString()))
            {
                MessageBox.Show("请选择机种!");
                return;
            }
            if (string.IsNullOrEmpty(cboFactory.Text.ToString()))
            {
                MessageBox.Show("请选择厂区!");
                return;
            }
            Regex rex = new Regex(@"^\d+$");
            if (string.IsNullOrEmpty(txtCQty.Text.ToString()) || !rex.IsMatch(txtCQty.Text.ToString()))
            {
                MessageBox.Show("请输入整数数量!");
                txtCQty.Focus();
                return;
            }
            string PN = string.Empty;
            string Rev = string.Empty;
            if (cboModel.Text.IndexOf("-") > -1)
            {
                PN = cboModel.Text.Substring(0, cboModel.Text.ToString().IndexOf("-"));
                Rev = cboModel.Text.Substring(cboModel.Text.IndexOf("-") + "-".Length);
            }
            dt = BeforeHandCheckBom.CheckOP();
            if (dt.Rows.Count > 0)
            {
                MessageBox.Show("当前有人正在checking Bom,请过几分钟后再试!");
                return;
            }
            dt.Clear();
            dt = BeforeHandCheckBom.beforehandCheckBom(cboFactory.Text, cboLine.Text, PN, Rev, Convert.ToInt32(txtCQty.Text));
            if (dt.Rows.Count > 0)
            {
                dt.Clear();
                dt = BeforeHandCheckBom.CheckBomFail();
                dgResult.DataSource = dt;
                MessageBox.Show("CheckBom Fail!");
                return;
            }
            else
            {
                dgResult.DataSource = dt;
                MessageBox.Show("CheckBom OK!");
                return;
            }



        }

        private void txtCQty_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && !string.IsNullOrEmpty(txtCQty.Text.ToString()))
            {
                btnCheckBom_Click(sender, e);
            }
        }
    }
}

