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
    public partial class frmQueryDIDNeedCut : Form
    {
        DbLibrary.Report.QueryDIDNeedCutProcess process = new DbLibrary.Report.QueryDIDNeedCutProcess();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        public frmQueryDIDNeedCut()
        {
            InitializeComponent();
        }

        private void cmdGetFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "*.xls|*.xlsx";
            dialog.FilterIndex = 0;
            dialog.RestoreDirectory = true;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = dialog.FileName;
            }
            cboSheetName.Items.Clear();
            string[] sheet = pubFunction.GetExcelSheetName(txtFilePath.Text);
            if (sheet.Length > 0)
            {
                pubFunction.BindComboBox(cboSheetName, sheet);
            }
            cboSheetName.Enabled = true;

        }

        private void txtPN_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (!(txtPN.Text!=null && txtPN.Text!=""))
                {
                    MessageBox.Show("Please Input PN!");
                    return;
                }
                System.Data.DataTable dt = process.GetByPN(txtPN.Text.ToString());
                if (dt.Rows.Count == 0)
                {
                    txtPN.Text = "";
                    MessageBox.Show("Can not find this Componet PN in system!");
                    txtPN.Focus();
                    return;
                }
                if (CheckExistsInPNList(txtPN.Text.ToString()) == false)
                {
                    lstPN.Items.Add(txtPN.Text.ToString());
                }
                txtPN.Text = "";
                txtPN.Focus();
            }
        }
        
        private void frmQueryDIDNeedCut_Load(object sender, EventArgs e)
        {
            dtFrom.Value = DateTime.Now.AddDays(-1);
            dtTo.Value = DateTime.Now.AddDays(1);
        }

        private void cmdLoad_Click(object sender, EventArgs e)
        {
            string compPN = string.Empty;
            DataTable dt=pubFunction.GetDataFromExcel(txtFilePath.Text.ToString(),cboSheetName.Text.ToString());
            foreach(DataRow dr in dt.Rows)
            {
                compPN=dr[0].ToString();
                if (CheckExistsInPNList(compPN) == false)
                {
                    DataTable dt1=process.GetByPN(compPN);
                    if (dt1.Rows.Count > 0)
                    {
                        lstPN.Items.Add(compPN);
                    }
                    else
                    {
                        MessageBox.Show("Can not find this PN:"+compPN);
                    }
                }
            }
        }

        private bool CheckExistsInPNList(string PN)
        {
            for (int i = 0; i < lstPN.Items.Count; i++)
            {
                if (lstPN.Items[i].ToString() == PN)
                {
                    return true;
                }
            }
            return false;
        }

        private void cmdClear_Click(object sender, EventArgs e)
        {
            lstPN.Items.Clear();
        }

        private void cmdQuery_Click(object sender, EventArgs e)
        {
            if (lstPN.Items.Count == 0)
            {
                MessageBox.Show("Please input component PN");
                return;
            }
            string strPNList="";
            for(int i = 0; i < lstPN.Items.Count; i++)
            {
                strPNList += "'" + lstPN.Items[i].ToString() + "',";
            }
            strPNList = "(" +strPNList.Substring(0,strPNList.Length-1)+ ")";
            string SDate=dtFrom.Value.ToString("yyyyMMddhhmmss");
            string EDate = dtTo.Value.ToString("yyyyMMddhhmmss");
            DataTable dt=process.QueryData(strPNList,SDate,EDate);
            dataGridView1.DataSource = dt;
        }

        private void cmdExist_Click(object sender, EventArgs e)
        {
            pubFunction.RemoveForm("frmQueryDIDNeedCut");
            this.Close();
        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            if (lstPN.Items.Count == 0)
            {
                MessageBox.Show("Please input component PN");
                return;
            }
            string strPNList = "";
            for (int i = 0; i < lstPN.Items.Count; i++)
            {
                strPNList += "'" + lstPN.Items[i].ToString() + "',";
            }
            strPNList = "(" + strPNList.Substring(0, strPNList.Length - 1) + ")";
            string SDate = dtFrom.Value.ToString("yyyyMMddhhmmss");
            string EDate = dtTo.Value.ToString("yyyyMMddhhmmss");
            DataTable dt = process.QueryDataToExcel(strPNList, SDate, EDate);
            pubFunction.doExport(dt);
        }

        private void frmQueryDIDNeedCut_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmQueryDIDNeedCut");
        }
    }
}
