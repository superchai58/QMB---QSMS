using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
namespace QSMS.QSMS.IPQC
{
    public partial class frmQuery : Form
    {
        DataTable dt = new DataTable();
        DataSet ds = new DataSet();
        DbLibrary.IPQC.IPQCProcess WorkStation = new DbLibrary.IPQC.IPQCProcess();
        BrLibrary.PublicFunction CopyExecl = new BrLibrary.PublicFunction();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        public frmQuery()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Regex reg = new Regex("^[0-9]*$");

            if (txtSTime.Text.Length > 6)
            {
                MessageBox.Show("超出了时间长度！！时间格式：时时分分秒秒");
                txtSTime.Clear();
                txtSTime.Focus();
                return;
            }
            else if (!reg.IsMatch(txtSTime.Text.ToString()))
            {
                MessageBox.Show("请输入数字！！");
                txtSTime.Clear();
                txtSTime.Focus();
                return;
            }
            if (txtETime.Text.Length > 6)
            {
                MessageBox.Show("超出了时间长度！！时间格式：时时分分秒秒");
                txtETime.Clear();
                txtETime.Focus();
                return;
            }
            else if (!reg.IsMatch(txtETime.Text.ToString()))
            {
                MessageBox.Show("请输入数字！！");
                txtETime.Clear();
                txtETime.Focus();
                return;
            }
            string SDateTime = dtSDate.Value.ToString("yyyyMMdd") + txtSTime.Text.Trim();
            string EDateTime = dtEDate.Value.ToString("yyyyMMdd") + txtETime.Text.Trim();
            dtGridView.DataSource = null;
            dt = WorkStation.QueryInSpect(textDID.Text.Trim().ToString(), SDateTime.Trim(), EDateTime.Trim());
            if (dt.Rows.Count != 0)
            {
                dtGridView.DataSource = dt;
                textDID.Clear();
                textDID.Focus();
            }
            else
            {
                textDID.Clear();
                textDID.Focus();
                MessageBox.Show("没有查询到数据！！");
                return;
            }
        }

        private void frmQuery_Load(object sender, EventArgs e)
        {
            this.dtSDate.Value = DateTime.Now.AddDays(-1);
            this.txtSTime.Text = "000000";
            this.txtETime.Text = "235900";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bool y = false;
            y = CopyExecl.CopyToExcel(dtGridView, "sheet1", true);
            if (y == true)
            {
                MessageBox.Show("数据导入Execl成功！！");
                this.dtGridView.DataSource = null;
                this.dtGridView.Rows.Clear();
                return;
            }
            else
            {
                MessageBox.Show("数据导入Execl失败！！");
                return;
            }
        }

        private void frmQuery_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmQuery");
        }
    }
}
