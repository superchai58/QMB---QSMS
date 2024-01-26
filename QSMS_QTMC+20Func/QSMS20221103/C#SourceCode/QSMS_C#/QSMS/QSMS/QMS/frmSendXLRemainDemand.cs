using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.QMS
{
    public partial class frmSendXLRemainDemand : Form
    {
        private DataSet ds = new DataSet();

        public frmSendXLRemainDemand()
        {
            InitializeComponent();
        }

        DbLibrary.Report.SendXLRemainDemand sendXLRemainDemand = new DbLibrary.Report.SendXLRemainDemand();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DataTable dt = new DataTable();

        private void frmSendXLRemainDemand_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmSendXLRemainDemand");
        }

        //private void btnExcel_Click(object sender, EventArgs e)
        //{
        //    if (dt.Rows.Count > 0)
        //    {
        //        pubFunction.doExport(dt);
        //    }
        //    else
        //    {
        //        MessageBox.Show("没有数据需要导出");
        //        return;
        //    }
        //}

        private void frmSendXLRemainDemand_Load(object sender, EventArgs e)
        {

        }

        private void btnQuery_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrEmpty(txtDate.Text.ToString()))
            {
                MessageBox.Show("请选择日期!");
                return;
            }
            if (string.IsNullOrEmpty(txtShift.Text.ToString()))
            {
                MessageBox.Show("请选择班别!");
                return;
            }
            if (string.IsNullOrEmpty(txtFactory.Text.ToString()))
            {
                MessageBox.Show("请输入厂区!");
                return;
            }

            ds = sendXLRemainDemand.QueryRemainDemand(txtDate.Text.ToString(), txtShift.Text.ToString(), txtFactory.Text.ToString());
            if (ds.Tables.Count == 3)
            {
                dgResult1.DataSource = ds.Tables[0];
                dgResult2.DataSource = ds.Tables[1];
                dgResult3.DataSource = ds.Tables[2];
            }
            else
            {
                MessageBox.Show("No Data");
            }


        }

        private void btnSend_Click(object sender, EventArgs e)
        {


            if (string.IsNullOrEmpty(txtDate.Text.ToString()))
            {
                MessageBox.Show("请选择日期!");
                return;
            }
            if (string.IsNullOrEmpty(txtShift.Text.ToString()))
            {
                MessageBox.Show("请选择班别!");
                return;
            }
            if (string.IsNullOrEmpty(txtFactory.Text.ToString()))
            {
                MessageBox.Show("请输入厂区!");
                return;
            }

            DataTable dt = new DataTable();

            dt = sendXLRemainDemand.SendRemainDemand(txtDate.Text.ToString(), txtShift.Text.ToString(), txtFactory.Text.ToString());
            if (dt.Rows[0]["Result"].ToString() == "0")
            {
                MessageBox.Show("Send OK");
            }
            else
            {
                MessageBox.Show("Send Fail");
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            //string[] sheetname = new string[3];
            //sheetname[0] = "Result1";
            //sheetname[1] = "Result2";
            //sheetname[2] = "Result3";

            if (ds.Tables[0].Rows.Count > 0)
            {
                ds.Tables[0].TableName = "Result1";
                ds.Tables[1].TableName = "Result2";
                ds.Tables[2].TableName = "Result3";

                string[] Names = new string[3] { "Result1", "Result2", "Result3" };

                pubFunction.ExportDataSetToExcel(ds, Names);
            }
            else
            {
                MessageBox.Show("没有数据需要导出");
                return;
            }
        }

    }
}
