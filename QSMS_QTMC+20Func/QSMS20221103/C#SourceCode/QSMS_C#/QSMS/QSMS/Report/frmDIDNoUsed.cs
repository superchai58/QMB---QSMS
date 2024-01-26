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
    public partial class frmDIDNoUsed : Form
    {
        public frmDIDNoUsed()
        {
            InitializeComponent();
        }
        DbLibrary.Report.PDReport Report = new DbLibrary.Report.PDReport();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DataTable dt = new DataTable();
        DataSet ds = new DataSet();
        private void frmDIDNoUsed_Load(object sender, EventArgs e)
        {
            DataTable dt = Report.QueryDispatch();
            cboLine.Items.Clear();
            cboLine.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cboLine.Items.Add(dt.Rows[i]["Line"].ToString().Trim());
            }
            cboDateRange.Items.Clear();
            cboDateRange.Items.Add(">=3 and <5");
            cboDateRange.Items.Add(">=5 and <10");
            cboDateRange.Items.Add(">=10");
            if (Parameter.BU == "NB5")
            {
                txtTime1.Visible = true;
                txtTime2.Visible = true;
                label3.Visible = true;
                label4.Visible = true;
            }
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            string DateType = "";

            if (cboLine.Text == "")
            {
                MessageBox.Show("Please input Line !");
                return;
            }

            if (cboDateRange.Text.Trim() == "" && txtTime1.Text == "" && txtTime2.Text == "")
            {
                MessageBox.Show("Please input DateRange !");
                return;
            }

            if (Parameter.BU == "NB5" && cboDateRange.Text.Trim() == "" && (txtTime1.Text != "" && txtTime2.Text != ""))
            {
                DateType = "4";
                ds = Report.Query_NOUseDID(DateType, cboLine.Text.Trim(), txtTime1.Text.Trim(), txtTime2.Text.Trim());
            }
            else
            {
                if (cboDateRange.Text == ">=3 and <5")
                {
                    DateType = "1";
                }
                else
                {
                    if (cboDateRange.Text == ">=5 and <10")
                    {
                        DateType = "1";
                    }
                    else if (cboDateRange.Text == ">=10")
                    {
                        DateType = "3";
                    }
                }
                ds = Report.Query_NOUseDID(DateType, cboLine.Text.Trim(), txtTime1.Text.Trim(), txtTime2.Text.Trim());
            }
            dt = ds.Tables[0];
            if (dt.Rows.Count > 0)
            {
                dataGridView1.DataSource = dt.DefaultView;
            }
            else
            {
                dataGridView1.DataSource = null;
            }
        }
        private void btnExcel_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {
                if (Parameter.BU == "NB5")
                {
                    dt = ds.Tables[1];

                    //ReportToExcel(dt, ds);
                }
                else
                {
                    pubFunction.CopyToExcel(dataGridView1, "DIDNoUsed", true);
                    //pubFunction.doExport(dt);
                }
            }
            else
            {
                MessageBox.Show("NO DATA !");
            }
        }

        private void frmDIDNoUsed_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmDIDNoUsed");
        }
    }
}
