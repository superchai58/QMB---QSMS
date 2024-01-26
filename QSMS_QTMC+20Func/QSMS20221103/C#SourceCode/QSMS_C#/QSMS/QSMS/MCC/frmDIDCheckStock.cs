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
    public partial class frmDIDCheckStock : Form
    {
        DbLibrary.MCC.DIDCheckStockProcess process = new DbLibrary.MCC.DIDCheckStockProcess();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        public static string FuncType="";
        public static DataTable rstCompPN = null;
        public frmDIDCheckStock()
        {
            InitializeComponent();
        }
        
        private void frmDIDCheckStock_Load(object sender, EventArgs e)
        {
            dtpSDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
            dtpEDate.Text = DateTime.Now.AddDays(1).ToString("yyyy/MM/dd");

            if(FuncType.ToUpper()== "AUTOCHK")
            {
                txtRefID.Enabled = false;
                chkStock.Enabled = false;
                chkStock.CheckState = CheckState.Checked;
                RefershData(rstCompPN);
            }
            else
            {
                txtRefID.Enabled = true;
                chkStock.Enabled = true;
                chkStock.CheckState = CheckState.Unchecked;
                lblMsg.Text = "";
            }
        }

        private void RefershData(DataTable dt)
        {
            if (chkStock.Checked == false)
            {
                label4.Text = "DID ToWH Info";
                dataGrid2.DataSource = dt;
                dataGrid2.Refresh();
            }
            else
            {
                label4.Text = "CompPN Not OK";
                DataView dv1 = dt.DefaultView;
                dv1.RowFilter = "IsToWH='Y'";
                dataGrid1.DataSource = dv1.ToTable();
                dataGrid1.Refresh();
                DataView dv2 = dt.DefaultView;
                dv2.RowFilter = "IsToWH<>'Y'";
                dataGrid2.DataSource = dv2.ToTable();
                dataGrid2.Refresh();
            }
        }

        private void txtRefID_KeyPress(object sender, KeyPressEventArgs e)
        {
            DataSet ds = null;
            if (e.KeyChar == 13)
            {
                if (txtRefID.Text == "")
                {
                    MessageBox.Show("请输入RefID!");
                    return;
                }
                string SDate = dtpSDate.Value.ToString("yyyyMMdd");
                string EDate = dtpEDate.Value.ToString("yyyyMMdd");
                if (Int32.Parse(SDate) < Int32.Parse(EDate))
                {
                    if (chkStock.Checked == false)
                    {
                        lblMsg.Text = "ReferenceID:" +txtRefID.Text+" ToWH info is below:";
                        ds = process.XL_DIDChkStockByRefID(txtRefID.Text, "", 0);
                    }
                    else
                    {
                        ds = process.XL_DIDChkStockByRefID(txtRefID.Text, Parameter.g_userName, 1);
                    }
                    if (ds.Tables.Count > 0)
                    {
                        if (ds.Tables[0].Rows[0]["Result"].ToString() != "0")
                        {
                            MessageBox.Show(ds.Tables[0].Rows[0]["Description"].ToString());
                            return;
                        }
                        lblMsg.Text = ds.Tables[0].Rows[0]["Description"].ToString();
                        RefershData(ds.Tables[1]);
                    }
                }
                else
                {
                    MessageBox.Show("The StartDate must be smaller than EndDate !");
                    return;
                }
                txtRefID.Text = "";
            }
        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            DataSet ds = null;
            string SDate = dtpSDate.Value.ToString("yyyyMMdd");
            string EDate = dtpEDate.Value.ToString("yyyyMMdd");
            if (Int32.Parse(SDate) < Int32.Parse(EDate))
            {
                if (chkStock.Checked == false)
                {
                    //if (txtRefID.Text != "")
                    //{
                        lblMsg.Text = "ReferenceID:" + txtRefID.Text + " ToWH info is below:";
                        ds = process.XL_DIDChkStockByRefID(txtRefID.Text, "", 0);
                        if (ds.Tables.Count > 0)
                        {
                            if (ds.Tables[0].Rows[0]["Result"].ToString() != "0")
                            {
                                MessageBox.Show(ds.Tables[0].Rows[0]["Description"].ToString());
                                return;
                            }
                            try
                            {
                                pubFunction.doExport(ds.Tables[1]);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                                return;
                            }
                        }
                    //}
                }
                else
                {
                    return;
                }
            }
            else
            {
                MessageBox.Show("The StartDate must be smaller than EndDate !");
                return;
            }
        }

        private void cmdQuery_Click(object sender, EventArgs e)
        {
            DataSet ds = null;
            string SDate = dtpSDate.Value.ToString("yyyyMMdd");
            string EDate = dtpEDate.Value.ToString("yyyyMMdd");
            if (Int32.Parse(SDate) < Int32.Parse(EDate))
            {
                ds = process.XL_CheckRefID(SDate,EDate, "MCC_CheckRefID");
                if (ds.Tables.Count > 0)
                {
                    if (ds.Tables[0].Rows[0][0].ToString() != "0")
                    {
                        MessageBox.Show(ds.Tables[0].Rows[0][1].ToString());
                        return;
                    }
                    label4.Text = "ReferenceID Not OK";
                    label3.Text = "ReferenceID is OK";
                    DataTable dt1 = ds.Tables[1];
                    DataTable dt2 = ds.Tables[1];
                    DataView dv1 = dt1.DefaultView;
                    DataView dv2 = dt2.DefaultView;
                    dv2.RowFilter = "IsPass<>'Y'";
                    dataGrid2.DataSource = dv2.ToTable();
                    dataGrid2.Refresh();
                    dv1.RowFilter = "IsPass='Y'";
                    dataGrid1.DataSource = dv1.ToTable();
                    dataGrid1.Refresh();
                    
                }
            }
            else
            {
                MessageBox.Show("The StartDate must be smaller than EndDate !");
                return;
            }
            }

        private void frmDIDCheckStock_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmDIDCheckStock");
        }
    }
}
