using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.PD
{
    public partial class frmChecKDID : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.PD.PDProcess PD = new DbLibrary.PD.PDProcess();
        DataTable dtResult = new DataTable();
        private string strGroupID = string.Empty;
        private string strCheckDIDByLine = string.Empty;
        private string strChkDID = string.Empty;

        public frmChecKDID()
        {
            InitializeComponent();
        }
        private void frmChecKDID_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmChecKDID");
        }

        private void frmChecKDID_Load(object sender, EventArgs e)
        {
            strCheckDIDByLine = pubFunction.ConfigListGetValue("CheckDIDByLine");
            strChkDID = pubFunction.ConfigListGetValue("ChkDID");

            DTPEndDate.Text = DateTime.Today.ToString("yyyy/MM/dd");
            DTPBeginDate.Text = DateTime.Today.AddDays(-1).ToString("yyyy/MM/dd");
            txtBeginT.Text = "0800";
            txtEndT.Text = "2000";

            lblSide.Visible = false;
            cboSide.Visible = false;
            lblGroupID.Visible = false;
            cboGroupID.Visible = false;
            lblTabel.Visible = false;
            cboTabel.Visible = false;
            lblLine.Visible = false;
            cboLine.Visible = false;

            if (strCheckDIDByLine == "Y")
            {
                lblLine.Visible = true;
                cboLine.Visible = true;
                GetLine();
            }

            if (strChkDID == "Y")
            {
                lblSide.Visible = true;
                cboSide.Visible = true;
                cboSide.Items.Clear();
                cboSide.Items.Add("");
                cboSide.Items.Add("S");
                cboSide.Items.Add("C");
                lblGroupID.Visible = true;
                cboGroupID.Visible = true;
                lblTabel.Visible = true;
                cboTabel.Visible = true;
                cboTabel.Items.Clear();
                cboTabel.Items.Add("");
                cboTabel.Items.Add("Tabel1");
                cboTabel.Items.Add("Tabel2");
                cboTabel.Items.Add("Tabel3");
                cboTabel.Items.Add("Tabel4");
            }

        }

        private void GetLine()
        {
            DataTable dt = PD.QSMS_PD_QueryDataByType("PD_GetLine", "", "", "", "", "");
            cboLine.Text = "";
            cboLine.Items.Clear();
            cboLine.Items.Add("");
            cboLine.Items.Add("ALL");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cboLine.Items.Add(dt.Rows[i]["Line"].ToString().ToUpper());
            }
        }

        private void cboLine_DropDownClosed(object sender, EventArgs e)
        {
            strGroupID = "";
            if (strChkDID == "Y")
            {
                DataTable dt = PD.QSMS_PD_QueryDataByType("PD_GetGroupID", "", "", cboLine.Text, "", "");
                cboGroupID.Items.Clear();
                cboGroupID.Items.Add("");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cboGroupID.Items.Add(dt.Rows[i]["GroupID"].ToString().ToUpper());
                }
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            pubFunction.doExport(dtResult);
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            if (pubFunction.IsNumeric(txtBeginT.Text, "DOUBLE") == false || pubFunction.IsNumeric(txtEndT.Text, "DOUBLE") == false)
            {
                MessageBox.Show("输入的时间格式不对!");
                return;
            }
            string BeginDate = Convert.ToDateTime(DTPBeginDate.Text).ToString("yyyyMMdd") + txtBeginT.Text;
            string EndDate = Convert.ToDateTime(DTPEndDate.Text).ToString("yyyyMMdd") + txtEndT.Text;

            if (BeginDate != "" || EndDate != "" || txtDID.Text != "")
            {
                if (Parameter.BU == "NB5")
                {
                    dtResult = PD.QSMS_CHeckDID_NB5(txtDID.Text, txtBarCode.Text, "Query", BeginDate, EndDate, cboLine.Text, cboGroupID.Text, cboSide.Text, Parameter.g_userName, cboTabel.Text);
                }
                else
                {
                    dtResult = PD.QSMS_CHeckDID(txtDID.Text, txtBarCode.Text, "Query", BeginDate, EndDate, cboLine.Text);
                }
                DG_Result.DataSource = dtResult;
            }
            else
            {
                MessageBox.Show("查询条件不能全为空(BeginDate/EndDate/DID)!");
                return;
            }
        }

        private void txtDID_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter && txtDID.Text != "" && txtDID.Text.Trim().Length <= 30)
            if (e.KeyCode == Keys.Enter && txtDID.Text != "")
            {
                DataTable dt = null;
                if (txtDID.Text.IndexOf(";") > 0)
                {
                    dt = PD.QSMS_GenUNID(txtDID.Text, "");
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["Result"].ToString().ToUpper() == "OK")
                        {
                            txtDID.Text = dt.Rows[0]["UNID"].ToString().ToUpper();
                        }
                        else
                        {
                            MessageBox.Show(dt.Rows[0]["Msg"].ToString().ToUpper());
                            return;
                        }
                    }
                }
                if (Parameter.BU == "NB5")
                {
                    if (txtDID.Text.Trim().IndexOf("'") > 0)
                    {
                        pubFunction.Sound("ERROR");
                        MessageBox.Show(txtDID.Text + "具有特殊字符(')");
                        txtBarCode.Enabled = false;
                        txtDID.Text = "";
                        txtDID.Focus();
                        return;
                    }
                }
                if (strCheckDIDByLine == "Y")
                {
                    if (cboLine.Text == "" && strChkDID != "Y")
                    {
                        MessageBox.Show("请先选择线别!");
                        return;
                    }
                    if (strChkDID == "Y" && (cboGroupID.Text != "" || cboLine.Text != "" || cboSide.Text != "" || cboTabel.Text != ""))
                    {
                        dt = null;
                        if (Parameter.BU == "NB5")
                        {
                            dt = PD.QSMS_CHeckDID_NB5(txtDID.Text, txtBarCode.Text, "CheckDID", "", "", cboLine.Text, cboGroupID.Text, cboSide.Text, Parameter.g_userName, cboTabel.Text);
                        }
                        else
                        {
                            dt = PD.QSMS_CHeckDID(txtDID.Text, txtBarCode.Text, "CheckDID", "", "", cboLine.Text);
                        }
                        if (dt.Rows.Count > 0)
                        {
                            if (dt.Rows[0]["Result"].ToString() == "1")
                            {
                                pubFunction.Sound("ERROR");
                                MessageBox.Show(dt.Rows[0]["Desc"].ToString());
                            }
                            else
                            {
                                pubFunction.Sound("OK");
                            }
                            txtBarCode.Enabled = false;
                            txtDID.Text = "";
                            txtDID.Focus();
                            return;
                        }
                    }
                }
                txtBarCode.Focus();
            }
            //if (txtDID.Text.Trim().Length > 30)
            //{
            //    MessageBox.Show("请输入正确的DID!");
            //    txtBarCode.Text = "";
            //    txtDID.Text = "";
            //    txtDID.Focus();
            //    return;
            //}
        }

        private void txtBarCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && txtBarCode.Text != "")
            {
                DataTable dt = new DataTable();
                if (Parameter.BU == "NB5")
                {
                    dt = PD.QSMS_CHeckDID_NB5(txtDID.Text, txtBarCode.Text, "Conf", "", "", cboLine.Text, cboGroupID.Text, cboSide.Text, Parameter.g_userName, cboTabel.Text);
                }
                else
                {
                    dt = PD.QSMS_CHeckDID(txtDID.Text, txtBarCode.Text, "Conf", "", "", cboLine.Text);
                }
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["Result"].ToString() == "0")
                    {
                        if (pubFunction.ConfigListGetValue("ChkDID") != "Y" && strCheckDIDByLine == "Y" && cboLine.Text != "ALL")
                        {
                            if (strGroupID == "")
                            {
                                strGroupID = dt.Rows[0]["GroupID"].ToString();
                            }
                            else
                            {
                                if (strGroupID.Trim().ToUpper() != dt.Rows[0]["GroupID"].ToString().Trim().ToUpper())
                                {
                                    pubFunction.Sound("ERROR");
                                    MessageBox.Show("相邻DID的GroupID不一致!");
                                    txtBarCode.Text = "";
                                    txtDID.Text = "";
                                    txtDID.Focus();
                                    return;
                                }
                            }
                        }
                        txtBarCode.Text = "";
                        txtDID.Text = "";
                        txtDID.Focus();
                        pubFunction.Sound("OK");
                    }
                    else
                    {
                        pubFunction.Sound("ERROR");
                        MessageBox.Show(dt.Rows[0]["Desc"].ToString());
                        txtBarCode.Text = "";
                        txtDID.Text = "";
                        txtDID.Focus();
                    }
                }
            }
        }
    }
}
