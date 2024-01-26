using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PrinterLib;
using System.IO;

namespace QSMS.QSMS.MCC
{
    public partial class frmDIDInteGration : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.MCC.MCCProcess MCC = new DbLibrary.MCC.MCCProcess();
        PrintLabel Print = new PrintLabel();
        DataTable dtQuery = new DataTable();

        private string strDID = string.Empty;
        private string PreDIDPrinted = string.Empty;
        private string strPrintPort = string.Empty;
        private string strCommSetting = string.Empty;
        private string strDIDPrintLabel = string.Empty;
        int NoEntireCompPNQty = 0;

        public frmDIDInteGration()
        {
            InitializeComponent();
        }

        private void frmDIDInteGration_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmDIDInteGration");
        }

        private void frmDIDInteGration_Load(object sender, EventArgs e)
        {
            DTPEndDate.Text = DateTime.Today.ToString("yyyy/MM/dd");
            DTPBeginDate.Text = DateTime.Today.AddDays(-1).ToString("yyyy/MM/dd");
            txtBeginT.Text = "0800";
            txtEndT.Text = "2000";
            groupboxLabel.Visible = false;
            Init();
            LockTheForm(true);
            strPrintPort = pubFunction.ConfigListGetValue("PrintPort");
            strCommSetting = pubFunction.ConfigListGetValue("CommSetting");
            strDIDPrintLabel = pubFunction.ConfigListGetValue("DIDLabel");
        }

        private void Init()
        {
            DataTable dt = MCC.QSMS_MCC_QueryDataByType("MCC_GetMachineLine", "", "", "", "", "");
            cboLine.Text = "";
            cboLine.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cboLine.Items.Add(dt.Rows[i]["Line"].ToString().ToUpper());
            }

            dt = null;
            dt = MCC.QSMS_MCC_QueryDataByType("QuerySite", "", "", "", "", "");
            cboFactory.Items.Clear();
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows.Count > 1)
                {
                    lblFactory.Visible = true;
                    cboFactory.Visible = true;
                }
                else
                {
                    lblFactory.Visible = false;
                    cboFactory.Visible = false;
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cboFactory.Items.Add(dt.Rows[i]["Factory"].ToString().ToUpper());
                }
            }

            cboSide.Text = "";
            cboSide.Items.Clear();
            cboSide.Items.Add("S");
            cboSide.Items.Add("C");
            cboSide.Items.Add("Q");
            cboSide.Items.Add("W");
        }

        private void LockTheForm(bool lockCtl)
        {
            txtCompPN.Enabled = lockCtl;
            txtVendorCode.Enabled = lockCtl;
            txtDateCode.Enabled = lockCtl;
            txtLotCode.Enabled = lockCtl;
            txtDID.Enabled = lockCtl;
            btnQuery.Enabled = lockCtl;
            btnSave.Enabled = lockCtl;
            btnCancel.Enabled = lockCtl;
            btnRefresh.Enabled = lockCtl;
            btnExit.Enabled = lockCtl;
            btnReprint.Enabled = lockCtl;
            DG_Result.Enabled = lockCtl;
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            if (pubFunction.IsNumeric(txtBeginT.Text, "DOUBLE") == false || pubFunction.IsNumeric(txtEndT.Text, "DOUBLE") == false)
            {
                MessageBox.Show("输入的时间格式[HHMM]不对!");
                return;
            }
            string BeginDate = Convert.ToDateTime(DTPBeginDate.Text).ToString("yyyyMMdd") + txtBeginT.Text;
            string EndDate = Convert.ToDateTime(DTPEndDate.Text).ToString("yyyyMMdd") + txtEndT.Text;

            if (cboFactory.Visible == true)
            {
                if (cboFactory.Text == "")
                {
                    MessageBox.Show("请选择Factory!");
                    return;
                }
            }
            dtQuery = MCC.QSMS_DIDIntegration("QUERY", txtCompPN.Text, "", cboLine.Text, cboSide.Text, "", "", "", cboFactory.Text, 0, 0, "", BeginDate, EndDate, Parameter.g_userName, "");
            DG_Result.DataSource = dtQuery;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            if (dtQuery != null)
            {
                pubFunction.doExport(dtQuery);
            }
            else
            {
                MessageBox.Show("No Data", "提示");
                return;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            ClearData();
        }

        private void ClearData()
        {
            txtCompPN.Text = "";
            txtVendorCode.Text = "";
            txtDateCode.Text = "";
            txtLotCode.Text = "";
            txtQty.Text = "";
            txtDID.Text = "";
            strDID = "";
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            btnQuery_Click(sender, e);
        }

        private void txtDID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && txtDID.Text != "")
            {
                DataTable dtDID = new DataTable();
                if (txtDID.Text.IndexOf(";") > 0)
                {
                    DataTable dt = MCC.QSMS_GenUNID(txtDID.Text, "");
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["Result"].ToString().ToUpper() == "OK")
                        {
                            txtDID.Text = dt.Rows[0]["UNID"].ToString().ToUpper();
                        }
                        else
                        {
                            lblMsg.Text = dt.Rows[0]["Msg"].ToString().ToUpper();
                            lblMsg.ForeColor = Color.Red;
                        }
                    }
                }
                if (strDID == "")
                {
                    lblMsg.Text = "";
                    dtDID = null;
                    dtDID = MCC.QSMS_MCC_QueryDataByType("IPQC_GetIPQCFlag", "", "", txtDID.Text, "", "");
                    if (dtDID.Rows.Count > 0)
                    {
                        txtCompPN.Text = dtDID.Rows[0]["CompPN"].ToString().ToUpper();
                        txtVendorCode.Text = dtDID.Rows[0]["VendorCode"].ToString().ToUpper();
                        txtDateCode.Text = dtDID.Rows[0]["DateCode"].ToString().ToUpper();
                        txtLotCode.Text = dtDID.Rows[0]["LotCode"].ToString().ToUpper();
                        cboLine.Text = dtDID.Rows[0]["Line"].ToString().ToUpper();
                        cboSide.Text = dtDID.Rows[0]["Side"].ToString().ToUpper();

                        dtDID = null;
                        dtDID = MCC.QSMS_MCC_QueryDataByType("MCC_GetNoEntireCompPNQty", "", "", txtCompPN.Text, "", "");
                        if (dtDID.Rows.Count > 0)
                        {
                            NoEntireCompPNQty = Convert.ToInt32(dtDID.Rows[0]["Qty"].ToString());
                        }
                        else
                        {
                            MessageBox.Show("Please Maintain Data in QSMS_NoEntireCompPNSetting(Mainmenu-->UpLoadBasicData-->QSMS_NoEntireCompPNSetting)!");
                            ClearData();
                            return;
                        }
                    }
                }
                dtDID = null;
                dtDID = MCC.QSMS_MCC_QueryDataByType("IPQC_GetIPQCFlag", "", "", txtDID.Text, "", "");
                if (dtDID.Rows.Count < 1)
                {
                    //MessageBox.Show("This DID is not exist,and Please MaintainDID at first !");
                    lblMsg.Text = "该DID不存在，请先定义！";
                    ClearData();
                    return;
                }
                else
                {
                    if (CheckValid(txtDID.Text.Trim().ToUpper()) == false)
                    {
                        ClearData();
                        return;
                    }
                    int DIDQty = 0;
                    if (txtQty.Text == "")
                    {
                        DIDQty = 0;
                    }
                    else
                    {
                        DIDQty = Convert.ToInt32(txtQty.Text.Trim());
                    }
                    txtQty.Text = (DIDQty + Convert.ToInt32(dtDID.Rows[0]["Qty"].ToString())).ToString();

                }
                strDID = strDID + txtDID.Text.Trim() + ";";

                lblMsg.Text = strDID;
                txtDID.Text = "";
                txtDID.Focus();
            }
        }

        private bool CheckValid(string DID)
        {
            try
            {
                DID = DID.Trim();
                if (DID.Substring(DID.Length - 1, 1) == ";")
                {
                    DID = DID.Substring(0, DID.Length - 1);
                }
                if (strDID.IndexOf(DID) > -1)
                {
                    //MessageBox.Show("Please do not input the same DID :" + DID);
                    lblMsg.Text = "请不要输入相同DID：" + DID;
                    return false;
                }
                string[] strInputDID = DID.Split(';');
                DataTable dtDID = MCC.QSMS_MCC_QueryDataByType("MCC_CheckNoEntireCompPNQty", "", "", txtCompPN.Text, NoEntireCompPNQty.ToString(), "");
                if (dtDID.Rows.Count < 1)
                {
                    MessageBox.Show(txtCompPN.Text + " ReelBaseQty必须大于系统定义的散料数量:" + NoEntireCompPNQty);
                    return false;
                }
                for (int i = 0; i <= strInputDID.GetUpperBound(0); i++)
                {
                    dtDID = null;
                    dtDID = MCC.QSMS_MCC_QueryDataByType("IPQC_GetIPQCFlag", "", "", strInputDID[i], "", "");
                    if (dtDID.Rows.Count < 1)
                    {
                        lblMsg.Text = strInputDID[i] + " 不存在!";
                        return false;
                    }
                    else
                    {
                        if (dtDID.Rows[0]["CompPN"].ToString().Trim().ToUpper() != txtCompPN.Text.Trim().ToUpper())
                        {
                            lblMsg.Text = strInputDID[i] + " 和第一个DID的CompPN不一致!";
                            return false;
                        }
                        if (Convert.ToInt32(dtDID.Rows[0]["Qty"].ToString().Trim()) > NoEntireCompPNQty && pubFunction.ConfigListGetValue("UnChkBaseReelQty") != "Y")
                        {
                            lblMsg.Text = strInputDID[i] + " 的Qty > " + NoEntireCompPNQty + "（系统中定义的散料数量）!";
                            return false;
                        }
                        if (dtDID.Rows[0]["Line"].ToString().Trim().ToUpper() != cboLine.Text.Trim().ToUpper())
                        {
                            lblMsg.Text = strInputDID[i] + " 和第一个DID的线别不一致!";
                            return false;
                        }
                        if (dtDID.Rows[0]["side"].ToString().Trim().ToUpper() != cboSide.Text.Trim().ToUpper())
                        {
                            lblMsg.Text = strInputDID[i] + " 和第一个DID的面别不一致!";
                            return false;
                        }
                    }
                    dtDID = null;
                    dtDID = MCC.QSMS_MCC_QueryDataByType("MCC_ChkDIDUse", "", "", strInputDID[i], "", "");
                    if (dtDID.Rows[0]["Result"].ToString() == "1")
                    {
                        lblMsg.Text = strInputDID[i] + " DID已被使用!";
                        return false;
                    }


                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "CheckValid");
                return false;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                bool InsertDIDOk;
                string TempDID = string.Empty;

                strDID = strDID.Trim();
                if (strDID.Substring(strDID.Length - 1, 1) == ";")
                {
                    strDID = strDID.Substring(0, strDID.Length - 1);
                }
                if (strDID == "" || cboLine.Text == "" || cboSide.Text == "")
                {
                    //MessageBox.Show("Please Input DID/Line/Side!");
                    lblMsg.Text = "请把 DID/Line/Side 信息输入完全!";
                    ClearData();
                    return;
                }

                string[] strInputDID = strDID.Split(';');
                if (strInputDID.GetUpperBound(0) < 1)
                {
                    MessageBox.Show("必须输入两个及其以上DID!");
                    ClearData();
                    return;
                }
                LockTheForm(false);
                InsertDIDOk = false;

                string TransDate = DateTime.Now.ToString("yyyyMMddHHmmss");
                TempDID = GetDID(txtCompPN.Text, TransDate);
                string VendorCode = string.Empty;
                if (txtVendorCode.Text.Trim().Length > 7)
                {
                    VendorCode = txtVendorCode.Text.Trim().Substring(0, 7);
                }
                else
                {
                    VendorCode = txtVendorCode.Text.Trim();
                }

                DataTable dt = MCC.QSMS_DIDIntegration("DIDIngration", txtCompPN.Text, TempDID, cboLine.Text, cboSide.Text, VendorCode, txtDateCode.Text.Trim(),
                    txtLotCode.Text.Trim(), Parameter.g_factory, Convert.ToInt32(txtQty.Text.Trim()),Convert.ToInt32(txtQty.Text.Trim()), strDID.Trim(), "", "", Parameter.g_userName, "");
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["result"].ToString() != "1")
                    {
                        lblMsg.Text = dt.Rows[0]["ErrDesc"].ToString();
                        dt = null;
                        dt = MCC.QSMS_DIDIntegration("RESTOREDISPATCH", txtCompPN.Text, TempDID, cboLine.Text, cboSide.Text, VendorCode, txtDateCode.Text.Trim(),
                    txtLotCode.Text.Trim(), Parameter.g_factory, Convert.ToInt32(txtQty.Text.Trim()), Convert.ToInt32(txtQty.Text.Trim()), strDID.Trim(), "", "", Parameter.g_userName, "");
                        if (dt.Rows.Count > 0)
                        {
                            LockTheForm(true);
                            ClearData();
                            lblMsg.Text = dt.Rows[0]["ErrDesc"].ToString();
                        }
                        else
                        {
                            LockTheForm(true);
                            ClearData();
                            lblMsg.Text = lblMsg.Text + "     Please  Call QMS !";
                            return;
                        }
                    }
                    else
                    {
                        DataTable dtdel = MCC.QSMS_MCC_QueryDataByType("MCC_DelQSMS_DID", "", "", TempDID, "", "");
                        if (dt.Rows.Count > 0)
                        {
                            if (dt.Rows[0]["Result"].ToString() == "1")
                            {
                                MessageBox.Show(dt.Rows[0]["Msg"].ToString());
                                return;
                            }
                        }
                        lblMsg.Text = "DID: " + TempDID + "  Dispatch success!";
                    }
                }
                else
                {
                    LockTheForm(true);
                    ClearData();
                    lblMsg.Text = "DID: " + TempDID + "  Dispatch  Fail,Please  Call QMS !";
                    return;
                }
                
                DataTable dtPrint = MCC.XL_GetDidPrintInfo(TempDID, Parameter.g_factory);
                if (dtPrint.Rows.Count == 0)
                {
                    MessageBox.Show("can not find the DID,Please check!");
                    LockTheForm(true);
                    ClearData();
                    return;
                }
                else
                {
                    InsertDIDOk = true;
                    txtDID.Text = TempDID;
                    if (PreDIDPrinted != txtDID.Text && InsertDIDOk == true)
                    {
                        PrintLabel(TempDID,txtQty.Text.Trim());
                    }
                    PreDIDPrinted = txtDID.Text;
                    LockTheForm(true);
                    ClearData();
                    lblMsg.Text = "DID: " + TempDID + " Dispatch  and print label success !";
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Save_Click");
                return;
            }

        }

        private string GetDID(string CompPN, string TransDate)
        {
            DataTable dtGetDID = MCC.QSMS_MCC_QueryDataByType("MCC_GetDID", "", "", CompPN, "", "");
            return dtGetDID.Rows[0]["DID"].ToString().ToUpper();
        }

        private void PrintLabel(string strDID, string strQty)
        {
            string Msg = string.Empty;

            if (File.Exists(strDIDPrintLabel) == false)
            {
                lblMsg.Text = "在路径[" + strDIDPrintLabel + "]没找到对应模板!";
                return;
            }
            StreamReader reader = new StreamReader(strDIDPrintLabel, Encoding.Default);
            string tmpPrintStr = reader.ReadToEnd();
            reader.Close();

            if (Print.LabelSetting(strCommSetting, strPrintPort, 1, ref Msg) == false)
            {
                MessageBox.Show(Msg);
                return;
            }
            DataTable dt = MCC.GetPrintInfo(strDID, "");
            if (dt.Rows.Count > 0)
            {
                if (Print.Print(tmpPrintStr, dt, ref Msg) == false)
                {
                    MessageBox.Show(Msg);
                    return;
                }
            }

        }

    }
}
