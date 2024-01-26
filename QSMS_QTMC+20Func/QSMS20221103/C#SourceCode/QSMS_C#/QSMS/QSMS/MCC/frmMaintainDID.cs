using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PrinterLib;
using Microsoft.Win32;

namespace QSMS.QSMS.MCC
{
    public partial class frmMaintainDID : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.MCC.MCCProcess MCC = new DbLibrary.MCC.MCCProcess();
        PrintLabel Print = new PrintLabel();
        private int CommandType = 0;
        private string UserRight = "N";
        private string RePrintDID = "N";
        private string strPrintPort = string.Empty;
        private string strCommSetting = string.Empty;
        private string strDIDPrintLabel = string.Empty;
        DataTable dt = new DataTable();

        public frmMaintainDID()
        {
            InitializeComponent();
        }

        private void frmMaintainDID_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmMaintainDID");
        }

        private void frmMaintainDID_Load(object sender, EventArgs e)
        {
            for (int i = Parameter.g_userRight.GetLowerBound(0); i <= Parameter.g_userRight.GetUpperBound(0); i++)
            {
                if (Parameter.g_userRight[i] == "DeleteDID")
                {
                    UserRight = "Y";
                }
                if (Parameter.g_userRight[i] == "RePrintDID")
                {
                    RePrintDID = "Y";
                }
            }
            strPrintPort = pubFunction.ConfigListGetValue("PrintPort");
            strCommSetting = pubFunction.ConfigListGetValue("CommSetting");
            strDIDPrintLabel = pubFunction.ConfigListGetValue("MaintainDIDLabel");
        }

        private void txtUNID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && txtUNID.Text != "")
            {
                try
                {
                    DataTable dtUNID = new DataTable();
                    string strErrMessage = string.Empty;
                    txtUNID.Text = txtUNID.Text.Replace("\r", "");
                    txtUNID.Text = txtUNID.Text.Replace("\n", "");
                    txtUNID.Text = txtUNID.Text.Replace(" ", "");
                    txtUNID.Text = txtUNID.Text.Trim().ToUpper();
                    if (txtUNID.Text.IndexOf(";") > 0)
                    {
                        dtUNID = null;
                        dtUNID = MCC.QSMS_GenUNID(txtUNID.Text, "");
                        if (dtUNID.Rows.Count > 0)
                        {
                            if (dtUNID.Rows[0]["Result"].ToString().ToUpper() == "OK")
                            {
                                txtCompPN.Text = dtUNID.Rows[0]["CompPN"].ToString().ToUpper();
                                txtVendorCode.Text = dtUNID.Rows[0]["VendorCode"].ToString().ToUpper();
                                txtDateCode.Text = dtUNID.Rows[0]["DateCode"].ToString().ToUpper();
                                txtLotCode.Text = dtUNID.Rows[0]["LotCode"].ToString().ToUpper();
                                txtUNID.Text = dtUNID.Rows[0]["UNID"].ToString().ToUpper();
                            }
                            else
                            {
                                MessageBox.Show(dt.Rows[0]["Msg"].ToString().ToUpper());
                                return;
                            }
                        }
                    }
                    strErrMessage = FunPartNumberCheck(txtCompPN.Text);
                    if (strErrMessage != "PASS")
                    {
                        MessageBox.Show(strErrMessage);
                        txtUNID.Text = "";
                        txtUNID.Focus();
                        return;
                    }
                    if (dt != null)
                    {
                        txtQty.Text = "";
                        txtQty.Focus();
                    }
                    else
                    {
                        txtVendorCode.Text = "";
                        txtVendorCode.Focus();
                    }
                    btnQuery_Click(sender, e);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
            }
        }

        private string FunPartNumberCheck(string PartNumber)
        {
            DataTable dtCheck = MCC.CheckFormat(PartNumber);
            if (dtCheck.Rows.Count > 0)
            {
                if (dtCheck.Rows[0]["ErrorCode"].ToString().ToUpper() == "0")
                {
                    return "PASS";
                }
                else
                {
                    return dtCheck.Rows[0]["Result"].ToString().ToUpper();
                }
            }
            else
            {
                return "FAIL";
            }
        }

        private void txtDateCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && txtDateCode.Text != "")
            {
                txtLotCode.Text = "";
                txtLotCode.Focus();
            }
        }

        private void txtLotCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && txtLotCode.Text != "")
            {
                txtQty.Text = "";
                txtQty.Focus();
            }
        }

        private void txtVendorCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && txtVendorCode.Text != "")
            {
                txtDateCode.Text = "";
                txtDateCode.Focus();
            }
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            dt = MCC.QSMS_MCC_QueryDataByType("MCC_GetQSMSDID_ByDID", "", "", txtCompPN.Text, txtUNID.Text, "");
            DG_Result.DataSource = dt;
            btnUpdate.Enabled = true;
            btnDel.Enabled = true;
            btnSave.Enabled = true;
            btnReprint.Enabled = true;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (pubFunction.IsNumeric(txtQty.Text.Trim(), "INT") == false)
            {
                MessageBox.Show("数量输入错误，必须为整数!", "提示");
                txtQty.Text = "";
                txtQty.Focus();
                return;
            }

            btnAdd.Enabled = false;
            btnUpdate.Enabled = true;
            btnDel.Enabled = true;
            btnSave.Enabled = true;
            btnCancel.Enabled = true;
            btnExit.Enabled = true;
            btnQuery.Enabled = true;
            txtCompPN.Enabled = true;
            txtVendorCode.Enabled = true;
            txtDateCode.Enabled = true;
            txtLotCode.Enabled = true;
            txtQty.Enabled = true;
            CommandType = 1;
            btnSave.Focus();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            txtCompPN.Text = "";
            txtVendorCode.Text = "";
            txtDateCode.Text = "";
            txtLotCode.Text = "";
            txtQty.Text = "";
            txtUNID.Text = "";
            txtGroupQty.Text = "";
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            if (dt != null)
            {
                pubFunction.doExport(dt);
            }
            else
            {
                MessageBox.Show("No Data", "提示");
                return;
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (UserRight == "N")
            {
                MessageBox.Show("你没有删除权限,请先维护(DeleteDID)!", "提示");
                return;
            }
            if (MessageBox.Show("确定要删除 QSMS系统中 该DID["+ txtUNID.Text + "]信息?", "提示", MessageBoxButtons.YesNo) == DialogResult.No)
            {
                return;
            }
            btnAdd.Enabled = true;
            btnUpdate.Enabled = true;
            btnDel.Enabled = true;
            btnSave.Enabled = true;
            btnCancel.Enabled = true;
            btnExit.Enabled = true;
            if (Parameter.BU == "")
            {
                MessageBox.Show("BU信息为空，请联系QMS！", "提示");
                return;
            }
            else
            {
                DataTable dtDel = MCC.DeleteDIDByBU(Parameter.BU, txtUNID.Text);
                if (dtDel.Rows.Count > 0)
                {
                    if (dtDel.Rows[0]["ErrorCode"].ToString() != "0")
                    {
                        MessageBox.Show(dtDel.Rows[0]["Result"].ToString(), "提示");
                        return;
                    }
                }
            }
            RefreshDg(txtCompPN.Text);
            txtVendorCode.Text = "";
            txtDateCode.Text = "";
            txtLotCode.Text = "";
            txtQty.Text = "";
            txtUNID.Text = "";

        }

        private void RefreshDg(string CompPN)
        {
            dt = MCC.QSMS_MCC_QueryDataByType("MCC_GetQSMSDID_ByCompPN", "", "", CompPN, "", "");
            DG_Result.DataSource = dt;
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            RefreshDg("");
        }

        private void btnReprint_Click(object sender, EventArgs e)
        {
            if (RePrintDID == "N")
            {
                MessageBox.Show("你没有删除权限,请先维护(RePrintDID) !", "提示");
                return;
            }
            if (txtUNID.Text == "")
            {
                MessageBox.Show("DID/UNID为空，请先输入!", "提示");
                return;
            }
            if (pubFunction.IsNumeric(txtQty.Text, "INT") == false)
            {
                MessageBox.Show("数量输入错误，必须为整数!", "提示");
                return;
            }
            DataTable dtDID = MCC.QSMS_MCC_QueryDataByType("IPQC_GetIPQCFlag", "", "", txtUNID.Text, "", "");
            if (dtDID.Rows.Count == 0)
            {
                MessageBox.Show("未找到该DID/UNID信息，请确认!", "提示");
                txtUNID.Text = "";
                txtUNID.Focus();
                return;
            }
            else if (dtDID.Rows[0]["FirstMachine"].ToString().Trim().ToUpper() == "RETURN" || dtDID.Rows[0]["FirstMachine"].ToString().Trim().ToUpper() == "CALLBACK")
            {
                MessageBox.Show("Can not do reprint, this DID has been " + dtDID.Rows[0]["FirstMachine"].ToString() + "ed !", "提示");
                txtUNID.Text = "";
                txtUNID.Focus();
                return;
            }
            PrintLabel();

        }

        private void PrintLabel()
        {
            string Msg = string.Empty;
            if (File.Exists(strDIDPrintLabel) == false)
            {
                MessageBox.Show("在路径[" + strDIDPrintLabel + "]没找到对应模板!", "提示");
                return;
            }
            StreamReader reader = new StreamReader(strDIDPrintLabel, Encoding.Default);
            string tmpPrintStr = reader.ReadToEnd();
            reader.Close();
            tmpPrintStr = tmpPrintStr.ToUpper();

            if (Print.LabelSetting(strCommSetting, strPrintPort, 1, ref Msg) == false)
            {
                MessageBox.Show(Msg, "提示");
                return;
            }
            //10001 begin
            DataTable dtPrint = new DataTable();
            if (Parameter.BU == "NB6")
            {
                dtPrint = MCC.QSMS_SaveCompPrintLog("GetMaintainDIDPrintInfo", txtCompPN.Text.Trim().ToUpper(), "", txtVendorCode.Text.Trim().ToUpper(),
                    txtDateCode.Text.Trim().ToUpper(), txtLotCode.Text.Trim().ToUpper(), Parameter.g_userName, "", txtUNID.Text.Trim().ToUpper(), txtQty.Text.Trim(), "", "", "", "");
            }
            else
            {
                dtPrint = MCC.QSMS_SaveCompPrintLog("GetMaintainDIDPrintInfo", txtCompPN.Text.Trim().ToUpper(), "", txtVendorCode.Text.Trim().ToUpper(),
                   txtDateCode.Text.Trim().ToUpper(), txtLotCode.Text.Trim().ToUpper(), Parameter.g_userName, "", txtUNID.Text.Trim().ToUpper(), txtQty.Text.Trim(), "", "", "");
            }

            //10001 end
            for (int i = 0; i < dtPrint.Rows.Count; i++)
            {
                DataTable dtPrinter = null;
                dtPrinter = dtPrint.Clone();
                dtPrinter.Clear();
                dtPrinter.ImportRow(dtPrint.Rows[i]);
                if (Print.Print(tmpPrintStr, dtPrinter, ref Msg) == false)
                {
                    MessageBox.Show(Msg, "提示");
                    return;
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string strErrMessage = string.Empty;
            string TempDID = string.Empty;
            string P2 = string.Empty;

            int RetryCnt;
            int P1;
            bool InsertDIDOk;
            btnAdd.Enabled = true;
            btnQuery.Enabled = true;
            btnUpdate.Enabled = true;
            btnDel.Enabled = true;
            btnSave.Enabled = true;
            btnCancel.Enabled = true;
            btnExit.Enabled = true;

            txtUNID.Text = txtUNID.Text.Replace("\r", "");
            txtUNID.Text = txtUNID.Text.Replace("\n", "");
            txtUNID.Text = txtUNID.Text.Replace(" ", "");

            if (pubFunction.IsNumeric(txtGroupQty.Text.Trim(), "INT") == false || pubFunction.IsNumeric(txtQty.Text.Trim(), "INT") == false)
            {
                MessageBox.Show("数量输入错误，必须为整数!", "提示");
                return;
            }
            txtVendorCode.Text = txtVendorCode.Text.Trim();
            if (txtVendorCode.Text.Length > 7)
            {
                MessageBox.Show("VendorCode的长度必须小于8!", "提示");
                return;
            }
            strErrMessage = FunPartNumberCheck(txtCompPN.Text);
            if (strErrMessage != "PASS")
            {
                MessageBox.Show(strErrMessage);
                txtCompPN.Text = "";
                txtCompPN.Focus();
                return;
            }
            if (txtVendorCode.Text == "" || txtDateCode.Text == "" || txtLotCode.Text == "" || int.Parse(txtQty.Text) > 60000)
            {
                MessageBox.Show("Verdorcode/datecode/lotcode为空,或者DID的数量超过60000!", "提示");
                txtCompPN.Text = "";
                txtCompPN.Focus();
                return;
            }
            if (ChkAVL(txtCompPN.Text, txtVendorCode.Text) == false)
            {
                MessageBox.Show("CompPN和VendorCode不匹配！请检查！", "提示");
                txtVendorCode.Text = "";
                txtVendorCode.Focus();
                return;
            }
            DataTable dtDID = MCC.QSMS_MCC_QueryDataByType("IPQC_GetIPQCFlag", "", "", txtUNID.Text, "", "");
            if (dtDID.Rows.Count > 0)
            {
                MessageBox.Show("该DID/UNID已经存在!", "提示");
                return;
            }

            if (txtVendorCode.Text.IndexOf("'") > -1 || txtDateCode.Text.IndexOf("'") > -1 || txtLotCode.Text.IndexOf("'") > -1)
            {
                MessageBox.Show("请检查输入VendorCode/DateCode/LotCode，不能有特殊字符['] !", "提示");
                return;
            }
            string TransDate = DateTime.Now.ToString("yyyyMMddHHmmss");
            try
            {
                switch (CommandType)
                {
                    case 1:
                        LockTheForm(false);
                        for (int i = 1; i <= int.Parse(txtGroupQty.Text); i++)
                        {
                            RetryCnt = 0;
                            InsertDIDOk = false;
                            while (InsertDIDOk == true && RetryCnt < 10)
                            {
                                TempDID = GetDID(txtCompPN.Text, TransDate);
                                dtDID = null;
                                dtDID = MCC.GenRegisterDID(TempDID, txtCompPN.Text, int.Parse(txtQty.Text), txtVendorCode.Text, txtDateCode.Text, txtLotCode.Text, "", "", Parameter.g_userName, TransDate, 0, "");
                                if (dtDID.Rows.Count > 0)
                                {
                                    P1 = int.Parse(dtDID.Rows[0]["RtnCode"].ToString());
                                    P2 = dtDID.Rows[0]["RtnMessage"].ToString().ToUpper();
                                    if (P1 < 0)
                                    {
                                        MessageBox.Show(P2);
                                        LockTheForm(true);
                                        btnReprint.Enabled = false;
                                        return;
                                    }
                                    else if (P1 > 0)
                                    {
                                        InsertDIDOk = true;
                                    }
                                }
                                RetryCnt = RetryCnt + 1;
                            }
                            txtUNID.Text = TempDID;
                            PrintLabel();
                        }
                        btnQuery_Click(sender, e);
                        LockTheForm(true);
                        btnReprint.Enabled = false;
                        break;
                    case 2:
                        if (txtUNID.Text == "")
                        {
                            MessageBox.Show("DID/UNID为空！");
                            txtUNID.Enabled = true;
                            txtUNID.Focus();
                            return;
                        }
                        if (ChkAVL(txtCompPN.Text, txtVendorCode.Text) == false)
                        {
                            txtVendorCode.Text = "";
                            txtVendorCode.Focus();
                            return;
                        }
                        TempDID = txtUNID.Text;
                        dtDID = null;
                        dtDID = MCC.QSMS_MCC_QueryDataByType("IPQC_GetIPQCFlag", "", "", txtUNID.Text, "", "");
                        if (dtDID.Rows.Count > 0)
                        {
                            if (dtDID.Rows[0]["RemainQty"].ToString() != dtDID.Rows[0]["Qty"].ToString())
                            {
                                MessageBox.Show("The DID is using,can not update");
                                return;
                            }
                        }
                        dtDID = null;
                        //10001 begin
                        if (Parameter.BU == "NB6")
                        {
                            dtDID = MCC.QSMS_SaveCompPrintLog("UpdateQSMS_DID", txtCompPN.Text, txtQty.Text, txtVendorCode.Text, txtDateCode.Text, txtLotCode.Text, Parameter.g_userName, "", txtUNID.Text, "", "", "", "", "");
                        }
                        else
                        {
                            dtDID = MCC.QSMS_SaveCompPrintLog("UpdateQSMS_DID", txtCompPN.Text, txtQty.Text, txtVendorCode.Text, txtDateCode.Text, txtLotCode.Text, Parameter.g_userName, "", txtUNID.Text, "", "", "", "");
                        }
                        //10001 end
                        if (txtUNID.Text == "")
                        {
                            txtUNID.Text = TempDID;
                        }
                        btnQuery_Click(sender, e);
                        break;
                }
                RefreshDg("");
                CommandType = 0;
                txtGroupQty.Text = "1";
                btnCancel_Click(sender, e);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            btnAdd.Enabled = true;
            btnUpdate.Enabled =  true;
            btnDel.Enabled = true;
            btnSave.Enabled = true;
            btnCancel.Enabled = true;
            btnExit.Enabled = true;
            btnQuery.Enabled = true;


            txtCompPN.Enabled = true;
            txtVendorCode.Enabled = true;
            txtDateCode.Enabled = true;
            txtLotCode.Enabled = true;
            txtQty.Enabled = true;
            CommandType = 2;
        }

        private bool ChkAVL(string CompPN, string VendorCode)
        {
            if (Parameter.Check_AVL != "Y")
            {
                return true;
            }
            else
            {
                DataTable dtChkAVL = MCC.QSMS_MCC_QueryDataByType("MCC_ChkAVL", "", "", CompPN, VendorCode, "");
                if (dt.Rows.Count == 0)
                {
                    return false;
                }
                return true;
            }
        }

        private void LockTheForm(bool lockCtl)
        {

            txtGroupQty.Enabled = lockCtl;
            txtCompPN.Enabled = lockCtl;
            txtVendorCode.Enabled = lockCtl;
            txtDateCode.Enabled = lockCtl;
            txtLotCode.Enabled = lockCtl;
            txtQty.Enabled = lockCtl;
            txtUNID.Enabled = lockCtl;
            btnQuery.Enabled = lockCtl;
            btnAdd.Enabled = lockCtl;
            btnUpdate.Enabled = lockCtl;
            btnDel.Enabled = lockCtl;
            btnSave.Enabled = lockCtl;
            btnCancel.Enabled = lockCtl;
            btnRefresh.Enabled = lockCtl;
            btnExit.Enabled = lockCtl;
            btnReprint.Enabled = lockCtl;
            DG_Result.Enabled = lockCtl;
        }

        private string GetDID(string CompPN,string TransDate)
        {
            DataTable dtGetDID = MCC.QSMS_MCC_QueryDataByType("MCC_GetDID", "", "", CompPN, "", "");
            return dtGetDID.Rows[0]["DID"].ToString().ToUpper();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btVenodrCode_Click(object sender, EventArgs e)
        {
            DataTable dt = MCC.QSMS_MCC_QueryDataByType("MCC_GetDIDByVendorCode", "", "", txtVendorCode.Text, "", "");
            if (dt != null)
            {
                pubFunction.doExport(dt);
            }
            else
            {
                MessageBox.Show("No Data", "提示");
                return;
            }
        }

        private void btnCommSave_Click(object sender, EventArgs e)
        {
            string SavePath = "HKEY_CURRENT_USER\\Software\\VB and VBA Program Settings\\SMT\\QSMS";
            string ReadPath = "Software\\VB and VBA Program Settings\\SMT\\QSMS";

            Registry.SetValue(SavePath, "Comm", TxtComm.Text.Trim(), RegistryValueKind.String);
            Registry.SetValue(SavePath, "CommPort", TxtCompPort.Text.Trim(), RegistryValueKind.String);
        }
    }
}
