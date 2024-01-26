using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using PrinterLib;
using Microsoft.Win32;


namespace QSMS.QSMS.MCC
{
    public partial class frmMaintainDIDAutoDispatch : Form
    {
        private DataTable dt;
        private DateTime dtime;
        private int timeSeq = 0;
        private string msg;
        private string dateCode;
        private string lotCode;
        private string vendorCode;
        private string str09Code;
        private string strDispatchType;
        private string strWOGroup;
        private string strPrintPort;
        private string strCommSetting;
        private string strCompPrintLabel;
        private string strLabelContent;
        private string strDIDPrintLabel;
        private string strUNIDPrintLabel;
        private string strDIDLabelContent;
        private string strISNUID;
        private bool ReturnDID = false;
        private string OldReturnDID = "";
        private int CommandType = 0;
        private string strFWImage = "";
        private string IsFWImage = "N";

        PrintLabel Print = new PrintLabel();
        DbLibrary.MCC.MCCProcess mccProcess = new DbLibrary.MCC.MCCProcess();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();

        public frmMaintainDIDAutoDispatch()
        {
            InitializeComponent();
        }

        private void frmMaintainDIDAutoDispatch_Load(object sender, EventArgs e)
        {
            OptNormal.Checked = true;
            strDispatchType = "0";
            strWOGroup = "";
            txtCompPN.Focus();
            optchgAction(false);
            strPrintPort = pubFunction.ConfigListGetValue("PrintPort");
            strCommSetting = pubFunction.ConfigListGetValue("CommSetting");
            strCompPrintLabel = pubFunction.ConfigListGetValue("CompPrintLabel");
            strDIDPrintLabel = pubFunction.ConfigListGetValue("DIDLabel");
            strUNIDPrintLabel = pubFunction.ConfigListGetValue("UNIDLabel");
            labInpectionNo.Visible = false;
            txtInspection.Visible = false;
            labelMSD.Visible = false;
            txtMSD.Visible = false;
            CompPrint.Visible = false;

            if (Parameter.StrBU == "AS") 
            { 
                labInpectionNo.Visible = true;
                txtInspection.Visible = true;
            }
            if (pubFunction.ConfigListGetValue("SCANMSD") == "Y") 
            {
                labelMSD.Visible = true;
                txtMSD.Visible = true;
            }
            if (pubFunction.ConfigListGetValue("DispatchCompPrint") == "Y")
            {
                CompPrint.Visible = true;
            }
            if (pubFunction.ConfigListGetValue("ChkOneByOneMaterial") == "N")
            {
                groupBox3.Visible = false;
                dgDMList.Visible = false;
                groupBox4.Top = groupBox3.Top;
                groupBox4.Height = groupBox4.Height + groupBox3.Height;
                dgDIDList.Top = dgDMList.Top;
                dgDIDList.Height = dgDIDList.Height + dgDMList.Height;
            }
        }

        private void OptNormal_CheckedChanged(object sender, EventArgs e)
        {
            if (OptNormal.Checked)
            {
                clearExtraData();
                optchgAction(false);
                strDispatchType = "0";
            }
        }

        private void optExtra_CheckedChanged(object sender, EventArgs e)
        {
            if (optExtra.Checked)
            {
                strDispatchType = "2";
                clearExtraData();
                optchgAction(true);
                dt = mccProcess.XL_GetAllWOInfoList("Line", "", "", "", "", "", txtCompPN.Text.ToString().Trim(), "");
                for(int i=0;i<dt.Rows.Count;i++)
                {
                    cmbLine.Items.Add(dt.Rows[i]["GroupValue"].ToString());
                }
            }
        }

        private void optSpecial_CheckedChanged(object sender, EventArgs e)
        {
            if (optSpecial.Checked)
            {
                strDispatchType = "3";
                clearExtraData();
                optchgAction(true);
                dt = mccProcess.XL_GetAllWOInfoList("Line", "", "", "", "", "", txtCompPN.Text.ToString().Trim(), "");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cmbLine.Items.Add(dt.Rows[i]["GroupValue"].ToString());
                }
            }
        }

        private void optToWO_CheckedChanged(object sender, EventArgs e)
        {
            if (optToWO.Checked)
            {
                strDispatchType = "5";
                clearExtraData();
                optchgAction(true);
                dt = mccProcess.XL_GetAllWOInfoList("Line", "", "", "", "", "", txtCompPN.Text.ToString().Trim(), "");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cmbLine.Items.Add(dt.Rows[i]["GroupValue"].ToString());
                }
            }
        }

        private void cmbLine_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbLR.Items.Clear();
            cmbMachine.Items.Clear();
            cmbSide.Items.Clear();
            cmbSlot.Items.Clear();
            cmbWO.Items.Clear();
            dt = mccProcess.XL_GetAllWOInfoList("WO", "", "", "", "", "", txtCompPN.Text.ToString().Trim(), cmbLine.Text.ToString());
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cmbWO.Items.Add(dt.Rows[i]["GroupValue"].ToString());
            }
        }

        private void cmbWO_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbLR.Items.Clear();
            cmbMachine.Items.Clear();
            cmbSide.Items.Clear();
            cmbSlot.Items.Clear();
            strWOGroup = "";
            DataSet ds = mccProcess.XL_GetAllWOInfoList_WO("Machine", cmbWO.Text.ToString().Trim(), "", "", "", "", txtCompPN.Text.ToString().Trim(), cmbLine.Text.ToString());
            dt = ds.Tables[0];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cmbMachine.Items.Add(dt.Rows[i]["GroupValue"].ToString());
            }
            dt = ds.Tables[1];
            strWOGroup = dt.Rows[0]["GroupID"].ToString();
        }

        private void cmbMachine_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbLR.Items.Clear();
            cmbSide.Items.Clear();
            cmbSlot.Items.Clear();
            dt = mccProcess.XL_GetAllWOInfoList("Side", cmbWO.Text.ToString().Trim(), cmbMachine.Text.ToString().Trim(), "", "", "", txtCompPN.Text.ToString().Trim(), cmbLine.Text.ToString());
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cmbSide.Items.Add(dt.Rows[i]["GroupValue"].ToString());
            }
        }

        private void cmbSide_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbLR.Items.Clear();
            cmbSlot.Items.Clear();
            dt = mccProcess.XL_GetAllWOInfoList("Slot", cmbWO.Text.ToString().Trim(), cmbMachine.Text.ToString().Trim(), cmbSide.Text.ToString().Trim(), "", "", txtCompPN.Text.ToString().Trim(), cmbLine.Text.ToString());
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cmbSlot.Items.Add(dt.Rows[i]["GroupValue"].ToString());
            }
        }

        private void cmbSlot_SelectedIndexChanged(object sender, EventArgs e)
        {
            cmbLR.Items.Clear();
            dt = mccProcess.XL_GetAllWOInfoList("LR", cmbWO.Text.ToString().Trim(), cmbMachine.Text.ToString().Trim(), cmbSide.Text.ToString().Trim(), cmbSlot.Text.ToString().Trim(), "", txtCompPN.Text.ToString().Trim(), cmbLine.Text.ToString());
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cmbLR.Items.Add(dt.Rows[i]["GroupValue"].ToString());
            }
        }

        private void optchgAction(bool value)
        {
            cmbLine.Enabled = value;
            cmbLR.Enabled = value;
            cmbMachine.Enabled = value;
            cmbSide.Enabled = value;
            cmbSlot.Enabled = value;
            cmbWO.Enabled = value;
        }


        private void LockTheForm(bool lockCtl)
        {
             OptComp.Enabled = lockCtl;
             OptPrint.Enabled = lockCtl;
             btnCommSave.Enabled = lockCtl;
             txtGroupQty.Enabled = lockCtl;
             txtCompPN.Enabled = lockCtl;
             CboVendorCode.Enabled = lockCtl;
             txtDateCode.Enabled = lockCtl;
             txtLotCode.Enabled = lockCtl;
             txtQty.Enabled = lockCtl;
             txtUNID.Enabled = lockCtl;
             btFind.Enabled = lockCtl;
             btnDel.Enabled = lockCtl;
             btsave.Enabled = lockCtl;
             btCancel.Enabled = lockCtl;
             btRefresh.Enabled = lockCtl;
             btExit.Enabled = lockCtl;
             btReprint.Enabled = lockCtl;
             dgDIDList.Enabled = lockCtl;
        }

        private void clearExtraData()
        {
            cmbLine.Text = "";
            cmbLR.Text = "";
            cmbMachine.Text = "";
            cmbSide.Text = "";
            cmbSlot.Text = "";
            cmbWO.Text = "";
            cmbLine.Items.Clear();
            cmbLR.Items.Clear();
            cmbMachine.Items.Clear();
            cmbSide.Items.Clear();
            cmbSlot.Items.Clear();
            cmbWO.Items.Clear();
            txtCompPN.Focus();
        }

        private void clearData()
        {
            txtUNID.Text = "";
            txtQty.Text = "";
            txtCompPN.Text = "";
            CboVendorCode.Text = "";
            txtDateCode.Text = "";
            txtLotCode.Text = "";
            txtMSD.Text = "";
            dateCode = "";
            lotCode = "";
            vendorCode = "";
            txtCompPN.Focus();
   
        }

        private void errorNotice(string msg)
        {
            pubFunction.Sound("Error");
            lblmsg.Text = msg;
            clearData();
            lblmsg.ForeColor = Color.Red;
        }

        //private bool isNumberic(string strValue)
        //{
        //    try
        //    {
        //        int result = Convert.ToInt32(strValue);
        //        return true;
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}

        private bool formatCheck()
        {
            if (optExtra.Checked == true || optSpecial.Checked == true || optToWO.Checked == true)
            {
                if (cmbLine.Text == "" || cmbLR.Text == "" || cmbMachine.Text == "" || cmbSide.Text == "" || cmbSlot.Text == "" || cmbWO.Text == "")
                {
                    errorNotice("Line/LR/Machine/Side/Slot/WO exists empty,please check !");
                    return false;
                }

                if (txtExtraQty.Text.ToString() == "")
                {
                    errorNotice("ExtraQty is empty !");
                    return false;
                }
                if (txtExtraQty.Text.ToString() == "0")
                {
                    errorNotice("ExtraQty can not be 0 !");
                    return false;
                }

                if (pubFunction.IsNumeric(txtExtraQty.Text.ToString(), "INT") == false)
                {
                    errorNotice("Extra dispatch qty is not numeric !");
                    return false;
                }

                if (Convert.ToInt32(txtExtraQty.Text.ToString()) > Convert.ToInt32(txtQty.Text.ToString()))
                {
                    errorNotice("Extra/Special dispatch qty bigger than did qty !");
                    return false;
                }
            }
            if (txtQty.Text.ToString() == "")
            {
                errorNotice("Qty is empty !");
                return false;
            }
            if (txtQty.Text.ToString() == "0")
            {
                errorNotice("Qty can not be 0 !");
                return false;
            }
            if (pubFunction.IsNumeric(txtQty.Text.ToString(), "INT") == false)
            {
                errorNotice("Qty is not numeric !");
                return false;
            }

            if (int.Parse(txtQty.Text.ToString()) > 120000)
            {
                errorNotice("Dispatch Qty is more than 120000,please check !");
                return false;
            }
            if (dateCode == "" || lotCode == "" || vendorCode == "")
            {
                errorNotice("DateCode/LotCode/VendorCode exists empty,please check !");
                return false;
            }
            if (vendorCode.Length > 7)
            {
                errorNotice("VendorCode length must be less than 8 !");
                return false;
            }
            
            return true;
        }

        private void txtComPN_KeyDown(object sender, KeyEventArgs e)
        {
           
            if (pubFunction.ConfigListGetValue("CheckScanner") == "Y")
            {
                DateTime dtimeNow = DateTime.Now;
                if (timeSeq == 0)
                {
                    dtime = dtimeNow;
                    timeSeq = 1;
                }
                dtimeNow = DateTime.Now;
                TimeSpan ts = dtimeNow - dtime;
                int t = ts.Milliseconds;
                if (t > 300)
                {
                    errorNotice("请使用刷枪作业");
                    timeSeq = 0;
                    return;
                }
            }
            if (e.KeyCode == Keys.Enter && txtCompPN.Text != "")
            {
                try
                {
                    if (pubFunction.ConfigListGetValue("CheckScanner") == "Y") 
                    {
                        DataTable dtChk2DCode = mccProcess.Check2DCode(txtCompPN.Text.Trim());
                        if (dtChk2DCode.Rows[0]["Result"].ToString() != "1")
                        {
                            errorNotice(dtChk2DCode.Rows[0]["ErrDesc"].ToString());
                            return;
                        } 
                    }
                    
                    if (CompPrint.Checked)
                    {
                        if (string.IsNullOrEmpty(strPrintPort) || string.IsNullOrEmpty(strCommSetting) || string.IsNullOrEmpty(strCompPrintLabel))
                        {
                            errorNotice("打印机端口,模板未设置");
                            return;
                        }
                        if (!File.Exists(strCompPrintLabel))
                        {
                            errorNotice("未找到模板:" + strCompPrintLabel);
                            return;
                        }
                    }
                    timeSeq = 0;
                    str09Code = "";
                    txtCompPN.Text = txtCompPN.Text.ToString().Trim().Replace(" ", "").Replace("\r", "").Replace("\t", "").Replace("\n", "");
                    if (txtCompPN.Text.ToString().IndexOf(';') > 0)
                    {
                        dt = mccProcess.getDID(txtCompPN.Text.ToString());
                        if (dt.Rows[0]["Result"].ToString() == "OK")
                        {
                            strISNUID = dt.Rows[0]["isUNID"].ToString();
                            txtCompPN.Text = dt.Rows[0]["CompPN"].ToString();
                            txtQty.Text = dt.Rows[0]["Qty"].ToString();
                            txtUNID.Text = dt.Rows[0]["UNID"].ToString();
                            CboVendorCode.Text = dt.Rows[0]["VendorCode"].ToString();
                            txtDateCode.Text = dt.Rows[0]["DateCode"].ToString();
                            txtLotCode.Text = dt.Rows[0]["LotCode"].ToString();
                            vendorCode = dt.Rows[0]["VendorCode"].ToString();
                            dateCode = dt.Rows[0]["DateCode"].ToString();
                            lotCode = dt.Rows[0]["LotCode"].ToString();

                            //Add 保存UniqueID 
                            if (pubFunction.ConfigListGetValue("SaveUNID") == "Y")
                            {
                                mccProcess.SaveUNID(txtCompPN.Text, dateCode,vendorCode,lotCode,txtQty.Text,txtUNID.Text,txtCompPN .Text);

                            }

                            if (dt.Rows[0]["Chk09Code"].ToString() == "Y")
                            {
                                string strCust09Code = dt.Rows[0]["Cust09Code"].ToString();
                                frmInput ipt09Code = new frmInput("09Code");
                                if (ipt09Code.ShowDialog() != DialogResult.OK)
                                {
                                    errorNotice("输入09Code错误");
                                    return;
                                }
                                str09Code = ipt09Code.strInPut;
                                DataTable dt09Code = mccProcess.QSMS_CheckDID(str09Code, strCust09Code, txtCompPN.Text.ToString(), Parameter.g_userName, "Check09Code");
                                if (dt09Code.Rows[0]["Result"].ToString() != "OK")
                                {
                                    errorNotice(dt.Rows[0]["Msg"].ToString());
                                    return;
                                }
                                if (str09Code.Length < 20)
                                {
                                    str09Code = "";
                                }
                            }
                            if (dt.Rows[0]["IsBSMaterial"].ToString() == "Y")
                            {
                                frmInput iptDatecode = new frmInput("DateCode");
                                if (iptDatecode.ShowDialog() != DialogResult.OK)
                                {
                                    errorNotice("输入DateCode错误");
                                    return;
                                }
                                dateCode = iptDatecode.strInPut;

                                frmInput iptLotCode = new frmInput("LotCode");
                                if (iptLotCode.ShowDialog() != DialogResult.OK)
                                {
                                    errorNotice("输入LotCode错误");
                                    return;
                                }
                                lotCode = iptLotCode.strInPut;
                            }
                        }
                        else
                        {
                           errorNotice(dt.Rows[0]["Msg"].ToString());
                            return;
                        }
                    }
                    else if (txtCompPN.Text.Trim().IndexOf("-") > 0 && txtCompPN.Text.Trim().Length > 15)
                    {
                        ReturnDID=true;
                        //OldDID = txtCompPN.Text.Trim();
                        dt=mccProcess.GetQSMS_DID_ToWH(txtCompPN.Text.Trim());
                        if (dt.Rows.Count > 0)
                        {
                            OldReturnDID = txtCompPN.Text.Trim();
                            txtCompPN.Text = dt.Rows[0]["compPN"].ToString().Trim();
                            CboVendorCode.Text = dt.Rows[0]["VendorCode"].ToString().Trim();
                            txtDateCode.Text = dt.Rows[0]["DateCode"].ToString().Trim();
                            txtLotCode.Text = dt.Rows[0]["LotCode"].ToString().Trim();
                            txtQty.Text = dt.Rows[0]["Qty"].ToString().Trim();    
                        }
                        else
                        {
                            MessageBox.Show("Can't find the information of this returnDID---" + txtCompPN.Text.Trim());
                            txtCompPN.Focus();
                            return;
                        }
                    }
                    else
                    {
                        errorNotice("请刷入正确的二维码BarCode");
                        return;
                    }

                    dt = mccProcess.CheckFormat("PartNumber", txtCompPN.Text.ToString().Trim());
                    if (dt.Rows[0]["ErrorCode"].ToString() == "1")
                    {
                        errorNotice(dt.Rows[0]["Result"].ToString());
                        return;
                    }
                    
                    if (pubFunction.ConfigListGetValue("CheckBCMS") == "Y")
                    {
                        if (mccProcess.CheckBCMS(txtCompPN.Text.ToString().Trim()) == false)
                        {
                            errorNotice("当前登录用户无发Bois材料的权限");
                            return;
                        }
                    }

                    if (pubFunction.ConfigListGetValue("OneByOneControl") == "Y")
                    {
                        if (mccProcess.CheckCompPN(txtCompPN.Text.ToString().Trim(),"") == true)
                        {
                            MessageBox.Show("请注意,此材料需一对一管控");
                        }
                    }

                    if (pubFunction.ConfigListGetValue("ChkVendorCode") == "Y")
                    {
                        DataTable dtChkVendorCode = mccProcess.CheckVendorCode(txtCompPN.Text.ToString().Trim(), CboVendorCode.Text.ToString().Trim());
                        
                        if (dtChkVendorCode.Rows[0]["Result"].ToString() != "1")
                        {
                            errorNotice(dtChkVendorCode.Rows[0]["ErrDesc"].ToString());
                            return;
                        }
                    }

                    dt = mccProcess.CheckNeedDispatch(OptNormal.Checked.ToString(), txtCompPN.Text.ToString().Trim(), Parameter.Factory);
                    if (dt.Rows[0]["Result"].ToString() != "1")
                    {
                        errorNotice(dt.Rows[0]["ErrDesc"].ToString());
                        return;
                    }
                    
                    if (pubFunction.ConfigListGetValue("CheckNeedMSD") == "Y")
                    {
                        if (mccProcess.CheckCompPN(txtCompPN.Text.ToString().Trim(), "IsNeedMSD") == true)
                        {
                            MessageBox.Show("这是MSD材料,请先烘烤");
                        }
                    }

                    if (pubFunction.ConfigListGetValue("ChkOneByOneMaterial") == "Y")
                    {
                        DataSet dss = mccProcess.XL_Dispatch_MaterialPrompt(txtCompPN.Text.ToString(), CboVendorCode.Text.Trim().ToString(), txtDateCode.Text.Trim().ToString(), txtLotCode.Text.Trim().ToString(), Parameter.Factory);
                        dt = dss.Tables[0];
                        if (dt.Rows[0]["Result"].ToString() == "0")
                        {
                            dt = dss.Tables[1];
                            if (dt.Rows.Count > 0)
                            {
                                dgDMList.DataSource = dt;
                                dgDMList.ClearSelection();
                            }
                        }
                        else
                        {
                            dgDMList.DataSource = null;
                            errorNotice(dt.Rows[0]["Description"].ToString());
                            return;
                        }
                    }

                    if (CheckQty.Checked)
                    {
                        btsave_Click(sender, e);
                    }
                    else
                    {
                        txtQty.Focus();
                    }
                }
                catch (Exception ex)
                {
                    pubFunction.Sound("Error");
                    lblmsg.Text = ex.Message.ToString();
                    lblmsg.BackColor = Color.Red;
                }
            }
        }

        private void txtQty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && txtQty.Text != "")
            {
                btsave_Click(sender, e);
            }
        }

        private void btsave_Click(object sender, EventArgs e)
        {
            string strBatch = string.Empty;
            string strErrMessage = string.Empty;
            string TempDID = string.Empty;
            string P2 = string.Empty;
            string BeginDID = string.Empty;
            string EndDID = string.Empty;
            string PreDIDPrinted = string.Empty;

            bool CHKAutoDispatchForAnotherBU = false;
            int RetryCnt;
            int P1;
            bool InsertDIDOk;
        
            try
            {
                if (formatCheck() == false)
                {
                    return;
                }

                if (Parameter.StrBU == "NB1") 
                {
                    if (mccProcess.CheckVendorPN(txtCompPN.Text.Trim().ToString())==true) 
                    {
                        string VendorPN = string.Empty;

                        frmInput iptVendorPN = new frmInput("对应厂商料号");
                        if (iptVendorPN.ShowDialog() != DialogResult.OK)
                        {
                            //errorNotice("输入厂商料号错误");
                            MessageBox.Show("输入厂商料号错误!");
                            return;
                        }
                        VendorPN = iptVendorPN.strInPut;
                        if (mccProcess.QSMS_CheckVendorPN(txtCompPN.Text.ToString(), VendorPN) == false)
                        {
                            //errorNotice("广达料号与厂商料号不对应!");
                            MessageBox.Show("广达料号与厂商料号不对应!");
                            return;  
                        }                       
                    }
                }

                if (pubFunction.ConfigListGetValue("CheckEMMCImageVersion") == "Y") 
                {
                    if (mccProcess.ChkEMMC(txtCompPN.Text.Trim().ToString()) == true) 
                    {
                        if (txtImgVersion.Text.Trim().ToString() == "")
                        {
                            //errorNotice("This CompPN is EMMC,please scan ImageVersion!");
                            //MessageBox.Show("This CompPN is EMMC,please scan ImageVersion!");
                            //pubFunction.Sound("Error");
                            MessageBox.Show("This CompPN is EMMC,please scan ImageVersion!");
                            lblmsg.Text = "This CompPN is EMMC,please scan ImageVersion!";
                            lblmsg.ForeColor = Color.Red;
                            return;  
                        }

                        if (optToWO.Checked == false) 
                        {
                            //errorNotice("This CompPN is EMMC,please choose ToWO dispatchtype!");
                            //MessageBox.Show("This CompPN is EMMC,please choose ToWO dispatchtype!");
                            //pubFunction.Sound("Error");
                            MessageBox.Show("This CompPN is EMMC,please choose ToWO dispatchtype!");
                            lblmsg.Text = "This CompPN is EMMC,please choose ToWO dispatchtype!";
                            lblmsg.ForeColor = Color.Red;
                            return;  
                        }

                       dt = mccProcess.GetWoinfoBasic(cmbWO.Text.Trim().ToString());
                       if (dt.Rows.Count > 0)
                       {
                           string MBPN = dt.Rows[0]["PN"].ToString();
                           if (mccProcess.QSMS_ChkEMMC(MBPN, txtImgVersion.Text.Trim().ToString()) == false) 
                           {
                                //errorNotice("This ImageVersion does not match the MBPN!");
                                //pubFunction.Sound("Error");
                                MessageBox.Show("This ImageVersion does not match the MBPN!");
                                lblmsg.Text = "This ImageVersion does not match the MBPN!";
                                lblmsg.ForeColor = Color.Red;
                                txtImgVersion.Text="";
                                return;  
                           }
                       }
                       else 
                       {
                           //errorNotice("Please choose WO!");
                           //return; 
                           //pubFunction.Sound("Error");
                           MessageBox.Show("Please choose WO!");
                           lblmsg.Text = "Please choose WO!";
                           lblmsg.ForeColor = Color.Red;
                           return;
                       }
                    }
                }

                if (pubFunction.ConfigListGetValue("CheckFWImage") == "Y")
                {
                    strFWImage = "";
                    IsFWImage = "N";
                    if (mccProcess.QSMS_ChkFWImage(txtCompPN.Text.Trim().ToString(), "FWImage") == true)
                    {
                        if (txtFWImage.Text.Trim().ToString() == "")
                        {
                            //MessageBox.Show("This CompPN is FW Image,please scan FWImage!");
                            //pubFunction.Sound("Error");
                            MessageBox.Show("This CompPN is FW Image,please scan FWImage!");
                            lblmsg.Text = "This CompPN is FW Image,please scan FWImage!";
                            lblmsg.ForeColor = Color.Red;
                            return;                        
                        }

                        if (optToWO.Checked == false)
                        {
                            //errorNotice("This CompPN is FW Image,please choose ToWO dispatchtype!");
                            //pubFunction.Sound("Error");
                            MessageBox.Show("This CompPN is FW Image,please choose ToWO dispatchtype!");
                            lblmsg.Text = "This CompPN is FW Image,please choose ToWO dispatchtype!";
                            lblmsg.ForeColor = Color.Red;                           
                            return;
                        }

                        dt = mccProcess.GetWoinfoBasic(cmbWO.Text.Trim().ToString());
                        if (dt.Rows.Count > 0)
                        {                          
                            strFWImage = mccProcess.QSMS_GetFWImage(cmbWO.Text.Trim().ToString());
                            if (strFWImage != "")
                            {                               
                                if (strFWImage!=txtFWImage.Text.Trim().ToString())
                                {
                                    //errorNotice("This FWImage does not match the WO!");
                                    //pubFunction.Sound("Error");
                                    MessageBox.Show("This FW Image does not match the WO!");
                                    lblmsg.Text = "This FW Image does not match the WO!";
                                    lblmsg.ForeColor = Color.Red;
                                    txtFWImage.Text = "";
                                    return;
                                }
                                IsFWImage = "Y";
                            }
                        }
                        else
                        {
                            //errorNotice("Please choose WO!");
                            //pubFunction.Sound("Error");
                            MessageBox.Show("Please choose WO!");
                            lblmsg.Text = "Please choose WO!";
                            lblmsg.ForeColor = Color.Red;
                            return;
                        }
                    }
                }

                if (ChkAVL(txtCompPN.Text.Trim(), CboVendorCode.Text.Trim()) == false)
                {
                    MessageBox.Show("CompPN and VendorCode not match!! please check ");
                    CboVendorCode.Text = "";
                    return;
                }

                if (Parameter.IC_CompChk == "Y") 
                {
                    dt = mccProcess.IC_CompNeedBurn(txtCompPN.Text.Trim());
                    if (dt.Rows[0]["Result"].ToString() == "0")
                    {
                        if (MessageBox.Show(dt.Rows[0]["Description"].ToString() + " DO you burn IC for it firstly!", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            return;
                        }
                    }
                }

                if (pubFunction.ConfigListGetValue("BatchControl") == "Y")
                {
                    dt = mccProcess.QSMS_ProcessCompBatch(txtCompPN.Text.Trim(), "", "", "", "CHECKBATCH");
                    if (dt.Rows[0]["result"].ToString().ToUpper() == "NEEDBATCH")
                    {
                        frmInput iptBatch = new frmInput(" the Batch of this CompPN ");
                        if (iptBatch.ShowDialog() != DialogResult.OK)
                        {
                            //errorNotice("Input Batch error");
                            MessageBox.Show("Input Batch error");
                            return;
                        }
                        strBatch = iptBatch.strInPut;
                        dt = mccProcess.QSMS_ProcessCompBatch(txtCompPN.Text.Trim(), "", strBatch, "", "CHECKBATCHVALUE");
                        if (dt.Rows[0]["result"].ToString().ToUpper() == "CHECKFAIL")
                        {
                            //errorNotice("Input Batch error, the Batch value must match with be defined!");
                            //return;
                            MessageBox.Show("Input Batch error, the Batch value must match with be defined!");
                            return;
                        }        
                    }
                }

                //if (Parameter.StrBU == "AS") 
                //{
                //    //CheckDataCode
                //    return;
                //}

                if (pubFunction.ConfigListGetValue("ChkDateCode") == "Y")
                {
                    if (CboVendorCode.Text.Trim() != "" && txtCompPN.Text.Trim() != "" && txtDateCode.Text.Trim() != "") 
                    {
                        dt = mccProcess.QSMS_ChkDateCodeSpecial(CboVendorCode.Text.Trim(), txtCompPN.Text.Trim(), txtDateCode.Text.Trim());
                        if (dt.Rows[0]["Result"].ToString().ToUpper() != "PASS")
                        {
                            MessageBox.Show(dt.Rows[0]["iMessage"].ToString());
                            return;
                        }    
                    }  
                }
                
                DataTable dtDID = mccProcess.XL_GetAllWOInfoList("QueryDID", "", "", "", "", "", "", "", txtUNID.Text.ToString());
                if (dtDID.Rows.Count > 0)
                {
                    errorNotice("DID:" + txtUNID.Text.ToString() + "已经存在");
                    return;
                }

                if (CboVendorCode.Text.IndexOf("'") > -1 || txtDateCode.Text.IndexOf("'") > -1 || txtLotCode.Text.IndexOf("'") > -1)
                {          
                    errorNotice("请检查输入VendorCode/DateCode/LotCode，不能有特殊字符['] !");
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
                                while (InsertDIDOk == false && RetryCnt < 10)
                                {
                                    TempDID = GetDID(txtCompPN.Text, TransDate);
                                    if (BeginDID == "")
                                    {
                                        BeginDID = TempDID;
                                    }
                                    EndDID = TempDID;

                                    if (ReturnDID == true && pubFunction.ConfigListGetValue("CheckMSDCallBack") == "Y")
                                    {

                                    }

                                    dtDID = null;
                                    dtDID = mccProcess.XL_DIDAutoDispatch(TempDID, txtCompPN.Text.ToString(), txtQty.Text.ToString(), CboVendorCode.Text.Trim(), txtDateCode.Text.Trim(), txtLotCode.Text.Trim(), txtInspection.Text.Trim(), "", Parameter.g_userName, strDispatchType, strWOGroup, cmbWO.Text.ToString(), cmbLine.Text.ToString(), cmbSide.Text.ToString(), cmbMachine.Text.ToString(), cmbSlot.Text.ToString(), cmbLR.Text.ToString(), str09Code);

                                    if (dtDID.Rows[0]["result"].ToString() != "1")
                                    {
                                        LockTheForm(true);
                                        btCancel_Click(sender, e);
                                        btExcel.Enabled = false;

                                        pubFunction.Sound("Error");
                                        lblmsg.Text = dt.Rows[0]["ErrDesc"].ToString();
                                        lblmsg.ForeColor = Color.Red;
                                        RefreshDg("", BeginDID, EndDID);
                                        return;
                                    }

                                    if (strDispatchType == "0")
                                    {
                                        if (dtDID.Rows[0]["DID"].ToString() != TempDID)
                                        {
                                            TempDID = dtDID.Rows[0]["DID"].ToString();
                                            CHKAutoDispatchForAnotherBU = true;
                                        }
                                    }


                                    DataTable dtPrint = mccProcess.XL_GetDidPrintInfo(TempDID, Parameter.g_factory);
                                    if (dtPrint.Rows.Count == 0)
                                    {
                                        LockTheForm(true);
                                        errorNotice("can not find the DID,Please check");
                                        return;
                                    }
                                    else
                                    {
                                        if (Parameter.StrBU == "ESBU")
                                        {
                                            if (dtPrint.Rows[0]["DateCodeFlag"].ToString().ToUpper() == "FAIL")
                                            {
                                                MessageBox.Show("DID DateCode > Defined DateCode,Please Check!");
                                            }
                                        }                    
                                    }

                                    InsertDIDOk = true;

                                    RetryCnt = RetryCnt + 1;
                                }

                                string NewcompFlag = "";
                                if (NewcompFlag == "Y" && Parameter.StrBU == "NB6") 
                                {
                                   dt = mccProcess.DIDTrace_SaveData(TempDID, "F");
                                }

                                txtUNID.Text = TempDID;
                                if (PreDIDPrinted != txtUNID.Text.Trim() && InsertDIDOk == true) 
                                {
                                    PrintLabel();
                                }


                                if (pubFunction.ConfigListGetValue("BatchControl") == "Y" && strBatch != "") 
                                {
                                    dt = mccProcess.QSMS_ProcessCompBatch("", TempDID, strBatch, Parameter.g_userName, "SAVEBATCH");                                   
                                }

                                if (pubFunction.ConfigListGetValue("DIDAutoOpen") == "Y")
                                {
                                    dt = mccProcess.QSMS_DIDAutoOpen(TempDID);
                                }

                                if (IsFWImage == "Y")
                                {
                                    dt = mccProcess.QSMS_SaveFWImage(TempDID, txtCompPN.Text.Trim(), cmbWO.Text.Trim(), "FW_Image", txtFWImage.Text.Trim(), Parameter.g_userName);
                                }


                                PreDIDPrinted = txtUNID.Text.Trim();
                            }

                            pubFunction.Sound("OK");
                            lblmsg.ForeColor = Color.Green;
                            clearData();
                            btFind_Click(sender, e);
                            LockTheForm(true);
                            btExcel.Enabled = false;
                            break;
                    }

                    RefreshDg("", BeginDID, EndDID);
                    CommandType = 0;
                    txtGroupQty.Text = "1";
                    btCancel_Click(sender, e);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }

            }
            
            catch (Exception ex)
            {
                pubFunction.Sound("Error");
                lblmsg.Text = ex.Message.ToString();
                lblmsg.BackColor = Color.Red;
            }
        }

        private void btCancel_Click(object sender, EventArgs e)
        {
            clearData();
        }

        private void btFind_Click(object sender, EventArgs e)
        {
            try
            {
                dt = mccProcess.XL_GetAllWOInfoList("refDIDList", "", "", "", "", "", txtCompPN.Text.ToString().Trim(), "",txtUNID.Text.ToString());
                dgDIDList.DataSource = dt;
                dgDIDList.ClearSelection();
                btnDel.Enabled = true;
                btsave.Enabled = true; 
                btExcel.Enabled = true;
            }
            catch (Exception ex)
            {
                pubFunction.Sound("Error");
                lblmsg.Text = ex.Message.ToString();
                lblmsg.BackColor = Color.Red;
            }

        }

        private void btExcel_Click(object sender, EventArgs e)
        {
            try
            {
                dt = mccProcess.XL_GetAllWOInfoList("refDIDList", "", "", "", "", "", txtCompPN.Text.ToString().Trim(), "", txtUNID.Text.ToString());
                pubFunction.doExport(dt);
            }
            catch (Exception ex)
            {
                pubFunction.Sound("Error");
                lblmsg.Text = ex.Message.ToString();
                lblmsg.BackColor = Color.Red;
            }
        }

        private void frmMaintainDIDAutoDispatch_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmMaintainDIDAutoDispatch");
        }

        private void btRefresh_Click(object sender, EventArgs e)
        {
            try
            {   RefreshDg("","","");
            }
            catch (Exception ex)
            {
                pubFunction.Sound("Error");
                lblmsg.Text = ex.Message.ToString();
                lblmsg.BackColor = Color.Red;
            } 
        }
    
        private void btnDel_Click(object sender, EventArgs e)
        {

        }

        private void btVenodrCode_Click(object sender, EventArgs e)
        {
            try
            {
                dt = mccProcess.QSMS_MCC_QueryDataByType("MCC_GetVendorCode", "", "", CboVendorCode.Text.Trim(), "", "");//need add MCC_GetVendorCode
                pubFunction.doExport(dt);
            }
            catch (Exception ex)
            {
                pubFunction.Sound("Error");
                lblmsg.Text = ex.Message.ToString();
                lblmsg.BackColor = Color.Red;
            }
        }

        private void btExit_Click(object sender, EventArgs e)
        {
            
        }

        private void RefreshDg(string CompPN, string BeginDID, string EndDID)
        {
            DataTable dtDIDList = mccProcess.QSMS_QueryDIDData(CompPN, BeginDID, EndDID);
            if (dtDIDList.Rows.Count > 0)
            {
                dgDIDList.DataSource = dtDIDList;
                dgDIDList.ClearSelection();
                lblCount.Text = "DIDQty: " + dtDIDList.Rows.Count.ToString();
            }
            else
            {
                lblCount.Text = "DIDQty: 0 ";
            }

        }

        private string FunPartNumberCheck(string PartNumber)
        {
            DataTable dtCheck = mccProcess.CheckFormat(PartNumber);
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

        private bool ChkAVL(string CompPN, string VendorCode)
        {
            if (Parameter.Check_AVL == "Y")
            {
                DataTable dtChkAVL = mccProcess.QSMS_MCC_QueryDataByType("MCC_ChkAVL", "", "", CompPN, VendorCode, "");
                if (dt.Rows.Count == 0)
                {
                    return false;
                }
                return true;
            }
            else
            {
                return true;
            }
        }

        private string GetDID(string CompPN, string TransDate)
        {
            DataTable dtGetDID = mccProcess.QSMS_MCC_QueryDataByType("MCC_GetDID", "", "", CompPN, "", "");
            return dtGetDID.Rows[0]["DID"].ToString().ToUpper();
        }

        private void PrintLabel()
        {
            PrinterLib.PrintLabel lblprint = new PrinterLib.PrintLabel();
            if (lblprint.LabelSetting(strCommSetting, strPrintPort, 1, ref msg) == false)
            {
                errorNotice(msg);
                return;
            }
            if (pubFunction.ConfigListGetValue("PrintDIDLabel") == "Y")
            {
                if (strISNUID == "Y")
                {
                    strDIDLabelContent = new StreamReader(strUNIDPrintLabel).ReadToEnd();
                }
                else
                {
                    strDIDLabelContent = new StreamReader(strDIDPrintLabel).ReadToEnd();
                }
                dt = mccProcess.GetPrintInfo(txtUNID.Text.ToString(), "");
                if (dt.Rows.Count > 0)
                {
                    if (lblprint.Print(strDIDLabelContent, dt, ref msg) == false)
                    {
                        errorNotice(msg);
                        return;
                    }
                }
            }

            if (CompPrint.Checked)
            {

                if (string.IsNullOrEmpty(strLabelContent))
                {
                    strLabelContent = new StreamReader(strCompPrintLabel).ReadToEnd();
                }
                dt = mccProcess.GetPrintInfo(txtUNID.Text.ToString(), "CompPrint");
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["Result"].ToString() != "OK")
                    {
                        errorNotice(dt.Rows[0]["Msg"].ToString());
                        return;
                    }
                    if (lblprint.Print(strLabelContent, dt, ref msg) == false)
                    {
                        errorNotice(msg);
                        return;
                    }
                }
            }
        }

        private string GetDIDLabelFile(string PrinterType, string PrintDpm, string LabelType)
        {
            string GetDIDLabelFile="";
            return GetDIDLabelFile;    
        }


        private void btnCommSave_Click(object sender, EventArgs e)
        {
            string SavePath = "HKEY_CURRENT_USER\\Software\\VB and VBA Program Settings\\SMT\\QSMS";
            string ReadPath = "Software\\VB and VBA Program Settings\\SMT\\QSMS";
            
            Registry.SetValue(SavePath, "Comm", TxtComm.Text.Trim(), RegistryValueKind.String);
            Registry.SetValue(SavePath, "CommPort", TxtCompPort.Text.Trim(), RegistryValueKind.String);
        }

        private void btReprint_Click(object sender, EventArgs e)
        {

        }
       

    }
}
