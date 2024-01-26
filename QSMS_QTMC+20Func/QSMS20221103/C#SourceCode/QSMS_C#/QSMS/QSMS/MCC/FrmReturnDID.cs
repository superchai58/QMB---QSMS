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
using QSMS.QSMS.MCC;
using Microsoft.Win32;
//20210128 -- 001  ju
//20210201 -- 002  ju
//20210202 -- 003  ju
namespace QSMS.QSMS.MCC
{
    public partial class FrmReturnDID : Form
    {
        public FrmReturnDID()
        {
            InitializeComponent();
        }
        DbLibrary.MCC.MCCProcess Process = new DbLibrary.MCC.MCCProcess();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.PD.PDProcess PD = new DbLibrary.PD.PDProcess();

        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        DataTable print = new DataTable();
        DataTable PrintData = new DataTable();
        private string strCheckScaner = "";
        private string IsAnotherBUDID = "";
        private string PreGroupID = "";
        private string TempDID = "";
        private DateTime dtime;
        private int timeSeq = 0;
        private string msg;
        private string IsNewWay;//是否新的退料方式


        private string strPrintPort;//新增flag设置打印属性


        private string strCommSetting;//新增flag设置打印属性


        private string PrintReturnDIDLabel;//新增flag设置打印路径及模板名称


        private string PrintReturnLabel;//新增flag设置打印路径及模板名称


        private string strUNIDPrintLabel;//新增flag设置打印路径及模板名称


        private string strLabelContent;
        private string strDIDLabelContent;
        private void FrmReturnDID_Load(object sender, EventArgs e)
        {
            dtpSDate.Text = DateTime.Now.ToShortDateString();
            dtpEDate.Text = DateTime.Now.ToShortDateString();
            strPrintPort = pubFunction.ConfigListGetValue("PrintPort");
            strCommSetting = pubFunction.ConfigListGetValue("CommSetting");
            PrintReturnDIDLabel = Application.StartupPath + "\\" + pubFunction.ConfigListGetValue("PrintReturnDIDLabel");
            PrintReturnLabel = Application.StartupPath + "\\" + pubFunction.ConfigListGetValue("PrintReturnLabel");

            //strCompPrintLabel = pubFunction.ConfigListGetValue("CompPrintLabel");
            //PrintReturnDIDLabel = pubFunction.ConfigListGetValue("DIDLabel");
            strUNIDPrintLabel = pubFunction.ConfigListGetValue("UNIDLabel");

            //strDIDPrintLabel = Application.StartupPath + "\\" + pubFunction.ConfigListGetValue("UNIDLabel");

            GetLine();
            CboReportType.Items.Add("SAP1");
            CboReportType.Items.Add("SAP2");
            CboReportType.Items.Add("ReturnDID");
            CboReportType.Items.Add("DispatchDID");
            CboReportType.Items.Add("Return_Dispatch");
            CboReportType.Items.Add("ReturnDIDByGroupID");
            CboReportType.Items.Add("ReturnDIDByWO");
            CboReportType.Items.Add("CastQty");
            strCheckScaner = QMSSDK.Br.FileSystem.Ini.ReadIniValue("QSMS", "DIDScan", AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "set.ini");
            if (Parameter.PrtCallBKandReturn != "Y")
            {
                cmdGetRefID.Visible = false;
            }
            if (Parameter.CheckDIDRemainQty == "Y")
            {
                ChkRemainQty.Visible = true;
                ChkRemainQty.Enabled = true;
            }
            else
            {
                ChkRemainQty.Visible = false;
                ChkRemainQty.Enabled = false;
            }

            sstabDID.TabIndex = 0;
            GetDIDQty();
        }

        private void GetLine()
        {
            dt = Process.QSMS_MCC_QueryDataByType("PD_GetLine", "", "", "", "", "");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cboLine.Items.Add(dt.Rows[i]["Line"].ToString().Trim());
            }
        }

        private void CboGroupID_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (ChkGroupClosed(CboGroupID.Text.Trim()) == true)
                {
                    MessageBox.Show("The Group has been closed,can not return DID");
                    return;
                }
                Process.QSMSGroupCompQty(CboGroupID.Text.Trim());
                GetGroupWO(CboGroupID.Text.Trim());
                GetReturned_NotReturnDID(CboGroupID.Text.Trim());
            }
            catch
            {
                MessageBox.Show("System Error,Please contact QMS");
                return;
            }
        }

        private Boolean ChkGroupClosed(string GroupID = "")
        {
            dt = Process.QSMS_MCC_QueryDataByType("ChkGroupClosed", "", "", GroupID.Trim(), "", "");
            if (dt.Rows.Count > 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private void GetGroupWO(string GroupID = "")
        {
            dt = Process.QSMS_MCC_QueryDataByType("PD_GetWOInfoByGroupID", "", "", "", GroupID, "");
            cboWO.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cboWO.Items.Add(dt.Rows[i]["Work_Order"].ToString().Trim());
            }
        }

        private void GetReturned_NotReturnDID(string GroupID = "")
        {
            //dt = Process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_GroupDID", "", "", GroupID, "", "");
            dt = Process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_GroupDID2", "", "", GroupID, "", "");
            DGDIDReturned.DataSource = dt.DefaultView;
            if (Parameter.PrtCallBKandReturn == "Y")
            {
                dt = Process.XL_DIDGetToWHInfo("Return", "", Parameter.Factory, "N");
                gridDIDtoWH.DataSource = dt.DefaultView;
            }
        }

        private void cboWO_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtWO.Text = cboWO.Text.Trim();
            GetWOInfo(txtWO.Text);
            dt = Process.QSMS_WONeedReturnDID(txtWO.Text.Trim());
            if (dt.Rows.Count > 0)
            {
                DGDIDNeedReturned.DataSource = dt.DefaultView;
            }
        }

        //private void txtChkDID_KeyPress(object sender, KeyPressEventArgs e)
        //{
        //    dt = Process.QSMS_MCC_QueryDataByType("MCC_GetReturnQty", "", "", "", txtChkDID.Text.Trim(), "");
        //    if (dt.Rows.Count > 0)
        //    {
        //        DGDIDInfo.DataSource = dt.DefaultView;
        //    }
        //    else
        //    {
        //        lblChk.Text = "The CompPN does not belong to the GroupID";
        //    }
        //}

        private void GetWOInfo(string WO = "")
        {
            DataTable dt = Process.QSMS_MCC_QueryDataByType("PD_GetWOInfoByWO", "", "", "", WO, "");
            if (dt.Rows.Count > 0)
            {
                txtMBPN.Text = dt.Rows[0]["PN"].ToString();
                txtWOQty.Text = dt.Rows[0]["Qty"].ToString();
                txtCustomer.Text = dt.Rows[0]["Customer"].ToString();
            }
        }

        private void txtDID_KeyPress(object sender, KeyPressEventArgs e)
        {
            string PreDID = "", strOldTray = "";
            long RestQty, TotalQty;
            DataTable dt = new DataTable();
            if (Parameter.strKeyInPNByManual == true)
            {
                strCheckScaner = "N";
            }
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
                    pubFunction.Sound("Error");
                    MessageBox.Show("请使用刷枪作业");
                    timeSeq = 0;
                    return;
                }
            }

            if (txtDID.Text.Trim() != "" && e.KeyChar == 13)
            {
                txtDID.Text = txtDID.Text.ToString().Trim().Replace(" ", "").Replace("\r", "").Replace("\t", "").Replace("\n", "");
                //if (txtDID.Text.ToString().IndexOf(';') > 0) //新DID格式抓取DID、CompPN等信息


                //{
                //    dt = Process.getDID(txtDID.Text.ToString());
                //    txtDID.Text = dt.Rows[0]["UNID"].ToString().Trim();
                //    txtCompPN.Text = dt.Rows[0]["CompPN"].ToString().Trim();
                //    IsNewWay = "Y";
                //}
                //else
                //{
                //    IsNewWay = "N";
                //}
                ////dt = Process.getDID(txtDID.Text.ToString());////UnID还没有导入YAN 20211228
                ////txtDID.Text = dt.Rows[0]["UNID"].ToString().Trim();
                ////if (txtDID.Text.Trim().Length > 30)
                ////{
                ////    IsNewWay = "Y";
                ////}
                ////else
                ////{
                ////    IsNewWay = "N";
                ////}
                if (pubFunction.ConfigListGetValue("ChkFujiSPL") == "Y")
                {
                    dt = Process.QSMS_MCC_QueryDataByType("MCC_GetDIDRealQty", "", "", "", txtDID.Text.Trim(), "");
                    if (dt.Rows.Count > 0)
                    {
                        PreDID = dt.Rows[0]["DID"].ToString().Trim();
                        RestQty = Convert.ToInt32(dt.Rows[0]["RealQty"].ToString().Trim());
                        dt = Process.QSMS_MCC_QueryDataByType("MCC_GetDIDInfo", "", "", "", txtDID.Text.Trim(), "");
                        MessageBox.Show("此DID有接料且没用完！前一个DID为：" + PreDID + ";数量为:" + RestQty.ToString() + "\r\n" + "后一个DID为：" + dt.Rows[0]["DID"].ToString().Trim() + ";数量为:" + dt.Rows[0]["RealQty"].ToString().Trim());
                    }
                }
                dt = Process.QSMS_MCC_QueryDataByType("MCC_GetCompPN_Data", "", "", "", TxtCompPN.Text.Trim(), "");
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("此DID的材料是独用料,请确认Return的数量是否和实际数量一致！");
                }
                if (Parameter.BU == "NB5" || Parameter.BU == "PU5")
                {
                    dt = Process.QSMS_MCC_QueryDataByType("MCC_Getmsd_data", "", "", txtDID.Text.Trim(), "", "");
                    if (dt.Rows.Count > 0)
                    {
                        MessageBox.Show("该材料需真空包装再退回库!");
                    }
                }
                dt = Process.QSMS_MCC_GetDIDRealQty(txtDID.Text.Trim());
                if (dt.Rows.Count > 0)
                {
                    txtReturnQty.Text = dt.Rows[0]["realqty"].ToString().Trim();
                    TotalQty = Convert.ToInt64(dt.Rows[0]["TotalQty"].ToString().Trim());
                    if (TotalQty.ToString() == txtReturnQty.Text.Trim())
                    {
                        MessageBox.Show("此DID尚未在产线使用过，请再三检查是否真的需要Return!!!");
                    }
                }
                if (Parameter.CheckOldNewPrintType == "Y")
                {
                    dt = Process.QSMS_MCC_DIDPrintTypeCheck(txtDID.Text.Trim(), strOldTray, Parameter.g_factory);
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["result"].ToString().Trim() == "1")
                        {
                            MessageBox.Show(dt.Rows[0]["Message"].ToString().Trim());
                            txtDID.Text = "";
                            return;
                        }
                    }
                }
                if (ChkHUA.Checked == true)
                {
                    dt = Process.QSMS_MCC_XL_DIDReturnCheck(txtDID.Text.Trim(), "Y");
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["result"].ToString() == "1")
                        {
                            MessageBox.Show(dt.Rows[0]["Message"].ToString().Trim());
                            txtDID.Text = "";
                            return;
                        }
                    }
                }
                else
                {
                    dt = Process.QSMS_MCC_XL_DIDReturnCheck(txtDID.Text.Trim(), "");
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["result"].ToString() == "1")
                        {
                            MessageBox.Show(dt.Rows[0]["Message"].ToString().Trim());
                            txtDID.Text = "";
                            return;
                        }
                    }
                }
                txtReturnQty.Focus();

            }
        }

        private void txtReturnQty_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                btnSave_Click(null, null);
            }
        }

        private void CmdQuery_Click(object sender, EventArgs e)
        {
            if (cboLine.Text == "")
            {
                MessageBox.Show("Please input line");
                return;
            }
            GetGroupID();
        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            Sap_Return(CboReportType.Text.Trim());
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string TransDate, sDID = "", intReturnQty, sProcessStatus = "", sNewDispDID;
            lblFeedBack.Text = "Qty FeedBack:";
            try
            {
                if (ChkErr() == false)
                {
                    goto Normal_Eixt;
                }
                sDID = txtDID.Text.Trim();
                intReturnQty = txtReturnQty.Text.Trim();
                dt = Process.QSMS_MCC_XL_CheckReturnQty(sDID, TxtCompPN.Text.Trim(), int.Parse(intReturnQty.Trim()), PreGroupID, IsAnotherBUDID.Trim(), pubFunction.ConfigListGetValue("CheckReturnForbiddenPN"));
                if (dt.Rows[0]["Result"].ToString().Trim() == "F")
                {
                    LblMessage.Text = dt.Rows[0]["Description"].ToString().Trim();
                    MessageBox.Show(dt.Rows[0]["Description"].ToString());
                    LblMessage.ForeColor = Color.Black;
                    goto Normal_Eixt;
                }
                if (dt.Rows[0]["Result"].ToString().Trim() == "0")
                {
                    if (MessageBox.Show(dt.Rows[0]["Description"].ToString().Trim(), "", MessageBoxButtons.YesNo) == DialogResult.No)
                    {
                        goto Normal_Eixt;
                    }
                    LblMessage.Text = dt.Rows[0]["Description"].ToString().Trim();
                    LblMessage.ForeColor = Color.Black;
                }
                if (UpdateReturnQty(TxtCompPN.Text.Trim(), sDID.Trim(), Convert.ToInt32(intReturnQty.Trim()), PreGroupID.Trim()) == false)
                {
                    goto Normal_Eixt;
                }

                if (Parameter.BU == "ESBU" || Parameter.IC_CompChk == "Y")
                {
                    if (IC_CompNeedBurn(TxtCompPN.Text.Trim()) == true)
                    {
                        goto Normal_Eixt;
                    }
                }

                if (Parameter.PrtCallBKandReturn == "Y")
                {
                    sProcessStatus = "Return Start";
                    ds = Process.QSMS_MCC_XL_DIDGetNewID("Return", sDID, (optGoodMaterial.Checked == true) ? "Y" : "N", int.Parse(intReturnQty), Parameter.g_userName, Parameter.Factory, IsAnotherBUDID);
                    dt = ds.Tables[0];
                    if (dt.Rows[0]["Result"].ToString() != "0")
                    {
                        LblMessage.Text = dt.Rows[0]["Description"].ToString();
                    }
                    else
                    {
                        if (IsNewWay == "Y")//003
                        {
                            goto Normal_Eixt;
                        }
                        print = ds.Tables[1];
                        if (print.Rows.Count <= 0)
                        {
                            LblMessage.Text = "Get DID information fail,print DID fail!!";
                            LblMessage.ForeColor = Color.Red;
                            goto Normal_Eixt;
                        }
                        lblFeedBack.Text = print.Rows[0]["QtyFeedback"].ToString().Trim();
                        if (lblFeedBack.Text.IndexOf("##") > 0)
                            lblFeedBack.Text = lblFeedBack.Text.Substring(lblFeedBack.Text.IndexOf("##"), lblFeedBack.Text.Trim().Length - lblFeedBack.Text.IndexOf("##"));

                        Parameter.DIDInfo.DID = print.Rows[0]["DID"].ToString().Trim();
                        Parameter.DIDInfo.compPN = print.Rows[0]["CompPN"].ToString().Trim();
                        Parameter.DIDInfo.Qty = int.Parse(print.Rows[0]["Qty"].ToString().Trim());
                        Parameter.DIDInfo.IsGood = print.Rows[0]["IsGood"].ToString().Trim();
                        Parameter.DIDInfo.VendorCode = print.Rows[0]["VendorCode"].ToString().Trim();
                        Parameter.DIDInfo.DateCode = print.Rows[0]["DateCode"].ToString().Trim();
                        Parameter.DIDInfo.LotCode = print.Rows[0]["LotCode"].ToString().Trim();
                        if (Parameter.BU == "NB5" || Parameter.BU == "PU5")
                        {
                            Parameter.DIDInfo.WareHouseID = print.Rows[0]["WareHouseID"].ToString().Trim();
                        }
                        if (pubFunction.ConfigListGetValue("ChkPrintDIDType") == "Y")
                        {
                            Parameter.DIDInfo.DIDType = print.Rows[0]["DIDType"].ToString().Trim();
                        }
                        else
                        {
                            Parameter.DIDInfo.DIDType = "";
                        }
                        if (Parameter.DIDInfo.IsGood == "Y")
                        {
                            TransDate = Process.QSMS_MCC_QueryDataByType("MCC_GetDate", "", "", "", "", "").Rows[0]["TransDateTime"].ToString();
                            TempDID = Process.QSMS_MCC_QueryDataByType("MCC_GetDID", "", "", Parameter.DIDInfo.compPN, TransDate, Parameter.DIDHead).Rows[0][0].ToString().Trim();//002
                            sProcessStatus = "Dispatch_Start";
                            dt = Process.Return_XL_DIDAutoDispatch(TempDID, Parameter.DIDInfo.compPN, Parameter.DIDInfo.Qty.ToString(), Parameter.DIDInfo.Qty.ToString(), Parameter.DIDInfo.VendorCode, Parameter.DIDInfo.DateCode, Parameter.DIDInfo.LotCode, "", "", Parameter.g_userName, "4", Parameter.Factory, sDID);
                            sProcessStatus = "Dispatch_End";
                            if (dt.Rows.Count > 0)
                            {
                                LblMessage.Text = dt.Rows[0]["ErrDesc"].ToString();
                                if (dt.Rows[0]["result"].ToString() != "1")
                                {
                                    dt = Process.QSMS_MCC_QueryDataByType("MCC_QSMS_Error_Log", "", "", sDID.Trim(), LblMessage.Text.Trim(), "");
                                    goto PrintLabel;
                                }
                                else
                                {

                                    sProcessStatus = "Del_ToWH_DID";
                                    sNewDispDID = dt.Rows[0]["DID"].ToString().Trim();
                                    TempDID = sNewDispDID;
                                    ds = Process.QSMS_MCC_XL_GetDidPrintInfo_Return(sNewDispDID, sDID, "N", "", pubFunction.ConfigListGetValue("PrinterType"), pubFunction.ConfigListGetValue("PrintDpm"));
                                    dt = ds.Tables[0];
                                    if (dt.Rows.Count <= 0)
                                    {
                                        LblMessage.Text = "Can not get auto dispatch did";
                                        goto PrintLabel;
                                    }
                                    else
                                    {
                                        if (dt.Rows[0]["Result"].ToString().Trim() != "0")
                                        {
                                            LblMessage.Text = dt.Rows[0]["Description"].ToString().Trim();
                                            goto PrintLabel;
                                        }
                                        else
                                        {
                                            dt = ds.Tables[1];
                                        }
                                        PrintData = dt;
                                        PrintAutoDispatchLabel();
                                        LblMessage.Text = "Return DID auto dispatch successful!";
                                        LblMessage.ForeColor = Color.Black;
                                        goto Normal_Eixt;
                                    }
                                }

                            }
                            else
                            {
                                LblMessage.Text = "Auto dispatch return DID fail!";
                                LblMessage.ForeColor = Color.Red;
                                goto Normal_Eixt;
                            }
                        }
                    }
                }
            PrintLabel:
                {
                    sProcessStatus = "Print_Label";
                    DIDPrintLabel();
                }
                if (IsAnotherBUDID != "Y")
                {
                    GetReturned_NotReturnDID(CboGroupID.Text.Trim());
                    GetDIDInfo(sDID, CboGroupID.Text.Trim());
                }
            Normal_Eixt:
                {
                    txtDID.Text = "";
                    txtReturnQty.Text = "";
                    txtDID.Focus();
                }
            }
            catch (Exception ex)
            {
                string sErrMsg;
                sErrMsg = "ErrDesc: " + ex.Message;
                dt = Process.QSMS_MCC_QueryDataByType("MCC_QSMS_Error_Log1", "", "", "DID:" + TempDID + ";SDID:" + sDID + ";ProcessStatus:" + sProcessStatus + ";ErrDetail:" + sErrMsg, Parameter.g_userName, "");
                if (sProcessStatus.ToUpper() == "Dispatch_Start".ToUpper() || sProcessStatus.ToUpper() == "Del_ToWH_DID".ToUpper())
                {
                    if (ChkDispatchIsOK(TempDID.Trim(), sDID, sErrMsg) == true)
                    {
                        sProcessStatus = "Del_ToWH_DID";
                        sNewDispDID = dt.Rows[0]["DID"].ToString().Trim();
                        TempDID = sNewDispDID;
                        ds = Process.QSMS_MCC_XL_GetDidPrintInfo_Return(sNewDispDID, sDID, "N", "", pubFunction.ConfigListGetValue("PrinterType"), pubFunction.ConfigListGetValue("PrintDpm"));
                        dt = ds.Tables[0];
                        if (dt.Rows.Count <= 0)
                        {
                            LblMessage.Text = "Can not get auto dispatch did";
                            goto PrintLabel;
                        }
                        else
                        {
                            if (dt.Rows[0]["Result"].ToString().Trim() != "0")
                            {
                                LblMessage.Text = dt.Rows[0]["Description"].ToString().Trim();
                                goto PrintLabel;
                            }
                            else
                            {
                                dt = ds.Tables[1];
                            }
                            PrintData = dt;
                            PrintAutoDispatchLabel();
                            LblMessage.Text = "Return DID auto dispatch successful!";
                            goto Normal_Eixt;
                        }
                    }
                PrintLabel:
                    {
                        sProcessStatus = "Print_Label";
                        DIDPrintLabel();
                    }
                    if (IsAnotherBUDID != "Y")
                    {
                        GetReturned_NotReturnDID(CboGroupID.Text.Trim());
                        GetDIDInfo(sDID, CboGroupID.Text.Trim());
                    }
                Normal_Eixt:
                    {
                        txtDID.Text = "";
                        txtReturnQty.Text = "";
                        txtDID.Focus();
                    }
                }
                else
                {
                    MessageBox.Show(sErrMsg + ",Please contact QSMS SMT Staff ");
                }
            }
        }

        private void cmdGetRefID_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            DataTable dt_temp = new DataTable();
            DataSet ds = new DataSet();
            string sCurrRefID, sMsg;
            dt = Process.QSMS_MCC_XL_DIDGetRefID("Return", (optGoodMaterial.Checked == true) ? "Y" : "N", Parameter.g_userName, Parameter.Factory, IsAnotherBUDID);
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["Result"].ToString() != "0")
                {
                    MessageBox.Show(dt.Rows[0]["Description"].ToString(), "Prompt");
                    return;
                }
                sMsg = dt.Rows[0]["Description"].ToString().Trim();
                sCurrRefID = pubFunction.DIDGetRefIDByResult(sMsg);
                Parameter.DIDInfo.DID = sCurrRefID;
                Parameter.DIDInfo.compPN = sCurrRefID;
                Parameter.DIDInfo.Qty = -100000;
                Parameter.DIDInfo.IsGood = (optGoodMaterial.Checked == true) ? "Y" : "N";
                Parameter.DIDInfo.DIDType = "";

                print.Rows[0]["DID"] = sCurrRefID;
                print.Rows[0]["CompPN"] = sCurrRefID;
                print.Rows[0]["Qty"] = -100000;
                print.Rows[0]["IsGood"] = Parameter.DIDInfo.IsGood;
                //if (print.Rows.Count > 0)   //20220107  Rain
                //{
                //    for (int i = 0; i < print.Columns.Count; i++)
                //    {
                //        if (print.Columns[i].ColumnName != "Qty")
                //        {
                //            dt_temp.Columns.Add(print.Columns[i].ColumnName);
                //        }
                //        else
                //        {
                //            dt_temp.Columns.Add("Qty");
                //        }
                //    }

                //    dt_temp.Rows.Add();
                //    for (int i = 0; i < print.Columns.Count; i++)
                //    {
                //        if (print.Columns[i].ColumnName != "Qty")
                //        {
                //            dt_temp.Rows[0][i] = print.Rows[0][i].ToString();
                //        }
                //        else
                //        {
                //            dt_temp.Rows[0][i] = "RefID";
                //        }
                //    }

                //    print = dt_temp;
                //    //print.Rows[0]["Qty"] = int.Parse("RefID");
                //}

                DIDPrintLabel();//DIDPrintLabel(OptZebra.Value, CInt(Trim(TxtCompPort)), Trim(TxtComm))  

                ds = Process.XL_DIDChkStockByRefID_set(sCurrRefID, Parameter.g_userName);
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["Result"].ToString() != "0")
                    {
                        MessageBox.Show(dt.Rows[0]["Description"].ToString(), "Prompt");
                        return;
                    }
                    frmDIDCheckStock DIDCheckStock = new frmDIDCheckStock();
                    frmDIDCheckStock.FuncType = "AutoChk";

                    dt = ds.Tables[1];
                    frmDIDCheckStock.rstCompPN = dt;

                    DIDCheckStock.Show();
                }
            }
        }

        private void GetGroupID()
        {
            string BeginDate, EndDate;
            BeginDate = dtpSDate.Text;
            BeginDate = BeginDate.Replace("-", "").Replace("/", "");
            EndDate = dtpEDate.Text;
            EndDate = EndDate.Replace("-", "").Replace("/", "");
            if (Parameter.BU == "NB5" || Parameter.BU == "PU5")
            {
                if (OptRelease.Checked == true)
                {
                    dt = Process.QSMS_MCC_QueryDataByType("PD_GetGroupIDByDate1_NB5", BeginDate, EndDate, cboLine.Text.Trim(), "", "");
                }
                else
                {
                    dt = Process.QSMS_MCC_QueryDataByType("PD_GetGroupIDByDate2_NB5", BeginDate, EndDate, cboLine.Text.Trim(), "", "");
                }
            }
            else
            {
                if (OptRelease.Checked == true)
                {
                    dt = Process.QSMS_MCC_QueryDataByType("PD_GetGroupIDByDate1", BeginDate, EndDate, cboLine.Text.Trim(), "", "");
                }
                else
                {
                    dt = Process.QSMS_MCC_QueryDataByType("PD_GetGroupIDByDate2", BeginDate, EndDate, cboLine.Text.Trim(), "", "");
                }
            }
            CboGroupID.Items.Clear();
            if (dt.Rows.Count > 0)
            {
                for (int n = 0; n < dt.Rows.Count; n++)
                {
                    CboGroupID.Items.Add(dt.Rows[n]["GroupID"].ToString().Trim());
                }
            }
            else
            {
                MessageBox.Show("No data");
            }
        }

        private void Sap_Return(string Report_Type = "")
        {
            switch (Report_Type)
            {
                case "SAP1":
                    dt = Process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_Sap", "", "", cboWO.Text.Trim(), "open", "");
                    break;
                case "SAP2":
                    dt = Process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_Sap", "", "", cboWO.Text.Trim(), "close", "");
                    break;
                case "ReturnDID":
                    dt = Process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_GroupDID3", "", "", CboGroupID.Text.Trim(), "", "");
                    break;
                case "DispatchDID":
                    dt = Process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_DispatchByGroup", "", "", CboGroupID.Text.Trim(), "", "");
                    break;
                case "ReturnDIDByGroupID":
                    dt = Process.QSMS_MCC_XL_ReturnDIDByGroupID(CboGroupID.Text.Trim());
                    break;
                case "ReturnDIDByWO":
                    dt = Process.QSMS_MCC_XL_ReturnDIDByWO(CboGroupID.Text.Trim());
                    break;
                case "Return_Dispatch":
                    dt = Process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_GroupCompQty", "", "", CboGroupID.Text.Trim(), "", "");
                    break;
                case "CastQty":
                    Process.QSMS_MCC_QSMSGetCastQty(CboGroupID.Text.Trim());
                    dt = Process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_GroupCompQty", "", "", CboGroupID.Text.Trim(), "", "");
                    break;
            }
            if (dt.Rows.Count > 0)
            {
                pubFunction.doExport(dt);
            }
            else
            {
                MessageBox.Show("No data");
            }
        }

        private void DIDPrintLabel()
        {
            string BU = "";
            PrinterLib.PrintLabel lblprint = new PrinterLib.PrintLabel();
            if (lblprint.LabelSetting(strCommSetting, strPrintPort, 1, ref msg) == false)
            {
                LblMessage.Text = msg;
                LblMessage.ForeColor = Color.Red;
                return;
            }

            if (PrintReturnDIDLabel.Substring(PrintReturnDIDLabel.Length - 3, 3).ToUpper() != "TXT") //20220106 Rain
            {
                if (Parameter.BU == "ESBU" && Process.CheckCompPN(TxtCompPN.Text.Trim(), "IsNeedMSD"))
                {
                    PrintReturnDIDLabel = GetDIDLabelFile("good_MSD");
                }
                else
                {
                    PrintReturnDIDLabel = GetDIDLabelFile((Parameter.DIDInfo.IsGood == "Y") ? "good" : "bad");
                }
            }

            if (File.Exists(PrintReturnDIDLabel) == false)
            {
                MessageBox.Show("File:" + PrintReturnDIDLabel + " not exists");
                LblMessage.ForeColor = Color.Red;
                return;
            }
            if (string.IsNullOrEmpty(strLabelContent))
            {
                strLabelContent = new StreamReader(PrintReturnDIDLabel).ReadToEnd();
            }

            if (lblprint.PrintReturn(strLabelContent, print, BU, ref msg) == false)
            {
                LblMessage.Text = msg;
                LblMessage.ForeColor = Color.Red;
                return;
            }

        }

        private Boolean ChkErr()
        {
            DataTable dt = new DataTable();
            if (txtDID.Text == "")
            {
                return false;
            }
            if (IsAnotherBUDID == "Y")
            {
                return true;
            }
            if (ChkDIDBelongToGroupID(CboGroupID.Text.Trim(), txtDID.Text.Trim()) == false)
            {
                return false;
            }
            if (txtReturnQty.Text.Trim() == "" || pubFunction.IsNumeric(txtReturnQty.Text.Trim(), "INT") == false)
            {
                MessageBox.Show("The Return Qty can not be empty or must be numeric");
                return false;
            }
            txtReturnQty.Text = ABS(txtReturnQty.Text.Trim());
            if (txtReturnQty.Text.Trim() == "0")
            {
                MessageBox.Show("The Return Qty must be >0 !!");
                return false;
            }
            if (ChkDIDInMachine(txtDID.Text.Trim()) == false)
            {
                return false;
            }
            if (pubFunction.ConfigListGetValue("CheckMSDCallBack") == "Y")
            {
                dt = Process.QSMS_MCC_QueryDataByType("MCC_Getmsd_data", "", "", TxtCompPN.Text.Trim(), "", "");
                if (dt.Rows.Count > 0)
                {
                    LblMessage.Text = "This is MSD Material! ";
                    dt = Process.QSMS_MCC_PD_MSD_LinkDIDAuto("", txtDID.Text.Trim(), TxtCompPN.Text.Trim(), "", "Y", Parameter.g_userName);
                    if (dt.Rows[0]["result"].ToString() == "CHECKFAIL")
                    {
                        MessageBox.Show("Message: " + dt.Rows[0]["ErrDesc"].ToString());
                        LblMessage.ForeColor = Color.Red;
                        return false;
                    }
                }
            }
            return true;
        }

        private Boolean ChkDIDBelongToGroupID(string GroupID = "", string DID = "")
        {
            DataTable dt = new DataTable();
            if (GroupID.Trim() == "")
            {
                return true;
            }
            dt = Process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_GroupDID4", "", "", GroupID.Trim(), DID.Trim(), "");
            if (dt.Rows.Count <= 0)
            {
                MessageBox.Show("The DID does not belong to the GroupID,Please check");
                return false;
            }
            return true;
        }

        private Boolean ChkDIDInMachine(string DID = "")
        {
            DataTable dt = Process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_FeederDID_Current", "", "", DID, "", "");
            if (dt.Rows.Count > 0)
            {
                MessageBox.Show("The DID is in Machine :" + dt.Rows[0]["Machine"].ToString().Trim() + " Feeder :" + dt.Rows[0]["Feeder"].ToString().Trim() + "  Please delete first");
                return false;
            }
            return true;
        }

        private string ABS(string str)
        {
            if (Convert.ToInt32(str) < 0)
            {
                return ((-1) * Convert.ToInt32(str)).ToString();
            }
            else
            {
                return str;
            }
        }

        private Boolean UpdateReturnQty(string compPN, string DID, long ReturnQty, string GroupID = "")
        {
            string transdatetime, ReturnDIDSeq, intPos;
            transdatetime = Process.QSMS_MCC_QueryDataByType("MCC_GetDate", "", "", "", "", "").Rows[0]["TransDateTime"].ToString().Trim();
            ReturnDIDSeq = compPN + "-A" + transdatetime;
            dt = Process.QSMS_MCC_QSMS_ReturnDID(ReturnDIDSeq.Trim(), DID, compPN, Convert.ToInt32(ReturnQty.ToString()), Parameter.g_userName, GroupID, transdatetime, (optGoodMaterial.Checked == true) ? "Y" : "N", Parameter.PrtCallBKandReturn, Parameter.Factory, IsAnotherBUDID);
            if (dt.Rows.Count > 0)
            {
                LblMessage.Text = dt.Rows[0]["Description"].ToString().Trim();
                if (dt.Rows[0]["Result"].ToString().Trim() == "0")
                {
                    intPos = dt.Rows[0]["Description"].ToString().Trim().Substring(0, dt.Rows[0]["Description"].ToString().Trim().IndexOf("PreGroupID:"));
                    intPos = dt.Rows[0]["Description"].ToString().Trim().IndexOf("PreGroupID:").ToString();
                    PreGroupID = dt.Rows[0]["Description"].ToString().Trim().Substring(int.Parse(intPos) + "PreGroupID:".Length);
                    return true;
                }
            }
            return false;
        }

        private void GetDIDInfo(string DID = "", string GroupID = "")
        {
            DataTable dt = new DataTable();
            if (txtDID.Text.Trim() == "")
                return;
            dt = Process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_GroupDID4", "", "", GroupID, DID, "");
            if (dt.Rows.Count > 0)
            {
                txtDIDTotalQty.Text = dt.Rows[0]["TotalQty"].ToString();
                txtDIDReturnedQty.Text = dt.Rows[0]["ReturnQty"].ToString();
                TxtCompPN.Text = dt.Rows[0]["compPN"].ToString();
            }
            else
            {
                dt = Process.QSMS_MCC_QueryDataByType("MCC_GetDIDInfo", "", "", "", DID, "");
                if (dt.Rows.Count > 0)
                {
                    txtDIDTotalQty.Text = dt.Rows[0]["Qty"].ToString();
                    txtDIDReturnedQty.Text = "0";
                    TxtCompPN.Text = dt.Rows[0]["compPN"].ToString();
                }
                else
                {
                    MessageBox.Show("DID:" + DID + " is not existed in QSMS_DID!!", "Prompt");
                    return;
                }
            }
            //dt = Process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_GroupCompQty2", "", "", GroupID, txtCompPN.Text.Trim(), "");
            //DGCompInfo.DataSource = dt.DefaultView;
            //DGCompInfo.Refresh();
            GetDIDQty();
        }

        private void GetDIDQty()  //20230518 Rain 统计当前已刷料盘数
        {
            DataTable dt = new DataTable();
            dt = Process.QSMS_MCC_QueryDataByType("MCC_GetDID_Qty", "", "", Parameter.g_userName, "", "");
            if (dt.Rows.Count > 0)
            {
                lbl_DIDQty.Text = dt.Rows[0]["Qty"].ToString();
            }
            else
            {
                return;                
            }

        }

        private Boolean ChkDispatchIsOK(string sNewDID, string sOldDID, string Msg)
        {
            DataTable dt = Process.QSMS_MCC_XL_DIDChk_ReturnDisp(sNewDID, sOldDID, Msg);
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["ReSult"].ToString() == "0")
                    return true;
            }
            return false;
        }

        private void PrintAutoDispatchLabel()
        {
            string BU = "";
            PrinterLib.PrintLabel lblprint = new PrinterLib.PrintLabel();
            if (lblprint.LabelSetting(strCommSetting, strPrintPort, 1, ref msg) == false)
            {
                LblMessage.Text = msg;
                LblMessage.ForeColor = Color.Red;
                return;
            }

            if (PrintReturnDIDLabel.Substring(PrintReturnDIDLabel.Length - 3, 3).ToUpper() != "TXT")        //20220106 Rain
            {
                PrintReturnDIDLabel = GetDIDLabelFile((opOldLabel.Checked == true) ? "OLD" : "NEW");
            }

            if (File.Exists(PrintReturnDIDLabel) == false)
            {
                MessageBox.Show("File:" + PrintReturnDIDLabel + " not exists");
                return;
            }
            if (string.IsNullOrEmpty(strDIDLabelContent))
            {
                strDIDLabelContent = new StreamReader(PrintReturnDIDLabel).ReadToEnd();
            }

            BU = (IsAnotherBUDID == "Y") ? Parameter.AutoDispatchForAnotherBU : Parameter.BUDIDShow;

            if (lblprint.PrintReturnDID(strDIDLabelContent, BU, pubFunction.ConfigListGetValue("PrintedVenderCode"), pubFunction.ConfigListGetValue("PrintedSeqID"), PrintData, ds, ref msg) == false)
            {
                LblMessage.Text = msg;
                LblMessage.ForeColor = Color.Red;
                return;
            }

        }

        private void FrmReturnDID_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("FrmReturnDID");
        }

        private void txtDID_Leave(object sender, EventArgs e)  //001
        {
            if (txtDID.Text == "")
            {
                return;
            }

            if (txtDID.Text.Trim().Length > 3 && txtDID.Text.Trim().Substring(txtDID.Text.Trim().Length - 3, 3).Substring(0, 1) == "R")
            {
                return;
            }

            IsAnotherBUDID = "N";
            if (Parameter.AutoDispatchForAnotherBU != "")
            {
                if (XL_ChkAnotherBUDID(txtDID.Text.Trim().ToUpper()) == false)
                {
                    txtDID.Text = "";
                    txtDID.Focus();
                    return;
                }
                if (IsAnotherBUDID == "Y")
                {
                    return;
                }
            }

            if (ChkDIDBelongToGroupID(CboGroupID.Text.Trim(), txtDID.Text.Trim()) == false)
            {
                txtDID.Text = "";
                txtDID.Focus();
                return;
            }

            if (ChkDIDInMachine(txtDID.Text.Trim()) == false)
            {
                txtDID.Text = "";
                txtDID.Focus();
                return;
            }

            GetDIDInfo(txtDID.Text.Trim(), CboGroupID.Text.Trim());
        }

        private Boolean XL_ChkAnotherBUDID(string sDID)  //001
        {
            DataSet ds = Process.QSMS_MCC_XL_ChkAnotherBUDID(sDID, IsAnotherBUDID, Parameter.Factory);
            DataTable dt = ds.Tables[0];
            if (dt.Rows[0]["Description"].ToString().Trim() != "0")
            {
                LblMessage.Text = dt.Rows[0]["Description"].ToString();
                LblMessage.BackColor = Color.White;
                LblMessage.ForeColor = Color.Red;
                return false;
            }
            else
            {
                LblMessage.BackColor = Color.White;
                LblMessage.ForeColor = Color.Black;
                dt = ds.Tables[1];
                txtDIDReturnedQty.Text = dt.Rows[0]["TotalQty"].ToString();
                txtDIDReturnedQty.Text = dt.Rows[0]["ReturnQty"].ToString();
                TxtCompPN.Text = dt.Rows[0]["compPN"].ToString();
                IsAnotherBUDID = dt.Rows[0]["IsAnotherBUDID"].ToString();
            }
            return true;
        }

        private void TxtChkDID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (TxtChkDID.Text.Trim() != "" && e.KeyChar == 13)
            {
                string SQLstr = "select A.DID,A.Qty as TotalQty,B.ReturnQty ,A.RealQty from QSMS_DID A left join QSMS_GroupDID B on  A.DID=B.DID where A.DID='" + TxtChkDID.Text.Trim() + "'";
                DataTable dtResult = PD.QSMS_EXE(SQLstr);

                if (dtResult == null || dtResult.Rows.Count == 0)
                {
                    LblChk.Text = "The CompPN does not belong to the GroupID";
                }
                else
                {
                    DGDIDInfo.DataSource = dtResult;
                }
            }
        }

        private bool IC_CompNeedBurn(string CompPN)
        {
            string SQLstr = "exec IC_CompNeedBurn  '" + CompPN + "'";

            try
            {
                DataTable dtResult = PD.QSMS_EXE(SQLstr);

                if (dtResult.Rows.Count > 0)
                {
                    if (dtResult.Rows[0]["Result"].ToString().Trim() == "0")
                    {
                        if (MessageBox.Show(dtResult.Rows[0]["Description"].ToString().Trim() + " DO you burn IC for it firstly!", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return false;
            }

            return false;

        }

        private void cmdReprint_Click(object sender, EventArgs e)
        {
            bool IsByDIDInput = true;

            if (TxtCompPort.Text == "" || TxtComm.Text == "")
            {
                MessageBox.Show("Printer have not set!!");
                return;
            }

            if (txtDID.Text == "")
            {
                LblMessage.Text = "Please select or Input DID to reprint!!";
                return;
            }

            if (gridDIDtoWH.Rows.Count > 0)
            {
                for (int i = 0; i < gridDIDtoWH.Rows.Count; i++)
                {
                    if (gridDIDtoWH.Rows[i].Cells[0].Value.ToString() == txtDID.Text)
                    {
                        Parameter.DIDInfo.DID = gridDIDtoWH.Rows[i].Cells[1].Value.ToString();
                        Parameter.DIDInfo.compPN = gridDIDtoWH.Rows[i].Cells[2].Value.ToString();
                        Parameter.DIDInfo.Qty = Convert.ToInt64(gridDIDtoWH.Rows[i].Cells[3].Value.ToString());
                        Parameter.DIDInfo.IsGood = gridDIDtoWH.Rows[i].Cells[10].Value.ToString();

                        if (pubFunction.ConfigListGetValue("ChkPrintDIDType") == "Y")
                        {
                            Parameter.DIDInfo.DIDType = gridDIDtoWH.Rows[i].Cells[14].Value.ToString();
                        }
                        else
                        {
                            Parameter.DIDInfo.DIDType = "";
                        }

                        IsByDIDInput = false;
                    }
                }
            }

            if (IsByDIDInput == true)
            {
                //DataTable dt = Process.XL_DIDGetToWHInfo("Return", txtDID.Text, Parameter.Factory, "N");
                print = Process.XL_DIDGetToWHInfo("Return", txtDID.Text, Parameter.Factory, "N");

                if (print.Rows.Count > 0)
                {
                    Parameter.DIDInfo.DID = print.Rows[0]["DID"].ToString();
                    Parameter.DIDInfo.compPN = print.Rows[0]["CompPN"].ToString();
                    Parameter.DIDInfo.Qty = Convert.ToInt64(print.Rows[0]["Qty"].ToString());
                    Parameter.DIDInfo.IsGood = print.Rows[0]["IsGood"].ToString();
                    Parameter.DIDInfo.DateCode = print.Rows[0]["DateCode"].ToString();
                    Parameter.DIDInfo.VendorCode = print.Rows[0]["VendorCode"].ToString();
                    Parameter.DIDInfo.LotCode = print.Rows[0]["LotCode"].ToString();

                    if (pubFunction.ConfigListGetValue("ChkPrintDIDType") == "Y")
                    {
                        Parameter.DIDInfo.DIDType = print.Rows[0]["DIDType"].ToString();
                    }
                    else
                    {
                        Parameter.DIDInfo.DIDType = "";
                        print.Rows[0]["DIDType"] = "";
                    }
                }
                else
                {
                    LblMessage.Text = "There is no DID:" + txtDID.Text + " !!";
                    txtDID.Text = "";
                    txtReturnQty.Text = "";
                    txtDID.Focus();
                    return;
                }

                DIDPrintLabel();
            }

        }

        private void DGDIDNeedReturned_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string DID = DGDIDNeedReturned.Rows[e.RowIndex].Cells[0].Value.ToString();
            TxtChkDID.Text = DID;

            TxtChkDID_KeyPress(sender, null);
        }

        private void DGDIDReturned_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string DID = DGDIDNeedReturned.Rows[e.RowIndex].Cells[0].Value.ToString();
            TxtChkDID.Text = DID;

            TxtChkDID_KeyPress(sender, null);
        }

        private void DGDIDInfo_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                txtDID.Text = DGDIDInfo.Rows[e.RowIndex].Cells[0].Value.ToString();
                txtDIDReturnedQty.Text = DGDIDInfo.Rows[e.RowIndex].Cells[1].Value.ToString();
                txtDIDTotalQty.Text = DGDIDInfo.Rows[e.RowIndex].Cells[2].Value.ToString();

                dt = Process.QSMS_MCC_QueryDataByType("MCC_GetDIDInfo", "", "", "", txtDID.Text, "");
                if (dt.Rows.Count > 0)
                {
                    //txtDIDTotalQty.Text = dt.Rows[0]["Qty"].ToString();
                    //txtDIDReturnedQty.Text = "0";
                    TxtCompPN.Text = dt.Rows[0]["compPN"].ToString();
                }

                //TxtCompPN.Text = txtDID.Text.Substring(0, 11);
            }
            catch (Exception exc)
            {
                txtDIDTotalQty.Text = "";
                txtDIDReturnedQty.Text = "";
            }
        }

        private void gridDIDtoWH_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtDID.Text = gridDIDtoWH.Rows[e.RowIndex].Cells[0].Value.ToString();
            txtDID.Focus();
        }

        private string GetDIDLabelFile(string LabelType) //20220106 Rain
        {
            string Setting = "200";
            try
            {
                if (Registry.GetValue("HKEY_CURRENT_USER\\Software\\VB and VBA Program Settings\\SMT\\QSMS", "DPM", "200").ToString() != null)
                {
                    Setting = Registry.GetValue("HKEY_CURRENT_USER\\Software\\VB and VBA Program Settings\\SMT\\QSMS", "DPM", "200").ToString();
                }
            }
            catch (Exception e)
            {
            }

            string DIDLabelFile = PrintReturnDIDLabel + "Zebra_" + Setting;

            if (LabelType != "")
            {
                DIDLabelFile = DIDLabelFile + "_" + LabelType + ".TXT";
            }
            else
            {
                DIDLabelFile = DIDLabelFile + ".TXT";
            }

            return DIDLabelFile;
        }

    }
}
