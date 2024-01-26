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

// 20210129  JU    002
namespace QSMS.QSMS.MCC
{
    public partial class FrmDIDCallBack : Form
    {
        public FrmDIDCallBack()
        {
            InitializeComponent();
        }

        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.MCC.MCCProcess process = new DbLibrary.MCC.MCCProcess();
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        DataTable PrintData = new DataTable();
        DataTable rstDIDCalled = new DataTable();
        DataTable rstDIDNotCall = new DataTable();
        DataTable rstWOGroup = new DataTable();
        DataTable rstDID = new DataTable();
        DataTable rstDIDtoWH = new DataTable();
        private string strLabelContent;
        private string strPrintPort;
        private string strCommSetting;
        private string PrintDIDCallBackLabel;//新增flag设置打印路径及模板名称
        private string msg;

        private string IsAnotherBUDID;

        private void FrmDIDCallBack_Load(object sender, EventArgs e)
        {
            dtpEDate.Text = DateTime.Now.ToShortDateString();
            dtpSDate.Text = DateTime.Now.ToShortDateString();
            GetLine();
            optGoodMaterial.Enabled = false;
            optBadMaterial.Enabled = false;
            if (Parameter.PrtCallBKandReturn != "Y")
            {
                cmdGetRefID.Visible = false;
                cmdReprint.Visible = false;      
            }
            PrintDIDCallBackLabel = Application.StartupPath + "\\" + pubFunction.ConfigListGetValue("PrintReturnLabel");
            strPrintPort = pubFunction.ConfigListGetValue("PrintPort");
            strCommSetting = pubFunction.ConfigListGetValue("CommSetting");
        }

        private void cboGroupID_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetGroupWO(cboGroupID.Text.Trim());
        }

        private void cboWO_SelectedIndexChanged(object sender, EventArgs e)
        {
            dt = process.QSMS_MCC_QueryDataByType("MCC_GetGroupIDByWO", "", "", cboWO.Text.Trim(), "", "");
            if (dt.Rows.Count > 0)
            {
                if (ChkGroupClosed(dt.Rows[0]["GroupID"].ToString().Trim()) == true)
                {
                    MessageBox.Show("The Group has been closed,can not return DID");
                    return;
                }
            }
            else
            {
                 MessageBox.Show("Can not find the GroupID for the Work Order:"+cboWO.Text.Trim());
                 return;
            }
            GetWOGroupInfo(cboWO.Text.Trim());
            process.QSMS_MCC_QSMSDIDCallBack(cboWO.Text.Trim());
            GetDIDInfoCallBack(cboWO.Text.Trim());
            txtDID.Text = "";
            txtReturnQty.Text = "";
        }

        private void txtDID_KeyPress(object sender, KeyPressEventArgs e)
        {
            DataTable dt = new DataTable();
            if (e.KeyChar == 13 && txtDID.Text != "")
            {
                if (txtDID.Text.Trim().Substring(2, 1) == "C")
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
                        txtReturnQty.Focus();
                        return;
                    }
                }
                if (ChkDIDBelongToPCB(cboWO.Text.Trim(), txtDID.Text.Trim()) == false)
                {
                    return;
                }
                GetDIDInfo(txtDID.Text.Trim(), cboWO.Text.Trim());
                if (txtCompPN.Text == "")
                {
                    return;
                }
                dt = process.QSMS_MCC_DIDSimilarDispByPCB(cboWO.Text.Trim(), txtDID.Text.Trim());
                if (dt.Rows.Count > 0)
                {
                    if (dt.Columns[0].ColumnName != "RESULT")
                    {
                        if (int.Parse(dt.Rows[0]["RemainQty"].ToString().Trim()) > 0 && dt.Rows[0]["DID"].ToString().Trim() != txtDID.Text.Trim())
                        {
                            MessageBox.Show("There are similar DID for your refrence!! ", "Prompt");
                        }
                    }
                }
                //gridSimilarDIDByPCB.DataSource = dt.DefaultView;
                string PreWO = "";
                dt = process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_DispatchByGroup2", "", "", cboWO.Text.Trim(), txtDID.Text.Trim(), "");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (PreWO == "" || PreWO != dt.Rows[i]["Work_Order"].ToString().Trim())
                    {
                        lstAvailableWO.Items.Add(dt.Rows[i]["Work_Order"].ToString().Trim());
                        PreWO = dt.Rows[i]["Work_Order"].ToString().Trim();
                    }
                }
                //gridDIDDispatched.DataSource = dt.DefaultView;
                txtReturnQty.Text = "";
                if (lstAvailableWO.Items.Count == 1)
                {
                    optRatebySelWO.Checked = true;
                    cmdADDALL_Click(sender, e);
                    txtReturnQty.Focus();
                }
            }
        }

        private void txtReturnQty_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && txtReturnQty.Text != "")
            {
                if (pubFunction.IsNumeric(txtReturnQty.Text.Trim(), "INT") == false)
                {
                    MessageBox.Show("Please input numeric!!", "Prompt");
                }
                else
                {
                    cmdSave.Focus();
                }
            }
        }

        private void optRatebySelWO_Click(object sender, EventArgs e)
        {
            setCtrlSelWOStatus(true);
        }

        private void optRatebyPCB_Click(object sender, EventArgs e)
        {
            if (lstAvailableWO.Items.Count + lstCallBackWO.Items.Count == 1)
            {
                MessageBox.Show("This PCB has only one WO dispathed!!", "Prompt");
                optRatebySelWO.Checked = true;
            }
            else
            {
                setCtrlSelWOStatus(false);
            }
        }

        private void optCallAll_Click(object sender, EventArgs e)
        {
            setCtrlSelWOStatus(false);
        }

        private void cmdQuery_Click(object sender, EventArgs e)
        {
            if (cboLine.Text.Trim() == "")
            {
                MessageBox.Show("Please input line");
                return;
            }
            GetGroupID();
        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            try
            {
                string sSelWo, sDID;
                long intReturnQty;
                cmdSave.Enabled = false;
                if (ChkErr() == false)
                {
                    goto Normal_Eixt;
                }
                sDID = txtDID.Text.Trim();
                intReturnQty = Convert.ToInt64(txtReturnQty.Text.Trim());
                if (optRatebySelWO.Checked == true)
                {
                    sSelWo = GetSelWO(lstCallBackWO);
                    if (sSelWo == "")
                        return;
                    if (chkCallBackQty(sDID, sSelWo, Convert.ToInt32(intReturnQty.ToString().Trim()), Convert.ToInt32(txtReturnQty.Text.Trim())) == false)
                    {
                        return;
                    }
                    dt = process.QSMS_MCC_DIDCallBackByType("CallBySelWO", sSelWo, txtCompPN.Text.Trim(), sDID.Trim(), Convert.ToInt32(intReturnQty), Parameter.g_userName, (optGoodMaterial.Checked == true) ? "Y" : "N", IsAnotherBUDID.Trim());
                }
                else if (optRatebyPCB.Checked == true)
                {
                    dt = process.QSMS_MCC_DIDCallBackByType("CallbyPCB", cboWO.Text.Trim(), txtCompPN.Text.Trim(), sDID.Trim(), Convert.ToInt32(intReturnQty), Parameter.g_userName, (optGoodMaterial.Checked == true) ? "Y" : "N", IsAnotherBUDID.Trim());
                }
                else
                {
                    if (MessageBox.Show("DID:" + sDID + ",Total = CallBackQty = " + txtDIDTotalQty.Text.Trim(), ",Are you sure CallBack All", MessageBoxButtons.YesNo) == DialogResult.No)
                    {
                        goto Normal_Eixt;
                    }
                    dt = process.QSMS_MCC_DIDCallBackByType("CallAll", cboWO.Text.Trim(), txtCompPN.Text.Trim(), sDID.Trim(), Convert.ToInt32(intReturnQty), Parameter.g_userName, (optGoodMaterial.Checked == true) ? "Y" : "N", IsAnotherBUDID.Trim());
                }
                if (dt.Rows.Count > 0)
                {
                    LblMessage.Text = dt.Rows[0]["Description"].ToString();
                    if (dt.Rows[0]["Result"].ToString() == "0")
                    {
                        if (pubFunction.ConfigListGetValue("PrtCallBKandReturn") == "Y")
                        {
                            ds = process.QSMS_MCC_XL_DIDGetNewID("CallBack", sDID, (optGoodMaterial.Checked == true) ? "Y" : "N", Convert.ToInt32(intReturnQty), Parameter.g_userName, Parameter.Factory, IsAnotherBUDID);
                            dt = ds.Tables[0];
                            PrintData = dt;
                            if (dt.Rows[0]["Result"].ToString() != "0")
                            {
                                LblMessage.Text = dt.Rows[0]["Description"].ToString();
                            }
                            else
                            {
                                dt = ds.Tables[1];
                                if (dt.Rows.Count <= 0)
                                {
                                    LblMessage.Text = "Get DID information fail,print DID fail!!";
                                    goto Normal_Eixt;
                                }
                                DIDPrintLabel();
                            }
                        }
                    }
                }
                GetDIDInfoCallBack(cboWO.Text.Trim());
            Normal_Eixt:
                {
                    cmdSave.Enabled = true;
                    txtDID.Text = "";
                    txtRemainQty.Text = "";
                    txtDIDTotalQty.Text = "";
                    txtCompPN.Text = "";
                    txtDIDReturnedQty.Text = "";
                    txtDID.Focus();
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Prompt");
                cmdSave.Enabled = true;
            }
        }

        private void cmdGetRefID_Click(object sender, EventArgs e)
        {
            string sMsg, sCurrRefID;
            dt = process.QSMS_MCC_XL_DIDGetRefID("CallBack", (optGoodMaterial.Checked == true) ? "Y" : "N",  Parameter.g_userName, Parameter.Factory, IsAnotherBUDID);
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["Result"].ToString() != "0")
                {
                    MessageBox.Show(dt.Rows[0]["Description"].ToString());
                    return;
                }
                sMsg = dt.Rows[0]["Description"].ToString();
                sCurrRefID = process.QSMS_MCC_QueryDataByType("DIDGetRefIDByResult", "", "", sMsg, "", "").Rows[0]["DIDGetRefIDByResult"].ToString();
                PrintData = dt;
                DIDPrintLabel();
                ds = process.XL_DIDChkStockByRefID_set(sCurrRefID, Parameter.g_userName);
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show(dt.Rows[0]["Description"].ToString(), "Prompt");
                    return;
                }
                frmDIDCheckStock DIDCheckStock = new frmDIDCheckStock();
                DIDCheckStock.Show();
            }
        }

        private void cmdReprint_Click(object sender, EventArgs e)
        {
            Boolean IsByDIDInput = false;
            if (strPrintPort == "" || strCommSetting == "")
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
                if (txtDID.Text.Trim() == gridDIDtoWH.Rows[0].Cells[0].ToString())
                {
                    Parameter.DIDInfo.DID = gridDIDtoWH.Rows[0].Cells[0].ToString();
                    Parameter.DIDInfo.compPN = gridDIDtoWH.Rows[0].Cells[1].ToString();
                    Parameter.DIDInfo.Qty = int.Parse(gridDIDtoWH.Rows[0].Cells[2].ToString());
                    Parameter.DIDInfo.IsGood = gridDIDtoWH.Rows[0].Cells[9].ToString();
                    if (pubFunction.ConfigListGetValue("ChkPrintDIDType") == "Y")
                    {
                        Parameter.DIDInfo.DIDType = gridDIDtoWH.Rows[0].Cells[13].ToString();
                    }
                    else
                    {
                        Parameter.DIDInfo.DIDType = "";
                    }
                }
                else
                {
                    IsByDIDInput = true;
                }
            }
            else
            {
                IsByDIDInput = true;
            }
            if (IsByDIDInput == true)
            {
                dt = process.XL_DIDGetToWHInfo("CallBack", txtDID.Text.Trim(), Parameter.Factory, "N");
                if (dt.Rows.Count <= 0)
                {
                    LblMessage.Text = "There is no DID:" + txtDID.Text.Trim() + " !!";
                    return;
                }
                else
                {
                    PrintData = dt;
                    Parameter.DIDInfo.DID = dt.Rows[0]["DID"].ToString().Trim();
                    Parameter.DIDInfo.compPN = dt.Rows[0]["CompPN"].ToString().Trim();
                    Parameter.DIDInfo.Qty = int.Parse(dt.Rows[0]["Qty"].ToString().Trim());
                    Parameter.DIDInfo.IsGood = dt.Rows[0]["IsGood"].ToString().Trim();
                    if (pubFunction.ConfigListGetValue("ChkPrintDIDType") == "Y")
                    {
                        Parameter.DIDInfo.DIDType = dt.Rows[0]["DIDType"].ToString().Trim();
                    }
                    else
                    {
                        Parameter.DIDInfo.DIDType = "";
                    }
                }
            }
            DIDPrintLabel();
            txtDID.Text = "";   
        }

        private void cmdADD_Click(object sender, EventArgs e)
        {
            int i;
            if (lstAvailableWO.Items.Count <= 0)
                return;
            if (lstAvailableWO.Items.Count < 0)
                return;
            i = lstAvailableWO.SelectedIndex;
            lstCallBackWO.Items.Add(lstAvailableWO.SelectedValue);
            lstAvailableWO.Items.Remove(lstAvailableWO.SelectedValue);
            if (lstAvailableWO.Items.Count > 0)
            {
                if (lstAvailableWO.Items.Count - 1 >= i)
                {
                    lstAvailableWO.SelectedIndex = i;
                }
                else
                {
                    lstAvailableWO.SelectedIndex = lstAvailableWO.Items.Count - 1;
                }
            }
        }

        private void cmdADDALL_Click(object sender, EventArgs e)
        {
            if (lstAvailableWO.Items.Count <= 0)
                return;
            while (lstAvailableWO.Items.Count > 0)
            {
                lstAvailableWO.SelectedIndex = 0;
                lstCallBackWO.Items.Add(lstAvailableWO.SelectedValue);
                lstAvailableWO.Items.Remove(lstAvailableWO.SelectedValue);
            }
        }

        private void cmdDEL_Click(object sender, EventArgs e)
        {
            int i;
            if (lstCallBackWO.Items.Count <= 0)
                return;
            i = lstCallBackWO.SelectedIndex;
            lstAvailableWO.Items.Add(lstCallBackWO.SelectedValue);
            lstCallBackWO.Items.Remove(lstCallBackWO.SelectedValue);
            if (lstCallBackWO.Items.Count > 0)
            {
                if (lstCallBackWO.Items.Count - 1 >= i)
                {
                    lstCallBackWO.SelectedIndex = i;
                }
                else
                {
                    lstCallBackWO.SelectedIndex = lstCallBackWO.Items.Count - 1;
                }
            }
        }

        private void cmdDELALL_Click(object sender, EventArgs e)
        {
            if (lstCallBackWO.Items.Count <= 0)
                return;
            while (lstCallBackWO.Items.Count > 0)
            {
                lstCallBackWO.SelectedIndex = 0;
                lstAvailableWO.Items.Add(lstCallBackWO.SelectedValue);
                lstCallBackWO.Items.Remove(lstCallBackWO.SelectedValue);
            }
        }

        private void GetLine()
        {
            dt = process.QSMS_MCC_QueryDataByType("MCC_GetLine", "", "", "", "", "");
            cboLine.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cboLine.Items.Add(dt.Rows[i]["Line"].ToString().Trim());
            }
        }

        private void GetGroupWO(string GroupID = "")
        {
            dt = process.QSMS_MCC_QueryDataByType("PD_GetWOInfoByGroupID", "", "", "", GroupID, "");
            cboWO.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cboWO.Items.Add(dt.Rows[i]["Work_Order"].ToString().Trim());
            }
        }

        private Boolean ChkGroupClosed(string GroupID = "")
        {
            if (process.QSMS_MCC_QueryDataByType("ChkGroupClosed", "", "", GroupID, "", "").Rows.Count > 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private void GetWOGroupInfo(string WO = "")
        {
            dt = process.QSMS_MCC_GetWOPCBStatus(WO);
            gridWOGroup.DataSource = dt.DefaultView;
            gridWOGroup.Refresh();
        }

        private void GetDIDInfoCallBack(string WO = "")
        {
            rstDIDCalled = process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_DIDCallBack", "", "", WO, "", "");
            gridDIDCalled.DataSource = rstDIDCalled.DefaultView;
            gridDIDCalled.Refresh();

            rstDIDNotCall = process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_DIDCallBack2", "", "", WO, "", "");
            gridDIDNotCall.DataSource = rstDIDNotCall.DefaultView;
            gridDIDNotCall.Refresh();

            rstDID = process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_DIDCallBack2", "", "", WO, "", "");
            gridDIDInfo.DataSource = rstDID.DefaultView;
            gridDIDInfo.Refresh();

            if (Parameter.PrtCallBKandReturn == "Y")
            {
                rstDIDtoWH = process.XL_DIDGetToWHInfo("CallBack", "", Parameter.Factory, "");
                gridDIDtoWH.DataSource = rstDIDtoWH.DefaultView;
                gridDIDtoWH.Refresh();
            }
        }

        private Boolean XL_ChkAnotherBUDID(string sDID)
        {
            ds = process.QSMS_MCC_DIDChkAnotherBU(sDID, IsAnotherBUDID, Parameter.Factory);
            dt = ds.Tables[0];
            if (dt.Rows[0]["Result"].ToString() != "0")
            {
                LblMessage.Text = dt.Rows[0]["Description"].ToString();
                return false;
            }
            else
            {
                dt = ds.Tables[1];
                if (dt.Rows.Count > 0)
                {
                    txtDIDTotalQty.Text = dt.Rows[0]["TotalQty"].ToString().Trim();
                    txtDIDReturnedQty.Text = dt.Rows[0]["ReturnQty"].ToString().Trim();
                    txtCompPN.Text = dt.Rows[0]["compPN"].ToString().Trim();
                    IsAnotherBUDID = dt.Rows[0]["IsAnotherBUDID"].ToString().Trim();
                }
                dt = ds.Tables[2];
                lstAvailableWO.Items.Clear();
                lstCallBackWO.Items.Clear();
                string PreWO = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (PreWO == "" || PreWO != dt.Rows[i]["Work_Order"].ToString().Trim())
                    {
                        lstAvailableWO.Items.Add(dt.Rows[i]["Work_Order"].ToString().Trim());
                        PreWO = dt.Rows[i]["Work_Order"].ToString().Trim();
                    }
                }
                //gridDIDDispatched.DataSource = dt.DefaultView;
                txtReturnQty.Text = "";
                if (lstAvailableWO.Items.Count == 1)
                {
                    optRatebySelWO.Checked = true;
                    cmdADDALL_Click(null, null);
                    txtReturnQty.Focus();
                }
            }
            return true;
        }

        private Boolean ChkDIDBelongToPCB(string WO, string DID)
        {
            dt = process.QSMS_MCC_DIDRestoreForCallBK(WO, DID);

            dt = process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_DIDCallBack4", "", "", WO, DID, "");
            if (dt.Rows.Count <= 0)
            {
                MessageBox.Show("The DID does not belong to the work order,Please check!!");
                return false;
            }
            return true;
        }

        private void GetDIDInfo(string sDID, string sWO)
        {
            ds = process.QSMS_MCC_DIDInfoForCallBK(sWO, sDID);
            dt=ds.Tables[0];
            if (dt.Rows.Count<=0)
            {
                MessageBox.Show("System can not get data,Please Retry or contact QMS!!", "Prompt");
            }
            else
            {
                if (dt.Rows[0]["Result"].ToString().Trim() != "0")
                {
                    MessageBox.Show(dt.Rows[0]["Description"].ToString().Trim(), "Prompt");
                }
                else
                {
                    dt = ds.Tables[1];
                    txtDIDTotalQty.Text = dt.Rows[0]["TotalQty"].ToString().Trim();
                    txtDIDReturnedQty.Text = dt.Rows[0]["ReturnQty"].ToString().Trim();
                    txtCompPN.Text = dt.Rows[0]["compPN"].ToString().Trim();
                    txtRemainQty.Text = dt.Rows[0]["RemainQty"].ToString().Trim();
                }
            }
        }

        private void setCtrlSelWOStatus(Boolean blnEnable)
        {
            lstAvailableWO.Enabled = blnEnable;
            lstCallBackWO.Enabled = blnEnable;
            cmdADD.Enabled = blnEnable;
            cmdADDALL.Enabled = blnEnable;
            cmdDEL.Enabled = blnEnable;
            cmdDELALL.Enabled = blnEnable;
        }

        private void GetGroupID()
        {
            string sSDate, sEDate;
            //sSDate = Convert.ToDateTime(dtpSDate.Text.ToString().Replace("/", "")).ToString("YYYYMMDD");
            //sEDate = Convert.ToDateTime(dtpEDate.Text.ToString().Replace("/", "")).ToString("YYYYMMDD");

            //string BeginDate, EndDate;
            sSDate = dtpSDate.Value.ToString("yyyy/MM/dd");
            sSDate = sSDate.Replace("-", "");
            sSDate = sSDate.Replace("/", "");
            sEDate = dtpEDate.Value.ToString("yyyy/MM/dd"); ;
            sEDate = sEDate.Replace("-", "");
            sEDate = sEDate.Replace("/", "");


            if (OptRelease.Checked == true)
            {
                dt = process.QSMS_MCC_QueryDataByType("PD_GetGroupIDByDate1", sSDate, sEDate, cboLine.Text.Trim(), "", "");
            }
            else
            {
                dt = process.QSMS_MCC_QueryDataByType("PD_GetGroupIDByDate2", sSDate, sEDate, cboLine.Text.Trim(), "", "");
            }
            cboGroupID.Items.Clear();
            if (dt.Rows.Count <= 0)
            {
                MessageBox.Show("No data");
            }
            else
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cboGroupID.Items.Add(dt.Rows[i]["GroupID"].ToString().Trim());
                }
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
            if (cboWO.Text == "")
            {
                return false;
            }
            if (ChkDIDBelongToPCB(cboWO.Text.Trim(), txtDID.Text.Trim()) == false)
            {
                return false;
            }
            if (txtDIDTotalQty.Text == "" || txtCompPN.Text.Trim() == "")
            {
                MessageBox.Show("DID total Qty or comppN Can not be empty,Please press enter key in DID txtbox");
                return false;
            }
            if (txtReturnQty.Text.Trim() == "" || pubFunction.IsNumeric(txtReturnQty.Text.Trim(), "INI") == false)
            {
                MessageBox.Show("The Return Qty can not be empty or must be numeric");
                return false;
            }
            txtReturnQty.Text = ((Convert.ToInt32(txtReturnQty.Text.Trim()) < 0) ? Convert.ToInt32(txtReturnQty.Text.Trim()) * (-1) : Convert.ToInt32(txtReturnQty.Text.Trim())).ToString();
            if (Convert.ToInt32(txtDIDTotalQty.Text) < Convert.ToInt32(txtReturnQty.Text))
            {
                MessageBox.Show("Return Qty can not larger than total qty");
                return false;
            }
            dt = process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_DispatchPCBQty", "", "", cboGroupID.Text.Trim(), txtDID.Text.Trim(), "");
            if (dt.Rows[0][0].ToString() == "")
            {
                MessageBox.Show("This DID did not dispatched!");
                return false;
            }
            dt = process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_DispatchByGroup3", "", "", txtDID.Text.Trim(), cboWO.Text.Trim(), "");
            if (dt.Rows.Count > 0)
            {
                if (Convert.ToInt32(txtReturnQty.Text.Trim()) > Convert.ToInt32(txtRemainQty.Text.Trim()) + Convert.ToInt32(dt.Rows[0]["PCBQty"].ToString()))
                {
                    MessageBox.Show("This DID has dispatched to more than one PCB,CallBack Qty can not larger than the dispatched Qty : " + (Convert.ToInt32(dt.Rows[0]["PCBQty"].ToString()) + Convert.ToInt32(txtRemainQty.Text.Trim()).ToString()).ToString());
                    return false;
                }
            }
            if (strPrintPort == "" || strCommSetting == "")
            {
                MessageBox.Show("Printer have not set!!");
                return false;
            }
            return true;
        }

        private string GetSelWO(ListBox lstB)
        {
            int i;
            string stempWO, GetSelWO;
            stempWO = "";
            GetSelWO = "";
            if (lstB.Items.Count <= 0)
            {
                MessageBox.Show("Please select WO to CallBack!!");
                return GetSelWO;
            }
            for (i = 0; i < lstB.Items.Count - 1; i++)
            {
                lstB.SelectedIndex = i;
                stempWO = stempWO + lstB.SelectedValue + ",";
            }
            stempWO = stempWO.Substring(0, stempWO.Length - 1);
            return stempWO;
        }

        private Boolean chkCallBackQty(string sDID, string sSelWo, long intCallQty, long intRemainQty, string sType = "BySelWO")
        {
            if (IsAnotherBUDID == "Y")
            {
                return true;
            }
            if (sType == "ByPCB")
            {
                dt = process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_DispatchCallQty", "", "", sSelWo, sDID, "");
            }
            else
            {
                dt = process.QSMS_MCC_QueryDataByType("MCC_GetQSMS_DispatchCallQty", "", "", sSelWo, sDID, "");
            }
            if (dt.Rows.Count > 0)
            {
                if (intCallQty - intRemainQty > Convert.ToInt32(dt.Rows[0]["CallQty"].ToString().Trim()))
                {
                    MessageBox.Show("CallBack Qty:" + intCallQty.ToString() + " > Qty:" + intRemainQty + Convert.ToInt32(dt.Rows[0]["CallQty"].ToString().Trim()) + " of these Selected WOs dispatched DID!!", "Prompt");
                    return false;
                }
            }
            else
            {
                return false;
            }
            return true;
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
            if (File.Exists(PrintDIDCallBackLabel) == false)
            {
                MessageBox.Show("File:" + PrintDIDCallBackLabel + " not exists");
                LblMessage.ForeColor = Color.Red;
                return;
            }
            if (string.IsNullOrEmpty(strLabelContent))
            {
                strLabelContent = new StreamReader(PrintDIDCallBackLabel).ReadToEnd();
            }
            if (PrintData.Rows[0]["Qty"].ToString().Trim() == "RefID")
            {
                BU = Parameter.BUDIDShow;
            }
            else
            {
                BU = (IsAnotherBUDID == "Y") ? Parameter.AutoDispatchForAnotherBU : Parameter.BUDIDShow;
            }
            if (lblprint.PrintReturn(strLabelContent, PrintData, BU, ref msg) == false)
            {
                LblMessage.Text = msg;
                LblMessage.ForeColor = Color.Red;
                return;
            }

        }
    }
}
