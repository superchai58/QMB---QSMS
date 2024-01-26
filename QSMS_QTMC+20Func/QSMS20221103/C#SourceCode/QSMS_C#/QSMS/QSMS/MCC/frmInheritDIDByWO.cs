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
    public partial class frmInheritDIDByWO : Form
    {
        DbLibrary.MCC.MCCProcess mccProcess = new DbLibrary.MCC.MCCProcess();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();

        public frmInheritDIDByWO()
        {
            InitializeComponent();
        }

        private void frmInheritDIDByWO_Load(object sender, EventArgs e)
        {
            DataTable dt = mccProcess.XL_GetAllWOInfoList("getLine", "", "", "", "", "", "", "");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                CboLine.Items.Add(dt.Rows[i]["Line"].ToString());
            }
            CboLine.Focus();
            CboLine.Select();
        }

        private void btFind_Click(object sender, EventArgs e)
        {
            try
            {
                if (CboLine.Text == "")
                {
                    errorNotice("线别为空");
                    return;
                }
                CboGroupID.Items.Clear();
                CboInheritingWO.Items.Clear();
                CboInheritWO.Items.Clear();
                CboNotChkBOM.Items.Clear();
                CboNotFinishedWO.Items.Clear();
                CboWO.Items.Clear();
                CboWO.Text = "";
                CboNotFinishedWO.Text = "";
                CboNotChkBOM.Text = "";
                CboInheritWO.Text = "";
                CboInheritingWO.Text = "";
                CboGroupID.Text = "";
                txtWO.Text = "";
                txtMBPN.Text = "";
                txtWOQty.Text = "";
                txtGroup.Text = "";
                string strBeginDate = dptBegin.Text.ToString();
                string strEndDate = dptEnd.Text.ToString();
                DataTable dt = mccProcess.XL_GetAllWOInfoList("getGroupID", "", "", "", "", "", "", CboLine.Text.ToString(), "", strBeginDate, strEndDate);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    CboGroupID.Items.Add(dt.Rows[i]["GroupID"].ToString());
                }
            }
            catch (Exception ex)
            {
                errorNotice(ex.Message.ToString());
            }
        }

        private void errorNotice(string msg)
        {
            pubFunction.Sound("Error");
            lblmsg.Text = msg;
            lblmsg.ForeColor = Color.Red;
        }

        private void CboGroupID_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                DataTable dt_tmp;
                CboWO.Items.Clear();
                CboNotChkBOM.Items.Clear();
                CboNotFinishedWO.Items.Clear();
                CboInheritWO.Items.Clear();
                DataTable dt = mccProcess.XL_GetAllWOInfoList("getGroupWO", "", "", "", "", "", "", "", "", "", "", CboGroupID.Text.ToString());
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dt_tmp = mccProcess.XL_GetAllWOInfoList("CheckMBWO", dt.Rows[i]["Work_Order"].ToString(), "", "", "", "", "", "", "", "", "", "");
                    if (dt_tmp.Rows.Count > 0)
                    {
                        dt_tmp = mccProcess.XL_GetAllWOInfoList("CheckQSMSWO", dt.Rows[i]["Work_Order"].ToString(), "", "", "", "", "", "", "", "", "", "");
                        if (dt_tmp.Rows.Count > 0)
                        {
                            CboInheritWO.Items.Add(dt.Rows[i]["Work_Order"].ToString());
                            dt_tmp = mccProcess.XL_GetAllWOInfoList("ChkWoFinished", dt.Rows[i]["Work_Order"].ToString(), "", "", "", "", "", "", "", "", "", "");
                            if (dt_tmp.Rows[0]["ChkWoFinished"].ToString() == "Y")
                            {
                                CboWO.Items.Add(dt.Rows[i]["Work_Order"].ToString());
                            }
                            else
                            {
                                CboInheritingWO.Items.Add(dt.Rows[i]["Work_Order"].ToString());
                                CboNotFinishedWO.Items.Add(dt.Rows[i]["Work_Order"].ToString());
                            }
                        }
                        else
                        {
                            CboNotChkBOM.Items.Add(dt.Rows[i]["Work_Order"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                errorNotice(ex.Message.ToString());
            }
        }

        private void CboNotChkBOM_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                txtWO.Text = CboNotChkBOM.Text.ToString();
                GetWoinfo(CboNotChkBOM.Text.ToString());
            }
            catch (Exception ex)
            {
                errorNotice(ex.Message.ToString());
            }
        }

        private void CboWO_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                txtWO.Text = CboWO.Text.ToString();
                GetWoinfo(CboWO.Text.ToString());
            }
            catch (Exception ex)
            {
                errorNotice(ex.Message.ToString());
            }
        }

        private void CboNotFinishedWO_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                txtWO.Text = CboNotFinishedWO.Text.ToString();
                GetWoinfo(CboNotFinishedWO.Text.ToString());
            }
            catch (Exception ex)
            {
                errorNotice(ex.Message.ToString());
            }
        }

        private void CboInheritWO_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                txtWO.Text = CboInheritWO.Text.ToString();
                GetWoinfo(CboInheritWO.Text.ToString());
            }
            catch (Exception ex)
            {
                errorNotice(ex.Message.ToString());
            }
        }

        private void CboInheritingWO_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                txtWO.Text = CboInheritingWO.Text.ToString();
                GetWoinfo(CboInheritingWO.Text.ToString());
            }
            catch (Exception ex)
            {
                errorNotice(ex.Message.ToString());
            }
        }

        private void btsave_Click(object sender, EventArgs e)
        {
            try
            {
                int FistTimeofDispatch = 1;
                if (OptMachine.Checked == false && OptSide.Checked == false)
                {
                    errorNotice("请选择承接方式:Side or Machine");
                    return;
                }
                DataTable dt = mccProcess.XL_GetAllWOInfoList("CheckWOGroup", CboInheritWO.Text.ToString(), CboInheritingWO.Text.ToString(), "", "", "", "", "", "", "", "", CboGroupID.Text.ToString());
                if (dt.Rows[0]["CheckWOGroup"].ToString() == "N")
                {
                    errorNotice("承接工单和被承接工单不属于同一个GroupID");
                    return;
                }
                if (chkIncludeXL.Checked)
                {
                    dt = mccProcess.XL_GetAllWOInfoList("getinheritDID_XL", CboInheritWO.Text.ToString(), "", "", "", "", "", "", "", "", "", "");
                }
                else
                {
                    dt = mccProcess.XL_GetAllWOInfoList("getinheritDID", CboInheritWO.Text.ToString(), "", "", "", "", "", "", "", "", "", "");
                }
                if (dt.Rows.Count == 0)
                {
                    errorNotice("没有可以承接的DID");
                    return;
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataTable dtinfo=null;
                    if (OptSide.Checked)
                    {
                        dtinfo = mccProcess.XL_GetAllWOInfoList("getinheritinfo_Side", CboInheritingWO.Text.ToString(), dt.Rows[i]["Machine"].ToString(), "", "", "", dt.Rows[i]["CompPN"].ToString(), CboLine.Text.ToString(), "", "", "", "");
                    }
                    if (OptMachine.Checked)
                    {
                        dtinfo = mccProcess.XL_GetAllWOInfoList("getinheritinfo_Machine", CboInheritingWO.Text.ToString(), dt.Rows[i]["Machine"].ToString(), "", "", "", dt.Rows[i]["CompPN"].ToString(), "", "", "", "", "");
                    }
                    if (dtinfo.Rows.Count > 0 && int.Parse(dt.Rows[i]["RemainQty"].ToString())>0)
                    {
                        for (int j = 0; j < dtinfo.Rows.Count; j++)
                        {
                            DataTable dtChkDIDDispatchedToWo = mccProcess.XL_GetAllWOInfoList("ChkDIDDispatchedToWo", "", dt.Rows[i]["Machine"].ToString(), "", dtinfo.Rows[j]["Slot"].ToString(), dtinfo.Rows[j]["LR"].ToString(), "", "", dt.Rows[i]["DID"].ToString(), "", "", "");
                            if (dtChkDIDDispatchedToWo.Rows[0]["ChkDIDDispatchedToWo"].ToString() == "Y")
                            {
                                int TempDIDQty = 0;
                                if (int.Parse(dt.Rows[i]["RemainQty"].ToString()) + int.Parse(dtinfo.Rows[j]["BalanceQty"].ToString()) > 0)
                                {
                                    TempDIDQty = 0 - int.Parse(dtinfo.Rows[j]["BalanceQty"].ToString());
                                }
                                else
                                {
                                    TempDIDQty = int.Parse(dt.Rows[i]["RemainQty"].ToString());
                                }
                                DataTable dtDispatch = mccProcess.QSMSInsertDispatch(dtinfo.Rows[j]["Work_Order"].ToString(), CboGroupID.Text.ToString(), CboLine.Text.ToString(), dtinfo.Rows[j]["WoQty"].ToString(), dtinfo.Rows[j]["JobPN"].ToString(), dtinfo.Rows[j]["Machine"].ToString(), dt.Rows[i]["CompPN"].ToString(), dtinfo.Rows[j]["Slot"].ToString(), dtinfo.Rows[j]["LR"].ToString(), dtinfo.Rows[j]["BaseQty"].ToString(), dtinfo.Rows[j]["TotalNeedQty"].ToString(), dt.Rows[i]["DID"].ToString(), dt.Rows[i]["Qty"].ToString(), TempDIDQty.ToString(), dt.Rows[i]["VendorCode"].ToString(), dt.Rows[i]["DateCode"].ToString(), dt.Rows[i]["LotCode"].ToString(), Parameter.g_userName, dt.Rows[i]["TransDateTime"].ToString(), CboInheritWO.Text.ToString(), dtinfo.Rows[j]["Item"].ToString(), dtinfo.Rows[j]["JobGroup"].ToString(), dtinfo.Rows[j]["Side"].ToString());
                                if (dt.Rows.Count == 0)
                                {
                                    errorNotice("承接DID异常,请联系QMS");
                                    return;
                                }
                                if (FistTimeofDispatch == 1)
                                {
                                    mccProcess.RecordDispatchFDT(dtinfo.Rows[j]["Work_Order"].ToString());

                                    FistTimeofDispatch = FistTimeofDispatch + 1;
                                }
                            }
                        }
                    }
                }
                mccProcess.RecordDispatchFDT(CboInheritingWO.Text.ToString(), "UpdateMachineFlag");
                mccProcess.RecordDispatchFDT(CboInheritingWO.Text.ToString(), "ChkWOItemFinished");
                lblmsg.Text = "承接OK";
                lblmsg.ForeColor = Color.Green;
            }
            catch (Exception ex)
            {
                errorNotice(ex.Message.ToString());
            }
        }

        private void GetWoinfo(string strWO)
        {
            DataTable dt = mccProcess.XL_GetAllWOInfoList("getWOinfo", strWO, "", "", "", "", "", "", "", "", "", "");
            if (dt.Rows.Count > 0)
            {
                txtMBPN.Text = dt.Rows[0]["PN"].ToString();
                txtWOQty.Text = dt.Rows[0]["Qty"].ToString();
                txtGroup.Text = dt.Rows[0]["Group"].ToString();
            }
        }

        private void GetMachine(string strWO)
        {

            DataTable dt = mccProcess.XL_GetAllWOInfoList("getMachine", strWO, "", "", "", "", "", "", "", "", "", "");
            if (dt.Rows.Count > 0)
            {
                txtMBPN.Text = dt.Rows[0]["PN"].ToString();
                txtWOQty.Text = dt.Rows[0]["Qty"].ToString();
                txtGroup.Text = dt.Rows[0]["Group"].ToString();
            }
        }

    }
}
