using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace QSMS.QSMS.PD
{
    public partial class frmCloseWO : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.PD.PDProcess PD = new DbLibrary.PD.PDProcess();

        public frmCloseWO()
        {
            InitializeComponent();
        }

        private void frmCloseWO_Load(object sender, EventArgs e)
        {
            frameCHK.Visible = false;
            OptRelease.Checked = false;
            OptGroup.Checked = false;
            FraSB.Visible = false;
            DTPEndDate.Text = DateTime.Today.ToString("yyyy/MM/dd");
            DTPBeginDate.Text = DateTime.Today.AddDays(-1).ToString("yyyy/MM/dd");
            GetLine();
            GetAuthority();
        }

        private void frmCloseWO_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmCloseWO");
        }

        private void GetLine()
        {
            DataTable dt = PD.QSMS_PD_QueryDataByType("PD_GetLine", "", "", "", "","");
            cboLine.Text = "";
            cboLine.Items.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cboLine.Items.Add(dt.Rows[i]["Line"].ToString().ToUpper());
            }
        }

        private void GetAuthority()
        {
            for (int i = Parameter.g_userRight.GetLowerBound(0); i <= Parameter.g_userRight.GetUpperBound(0); i++)
            {
                if (Parameter.g_userRight[i] == "PowerCloseWO")
                {
                    frameCHK.Visible = true;
                    return;
                }
                if (Parameter.g_userRight[i] == "UnChkDispCloseWO")
                {
                    frameCHK.Visible = true;
                    checkBox2.Enabled = false;
                    checkBox3.Enabled = false;
                    checkBox4.Enabled = false;
                }
            }
        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            if (cboLine.Text == "")
            {
                MessageBox.Show("Please select Line!");
                return;
            }
            GetGroupID();
        }

        private void GetGroupID()
        {
            string BeginDate = Convert.ToDateTime(DTPBeginDate.Text).ToString("yyyyMMdd");
            string EndDate = Convert.ToDateTime(DTPEndDate.Text).ToString("yyyyMMdd");
            DataTable dt = new DataTable();
            if (Parameter.BU == "NB5")
            {
                if (OptRelease.Checked == true)
                {
                    dt = PD.QSMS_PD_QueryDataByType("PD_GetGroupIDByDate1_NB5", BeginDate, EndDate, cboLine.Text, "","");
                }
                else
                {
                    dt = PD.QSMS_PD_QueryDataByType("PD_GetGroupIDByDate2_NB5", BeginDate, EndDate, cboLine.Text, "", "");
                }
            }
            else
            {
                if (OptRelease.Checked == true)
                {
                    dt = PD.QSMS_PD_QueryDataByType("PD_GetGroupIDByDate1", BeginDate, EndDate, cboLine.Text, "", "");
                }
                else
                {
                    dt = PD.QSMS_PD_QueryDataByType("PD_GetGroupIDByDate2", BeginDate, EndDate, cboLine.Text, "", "");
                }
            }
            cboGroupID.Text = "";
            cboGroupID.Items.Clear();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cboGroupID.Items.Add(dt.Rows[i]["GroupID"].ToString().ToUpper());
                }
            }
            else
            {
                MessageBox.Show("No Data!");
                return;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            btnSave.Enabled = false;
            if (lstWO_SELECT.Items.Count <= 0)
            {
                btnSave.Enabled = true;
                return;
            }
            while (lstWO_SELECT.Items.Count > 0)
            {
                lstWO_SELECT.Items[0].Selected = true;
                if (MessageBox.Show("do you make sure to close the work order by manual " + lstWO_SELECT.Items[0].Text.ToString(), "提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    if (CloseWoByManual(lstWO_SELECT.Items[0].Text.ToString(), "Manual") == true)
                    {
                        ListViewItem lst = new ListViewItem(lstWO_SELECT.Items[0].SubItems[0].Text);
                        lstWOClosed.Items.Add(lst);
                        lstWO_SELECT.Items.RemoveAt(0);
                    }
                    else
                    {
                        btnSave.Enabled = true;
                        return;
                    }
                }
                else
                {
                    btnSave.Enabled = true;
                    return;
                }
            }
            btnSave.Enabled = true;
        }

        //private bool CloseWoByManual(string WO, string CloseType)//共用方法
        //{
        //    string Dispatch_Flag = string.Empty;
        //    string AOI_Flag = string.Empty;
        //    string SAP1_Flag = string.Empty;
        //    string SAP2_Flag = string.Empty;
        //    string XBoardQtyInput = string.Empty;
        //    int XBoardQty = 0;
        //    try
        //    {
        //        if (checkBox1.Checked == false)
        //        {
        //            Dispatch_Flag = "N";
        //        }
        //        else
        //        {
        //            Dispatch_Flag = "Y";
        //        }
        //        if (checkBox2.Checked == false)
        //        {
        //            AOI_Flag = "N";
        //        }
        //        else
        //        {
        //            AOI_Flag = "Y";
        //        }
        //        if (checkBox3.Checked == false)
        //        {
        //            SAP1_Flag = "N";
        //        }
        //        else
        //        {
        //            SAP1_Flag = "Y";
        //        }
        //        if (checkBox4.Checked == false)
        //        {
        //            SAP2_Flag = "N";
        //        }
        //        else
        //        {
        //            SAP2_Flag = "Y";
        //        }
        //        DataTable dt = PD.QSMSChkCloseWOByManual(WO, Dispatch_Flag, AOI_Flag, SAP1_Flag, SAP2_Flag);
        //        if (dt.Rows.Count > 0)
        //        {
        //            if (dt.Rows[0][0].ToString().ToUpper() != "PASS")
        //            {
        //                pubFunction.doExport(dt);
        //                MessageBox.Show(dt.Rows[0][0].ToString());
        //                return false;
        //            }
        //        }
        //        //Check If have any DID need to return
        //        dt = PD.QSMS_WONeedReturnDID(WO);
        //        if (dt.Rows.Count > 0)
        //        {
        //            DialogResult strResultNB5;
        //            if (Parameter.BU == "NB5" || Parameter.BU == "NB3")
        //            {
        //                strResultNB5 = MessageBox.Show("There are some DID need to return by the WO,please check!!" + "\r\n" +
        //                    "1.[Yes]close WO and delete DID;" + "\r\n" +
        //                    "2.[No]close WO and not delete DID;" + "\r\n" +
        //                    "3.[Cancel]Do not close WO", "Message!", MessageBoxButtons.YesNoCancel);
        //                if (strResultNB5 == DialogResult.Yes)
        //                {
        //                    PD.QSMSCloseWODelDID(WO);
        //                }
        //                else if (strResultNB5 == DialogResult.Cancel)
        //                {
        //                    pubFunction.doExport(dt);
        //                    return false;
        //                }
        //            }
        //            else
        //            {
        //                if (MessageBox.Show("There are some DID need to return by the WO,please check the result!!Do you still want to Close the WO(If you close the WO, the DID will be delete which need to return!!)?", "Message!", MessageBoxButtons.OKCancel) == DialogResult.OK)
        //                {
        //                    PD.QSMSCloseWODelDID(WO);
        //                }
        //                else
        //                {
        //                    pubFunction.doExport(dt);
        //                    return false;
        //                }
        //            }
        //        }
        //        //MBU Xborad自动收缩C和S面材料使用量，发料量，需求量
        //        if (pubFunction.ConfigListGetValue("CheckWOIFReduceXboard") == "Y")
        //        {
        //            dt = PD.QSMS_CloseWO_CheckWOIFReduceXboard(WO);
        //            if (dt.Rows.Count > 0)
        //            {
        //                if (dt.Rows[0][0].ToString() == "1")
        //                {
        //                    XBoardQtyInput = Interaction.InputBox("请输入" + WO + "的X板数量：", "CloseWO", "0");
        //                    XBoardQty = int.Parse(XBoardQtyInput);
        //                    dt = PD.QSMS_CloseWO_ReduceXboard(WO, XBoardQty);
        //                    if (dt.Rows.Count > 0)
        //                    {
        //                        if (dt.Rows[0][0].ToString().ToUpper() != "PASS")
        //                        {
        //                            pubFunction.doExport(dt);
        //                            MessageBox.Show(dt.Rows[0][0].ToString());
        //                            return false;
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        //send SAP1 data ,include lost data and sended more data
        //        dt = PD.QSMS_SapCostPacking(WO, CloseType);
        //        if (dt.Rows.Count > 0)
        //        {
        //            if (dt.Rows[0][0].ToString().ToUpper() != "PASS")
        //            {
        //                pubFunction.doExport(dt);
        //                MessageBox.Show(dt.Rows[0][0].ToString());
        //                return false;
        //            }
        //        }
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message + " Please Call QMS!");
        //        return false;
        //    }
        //}
        private bool CloseWoByManual(string WO, string CloseType)//共用方法
        {
            string Dispatch_Flag = string.Empty;
            string AOI_Flag = string.Empty;
            string SAP1_Flag = string.Empty;
            string SAP2_Flag = string.Empty;
            string XBoardQtyInput = string.Empty;
            int XBoardQty = 0;
            try
            {
                if (checkBox1.Checked == false)
                {
                    Dispatch_Flag = "N";
                }
                else
                {
                    Dispatch_Flag = "Y";
                }
                if (checkBox2.Checked == false)
                {
                    AOI_Flag = "N";
                }
                else
                {
                    AOI_Flag = "Y";
                }
                if (checkBox3.Checked == false)
                {
                    SAP1_Flag = "N";
                }
                else
                {
                    SAP1_Flag = "Y";
                }
                if (checkBox4.Checked == false)
                {
                    SAP2_Flag = "N";
                }
                else
                {
                    SAP2_Flag = "Y";
                }
                DataTable dt = PD.QSMSChkCloseWOByManual(WO, Dispatch_Flag, AOI_Flag, SAP1_Flag, SAP2_Flag);
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0][0].ToString().ToUpper() != "PASS")
                    {
                        pubFunction.doExport(dt);
                        MessageBox.Show(dt.Rows[0][0].ToString());
                        return false;
                    }
                }
                //Check If have any DID need to return
                dt = PD.QSMS_WONeedReturnDID(WO);
                if (dt.Rows.Count > 0)
                {
                    DialogResult strResultNB5;
                    if (Parameter.BU == "NB5" || Parameter.BU == "NB3")
                    {
                        if (pubFunction.ConfigListGetValue("CancelDeleteDID") == "Y")  //0001
                        {
                            strResultNB5 = DialogResult.No;
                        }
                        else
                        {
                            strResultNB5 = MessageBox.Show("There are some DID need to return by the WO,please check!!" + "\r\n" +
                          "1.[Yes]close WO and delete DID;" + "\r\n" +
                          "2.[No]close WO and not delete DID;" + "\r\n" +
                          "3.[Cancel]Do not close WO", "Message!", MessageBoxButtons.YesNoCancel);
                        }

                        if (strResultNB5 == DialogResult.Yes)
                        {
                            PD.QSMSCloseWODelDID(WO);
                        }
                        else if (strResultNB5 == DialogResult.Cancel)
                        {
                            pubFunction.doExport(dt);
                            return false;
                        }
                    }
                    else
                    {
                        //if (pubFunction.ConfigListGetValue("CancelDeleteDID") == "Y")  //0001
                        //{
                        //    pubFunction.doExport(dt);
                        //    return false;
                        //}
                        //else
                        //{

                            if (MessageBox.Show("There are some DID need to return by the WO,please check the result!!Do you still want to Close the WO(If you close the WO, the DID will be delete which need to return!!)?", "Message!", MessageBoxButtons.OKCancel) == DialogResult.OK)
                            {
                                PD.QSMSCloseWODelDID(WO);
                            }
                            else
                            {
                                pubFunction.doExport(dt);
                                return false;
                            }
                        //}

                    }
                }
                //MBU Xborad自动收缩C和S面材料使用量，发料量，需求量
                if (pubFunction.ConfigListGetValue("CheckWOIFReduceXboard") == "Y")
                {
                    dt = PD.QSMS_CloseWO_CheckWOIFReduceXboard(WO);
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0][0].ToString() == "1")
                        {
                            XBoardQtyInput = Interaction.InputBox("请输入" + WO + "的X板数量：", "CloseWO", "0");
                            XBoardQty = int.Parse(XBoardQtyInput);
                            dt = PD.QSMS_CloseWO_ReduceXboard(WO, XBoardQty);
                            if (dt.Rows.Count > 0)
                            {
                                if (dt.Rows[0][0].ToString().ToUpper() != "PASS")
                                {
                                    pubFunction.doExport(dt);
                                    MessageBox.Show(dt.Rows[0][0].ToString());
                                    return false;
                                }
                            }
                        }
                    }
                }
                //send SAP1 data ,include lost data and sended more data
                dt = PD.QSMS_SapCostPacking(WO, CloseType);
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0][0].ToString().ToUpper() != "PASS")
                    {
                        pubFunction.doExport(dt);
                        MessageBox.Show(dt.Rows[0][0].ToString());
                        return false;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " Please Call QMS!");
                return false;
            }
        }

        private bool ChkMBWo(string WO)//共用方法
        {
            DataTable dt = PD.QSMS_PD_QueryDataByType("PD_ChkMBWo", "", "", "", WO, "");
            if (dt.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }

        private bool ChkIfSmallBoardExist(string WO)
        {
            DataTable dt = PD.QSMS_PD_QueryDataByType("PD_ChkIfSmallBoardExist", "", "", "", WO, "");
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < lstWOUnClose.Items.Count; j++)
                    {
                        if (lstWOUnClose.Items[j].Text.ToString() == dt.Rows[i]["Wo"].ToString())
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        private void GetGroupWO(string GroupID)
        {
            if (GroupID == "")
            {
                return;
            }
            DataTable dt = PD.QSMS_PD_QueryDataByType("PD_GetWOInfoByGroupID", "", "", "", GroupID, "");
            lstWOClosed.Items.Clear();
            lstWOUnClose.Items.Clear();
            lstWO_SELECT.Items.Clear();
            cboWO.Text = "";
            cboWO.Items.Clear();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (ChkMBWo(dt.Rows[i]["Work_Order"].ToString()) == true)
                    {
                        if (ChkIfSmallBoardExist(dt.Rows[i]["Work_Order"].ToString()) == false)
                        {
                            cboWO.Items.Add(dt.Rows[i]["Work_Order"].ToString());
                            ListViewItem lst = new ListViewItem(dt.Rows[i]["Work_Order"].ToString());
                            if (dt.Rows[i]["ClosedFlag"].ToString().Trim().ToUpper() == "Y")
                            {

                                lstWOClosed.Items.Add(lst);
                            }
                            else
                            {
                                lstWOUnClose.Items.Add(lst);
                            }
                        }
                    }
                }
            }
            if (lstWOUnClose.Items.Count <= 0)
            {
                return;
            }
            this.lstWOUnClose.Focus();
            this.lstWOUnClose.Items[lstWOUnClose.Items.Count - 1].Selected = true;
        }

        private void cboGroupID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && cboGroupID.Text != "")
            {
                cboGroupID_SelectedIndexChanged(sender, e);
                DeleteWOQSMS(cboGroupID.Text);
            }
        }

        private void DeleteWOQSMS(string GroupID)
        {
            string strWO = string.Empty;
            DataTable dt = PD.QSMS_PD_QueryDataByType("PD_GetWOByGroupID", "", "", "", GroupID, "");
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    strWO = dt.Rows[i]["Work_Order"].ToString() + ";" + strWO;
                }
                if (MessageBox.Show("The SMT part have been deleted of the WO: " + strWO + ",do you want to delete the QSMS part of the WO ?", "提示", MessageBoxButtons.OKCancel) == DialogResult.OK)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        PD.PD_QSMS_DelWO(dt.Rows[i]["Work_Order"].ToString());
                    }
                }
            }

        }

        private void GetWoinfo(string WO)
        {
            DataTable dt = PD.QSMS_PD_QueryDataByType("PD_GetWOInfoByWO", "", "", "", WO, "");
            if (dt.Rows.Count > 0)
            {
                txtMBPN.Text = dt.Rows[0]["PN"].ToString().ToUpper();
                txtWOQty.Text = dt.Rows[0]["Qty"].ToString();
                txtCustomer.Text = dt.Rows[0]["Customer"].ToString().ToUpper();
            }
        }

        private void cboWO_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtWO.Text = cboWO.Text;
            GetWoinfo(txtWO.Text);
            GetSBWO(txtWO.Text);
        }

        private void cboGroupID_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetGroupWO(cboGroupID.Text.ToString());
        }

        private void GetSBWO(string WO)
        {
            int I = 0;
            cboSBWO.Text = "";
            cboSBWO.Items.Clear();
            FraSB.Visible = false;
            DataTable dt = PD.QSMS_PD_QueryDataByType("PD_GetSBWO", "", "", "", WO, "");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cboSBWO.Items.Add(dt.Rows[i]["WO"].ToString().Trim());
                I = I + 1;
            }
            if (I > 0)
            {
                FraSB.Visible = true;
            }
        }

        private void cboWO_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && cboWO.Text != "")
            {
                cboWO_SelectedIndexChanged(sender, e);
            }
        }

        private void checkBox1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == false)
            {
                if (MessageBox.Show("Are you sure to Un-check dispatch ?", "提示", MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    checkBox1.Checked = true;
                    return;
                }
            }
            if (checkBox1.Checked == false)
            {
                checkBox1.ForeColor = Color.Red;
            }
            else
            {
                checkBox1.ForeColor = Color.Black;
            }
        }

        private void checkBox2_Click(object sender, EventArgs e)
        {
            if (checkBox2.Checked == false)
            {
                if (MessageBox.Show("Are you sure to Un-check AOI Qty ?", "提示", MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    checkBox2.Checked = true;
                    return;
                }
            }
            if (checkBox2.Checked == false)
            {
                checkBox2.ForeColor = Color.Red;
            }
            else
            {
                checkBox2.ForeColor = Color.Black;
            }
        }

        private void checkBox3_Click(object sender, EventArgs e)
        {
            if (checkBox3.Checked == false)
            {
                if (MessageBox.Show("Are you sure to Un-check if has been sent SAP1 ?", "提示", MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    checkBox3.Checked = true;
                    return;
                }
            }
            if (checkBox3.Checked == false)
            {
                checkBox3.ForeColor = Color.Red;
            }
            else
            {
                checkBox3.ForeColor = Color.Black;
            }
        }

        private void checkBox4_Click(object sender, EventArgs e)
        {
            if (checkBox4.Checked == false)
            {
                if (MessageBox.Show("Are you sure to Un-check if has been sent SAP2 ?", "提示", MessageBoxButtons.YesNo) == DialogResult.No)
                {
                    checkBox4.Checked = true;
                    return;
                }
            }
            if (checkBox4.Checked == false)
            {
                checkBox4.ForeColor = Color.Red;
            }
            else
            {
                checkBox4.ForeColor = Color.Black;
            }
        }

        private void cmdADD_Click(object sender, EventArgs e)
        {
            if (lstWOUnClose.SelectedItems.Count <= 0)
            {
                return;
            }
            if (lstWOUnClose.Items.Count <= 0)
            {
                return;
            }
            int index = lstWOUnClose.SelectedItems[0].Index;
            if (index < 0)
            {
                return;
            }
            ListViewItem lst = new ListViewItem(lstWOUnClose.Items[index].SubItems[0].Text);
            lstWO_SELECT.Items.Add(lst);
            lstWOUnClose.Items.RemoveAt(index);
            if (lstWOUnClose.Items.Count != index)
            {
                this.lstWOUnClose.Focus();
                this.lstWOUnClose.Items[0].Selected = true;
            }
        }

        private void cmdADDALL_Click(object sender, EventArgs e)
        {
            if (lstWOUnClose.Items.Count <= 0)
            {
                return;
            }
            for (int i = 0; i < lstWOUnClose.Items.Count; i++)
            {
                ListViewItem lst = new ListViewItem(lstWOUnClose.Items[i].SubItems[0].Text);
                lstWO_SELECT.Items.Add(lst);
            }
            lstWOUnClose.Items.Clear();
        }

        private void cmdDEL_Click(object sender, EventArgs e)
        {
            if (lstWO_SELECT.SelectedItems.Count <= 0)
            {
                return;
            }
            if (lstWO_SELECT.Items.Count <= 0)
            {
                return;
            }
            int index = lstWO_SELECT.SelectedItems[0].Index;
            if (index < 0)
            {
                return;
            }
            ListViewItem lst = new ListViewItem(lstWO_SELECT.Items[index].SubItems[0].Text);
            lstWOUnClose.Items.Add(lst);
            lstWO_SELECT.Items.RemoveAt(index);
            if (lstWO_SELECT.Items.Count != index)
            {
                this.lstWO_SELECT.Focus();
                this.lstWO_SELECT.Items[index].Selected = true;
            }
        }

        private void cmdDELALL_Click(object sender, EventArgs e)
        {
            if (lstWO_SELECT.Items.Count <= 0)
            {
                return;
            }
            for (int i = 0; i < lstWO_SELECT.Items.Count; i++)
            {
                ListViewItem lst = new ListViewItem(lstWO_SELECT.Items[i].SubItems[0].Text);
                lstWOUnClose.Items.Add(lst);
            }
            lstWO_SELECT.Items.Clear();
        }

    }
}
