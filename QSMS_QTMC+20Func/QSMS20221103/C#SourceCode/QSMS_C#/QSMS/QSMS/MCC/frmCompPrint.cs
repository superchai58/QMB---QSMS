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

namespace QSMS.QSMS.MCC
{
    public partial class frmCompPrint : Form
    {
        string msg = string.Empty;
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.MCC.MCCProcess MCC = new DbLibrary.MCC.MCCProcess();
        PrintLabel Print = new PrintLabel();

        private string strPrintPort = string.Empty;
        private string strCommSetting = string.Empty;
        private string strDIDPrintLabel = string.Empty;
        private string strISUNID = "N";

        public frmCompPrint()
        {
            InitializeComponent();
        }

        private void frmCompPrint_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmCompPrint");
        }

        private void frmCompPrint_Load(object sender, EventArgs e)
        {
            reFreshData();
            txtUserID.Text = Parameter.g_userName;
            txtUserID.ForeColor = Color.Red;
            if (Parameter.BU == "NB5")
            {
                groupboxKF.Visible = false;
            }
            
            strPrintPort = pubFunction.ConfigListGetValue("PrintPort");
            strCommSetting = pubFunction.ConfigListGetValue("CommSetting");

            if (Parameter.BU == "NB6")
            {
                lblLinkID.Visible = true;
                cboLinkID.Visible = true;
            }
            else
            {
                lblLinkID.Visible = false;
                cboLinkID.Visible = false;
            }
        }

        private void reFreshData()
        {
            DataTable dt = MCC.QSMS_MCC_QueryDataByType("MCC_GetCompPrintLog", "", "", "", "", "");
            DG_Result.DataSource = dt;
        }

        private void txtCompPN_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && txtCompPN.Text != "")
            {
                try
                {
                    if (Parameter.BU == "NB6")//手动刷comppn Ada 0002
                    {
                        DataTable dt = MCC.QSMS_MCC_QueryVendorCode(txtCompPN.Text,"");
                        if (dt.Rows.Count ==0)
                        {
                            lblMsg.Text = "Please Check  Uniupload --> 上传SAP_CompPN_Info!!";
                            txtCompPN.Text = "";
                            txtCompPN.Focus();
                            return;
                        }

                        lblMsg.Text = "";
                        comboBox1.Items.Clear();
                        comboBox1.Items.Add("");
                        comboBox1.Text = "";
                        foreach (DataRow dr in dt.Rows)
                        {
                            comboBox1.Items.Add(dr[0].ToString());
                        }

                        if (Parameter.BU == "NB6")//10001
                        {
                            DataTable dtall = MCC.QSMS_MCC_QueryVendorCode(txtCompPN.Text, "all");
                            txtCompPN.Text = dtall.Rows[0]["CompPN"].ToString();
                        }
                        else
                        {
                            DataTable dtall = MCC.QSMS_MCC_QueryVendorCode(txtCompPN.Text, "all");
                            txtSpec.Text = txtCompPN.Text;
                            txtCompPN.Text = dtall.Rows[0]["CompPN"].ToString();
                        }

                        DataTable dtmsd = MCC.QSMS_MCC_QueryVendorCode(txtCompPN.Text, "MSD");
                        foreach (DataRow dr in dtmsd.Rows)
                        {
                            comboBox2.Items.Add(dr[0].ToString());
                        }

                        //10001 begin
                        if (Parameter.BU == "NB6")
                        {
                            DataTable dtGetSize = MCC.QSMS_MCC_QueryVendorCode(txtCompPN.Text.Trim(), "GetSize");
                            if (dtGetSize.Rows.Count != 0)
                            {
                                txtSpec.Text = dtGetSize.Rows[0][0].ToString();
                            }
                        }
                        //10001 end
                        txtSpec.Focus();                     
                    }

                    txtVendorCode.Text = "";                 
                    txtDateCode.Focus();
                }
                catch (Exception ex)
                {
                    lblMsg.Text = ex.Message;
                    lblMsg.ForeColor = Color.Red;
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)//手动刷comppn Ada 0002
        {
            txtVendorCode.Text = this.comboBox1.SelectedItem.ToString();
        }

        private void txtUNID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && txtUNID.Text != "")
            {
                try
                {
                    DataTable dt;
                    if (txtUNID.Text.IndexOf(";") > 0)
                    {
                        dt = null;
                        dt = MCC.QSMS_GenUNID(txtUNID.Text, "");
                        if (dt.Rows.Count > 0)
                        {
                            if (dt.Rows[0]["Result"].ToString().ToUpper() == "OK")
                            {
                                txtUniqueID.Enabled = true;
                                txtCompPN.Text = dt.Rows[0]["CompPN"].ToString().ToUpper();
                                txtVendorCode.Text = dt.Rows[0]["VendorCode"].ToString().ToUpper();
                                txtDateCode.Text = dt.Rows[0]["DateCode"].ToString().ToUpper();
                                txtLotCode.Text = dt.Rows[0]["LotCode"].ToString().ToUpper();
                                txtUNID.Text = dt.Rows[0]["UNID"].ToString().ToUpper();
                                txtQty.Text = dt.Rows[0]["Qty"].ToString().ToUpper();
                                txtSpec.Text = dt.Rows[0]["Spec"].ToString().ToUpper();
                                txtMfrSite.Text = dt.Rows[0]["MfrSite"].ToString().ToUpper();
                                txtMark.Text = dt.Rows[0]["Mark"].ToString().ToUpper();
                                txtUniqueID.Text = dt.Rows[0]["UniqueID"].ToString().ToUpper();
                                if (dt.Rows[0]["UNID"].ToString() != "")
                                {
                                    strISUNID = "Y";
                                }
                            }
                            else
                            {
                                lblMsg.Text = dt.Rows[0]["Msg"].ToString().ToUpper();
                                lblMsg.ForeColor = Color.Red;
                            }
                            txtUniqueID.Enabled = false;

                        }
                    }
                    else
                    {
                        dt = null;
                        dt = MCC.QSMS_PrintDID(txtUNID.Text);
                        //10001 begin
                        //if (dt.Rows.Count > 0)
                        //{
                        //    txtCompPN.Text = dt.Rows[0]["CompPN"].ToString().ToUpper();
                        //    txtVendorCode.Text = dt.Rows[0]["VendorCode"].ToString().ToUpper();
                        //    txtDateCode.Text = dt.Rows[0]["DateCode"].ToString().ToUpper();
                        //    txtLotCode.Text = dt.Rows[0]["LotCode"].ToString().ToUpper();
                        //}
                        if (Parameter.BU == "NB6")
                        {
                            if (dt.Rows.Count > 0 )
                            {
                                if (dt.Rows[0][0].ToString().ToUpper() == "OK")
                                {
                                    txtCompPN.Text = dt.Rows[0]["CompPN"].ToString().ToUpper();
                                    txtVendorCode.Text = dt.Rows[0]["VendorCode"].ToString().ToUpper();
                                    txtDateCode.Text = dt.Rows[0]["DateCode"].ToString().ToUpper();
                                    txtLotCode.Text = dt.Rows[0]["LotCode"].ToString().ToUpper();
                                    txtUNID.Text = dt.Rows[0]["UNID"].ToString().ToUpper();
                                    txtQty.Text = dt.Rows[0]["Qty"].ToString().ToUpper();
                                    txtSpec.Text = dt.Rows[0]["Spec"].ToString().ToUpper();
                                    txtMfrSite.Text = dt.Rows[0]["MfrSite"].ToString().ToUpper();
                                    txtMark.Text = dt.Rows[0]["Mark"].ToString().ToUpper();
                                    txtUniqueID.Text = dt.Rows[0]["UniqueID"].ToString().ToUpper();
                                }
                                else
                                {
                                    MessageBox.Show(dt.Rows[0][1].ToString());
                                }
                            }
                        }
                        else
                        {
                            if (dt.Rows.Count > 0)
                            {
                                txtCompPN.Text = dt.Rows[0]["CompPN"].ToString().ToUpper();
                                txtVendorCode.Text = dt.Rows[0]["VendorCode"].ToString().ToUpper();
                                txtDateCode.Text = dt.Rows[0]["DateCode"].ToString().ToUpper();
                                txtLotCode.Text = dt.Rows[0]["LotCode"].ToString().ToUpper();
                            }
                        }
                        //10001 end
                    }
                    //txtQty.Text = "";
                    txtQty.Focus();
                    //txtQty.SelectAll();
                    //txtMark.Text = "";
                    //txtMark.Focus();
                }
                catch (Exception ex)
                {
                    lblMsg.Text = ex.Message + ",请联系QMS!";
                }
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

        private void txtQty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && txtQty.Text != "")
            {
                if (pubFunction.IsNumeric(txtQty.Text, "INT") == false)
                {
                    lblMsg.Text = "请输入正确的数字!";
                    txtQty.Text = "";
                    txtQty.Focus();
                    return;
                }
                btnPrint.Focus();
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            string Msg = string.Empty;
            try
            {
                if (ValidateData() == false)
                {
                    return;
                }

                if (strISUNID == "Y")
                {
                    //strDIDPrintLabel = pubFunction.ConfigListGetValue("UNIDLabel");
                    strDIDPrintLabel = Application.StartupPath + "\\" + pubFunction.ConfigListGetValue("UNIDLabel");
                }
                else
                {
                    //strDIDPrintLabel = pubFunction.ConfigListGetValue("CompPrintLabel");
                    strDIDPrintLabel = Application.StartupPath + "\\" + pubFunction.ConfigListGetValue("CompPrintLabel");
                }

                if (File.Exists(strDIDPrintLabel) == false)
                {
                    lblMsg.Text = "在路径[" + strDIDPrintLabel + "]没找到对应模板!";
                    return;
                }
                StreamReader reader = new StreamReader(strDIDPrintLabel, Encoding.Default);
                string tmpPrintStr = reader.ReadToEnd();
                reader.Close();
                tmpPrintStr = tmpPrintStr.ToUpper();
                if (Print.LabelSetting(strCommSetting, strPrintPort, 1, ref Msg) == false)
                {
                    lblMsg.Text = Msg;
                    return;
                }
                //10001 begin
                //DataTable dt = MCC.QSMS_SaveCompPrintLog("GetPrintInfo", txtCompPN.Text.Trim().ToUpper(), txtQty.Text.Trim(), txtVendorCode.Text.Trim().ToUpper(), txtDateCode.Text.Trim().ToUpper(), txtLotCode.Text.Trim().ToUpper(),
                //                                                                          Parameter.g_userName, txtMark.Text.Trim().ToUpper(), txtUNID.Text.Trim().ToUpper(), "1", txtUniqueID.Text.Trim(), txtSpec.Text.Trim(), txtMfrSite.Text.Trim());

                DataTable dt = new DataTable();
                if (Parameter.BU == "NB6")
                {
                     dt = MCC.QSMS_SaveCompPrintLog("GetPrintInfo", txtCompPN.Text.Trim().ToUpper(), txtQty.Text.Trim(), txtVendorCode.Text.Trim().ToUpper(),txtDateCode.Text.Trim().ToUpper(), txtLotCode.Text.Trim().ToUpper(), 
                                                                                              Parameter.g_userName, txtMark.Text.Trim().ToUpper(), txtUNID.Text.Trim().ToUpper(), "1", txtUniqueID.Text.Trim(), txtSpec.Text.Trim(), txtMfrSite.Text.Trim(), cboLinkID.Text.Trim());
                }
                else
                {
                     dt = MCC.QSMS_SaveCompPrintLog("GetPrintInfo", txtCompPN.Text.Trim().ToUpper(), txtQty.Text.Trim(), txtVendorCode.Text.Trim().ToUpper(),txtDateCode.Text.Trim().ToUpper(), txtLotCode.Text.Trim().ToUpper(), 
                                                                                              Parameter.g_userName, txtMark.Text.Trim().ToUpper(), txtUNID.Text.Trim().ToUpper(), "1", txtUniqueID.Text.Trim(), txtSpec.Text.Trim(), txtMfrSite.Text.Trim());
                }
                //10001 end

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataTable dtPrint = null;
                    dtPrint = dt.Clone();
                    dtPrint.Clear();
                    dtPrint.ImportRow(dt.Rows[i]);
                    if (Print.Print(tmpPrintStr, dtPrint, ref Msg) == false)
                    {
                        lblMsg.Text = Msg;
                        return;
                    }

                    //10001 begin
                    if (Parameter.BU == "NB6")
                    {
                        MCC.QSMS_SaveCompPrintLog("SaveLog", txtCompPN.Text.Trim().ToUpper(), txtQty.Text.Trim(), txtVendorCode.Text.Trim().ToUpper(), txtDateCode.Text.Trim().ToUpper(), txtLotCode.Text.Trim().ToUpper(), Parameter.g_userName, txtMark.Text.Trim().ToUpper(), txtUNID.Text.Trim().ToUpper(), "1", txtUniqueID.Text.Trim(), txtSpec.Text.Trim(), txtMfrSite.Text.Trim(), "");
                    }
                    else
                    {
                        MCC.QSMS_SaveCompPrintLog("SaveLog", txtCompPN.Text.Trim().ToUpper(), txtQty.Text.Trim(), txtVendorCode.Text.Trim().ToUpper(), txtDateCode.Text.Trim().ToUpper(), txtLotCode.Text.Trim().ToUpper(), Parameter.g_userName, txtMark.Text.Trim().ToUpper(), txtUNID.Text.Trim().ToUpper(), "1", txtUniqueID.Text.Trim(), txtSpec.Text.Trim(), txtMfrSite.Text.Trim());
                    }
                    //10001 end
                }
                lblMsg.Text = "打印成功!";
                lblMsg.ForeColor = Color.Green;

                reFreshData();
                ClearText();
            }
            catch (Exception ex)
            {
                lblMsg.Text = ex.Message + ",请联系QMS!";
            }

        }

        private bool ValidateData()
        {
            if (txtUserID.Text == "")
            {
                lblMsg.Text = "UserID 为空!";
                return false;
            }
            if (txtCompPN.Text == "")
            {
                lblMsg.Text = "料号为空!";
                return false;
            }
            if (txtCompPN.Text.Trim().Length <11)
            {
                lblMsg.Text = "料号长度不能小于11 !";
                return false;
            }
            if (txtVendorCode.Text == "")
            {
                lblMsg.Text = "厂商代码为空!";
                return false;
            }
            if (txtDateCode.Text == "")
            {
                lblMsg.Text = "生产日期为空!";
                return false;
            }
            if (txtLotCode.Text == "")
            {
                lblMsg.Text = "生产批号为空!";
                return false;
            }
            if (txtQty.Text == "" || pubFunction.IsNumeric(txtQty.Text.Trim(), "INT") == false)
            {
                lblMsg.Text = "输入的数量错误，必须为整数!";
                return false;
            }
            if (int.Parse(txtQty.Text.Trim())<0)
            {
                lblMsg.Text = "数量必须大于0!";
                return false;
            }
            return true;
        }

        private void ClearText()
        {
            txtUNID.Text = "";
            txtCompPN.Text = "";
            txtVendorCode.Text = "";
            txtDateCode.Text = "";
            txtLotCode.Text = "";
            txtSpec.Text = "";
            txtMark.Text = "";
            txtQty.Text = "";
            txtUniqueID.Text = "";
            txtMfrSite.Text = "";
            txtUNID.Focus();
            comboBox1.Text = "";
            comboBox1.Items.Clear();
            cboLinkID.Text = "";
            cboLinkID.Items.Clear();
        }

        private void txtKFqty_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && txtKFqty.Text != "")
            {
                if (pubFunction.IsNumeric(txtKFqty.Text, "INT") == false)
                {
                    lblMsg.Text = "请输入正确的数字!";
                    txtKFqty.Text = "";
                    txtKFqty.Focus();
                    return;
                }

                string KFLabelFile = string.Empty;
                string KFMsg = string.Empty;

                KFLabelFile = pubFunction.ConfigListGetValue("KFLabel");
                if (File.Exists(KFLabelFile) == false)
                {
                    lblMsg.Text = "在路径[" + KFLabelFile + "]没找到对应模板!";
                    return;
                }
                StreamReader reader = new StreamReader(KFLabelFile, Encoding.Default);
                string tmpPrintStr = reader.ReadToEnd();
                reader.Close();

                //10001 begin
                DataTable dt = new DataTable();
                if (Parameter.BU == "NB6")
                {
                    dt = MCC.QSMS_SaveCompPrintLog("GetLabelSetting", txtCompPN.Text.Trim().ToUpper(), txtQty.Text.Trim(), txtVendorCode.Text.Trim().ToUpper(), txtDateCode.Text.Trim().ToUpper(),
                     txtLotCode.Text.Trim().ToUpper(), Parameter.g_userName, txtMark.Text.Trim().ToUpper(), txtUNID.Text.Trim().ToUpper(), "1", txtUniqueID.Text.Trim(), txtSpec.Text.Trim(), txtMfrSite.Text.Trim(), "");
                }
                else
                {
                    dt = MCC.QSMS_SaveCompPrintLog("GetLabelSetting", txtCompPN.Text.Trim().ToUpper(), txtQty.Text.Trim(), txtVendorCode.Text.Trim().ToUpper(), txtDateCode.Text.Trim().ToUpper(),
                      txtLotCode.Text.Trim().ToUpper(), Parameter.g_userName, txtMark.Text.Trim().ToUpper(), txtUNID.Text.Trim().ToUpper(), "1", txtUniqueID.Text.Trim(), txtSpec.Text.Trim(), txtMfrSite.Text.Trim());
                }
                //10001 end
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["Result"].ToString() == "1")
                    {
                        lblMsg.Text = dt.Rows[0]["Msg"].ToString();
                        return;
                    }
                    else
                    {
                        Print.LabelSetting(dt.Rows[0]["Setting"].ToString(), dt.Rows[0]["Port"].ToString(), Convert.ToInt32(dt.Rows[0]["Qty"].ToString()), ref KFMsg);
                        for (int i = 0; i < int.Parse(txtKFqty.Text); i++)
                        {
                            if (Print.Print(tmpPrintStr, dt, ref KFMsg) == false)
                            {
                                lblMsg.Text = KFMsg;
                                return;
                            }
                        }
                        lblMsg.Text = "打印成功!";
                        lblMsg.ForeColor = Color.Green;
                    }
                }
                txtKFqty.Text = "";
                txtKFqty.Focus();
            }
        }

        private void txtMark_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnPrint_Click(sender, e);
            }
        }

        private void cboLinkID_KeyDown(object sender, KeyEventArgs e)//10001
        {
            try
            {
                if (e.KeyCode == Keys.Enter && !string.IsNullOrEmpty(cboLinkID.Text.Trim()))
                {
                    string strLinkID = string.Empty;
                    strLinkID = cboLinkID.Text.Trim();

                    DataTable dtLinkID = MCC.CompPrint_GetLinkID(strLinkID);
                    if (dtLinkID.Rows.Count != 0)
                    {
                        cboLinkID.Text = "";
                        cboLinkID.Items.Clear();

                        for (int ik = 0; ik < dtLinkID.Rows.Count; ik++)
                        {
                            cboLinkID.Items.Add(dtLinkID.Rows[ik][0].ToString());
                        }

                        cboLinkID.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

    }
}
