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
    public partial class FrmReturnComp : Form
    {
        public FrmReturnComp()
        {
            InitializeComponent();
        }

        BrLibrary.PublicFunction funtion = new BrLibrary.PublicFunction();
        DbLibrary.MCC.MCCProcess Process = new DbLibrary.MCC.MCCProcess();
        DataTable dt = new DataTable();
        DataTable printdt = new DataTable();
        DataSet ds = new DataSet();
        private string msg;
        private string strPrintPort;//新增flag设置打印属性
        private string strCommSetting;//新增flag设置打印属性
        private string PrintReturnCompLabel;//新增flag设置打印路径及模板名称
        private string strLabelContent;
        private string OldDID;
        private string UnID;

        private void FrmReturnComp_Load(object sender, EventArgs e)
        {
            strPrintPort = funtion.ConfigListGetValue("PrintPort");
            strCommSetting = funtion.ConfigListGetValue("CommSetting");
            PrintReturnCompLabel = Application.StartupPath + "\\" + funtion.ConfigListGetValue("PrintReturnLabel");

            reFreshData();
            OptZebra.Checked = true;
            GetPrinterSetting();
            if (funtion.ConfigListGetValue("BGAWarehouse").ToString().Trim() == "Y")
            {
                ChkBGA.Visible = true;
            }
            this.optBadMaterial.Enabled = false;
            this.optGoodMaterial.Enabled = false;
        }

        private void reFreshData()
        {
            DataTable dt = Process.XL_ReturnCompRefresh(Parameter.Factory);
            gridReturnComp.DataSource = dt.DefaultView;
        }

        public void GetPrinterSetting()
        {
            try
            {
                if (funtion.ConfigListGetValue("Printer").ToString().Trim() == "Zebra")
                {
                    this.OptZebra.Checked = true;
                }
                else
                {
                    this.OptSATO.Checked = true;
                }
                if (funtion.ConfigListGetValue("Port").ToString().Trim() == "COM")
                {
                    this.OptComp.Checked = true;
                }
                else if (funtion.ConfigListGetValue("Port").ToString().Trim() == "LPT")
                {
                    this.OptPrint.Checked = true;
                }
                else
                    this.optNetwork.Checked = true;
                if (funtion.ConfigListGetValue("CommPort").ToString().Trim() != "")
                {
                    this.txtCompPort.Text = funtion.ConfigListGetValue("CommPort").ToString().Trim();
                }
                else
                {
                    this.txtCompPort.Text = "1";
                }
                if (funtion.ConfigListGetValue("Comm").ToString().Trim() != "")
                {
                    this.txtComm.Text = funtion.ConfigListGetValue("Comm").ToString().Trim();
                }
                else
                {
                    this.txtComm.Text = "9600,N,8,1";
                }
                this.OptZebra.Enabled = false;
                this.OptSATO.Enabled = false;
                this.OptComp.Enabled = false;
                this.OptPrint.Enabled = false;
                this.optNetwork.Enabled = false;
                this.txtCompPort.Enabled = false;
                this.txtComm.Enabled = false;
                this.CmdCommSave.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void txtCompPN_KeyPress(object sender, KeyPressEventArgs e)
        {
            string[] NewComp = new string[] { };
            int index;
            string IsBSMaterial;
            if (e.KeyChar == 13 && txtCompPN.Text.Trim() != "")
            {
                if (txtCompPN.Text.Trim().IndexOf(";") > 0)
                {
                    NewComp = txtCompPN.Text.Trim().Split(';');
                    dt = Process.getDID(txtCompPN.Text.ToString());
                    if (dt.Rows.Count > 0)
                    {
                        txtCompPN.Text = dt.Rows[0]["CompPN"].ToString().Trim();
                        txtDateCode.Text = dt.Rows[0]["DateCode"].ToString().Trim();
                        txtVendorCode.Text = dt.Rows[0]["VendorCode"].ToString().Trim();
                        txtLotCode.Text = dt.Rows[0]["LotCode"].ToString().Trim();
                        txtQty.Text = dt.Rows[0]["Qty"].ToString().Trim();
                        UnID = dt.Rows[0]["UNID"].ToString().Trim();
                        if (UnID.Length == 38)
                        {
                            cmdOK_Click(sender, e);
                        }
                    }
                    //for (index = 0; index < NewComp.Length; index++)
                    //{
                    //    if (index == 0)
                    //    {
                    //        txtCompPN.Text = NewComp[index].Trim();
                    //    }
                    //    if (index == 1)
                    //    {
                    //        txtDateCode.Text = NewComp[index].Trim();
                    //    }
                    //    if (index == 2)
                    //    {
                    //        txtVendorCode.Text = NewComp[index].Trim();
                    //    }
                    //    if (index == 3)
                    //    {
                    //        txtLotCode.Text = NewComp[index].Trim();
                    //    }
                    //    if (index == 4)
                    //    {
                    //        txtQty.Text = NewComp[index].Trim();
                    //    }
                    //}
                    if (funtion.ConfigListGetValue("CheckBSMaterial") == "Y")
                    {
                        IsBSMaterial = "N";
                        if (txtLotCode.Text.Trim() == "@@@@")
                        {
                            IsBSMaterial = "Y";
                            txtQty.Text = txtQty.Text.Trim().Replace("PCS", "");
                        }
                        else
                        {
                            dt = Process.GetComponent_Data(txtCompPN.Text.Trim());
                            if (dt.Rows.Count > 0)
                            {
                                IsBSMaterial = "Y";
                            }
                        }
                        if (IsBSMaterial == "Y")
                        {
                            txtDateCode.Text = "";
                            txtLotCode.Text = "";
                            txtDateCode.Text = "";
                            //txtDateCode.Text = UCase(Trim(InputBox("请刷入Datecode:", "Input Datecode")))
                            //txtLotCode.Text = UCase(Trim(InputBox("请刷入Lotcode:", "Input Lotcode")))
                            if (txtDateCode.Text == "" || txtLotCode.Text == "")
                            {
                                MessageBox.Show("DateCode Or LotCode is empty!");
                                return;
                            }
                        }
                    }
                    txtQty.Focus();
                }
                else if (txtCompPN.Text.Trim().IndexOf("-") > 0 && txtCompPN.Text.Trim().Length > 15)
                {
                    OldDID = txtCompPN.Text.Trim();
                    dt=Process.GetQSMS_DID_ToWH(txtCompPN.Text.Trim());
                    if (dt.Rows.Count > 0)
                    {
                        txtCompPN.Text = dt.Rows[0]["compPN"].ToString().Trim();
                        txtVendorCode.Text = dt.Rows[0]["VendorCode"].ToString().Trim();
                        txtDateCode.Text = dt.Rows[0]["DateCode"].ToString().Trim();
                        txtLotCode.Text = dt.Rows[0]["LotCode"].ToString().Trim();
                        txtQty.Text = dt.Rows[0]["Qty"].ToString().Trim();
                        cmdOK_Click(sender, e);
                    }
                    else
                    {
                        MessageBox.Show("Can't find the information of this returnDID---" + txtCompPN.Text.Trim());
                        txtCompPN.Focus();
                        return;
                    }
                }
                if (txtCompPN.Text == "")
                {
                    txtVendorCode.Focus();
                }
            }
        }

        private void txtDateCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && txtDateCode.Text != "")
            {
                txtLotCode.Focus();
            }
        }

        private void txtQty_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && txtQty.Text != "")
            {
                if (txtQty.Text.Trim() != "")//是数字
                {
                    if (funtion.IsNumeric(txtQty.Text.Trim(), "INT") == true)
                    {
                        txtQty.Text = txtQty.Text;
                        cmdOK_Click(null, null);
                    }
                }
                else
                {
                    txtQty.Focus();
                }
            }
        }

        private void txtVendorCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && txtVendorCode.Text != "")
            {
                txtDateCode.Focus();
            }
        }

        private void txtLotCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && txtLotCode.Text != "")
            {
                txtQty.Focus();
            }
        }

        private void cmdOK_Click(object sender, EventArgs e)
        {
            try
            {
                string sCompPN, intReturnQty = "";
                cmdOK.Enabled = false;
                LblMessage.Text = "";
                lblFeedBack.Text = "Qty FeedBack:";
                if (ChkValidData() == false)
                {
                    goto Normal_Eixt;
                }
                intReturnQty = txtQty.Text.Trim();
                sCompPN = txtCompPN.Text.Trim();
                dt = Process.QSMS_MCC_QueryDataByType("MCC_Getmsd_data1", "", "", txtCompPN.Text.Trim(),"","");
                if (ChkBGA.Checked == true)
                {
                    ds = Process.QSMS_MCC_XL_ReturnComp(sCompPN, txtVendorCode.Text.Trim(), txtDateCode.Text.Trim(), txtLotCode.Text.Trim(), (optGoodMaterial.Checked == true) ? "YBGA" : "NBGA", int.Parse(intReturnQty.Trim()), Parameter.g_userName, Parameter.Factory, funtion.ConfigListGetValue("CheckReturnForbiddenPN"), OldDID);
                }
                else if (ChkHUA.Checked == true)
                {
                    ds = Process.QSMS_MCC_XL_ReturnComp(sCompPN, txtVendorCode.Text.Trim(), txtDateCode.Text.Trim(), txtLotCode.Text.Trim(), (optGoodMaterial.Checked == true) ? "YHUA" : "NHUA", int.Parse(intReturnQty.Trim()), Parameter.g_userName, Parameter.Factory, funtion.ConfigListGetValue("CheckReturnForbiddenPN"), OldDID);
                }
                else
                {
                    ds = Process.QSMS_MCC_XL_ReturnComp(sCompPN, txtVendorCode.Text.Trim(), txtDateCode.Text.Trim(), txtLotCode.Text.Trim(), (optGoodMaterial.Checked == true) ? "Y" : "N", int.Parse(intReturnQty.Trim()), Parameter.g_userName, Parameter.Factory, funtion.ConfigListGetValue("CheckReturnForbiddenPN"),OldDID);
                }
                dt = ds.Tables[0];
                if (dt.Rows[0]["Result"].ToString() != "0")
                {
                    LblMessage.Text = dt.Rows[0]["Description"].ToString();
                }
                else
                {
                    LblMessage.BackColor = Color.Green;
                    dt = ds.Tables[1];
                    if (dt.Rows.Count <= 0)
                    {
                        LblMessage.Text = "Get DID information fail,print DID fail!!";
                        goto Normal_Eixt;
                    }
                    printdt = dt;
                    lblFeedBack.Text = dt.Rows[0]["QtyFeedback"].ToString();
                    lblFeedBack.Text = lblFeedBack.Text.Trim().Substring(lblFeedBack.Text.Trim().IndexOf("##"), lblFeedBack.Text.Trim().Length - lblFeedBack.Text.Trim().IndexOf("##") - 1);
                    Parameter.DIDInfo.DID = dt.Rows[0]["DID"].ToString().Trim();
                    Parameter.DIDInfo.compPN = dt.Rows[0]["CompPN"].ToString().Trim();
                    Parameter.DIDInfo.Qty = Convert.ToInt32(dt.Rows[0]["Qty"].ToString().Trim());
                    Parameter.DIDInfo.IsGood = dt.Rows[0]["IsGood"].ToString().Trim();
                    Parameter.DIDInfo.VendorCode = dt.Rows[0]["VendorCode"].ToString().Trim();
                    Parameter.DIDInfo.DateCode = dt.Rows[0]["DateCode"].ToString().Trim();
                    Parameter.DIDInfo.LotCode = dt.Rows[0]["LotCode"].ToString().Trim();
                    if (Parameter.BU == "NB5" || Parameter.BU == "PU5")
                    {
                        Parameter.DIDInfo.WareHouseID = dt.Rows[0]["WareHouseID"].ToString().Trim();
                    }
                    if (Parameter.BU == "NB3" || Parameter.BU == "PU3")
                    {
                        Parameter.DIDInfo.VendorCode = gridReturnComp.Columns[4].ToString();
                        Parameter.DIDInfo.DateCode = gridReturnComp.Columns[5].ToString();
                        Parameter.DIDInfo.LotCode = gridReturnComp.Columns[6].ToString();
                    }
                    if (funtion.ConfigListGetValue("ChkPrintDIDType") == "Y")
                    {
                        Parameter.DIDInfo.DIDType = dt.Rows[0]["DIDType"].ToString().Trim();
                    }
                    else
                    {
                        Parameter.DIDInfo.DIDType = "";
                    }
                    DIDPrintLabel(OptZebra.Checked, int.Parse(txtCompPort.Text.Trim()), txtComm.Text.Trim());
                    reFreshData();
                }
                txtCompPN.Focus();
            }
            catch (Exception EX)
            {
                LblMessage.Text = EX.Message;
                cmdOK.Enabled = true;
            }
        Normal_Eixt:
            {
                txtVendorCode.Text = "";
                txtLotCode.Text = "";
                txtDateCode.Text = "";
                txtQty.Text = "";
                txtCompPN.Text = "";
                txtCompPN.Focus();
                cmdOK.Enabled = true;
            }
        }

        private void cmdReprint_Click(object sender, EventArgs e)
        {
            if (txtCompPort.Text == "" || txtComm.Text == "")
            {
                MessageBox.Show("Printer have not set!!");
                txtCompPN.Focus();
                return;
            }
            if (gridReturnComp.Rows.Count < 0)
            {
                return;
            }
            if (gridReturnComp.Columns[0].ToString() != "")
            {
                Parameter.DIDInfo.DID = gridReturnComp.Columns[0].ToString().Trim();
                Parameter.DIDInfo.compPN = gridReturnComp.Columns[1].ToString().Trim();
                Parameter.DIDInfo.Qty = long.Parse(gridReturnComp.Columns[2].ToString().Trim());
                Parameter.DIDInfo.IsGood = gridReturnComp.Columns[3].ToString().Trim();
                if (Parameter.BU == "NB5")
                {
                    Parameter.DIDInfo.WareHouseID = gridReturnComp.Columns[20].ToString().Trim();
                }
                if (funtion.ConfigListGetValue("ChkPrintDIDType") == "Y")
                {
                    Parameter.DIDInfo.DIDType = gridReturnComp.Columns[8].ToString().Trim();
                }
                else
                {
                    Parameter.DIDInfo.DIDType = "";
                }
                if (Parameter.BU == "NB3" || Parameter.BU == "NB5")
                {
                    Parameter.DIDInfo.VendorCode = gridReturnComp.Columns[4].ToString().Trim();
                    Parameter.DIDInfo.DateCode = gridReturnComp.Columns[5].ToString().Trim();
                    Parameter.DIDInfo.LotCode = gridReturnComp.Columns[6].ToString().Trim();
                }
                DIDPrintLabel(OptZebra.Checked, int.Parse(txtCompPort.Text.Trim()), txtComm.Text.Trim());
            }
        }

        private void cmdGetRefID_Click(object sender, EventArgs e)
        {
            string sCurrRefID, sMsg;
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            dt = Process.XL_DIDGetRefID(((optGoodMaterial.Checked == true) ? "Y" : "N"), Parameter.g_userName, Parameter.Factory);
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["Result"].ToString() != "0")
                {
                    MessageBox.Show(dt.Rows[0]["Description"].ToString(), "Prompt");
                    txtCompPN.Focus();
                    return;
                }
                sMsg = dt.Rows[0]["Description"].ToString().Trim();
                sCurrRefID = funtion.DIDGetRefIDByResult(sMsg);
                Parameter.DIDInfo.DID = sCurrRefID;
                Parameter.DIDInfo.compPN = sCurrRefID;
                Parameter.DIDInfo.Qty = -10000;
                Parameter.DIDInfo.IsGood = (optGoodMaterial.Checked == true) ? "Y" : "N";
                Parameter.DIDInfo.DIDType = "";
                DIDPrintLabel(OptZebra.Checked, int.Parse(txtCompPort.Text.Trim()), txtComm.Text.Trim());
                ds = Process.XL_DIDChkStockByRefID_set(sCurrRefID, Parameter.g_userName);
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["Result"].ToString() != "0")
                    {
                        MessageBox.Show(dt.Rows[0]["Description"].ToString(), "Prompt");
                        txtCompPN.Focus();
                        return;
                    }
                    dt = ds.Tables[1];
                    frmDIDCheckStock CheckDIDStock = new frmDIDCheckStock();
                    frmDIDCheckStock.FuncType = "AutoChk";
                    CheckDIDStock.Show();
                }
            }
        }

        private void FrmReturnComp_FormClosed(object sender, FormClosedEventArgs e)
        {
            funtion.RemoveForm("FrmReturnComp");
        }

        private void DIDPrintLabel(Boolean blnZebra, int intCompPort, string sCommString)
        {
            string BU = "";
            PrinterLib.PrintLabel lblprint = new PrinterLib.PrintLabel();
            if (lblprint.LabelSetting(strCommSetting, strPrintPort, 1, ref msg) == false)
            {
                LblMessage.Text = msg;
                return;
            }
            if (File.Exists(PrintReturnCompLabel) == false)
            {
                MessageBox.Show("File:" + PrintReturnCompLabel + " not exists");
                txtCompPN.Focus();
                return;
            }
            if (string.IsNullOrEmpty(strLabelContent))
            {
                strLabelContent = new StreamReader(PrintReturnCompLabel).ReadToEnd();
            }
            if (lblprint.PrintReturn(strLabelContent, printdt, BU, ref msg) == false)
            {
                LblMessage.Text = msg;
                txtCompPN.Focus();
                return;
            }
        }

        private Boolean ChkValidData()
        {
            if (txtVendorCode.Text == "")
            {
                LblMessage.Text = "Vendor Code is blank!!";
                return false;
            }
            if (txtLotCode.Text == "")
            {
                LblMessage.Text = "Lot Code is blank!!";
                return false;
            }
            if (txtDateCode.Text == "")
            {
                LblMessage.Text = "Date Code is blank!!";
                return false;
            }
            if (txtCompPN.Text.Trim() == "")
            {
                LblMessage.Text = "CompPN is blank!!";
                return false;
            }
            if (txtQty.Text == "")
            {
                LblMessage.Text = "Qty is blank!!";
                return false;
            }
            if (int.Parse(txtQty.Text.Trim()) < 0)
            {
                txtQty.Text = ((-1) * int.Parse(txtQty.Text.Trim())).ToString();
            }
            if (int.Parse(txtQty.Text.Trim()) <= 0)
            {
                LblMessage.Text = "The Return Qty must be >0 !!";
                return false;
            }
            txtVendorCode.Text = txtVendorCode.Text.ToUpper().Replace("'", "");
            txtLotCode.Text = txtLotCode.Text.ToUpper().Replace("'", "");
            txtDateCode.Text = txtDateCode.Text.ToUpper().Replace("'", "");
            txtCompPN.Text = txtCompPN.Text.ToUpper().Replace("'", "");
            return true;
        }

        private void CmdCommSave_Click(object sender, EventArgs e)
        {

        }

        private void txtCompPN_TextChanged(object sender, EventArgs e)
        {
            if (Parameter.BU == "NB5" || Parameter.BU == "PU5")
            {
                if (txtCompPN.Text.Trim() != "")
                {
                    if (txtCompPN.Text.Trim().Substring(0, 1).ToUpper().ToString() == "A" || txtCompPN.Text.Trim().Substring(0, 1).ToUpper().ToString() == "B")
                    {
                        if (Process.CheckReturnRight() == false)
                        {
                            MessageBox.Show("当前用户没有Return料号'A'或'B'开头的权限！！");
                            txtCompPN.Text = "";
                            txtCompPN.Focus();
                            return;
                        }
                    }
                }
            }
        }
    }
}
