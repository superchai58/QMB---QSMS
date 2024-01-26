using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Threading;
using System.IO.Ports;

namespace QSMS.QSMS.IPQC
{
    public partial class frmInSpection : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.IPQC.IPQCProcess IPQC = new DbLibrary.IPQC.IPQCProcess();
        private SerialPort serialPort = new SerialPort();

        private string ReceiveData = string.Empty;
        private string strDID = string.Empty;
        private string DIDQty = string.Empty;
        private string VendorPN = string.Empty;
        private string strLotNo = string.Empty;
        private string Result = string.Empty;
        private string PreCompPN = string.Empty;
        private string IPQCFlag = string.Empty;
        private string ImagePath = string.Empty;
        private string strScanCompPN = string.Empty;
        private string PreStr = string.Empty;

        public frmInSpection()
        {
            InitializeComponent();
        }

        private void frmInSpection_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.serialPort.Close();
            pubFunction.RemoveForm("frmInSpection");
        }

        private void frmInSpection_Load(object sender, EventArgs e)
        {
            strScanCompPN= pubFunction.ConfigListGetValue("ScanCompPN");
            //Image1.Load("D:\\3H97MB06X0.jpg");
            txtDID.Text = "";
            txtDID.Focus();
            lblDelaytime.Visible = false;
            txtDelaytime.Visible = false;
            txtDelaytime.Text = "100";

            txtUid.Text = Parameter.g_userName;
            txtBU.Text = Parameter.BU;

            cboEquipType.Items.Add("34401");
            cboEquipType.Items.Add("3302");
            cboEquipType.Items.Add("4235");
            cboEquipType.Items.Add("4300");
            cboEquipType.Items.Add("6420");
            cboEquipType.Items.Add("8110G");
            cboEquipType.Items.Add("3523");
            cboEquipType.Items.Add("TH2832");  //Leslie  PU6 Add TH2832  0001
            if (Parameter.BU == "NB5")
            {
                cboEquipType.Text = "4300";
            }
            else if (Parameter.BU == "NB3")
            {
                cboEquipType.Text = "8110G";
            }

            cboCom.Items.Add("COM1");
            cboCom.Items.Add("COM2");
            cboCom.Items.Add("COM3");
            cboCom.Items.Add("COM4");
            cboCom.Items.Add("COM5");
            cboCom.Items.Add("COM6");
            cboCom.Text = "COM1";
            //cboCom.Items.Add("GPIB22");

            if (strScanCompPN == "Y")
            {
                lblLotNo.Visible = true;
                txtLotNo.Visible = true;
            }
            txtimagePath.Text = pubFunction.ConfigListGetValue("ImagePath").ToUpper();
            btnStart.Enabled = false;
        }

        private void timer_Time_Tick(object sender, EventArgs e)
        {
            txtDatetime.Text = DateTime.Now.ToString();
            //btnStart.Enabled = false;
        }

        private void txtDID_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter && txtDID.Text != "")
            if (e.KeyCode == Keys.Enter && !string.IsNullOrEmpty(txtDID.Text.Trim()))
            {
                //txtDID.Text = txtDID.Text.Replace("\r", "");
                //txtDID.Text = txtDID.Text.Replace("\n", "");
                //txtDID.Text = txtDID.Text.Trim().ToUpper();
                txtDID.Text = txtDID.Text.Replace("\r", "").Replace("\n", "").Trim().ToUpper();
                C_Interface();
                DG_Result.DataSource = null;

                //get DID basc info
                if (strScanCompPN != "Y")
                {
                    #region QSMC 不存在这段代码
                    //if (txtDID.Text.IndexOf(";") > 0)
                    //{
                    //    DataTable dt = IPQC.QSMS_GenUNID(txtDID.Text, "");
                    //    if (dt.Rows.Count > 0)
                    //    {
                    //        if (dt.Rows[0]["Result"].ToString().ToUpper() == "OK")
                    //        {
                    //            //txtCompPN.Text = dt.Rows[0]["CompPN"].ToString().ToUpper();
                    //            //txtVendor.Text = dt.Rows[0]["VendorCode"].ToString().ToUpper();
                    //            //txtDateCode.Text = dt.Rows[0]["DateCode"].ToString().ToUpper();
                    //            //txtLotCode.Text = dt.Rows[0]["LotCode"].ToString().ToUpper();
                    //            txtDID.Text = dt.Rows[0]["UNID"].ToString().ToUpper();
                    //            //txtDIDQty.Text = dt.Rows[0]["Qty"].ToString().ToUpper();
                    //        }
                    //        else
                    //        {
                    //            lblMsg.Text = dt.Rows[0]["Msg"].ToString().ToUpper();
                    //            lblMsg.ForeColor = Color.Red;
                    //            txtDID.Text = "";
                    //            txtDID.Focus();
                    //        }
                    #endregion

                    strDID = txtDID.Text;
                    if (GetBaseInfoByDID(strDID) == false)
                    {
                        return;
                    }
                }
                else
                {
                    txtCompPN.Text = txtDID.Text;
                    strLotNo = GetLotNo();
                    strDID = strLotNo;
                    txtLotNo.Text = strLotNo;
                }
                if (txtDID.Text.Trim().Length < 12)
                {
                    txtCompPN.Text = txtDID.Text.Trim();
                }
                else
                {
                    txtCompPN.Text = txtDID.Text.Trim().Substring(0, 11);
                }

                //get rule of comppn
                GetRuleByCompPN(txtCompPN.Text);
                ShowPicture(txtCompPN.Text, txtVendor.Text);
                if (txtunit.Text == "R" || txtunit.Text == "C" || txtunit.Text == "L" || txtunit.Text == "Z")
                {
                    if (MessageBox.Show(@"请核对DID\CompPN信息是否正确？", "提示", MessageBoxButtons.YesNo) != DialogResult.Yes)
                    {
                        txtDID.SelectAll();
                        return;
                    }
                }
                btnStart.Enabled = true;
                btnStart.Focus();
                lblMsg.Text = "准备测试!";
            }
        }

        private void C_Interface()
        {
            txtCompPN.Text = "";
            txtVendor.Text = "";
            txtLotCode.Text = "";
            txtDateCode.Text = "";
            txtDIDQty.Text = "";
            txtChkNum.Text = "";
            txtDIDUID.Text = "";
            txtSpec.Text = "";
            txtunit.Text = "";
            VendorPN = "";
        }

        private bool GetBaseInfoByDID(string DID)
        {
            DataTable dt = IPQC.QSMS_PD_QueryDataByType("IPQC_GetBaseInfoByDID", "", "", "", DID, "");
            if (dt.Rows.Count > 0)
            {
                txtCompPN.Text = dt.Rows[0]["CompPN"].ToString().Trim().ToUpper();
                txtVendor.Text = dt.Rows[0]["VendorCode"].ToString().Trim().ToUpper();
                txtLotCode.Text = dt.Rows[0]["LotCode"].ToString().Trim().ToUpper();
                txtDateCode.Text = dt.Rows[0]["DateCode"].ToString().Trim().ToUpper();
                txtDIDQty.Text = dt.Rows[0]["Qty"].ToString().Trim().ToUpper();
                DIDQty = dt.Rows[0]["Qty"].ToString().Trim().ToUpper();
                txtDIDUID.Text = dt.Rows[0]["UID"].ToString().Trim().ToUpper();
            }
            else
            {
                lblMsg.Text = "该DID[" + DID + "]不存在!";
                lblMsg.ForeColor = Color.Red;
                txtDID.Enabled = true;
                btnStart.Enabled = false;
                txtDID.Text = "";
                txtDID.Focus();
                return false;
            }

            //if (Parameter.IPQC_ChkVendorPN == "Y")
            //{
            //    dt = IPQC.QSMS_PD_QueryDataByType("IPQC_GetVendorPN", "", "", txtCompPN.Text, txtVendor.Text, "");
            //    if (dt.Rows.Count > 0)
            //    {
            //        VendorPN = dt.Rows[0]["VendorPN"].ToString().Trim().ToUpper();
            //    }
            //    else
            //    {
            //        lblMsg.Text = "没有找到 Vendor PN !";
            //        lblMsg.ForeColor = Color.Red;
            //        return false;
            //    }
            //}
            return true;
        }

        private string GetLotNo()
        {
            DataTable dt = IPQC.GetLotNo();
            if (dt.Rows.Count > 0)
            {
                return dt.Rows[0]["LotNo"].ToString().Trim().ToUpper();
            }
            return "";
        }

        private void GetRuleByCompPN(string strPN)
        {
            try
            {
                string Upper = string.Empty;
                string Lower = string.Empty;
                DataTable dt = IPQC.QSMS_PD_QueryDataByType("IPQC_GetRuleByCompPN", "", "", strPN, "", "");
                if (dt.Rows.Count > 0)
                {
                    Upper = Convertdouble(dt.Rows[0]["Upper"].ToString()).ToString();
                    Lower = Convertdouble(dt.Rows[0]["Lower"].ToString()).ToString();
                    txtFrequency.Text = dt.Rows[0]["Hz"].ToString().Trim().ToUpper();
                    txtVoltage.Text = dt.Rows[0]["Volt"].ToString().Trim().ToUpper();
                    txtcurrent.Text = dt.Rows[0]["Ampere"].ToString().Trim().ToUpper();
                    txtunit.Text = dt.Rows[0]["Unit"].ToString().Trim().ToUpper();
                    txtChkNum.Text= dt.Rows[0]["CHKNum"].ToString().Trim().ToUpper();

                    if (txtunit.Text == "IC" || txtunit.Text == "CON")
                    {
                        txtChkNum.Text = "1";
                    }
                    if (txtunit.Text == "CON")
                    {
                        txtSpec.Text = Lower;
                    }
                    else
                    {
                        txtSpec.Text = Lower + "-" + Upper;
                    }

                    if (strScanCompPN != "Y" && DIDQty != "")
                    {
                        if (long.Parse(DIDQty) < long.Parse(dt.Rows[0]["BaseQty"].ToString().Trim()))
                        {
                            txtChkNum.Text = (int.Parse(dt.Rows[0]["CHKNum"].ToString().Trim()) - 1).ToString();
                        }
                    }
                    else
                    {
                        txtChkNum.Text = dt.Rows[0]["CHKNum"].ToString().Trim().ToUpper();
                    }
                    if (txtChkNum.Text == "0")
                    {
                        txtChkNum.Text = "1";
                    }

                }
                else
                {
                    lblMsg.Text = "该CompPN[" + strPN + "]测试标准不存在!";
                    lblMsg.ForeColor = Color.Red;
                    txtDID.Enabled = true;
                    btnStart.Enabled = false;
                    return;
                }
            }
            catch (Exception ex)
            {
                lblMsg.Text = ex.Message;
                lblMsg.ForeColor = Color.Red;
                txtDID.Enabled = true;
                btnStart.Enabled = false;
                return;
            }
        }

        public Decimal Convertdouble(string txt)
        {
            Decimal dData = 0.0M;
            if (txt.Contains("E"))
            {
                dData = Decimal.Parse(txt, System.Globalization.NumberStyles.Float);
            }
            else
            {
                dData = Decimal.Parse(txt);
            }
            return dData;
        } 

        private void ShowPicture(string compPN, string Vendor)
        {
            string PicFilename = string.Empty;

            if (pubFunction.ConfigListGetValue("ImagePath") == "")
            {
                ImagePath = Application.StartupPath.Trim();
            }
            if (compPN != "" && Vendor != "")
            {
                PicFilename = compPN + "-" + Vendor + ".jpg";
                if (File.Exists(ImagePath + "\\" + PicFilename) == false)
                {
                    //if (Parameter.IPQC_ChkVendorPN == "Y")
                    //{
                    //    MessageBox.Show("没有找到文件或路径["+ Parameter.imagePath + "\\" + PicFilename + "]!");
                    //    return;
                    //}
                    //else
                    //{
                    return;
                    //}
                }
                Image1.Load(ImagePath + "\\" + PicFilename);
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            string TestStr = string.Empty;
            string Errcode = string.Empty;
            string ErrDesc = string.Empty;
            string Unit = string.Empty;
            string sGetStr = string.Empty;
            Double TestValue = new double();
            try
            {
                if (txtDID.Text == "")
                {
                    MessageBox.Show("DID为空！");
                    return;
                }
                if (txtFrequency.Text == "" || txtVoltage.Text == "")
                {
                    MessageBox.Show("请刷入正确DID！");
                    return;
                }
                if (strScanCompPN == "Y")
                {
                    txtDID.Enabled = false;
                }
                btnStart.Enabled = false;
                if (cboEquipType.Text == "")
                {
                    lblMsg.Text = "没有选择设备类型!";
                    lblMsg.ForeColor = Color.Red;
                    pubFunction.Sound("ERROR");
                    btnStart.Enabled = true;
                    cboEquipType.Focus();
                    return;
                }
                if (strScanCompPN == "Y")
                {
                    lblMsg.Text = txtDID.Text + "允许测试";
                }
                else
                {
                    DataTable dt = IPQC.QSMS_PD_QueryDataByType("IPQC_GetIPQCFlag", "", "", txtDID.Text, "", "");
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["IPQCFlag"].ToString().ToUpper() == "Y")
                        {
                            if (chkIPQCRetest(txtUid.Text, txtDID.Text) == false)
                            {
                                lblMsg.Text = "已测试PASS，没有权限(IPQCRetest)重测!";
                                lblMsg.ForeColor = Color.Red;
                                txtDID.Text = "";
                                txtDID.Enabled = true;
                                txtDID.Focus();
                                return;
                            }
                        }
                        else if (dt.Rows[0]["IPQCFlag"].ToString().ToUpper() == "N")
                        {
                            if (chkIPQCRetest(txtUid.Text, txtDID.Text) == false)
                            {
                                lblMsg.Text = "已测试FAIL，没有权限(IPQCRetest)重测!";
                                lblMsg.ForeColor = Color.Red;
                                txtDID.Text = "";
                                txtDID.Enabled = true;
                                txtDID.Focus();
                                return;
                            }
                        }
                        else
                        {
                            lblMsg.Text = txtDID.Text + "允许测试";
                        }
                    }
                }

                if (txtunit.Text.ToUpper() == "IC")
                {
                    if (MessageBox.Show("请核对IC信息是否正确?", "IC 测试", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        TestStr = "0";  //0-->pass,1-->fail
                        Errcode = "PASS";
                    }
                    else
                    {
                        TestStr = "1";
                        Errcode = Interaction.InputBox("请输入错误代码：", "IC 测试", "");
                        DataTable dt = IPQC.QSMS_PD_QueryDataByType("IPQC_ChkErrCode", "", "", Errcode, "", "");
                        if (dt.Rows.Count == 0)
                        {
                            lblMsg.Text = "请输入正确的错误代码!";
                            lblMsg.ForeColor = Color.Red;
                            pubFunction.Sound("ERROR");
                            btnStart.Enabled = true;
                            return;
                        }
                    }
                }
                else if (txtunit.Text.ToUpper() == "CON")
                {
                    if (MessageBox.Show("Connecter PIN:" + txtSpec.Text + "  请核对实物是否相符?", "CON 测试", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        TestStr = "0";  //0-->pass,1-->fail
                        Errcode = "PASS";
                    }
                    else
                    {
                        TestStr = "1";
                        Errcode = Interaction.InputBox("请输入错误代码：", "CON 测试", "");
                        DataTable dt = IPQC.QSMS_PD_QueryDataByType("IPQC_ChkErrCode", "", "", Errcode, "", "");
                        if (dt.Rows.Count == 0)
                        {
                            lblMsg.Text = "请输入正确的错误代码!";
                            lblMsg.ForeColor = Color.Red;
                            pubFunction.Sound("ERROR");
                            btnStart.Enabled = true;
                            return;
                        }
                        else
                        {
                            ErrDesc = dt.Rows[0]["CHErrDesc"].ToString();
                        }
                    }
                }
                else
                {
                    Thread.Sleep(Convert.ToInt32(txtDelaytime.Text));
                    Unit = txtunit.Text.ToUpper();
                    if (Unit == "TRIODE")
                    {
                        Unit = "DIODE";
                    }
                    if (cboCom.Text != "")
                    {
                        if (cboCom.Text.ToUpper() == "GPIB22")
                        {
                            //仪器端口22, 'D'测试用参数 ,'R'电阻
                            //sGetStr = MeasureE_GPIB(cboEquipType.Text, 22, Unit, Convert.ToDouble(txtFrequency.Text), Convert.ToDouble(txtVoltage.Text), Convert.ToDouble(txtcurrent.Text)); 
                            sGetStr = "0";
                        }
                        else
                        {
                            if (cboEquipType.Text == "4300" || cboEquipType.Text == "4235")
                            {
                                TestStr = MeasureE4300_RS232(cboEquipType.Text, Unit, Convert.ToDouble(txtFrequency.Text), Convert.ToDouble(txtVoltage.Text));
                            }
                            else if (cboEquipType.Text == "6420")
                            {
                                TestStr = MeasureE6420_RS232(cboEquipType.Text, Unit, Convert.ToDouble(txtFrequency.Text), Convert.ToDouble(txtVoltage.Text));
                                if (Unit.ToUpper() == "C")
                                {
                                    TestStr = (Convert.ToDouble(TestStr) * 1000000000000).ToString();
                                }
                            }
                            else if (cboEquipType.Text == "3523")
                            {
                                TestStr = Measure3523_RS232(cboEquipType.Text, Unit, Convert.ToDouble(txtFrequency.Text), Convert.ToDouble(txtVoltage.Text));
                            }
                            else if (cboEquipType.Text == "8110G")
                            {
                                TestStr = MeasureE8110G_RS232(cboEquipType.Text, Unit, Convert.ToDouble(txtFrequency.Text), Convert.ToDouble(txtVoltage.Text));
                            }
                            else if (cboEquipType.Text == "TH2832") //Leslie  add TH3832  0001
                            {
                                TestStr=Measure_TH2832(cboEquipType.Text, Unit, Convert.ToDouble(txtFrequency.Text), Convert.ToDouble(txtVoltage.Text));                        
                            }
                            else
                            {
                                TestStr = "0";
                            }
                        }
                    }
                    else
                    {
                        lblMsg.Text = "请选择仪器端口!";
                        lblMsg.ForeColor = Color.Red;
                        pubFunction.Sound("ERROR");
                        btnStart.Enabled = true;
                        return;
                    }
                    if (cboEquipType.Text != "4300" && cboEquipType.Text != "6420" && cboEquipType.Text != "3523" && cboEquipType.Text != "8110G" && cboEquipType.Text != "4235")
                    {
                        //
                    }
                }
                if (TestStr == "")
                {
                    TestStr = "0";
                }
                TestValue = Convert.ToDouble(TestStr);
                DataSet ds = IPQC.QSMSDIDInSpect(strDID, txtCompPN.Text, txtVendor.Text, TestValue.ToString(), Errcode, strScanCompPN);
                DataTable dtResult = ds.Tables[0];
                DataTable dtDG_Result = ds.Tables[1];
                DG_Result.DataSource = null;
                DG_Result.DataSource = dtDG_Result;

                if (dtResult.Rows.Count > 0)
                {
                    IPQCFlag = dtResult.Rows[0]["result"].ToString().Trim().ToUpper();
                }

                if (dtDG_Result.Rows.Count > 0)
                {
                    int index = dtDG_Result.Rows.Count - 1;
                    if (dtDG_Result.Rows[index]["testorder"].ToString().ToUpper() != txtChkNum.Text.ToUpper())
                    {
                        lblMsg.Text = "该DID第" + dtDG_Result.Rows[index]["testorder"].ToString().ToUpper() + " 颗测试结果为:"
                            + dtDG_Result.Rows[index]["testresult"].ToString().ToUpper() + " " + ErrDesc;
                        if (dtDG_Result.Rows[index]["testresult"].ToString().ToUpper() == "PASS")
                        {
                            lblMsg.ForeColor = Color.Green;
                            pubFunction.Sound("OK");
                        }
                        else
                        {
                            lblMsg.ForeColor = Color.Red;
                            pubFunction.Sound("ERROR");
                        }
                        if (strScanCompPN == "Y")
                        {
                            txtDID.Enabled = true;
                        }
                        btnStart.Enabled = true;
                        btnStart.Focus();
                        return;
                    }
                    else
                    {
                        if (IPQCFlag == "PASS")
                        {
                            DataTable dt = IPQC.QSMS_PD_QueryDataByType("IPQC_UpdIPQCFlag", "", "", txtDID.Text, "Y", "");
                            if (dt.Rows.Count > 0)
                            {
                                if (dt.Rows[0]["Result"].ToString() == "0")
                                {
                                    lblMsg.ForeColor = Color.Green;
                                    pubFunction.Sound("OK");
                                }
                                else
                                {
                                    MessageBox.Show(dt.Rows[0]["Msg"].ToString());
                                    lblMsg.ForeColor = Color.Red;
                                    pubFunction.Sound("ERROR");
                                }
                            }
                        }
                        else
                        {
                            DataTable dt = IPQC.QSMS_PD_QueryDataByType("IPQC_UpdIPQCFlag", "", "", txtDID.Text, "N", "");
                            if (dt.Rows.Count > 0)
                            {
                                if (dt.Rows[0]["Result"].ToString() == "0")
                                {
                                    lblMsg.ForeColor = Color.Red;
                                    pubFunction.Sound("ERROR");
                                }
                                else
                                {
                                    MessageBox.Show(dt.Rows[0]["Msg"].ToString());
                                    lblMsg.ForeColor = Color.Red;
                                    pubFunction.Sound("ERROR");
                                }
                            }
                        }
                        if (DG_Result.Rows[0].Cells[3].Value.ToString() == "CON" || DG_Result.Rows[0].Cells[3].Value.ToString() == "IC")
                        {
                            lblMsg.Text = txtDID.Text + "目检结果为:" + IPQCFlag + "目检完成!";
                            if (IPQCFlag.ToUpper() == "PASS")
                            {
                                pubFunction.Sound("OK");
                            }
                            else
                            {
                                pubFunction.Sound("ERROR");
                            }
                            
                        }
                        else
                        {
                            lblMsg.Text = txtDID.Text + "测试结果为:" + IPQCFlag + "测试完成!";
                            if (IPQCFlag.ToUpper() == "PASS")
                            {
                                pubFunction.Sound("OK");
                            }
                            else
                            {
                                pubFunction.Sound("ERROR");
                            }
                        }
                        txtDID.Enabled = true;
                        txtDID.Text = "";
                        btnStart.Enabled = true;
                        txtDID.Focus();
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + '[' + TestStr + ']', "Start_Click");
                txtDID.Enabled = true;
                txtDID.Text = "";
                txtDID.Focus();
                return;
            }
        }

        private bool chkIPQCRetest(string UID, string DID)
        {
            DataTable dt = IPQC.QSMS_PD_QueryDataByType("IPQC_ChkRight", "", "", UID, "IPQCRetest", "");
            if (dt.Rows.Count > 0)
            {
                if (MessageBox.Show("确定要重测该DID[" + DID + "]吗？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }

        }

        private string MeasureE4300_RS232(string sEquipType, string sUnit, double dblFrequency, double dblVoltage)
        {
            try
            {
                if (!serialPort.IsOpen)
                {
                    InitPort(5);
                }
                ReceiveData = "";
                string[] bytes = new string[4];
                bytes[0] = ":MEAS:FUNC1 " + sUnit;
                bytes[1] = ":MEAS:FREQ " + dblFrequency;
                bytes[2] = ":MEAS:LEV " + dblVoltage;
                bytes[3] = ":MEAS:TRIG";

                if (PreCompPN == "" || PreCompPN != txtCompPN.Text.ToUpper())
                {
                    InitPort(5);
                    PreCompPN = txtCompPN.Text.ToUpper();
                    serialPort.WriteLine(bytes[0]);
                    serialPort.WriteLine(bytes[1]);
                    serialPort.WriteLine(bytes[2]);
                    Thread.Sleep(1000);
                }
                serialPort.WriteLine(bytes[3]);
                Thread.Sleep(1000);

                while (Asc(Result.Substring(Result.Length - 1, 1)) < 20)
                {
                    Result = Result.Substring(0, Result.Length - 1);
                }

                return SplitString(Result);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + '[' + Result + ']', "MeasureE4300_RS232");
                return "0";
            }

        }

        private string MeasureE6420_RS232(string sEquipType, string sUnit, double dblFrequency, double dblVoltage)
        {
            try
            {
                if (!serialPort.IsOpen)
                {
                    InitPort(5);
                }
                ReceiveData = "";
                string[] bytes = new string[6];
                bytes[0] = ":MEAS";
                bytes[1] = ":MEAS:TEST:AC";
                bytes[2] = ":MEAS:FUNC:" + sUnit;
                bytes[3] = ":MEAS:LEVEL " + dblVoltage + ";FREQ " + dblFrequency + ";";
                bytes[4] = ":MEAS:RANGE AUTO";
                bytes[5] = ":MEAS:TRIG";

                if (PreCompPN == "" || PreCompPN != txtCompPN.Text.ToUpper())
                {
                    InitPort(5);
                    PreCompPN = txtCompPN.Text.ToUpper();
                    serialPort.WriteLine(bytes[0]);
                    serialPort.WriteLine(bytes[1]);
                    serialPort.WriteLine(bytes[2]);
                    serialPort.WriteLine(bytes[3]);
                    serialPort.WriteLine(bytes[4]);
                    Thread.Sleep(1000);
                }

                serialPort.WriteLine(bytes[5]);
                Thread.Sleep(1000);
                return SplitString(Result);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + '[' + Result + ']', "MeasureE6420_RS232");
                return "0";
            }
        }

        private string Measure3523_RS232(string sEquipType, string sUnit, double dblFrequency, double dblVoltage)
        {
            try
            {
                if (!serialPort.IsOpen)
                {
                    InitPort(5);
                }
                ReceiveData = "";
                string[] bytes = new string[6];
                bytes[0] = "*RST" + "\r\n";
                bytes[1] = ":TRIGger EXTernal" + "\r\n";
                bytes[2] = ":FREQuency 120" + "\r\n";
                if (sUnit.ToUpper() == "C")
                {
                    bytes[3] = ":PARameter1 CS;:PARameter2 RS" + "\r\n";
                }
                else if (sUnit.ToUpper() == "R")
                {
                    bytes[3] = ":PARameter1 RS;:PARameter2 CS" + "\r\n";
                }
                else if (sUnit.ToUpper() == "L")
                {
                    bytes[3] = ":PARameter1 LS;:PARameter2 RS" + "\r\n";
                }
                bytes[4] = "*TRG" + "\r\n";
                bytes[5] = ":MEASure?" + "\r\n";

                if (PreCompPN == "" || PreCompPN != txtCompPN.Text.ToUpper())
                {
                    InitPort(5);
                    PreCompPN = txtCompPN.Text.ToUpper();
                    serialPort.WriteLine(bytes[0]);
                    serialPort.WriteLine(bytes[1]);
                    serialPort.WriteLine(bytes[2]);
                    serialPort.WriteLine(bytes[3]);
                    serialPort.WriteLine(bytes[4]);
                    Thread.Sleep(1000);
                }

                serialPort.WriteLine(bytes[5]);
                Thread.Sleep(1000);
                return SplitString(Result);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + '[' + Result + ']', "Measure3523_RS232");
                return "0";
            }
        }

        private string MeasureE8110G_RS232(string sEquipType, string sUnit, double dblFrequency, double dblVoltage)
        {
            try
            {

                if (!serialPort.IsOpen)
                {
                    InitPort(1);
                }
                ReceiveData = "";
                string[] bytes = new string[6];
                bytes[0] = ":MEAS:SPEED SLOW";
                bytes[1] = ":MEAS:TEST:AC";
                bytes[2] = ":MEAS:FUNC " + sUnit;
                bytes[3] = ":MEAS:LEVEL " + dblVoltage + ";FREQ " + dblFrequency + ";";
                bytes[4] = ":MEAS:RANGE AUTO";
                bytes[5] = ":MEAS:TRIG";

                if (PreCompPN == "" || PreCompPN != txtCompPN.Text.ToUpper())
                {
                    InitPort(1);
                    PreCompPN = txtCompPN.Text.ToUpper();
                    serialPort.WriteLine(bytes[0]);
                    serialPort.WriteLine(bytes[1]);
                    serialPort.WriteLine(bytes[2]);
                    serialPort.WriteLine(bytes[3]);
                    serialPort.WriteLine(bytes[4]);
                    Thread.Sleep(1000);
                }

                serialPort.WriteLine(bytes[5]);
                Thread.Sleep(2000);

                //if (PreStr == "")
                //{
                //    Thread.Sleep(3000);
                //    PreStr = "Y";
                //}
                //else
                //{
                //    Thread.Sleep(1000);
                //}
                if (sUnit == "C")
                {
                    return SplitStringNum(Result, 0);
                }
                else
                {
                    return SplitStringNum(Result, 1);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + '[' + Result + ']', "MeasureE8110G_RS232");
                return "0";
            }
        }
        private string Measure_TH2832(string sEquipType, string sUnit, double dblFrequency, double dblVoltage)  //Leslie  add TH3832  0001
        {
            try
            {
                if (!serialPort.IsOpen)
                {
                    InitPort(5);
                }
                ReceiveData = "";
                string[] bytes = new string[1];
                bytes[0] = "fetch?";
                //bytes[1] = ":TRIGger EXTernal" + "\r\n";
                //bytes[2] = ":FREQuency 120" + "\r\n";
                //if (sUnit.ToUpper() == "C")
                //{
                //    bytes[3] = ":PARameter1 CS;:PARameter2 RS" + "\r\n";
                //}
                //else if (sUnit.ToUpper() == "R")
                //{
                //    bytes[3] = ":PARameter1 RS;:PARameter2 CS" + "\r\n";
                //}
                //else if (sUnit.ToUpper() == "L")
                //{
                //    bytes[3] = ":PARameter1 LS;:PARameter2 RS" + "\r\n";
                //}
                //bytes[4] = "*TRG" + "\r\n";
                //bytes[5] = ":MEASure?" + "\r\n";

                if (PreCompPN == "" || PreCompPN != txtCompPN.Text.ToUpper())
                {
                    InitPort(5);
                    PreCompPN = txtCompPN.Text.ToUpper();
                    serialPort.WriteLine(bytes[0]);
                    //serialPort.WriteLine(bytes[1]);
                    //serialPort.WriteLine(bytes[2]);
                    //serialPort.WriteLine(bytes[3]);
                    //serialPort.WriteLine(bytes[4]);
                    Thread.Sleep(1000);
                }

                serialPort.WriteLine(bytes[0]);
                Thread.Sleep(1000);             
                return SplitStringNum(Result, 0);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + '[' + Result + ']', "MeasureTH_TH2832-Leslie");             
                return "0";
            }
        }

        private void InitPort(int Threshold)
        {
            try
            {
                CloseSerialPort();
                serialPort.BaudRate = 9600;
                serialPort.PortName = cboCom.Text.ToString().Trim().ToUpper(); ;
                serialPort.DataBits = 8;
                serialPort.Parity = Parity.None;
                serialPort.StopBits = StopBits.One;
                serialPort.WriteTimeout = 1000;
                serialPort.ReadTimeout = 2000;

                serialPort.ReceivedBytesThreshold = Threshold;
                serialPort.WriteBufferSize = 512;
                serialPort.ReadBufferSize = 512;

                serialPort.DtrEnable = true;
                serialPort.RtsEnable = true;
                serialPort.DataReceived += new SerialDataReceivedEventHandler(COM_DataReceived);
                if (!serialPort.IsOpen)
                {
                    serialPort.Open();
                }
                else
                {
                    MessageBox.Show("Please Open Relative SerialPort");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "InitPort");
            }
        }

        private void COM_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                if (!serialPort.IsOpen)
                {
                    serialPort.Open();
                }
                Thread.Sleep(80);
                string str = "";
                byte[] buffer = new byte[serialPort.BytesToRead];
                serialPort.Read(buffer, 0, buffer.Length);
                str = System.Text.Encoding.ASCII.GetString(buffer);

                ReceiveData = ReceiveData + str;
                ReceiveData.Replace("\n", "");
                ReceiveData.Replace("\r\n", "");
                ReceiveData.Replace(" ", "");
                Result = ReceiveData;
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + '[' + Result + ']', "COM_DataReceived");
            }
        }

        private string SplitString(string strResult)
        {
            if (strResult == "")
            {
                return "";
            }
            char[] a = { ',' };
            string[] str = strResult.Split(a, StringSplitOptions.RemoveEmptyEntries);
            return str[0];
        }

        private string SplitStringNum(string strResult, int Num)
        {
            try
            {
                if (strResult == "")
                {
                    return "";
                }
                char[] a = { ',' };
                string[] str = strResult.Split(a, StringSplitOptions.RemoveEmptyEntries);
                return str[Num];
            }
            catch
            {
                return "";
            }
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            try
            {
                folderBrowserDialog1.SelectedPath = "C:\\Desktop";
                if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                {
                    txtimagePath.Text = folderBrowserDialog1.SelectedPath;
                    ImagePath = txtimagePath.Text;
                }
                else
                {
                    return;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CloseSerialPort()
        {
            if (serialPort.IsOpen == true)
            {
                serialPort.Close();
            }
        }

        private int Asc(string character)
        {
            if (character.Length == 1)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                int intAsciiCode = (int)asciiEncoding.GetBytes(character)[0];
                return (intAsciiCode);
            }
            else
            {
                throw new Exception("Character is not valid.");
            }

        }

    }
}
