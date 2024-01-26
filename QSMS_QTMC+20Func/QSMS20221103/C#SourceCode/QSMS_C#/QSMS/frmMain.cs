 using System;
using System.Net;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS
{
    public partial class frmMain : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        public frmMain()
        {
            InitializeComponent();
        }
        //private static bool returnDIDflag;  //改为公共变量 Parameter.returnDIDflag；
        DbLibrary.DbLogin login = new DbLibrary.DbLogin();

        private void frmMain_Load(object sender, EventArgs e)
        {
            try
            {
                string[] qwe = { "7","4","4","1","1"};
                int i = qwe.GetUpperBound(0); //4

                if (Application.StartupPath.ToString().Substring(Application.StartupPath.ToString().Length - 1) == "\\")
                {
                    Parameter.WorkDir = Application.StartupPath.ToString().Substring(0, Application.StartupPath.ToString().Length - 1);
                }
                else
                {
                    Parameter.WorkDir = Application.StartupPath.ToString() + "\\";
                }
                Parameter.Profile = Parameter.WorkDir + "set.ini";
                Parameter.hSECTION = "COMMON";
                //GetSettings Profile, hSECTION    begin
                Parameter.Settings.PRNa_Port = QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "PRNa_Port", Parameter.Profile);
                Parameter.Settings.PRNa_Settings = QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "PRNa_Settings", Parameter.Profile);
                Parameter.Settings.LabelAFile = QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "LabelAFile", Parameter.Profile);
                Parameter.Settings.LabelSATOFIle = QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "LabelFIle_SATO", Parameter.Profile);
                Parameter.Settings.ChkDIDDispatch = QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "CheckDIDDispatch", Parameter.Profile);
                Parameter.Settings.UpdateJobSide = "N";
                Parameter.Settings.UpdateJobSide = QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "UpdateJobSide", Parameter.Profile);
                Parameter.Settings.AutoDispatchLabel = Parameter.WorkDir + QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "AutoDispatchLabel", Parameter.Profile);
                Parameter.Settings.AutoDispatchSatoLabel =Parameter.WorkDir + QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "AutoDispatchSatoLabel", Parameter.Profile);
                Parameter.Settings.DIDLabelGood = Parameter.WorkDir + QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "DIDLabelGood", Parameter.Profile);
                Parameter.Settings.DIDLabelBad = Parameter.WorkDir + QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "DIDLabelBad", Parameter.Profile);
                Parameter.Settings.DIDLabelSATOGood = Parameter.WorkDir + QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "DIDLabelSATOGood", Parameter.Profile);
                Parameter.Settings.DIDLabelSATOBad = Parameter.WorkDir + QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "DIDLabelSATOBad", Parameter.Profile);
                Parameter.Settings.CompPrintLabel = Parameter.WorkDir + QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "CompPrintLabel", Parameter.Profile);
                Parameter.Settings.CompPNLabelPrint = Parameter.WorkDir + QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "CompPNLabelPrint", Parameter.Profile);
                Parameter.Settings.KFLabel = Parameter.WorkDir + QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "KFLabel", Parameter.Profile);
                Parameter.Settings.AutoDispatchNewLabel = Parameter.WorkDir + QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "AutoDispatchNewLabel", Parameter.Profile);
                Parameter.Settings.AutoDispatchSatoNewLabel = Parameter.WorkDir + QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "AutoDispatchSatoNewLabel", Parameter.Profile);
                Parameter.Settings.DIDLabelPath = Parameter.WorkDir + QMSSDK.Br.FileSystem.Ini.ReadIniValue(Parameter.hSECTION, "DIDLabelPath", Parameter.Profile);
                // END
                Parameter.Check_NonAVL = QMSSDK.Br.FileSystem.Ini.ReadIniValue("QSMS", "Check_NonAVL", Parameter.Profile);
                Parameter.Check_AVL = QMSSDK.Br.FileSystem.Ini.ReadIniValue("QSMS", "Check_AVL", Parameter.Profile);
                Parameter.Check_DID = QMSSDK.Br.FileSystem.Ini.ReadIniValue("QSMS", "Check_DID", Parameter.Profile);
                Parameter.imagePath = QMSSDK.Br.FileSystem.Ini.ReadIniValue("QSMS", "ImagePath", Parameter.Profile);
                Parameter.TestFilepath = QMSSDK.Br.FileSystem.Ini.ReadIniValue("QSMS", "TestFilepath", Parameter.Profile);
                Parameter.IPQC_ChkVendorPN = QMSSDK.Br.FileSystem.Ini.ReadIniValue("QSMS", "IPQC_ChkVendorPN", Parameter.Profile);
                Parameter.IC_CompChk = QMSSDK.Br.FileSystem.Ini.ReadIniValue("QSMS", "IC_CompChk", Parameter.Profile);
                Parameter.PrtCallBKandReturn = QMSSDK.Br.FileSystem.Ini.ReadIniValue("QSMS", "PrtCallBKandReturn", Parameter.Profile);

                Parameter.IPDefine = QMSSDK.Br.FileSystem.Ini.ReadIniValue("QSMS", "IPDefine", Parameter.Profile);

                DataTable dt = login.GetXL_SiteData(Parameter.Factory, Parameter.PrtCallBKandReturn);
                if (dt.Rows.Count > 0)
                {
                    Parameter.PrtCallBKandReturn = dt.Rows[0]["PrtCallBKandReturn"].ToString().Trim();
                    Parameter.DIDHead = dt.Rows[0]["DIDHead"].ToString().Trim();
                    Parameter.AutoDispatchForAnotherBU = dt.Rows[0]["AutoDispatchForAnotherBU"].ToString().Trim();
                    Parameter.CheckPNGroup = dt.Rows[0]["CheckPNGroup"].ToString().Trim();
                    Parameter.BUDIDShow = dt.Rows[0]["BUDIDShow"].ToString().Trim();
                    Parameter.DIDnotToQWMS = dt.Rows[0]["DIDnotToQWMS"].ToString().Trim();
                    if (Parameter.DIDHead == "")
                    {
                        MessageBox.Show("Can't get DID Head from table [Site], please define it first!");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Can't get DID Head from table [Site], please define it first!");
                    return;
                }

                setMenu();                
                Parameter.LocalIP = LocalIP();


                if (Parameter.IPDefine == "Y")///20210922 Yan 改为定义
                {
                    CheckFacIPByDefine();
                }

                else
                {
                    CheckFacIP();
                }


                if (Parameter.Factory != "")
                {
                    LabFac.Text = "Your QSMS program Factory: " + Parameter.Factory + ", IP:" + Parameter.LocalIP;
                }
                else
                {
                    LabFac.Text = "Your IP:" + Parameter.LocalIP + ", it is not defined";
                }
                if (Parameter.CreateDIDFlag == "N")
                {
                    menuAutoDispatch.Enabled = false;
                    mnuReturnDID.Enabled = false;
                    mnuReturnDIDALL.Enabled = false;
                    mnuReturnComp.Enabled = false;
                }
                this.Text = this.Text + Parameter.Version + " IP: " + Parameter.IP + "; Factory:" + Parameter.Factory;

                Parameter.UID = Parameter.g_userName;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        //private void setMenu()
        //{
        //    mnuQueryByReturnDIDed.Visible = true;
        //    mnuTransferFujiNexim.Visible = false;
        //    mnuMCCPreMaterial.Visible = true;
        //    mnuInheritDIDByWO.Visible = true;
        //    mnuDispatchDIDAdditionnal.Visible = false;
        //    mnuUpLoadBom.Visible = false;
        //    mnuMaintainDID.Visible = true;
        //    mnuModfyDIDTotalQty.Visible = true;
        //    mnuReturnDID.Visible = true;
        //    mnuReturnDIDALL.Visible = false;
        //    mnuDIDCallBack.Visible = true;
        //    mnuDIDChkStock.Visible = true;
        //    mnuTransferFujiXML.Visible = true;
        //    mnuTransferPanaMSF.Visible = true;
        //    mnuTransferPanaAMI.Visible = true;
        //    mnuSAP1DataPatch.Visible = false;
        //    mnuDefineBuildType.Visible = false;
        //    menuAutoDispatch.Visible = true;
        //    mmuWOInputPlan.Visible = false;
        //    mmuQSMS_SapHis.Visible = false;
        //    mnuReturnComp.Visible = true;
        //    mmuDIDIntegration.Visible = true;
        //    mnuCompPrint.Visible = true;
        //    mnuFixDispatchData.Visible = false;
        //    mnuSpecialReturnArchive.Visible = false;
        //    mnuCostCenter.Visible = false;
        //    mnuIC_Burn.Visible = false;
        //    mmuQSMS_Record_DIDInfo.Visible = false;

        //    //PMC
        //    mnumaintainWOSeq.Visible = true;
        //    mnuQueryWOGroup.Visible = true;
        //    //PD
        //    mnuupdRealqty.Visible = true;
        //    mnumaintainFeeder.Visible = true;
        //    mnuVerifyFeederSlot.Visible = true;
        //    mnuDeleteFeeder.Visible = true;
        //    mnuCloseWO.Visible = true;
        //    mnuDIDBake.Visible = true;
        //    mmuUnlockCompPNCompare.Visible = false;
        //    mnuIPQCRelieve.Visible = false;
        //    //Report
        //    mnfrmBeforeCheckBom.Visible = true;
        //    mnuTraceReport.Visible = true;
        //    mnuWipReport.Visible = true;
        //    mnuQueryDID.Visible = true;
        //    munPanelDiff.Visible = true;
        //    munQueryReplacePN.Visible = true;

        //    //IPQC
        //    mnuDelete.Visible = true;

        //    //QMS
        //    mnuSetDIOandInterlock.Visible = true;
        //    mnuCheckDispatchQty.Visible = false;
        //    mnuDeleteME_BOM.Visible = true;
        //    mnuPrinterSetting.Visible = true;

        //    //Special Case
        //    mnuUnChkWO.Visible = true;
        //}

        private void setMenu()
        {
            mnuMCCPreMaterial.Enabled = false;
            mnuInheritDIDByWO.Enabled = false;
            mnuTransferDispatchDID.Enabled = false;
            mnuUpLoadBom.Enabled = false;
            mnuMaintainDID.Enabled = false;
            mnuModfyDIDTotalQty.Enabled = false;
            mnuReturnDID.Enabled = false;
            mnuReturnDIDALL.Enabled = false;
            mnuDIDCallBack.Enabled = false;
            mnuDIDChkStock.Enabled = false;
            mnuTransferFujiXML.Enabled = false;
            mnuTransferPanaMSF.Enabled = false;
            mnuTransferPanaAMI.Enabled = false;
            mnuSAP1DataPatch.Enabled = false;
            mnuDefineBuildType.Enabled = false;
            menuAutoDispatch.Enabled = false;
            mmuWOInputPlan.Enabled = false;
            mmuQSMS_SapHis.Enabled = false;
            mnuReturnComp.Enabled = false;
            mmuDIDIntegration.Enabled = false;
            mnuCompPrint.Enabled = true;
            mnuFixDispatchData.Enabled = false;
            mnuSpecialReturnArchive.Enabled = false;

            //PMC
            mnumaintainWOSeq.Enabled = false;
            mnuQueryWOGroup.Enabled = false;
            //PD
            mnuupdRealqty.Enabled = false;
            mnumaintainFeeder.Enabled = false;
            mnuVerifyFeederSlot.Enabled = false;
            mnuDeleteFeeder.Enabled = false;
            mnuCloseWO.Enabled = false;
            mnuDIDBake.Enabled = false;
            mmuUnlockCompPNCompare.Enabled = false;

            //Report
            mnuWipReport.Enabled = false;
            mnuQueryDID.Enabled = false;

            //IPQC
            mnuInSpection.Enabled = false;

            //QMS
            mnuSetDIOandInterlock.Enabled = false;
            mnuCheckDispatchQty.Enabled = false;
            mnuSendXLRemainDemand.Enabled = false;
            mnuDeleteME_BOM.Enabled = false;
            mnuPrinterSetting.Enabled = true;

            //Special Case
            mnuUrgentInsertWO.Enabled = false;
            mnuGenXLMD.Enabled = false;
            mnuTransferFujiAVL.Enabled = false;
            mnuStartSplitLineMC.Enabled = false;
            mnuUpdateUID.Enabled = false;
            mmuQSMS_Record_DIDInfo.Enabled = false;
            Parameter.strKeyInPNByManual = false;

            for (int i = Parameter.g_userRight.GetLowerBound(0); i <= Parameter.g_userRight.GetUpperBound(0); i++)
            {
                if (Parameter.g_userRight[i] == "mnuMCCPreMaterial")
                {
                    mnuMCCPreMaterial.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuInheritDIDByWO")
                {
                    mnuInheritDIDByWO.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuDispatchDIDAdditionnal")
                {
                    mnuTransferDispatchDID.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "MnuUpLoadBom")
                {
                    mnuUpLoadBom.Enabled = true;
                    mnuDeleteME_BOM.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuMaintainDID")
                {
                    mnuMaintainDID.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuModfyDIDTotalQty")
                {
                    mnuModfyDIDTotalQty.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuReturnDID")
                {
                    mnuReturnDID.Enabled = true;
                    mnuReturnComp.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuDIDChkStock")
                {
                    mnuDIDChkStock.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuReturnDIDALL")
                {
                    mnuReturnDIDALL.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuDIDCallBack")
                {
                    mnuDIDCallBack.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuTransferFujiXML")
                {
                    mnuTransferFujiXML.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuTransferFujiNexim")
                {
                    mnuTransferFujiNexim.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuTransferPanaAMI")
                {
                    mnuTransferPanaAMI.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuTransferPanaMSF")
                {
                    mnuTransferPanaMSF.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuSAP1DataPatch")
                {
                    mnuSAP1DataPatch.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuDefineBuildType")
                {
                    mnuDefineBuildType.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuMaintainDIDAutoDispatch")
                {
                    menuAutoDispatch.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuDIDBake")
                {
                    mnuDIDBake.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mmuWOInputPlan")
                {
                    mmuWOInputPlan.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mmuQSMS_SapHis")
                {
                    mmuQSMS_SapHis.Enabled = true;
                }
                //PMC
                if (Parameter.g_userRight[i] == "mnumaintainWOSeq")
                {
                    mnumaintainWOSeq.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuQueryWOGroup")
                {
                    mnuQueryWOGroup.Enabled = true;
                }
                //PD
                if (Parameter.g_userRight[i] == "mnumaintainFeeder")
                {
                    mnumaintainFeeder.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuVerifyFeederSlot")
                {
                    mnuVerifyFeederSlot.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuDeleteFeeder")
                {
                    mnuDeleteFeeder.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuClearDIDSplicing")
                {
                    mnuDeleteFeeder.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuCloseWO")
                {
                    mnuCloseWO.Enabled = true;
                }
                //Report
                if (Parameter.g_userRight[i] == "mnuWipReport")
                {
                    mnuWipReport.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuQueryDID")
                {
                    mnuQueryDID.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "returnDIDflag")
                {
                    Parameter.returnDIDflag = true;
                }
                //IPQC
                if (Parameter.g_userRight[i].ToUpper() == "MNUINSPECTION")
                {
                    mnuInSpection.Enabled = true;
                }
                if (Parameter.g_userRight[i].ToUpper() == "MNUDELETE")
                {
                    mnuDelete.Enabled = true;
                }
                if (Parameter.g_userRight[i].ToUpper() == "MNURELIEVE")
                {
                    mnuIPQCRelieve.Enabled = true;
                }
                //QMS
                if (Parameter.g_userRight[i] == "mnuSetDIOandInterlock")
                {
                    mnuSetDIOandInterlock.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuCheckDispatchQty")
                {
                    mnuCheckDispatchQty.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuSendXLRemainDemand")
                {
                    mnuSendXLRemainDemand.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mmuUnlockCompPNCompare")
                {
                    mmuUnlockCompPNCompare.Enabled = true;
                }
                //Special Case
                if (Parameter.g_userRight[i] == "mnuUrgentInsertWO")
                {
                    mnuUrgentInsertWO.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuGenXLMD")
                {
                    mnuGenXLMD.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuUrgentDIDToWH")
                {
                    mnuUrgentDIDToWH.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuUpdRealQty")
                {
                    mnuupdRealqty.Enabled = true;
                }
                if (Parameter.g_userRight[i] == "mnuTransferFujiAVL")
                {
                    mnuTransferFujiAVL.Enabled = true;
                }
                if (Parameter.g_userRight[i].ToUpper() == "CHECKBOM")
                {
                    Parameter.CheckBomRight = true;
                }
                if (Parameter.g_userRight[i].ToUpper() == "MNUFIXDISPATCHDATA")
                {
                    mnuFixDispatchData.Enabled = true;
                }
                if (Parameter.g_userRight[i].ToUpper() == ("KeyInPNByManual").ToUpper())
                {
                    Parameter.strKeyInPNByManual = true;
                }
                if (Parameter.g_userRight[i].ToUpper() == ("mmuDIDintegration").ToUpper())
                {
                    mmuDIDIntegration.Enabled = true;
                }
                if (Parameter.g_userRight[i].ToUpper() == ("DeleteMeBomByLine").ToUpper())
                {
                    Parameter.DeleteMeBomByLine = true;
                }
                if (Parameter.g_userRight[i].ToUpper() == ("mnuStartSplitLineMC").ToUpper())
                {
                    mnuStartSplitLineMC.Enabled = true;
                }
                if (Parameter.g_userRight[i].ToUpper() == ("mnuUpdateUID").ToUpper())
                {
                    mnuUpdateUID.Enabled = true;
                }
                if (Parameter.g_userRight[i].ToUpper() == ("mmuQSMS_Record_DIDInfo").ToUpper())
                {
                    mmuQSMS_Record_DIDInfo.Enabled = true;
                }
                if (Parameter.g_userRight[i].ToUpper() == ("mnuSpecialReturnArchive").ToUpper())
                {
                    mnuSpecialReturnArchive.Enabled = true;
                }
            }
        }

        private string LocalIP()
        {
            string AddressIP = string.Empty;
            foreach (IPAddress _IPAddress in Dns.GetHostEntry(Dns.GetHostName()).AddressList)
            {
                if (_IPAddress.AddressFamily.ToString() == "InterNetwork")
                {
                    AddressIP = _IPAddress.ToString();
                }
            }

            return AddressIP;
        }

        private bool CheckFacIP()
        {
            string[] strIP;
            Parameter.Factory = "";
            Parameter.CreateDIDFlag = "N";
            DataTable dt = login.CheckFacIP();
            if (dt.Rows.Count < 1)
            {
                MessageBox.Show("The Factory is empty,please connect with QMS for set the Factory in the Site table!");
                return false;
            }
            if(dt.Rows.Count > 1)
            {
                string[,] FactoryID = new string[dt.Rows.Count,2];
                for (int i = 0; i < dt.Rows.Count;i++ )
                {
                    FactoryID[i, 0] = dt.Rows[i]["Factory"].ToString();
                    //FactoryID[i, 1] = QMSSDK.Br.FileSystem.Ini.ReadIniValue("QSMS", "Factory", Parameter.Profile);
                    FactoryID[i, 1] = QMSSDK.Br.FileSystem.Ini.ReadIniValue("QSMS", FactoryID[i, 0], Parameter.Profile);
                }
                for (int i = 0; i <= FactoryID.GetUpperBound(0);i++ )
                {
                    if(FactoryID[i,0].Trim() != "" && FactoryID[i,1].Trim() == "")
                    {
                        MessageBox.Show("Your BU produce in " + FactoryID[i, 0] + " factories,please connect with QMS for set the " + FactoryID[i, 0] + " IP!");
                        return false;
                    }
                    strIP = FactoryID[i, 1].Split(new char[] { ';' });
                    for (int j = 0; j <= strIP.GetUpperBound(0);j++ )
                    {
                        if (strIP[j] == Parameter.LocalIP.Substring(0, strIP[j].Length) && strIP[j] !="")
                        {
                            if (Parameter.Factory.Trim() != "")
                            {
                                MessageBox.Show("Your IP " + Parameter.LocalIP + " is exist in different factory,please connect with QMS check!");
                                return false;
                            }
                            else
                            {
                                Parameter.Factory = FactoryID[i, 0];
                                Parameter.CreateDIDFlag = "Y";
                            }
                        }
                    }
                }
            }
            else
            {
                Parameter.Factory = dt.Rows[0]["Factory"].ToString().Trim();
                Parameter.CreateDIDFlag = "Y";
            }
            return true;
        }


        private bool CheckFacIPByDefine()
        {
            Parameter.Factory = "";
            Parameter.CreateDIDFlag = "N";

            DataTable dt1 = login.QSMS_IPFactory(Parameter.Factory, Parameter.LocalIP);
            if (dt1.Rows[0]["RE"].ToString().Trim()=="Fail")
            {
                MessageBox.Show("The local IP does not match the defined factory area!");
                return false;
            }
            else
            {
                Parameter.Factory = dt1.Rows[0]["Factory"].ToString().Trim();
                Parameter.CreateDIDFlag = "Y";
            }

            return true;
        }



        private void mnuUpLoadBom_Click(object sender, EventArgs e)
        {
            QSMS.MCC.frmUpLoadBom frmUpLoadBom = new QSMS.MCC.frmUpLoadBom();
            pubFunction.HaveOpened(frmUpLoadBom, "frmUpLoadBom");            
        }

        private void menuAutoDispatch_Click(object sender, EventArgs e)
        {
            QSMS.MCC.frmMaintainDIDAutoDispatch MaintainDIDAutoDispatch = new QSMS.MCC.frmMaintainDIDAutoDispatch();
            pubFunction.HaveOpened(MaintainDIDAutoDispatch, "frmMaintainDIDAutoDispatch");
        }

        private void mnuTransferPanaMSF_Click(object sender, EventArgs e)
        {
            QSMS.MCC.frmTransferPanaMSF frmTransferPanaMSF = new QSMS.MCC.frmTransferPanaMSF();
            frmTransferPanaMSF.Show();
        }

        private void mmuCheckDID_Click(object sender, EventArgs e)
        {
            QSMS.PD.frmChecKDID frm = new QSMS.PD.frmChecKDID();
            pubFunction.HaveOpened(frm, "frmChecKDID");
        }

        private void mnuCloseWO_Click(object sender, EventArgs e)
        {
            QSMS.PD.frmCloseWO frm = new QSMS.PD.frmCloseWO();
            pubFunction.HaveOpened(frm, "frmCloseWO");
        }

        private void mnuInSpection_Click(object sender, EventArgs e)
        {
            QSMS.IPQC.frmInSpection frm = new QSMS.IPQC.frmInSpection();
            pubFunction.HaveOpened(frm, "frmInSpection");
        }

        private void mnuCompPrint_Click(object sender, EventArgs e)
        {
            QSMS.MCC.frmCompPrint frm = new QSMS.MCC.frmCompPrint();
            pubFunction.HaveOpened(frm, "frmCompPrint");
        }

        private void mnuMaintainDID_Click(object sender, EventArgs e)
        {
            QSMS.MCC.frmMaintainDID frm = new QSMS.MCC.frmMaintainDID();
            pubFunction.HaveOpened(frm, "frmMaintainDID");
        }

        private void mnuSetDIOandInterlock_Click(object sender, EventArgs e)
        {
            QSMS.QMS.frmSetInterDIO frm = new QSMS.QMS.frmSetInterDIO();
            pubFunction.HaveOpened(frm, "frmSetInterDIO");
        }

        private void mnuWipReport_Click(object sender, EventArgs e)
        {
            QSMS.Report.frmReport frmReport = new QSMS.Report.frmReport();
            //pubFunction.HaveOpened(frmReport, "frmReport");
            frmReport.Show();
        }

        private void mnuUrgentDIDToWH_Click(object sender, EventArgs e)
        {
            QSMS.SpecialCase.frmUrgentDIDToWH frm = new QSMS.SpecialCase.frmUrgentDIDToWH();
            pubFunction.HaveOpened(frm, "frmUrgentDIDToWH");
        }

        private void mnuUpdateUID_Click(object sender, EventArgs e)
        {
            QSMS.SpecialCase.frmUpdateUID frmUpdateUID = new QSMS.SpecialCase.frmUpdateUID();
            pubFunction.HaveOpened(frmUpdateUID, "frmUpdateUID");  
        }

        private void mnuUrgentInsertWO_Click(object sender, EventArgs e)
        {
            QSMS.SpecialCase.frmUrgentWO frmUrgentWO = new QSMS.SpecialCase.frmUrgentWO();
            pubFunction.HaveOpened(frmUrgentWO, "frmUrgentWO"); 
        }

        private void mnuStartSplitLineMC_Click(object sender, EventArgs e)
        {
            QSMS.SpecialCase.frmStartSplitLineMC frmStartSplitLineMC = new QSMS.SpecialCase.frmStartSplitLineMC();
            pubFunction.HaveOpened(frmStartSplitLineMC, "frmStartSplitLineMC");
        }

        private void mnuupdRealqty_Click(object sender, EventArgs e)
        {
            QSMS.PD.frmUpdateRealQty frm = new QSMS.PD.frmUpdateRealQty();
            pubFunction.HaveOpened(frm, "frmUpdateRealQty");
        }

        private void mnuQueryWOGroup_Click(object sender, EventArgs e)
        {
            QSMS.PMC.frmQueryWOGroup frm = new QSMS.PMC.frmQueryWOGroup();
            pubFunction.HaveOpened(frm, "frmQueryWOGroup");
        }

        private void mnuQDIDNeedCut_Click(object sender, EventArgs e)
        {
            QSMS.Report.frmQueryDIDNeedCut frm = new QSMS.Report.frmQueryDIDNeedCut();
            pubFunction.HaveOpened(frm, "frmQueryDIDNeedCut");
        }

        private void mnuGenXLMD_Click(object sender, EventArgs e)
        {
            QSMS.SpecialCase.frmGenXLMD frm = new QSMS.SpecialCase.frmGenXLMD();
            pubFunction.HaveOpened(frm, "GenXLMaterialDemand");
        }

        private void mnuDIDNoUsed_Click(object sender, EventArgs e)
        {
            QSMS.Report.frmDIDNoUsed frmDIDNoUsed = new QSMS.Report.frmDIDNoUsed();
            pubFunction.HaveOpened(frmDIDNoUsed, "frmDIDNoUsed");
        }

        private void mnuCompPNReport_Click(object sender, EventArgs e)
        {
            QSMS.Report.frmComppnReport frmComppnReport = new QSMS.Report.frmComppnReport();
            pubFunction.HaveOpened(frmComppnReport, "frmComppnReport");
        }

        private void mnuQuery_Click(object sender, EventArgs e)
        {
            QSMS.IPQC.frmQuery frmComppnReport = new QSMS.IPQC.frmQuery();
            pubFunction.HaveOpened(frmComppnReport, "frmQuery");
        }

        private void mnuReturnComp_Click(object sender, EventArgs e)
        {
            QSMS.MCC.FrmReturnComp FrmReturnComp = new QSMS.MCC.FrmReturnComp();
            pubFunction.HaveOpened(FrmReturnComp, "FrmReturnComp");
        }

        private void mnuDIDBake_Click(object sender, EventArgs e)
        {
            QSMS.PD.FrmDIDBake FrmDIDBake = new QSMS.PD.FrmDIDBake();
            pubFunction.HaveOpened(FrmDIDBake, "FrmDIDBake");
        }

        private void mnumaintainWOSeq_Click(object sender, EventArgs e)
        {
            QSMS.PMC.frmMaintainWOSeq frmMaintainWOSeq = new QSMS.PMC.frmMaintainWOSeq();
            pubFunction.HaveOpened(frmMaintainWOSeq, "frmMaintainWOSeq");
        }

        private void mnuQueryDID_Click(object sender, EventArgs e)
        {
            QSMS.Report.frmQueryDID frmQueryDID = new QSMS.Report.frmQueryDID();
            pubFunction.HaveOpened(frmQueryDID, "frmQueryDID");
        }

        private void mnuUnChkWO_Click(object sender, EventArgs e)
        {
            QSMS.SpecialCase.frmUnChkWO frmUnChkWO = new QSMS.SpecialCase.frmUnChkWO();
            pubFunction.HaveOpened(frmUnChkWO, "frmUnChkWO");
        }

        private void mnuDelete_Click(object sender, EventArgs e)
        {
            QSMS.IPQC.frmInspection_Del frmInspection_Del = new QSMS.IPQC.frmInspection_Del();
            pubFunction.HaveOpened(frmInspection_Del, "frmInspection_Del");
        }

        private void mnuQueryCheckBOM_Click(object sender, EventArgs e)
        {
            QSMS.Report.frmQueryCheckBOM frmQueryCheckBOM = new QSMS.Report.frmQueryCheckBOM();
            pubFunction.HaveOpened(frmQueryCheckBOM, "frmQueryCheckBOM");
        }

        private void mnuCheckDispatchQty_Click(object sender, EventArgs e)
        {
            QSMS.QMS.frmCompDiff frmCompDiff = new QSMS.QMS.frmCompDiff();
            pubFunction.HaveOpened(frmCompDiff, "frmCompDiff");
        }

        private void mnuSendXLRemainDemand_Click(object sender, EventArgs e)
        {
            QSMS.QMS.frmSendXLRemainDemand frmSendXLRemainDemand = new QSMS.QMS.frmSendXLRemainDemand();
            pubFunction.HaveOpened(frmSendXLRemainDemand, "frmSendXLRemainDemand");
        }

        private void mnuIPQCRelieve_Click(object sender, EventArgs e)
        {
            QSMS.IPQC.frmInRelieve frmInRelieve = new QSMS.IPQC.frmInRelieve();
            pubFunction.HaveOpened(frmInRelieve, "frmInRelieve");
        }

        private void mnuModfyDIDTotalQty_Click(object sender, EventArgs e)
        {
            QSMS.MCC.FrmModifyDIDTotalQty FrmModifyDIDTotalQty = new QSMS.MCC.FrmModifyDIDTotalQty();
            pubFunction.HaveOpened(FrmModifyDIDTotalQty,"FrmModifyDIDTotalQty");
        }

        private void mnuDeleteME_BOM_Click(object sender, EventArgs e)
        {
            QSMS.MCC.FrmDeleteME_BOM FrmDeleteME_BOM = new QSMS.MCC.FrmDeleteME_BOM();
            pubFunction.HaveOpened(FrmDeleteME_BOM, "FrmDeleteME_BOM");
        }

        private void mnuTransferPanaAMI_Click(object sender, EventArgs e)
        {
            QSMS.MCC.frmTransferPanaAMI frmTransferPanaAMI = new QSMS.MCC.frmTransferPanaAMI();
            pubFunction.HaveOpened(frmTransferPanaAMI, "frmTransferPanaAMI");
        }

        private void mnuInheritDIDByWO_Click(object sender, EventArgs e)
        {
            QSMS.MCC.frmInheritDIDByWO FrmInheritDIDByWO = new QSMS.MCC.frmInheritDIDByWO();
            FrmInheritDIDByWO.Show();
        }

        private void mnuReturnDID_Click(object sender, EventArgs e)
        {
            QSMS.MCC.FrmReturnDID FrmReturnDID = new QSMS.MCC.FrmReturnDID();
            pubFunction.HaveOpened(FrmReturnDID, "FrmReturnDID");
        }

        private void mnuMCCPreMaterial_Click(object sender, EventArgs e)
        {
            QSMS.MCC.frmMCCPreMaterial frmMCCPreMaterial = new QSMS.MCC.frmMCCPreMaterial();
            pubFunction.HaveOpened(frmMCCPreMaterial, "frmMCCPreMaterial");
        }

        private void mnuDIDChkStock_Click(object sender, EventArgs e)
        {
            QSMS.MCC.frmDIDCheckStock frmDIDCheckStock = new QSMS.MCC.frmDIDCheckStock();
            pubFunction.HaveOpened(frmDIDCheckStock, "frmDIDCheckStock");
        }

        private void mnuVerifyFeederSlot_Click(object sender, EventArgs e)
        {
            QSMS.PD.frmVerifyFeeder frmVerifyFeeder = new QSMS.PD.frmVerifyFeeder();
            pubFunction.HaveOpened(frmVerifyFeeder, "frmVerifyFeeder");
        }

        private void mnuTransferFujiXML_Click(object sender, EventArgs e)
        {
            QSMS.MCC.frmTransferFujiXML frmTransferFujiXML = new QSMS.MCC.frmTransferFujiXML();
            pubFunction.HaveOpened(frmTransferFujiXML, "frmTransferFujiXML");
        }

        private void mmuDIDIntegration_Click(object sender, EventArgs e)
        {
            QSMS.MCC.frmDIDInteGration frmDIDInteGration = new QSMS.MCC.frmDIDInteGration();
            pubFunction.HaveOpened(frmDIDInteGration, "frmDIDInteGration");
        }

        private void mnumaintainFeeder_Click(object sender, EventArgs e)
        {
            QSMS.PD.frmMaintainFeeder frmMaintainFeeder = new QSMS.PD.frmMaintainFeeder();
            pubFunction.HaveOpened(frmMaintainFeeder, "frmMaintainFeeder");
        }

        private void mnuTraceReport_Click(object sender, EventArgs e)
        {
            QSMS.Report.frmTraceReport frmTraceReport = new QSMS.Report.frmTraceReport();
            pubFunction.HaveOpened(frmTraceReport, "frmTraceReport");
        }

        private void mnuQueryByReturnDIDed_Click(object sender, EventArgs e)
        {
            QSMS.MCC.frmQryReturnedDID frmQryReturnedDID = new QSMS.MCC.frmQryReturnedDID();
            pubFunction.HaveOpened(frmQryReturnedDID, "frmQryReturnedDID");
        }

        private void mnuDeleteFeeder_Click(object sender, EventArgs e)
        {
            QSMS.PD.frmUnlinkFeederDID frmUnlinkFeederDID = new QSMS.PD.frmUnlinkFeederDID();
            pubFunction.HaveOpened(frmUnlinkFeederDID,"frmUnlinkFeederDID");
        }

        private void mnfrmBeforeCheckBom_Click(object sender, EventArgs e)
        {
            QSMS.Report.frmBeforeHandCheckBom frmBeforeHandCheckBom = new QSMS.Report.frmBeforeHandCheckBom();
            pubFunction.HaveOpened(frmBeforeHandCheckBom, "frmBeforeHandCheckBom");
        }

        private void munQueryReplacePN_Click(object sender, EventArgs e)
        {
            QSMS.Report.frmQueryReplacePN frmQueryReplacePN = new QSMS.Report.frmQueryReplacePN();
            pubFunction.HaveOpened(frmQueryReplacePN, "frmQueryReplacePN");
        }

        private void mnuDIDCallBack_Click(object sender, EventArgs e)
        {
            QSMS.MCC.FrmDIDCallBack FrmDIDCallBack = new QSMS.MCC.FrmDIDCallBack();
            pubFunction.HaveOpened(FrmDIDCallBack, "FrmDIDCallBack");
        }

        private void unlockCompPNCompare_Click(object sender, EventArgs e) //Rain 20210901
        {
            QSMS.PD.FrmUnlockCompPNCompare UnlockCompPNCompare = new QSMS.PD.FrmUnlockCompPNCompare();
            pubFunction.HaveOpened(UnlockCompPNCompare, "FrmUnlockCompPNCompare");
        }

        private void frmPNCompare_Click(object sender, EventArgs e) //Rain 20210901
        {
            QSMS.PD.FrmPNCompare FrmPNCompare = new QSMS.PD.FrmPNCompare();
            pubFunction.HaveOpened(FrmPNCompare, "frmPNCompare");
        }

        private void mnuPrinterSetting_Click(object sender, EventArgs e)
        {
            QSMS.frmPrinterSetting frmPrinterSetting = new QSMS.frmPrinterSetting();
            pubFunction.HaveOpened(frmPrinterSetting, "frmPrinterSetting");
        }

        private void mnuTransferDispatchDID_Click(object sender, EventArgs e)                                   //Aris  Add 20210915
        {
            QSMS.MCC.frmTransferDispatchDID frmTransferDispatchDID = new QSMS.MCC.frmTransferDispatchDID();         
            pubFunction.HaveOpened(frmTransferDispatchDID, "frmTransferDispatchDID");
        }

        private void mnuTransferFujiNexim_Click(object sender, EventArgs e)
        {
            QSMS.MCC.frmTransferFujiNexim frmTransferFujiNexim = new QSMS.MCC.frmTransferFujiNexim();
            pubFunction.HaveOpened(frmTransferFujiNexim, "frmTransferFujiNexim");
        }

        private void mnuTransferFujiNexim_MI_Click(object sender, EventArgs e)
        {
            QSMS.MCC.frmTransferFujiNexim_MI frmTransferFujiNexim_MI = new QSMS.MCC.frmTransferFujiNexim_MI();
            pubFunction.HaveOpened(frmTransferFujiNexim_MI, "frmTransferFujiNexim_MI");
        }

        private void deleteMEBOMToolStripMenuItem_Click(object sender, EventArgs e)
        {
            QSMS.MCC.FrmDefineBuildType FrmDefineBuildType = new QSMS.MCC.FrmDefineBuildType();
            pubFunction.HaveOpened(FrmDefineBuildType, "FrmDefineBuildType");
        }

        private void defineBuildTypeToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            QSMS.MCC.FrmDefineBuildType FrmDefineBuildType = new QSMS.MCC.FrmDefineBuildType();
            pubFunction.HaveOpened(FrmDefineBuildType, "FrmDefineBuildType");
        }

        private void mnuIC_Burn_Click(object sender, EventArgs e)
        {
            QSMS.MCC.IC_Burn IC_Burn = new QSMS.MCC.IC_Burn();
            pubFunction.HaveOpened(IC_Burn, "IC_Burn");
        }
        private void mnuCompPNPrint_Click(object sender, EventArgs e)
        {
            QSMS.MCC.CompPNPrint CompPNPrint = new QSMS.MCC.CompPNPrint();
            pubFunction.HaveOpened(CompPNPrint, "CompPNPrint");
        }

        private void compPNPrintToolStripMenuItem_Click(object sender, EventArgs e)
        {
            QSMS.MCC.CompPNPrint CompPNPrint = new QSMS.MCC.CompPNPrint();
            pubFunction.HaveOpened(CompPNPrint, "CompPNPrint");
        }

        private void mnuSingleSideBrdConfirm_Click(object sender, EventArgs e)
        {
               QSMS.MCC.FrmSingleSideBrdConfirm FrmSingleSideBrdConfirm = new QSMS.MCC.FrmSingleSideBrdConfirm();
               pubFunction.HaveOpened(FrmSingleSideBrdConfirm, "FrmSingleSideBrdConfirm");
        }

        //private void uploadXLPlanAndMaintainWOToolStripMenuItem_Click(object sender, EventArgs e)
        //{

        //}
        //private void singleSideBrdConfirmToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    QSMS.MCC.FrmSingleSideBrdConfirm FrmSingleSideBrdConfirm = new QSMS.MCC.FrmSingleSideBrdConfirm();
        //    pubFunction.HaveOpened(FrmSingleSideBrdConfirm, "FrmSingleSideBrdConfirm");
        //}

        private void uploadXLPlanAndMaintainWOToolStripMenuItem_Click(object sender, EventArgs e)  //20230314 Rain 同步台湾导入自动建Group
        {
            QSMS.MCC.frmUploadXLSchedule FrmUploadXLSchedule = new QSMS.MCC.frmUploadXLSchedule();
            pubFunction.HaveOpened(FrmUploadXLSchedule, "FrmUploadXLSchedule");
        }
    }
}
