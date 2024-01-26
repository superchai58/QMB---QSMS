using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using QSMS.QSMS.QMS;

namespace QSMS
{
    static class Program
    {
        
        private static string strTemp;
        private static DataTable dt;
        private static bool IsInDesign = false;
        
        
        /// <summary>
        /// 应用程序的主入口点。

        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            DbLibrary.DbLogin login = new DbLibrary.DbLogin();
            string strCommand = string.Empty;
            if (args.Length == 0)
            {
                if (!IsInDesign)
                {
                    //strCommand = "<APPNAME=QSMS><LINE=All><STATION=QSMS><CONN=PROVIDER=SQLOLEDB;UID=sa;SERVER=172.26.16.4;DATABASE=SMT;NETWORK LIBRARY=DBMSSOCN;pwd=East#86><USERID=QMS><BU=NB6><SERVERBU=NB6><CHKLOGIN=N><RIGHT=CheckBom,DeleteDID,DIDSlot,IPQCRetest,KeyInPNByManual,Login,mmuCompPNCompare,mmuDIDintegration,mmuQSMS_SapHis,mmuUnlockCompPNCompare,mmuWOInputPlan,mnuChangeExtraDIDslot,mnuCheckDispatchQty,mnuCloseWO,mnuCompPrint,mnuDefineBuildType,mnuDeleteDID,mnuDeleteFeeder,mnuDIDCallBack,mnuDIDChkStock,mnuDispatchDIDAdditionnal,mnuFixDispatchData,mnuForceDeleteDID,mnuGenXLMD,mnuInheritDIDByWO,MnuInSpection,mnuMaintainDID,mnuMaintainDIDAutoDispatch,mnumaintainFeeder,mnumaintainWOSeq,mnuMCCPreMaterial,mnuModfyDIDTotalQty,MnuPrematerial,mnuQueryDID,mnuQueryKP,mnuQueryWOGroup,mnuReturnDID,mnuReturnDIDALL,mnuSetDIOandInterlock,mnuSingleSideBrdConfirm,mnuStartSplitLineMC,mnuTransferDispatchDID,mnuTransferFujiNexim,mnuTransferFujiXML,mnuTransferPanaAMI,mnuTransferPanaMSF,mnuUnChkWO,mnuUpdRealQty,MnuUpLoadBom,MnuUpLoadMachineType,mnuUrgentDIDToWH,mnuUrgentInsertWO,mnuVendorCodeMapping,mnuVerifyFeederSlot,mnuWipReport,PowerCloseWO,RePrintDID,returnDIDflag,ReturnFlag,UploadReplacePN,XL_PMC,mnuSendXLRemainDemand><FACTORY=F7>";
                    strCommand = "<APPNAME=QSMS><LINE=All><STATION=QSMS><CONN=PROVIDER=SQLOLEDB;UID=sa;SERVER=10.94.7.11;DATABASE=SMT;NETWORK LIBRARY=DBMSSOCN;pwd=pqmb#7sa><USERID=QMS><BU=PU9><SERVERBU=PU9><CHKLOGIN=N><RIGHT=CheckBom,DeleteDID,DIDSlot,IPQCRetest,KeyInPNByManual,Login,mmuCompPNCompare,mmuDIDintegration,mmuQSMS_SapHis,mmuUnlockCompPNCompare,mmuWOInputPlan,mnuChangeExtraDIDslot,mnuCheckDispatchQty,mnuCloseWO,mnuCompPrint,mnuDefineBuildType,mnuDeleteDID,mnuDeleteFeeder,mnuDIDCallBack,mnuDIDChkStock,mnuDispatchDIDAdditionnal,mnuFixDispatchData,mnuForceDeleteDID,mnuGenXLMD,mnuInheritDIDByWO,MnuInSpection,mnuMaintainDID,mnuMaintainDIDAutoDispatch,mnumaintainFeeder,mnumaintainWOSeq,mnuMCCPreMaterial,mnuModfyDIDTotalQty,MnuPrematerial,mnuQueryDID,mnuQueryKP,mnuQueryWOGroup,mnuReturnDID,mnuReturnDIDALL,mnuSetDIOandInterlock,mnuSingleSideBrdConfirm,mnuStartSplitLineMC,mnuTransferDispatchDID,mnuTransferFujiNexim,mnuTransferFujiXML,mnuTransferPanaAMI,mnuTransferPanaMSF,mnuUnChkWO,mnuUpdRealQty,MnuUpLoadBom,MnuUpLoadMachineType,mnuUrgentDIDToWH,mnuUrgentInsertWO,mnuVendorCodeMapping,mnuVerifyFeederSlot,mnuWipReport,PowerCloseWO,RePrintDID,returnDIDflag,ReturnFlag,UploadReplacePN,XL_PMC,mnuSendXLRemainDemand><FACTORY=F7>";
                    //strCommand = "<APPNAME=QSMS><LINE=All><STATION=QSMS><CONN=PROVIDER=SQLOLEDB;UID=sa;SERVER=10.26.1.31;DATABASE=SMT;NETWORK LIBRARY=DBMSSOCN;pwd=pqms#9vd><USERID=A1082696><BU=PU8><SERVERBU=PU8><CHKLOGIN=N><RIGHT=mnuReturnDID,CheckBom,ClearMachine,CompCompare,CycleTime,DeleteDIDmmuCompPNCompare,mmuDIDintegration,mmuQSMS_SapHis,mmuUnlockCompPNCompare,mmuWOInputPlan,MnuAutoDispatch,mnuChangeExtraDIDslot,mnuCheckDispatchQty,mnuCloseWO,mnuCompPrint,mnuDefineBuildType,mnuWipReport,mnuInSpection,MnuUpLoadBom,mnumaintainWOSeq,mnumaintainFeeder,mnuTransferPanaAMI,mnuTransferPanaMSF,mnuTransferFujiNexim,mnuMaintainDIDAutoDispatch><FACTORY=F5>";
                }
                else
                {
                    MessageBox.Show("请使用MainMenu打开...", "打开错误");
                    return;
                }
            }

            for (int i = 0; i < args.Length; i++)
            {
                strCommand += args[i];
            }


            if (strCommand.IndexOf("<CONN=") > -1)
            {
                strTemp = strCommand.Substring(strCommand.IndexOf("<CONN="), strCommand.Length - strCommand.IndexOf("<CONN="));
                Parameter.ConnSMT = strTemp.Substring(strTemp.IndexOf("<CONN="), strTemp.Length).Substring("<CONN=".Length, strTemp.IndexOf(">") - "<CONN=".Length);
                Parameter.ConnSMT = Parameter.ConnSMT.Replace("PROVIDER=SQLOLEDB;", "");
            }

            if (strCommand.IndexOf("SERVER=") > -1)
            {
                strTemp = strCommand.Substring(strCommand.IndexOf("SERVER="), strCommand.Length - strCommand.IndexOf("SERVER="));
                Parameter.SMTServer = strTemp.Substring(strTemp.IndexOf("SERVER="), strTemp.Length).Substring("SERVER=".Length, strTemp.IndexOf(";") - "SERVER=".Length);
            }

            if (strCommand.IndexOf("<RIGHT=") > -1)
            {
                strTemp = strCommand.Substring(strCommand.IndexOf("<RIGHT="), strCommand.Length - strCommand.IndexOf("<RIGHT="));
                Parameter.strRights = strTemp.Substring("<RIGHT=".Length, strTemp.IndexOf(">") - "<RIGHT=".Length);
            }


            if (strCommand.IndexOf("STATION") > -1)
            {
                strTemp = strCommand.Substring(strCommand.IndexOf("<STATION="), strCommand.Length - strCommand.IndexOf("<STATION="));
                Parameter.strStation = strTemp.Substring("<STATION=".Length, strTemp.IndexOf(">") - "<STATION=".Length);

            }

            if (strCommand.IndexOf("USERID=") > -1)
            {
                strTemp = strCommand.Substring(strCommand.IndexOf("<USERID="), strCommand.Length - strCommand.IndexOf("<USERID="));
                Parameter.g_userName = strTemp.Substring("<USERID=".Length, strTemp.IndexOf(">") - "<USERID=".Length);
            }

            if (strCommand.IndexOf("SERVERBU=") > -1)
            {
                strTemp = strCommand.Substring(strCommand.IndexOf("SERVERBU="), strCommand.Length - strCommand.IndexOf("SERVERBU="));
                Parameter.BU = strTemp.Substring(strTemp.IndexOf("SERVERBU="), strTemp.Length).Substring("SERVERBU=".Length, strTemp.IndexOf(">") - "SERVERBU=".Length);
            }

            if (strCommand.IndexOf("LINE=") > -1)
            {
                strTemp = strCommand.Substring(strCommand.IndexOf("LINE="), strCommand.Length - strCommand.IndexOf("LINE="));
                Parameter.strLine = strTemp.Substring(strTemp.IndexOf("LINE="), strTemp.Length).Substring("LINE=".Length, strTemp.IndexOf(">") - "LINE=".Length);

            }
            if (strCommand.IndexOf("FACTORY=") > -1)
            {
                strTemp = strCommand.Substring(strCommand.IndexOf("FACTORY="), strCommand.Length - strCommand.IndexOf("FACTORY="));
                Parameter.g_factory = strTemp.Substring(strTemp.IndexOf("FACTORY="), strTemp.Length).Substring("FACTORY=".Length, strTemp.IndexOf(">") - "FACTORY=".Length);

            }

            if (strCommand.IndexOf("DATA SOURCE=") > -1)
            {
                strTemp = strCommand.Substring(strCommand.IndexOf("DATA SOURCE="), strCommand.Length - strCommand.IndexOf("DATA SOURCE="));
                Parameter.SMTDB = strTemp.Substring(strTemp.IndexOf("DATA SOURCE="), strTemp.Length).Substring("DATA SOURCE=".Length, strTemp.IndexOf(";") - "DATA SOURCE=".Length);
            }
            Parameter.g_userRight = Parameter.strRights.Split(new char[] { ',' });

            if (Parameter.SMTServer == "")
            {
                MessageBox.Show("Cant't get SMT Server information !! Call QMS please!");
                return;
            }
            else
            {
                QMSSDK.Db.Connections.CreateCn(Parameter.ConnSMT);
                //获取QSMSDB
                dt = login.GetQSMSServer(Parameter.SMTServer);
                Parameter.SMTDB = dt.Rows[0]["smt_db"].ToString();
                Parameter.QSMSDB = dt.Rows[0]["qsms_db"].ToString();
                Parameter.QSMSServer = dt.Rows[0]["QSMS_Server"].ToString();
            }

            if (Parameter.QSMSServer == "")
            {
                MessageBox.Show("Can't get QSMS Server information ! Call QMS please!");
                return;
            }
            else
            {
                Parameter.IP = Parameter.QSMSServer;
                Parameter.ConnQSMS = Parameter.ConnSMT.Replace(Parameter.SMTDB, Parameter.QSMSDB);
                Parameter.ConnQSMS = Parameter.ConnQSMS.Replace(Parameter.SMTServer, Parameter.QSMSServer);
            }

            QMSSDK.Db.Connections.CreateCn(Parameter.ConnQSMS);

            dt = login.GetQSMS_ProConfig();
            for(int i = 0;i< dt.Rows.Count ; i++)
            {
                Parameter.ConfigList.Add(dt.Rows[i]["Key"].ToString().Trim().ToUpper(), dt.Rows[i]["Value"].ToString().Trim().ToUpper());
                #region  取消以下逻辑，用Parameter.ConfigList字典替代；

                //switch(dt.Rows[i]["Key"].ToString().ToUpper())
                //{
                //    case "SCANCOMPPN":
                //        Parameter.ScanCompPN = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case "SCANMSD":
                //        Parameter.ScanMSD = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case "CHECKBOMLOGON":
                //        Parameter.CheckBomLogon = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case "CHECKRETURNFORBIDDENPN":
                //        Parameter.CheckReturnForbiddenPN = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHKOLDDIDLABELQTY") :
                //        Parameter.ChkOldDIDLabelQty = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHKONEBYONEMATERIAL") : 
                //        Parameter.ChkOneByOneMaterial = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("NPMMACHINETYPE")  :
                //        Parameter.NPMMachineType = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("MAINTAINFEEDERDID")  :
                //        Parameter.MaintainFeederDID = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHKFUJISPL")  :
                //        Parameter.ChkFujiSPL = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHKWOGROUPID")  :
                //        Parameter.ChkWOGroupID = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHKPRINTDIDTYPE"):
                //        Parameter.ChkPrintDIDType = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("PRINTEDSEQID"):       
                //        Parameter.PrintedSeqID = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("BATCHCONTROL"):      
                //        Parameter.BatchControl = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("UnChkCompPN"):       
                //        Parameter.UnChkCompPN = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHECKNEEDMSD"):
                //        Parameter.CheckNeedMSD = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHECKWOIFREDUCEXBOARD"):
                //        Parameter.CheckWOIFReduceXboard = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHECKMSDCALLBACK"):     
                //        Parameter.CheckMSDCallBack = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHECKBURNDID"):    
                //        Parameter.CheckBurnDID = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("NOKEEPPWD"):     
                //        Parameter.NoKeepPWD = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("BGAWAREHOUSE"):   
                //        Parameter.BGAWarehouse = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHKPNCQ"):
                //        Parameter.ChkPNCQ = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHECKBSMATERIAL"):
                //        Parameter.CheckBSMaterial = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHKEQPROGRAM"):
                //        Parameter.ChkEQProgram = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHKDATECODE"):
                //        Parameter.ChkDateCode = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHECKDIDBYLINE"):
                //        Parameter.strChkDIDByLine =dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("PRINTEDVENDERCODE"):     
                //        Parameter.PrintedVenderCode = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("NEWGROUPIDRULE"):        
                //        Parameter.NewGroupIDRule = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("UNCHKBASEREELQTY"):       
                //        Parameter.UnChkBaseReelQty = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("CHKMEBOM_LOCATION"):
                //        Parameter.ChkMEBOM_Location = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("DIDAUTOOPEN"):
                //        Parameter.DIDAutoOpen = dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //    case ("LABELPRINTCHECK"):   
                //        Parameter.LabelPrintCheck =dt.Rows[i]["Value"].ToString().ToUpper();
                //        break;
                //}
                #endregion  获取Excel中指定位置的值；
            }
            Parameter.chkQty = QMSSDK.Br.FileSystem.Ini.ReadIniValue("QSMS", "MaxDIDGroupQty", AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "set.ini");
            Parameter.StrBU = QMSSDK.Br.FileSystem.Ini.ReadIniValue("COMMON", "BU", AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "set.ini");
            if (Parameter.g_factory == "")
            {
                Parameter.CreateDIDFlag = "N";
                Parameter.Factory = Parameter.g_factory;
            }
            else
            {
                Parameter.CreateDIDFlag = "Y";
                Parameter.Factory = Parameter.g_factory;
            }
            
            Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new frmMain());
            Application.Run(new frmSendXLRemainDemand());       //superchai add 20231003
            
        }
    }
}
