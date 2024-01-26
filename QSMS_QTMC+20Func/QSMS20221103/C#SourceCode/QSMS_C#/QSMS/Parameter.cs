using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace QSMS
{
    class Parameter
    {
        public static Dictionary<string, string> ConfigList = new Dictionary<string, string>();         
        public static string Version = "20201118";
        public static string SMTServer;
        public static string SMTDB, QSMSDB;
        public static string QSMSServer;
        public static string strLine, strStation,ConnSMT,ConnQSMS, strRights;
        public static string IP;
        public static string Factory;
        public static string LocalIP;
        public static string StrBU;
        public static string UID;

        public static bool returnDIDflag;
        public static string chkDomain;
        public static string g_userName, g_factory,BU;
        public static string[] g_userRight;

        public static List<string> Openforms = new List<string>();
        public struct Settings_DataType
        {
            public string PRNa_Port;
            public string PRNa_Settings;
            public string LabelAFile;
            public string LabelSATOFIle;
            public string AutoDispatchLabel;
            public string AutoDispatchSatoLabel;
            public string ChkDIDDispatch;
            public string UpdateJobSide;
            public string DIDLabelGood;
            public string DIDLabelBad;
            public string DIDLabelSATOGood;
            public string DIDLabelSATOBad;
            public string CompPrintLabel;
            public string CompPNLabelPrint;
            public string KFLabel;
            public string AutoDispatchNewLabel;
            public string AutoDispatchSatoNewLabel;
            public string DIDLabelPath;
        }
        public static Settings_DataType Settings = new Settings_DataType();
        public struct DIDBasic
        {
            public string compPN;
            public string DID;
            public string VendorCode;
            public string DateCode;
            public string LotCode;
            public long Qty;
            public string IsGood;
            public string DIDType;
            public string location;
            public string Mark;
            public string WareHouseID;
            public string jobgroup;
        }
        public static DIDBasic DIDInfo = new DIDBasic();

        public static bool strKeyInPNByManual;
        public static bool CheckBomRight;
        public static bool DeleteMeBomByLine;

        #region 取消下面单独的Flag变量，改为ConfigList
        //public static string ScanCompPN;
        //public static string ScanMSD;
        //public static string CheckBomLogon;
        //public static string CheckReturnForbiddenPN;
        //public static string ChkOldDIDLabelQty;        
        //public static string ChkOneByOneMaterial;
        //public static string NPMMachineType;
        //public static string MaintainFeederDID;
        //public static string ChkFujiSPL;
        //public static string ChkWOGroupID;
        //public static string ChkPrintDIDType;
        //public static string PrintedSeqID;
        //public static string BatchControl;
        //public static string UnChkCompPN;
        //public static string CheckNeedMSD;
        //public static string CheckWOIFReduceXboard;
        //public static string CheckMSDCallBack;
        //public static string CheckBurnDID;
        //public static string NoKeepPWD;
        //public static string BGAWarehouse;
        //public static string ChkPNCQ;
        //public static string CheckBSMaterial;
        //public static string ChkEQProgram;
        //public static string ChkDateCode;
        //public static string strChkDIDByLine;
        //public static string PrintedVenderCode;
        //public static string NewGroupIDRule;
        //public static string UnChkBaseReelQty;
        //public static string ChkMEBOM_Location;
        //public static string DIDAutoOpen;
        //public static string LabelPrintCheck;
        #endregion

        public static string CreateDIDFlag;
        public static string chkQty;
        public static string WorkDir;
        public static string Profile;
        public static string hSECTION;

        public static string Check_NonAVL;
        public static string Check_AVL;
        public static string IPDefine;

        public static string Check_DID;
        public static string DIDHead;
        public static string AutoDispatchForAnotherBU;
        public static string CheckPNGroup;
        public static string BUDIDShow;
        public static string DIDnotToQWMS;
        public static string imagePath;
        public static string TestFilepath;
        public static string IPQC_ChkVendorPN;
        public static string IC_CompChk;
        public static string PrtCallBKandReturn;
        public static string CheckDIDRemainQty;
        public static string CheckOldNewPrintType;
    }
}
