using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace QSMS.DbLibrary.MCC
{
    public class MCCProcess : QMSSDK.Db.WinForm
    {
        string msg = "";

        public DataTable CheckUploadReplacePNRight()
        {
            string strSQL = "Select * from UserRight where UserName='" + Parameter.g_userName.Trim() + "' and UserRight='UploadReplacePN'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetDataByFuncType(string FuncType)
        {
            string strSQL = "";
            DataTable dt = new DataTable();
            bool Isnull = false;
            switch (FuncType.ToUpper())
            {
                case "QSMS_MEBOM":
                    strSQL = "select   * from QSMS_MEBom order by Machine,JObPN,Version,CompPN";
                    break;
                case "SINGLESIDEBRD":
                    strSQL = "select   * from QSMS_SingleSideBrd order by MBPN";
                    break;
                case "REPLACEPN":
                    strSQL = "select   * from QSMS_ReplacePN order by JobPN,Version,ID";
                    break;
                case "UNCHKCOMP":
                    strSQL = "select   * from QSMS_UnChkComp order by CompHead";
                    break;
                case "FUJIBRDSEQMAPPING":
                    strSQL = "select   * from FujiBrdSeqMapping";
                    break;
                case "TRAYSLOT":
                    strSQL = "Select   * from TraySlot order by machine";
                    break;
                case "CASTRATE":
                    strSQL = "select   * from QSMS_CastRate order by CompHead";
                    break;
                case "ONEBYONE":
                    strSQL = "select   * from QSMS_OneByOne order by CompPN";
                    break;
                case "DID":
                    strSQL = "select   * from QSMS_DID order by DID";
                    break;
                case "NOMACHINEDROPCOMPPN":
                    strSQL = "select   * from QSMS_UnCheckCompPN where Type='NOMDrop' order by CompPN";
                    break;
                case "COMPPNINSPECTRULE":
                    strSQL = "select   * from QSMS_InSpect_Rule";
                    break;
                case "XL_WOPLANSEQ":
                    strSQL = "select top(200) Date,Shift,Line,WO,PlanQty,SeqID,TransDateTime,OPID,InputQty,Factory from XL_WOPlanSeq order by Date Desc,shift,Line,SeqID";
                    break;
                case "XL_WOPLANLINE":
                    strSQL = "select   * from XL_WOPlanLine";
                    break;
                case "XL_IMPLEMENTPN":
                    strSQL = "select distinct PrefixPN,UID,TransDateTime from XL_ImplementPN";
                    break;
                case "MATERIALTOWHID":
                    strSQL = "select   * from MATERIALTOWHID";
                    break; ;
                case "XL_PNONEBYONE":
                    strSQL = "select   * from xl_PNOneByOne";
                    break;
                case "XL_DOUBLETABLES":
                    strSQL = "select   * from DoubleTables";
                    break;
                case "XL_MAXDIDMAINTAINQTY":
                    strSQL = "select * from XL_MaxDIDMaintainQty";
                    break;
                case "NOCHECKREPLACEPNSPLICING":
                    strSQL = "SELECT * FROM QSMS_NOCheckReplacePNSplicing";
                    break;
                case "COMPONENT_DATA":
                    strSQL = "select * from Component_data";
                    break;
                case "MACHINE_DATA":
                    strSQL = "select * from Machine_data";
                    break;
                case "COMPPN_SPACER":
                    strSQL = "select CompPN,Value,UID,TransdateTime from CompPN_BaseData WHERE TYPE='Spacer'";
                    break;
                default:
                    Isnull = true;
                    MessageBox.Show("Please check the right sheet name.");
                    break;
            }
            if (Isnull == true)
            {
                return dt;
            }
            else
            {
                dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
                return dt;
            }

        }

        public DataTable getDID(string barCode)
        {
            string strSQL = "QSMS_GenUNID";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@2DBarCode", SqlDbType.NVarChar) { Value = barCode };
            return SqlHelper.ExecuteDataTable(strSQL, paras, Parameter.ConnQSMS);
        }

        public DataTable CheckFormat(string Category, string Value)
        {
            string strSQL = "CheckFormat";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Category", SqlDbType.VarChar) { Value = Category };
            paras[1] = new SqlParameter("@Value", SqlDbType.VarChar) { Value = Value };
            return SqlHelper.ExecuteDataTable(strSQL, paras, Parameter.ConnQSMS);
        }

        public DataTable SaveUNID(string CompPN, string DateCode, string VendorCode, string LotCode, string Qty, string UNID, string BarCode)
        {
            string strSQL = "XL_DIDSaveDIDLabel_Data";
            SqlParameter[] paras = new SqlParameter[7];
            paras[0] = new SqlParameter("@CompPN", SqlDbType.VarChar) { Value = CompPN };
            paras[1] = new SqlParameter("@DateCode", SqlDbType.VarChar) { Value = DateCode };
            paras[2] = new SqlParameter("@VendorCode", SqlDbType.VarChar) { Value = VendorCode };
            paras[3] = new SqlParameter("@LotCode", SqlDbType.VarChar) { Value = LotCode };
            paras[4] = new SqlParameter("@Qty", SqlDbType.VarChar) { Value = Qty };
            paras[5] = new SqlParameter("@UniqueID", SqlDbType.VarChar) { Value = UNID };
            paras[6] = new SqlParameter("@2DBarCode", SqlDbType.VarChar) { Value = BarCode };
            return SqlHelper.ExecuteDataTable(strSQL, paras, Parameter.ConnQSMS);
        }

        //public DataTable CheckNeedDispatch(string Type, string CompPN)
        //{
        //    string strSQL = "XL_CheckNeedDispatch";
        //    SqlParameter[] paras = new SqlParameter[2];
        //    paras[0] = new SqlParameter("@Type", SqlDbType.VarChar) { Value = Type };
        //    paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar) { Value = CompPN };
        //    return SqlHelper.ExecuteDataTable(strSQL, paras, Parameter.ConnQSMS);
        //}

        public DataTable CheckNeedDispatch(string Type, string CompPN, string Factory)
        {
            string strSQL = "XL_CheckNeedDispatch";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar) { Value = Type };
            paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar) { Value = CompPN };
            paras[2] = new SqlParameter("@Factory", SqlDbType.VarChar) { Value = Factory };
            return SqlHelper.ExecuteDataTable(strSQL, paras, Parameter.ConnQSMS);
        }

        public DataTable XL_GetAllWOInfoList(string Type, string WO, string Machine, string Side, string Slot, string LR, string CompPN, string Line, string DID = "", string BeginDate = "", string EndDate = "", string GroupID = "")
        {
            string strSQL = "XL_GetAllWOInfoList";
            SqlParameter[] paras = new SqlParameter[12];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar) { Value = Type };
            paras[1] = new SqlParameter("@WO", SqlDbType.VarChar) { Value = WO };
            paras[2] = new SqlParameter("@Machine", SqlDbType.VarChar) { Value = Machine };
            paras[3] = new SqlParameter("@Side", SqlDbType.VarChar) { Value = Side };
            paras[4] = new SqlParameter("@Slot", SqlDbType.VarChar) { Value = Slot };
            paras[5] = new SqlParameter("@LR", SqlDbType.VarChar) { Value = LR };
            paras[6] = new SqlParameter("@CompPN", SqlDbType.VarChar) { Value = CompPN };
            paras[7] = new SqlParameter("@Line", SqlDbType.VarChar) { Value = Line };
            paras[8] = new SqlParameter("@DID", SqlDbType.VarChar) { Value = DID };
            paras[9] = new SqlParameter("@BeginDate", SqlDbType.VarChar) { Value = BeginDate };
            paras[10] = new SqlParameter("@EndDate", SqlDbType.VarChar) { Value = EndDate };
            paras[11] = new SqlParameter("@GroupID", SqlDbType.VarChar) { Value = GroupID };
            DataSet ds = SqlHelper.ExecuteDataSet(strSQL, paras, Parameter.ConnQSMS);
            return ds.Tables[0];
        }

        public DataSet XL_GetAllWOInfoList_WO(string Type, string WO, string Machine, string Side, string Slot, string LR, string CompPN, string Line)
        {
            string strSQL = "XL_GetAllWOInfoList";
            SqlParameter[] paras = new SqlParameter[8];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar) { Value = Type };
            paras[1] = new SqlParameter("@WO", SqlDbType.VarChar) { Value = WO };
            paras[2] = new SqlParameter("@Machine", SqlDbType.VarChar) { Value = Machine };
            paras[3] = new SqlParameter("@Side", SqlDbType.VarChar) { Value = Side };
            paras[4] = new SqlParameter("@Slot", SqlDbType.VarChar) { Value = Slot };
            paras[5] = new SqlParameter("@LR", SqlDbType.VarChar) { Value = LR };
            paras[6] = new SqlParameter("@CompPN", SqlDbType.VarChar) { Value = CompPN };
            paras[7] = new SqlParameter("@Line", SqlDbType.VarChar) { Value = Line };
            return SqlHelper.ExecuteDataSet(strSQL, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMSInsertDispatch(string WO, string GroupID, string Line, string WoQty, string JobPN, string Machine, string CompPN, string Slot, string LR, string BaseQty, string NeedQty, string DID, string TotalQty, string DIDQty, string VendorCode, string DateCode, string LotCode, string UID, string DIDDateTime, string InheritWO, string Item, string JobGroup, string Side)
        {
            string strSQL = "QSMSInsertDispatch";
            SqlParameter[] paras = new SqlParameter[23];
            paras[0] = new SqlParameter("@Work_Order", SqlDbType.VarChar) { Value = WO };
            paras[1] = new SqlParameter("@GroupID", SqlDbType.VarChar) { Value = GroupID };
            paras[2] = new SqlParameter("@Line", SqlDbType.VarChar) { Value = Line };
            paras[3] = new SqlParameter("@WoQty", SqlDbType.VarChar) { Value = WoQty };
            paras[4] = new SqlParameter("@JobPN", SqlDbType.VarChar) { Value = JobPN };
            paras[5] = new SqlParameter("@Machine", SqlDbType.VarChar) { Value = Machine };
            paras[6] = new SqlParameter("@CompPN", SqlDbType.VarChar) { Value = CompPN };
            paras[7] = new SqlParameter("@Slot", SqlDbType.VarChar) { Value = Slot };
            paras[8] = new SqlParameter("@LR", SqlDbType.VarChar) { Value = LR };
            paras[9] = new SqlParameter("@BaseQty", SqlDbType.VarChar) { Value = BaseQty };
            paras[10] = new SqlParameter("@NeedQty", SqlDbType.VarChar) { Value = NeedQty };
            paras[11] = new SqlParameter("@DID", SqlDbType.VarChar) { Value = DID };
            paras[12] = new SqlParameter("@TotalQty", SqlDbType.VarChar) { Value = TotalQty };
            paras[13] = new SqlParameter("@DIDQty", SqlDbType.VarChar) { Value = DIDQty };
            paras[14] = new SqlParameter("@VendorCode", SqlDbType.VarChar) { Value = VendorCode };
            paras[15] = new SqlParameter("@DateCode", SqlDbType.VarChar) { Value = DateCode };
            paras[16] = new SqlParameter("@LotCode", SqlDbType.VarChar) { Value = LotCode };
            paras[17] = new SqlParameter("@UID", SqlDbType.VarChar) { Value = UID };
            paras[18] = new SqlParameter("@DIDDateTime", SqlDbType.VarChar) { Value = DIDDateTime };
            paras[19] = new SqlParameter("@InheritWO", SqlDbType.VarChar) { Value = InheritWO };
            paras[20] = new SqlParameter("@Item", SqlDbType.VarChar) { Value = Item };
            paras[21] = new SqlParameter("@JobGroup", SqlDbType.VarChar) { Value = JobGroup };
            paras[22] = new SqlParameter("@Side", SqlDbType.VarChar) { Value = Side };
            DataSet ds = SqlHelper.ExecuteDataSet(strSQL, paras, Parameter.ConnQSMS);
            return ds.Tables[0];
        }

        public void RecordDispatchFDT(string WO, string Type = "")
        {
            string strSQL = "RecordDispatchFDT";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Work_Order", SqlDbType.VarChar) { Value = WO };
            paras[1] = new SqlParameter("@Type", SqlDbType.VarChar) { Value = Type };
            SqlHelper.Executeless(strSQL, CommandType.StoredProcedure, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_CheckDID(string DID, string strBarCode, string strCompPN, string strOPID, string Type)
        {
            string strSQL = "QSMS_CheckDID";
            SqlParameter[] paras = new SqlParameter[5];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar) { Value = DID };
            paras[1] = new SqlParameter("@BarCode", SqlDbType.VarChar) { Value = strBarCode };
            paras[2] = new SqlParameter("@CompPN", SqlDbType.VarChar) { Value = strCompPN };
            paras[3] = new SqlParameter("@Type", SqlDbType.VarChar) { Value = Type };
            paras[4] = new SqlParameter("@OPID", SqlDbType.VarChar) { Value = strOPID };
            return SqlHelper.ExecuteDataTable(strSQL, paras, Parameter.ConnQSMS);
        }

        public DataSet XL_Dispatch_MaterialPrompt(string strCompPN, string strVendorCode, string strDateCode, string strLotCode, string strFactory)
        {
            string strSQL = "XL_Dispatch_MaterialPrompt";
            SqlParameter[] paras = new SqlParameter[5];
            paras[0] = new SqlParameter("@CompPN", SqlDbType.VarChar) { Value = strCompPN };
            paras[1] = new SqlParameter("@VendorCode", SqlDbType.VarChar) { Value = strVendorCode };
            paras[2] = new SqlParameter("@DateCode", SqlDbType.VarChar) { Value = strDateCode };
            paras[3] = new SqlParameter("@LotCode", SqlDbType.VarChar) { Value = strLotCode };
            paras[4] = new SqlParameter("@Factory", SqlDbType.VarChar) { Value = strFactory };
            return SqlHelper.ExecuteDataSet(strSQL, paras, Parameter.ConnQSMS);
        }

        public DataTable GetPrintInfo(string DID, string Type)
        {
            string strSQL = "QSMS_GetDIDPrintInfo";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar) { Value = DID };
            paras[1] = new SqlParameter("@Type", SqlDbType.VarChar) { Value = Type };
            return SqlHelper.ExecuteDataTable(strSQL, paras, Parameter.ConnQSMS);
        }

        //public DataTable XL_DIDAutoDispatch(string DID, string strCompPN, string strQty, string strVendorCode, string strDateCode, string strLotCode, string DIDLoc, string strUID, string strDispatchType, string strWOGroup, string strWO, string strLine, string strSide, string strMachine, string strSlot, string strLR, string str09Code)
        //{
        //    string strSQL = "XL_DIDAutoDispatch";
        //    SqlParameter[] paras = new SqlParameter[19];
        //    paras[0] = new SqlParameter("@DID", SqlDbType.VarChar) { Value = DID };
        //    paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar) { Value = strCompPN };
        //    paras[2] = new SqlParameter("@Qty", SqlDbType.VarChar) { Value = strQty };
        //    paras[3] = new SqlParameter("@RemainQty", SqlDbType.VarChar) { Value = strQty };
        //    paras[4] = new SqlParameter("@VendorCode", SqlDbType.VarChar) { Value = strVendorCode };
        //    paras[5] = new SqlParameter("@DateCode", SqlDbType.VarChar) { Value = strDateCode };
        //    paras[6] = new SqlParameter("@LotCode", SqlDbType.VarChar) { Value = strLotCode };
        //    paras[7] = new SqlParameter("@UID", SqlDbType.VarChar) { Value = strUID };
        //    paras[8] = new SqlParameter("@Type", SqlDbType.VarChar) { Value = strDispatchType };
        //    paras[9] = new SqlParameter("@extraWOGroup", SqlDbType.VarChar) { Value = strWOGroup };
        //    paras[10] = new SqlParameter("@extraWO", SqlDbType.VarChar) { Value = strWO };
        //    paras[11] = new SqlParameter("@extraLine", SqlDbType.VarChar) { Value = strLine };
        //    paras[12] = new SqlParameter("@extraSide", SqlDbType.VarChar) { Value = strSide };
        //    paras[13] = new SqlParameter("@extraMachine", SqlDbType.VarChar) { Value = strMachine };
        //    paras[14] = new SqlParameter("@extraSlot", SqlDbType.VarChar) { Value = strSlot };
        //    paras[15] = new SqlParameter("@extraLR", SqlDbType.VarChar) { Value = strLR };
        //    paras[16] = new SqlParameter("@extraQty", SqlDbType.VarChar) { Value = strQty };
        //    paras[17] = new SqlParameter("@str09Code", SqlDbType.VarChar) { Value = str09Code };
        //    paras[18] = new SqlParameter("@DIDLoc", SqlDbType.VarChar) { Value = DIDLoc };

        //    return SqlHelper.ExecuteDataTable(strSQL, paras, Parameter.ConnQSMS);
        //}

        public DataTable XL_DIDAutoDispatch(string DID, string strCompPN, string strQty, string strVendorCode, string strDateCode, string strLotCode, string strUID, string strDispatchType, string strWOGroup, string strWO, string strLine, string strSide, string strMachine, string strSlot, string strLR, string str09Code, string Factory = "", string OldDID = "")
        {
            string strSQL = "XL_DIDAutoDispatch";
            SqlParameter[] paras = new SqlParameter[20];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar) { Value = DID };
            paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar) { Value = strCompPN };
            paras[2] = new SqlParameter("@Qty", SqlDbType.VarChar) { Value = strQty };
            paras[3] = new SqlParameter("@RemainQty", SqlDbType.VarChar) { Value = strQty };
            paras[4] = new SqlParameter("@VendorCode", SqlDbType.VarChar) { Value = strVendorCode };
            paras[5] = new SqlParameter("@DateCode", SqlDbType.VarChar) { Value = strDateCode };
            paras[6] = new SqlParameter("@LotCode", SqlDbType.VarChar) { Value = strLotCode };
            paras[7] = new SqlParameter("@UID", SqlDbType.VarChar) { Value = strUID };
            paras[8] = new SqlParameter("@Type", SqlDbType.VarChar) { Value = strDispatchType };
            paras[9] = new SqlParameter("@extraWOGroup", SqlDbType.VarChar) { Value = strWOGroup };
            paras[10] = new SqlParameter("@extraWO", SqlDbType.VarChar) { Value = strWO };
            paras[11] = new SqlParameter("@extraLine", SqlDbType.VarChar) { Value = strLine };
            paras[12] = new SqlParameter("@extraSide", SqlDbType.VarChar) { Value = strSide };
            paras[13] = new SqlParameter("@extraMachine", SqlDbType.VarChar) { Value = strMachine };
            paras[14] = new SqlParameter("@extraSlot", SqlDbType.VarChar) { Value = strSlot };
            paras[15] = new SqlParameter("@extraLR", SqlDbType.VarChar) { Value = strLR };
            paras[16] = new SqlParameter("@extraQty", SqlDbType.VarChar) { Value = strQty };
            paras[17] = new SqlParameter("@str09Code", SqlDbType.VarChar) { Value = str09Code };
            paras[18] = new SqlParameter("@Factory", SqlDbType.VarChar) { Value = Factory };
            paras[19] = new SqlParameter("@OldDID", SqlDbType.VarChar) { Value = OldDID };
            return SqlHelper.ExecuteDataTable(strSQL, paras, Parameter.ConnQSMS);
        }

        public DataTable Return_XL_DIDAutoDispatch(string DID, string strCompPN, string strQty, string RemainQty, string strVendorCode, string strDateCode, string strLotCode, string DIDLoc, string DIDMEM, string strUID, string strDispatchType, string Factory = "", string OldDID = "")
        {
            string strSQL = "XL_DIDAutoDispatch";
            SqlParameter[] paras = new SqlParameter[13];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar) { Value = DID };
            paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar) { Value = strCompPN };
            paras[2] = new SqlParameter("@Qty", SqlDbType.VarChar) { Value = strQty };
            paras[3] = new SqlParameter("@RemainQty", SqlDbType.VarChar) { Value = strQty };
            paras[4] = new SqlParameter("@VendorCode", SqlDbType.VarChar) { Value = strVendorCode };
            paras[5] = new SqlParameter("@DateCode", SqlDbType.VarChar) { Value = strDateCode };
            paras[6] = new SqlParameter("@LotCode", SqlDbType.VarChar) { Value = strLotCode };
            paras[7] = new SqlParameter("@DIDLoc", SqlDbType.VarChar) { Value = strLotCode };
            paras[8] = new SqlParameter("@DIDMEM", SqlDbType.VarChar) { Value = strLotCode };
            paras[9] = new SqlParameter("@UID", SqlDbType.VarChar) { Value = strUID };
            paras[10] = new SqlParameter("@Type", SqlDbType.VarChar) { Value = strDispatchType };
            paras[11] = new SqlParameter("@Factory", SqlDbType.VarChar) { Value = Factory };
            paras[12] = new SqlParameter("@OldDID", SqlDbType.VarChar) { Value = OldDID };
            return SqlHelper.ExecuteDataTable(strSQL, paras, Parameter.ConnQSMS);
        }

        public bool CheckBCMS(string strCompPN)
        {
            string strSQL = "Select Top 1 0 From BCMS_Bios Where ChipPN like '%" + strCompPN + "%'";
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnSMT);
            if (dt.Rows.Count > 0)
            {
                strSQL = "Select Top 1 0 From UserRight Where AppName='QSMS' and Userright='CheckBiosLogin' and UserName='" + Parameter.g_userName.Trim() + "'";
                dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnSMT);
                if (dt.Rows.Count > 0)
                {
                    return true;
                }
                return false;
            }
            return true;
        }

        public bool CheckCompPN(string strCompPN, string type = "")
        {
            string strSQL = "";
            strSQL = "Select Top 1 0 From OneToOneControl Where CompPN like '%" + strCompPN + "%'";
            if (type == "IsNeedMSD")
            {
                strSQL = "Select Top 1 0 From MSD_Data Where CompPN = '" + strCompPN + "'";
            }
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            if (dt.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }
        //Paul Add
        public DataTable QSMS_MCC_QueryDataByType(string Type, string BeginDate, string EndDate, string Item1, string Item2, string Item3)
        {
            this.spName = "QSMS_PD_QueryDataByType";
            SqlParameter[] paras = new SqlParameter[6];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = Type };
            paras[1] = new SqlParameter("@BeginDate", SqlDbType.VarChar, 14) { Value = BeginDate };
            paras[2] = new SqlParameter("@EndDate", SqlDbType.VarChar, 14) { Value = EndDate };
            paras[3] = new SqlParameter("@Item1", SqlDbType.VarChar, 200) { Value = Item1 };
            paras[4] = new SqlParameter("@Item2", SqlDbType.VarChar, 200) { Value = Item2 };
            paras[5] = new SqlParameter("@Item3", SqlDbType.VarChar, 200) { Value = Item3 };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_QueryVendorCode(string CompPN, string type)//0002 Ada
        {
            this.spName = "QSMS_PD_QueryVendorCode";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@CompPN", SqlDbType.VarChar, 50) { Value = CompPN };
            paras[1] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = type };
            //msg = "";
            //DataTable dt = SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
            //msg = dt.Rows[0]["Des"].ToString();
            //if (dt.Rows[0]["Result"].ToString().ToUpper().Trim() != "P")
            //{
            //    return false;
            //}

            //return true;
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);

        }
        //Paul Add
        public DataTable QSMS_PrintDID(string UNID)
        {
            this.spName = "QSMS_PrintDID";//10001
            //this.spName = "_QSMS_PrintDID_F2023";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 100) { Value = UNID };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        //Paul Add

        public DataTable QSMS_GenUNID(string BarCode, string Type)
        {
            this.spName = "QSMS_GenUNID";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@2DBarCode", SqlDbType.NVarChar, 4000) { Value = BarCode };
            paras[1] = new SqlParameter("@Type", SqlDbType.VarChar, 20) { Value = Type };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        //Paul Add
        //10001 begin
        public DataTable QSMS_SaveCompPrintLog(string Type, string CompPN, string Qty, string VendorCode, string DateCode, string LotCode, string OPID, string Mark, string UNID, string PrintQty, string UniqueID, string Spec, string MfrSite)       
        {
            this.spName = "QSMS_SaveCompPrintLog";

            SqlParameter[] paras = new SqlParameter[13];

            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = Type };
            paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar, 30) { Value = CompPN };
            paras[2] = new SqlParameter("@Qty", SqlDbType.VarChar, 10) { Value = Qty };
            paras[3] = new SqlParameter("@VendorCode", SqlDbType.VarChar, 50) { Value = VendorCode };
            paras[4] = new SqlParameter("@DateCode", SqlDbType.VarChar, 50) { Value = DateCode };
            paras[5] = new SqlParameter("@LotCode", SqlDbType.VarChar, 50) { Value = LotCode };
            paras[6] = new SqlParameter("@OPID", SqlDbType.VarChar, 20) { Value = OPID };
            paras[7] = new SqlParameter("@Mark", SqlDbType.VarChar, 50) { Value = Mark };
            paras[8] = new SqlParameter("@UNID", SqlDbType.VarChar, 100) { Value = UNID };
            paras[9] = new SqlParameter("@PrintQty", SqlDbType.VarChar, 10) { Value = PrintQty };
            paras[10] = new SqlParameter("@UniqueID", SqlDbType.VarChar, 50) { Value = UniqueID };
            paras[11] = new SqlParameter("@Spec", SqlDbType.VarChar, 100) { Value = Spec };
            paras[12] = new SqlParameter("@MfrSite", SqlDbType.VarChar, 100) { Value = MfrSite };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_SaveCompPrintLog(string Type, string CompPN, string Qty, string VendorCode, string DateCode, string LotCode, string OPID, string Mark, string UNID, string PrintQty, string UniqueID, string Spec, string MfrSite, string strLinkID)
        {
            this.spName = "QSMS_SaveCompPrintLog";
            //this.spName = "_QSMS_SaveCompPrintLog_F2023";

            SqlParameter[] paras = new SqlParameter[14];

            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = Type };
            paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar, 30) { Value = CompPN };
            paras[2] = new SqlParameter("@Qty", SqlDbType.VarChar, 10) { Value = Qty };
            paras[3] = new SqlParameter("@VendorCode", SqlDbType.VarChar, 50) { Value = VendorCode };
            paras[4] = new SqlParameter("@DateCode", SqlDbType.VarChar, 50) { Value = DateCode };
            paras[5] = new SqlParameter("@LotCode", SqlDbType.VarChar, 50) { Value = LotCode };
            paras[6] = new SqlParameter("@OPID", SqlDbType.VarChar, 20) { Value = OPID };
            paras[7] = new SqlParameter("@Mark", SqlDbType.VarChar, 50) { Value = Mark };
            paras[8] = new SqlParameter("@UNID", SqlDbType.VarChar, 100) { Value = UNID };
            paras[9] = new SqlParameter("@PrintQty", SqlDbType.VarChar, 10) { Value = PrintQty };
            paras[10] = new SqlParameter("@UniqueID", SqlDbType.VarChar, 50) { Value = UniqueID };
            paras[11] = new SqlParameter("@Spec", SqlDbType.VarChar, 100) { Value = Spec };
            paras[12] = new SqlParameter("@MfrSite", SqlDbType.VarChar, 100) { Value = MfrSite };
            paras[13] = new SqlParameter("@LinkID", SqlDbType.VarChar, 50) { Value = strLinkID };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        //10001 end
        public DataTable QSMS_CompPN(string CompPN)
        {
            this.spName = "QSMS_CompPN";
            SqlParameter[] paras = new SqlParameter[1];


            paras[0] = new SqlParameter("@CompPN", SqlDbType.VarChar, 30) { Value = CompPN };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable CheckFormat(string PartNumber)
        {
            string strSQL = "Exec CheckFormat 'PARTNUMBER','" + PartNumber + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable DeleteDIDByBU(string BU, string DID)
        {
            string strSQL = "Exec DeleteDIDByBU '" + BU + "','" + Parameter.g_userName + "','" + DID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GenRegisterDID(string UNID, string CompPN, int Qty, string VendorCode, string DateCode, string LotCode, string Inspection, string DIDMEM, string OPID, string TransDate, int Type, string MSD)
        {
            this.spName = "MCC_GenRegisterDID";
            SqlParameter[] paras = new SqlParameter[12];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = UNID };
            paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar, 50) { Value = CompPN };
            paras[2] = new SqlParameter("@Qty", SqlDbType.Int) { Value = Qty };
            paras[3] = new SqlParameter("@VendorCode", SqlDbType.VarChar, 50) { Value = VendorCode };
            paras[4] = new SqlParameter("@DateCode", SqlDbType.VarChar, 50) { Value = DateCode };
            paras[5] = new SqlParameter("@LotCode", SqlDbType.VarChar, 50) { Value = LotCode };
            paras[6] = new SqlParameter("@DIDLoc", SqlDbType.VarChar, 50) { Value = Inspection };
            paras[7] = new SqlParameter("@DIDMEM", SqlDbType.VarChar, 50) { Value = DIDMEM };
            paras[8] = new SqlParameter("@UID", SqlDbType.VarChar, 20) { Value = OPID };
            paras[9] = new SqlParameter("@TransDateTime", SqlDbType.VarChar, 14) { Value = TransDate };
            paras[10] = new SqlParameter("@Type", SqlDbType.Int) { Value = Type };
            paras[11] = new SqlParameter("@MSD", SqlDbType.VarChar, 30) { Value = MSD };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public void QSMSGroupCompQty(string GroupID)
        {
            this.spName = "QSMSGroupCompQty";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@GroupID", SqlDbType.VarChar, 50) { Value = GroupID };
            SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable XL_DIDGetToWHInfo(string Type, string DID, string Factory, string IsAnotherBUDID)
        {
            this.spName = "XL_DIDGetToWHInfo";
            SqlParameter[] paras = new SqlParameter[4];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = Type };
            paras[1] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            paras[2] = new SqlParameter("@Factory", SqlDbType.VarChar, 50) { Value = Factory };
            paras[3] = new SqlParameter("@IsAnotherBUDID", SqlDbType.VarChar, 50) { Value = IsAnotherBUDID };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_WONeedReturnDID(string WO)
        {
            this.spName = "QSMS_WONeedReturnDID";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@WO", SqlDbType.VarChar, 50) { Value = WO };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_GetDIDRealQty(string DID)
        {
            this.spName = "GetDIDRealQty";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_DIDPrintTypeCheck(string DID, string ChkOldLabel, string Factory)
        {
            this.spName = "DIDPrintTypeCheck";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            paras[1] = new SqlParameter("@ChkOldLabel", SqlDbType.VarChar, 50) { Value = ChkOldLabel };
            paras[2] = new SqlParameter("@Factory", SqlDbType.VarChar, 50) { Value = Factory };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_XL_DIDReturnCheck(string DID, string ChkPN)
        {
            this.spName = "XL_DIDReturnCheck";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 100) { Value = DID };
            paras[1] = new SqlParameter("@ChkPN", SqlDbType.VarChar, 50) { Value = ChkPN };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_XL_ReturnDIDByGroupID(string GroupID)
        {
            this.spName = "XL_ReturnDIDByGroupID";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@GroupID", SqlDbType.VarChar, 100) { Value = GroupID };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_XL_ReturnDIDByWO(string GroupID)
        {
            this.spName = "XL_ReturnDIDByWO";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@GroupID", SqlDbType.VarChar, 100) { Value = GroupID };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public void QSMS_MCC_QSMSGetCastQty(string GroupID)
        {
            this.spName = "QSMSGetCastQty";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@GroupID", SqlDbType.VarChar, 100) { Value = GroupID };

            SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_XL_DIDGetRefID(string Type, string IsGood, string UserName, string Factory, string IsAnotherBUDID)
        {
            this.spName = "XL_DIDGetRefID";
            SqlParameter[] paras = new SqlParameter[5];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@IsGood", SqlDbType.VarChar, 2) { Value = IsGood };
            paras[2] = new SqlParameter("@UserName", SqlDbType.VarChar, 30) { Value = UserName };
            paras[3] = new SqlParameter("@Factory", SqlDbType.VarChar, 30) { Value = Factory };
            paras[4] = new SqlParameter("@IsAnotherBUDID", SqlDbType.VarChar, 30) { Value = IsAnotherBUDID };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_PD_MSD_LinkDIDAuto(string DID, string ReturnDID, string CompPN, string Inherit_WO, string ReturnFlag, string g_userName)
        {
            this.spName = "PD_MSD_LinkDIDAuto";
            SqlParameter[] paras = new SqlParameter[6];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 30) { Value = DID };
            paras[1] = new SqlParameter("@ReturnDID", SqlDbType.VarChar, 2) { Value = ReturnDID };
            paras[2] = new SqlParameter("@CompPN", SqlDbType.VarChar, 30) { Value = CompPN };
            paras[3] = new SqlParameter("@Inherit_WO", SqlDbType.VarChar, 30) { Value = Inherit_WO };
            paras[4] = new SqlParameter("@ReturnFlag", SqlDbType.VarChar, 30) { Value = ReturnFlag };
            paras[5] = new SqlParameter("@UID", SqlDbType.VarChar, 30) { Value = g_userName };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_XL_CheckReturnQty(string DID, string CompPN, int ReturnQty, string GroupID, string IsAnotherBUDID, string CheckForbiddenPN)
        {
            this.spName = "XL_CheckReturnQty";
            SqlParameter[] paras = new SqlParameter[6];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar, 20) { Value = CompPN };
            paras[2] = new SqlParameter("@ReturnQty", SqlDbType.Int) { Value = ReturnQty };
            paras[3] = new SqlParameter("@GroupID", SqlDbType.VarChar, 50) { Value = GroupID };
            paras[4] = new SqlParameter("@IsAnotherBUDID", SqlDbType.VarChar, 20) { Value = IsAnotherBUDID };
            paras[5] = new SqlParameter("@CheckForbiddenPN", SqlDbType.VarChar, 5) { Value = CheckForbiddenPN };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_QSMS_ReturnDID(string ReturnDID, string DID, string CompPN, int ReturnQty, string UID, string GroupID, string Transdatetime, string IsGood, string PrtCallBKandReturn, string Factory, string IsAnotherBUDID)
        {
            this.spName = "QSMS_ReturnDID";
            SqlParameter[] paras = new SqlParameter[11];
            paras[0] = new SqlParameter("@ReturnDID", SqlDbType.VarChar, 50) { Value = ReturnDID };
            paras[1] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            paras[2] = new SqlParameter("@CompPN", SqlDbType.VarChar, 20) { Value = CompPN };
            paras[3] = new SqlParameter("@ReturnQty", SqlDbType.Int) { Value = ReturnQty };
            paras[4] = new SqlParameter("@UID", SqlDbType.VarChar, 20) { Value = UID };
            paras[5] = new SqlParameter("@GroupID", SqlDbType.VarChar, 50) { Value = GroupID };
            paras[6] = new SqlParameter("@Transdatetime", SqlDbType.VarChar, 20) { Value = Transdatetime };
            paras[7] = new SqlParameter("@IsGood", SqlDbType.VarChar, 20) { Value = GroupID };
            paras[8] = new SqlParameter("@PrtCallBKandReturn", SqlDbType.VarChar, 20) { Value = PrtCallBKandReturn };
            paras[9] = new SqlParameter("@Factory", SqlDbType.VarChar, 20) { Value = Factory };
            paras[10] = new SqlParameter("@IsAnotherBUDID", SqlDbType.VarChar, 20) { Value = IsAnotherBUDID };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataSet QSMS_MCC_XL_DIDGetNewID(string Type, string DID, string IsGood, int ReturnQty, string UserName, string Factory, string IsAnotherBUDID)
        {
            this.spName = "XL_DIDGetNewID";
            SqlParameter[] paras = new SqlParameter[7];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 20) { Value = Type };
            paras[1] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            paras[2] = new SqlParameter("@IsGood", SqlDbType.Char, 1) { Value = IsGood };
            paras[3] = new SqlParameter("@ReturnQty", SqlDbType.Int) { Value = ReturnQty };
            paras[4] = new SqlParameter("@UserName", SqlDbType.VarChar, 20) { Value = UserName };
            paras[5] = new SqlParameter("@Factory", SqlDbType.VarChar, 20) { Value = Factory };
            paras[6] = new SqlParameter("@IsAnotherBUDID", SqlDbType.VarChar, 20) { Value = IsAnotherBUDID };

            return SqlHelper.ExecuteDataSet(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataSet QSMS_MCC_XL_GetDidPrintInfo_Return(string DID, string OldDID, string IsAnotherBU = "N", string Factory = "", string PrinterType = "", string PrintDpm = "")
        {
            this.spName = "XL_GetDidPrintInfo_Return";
            SqlParameter[] paras = new SqlParameter[6];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            paras[1] = new SqlParameter("@OldDID", SqlDbType.VarChar, 50) { Value = OldDID };
            paras[2] = new SqlParameter("@IsAnotherBU", SqlDbType.VarChar, 20) { Value = IsAnotherBU };
            paras[3] = new SqlParameter("@Factory", SqlDbType.VarChar, 20) { Value = Factory };
            paras[4] = new SqlParameter("@PrinterType", SqlDbType.VarChar, 20) { Value = PrinterType };
            paras[5] = new SqlParameter("@PrintDpm", SqlDbType.VarChar, 20) { Value = PrintDpm };

            return SqlHelper.ExecuteDataSet(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_XL_DIDChk_ReturnDisp(string NewDID, string OldDID, string ErrMsg, string IsAnotherBUDID = "N")
        {
            this.spName = "XL_DIDChk_ReturnDisp";
            SqlParameter[] paras = new SqlParameter[4];
            paras[0] = new SqlParameter("@NewDID", SqlDbType.VarChar, 50) { Value = NewDID };
            paras[1] = new SqlParameter("@OldDID", SqlDbType.VarChar, 50) { Value = OldDID };
            paras[2] = new SqlParameter("@ErrMsg", SqlDbType.VarChar, 2000) { Value = ErrMsg };
            paras[3] = new SqlParameter("@IsAnotherBUDID", SqlDbType.VarChar, 20) { Value = IsAnotherBUDID };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataSet QSMS_MCC_XL_ReturnComp(string CompPN, string VendorCode, string DateCode, string LotCode, string IsGood, int Qty, string UserName, string Factory, string CheckForbiddenPN, string OldDID)
        {
            this.spName = "XL_ReturnComp";
            SqlParameter[] paras = new SqlParameter[10];
            paras[0] = new SqlParameter("@CompPN", SqlDbType.VarChar, 30) { Value = CompPN };
            paras[1] = new SqlParameter("@VendorCode", SqlDbType.VarChar, 50) { Value = VendorCode };
            paras[2] = new SqlParameter("@DateCode", SqlDbType.VarChar, 50) { Value = DateCode };
            paras[3] = new SqlParameter("@LotCode", SqlDbType.VarChar, 50) { Value = LotCode };
            paras[4] = new SqlParameter("@IsGood", SqlDbType.VarChar, 5) { Value = IsGood };
            paras[5] = new SqlParameter("@Qty", SqlDbType.Int) { Value = Qty };
            paras[6] = new SqlParameter("@UserName", SqlDbType.VarChar, 20) { Value = UserName };
            paras[7] = new SqlParameter("@Factory", SqlDbType.VarChar, 20) { Value = Factory };
            paras[8] = new SqlParameter("@CheckForbiddenPN", SqlDbType.VarChar, 5) { Value = CheckForbiddenPN };
            paras[9] = new SqlParameter("@OldDID", SqlDbType.VarChar, 5) { Value = OldDID };

            return SqlHelper.ExecuteDataSet(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataSet QSMS_MCC_XL_ChkAnotherBUDID(string DID, string IsAnotherBUDID, string Factory) //001
        {
            this.spName = "XL_ChkAnotherBUDID";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 100) { Value = DID };
            paras[1] = new SqlParameter("@IsAnotherBUDID", SqlDbType.VarChar, 20) { Value = IsAnotherBUDID };
            paras[2] = new SqlParameter("@Factory", SqlDbType.VarChar, 20) { Value = Factory };
            return SqlHelper.ExecuteDataSet(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable CheckMachine(string Line, string Machine, string Side)
        {
            string strSQL = "select machine from machine where machine='" + Machine + "' and line ='" + Line + "' and side='" + Side + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable XL_ReturnCompRefresh(string Factory)
        {
            string strSQL = "exec XL_ReturnCompRefresh '" + Factory + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetComponent_Data(string CompPN)
        {
            string strSQL = "select 1 from Component_Data where comppn='" + CompPN + "' and [Functype]='BSMaterial'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetQSMS_DID_ToWH(string CompPN)
        {
            string strSQL = "select CompPN,VendorCode,DateCode ,LotCode ,Qty from QSMS_DID_ToWH where DID = '" + CompPN + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable XL_DIDGetRefID(string optGoodMaterial, string UserName, string Factory)
        {
            string strSQL = "exec XL_DIDGetRefID @Type='Return',@IsGood='" + optGoodMaterial + "',@UserName='" + UserName + "',@Factory='" + Factory + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataSet XL_DIDChkStockByRefID_set(string sCurrRefID, string UserName)
        {
            string strSQL = "exec XL_DIDChkStockByRefID @Type='Auto',@RefID='" + sCurrRefID + "',@UserName='" + UserName + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }

        public DataTable XL_DIDChkStockByRefID(string sCurrRefID, string UserName)
        {
            string strSQL = "exec XL_DIDChkStockByRefID @Type='Auto',@RefID='" + sCurrRefID + "',@UserName='" + UserName + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public void Del_QSMSMEBom(string Machine, string JobPN, string JobGroup, string Version, string BuildType, string Factory, string Line)
        {
            string strSQL = "delete QSMS_MEBom where JobGroup='" + JobGroup + "' and Machine='" + Machine + "' and JobPN='" + JobPN + "' and Version='" + Version + "' and BuildType='" + BuildType + "' and Factory='" + Factory + "' and Line='" + Line + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public void Insert_QSMSMEBom(string Machine, string JobPN, string JobGroup, string Version, string CompPN, string LR, string Slot, string Qty, string BuildType, string Side, string UID, string Factory, string Line)
        {
            string strSQL = "Insert Into QSMS_MEBom(Machine,JobPN,JobGroup,Version,CompPN,LR,Slot,Qty,BuildType,Side,UID,Factory,Line) values('" + Machine + "','" + JobPN + "','" + JobGroup + "','" + Version + "','" + CompPN + "','" + LR + "','" + Slot + "','" + Qty + "','" + BuildType + "','" + Side + "','" + UID + "','" + Factory + "','" + Line + "')";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public void Save_Log(string UID, string Name)
        {
            string strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_PanaMSF','" + Name + "','" + UID + "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable QSMS_DIDIntegration(string Item, string CompPN, string NEWDID, string Line, string Side, string Vendorcode, string Datecode, string Lotcode, string Factory, int Qty, int RemainQty, string ALLDID, string BeginTime, string EndTime, string UID, string GroupID)
        {
            this.spName = "QSMS_DIDIntegration";
            SqlParameter[] paras = new SqlParameter[16];
            paras[0] = new SqlParameter("@Item", SqlDbType.VarChar, 20) { Value = Item };
            paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar, 50) { Value = CompPN };
            paras[2] = new SqlParameter("@NEWDID", SqlDbType.VarChar, 50) { Value = NEWDID };
            paras[3] = new SqlParameter("@Line", SqlDbType.VarChar, 20) { Value = Line };
            paras[4] = new SqlParameter("@Side", SqlDbType.VarChar, 5) { Value = Side };
            paras[5] = new SqlParameter("@Vendorcode", SqlDbType.VarChar, 20) { Value = Vendorcode };
            paras[6] = new SqlParameter("@Datecode", SqlDbType.VarChar, 20) { Value = Datecode };
            paras[7] = new SqlParameter("@Lotcode", SqlDbType.VarChar, 20) { Value = Lotcode };
            paras[8] = new SqlParameter("@Factory", SqlDbType.VarChar, 8) { Value = Factory };
            paras[9] = new SqlParameter("@Qty", SqlDbType.Int) { Value = Qty };
            paras[10] = new SqlParameter("@RemainQty", SqlDbType.Int) { Value = RemainQty };
            paras[11] = new SqlParameter("@ALLDID", SqlDbType.VarChar, 300) { Value = ALLDID };
            paras[12] = new SqlParameter("@BeginTime", SqlDbType.VarChar, 14) { Value = BeginTime };
            paras[13] = new SqlParameter("@EndTime", SqlDbType.VarChar, 14) { Value = EndTime };
            paras[14] = new SqlParameter("@UID", SqlDbType.VarChar, 20) { Value = UID };
            paras[15] = new SqlParameter("@GroupID", SqlDbType.VarChar, 50) { Value = GroupID };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable XL_GetDidPrintInfo(string DID, string Factory)
        {
            string strSQL = "XL_GetDidPrintInfo";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            paras[1] = new SqlParameter("@Factory", SqlDbType.VarChar, 50) { Value = Factory };
            return SqlHelper.ExecuteDataTable(strSQL, paras, Parameter.ConnQSMS);
        }

        public DataTable MCC_QueryDataByType(string Type, string BeginDate, string EndDate, string Item1, string Item2, string Item3, string Item4, string Item5, string Item6, string Item7)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + Type + "','','','" + Item1 + "','" + Item2 + "','" + Item3 + "','" + Item4 + "','" + Item5 + "','" + Item6 + "','" + Item7 + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public void InsertIntoQSMS_MEBom(string Machine, string JobPN, string JobGroup, string Version, string CompPN, string LR, string Slot, string Qty, string BuildType, string Side, string UID, string Factory, string Line, string Location)
        {
            string strSQL = "insert into QSMS_MEBom(Machine,JobPN,JobGroup,Version,CompPN,LR,Slot,Qty,BuildType,Side,UID,Factory,Line,Location)" +
               " values('" + Machine + "','" + JobPN + "','" + JobGroup + "','" + Version + "','" + CompPN + "','" + LR + "','" + Slot + "','" + Qty + "','" + BuildType + "','" + Side + "','" + UID + "','" + Factory + "','" + Line + "','" + Location + "')";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataSet QSMS_ReturnQry(string Type, string ReturnDID)
        {
            this.spName = "QSMS_ReturnQry";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 20) { Value = Type };
            paras[1] = new SqlParameter("@ReturnDID", SqlDbType.VarChar, 50) { Value = ReturnDID };
            return SqlHelper.ExecuteDataSet(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_InsertMEBom_Nexim(string Factory, string Line, string JobGroup, string Revision, string BuildType, string Side, string Username)
        {
            this.spName = "QSMS_InsertMEBom_Nexim";
            SqlParameter[] paras = new SqlParameter[7];
            paras[0] = new SqlParameter("@Factory", SqlDbType.VarChar, 10) { Value = Factory };
            paras[1] = new SqlParameter("@Line", SqlDbType.VarChar, 10) { Value = Line };
            paras[2] = new SqlParameter("@JobPN", SqlDbType.VarChar, 50) { Value = JobGroup };
            paras[3] = new SqlParameter("@Version", SqlDbType.VarChar, 10) { Value = Revision };
            paras[4] = new SqlParameter("@BuildType", SqlDbType.VarChar, 5) { Value = BuildType };
            paras[5] = new SqlParameter("@Side", SqlDbType.VarChar, 5) { Value = Side };
            paras[6] = new SqlParameter("@UID", SqlDbType.VarChar, 20) { Value = Username };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_InsertMEBom_Nexim_MI(string Factory, string Line, string JobGroup, string Revision, string BuildType, string Side, string Type, string Username)
        {
            this.spName = "QSMS_InsertMEBom_Nexim_MI";
            SqlParameter[] paras = new SqlParameter[8];
            paras[0] = new SqlParameter("@Factory", SqlDbType.VarChar, 10) { Value = Factory };
            paras[1] = new SqlParameter("@Line", SqlDbType.VarChar, 10) { Value = Line };
            paras[2] = new SqlParameter("@JobPN", SqlDbType.VarChar, 50) { Value = JobGroup };
            paras[3] = new SqlParameter("@Version", SqlDbType.VarChar, 10) { Value = Revision };
            paras[4] = new SqlParameter("@BuildType", SqlDbType.VarChar, 5) { Value = BuildType };
            paras[5] = new SqlParameter("@Side", SqlDbType.VarChar, 5) { Value = Side };
            paras[6] = new SqlParameter("@Type", SqlDbType.VarChar, 5) { Value = Type };
            paras[7] = new SqlParameter("@UID", SqlDbType.VarChar, 20) { Value = Username };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public void QSMS_MCC_QSMSDIDCallBack(string WO) //002
        {
            this.spName = "QSMSDIDCallBack";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@WO", SqlDbType.VarChar, 50) { Value = WO };

            SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_GetWOPCBStatus(string WO) //002
        {
            this.spName = "GetWOPCBStatus";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@WO", SqlDbType.VarChar, 30) { Value = WO };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_DIDSimilarDispByPCB(string WO, string DID) //002
        {
            this.spName = "DIDSimilarDispByPCB";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@WO", SqlDbType.VarChar, 30) { Value = WO };
            paras[1] = new SqlParameter("@DID", SqlDbType.VarChar, 100) { Value = DID };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataSet QSMS_MCC_DIDChkAnotherBU(string DID, string IsAnotherBUDID, string Factory) //002
        {
            this.spName = "DIDChkAnotherBU";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 100) { Value = DID };
            paras[1] = new SqlParameter("@IsAnotherBUDID", SqlDbType.VarChar, 100) { Value = IsAnotherBUDID };
            paras[2] = new SqlParameter("@Factory", SqlDbType.VarChar, 100) { Value = Factory };
            return SqlHelper.ExecuteDataSet(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_DIDRestoreForCallBK(string WO, string DID) //002
        {
            this.spName = "DIDRestoreForCallBK";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@WO", SqlDbType.VarChar, 20) { Value = WO };
            paras[1] = new SqlParameter("@DID", SqlDbType.VarChar, 100) { Value = DID };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataSet QSMS_MCC_DIDInfoForCallBK(string WO, string DID) //002
        {
            this.spName = "DIDInfoForCallBK";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@WO", SqlDbType.VarChar, 20) { Value = WO };
            paras[1] = new SqlParameter("@DID", SqlDbType.VarChar, 100) { Value = DID };

            return SqlHelper.ExecuteDataSet(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_MCC_DIDCallBackByType(string CallType, string WO, string CompPN, string DID, int ReturnQty, string UserName, string IsGood, string IsAnotherBUDID) //002
        {
            this.spName = "DIDCallBackByType";
            SqlParameter[] paras = new SqlParameter[8];
            paras[0] = new SqlParameter("@CallType", SqlDbType.VarChar, 20) { Value = CallType };
            paras[1] = new SqlParameter("@WO", SqlDbType.VarChar, 200) { Value = WO };
            paras[2] = new SqlParameter("@CompPN", SqlDbType.VarChar, 30) { Value = CompPN };
            paras[3] = new SqlParameter("@DID", SqlDbType.VarChar, 200) { Value = DID };
            paras[4] = new SqlParameter("@ReturnQty", SqlDbType.Int) { Value = ReturnQty };
            paras[5] = new SqlParameter("@UserName", SqlDbType.VarChar, 20) { Value = UserName };
            paras[6] = new SqlParameter("@IsGood", SqlDbType.VarChar, 10) { Value = IsGood };
            paras[7] = new SqlParameter("@IsAnotherBUDID", SqlDbType.VarChar, 20) { Value = IsAnotherBUDID };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public bool CheckReturnRight()
        {
            string strSQL = "Select Top 1 0 From UserRight Where AppName='QSMS' and Userright='ReturnCompPN' and UserName='" + Parameter.g_userName.Trim() + "'";
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnSMT);
            if (dt.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }

        public DataTable GetLine() //FrmDefineBuildType 
        {
            string strSQL = "select distinct Line from Sap_WO_List with(nolock) where Trans_Date>dbo.FormatDate(getdate()-180,'YYYYMMDDHHNNSS')";
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnSMT);
            return dt;
        }

        public DataTable GetDataGDetail() //FrmDefineBuildType 
        {
            string strSQL = "select top 500 WO,PN,MB_Rev,BuildType,Line,Qty,WO_Type,CostBU,Trans_Date from Sap_Wo_List where BuildType<>'1' order by Trans_Date desc";
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnSMT);
            return dt;
        }

        public DataTable GetDataWOMulti() //FrmDefineBuildType 
        {
            string strSQL = "select top 500 * from WO_MultiLine order by TransDateTime desc";
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnSMT);
            return dt;
        }

        public DataTable GetWOByLine(string line) //FrmDefineBuildType 
        {
            string strSQL = "select WO from Sap_Wo_List where Line='" + line + "' and InitAOIFlag='Y' and Trans_Date>dbo.FormatDate(getdate()-60,'YYYYMMDDHHNNSS') and QCCnt <=0 order by Trans_Date DESC";
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnSMT);
            return dt;
        }

        public DataTable GetWoinfoBasic(string WO) //FrmDefineBuildType 
        {
            string strSQL = "select PN, Qty ,MB_Rev,WO_Type,Line,[Group],BuildType from Sap_Wo_List where WO='" + WO + "'";
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnSMT);
            return dt;
        }

        public DataTable CheckWODispatch(string Group) //FrmDefineBuildType 
        {
            string strSQL = "Select distinct Work_Order from QSMS_Dispatch with(nolock) where Work_Order in(Select Wo from Sap_WO_List where [Group]='" + Group + "'";
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnSMT);
            return dt;

        }
        //Parameter.g_userName.Trim()
        public DataTable QSMS_SetBuildType(string WO, string buidType, string line, string side, string station) //FrmDefineBuildType 
        {
            //@WO varchar(20),@BuildType varchar(20),@Line varchar(20)='', @Side varchar(20) ='',@UserName varchar(10)='',@Station varchar(50)='' AS
            this.spName = "QSMS_SetBuildType";
            SqlParameter[] paras = new SqlParameter[6];
            paras[0] = new SqlParameter("@WO", SqlDbType.VarChar, 20) { Value = WO };
            paras[1] = new SqlParameter("@BuildType", SqlDbType.VarChar, 200) { Value = buidType };
            paras[2] = new SqlParameter("@Line", SqlDbType.VarChar, 30) { Value = line };
            paras[3] = new SqlParameter("@Side", SqlDbType.VarChar, 200) { Value = side };
            paras[4] = new SqlParameter("@UserName", SqlDbType.Int) { Value = Parameter.g_userName.Trim() };
            paras[5] = new SqlParameter("@Station", SqlDbType.VarChar, 20) { Value = station };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public void DeleteQSMS_WO(string Group) //FrmDefineBuildType 
        {
            string strSQL = "Delete QSMS_WO where Work_Order in(Select Wo from Sap_WO_List where [Group]='" + Group + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnSMT);

        }

        public DataTable GetWoinfoByGroup(string Group) //FrmDefineBuildType 
        {
            string strSQL = "select WO,BuildType from Sap_Wo_List where Group='" + Group + "'";
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnSMT);
            return dt;
        }

        public bool CheckBom(string WO, string BuildType) //FrmDefineBuildType 
        {
            //@WO varchar(20), @RefreshFlag varchar(20) = 'N',@BuildType varchar(5) = '1'

            this.spName = "QSMS_CheckBomSP";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@WO", SqlDbType.VarChar, 20) { Value = WO };
            paras[1] = new SqlParameter("@RefreshFlag", SqlDbType.VarChar, 20) { Value = "N" };
            paras[2] = new SqlParameter("@BuildType", SqlDbType.VarChar, 200) { Value = BuildType };
            if (SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS).Rows.Count > 0)
            {
                return false;
            }
            return true;
        }

        public DataTable CheckSapBomFail(string Group) //FrmDefineBuildType 
        {
            string strSQL = "select *  from Sap_BOM_Fail  where Work_Order in (select wo from sap_wo_list where [group]='" + Group + "')";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnSMT);
        }

        public DataTable Load()
        {
            string strSQL = "select distinct PN from ModelName";
            DataTable dt = this.Execute(strSQL, null);
            return dt;
        }

        public DataTable DID(string DID)
        {
            string strSQL = "select distinct CompPN from QSMS_DID where DID= '" + DID + "'";
            DataTable dt = this.Execute(strSQL, null);
            return dt;
        }

        public DataTable Model(string PN)
        {
            string strSQL = "select distinct ModelName from ModelName where PN= '" + PN + "'";
            DataTable dt = this.Execute(strSQL, null);
            return dt;
        }

        public DataTable ShearPinLinkDID(string DID, string Model, string PN, string CompPN, string UID)
        {
            this.spName = "IC_ShearPinLinkDID ";
            SqlParameter[] paras = new SqlParameter[5];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            paras[1] = new SqlParameter("@Model", SqlDbType.VarChar, 20) { Value = Model };
            paras[2] = new SqlParameter("@PN", SqlDbType.VarChar, 30) { Value = PN };
            paras[3] = new SqlParameter("@CompPN", SqlDbType.VarChar, 30) { Value = CompPN };
            paras[4] = new SqlParameter("@UID", SqlDbType.VarChar, 20) { Value = UID };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable ccshearpin(string PN, string Model)
        {
            string strSQL = "select * from IC_ShearPin where PN like '" + PN + "%' and ModelName like '" + Model + "%'";
            DataTable dt = this.Execute(strSQL, null);
            return dt;
        }

        public DataTable QSMS_QueryDIDData(string CompPN, string BeginDID, string EndDID)
        {
            if (BeginDID == "" || EndDID == "")
            {
                string strSQL = "select top 150 DID,CompPN,VendorCode,DateCode,LotCode,Qty,RemainQty,UID,TransDateTime,Line,Side,FirstMachine from QSMS_DID where CompPN like '" + CompPN + "%' and Qty<>0 order by TransDateTime desc";
                DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
                return dt;
            }
            else
            {
                string strSQL = "select TOP 150 DID,CompPN,VendorCode,DateCode,LotCode,Qty,RemainQty,UID,TransDateTime,Line,Side,FirstMachine from QSMS_DID where did between '" + BeginDID + "' and '" + EndDID + "' and DIDHostName=left(Host_Name(),20) and Qty<>0 order by TransDateTime desc";
                DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
                return dt;
            }
        }

        public DataTable QSMS_ProcessCompBatch(string CompPN, string DID, string Batch, string UID, string Type)
        {
            this.spName = "QSMS_ProcessCompBatch";
            SqlParameter[] paras = new SqlParameter[5];
            paras[0] = new SqlParameter("@CompPN", SqlDbType.VarChar, 50) { Value = CompPN };
            paras[1] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            paras[2] = new SqlParameter("@Batch", SqlDbType.VarChar, 50) { Value = Batch };
            paras[3] = new SqlParameter("@UID", SqlDbType.VarChar, 50) { Value = UID };
            paras[4] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = Type };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_DIDAutoOpen(string DID)
        {
            this.spName = "QSMS_DIDAutoOpen";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public bool QSMS_ChkEMMC(string MBPN, string ImageVersion)
        {
            string strSQL = "select 0 from EMMC with(nolock) where MBPN='" + MBPN + "' and ImageVersion= '" + ImageVersion + "";
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            if (dt.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }

        public DataTable QSMS_ChkDateCodeSpecial(string Vendor, string CompPN, string DateCode)
        {
            this.spName = "QSMS_ChkDateCodeSpecial";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@Vendor", SqlDbType.VarChar, 50) { Value = Vendor };
            paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar, 50) { Value = CompPN };
            paras[2] = new SqlParameter("@DateCode", SqlDbType.VarChar, 20) { Value = DateCode };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public bool QSMS_CheckVendorPN(string CompPN, string VendorPN)
        {
            string strSQL = "select 0 from Mapping_VendorPN with(nolock) where QuantaPN='" + CompPN + "' and VendorPN='" + VendorPN + "'";
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            if (dt.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }

        public DataTable IC_CompNeedBurn(string CompPN)
        {
            this.spName = "IC_CompNeedBurn";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@CompPN", SqlDbType.VarChar, 50) { Value = CompPN };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable DIDTrace_SaveData(string DID, string Type)
        {
            this.spName = "DIDTrace_SaveData";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            paras[1] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = Type };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public bool ChkEMMC(string CompPN)
        {
            string strSQL = "select 0 from EMMC with(nolock) where EMMCPN='" + CompPN + "'";
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            if (dt.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }

        public bool QSMS_ChkFWImage(string CompPN, string Type)
        {
            string strSQL = "select 0 from CompPN_Data with(nolock) where Type='" + Type + "' and CompPN='" + CompPN + "' ";
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            if (dt.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }

        public string QSMS_GetFWImage(string WO)
        {
            string strSQL = "select WO, Line, PN, ModelName, Item, Value, IsFirst, Chk_ID, Chk_Name, Chk_Result, Chk_Detail, Chk_Time, UID, TransDateTime  from WO_LinkItemData with(nolock) where WO='" + WO + "' and item='FW_Image' ";

            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnSMT);
            if (dt.Rows.Count > 0)
            {
                return dt.Rows[0]["Value"].ToString();
            }
            return "";
        }

        public DataTable QSMS_SaveFWImage(string DID, string CompPN, string WO, string Item, string Value, string UID)
        {
            this.spName = "QSMS_SaveFWImage";
            SqlParameter[] paras = new SqlParameter[6];

            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar) { Value = DID };
            paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar) { Value = CompPN };
            paras[2] = new SqlParameter("@WO", SqlDbType.VarChar) { Value = WO };
            paras[3] = new SqlParameter("@Item", SqlDbType.VarChar) { Value = Item };
            paras[4] = new SqlParameter("@Value", SqlDbType.VarChar) { Value = Value };
            paras[5] = new SqlParameter("@UID", SqlDbType.VarChar) { Value = UID };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable CheckVendorCode(string CompPN, string VendorCode)
        {
            this.spName = "QSMS_VendorCode_Check";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@CompPN", SqlDbType.VarChar, 50) { Value = CompPN };
            paras[1] = new SqlParameter("@VendorCode", SqlDbType.VarChar, 50) { Value = VendorCode };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public bool CheckVendorPN(string CompPN)
        {
            string strSQL = "select 0 from Mapping_VendorPN with(nolock) where QuantaPN='" + CompPN + "' ";
            DataTable dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            if (dt.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }

        public DataTable Check2DCode(string CompPN)
        {
            this.spName = "Check2DCode";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@2DBarCode", SqlDbType.VarChar, 200) { Value = CompPN };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable CompPrint_GetLinkID(string strLinkID)//10001
        {
            this.spName = "QSMS_CompPrint_GetLinkID";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@LinkID", SqlDbType.VarChar,50) { Value = strLinkID };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnSMT);
        }
    }
}
