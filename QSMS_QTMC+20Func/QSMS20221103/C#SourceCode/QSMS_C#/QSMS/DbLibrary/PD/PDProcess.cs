using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace QSMS.DbLibrary.PD
{
    public class PDProcess : QMSSDK.Db.WinForm
    {
        public DataTable QSMS_PD_QueryDataByType(string Type, string BeginDate, string EndDate, string Item1, string Item2, string Item3)
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
            //return this.Execute(paras).Tables[0];
        }

        public DataTable QSMS_CHeckDID_NB5(string DID,string BarCode, string Type, string BeginDate, string EndDate, string Line, string GroupID, string Side, string OPID, string Tabel)
        {
            this.spName = "QSMS_CHeckDID";
            SqlParameter[] paras = new SqlParameter[10];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar, 500) { Value = DID };
            paras[1] = new SqlParameter("@BarCode", SqlDbType.VarChar, 500) { Value = BarCode };
            paras[2] = new SqlParameter("@Type", SqlDbType.VarChar, 10) { Value = Type };
            paras[3] = new SqlParameter("@BeginDate", SqlDbType.VarChar, 50) { Value = BeginDate };
            paras[4] = new SqlParameter("@EndDate", SqlDbType.VarChar, 50) { Value = EndDate };
            paras[5] = new SqlParameter("@Line", SqlDbType.VarChar, 20) { Value = Line };
            paras[6] = new SqlParameter("@GroupID", SqlDbType.VarChar, 30) { Value = GroupID };
            paras[7] = new SqlParameter("@Side", SqlDbType.VarChar, 10) { Value = Side };
            paras[8] = new SqlParameter("@OPID", SqlDbType.VarChar, 20) { Value = OPID };
            paras[9] = new SqlParameter("@Tabel", SqlDbType.VarChar, 20) { Value = Tabel };
            //return this.Execute(paras).Tables[0];
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_CHeckDID(string DID, string BarCode, string Type, string BeginDate, string EndDate, string Line)
        {
            string strSQL = "exec QSMS_CheckDID '" + DID + "','" + BarCode + "','" + Type + "','" + BeginDate + "','" + EndDate + "','" + Line + "','" + Parameter.g_userName + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public void PD_QSMS_DelWO(string WO)
        {
            this.spName = "QSMS_DelWO";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 20) { Value = WO };
            SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMSChkCloseWOByManual(string WO, string Dispatch_Flag, string AOI_Flag, string SAP1_Flag, string SAP2_Flag)
        {
            string strSQL = "exec QSMSChkCloseWOByManual '" + WO + "','" + Dispatch_Flag + "','" + AOI_Flag + "','" + SAP1_Flag + "','" + SAP2_Flag + "','" + Parameter.g_userName + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable QSMS_WONeedReturnDID(string WO)
        {
            this.spName = "QSMS_WONeedReturnDID";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@WO", SqlDbType.VarChar, 30) { Value = WO };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public void QSMSCloseWODelDID(string WO)
        {
            string strSQL = "exec QSMSCloseWODelDID '" + WO + "','" + Parameter.g_userName + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable QSMS_CloseWO_CheckWOIFReduceXboard(string WO)
        {
            string strSQL = "exec QSMS_CloseWO_CheckWOIFReduceXboard '" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable QSMS_CloseWO_ReduceXboard(string WO,int XBoardQty)
        {
            string strSQL = "exec QSMS_CloseWO_ReduceXboard '" + WO + "'," + XBoardQty + ",'" + Parameter.g_userName + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable QSMS_SapCostPacking(string WO, string CloseType)
        {
            string strSQL = "exec QSMS_SapCostPacking '" + WO + "','" + Parameter.g_userName + "','" + CloseType + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable QSMS_GenUNID(string BarCode, string Type)
        {
            this.spName = "QSMS_GenUNID";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@2DBarCode", SqlDbType.NVarChar, 4000) { Value = BarCode };
            paras[1] = new SqlParameter("@Type", SqlDbType.VarChar, 20) { Value = Type };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable QSMS_DIDBake(string DID, string UID, string Type)
        {
            string strSQL = "exec QSMS_DIDBake '" + DID + "','" + UID + "','" + Type + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetBaseDIDInfo(string DID)
        {
            string strSQL = "Select DID,Qty,RealQty,Line,Side from QSMS_DID where DID='" + DID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void UpdateRealQty(string updateToQty, string DID)
        {
            string strSQL = "Update QSMS_DID set RealQty='" + updateToQty + "' where DID='" + DID + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void InsertLog(string str,string DID,string str2)
        {
            //string strSQL = "Insert into QSMS_Log(System_Name, Event_No, DID, User_Name, ReturnQty, Trans_Date) values('SMT_QSMS', 'UpdateRealQty', '" + str + "', '" + Parameter.g_userName + "', '0', '" + DateTime.Now.ToString("yyyyMMddhhmmss") + "')";
            string strSQL = "Insert into QSMS_Log(System_Name, Event_No, DID, User_Name, ReturnQty, Trans_Date) values('SMT_QSMS',  '" + str + "', '" + DID + "', '" + Parameter.g_userName + str2 + "', '0', '" + DateTime.Now.ToString("yyyyMMddhhmmss") + "')";   //20230116 Ellen

            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GeiMachineData(string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable GeiLine(string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[1];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable GetGroupIDByLine(string BeginDate, string EndDate, string Line, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[4];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@BeginDate", SqlDbType.VarChar, 8) { Value = BeginDate };
            paras[2] = new SqlParameter("@EndDate", SqlDbType.VarChar, 8) { Value = EndDate };
            paras[3] = new SqlParameter("@Line", SqlDbType.VarChar, 10) { Value = Line };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable GetWoByGroupID(string GroupID, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@GroupID", SqlDbType.VarChar, 50) { Value = GroupID };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable GetGroup(string WO, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@WO", SqlDbType.VarChar, 30) { Value = WO };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable GetSBWO(string WO, string Group, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@WO", SqlDbType.VarChar, 30) { Value = WO };
            paras[2] = new SqlParameter("@GroupID", SqlDbType.VarChar, 30) { Value = Group };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable GetMachine(string Group, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@GroupID", SqlDbType.VarChar, 30) { Value = Group };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable GetJobByMachine(string Machine, string Group, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@GroupID", SqlDbType.VarChar, 30) { Value = Group };
            paras[2] = new SqlParameter("@Machine", SqlDbType.VarChar, 30) { Value = Machine };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataSet MachineFeeder(string JobGroup, string Machine, string WO, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[4];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@GroupID", SqlDbType.VarChar, 30) { Value = JobGroup };
            paras[2] = new SqlParameter("@Machine", SqlDbType.VarChar, 30) { Value = Machine };
            paras[3] = new SqlParameter("@WO", SqlDbType.VarChar, 30) { Value = WO };
            return SqlHelper.ExecuteDataSet(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable CopyToExcel(string JobGroup, string Machine, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@GroupID", SqlDbType.VarChar, 30) { Value = JobGroup };
            paras[2] = new SqlParameter("@Machine", SqlDbType.VarChar, 30) { Value = Machine };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable QueryDID(string DID, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable GenUNID(string DID, string Type)
        {
            this.spName = "QSMS_GenUNID";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@2DBarCode", SqlDbType.VarChar, 500) { Value = DID };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable ChkNonAVL(string DID, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@DID", SqlDbType.VarChar, 500) { Value = DID };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable CheckDIDValidity(string DID, string Machine, string GroupID, string WO, string Line, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[6];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@DID", SqlDbType.VarChar, 500) { Value = DID };
            paras[2] = new SqlParameter("@Machine", SqlDbType.VarChar, 30) { Value = Machine };
            paras[3] = new SqlParameter("@GroupID", SqlDbType.VarChar, 30) { Value = GroupID };
            paras[4] = new SqlParameter("@Line", SqlDbType.VarChar, 30) { Value = Line };
            paras[5] = new SqlParameter("@WO", SqlDbType.VarChar, 30) { Value = WO };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable ChkIfInCurretnFeeder(string Feeder, string DID, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@DID", SqlDbType.VarChar, 500) { Value = DID };
            paras[2] = new SqlParameter("@Feeder", SqlDbType.VarChar, 30) { Value = Feeder };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable CheckFeeder(string DID, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@DID", SqlDbType.VarChar, 500) { Value = DID };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable CheckFeederLine(string Feeder, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@Feeder", SqlDbType.VarChar, 50) { Value = Feeder };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable ChkNonAVLData(string Customer, string CompPN, string MBPN, string Model, string VendorCode, string DateCode, string LotCode, string WO, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[9];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@Customer", SqlDbType.VarChar, 50) { Value = Customer };
            paras[2] = new SqlParameter("@CompPN", SqlDbType.VarChar, 50) { Value = CompPN };
            paras[3] = new SqlParameter("@MBPN", SqlDbType.VarChar, 50) { Value = MBPN };
            paras[4] = new SqlParameter("@Model", SqlDbType.VarChar, 50) { Value = Model };
            paras[5] = new SqlParameter("@VendorCode", SqlDbType.VarChar, 50) { Value = VendorCode };
            paras[6] = new SqlParameter("@DateCode", SqlDbType.VarChar, 50) { Value = DateCode };
            paras[7] = new SqlParameter("@LotCode", SqlDbType.VarChar, 50) { Value = LotCode };
            paras[8] = new SqlParameter("@WO", SqlDbType.VarChar, 50) { Value = WO };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable QueryWOInfo(string JobGroup, string CompPN, string Machine, string Group, int Qty, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[6];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@JobGroup", SqlDbType.VarChar, 50) { Value = JobGroup };
            paras[2] = new SqlParameter("@CompPN", SqlDbType.VarChar, 50) { Value = CompPN };
            paras[3] = new SqlParameter("@Machine", SqlDbType.VarChar, 50) { Value = Machine };
            paras[4] = new SqlParameter("@GroupID", SqlDbType.VarChar, 50) { Value = Group };
            paras[5] = new SqlParameter("@Qty", SqlDbType.Int) { Value = Qty };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable QueryWO(string CompPN, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar, 50) { Value = CompPN };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable QueryProConfig(string Feeder, string Line, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@Line", SqlDbType.VarChar, 50) { Value = Line };
            paras[2] = new SqlParameter("@Feeder", SqlDbType.VarChar, 50) { Value = Feeder };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable SevenFeederData(string Machine, string JobGroup, string DID, string CompPN, string VendorCode, string DateCode, string LotCode, string Feeder, string Slot, string LR, string UserName, bool Panal, string Line, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[14];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@Machine", SqlDbType.VarChar, 50) { Value = Machine };
            paras[2] = new SqlParameter("@JobGroup", SqlDbType.VarChar, 50) { Value = JobGroup };
            paras[3] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            paras[4] = new SqlParameter("@CompPN", SqlDbType.VarChar, 50) { Value = CompPN };
            paras[5] = new SqlParameter("@VendorCode", SqlDbType.VarChar, 50) { Value = VendorCode };
            paras[6] = new SqlParameter("@DateCode", SqlDbType.VarChar, 50) { Value = DateCode };
            paras[7] = new SqlParameter("@LotCode", SqlDbType.VarChar, 50) { Value = LotCode };
            paras[8] = new SqlParameter("@Feeder", SqlDbType.VarChar, 50) { Value = Feeder };
            paras[9] = new SqlParameter("@Slot", SqlDbType.VarChar, 50) { Value = Slot };
            paras[10] = new SqlParameter("@LR", SqlDbType.VarChar, 50) { Value = LR };
            paras[11] = new SqlParameter("@UserName", SqlDbType.VarChar, 50) { Value = UserName };
            paras[12] = new SqlParameter("@Panal", SqlDbType.VarChar, 50) { Value = Panal };
            paras[13] = new SqlParameter("@Line", SqlDbType.VarChar, 50) { Value = Line };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable SevenLOG(string Line, string DID, string Machine, string Feeder, string userName, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[6];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@Line", SqlDbType.VarChar, 50) { Value = Line };
            paras[2] = new SqlParameter("@DID", SqlDbType.VarChar, 50) { Value = DID };
            paras[3] = new SqlParameter("@Machine", SqlDbType.VarChar, 50) { Value = Machine };
            paras[4] = new SqlParameter("@Feeder", SqlDbType.VarChar, 50) { Value = Feeder };
            paras[5] = new SqlParameter("@userName", SqlDbType.VarChar, 50) { Value = userName };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable AVL_Vendor(string Customer, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@Customer", SqlDbType.VarChar, 50) { Value = Customer };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable QueryAVL(string CompPN, string VendorCode, string Customer, string Model, int Qty, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[6];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar, 50) { Value = CompPN };
            paras[2] = new SqlParameter("@VendorCode", SqlDbType.VarChar, 50) { Value = VendorCode };
            paras[3] = new SqlParameter("@Customer", SqlDbType.VarChar, 50) { Value = Customer };
            paras[4] = new SqlParameter("@Model", SqlDbType.VarChar, 50) { Value = Model };
            paras[5] = new SqlParameter("@Qty", SqlDbType.Int) { Value = Qty };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public DataTable QueryControlPart(string CompPN, string Model, string Type)
        {
            this.spName = "QSMS_MaintainFeeder";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 30) { Value = Type };
            paras[1] = new SqlParameter("@CompPN", SqlDbType.VarChar, 50) { Value = CompPN };
            paras[2] = new SqlParameter("@Model", SqlDbType.VarChar, 50) { Value = Model };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QSMS_EXE(string SQL)  //2021/09/01 Rain 
        {
            string strSQL = SQL;
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
    }
}
