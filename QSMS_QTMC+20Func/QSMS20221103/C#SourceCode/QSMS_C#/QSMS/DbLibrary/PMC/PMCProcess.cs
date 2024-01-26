using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace QSMS.DbLibrary.PMC
{
    public class PMCProcess : QMSSDK.Db.WinForm
    {
        public DataTable GetDataByWO(string type, string wo, string line)
        {
            if (type == "UnClosedByWO")
            {
                string strSQL = "select a.GroupID,a.Seq_NO,a.Work_Order,a.Line,b.PN,B.MB_Rev,B.Qty,a.Wo_TransDateTime,a.Group_TransDateTime,a.DispatchFlag,a.Sap1Flag,a.ClosedFlag,a.ClosedType,a.UID,a.CloseDateTime from QSMS_WOGroup a with(nolock),Sap_Wo_List b with(nolock) where a.work_order=b.wo and a.groupID in (select Groupid from qsms_woGroup where work_order='" + wo + "') and ClosedFlag='N' and a.line like '" + line + "%' order by groupID";
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            else
            {
                string strSQL = "select a.GroupID,a.Seq_NO,a.Work_Order,a.Line,b.PN,B.MB_Rev,B.Qty,a.Wo_TransDateTime,a.Group_TransDateTime,a.DispatchFlag,a.Sap1Flag,a.ClosedFlag,a.ClosedType,a.UID,a.CloseDateTime from QSMS_WOGroup a with(nolock),Sap_Wo_List b with(nolock) where a.work_order=b.wo and a.groupID in (select Groupid from QSMS_woGroup where work_order='" + wo + "') and ClosedFlag='Y' and a.line like '" + line + "%' order by GroupID";
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
        }
        public DataTable GetData(string type, string sDate, string eDate, string line)
        {
            if (type == "UnClosedWO")
            {
                string strSQL = "select a.GroupID,a.Seq_NO,a.Work_Order,a.Line,b.PN,B.MB_Rev,B.Qty,a.Wo_TransDateTime,a.Group_TransDateTime,a.DispatchFlag,a.Sap1Flag,a.ClosedFlag,a.ClosedType,a.UID,a.CloseDateTime from qsms_WOGroup a with(nolock),Sap_Wo_List b  where a.Work_Order=b.Wo and substring(Group_TransDateTime,1,8) between '" + sDate + "' and '" + eDate + "' and ClosedFlag='N' and a.line like '" + line + "%' order by groupID";
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            else
            {
                string strSQL = "select a.GroupID,a.Seq_NO,a.Work_Order,a.Line,b.PN,B.MB_Rev,B.Qty,a.Wo_TransDateTime,a.Group_TransDateTime,a.DispatchFlag,a.Sap1Flag,a.ClosedFlag,a.ClosedType,a.UID,a.CloseDateTime from qsms_WOGroup a with(nolock),Sap_Wo_List b  where a.Work_Order=b.Wo and substring(Group_TransDateTime,1,8) between '" + sDate + "' and '" + eDate + "' and ClosedFlag='Y' and a.line like '" + line + "%' order by GroupID";
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
        }

        public DataTable GetLine()
        {
            string strSQL = "select distinct line from QSMS_WoGroup with(nolock) order by line";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetLinefromMachine()
        {
            string strSQL = "select distinct Line from Machine order by Line";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetGroupWO(string GroupID)
        {
            string strSQL = "select Work_Order from QSMS_WOGroup  where GroupID= '" + GroupID + "' order by Seq_NO";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetWoInfo (string WO)
        {
            string strSQL = "select Line,PN ,Qty ,Trans_date from SAP_WO_list Where WO='" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetListWO(string BU, string BeginDate, string EndDate, string Line)
        {
            string strSQL = "";
            if (BU == "NB3" || BU == "NB5")
            {
                strSQL = "select  WO from Sap_Wo_List where WO_Type='PP10' And Trans_Date between '" + BeginDate + "' and '" + EndDate + "' and line like '" + Line + "%'";
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            else
            {
                strSQL = "select  WO from Sap_Wo_List where Trans_Date between '" + BeginDate + "' and '" + EndDate + "' and line like '" + Line + "%'";
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
        }
        public DataTable GetGroupID(string Line, string BeginDate, string EndDate)
        {
            string strSQL = "select distinct GroupID from QSMS_WOGroup A where substring(Group_TransDateTime,1,8) between '" + BeginDate + "' and '" + EndDate + "' and Line='"+Line+ "' and exists(select 0 from QSMS_WOGroup B where a.GroupID=b.GroupID and b.Closedflag='N') ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetGroupIDbyRelease(string Line,string BeginDate,string EndDate)
        {
            string strSQL = "select distinct GroupID from QSMS_WOGroup A  where wo_TransDateTime between '" + BeginDate + "' and '" + EndDate + "' and Line='" + Line + "' and exists(select 0 from QSMS_WOGroup B where a.GroupID=b.GroupID and b.Closedflag='N')";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable CheckWOGroup(string WO)
        {
            string strSQL = "select * from QSMS_WoGroup where Work_Order='" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable CheckDispatch(string WO)
        {
            string strSQL = "Select Count(*) as Qty From QSMS_Dispatch where Work_Order='" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable DeleteWOByGroup(string GroupID, string WO ,string UID)
        {
            string strSQL = "exec DeleteWOByGroup '" +GroupID+ "','"+ WO +"','" + UID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable DblChkLine(string WO, string Line)
        {
            string strSQL = "select WO from SAP_WO_list Where WO='" + WO + "' and line='" + Line + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GenGroupID(string TempGroupHead)
        {
            string strSQL = "select top 1 GroupID  from QSMS_WOGroup  where GroupID like '" + TempGroupHead + "%' order by GroupID desc";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable CHKMaintainWO(string WOList,string Line,string GroupID)
        {
            string strSQL = "Exec CHKMaintainWO '"+WOList+ "','"+Line+"','" + GroupID+ "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataSet ChkPNGroup(string WOList)
        {
            string strSQL = "Exec ChkPNGroup '" + WOList + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
        public DataTable CheckWOGroupID(string WO, string GroupID)
        {
            string strSQL = "Exec CheckWOGroupID '" + WO + "','" + GroupID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_ChkWOGroupID(string WO, string GroupID)
        {
            string strSQL = "Exec QSMS_ChkWOGroupID '" + WO + "','" + GroupID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_ChkGroupID( string GroupID)
        {
            string strSQL = "Exec QSMS_CheckGroupID '" + GroupID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable ChkWOGroup_His(string WO)
        {
            string strSQL = "select Work_Order ,GroupID from QSMS_WoGroup where Work_Order='" + WO + "' union all select Work_Order ,GroupID from QSMS_History.dbo.qsms_wogroup where Work_Order='" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable ChkMBWo(string WO)
        {
            string strSQL = "select WO from Sap_Wo_list where wo='" + WO + "' and InitAOIFlag='Y'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void Insert_QSMSWoGroup(string WO, string PN, string MBFlag, string Line, string Seq_No, string GroupID, string TransDate, string GroupDateTime, string UID)
        {
            string strSQL = "insert into QSMS_WoGroup(Work_Order,MBPN,MBFlag,Line,Seq_No,GroupID,WO_TransDateTime,Group_TransDateTime,Sap1Flag,ClosedFlag,ClosedType,UID)values ('" + WO + "','" + PN + "','" + MBFlag + "','" + Line + "','" + Seq_No + "','" + GroupID + "','" + TransDate + "','" + GroupDateTime + "','N','N','','" + UID + "')";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable XL_CheckWOGroupID(string WO,string GroupID)
        {
            string strSQL = "EXEC XL_CheckWOGroupID '" + WO+ "','" + GroupID+ "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetMaxSeq(string GroupID)
        {
            string strSQL = "Select Max(Seq_NO) as Max from QSMS_WoGroup where GroupID='" + GroupID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void Update_QSMSWoGroup(string WO, string Seq_No)
        {
            string strSQL = "Update QSMS_WoGroup set Seq_NO='" + Seq_No + "' where Work_Order='" + WO + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable Get_GroupID(string WO)
        {
            string strSQL = "select  GroupID from QSMS_WoGroup where Work_Order='" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
    }
}
