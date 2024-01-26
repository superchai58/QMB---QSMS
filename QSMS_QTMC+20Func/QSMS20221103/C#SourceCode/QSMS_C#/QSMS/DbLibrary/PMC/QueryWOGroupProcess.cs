using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace QSMS.DbLibrary.PMC
{
    public class QueryWOGroupProcess
    {
        public DataTable GetDataByWO(string type,string wo,string line)
        {
            if (type == "UnClosedByWO")
            {
                string strSQL = "select a.GroupID,a.Seq_NO,a.Work_Order,a.Line,b.PN,B.MB_Rev,B.Qty,a.Wo_TransDateTime,a.Group_TransDateTime,a.DispatchFlag,a.Sap1Flag,a.ClosedFlag,a.ClosedType,a.UID,a.CloseDateTime from QSMS_WOGroup a with(nolock),Sap_Wo_List b with(nolock) where a.work_order=b.wo and a.groupID in (select Groupid from qsms_woGroup where work_order='"+wo+"') and ClosedFlag='N' and a.line like '"+line+"%' order by groupID";
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            else
            {
                string strSQL = "select a.GroupID,a.Seq_NO,a.Work_Order,a.Line,b.PN,B.MB_Rev,B.Qty,a.Wo_TransDateTime,a.Group_TransDateTime,a.DispatchFlag,a.Sap1Flag,a.ClosedFlag,a.ClosedType,a.UID,a.CloseDateTime from QSMS_WOGroup a with(nolock),Sap_Wo_List b with(nolock) where a.work_order=b.wo and a.groupID in (select Groupid from QSMS_woGroup where work_order='"+wo+"') and ClosedFlag='Y' and a.line like '"+line+"%' order by GroupID";
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
        }
        public DataTable GetData(string type, string sDate, string eDate,string line)
        {
            if (type == "UnClosedWO")
            {
                string strSQL = "select a.GroupID,a.Seq_NO,a.Work_Order,a.Line,b.PN,B.MB_Rev,B.Qty,a.Wo_TransDateTime,a.Group_TransDateTime,a.DispatchFlag,a.Sap1Flag,a.ClosedFlag,a.ClosedType,a.UID,a.CloseDateTime from qsms_WOGroup a with(nolock),Sap_Wo_List b  where a.Work_Order=b.Wo and substring(Group_TransDateTime,1,8) between '" + sDate+"' and '"+eDate+"' and ClosedFlag='N' and a.line like '"+line+"%' order by groupID";
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            else
            {
                string strSQL = "select a.GroupID,a.Seq_NO,a.Work_Order,a.Line,b.PN,B.MB_Rev,B.Qty,a.Wo_TransDateTime,a.Group_TransDateTime,a.DispatchFlag,a.Sap1Flag,a.ClosedFlag,a.ClosedType,a.UID,a.CloseDateTime from qsms_WOGroup a with(nolock),Sap_Wo_List b  where a.Work_Order=b.Wo and substring(Group_TransDateTime,1,8) between '" + sDate+"' and '"+eDate+"' and ClosedFlag='Y' and a.line like '"+line+"%' order by GroupID";
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
        }

        public DataTable GetLine()
        {
            string strSQL = "select distinct line from QSMS_WoGroup with(nolock) order by line";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
    }
}
