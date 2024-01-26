using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace QSMS.DbLibrary.QMS
{
    public class CompDiffProcess
    {
        public DataTable GetLine()
        {
            string strSql = "SELECT DISTINCT Line FROM dbo.QSMS_WoGroup WITH(NOLOCK)";
            return SqlHelper.ExecuteTable(strSql, Parameter.ConnQSMS);
        }

        public DataTable GetGroupID(string BeginDate, string EndDate, string Line)
        {
            string strSql = "SELECT DISTINCT GroupID FROM dbo.QSMS_WoGroup WITH(NOLOCK) WHERE Wo_TransDateTime BETWEEN '" + BeginDate + "'AND '" + EndDate + "' AND Line='" + Line + "'";
            return SqlHelper.ExecuteTable(strSql, Parameter.ConnQSMS);
        }
        public DataTable GetGroupWO(string GroupID)
        {
            string strSql = "SELECT DISTINCT Work_Order FROM dbo.QSMS_WoGroup WITH(NOLOCK) WHERE GroupID='" + GroupID + "'";
            return SqlHelper.ExecuteTable(strSql, Parameter.ConnQSMS);
        }

        public DataTable GetWoInfo(string WO)
        {
            string strSql = "SELECT PN,Qty,MB_Rev,Line FROM dbo.Sap_Wo_List WITH(NOLOCK) WHERE WO='"+WO+"'";
            return SqlHelper.ExecuteTable(strSql, Parameter.ConnQSMS);
        }

        public DataTable ChkWO(string WO)
        {
            string strSql = "EXEC [dbo].[QsmsComp_diff] '" + WO + "'";
            return SqlHelper.ExecuteTable(strSql, Parameter.ConnQSMS);
        }

    }
}
