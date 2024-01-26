using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace QSMS.DbLibrary.SpecialCase
{
    class UpdateUID
    {
        public DataTable QueryUID()
        {
            string strSQL = "select distinct UID  from QSMS_DID_ToWH with(nolock) where UID NOT IN (SELECT Username from userdetail with(nolock))and UID<>''";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void updateUID(string NewUID,string OldUID)
        {
            string strSQL = "update QSMS_DID_ToWH  set UID='" + NewUID+ "' where UID='" + OldUID + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
    }
}
