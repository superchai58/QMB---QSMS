using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Text;

namespace QSMS.DbLibrary.SpecialCase
{
    public class CloseUnCheckWO
    {
        public DataTable CloseWO_UnCheck()
        {
            string strSQL = "SELECT * FROM dbo.CloseWO_UnCheck WITH(NOLOCK) ORDER BY TransDateTime DESC";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetCloseWO_UnCheck(string WO)
        {
            string strSQL = "SELECT * FROM dbo.CloseWO_UnCheck WITH(NOLOCK) WHERE WO='" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void Insert_data(string WO, string UID)
        {
            string strSQL = "Insert into CloseWO_UnCheck(WO,UID,TransDateTime) values('" + WO + "','" + UID + "',dbo.formatdate(getdate(),'YYYYMMDDHHNNSS'))";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void Delete_data(string WO)
        {
            string strSQL = "Delete CloseWO_UnCheck where WO='" + WO + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
    }
}
