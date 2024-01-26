using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace QSMS.DbLibrary.SpecialCase
{
    public class UrgentDIDToWH
    {
        public DataTable QSMS_DID_ToWH(string txtRefID)
        {
            string strSQL = "SELECT * FROM dbo.QSMS_DID_ToWH WITH(NOLOCK) WHERE ReferenceID='" + txtRefID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable XL_UrgentToWH(string txtRefID, string OPID, string Type)
        {
            string strSQL = "Exec XL_UrgentToWH  '" + txtRefID + "','"+ OPID + "','"+ Type + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        
    }
}
