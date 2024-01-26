using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QSMS.DbLibrary
{
    class DbLogin
    {
        public DataTable GetQSMSServer(string SMTServer)
        {
            string strSQL = "SELECT SMT_DB,QSMS_DB,QSMS_Server FROM dbo.QSMS_SMT_DB WITH(NOLOCK) WHERE SMT_Server='" + SMTServer + "'";
            return SqlHelper.ExecuteTable(strSQL,Parameter.ConnSMT);
        }
        public DataTable GetQSMS_ProConfig()
        {
            string strSQL = "SELECT * FROM dbo.QSMS_ProConfig WITH(NOLOCK) WHERE Line='All' AND station='QSMS'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetXL_SiteData(string Factory, string PrtCallBKandReturn)
        {
            string strSQL = "exec XL_SiteData '" + Factory + "','" + PrtCallBKandReturn + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable CheckFacIP()
        {
            string strSQL = "SELECT DISTINCT Factory FROM dbo.SITE WITH(NOLOCK)";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable QSMS_IPFactory(string Factory, string PrtCallBKandReturn)
        {
            string strSQL = "exec QSMS_IPFactory '" + Factory + "','" + PrtCallBKandReturn + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

    }
}
