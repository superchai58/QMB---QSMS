using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace QSMS.DbLibrary.Report
{
    public class QueryDIDNeedCutProcess
    {
        public DataTable GetByPN(string PN)
        {
            string strSQL = "select top 1 * from qsms_did where comppn='"+PN+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable QueryData(string strPNList, string sDate, string eDate)
        {
            string strSQL = "select top 5000 * from qsms_did with(nolock) where remainqty>0 and transdatetime between '"+sDate+"' and '"+eDate+"' and did not like '%-A%' AND QTY<>Remainqty and realqty>0 and comppn in "+strPNList+" and wogroup in (select distinct groupid from QSMS_WOGroup where ClosedFlag<>'Y') order by comppn,line,transdatetime";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QueryDataToExcel(string strPNList, string sDate, string eDate)
        {
            string strSQL = "select top 5000 * from qsms_did with(nolock) where remainqty>0 and transdatetime between '" + sDate + "' and '" + eDate + "' and did not like '%-A%' AND QTY<>Remainqty and realqty>0 and comppn in " + strPNList + " order by comppn,line,transdatetime";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
    }
}
