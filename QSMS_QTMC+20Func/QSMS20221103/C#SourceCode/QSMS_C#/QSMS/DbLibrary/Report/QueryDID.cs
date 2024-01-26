using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace QSMS.DbLibrary.Report
{
    public class QueryDID
    {
        public DataTable GetMachine()
        {
            string strSQL = "SELECT distinct Machine FROM QSMS_Verify where Machine>'' order by Machine";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetLine()
        {
            string strSQL = "select distinct left(machine,1) as line from qsms_verify where machine>'' order by 1";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetSlot()
        {
            string strSQL = "select distinct Slot from QSMS_Verify order by Slot";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetComPN()
        {
            string strSQL = "select distinct CompPN from QSMS_Verify order by CompPN";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable Query(string strsql)
        {
            string strSQL = "select A.*,Isnull(B.SplicingDT,Isnull(C.SplicingDT,'')) as SplicingDT,IsNull(B.Qty,Isnull(C.Qty,-1)) as Qty from QSMS_Verify A left Join QSMS_DID B ON A.DID=B.DID Left join QSMS_DID_Log C on A.DID=C.DID where "+strsql;
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataSet QueryDIDUse(string DID)
        {
            string strSQL = "EXEC QueryDIDUse '" + DID + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
    }
}
