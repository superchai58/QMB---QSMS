using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace QSMS.DbLibrary.Report
{
    class PDReport
    {
        public DataTable QueryDispatch()
        {
            string strSQL = "select distinct line from QSMS_Dispatch";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataSet Query_NOUseDID(string Type, string Line,string Timer1,string Timer2)
        {
            string strSQL = "exec Query_NOUseDID '" + Type + "','" + Line + "','" + Timer1 + "','" + Timer2 + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
    }
}
