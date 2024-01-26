using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;

namespace QSMS.DbLibrary.Report
{
   public  class QueryCheckBOMProcess
    {
        public DataTable GetLine()
        {

            string strSQL = "select distinct line from QSMS_WoGroup  order by line";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);

        }
        public DataTable QSMS_QueryCheckBOM(string wo,string line,string BeginDate,string EndDate)
        {

            string strSQL = "EXEC QSMS_QueryCheckBOM '"+ wo + "','"+ line + "','"+ BeginDate + "','"+ EndDate + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);

        }
    }
}
