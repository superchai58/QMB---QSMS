using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QSMS.DbLibrary.SpecialCase
{
    class SpecialCaseProcess
    {
        public void QSMS_SplitLineMC(string UID)
        {
            string strSQL = "exec QSMS_SplitLineMC '" + UID + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
    }
}
