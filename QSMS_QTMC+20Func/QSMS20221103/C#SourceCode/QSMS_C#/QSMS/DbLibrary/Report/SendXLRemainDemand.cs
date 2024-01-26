using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;


namespace QSMS.DbLibrary.Report
{
    public class SendXLRemainDemand
    {
        public DataSet QueryRemainDemand(string Date, string Shift, string Factory)
        {
            string strSQL = "EXEC QueryXLRemainDemand '" + Date + "','" + Shift + "','" + Factory + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }

        public DataTable SendRemainDemand(string Date, string Shift, string Factory)
        {
            string strSQL = "EXEC SendXLRemainDemand '" + Date + "','" + Shift + "','" + Factory + "'";
            //SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
    }
}
