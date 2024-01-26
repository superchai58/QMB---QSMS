using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace QSMS.DbLibrary.MCC
{
    public class DIDCheckStockProcess
    {
        public DataSet XL_DIDChkStockByRefID(string refID,string userName,int type)
        {
            string strSQL = "";
            if (type == 0)
            {
                strSQL = "exec XL_DIDChkStockByRefID @Type='Query',@RefID='" + refID + "'";
            }
            else
            {
                strSQL = "exec XL_DIDChkStockByRefID @Type='Manual',@RefID='" + refID + "',@UserName='"+userName+"'";
            }
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }

        public DataSet XL_CheckRefID(string sDate, string eDate, string v)
        {
            string strSQL = "exec XL_CheckRefID @BeginDate='"+sDate+"',@EndDate='"+eDate+"'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
    }
}
