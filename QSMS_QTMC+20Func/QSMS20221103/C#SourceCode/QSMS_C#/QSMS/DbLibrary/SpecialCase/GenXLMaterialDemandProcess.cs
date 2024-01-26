using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace QSMS.DbLibrary.SpecialCase
{
    public class GenXLMaterialDemandProcess
    {
        public DataTable GetSite()
        {
            string strSQL = "Select distinct Factory from Site";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetXL_Type()
        {
            string strSQL = "Select  XL_Type from XL_TypeDateTime order by cast(XL_Type as int )";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable XL_Job_12H_NB4(string OPID, string factory, string type)
        {
            string strSQL = "exec XL_JOB_12Hours_GenMD @OPID='" + OPID + "',@Factory='" + factory + "',@PNInterval='" + type + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable XL_JOB_PO(string OPID, string factory)
        {
            string strSQL = "exec XL_JOB_8Hours_GenMD @OPID='" + OPID + "',@Factory='" + factory + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable XL_JOB_Others(string OPID, string factory)
        {
            string strSQL = "exec XL_JOB_12Hours_GenMD @OPID='" + OPID + "',@Factory='" + factory + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
    }
}
