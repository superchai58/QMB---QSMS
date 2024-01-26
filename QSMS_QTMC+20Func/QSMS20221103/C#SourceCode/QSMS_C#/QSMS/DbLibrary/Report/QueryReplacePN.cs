using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace QSMS.DbLibrary.Report
{
    class QueryReplacePN
    {
        public DataSet QuerySAPBOM(string PN, string Model)
        {
            string SPName = "QSMS_QuerySAP_BOM";

            SqlParameter[] paras = new SqlParameter[2];
            try
            {
                paras[0] = new SqlParameter("@PN", SqlDbType.VarChar) { Value = PN };
                paras[1] = new SqlParameter("@Model", SqlDbType.VarChar) { Value = Model };
                return SqlHelper.ExecuteDataSet(SPName, paras, Parameter.ConnQSMS);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }
    }
}
