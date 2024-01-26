using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace QSMS.DbLibrary.Report
{
    class BeforeHandCheckBom
    {
        public DataTable GetLine()
        {
            string strSQL = "select distinct Line from QSMS_woGroup order by line";


            try
            {
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }
        public DataTable GetModel()
        {
            string strSQL = "select distinct MBPN +'-'+MBRev as Model from SAPBOM order by Model";
            try
            {
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }
        public DataTable CheckOP()
        {
            string strSQL = "select * from Sap_Wo_List where WO='VIRTUALWO' and Trans_Date>dbo.formatdate(dateadd(N,-8,getdate()),'YYYYMMDDHHNNSS')";
            try
            {
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }
        public DataTable beforehandCheckBom(string Factory, string Line, string PN, string Rev, int CombineQty)
        {
            string SPName = "QSMS_BeforehandCheckBom";

            SqlParameter[] paras = new SqlParameter[5];
            try
            {
                paras[0] = new SqlParameter("@Factory", SqlDbType.VarChar) { Value = Factory };
                paras[1] = new SqlParameter("@Line", SqlDbType.VarChar) { Value = Line };
                paras[2] = new SqlParameter("@PN", SqlDbType.VarChar) { Value = PN };
                paras[3] = new SqlParameter("@Rev", SqlDbType.VarChar) { Value = Rev };
                paras[4] = new SqlParameter("@CombineQty", SqlDbType.Int) { Value = CombineQty };
                return SqlHelper.ExecuteDataTable(SPName, paras, Parameter.ConnQSMS);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }
        public DataTable CheckBomFail()
        {
            string strSQL = "select * from Sap_Bom_Fail where Work_Order='VIRTUALWO'";
            try
            {
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }
    }
}
