using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace QSMS.DbLibrary.SpecialCase
{
    class UrgentWO
    {
        public DataSet QueryWOSeq(string WO)
        {
            //string SPName = "XL_SpecialCaseByWO";    
            //QSMC 启用新的SP
            string SPName = "XL_SpecialCaseByWO_New"; 
            SqlParameter[] paras = new SqlParameter[1];
            try
            {
                paras[0] = new SqlParameter("@WO", SqlDbType.VarChar) { Value = WO };               
                
                return SqlHelper.ExecuteDataSet(SPName,paras,Parameter.ConnQSMS);
            }
            catch(Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }
        public DataTable CheckUrgentWO(string WO)
        {
            string SPName = "ChkUrgentWO";

            SqlParameter[] paras = new SqlParameter[1];
            try
            {
                paras[0] = new SqlParameter("@WO", SqlDbType.VarChar) { Value = WO };

                return SqlHelper.ExecuteDataTable(SPName, paras, Parameter.ConnQSMS);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }
        public void InsertUrgentWO(string UID,string WO,string WorkDate,string Shift)
        {
            //string SPName = "XL_SpecialCaseByWO"; 
            //QSMC 启用新的SP
            string SPName = "XL_SpecialCaseByWO_New";

            SqlParameter[] paras = new SqlParameter[4];
            try
            {
                paras[0] = new SqlParameter("@UID", SqlDbType.VarChar) { Value = UID };
                paras[1] = new SqlParameter("@WO", SqlDbType.VarChar) { Value = WO };
                paras[2] = new SqlParameter("@WorkDate", SqlDbType.VarChar) { Value = WorkDate };
                paras[3] = new SqlParameter("@Shift", SqlDbType.VarChar) { Value = Shift };
                SqlHelper.ExecuteDataTable(SPName, paras, Parameter.ConnQSMS);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }
    }
}
