using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace QSMS.DbLibrary.PD
{
    class UnlinkFeederDID
    {
        public DataTable QueryFeederDID(string Feeder)
        {
            string strSQL = "select DID from QSMS_Feeder where Feeder='" + Feeder + "'";
            try
            {
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }
        public DataTable UnlinkFeeder(string Feeder, string UID, string Type)
        {
            string SPName = "PDA_UnLinkFeeder";

            SqlParameter[] paras = new SqlParameter[3];
            try
            {
                paras[0] = new SqlParameter("@Feeder", SqlDbType.VarChar) { Value = Feeder };
                paras[1] = new SqlParameter("@OPID", SqlDbType.VarChar) { Value = UID };
                paras[2] = new SqlParameter("@Type", SqlDbType.VarChar) { Value = Type };
                return SqlHelper.ExecuteDataTable(SPName, paras, Parameter.ConnQSMS);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }
        public void DeleteFeederDID(string UID, string DID)
        {
            string strSQL = "Insert into QSMS_Feeder_Delete(Machine,JobPN,Version,DID,VendorCode,DateCode,LotCode,Feeder,Slot,LR,[UID],TransDateTime,DeleteDateTime) Select Machine,JobPN,Version,a.DID,VendorCode,DateCode,LotCode,Feeder,Slot,LR," + "'" + UID + "'" + ",TransDateTime,dbo.FormatDate(GETDATE(),'YYYYMMDDHHNNSS') from QSMS_Feeder a where a.DID='" + DID + "'";
            try
            {
                SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
        }

        public void DeleteFromFeederDID(string DID)     //0001
        {
            string strSQL = "delete from QSMS_Feeder where DID='" + DID + "'";
            try
            {
                SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }

        }

        public void DeleteFromFeederDID_Current(string DID)     //20220628  增加删除QSMS_FeederDID_Current表数据的动作
        {
            string strSQL = "delete from QSMS_FeederDID_Current where DID='" + DID + "'";
            try
            {
                SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }

        }


    }
}

