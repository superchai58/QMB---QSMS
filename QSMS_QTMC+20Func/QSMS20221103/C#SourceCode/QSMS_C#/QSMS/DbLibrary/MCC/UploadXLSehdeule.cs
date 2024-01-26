using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace QSMS.DbLibrary.MCC
{
    public class UploadXLSehdeule : QMSSDK.Db.WinForm
    {
        public DataTable GetXL_WOPlanSeq_Tmp()
        {
            this.spName = "XL_UploadPlanSeq_AutoGroup";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = "GetXL_WOPlanSeq_Tmp" };
            paras[1] = new SqlParameter("@UID", SqlDbType.VarChar, 50) { Value = Parameter.g_userName };

            return SqlHelper.ExecuteDataTable(spName, paras, Parameter.ConnQSMS);
        }
        public DataSet UploadTemp(string Date,string Factory,string xml)
        {

            this.spName = "XL_UploadPlanSeq_AutoGroup";
            SqlParameter[] paras = new SqlParameter[5];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = "InsertTemp" };
            paras[1] = new SqlParameter("@UID", SqlDbType.VarChar, 50) { Value = Parameter.g_userName };
            paras[2] = new SqlParameter("@Date", SqlDbType.VarChar, 20) { Value = Date };
            paras[3] = new SqlParameter("@Factory", SqlDbType.VarChar, 20) { Value = Factory };
            paras[4] = new SqlParameter("@Excel", SqlDbType.Xml) { Value = xml };

            //return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
           
            return SqlHelper.ExecuteDataSet(spName, paras, Parameter.ConnQSMS);
        }
        public DataTable GetWOAnalyze(string WO)
        {
            //string strSQL = "Exec XL_UploadPlanSeq_AutoGroup @Type = 'GetAnalyze',@WO = '" + WO + "'";

            this.spName = "XL_UploadPlanSeq_AutoGroup";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = "GetAnalyze" };
            paras[1] = new SqlParameter("@WO", SqlDbType.VarChar, 50) { Value = WO };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }
        public int UpdateUpdateXL_WOPlanSeq_Tmp(string WO,string GroupID,ref string msg,string flag = "")
        {
            DataTable dt;
            //string strSQL = "Exec XL_UploadPlanSeq_AutoGroup @Type = 'UpdateXL_WOPlanSeq_Tmp',@WO = '" + WO + "',@GroupID = '"+ GroupID + "'";
            this.spName = "XL_UploadPlanSeq_AutoGroup";
            SqlParameter[] paras = new SqlParameter[4];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = "UpdateXL_WOPlanSeq_Tmp" };
            paras[1] = new SqlParameter("@WO", SqlDbType.VarChar, 50) { Value = WO };
            paras[2] = new SqlParameter("@GroupID", SqlDbType.VarChar, 50) { Value = GroupID };
            paras[3] = new SqlParameter("@Flag", SqlDbType.VarChar, 50) { Value = flag };

            dt = SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
            msg = dt.Rows[0]["Msg"].ToString();
            if (dt.Rows[0]["Result"].ToString().ToUpper() == "PASS")
            {              
                return 1;
            }
            else if(dt.Rows[0]["Result"].ToString().ToUpper() == "Warning")
            {
                return 2;
            }
            else
            {
                return 3;
            }
        }
        public string GetNewGroup(string WO)
        {
            DataTable dt = new DataTable();
            this.spName = "XL_UploadPlanSeq_AutoGroup";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = "GetNewGroup" };
            paras[1] = new SqlParameter("@WO", SqlDbType.VarChar, 50) { Value = WO };


            dt = SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);

            //string strSQL = "Exec XL_UploadPlanSeq_AutoGroup @Type = 'GetNewGroup',@WO = '" + WO + "'";
            //dt = SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            if(dt.Rows.Count > 0)
            {
                return dt.Rows[0]["Group"].ToString();
            }
            else
            {
                return "";
            }
        }
        public DataTable AssignGroupID()
        {
            this.spName = "XL_UploadPlanSeq_AutoGroup";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = "AssignGroupID" };
            paras[1] = new SqlParameter("@UID", SqlDbType.VarChar, 50) { Value = Parameter.g_userName };

            //return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);

            return SqlHelper.ExecuteDataTable(spName, paras, Parameter.ConnQSMS);

        }
        public DataTable ReAssignGroupID()
        {
            this.spName = "XL_UploadPlanSeq_AutoGroup";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = "ReAssignGroupID" };
            paras[1] = new SqlParameter("@UID", SqlDbType.VarChar, 50) { Value = Parameter.g_userName };

            //return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);

            return SqlHelper.ExecuteDataTable(spName, paras, Parameter.ConnQSMS);

        }
        public DataTable SaveSchedule()
        {
            this.spName = "XL_UploadPlanSeq_AutoGroup";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = "UploadXL_WOPlanSeq" };
            paras[1] = new SqlParameter("@UID", SqlDbType.VarChar, 50) { Value = Parameter.g_userName };
            paras[2] = new SqlParameter("@Factory", SqlDbType.VarChar, 20) { Value = Parameter.Factory };

            //return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);

            return SqlHelper.ExecuteDataTable(spName, paras, Parameter.ConnQSMS);

        }
        public bool CheckUpload(out string msg)
        {
            DataTable dt = new DataTable();
            this.spName = "XL_UploadPlanSeq_AutoGroup";
            SqlParameter[] paras = new SqlParameter[3];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = "ChkUpload" };
            paras[1] = new SqlParameter("@UID", SqlDbType.VarChar, 50) { Value = Parameter.g_userName };
            paras[2] = new SqlParameter("@Factory", SqlDbType.VarChar, 20) { Value = Parameter.Factory };

            dt = SqlHelper.ExecuteDataTable(spName, paras, Parameter.ConnQSMS);

            if (dt.Rows[0]["Result"].ToString().ToUpper() == "PASS")
            {
                msg = "";
                return true;
            }
            else if(dt.Rows[0]["Result"].ToString().ToUpper() == "NOTICE")
            {
                msg = dt.Rows[0]["Msg"].ToString();
                return true;
            }
            else
            {
                msg = dt.Rows[0]["Msg"].ToString();
                return false;
            }

        }
        public bool DelUploadTemp(out string msg)
        {
            DataTable dt;
            bool bolReturn = false;
            msg = "";
            this.spName = "XL_UploadPlanSeq_AutoGroup";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = "DelUploadTemp" };
            paras[1] = new SqlParameter("@UID", SqlDbType.VarChar, 50) { Value = Parameter.g_userName };

            dt = SqlHelper.ExecuteDataTable(spName, paras, Parameter.ConnQSMS);

            if(dt.Rows.Count > 0)
            {
                if(dt.Rows[0]["Result"].ToString().ToUpper() == "PASS")
                {
                    bolReturn = true;
                }
                else
                {
                    msg = dt.Rows[0]["MSG"].ToString();
                }
            }

            return bolReturn;
        }
        public DataTable QuerySchedule(string Date)
        {
            this.spName = "XL_UploadPlanSeq_AutoGroup";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = "QuerySchedule" };
            paras[1] = new SqlParameter("@Date", SqlDbType.VarChar, 50) { Value = Date };

            //return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);

            return SqlHelper.ExecuteDataTable(spName, paras, Parameter.ConnQSMS);

        }
    }
}
