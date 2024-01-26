using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;


namespace QSMS.DbLibrary.IPQC
{
    public class IPQCProcess : QMSSDK.Db.WinForm
    {
        public DataTable QSMS_PD_QueryDataByType(string Type, string BeginDate, string EndDate, string Item1, string Item2,string Item3)
        {
            this.spName = "QSMS_PD_QueryDataByType";
            SqlParameter[] paras = new SqlParameter[6];
            paras[0] = new SqlParameter("@Type", SqlDbType.VarChar, 50) { Value = Type };
            paras[1] = new SqlParameter("@BeginDate", SqlDbType.VarChar, 14) { Value = BeginDate };
            paras[2] = new SqlParameter("@EndDate", SqlDbType.VarChar, 14) { Value = EndDate };
            paras[3] = new SqlParameter("@Item1", SqlDbType.VarChar, 200) { Value = Item1 };
            paras[4] = new SqlParameter("@Item2", SqlDbType.VarChar, 200) { Value = Item2 };
            paras[5] = new SqlParameter("@Item3", SqlDbType.VarChar, 200) { Value = Item3 };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable GetLotNo()
        {
            string strSQL = "exec GetLotNo";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataSet QSMSDIDInSpect(string strDID, string lblcomppn, string lblVendor, string TestValue, string Errcode, string ScanCompPN)
        {
            string strSQL = "exec QSMSDIDInSpect '" +strDID + "','" + lblcomppn + "','" + lblVendor + "','" + TestValue + "' ,'" + Errcode + "','" + Parameter.g_userName + "','" + ScanCompPN + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }

        public DataTable QSMS_GenUNID(string BarCode, string Type)
        {
            this.spName = "QSMS_GenUNID";
            SqlParameter[] paras = new SqlParameter[2];
            paras[0] = new SqlParameter("@2DBarCode", SqlDbType.NVarChar, 4000) { Value = BarCode };
            paras[1] = new SqlParameter("@Type", SqlDbType.VarChar, 20) { Value = Type };
            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);
        }

        public DataTable QueryInSpect(string DID, string SDateTime, string EDateTime)
        {
            this.spName = "PD_QSMS_QueryInSpect";

            SqlParameter[] paras = new SqlParameter[3];

            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar) { Value = DID };
            paras[1] = new SqlParameter("@SDate", SqlDbType.VarChar) { Value = SDateTime };
            paras[2] = new SqlParameter("@EDate", SqlDbType.VarChar) { Value = EDateTime };

            return SqlHelper.ExecuteDataTable(this.spName, paras, Parameter.ConnQSMS);

            //string strSQL = "SELECT * FROM dbo.QSMS_DID_InSpect WITH(NOLOCK) WHERE (DID like '%" + DID + "%' or CompPN like '%" + DID + "%') and transDatetime between '" + SDateTime + "' and '" + EDateTime + "' ORDER BY TransDateTime DESC";
            //return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void DelInSpect(string CompPN, string Transdatetime)
        {
            string strSQL = "delete QSMS_DID_InSpect where compPn='" + CompPN + "' and Transdatetime='" + Transdatetime + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataSet InRelieve(string DID, string SDTime, string EDTime, string TransDate, string QPN, string Vendor, string Errcode, string CompPN, float TestValue, string ScanCompPN, string UID, string FuncType)
        {
            if (FuncType.ToUpper() == "QUERYBYDID")
            {
                DID = "%" + DID + "%";
            }

            this.spName = "QSMS_Inspection_Action";
            SqlParameter[] paras = new SqlParameter[12];
            paras[0] = new SqlParameter("@DID", SqlDbType.VarChar) { Value = DID };
            paras[1] = new SqlParameter("@SDTime", SqlDbType.VarChar) { Value = SDTime };
            paras[2] = new SqlParameter("@EDTime", SqlDbType.VarChar) { Value = EDTime };
            paras[3] = new SqlParameter("@TransDate", SqlDbType.VarChar) { Value = TransDate };
            paras[4] = new SqlParameter("@QPN", SqlDbType.VarChar) { Value = QPN };
            paras[5] = new SqlParameter("@Vendor", SqlDbType.VarChar) { Value = Vendor };
            paras[6] = new SqlParameter("@Errcode", SqlDbType.VarChar) { Value = Errcode };
            paras[7] = new SqlParameter("@CompPN", SqlDbType.VarChar) { Value = CompPN };
            paras[8] = new SqlParameter("@TestValue", SqlDbType.Float) { Value = TestValue };
            paras[9] = new SqlParameter("@ScanCompPN", SqlDbType.VarChar) { Value = ScanCompPN };
            paras[10] = new SqlParameter("@UID", SqlDbType.VarChar) { Value = UID };
            paras[11] = new SqlParameter("@FuncType", SqlDbType.VarChar) { Value = FuncType };

            return SqlHelper.ExecuteDataSet(this.spName, Parameter.ConnQSMS);
        }
    }
}
