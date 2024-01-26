using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace QSMS.DbLibrary.MCC
{
    public class ModifyDIDTotalQty
    {
        public DataTable RefreshDg(string compPN)
        {
            string strSQL = "select top 20 DID,CompPN,VendorCode,DateCode,LotCode,Qty,RemainQty,UID,TransDateTime from QSMS_DID with(nolock) where CompPN like '" + compPN + "%' order by DID";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetQSMS_DID(string DID)
        {
            string strSQL = "Select DID,CompPN,VendorCode,DateCode,LotCode,Qty,UID,RemainQty,TransDateTime,UsedFlag From QSMS_DID with(nolock) where DID like '" + DID + "%' Order by CompPN,DID  ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable CheckFormat(string PartNumber)
        {
            string strSQL = "Exec CheckFormat 'PARTNUMBER','" + PartNumber + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetDate()
        {
            string strSQL = "select dbo.FormatDate(getdate(),'YYYYMMDDHHNNSS')";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetRemainQty(string TempDID)
        {
            string strSQL = "Select Qty,RemainQty from QSMS_DID where DID='" +TempDID+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void updateQSMS_DID(string TxtQty,string CboDID,string UID, Int64 intDIDInitQty)
        {
            string strSQL = "Update QSMS_DID Set UID='" + UID + "',Qty='" +TxtQty+"',RealQty='" +TxtQty+ "' Where DID='" + CboDID + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            strSQL = "QTY: " + intDIDInitQty.ToString() + " -> " + strSQL;
            string strlog = "insert into qms_log(system_name,event_no,sn,user_name,desc1,trans_date) values('ModifyDIDTotalQty','1','" + CboDID + "','" + UID + "','" + strSQL + "',dbo.formatdate(getdate(),'YYYYMMDDHHNNSS'))";
            
            SqlHelper.ExecuteTable(strlog, Parameter.ConnQSMS);
        }
       

    }
}
