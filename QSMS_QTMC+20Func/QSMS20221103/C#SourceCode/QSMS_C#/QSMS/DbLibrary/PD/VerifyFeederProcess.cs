using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace QSMS.DbLibrary.PD
{
    public class VerifyFeederProcess
    {
        public DataTable GetWoByGroupID(string groupID)
        {
            string strSQL = "Select Work_Order,Sap1Flag,ClosedFlag from QSMS_WoGroup with(nolock) where GroupID='"+groupID+"' ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetJobByMachine(string machine,string wo)
        {
            string strSQL = "Select Distinct JobPN from QSMS_MEBom with(nolock) where Machine='"+machine+"' and JobPN in (select Jobpn from QSMS_JobBOM where Work_order='"+wo+"') ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetGroupIDByLine(string sDate,string Line)
        {
            string strSQL = "Select distinct GroupID from QSMS_WoGroup where GroupID>'"+sDate+Line+"' ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetMachineByWo(string wo,string rev,string type)
        {
            string strSQL = "";
            if (type == "GetRev")
            {
                strSQL= "Select Mb_Rev from Sap_Wo_List where Wo='" + wo + "' ";
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            strSQL = "select distinct Machine From QSMS_MEbom with(nolock) where JobPN in (select JobPN from QSMS_JobBOM with(nolock) where Work_Order='"+wo+"') and  Version='"+rev+"' ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetCompPN(string machine,string jobPN,string Version,string slot,string LR)
        {
            string strSQL = "select CompPN from QSMS_MEBom with(nolock) where Machine='"+machine+"' and JObpN='"+jobPN+"' and version='"+Version+"' and Slot='"+slot.Trim(' ')+"' and LR='"+LR.Trim(' ')+"' ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable UpdateQSMS_Feeder(string slot,string feeder)
        {
            string strSQL = "Update QSMS_Feeder set Slot='"+slot.Trim(' ')+"' where Feeder='"+feeder+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable ChkMachineVerifyFinished(string machine, string jobPN, string Version)
        {
            string strSQL = "select CompPN from QSMS_MEBom with(nolock) where Machine='" + machine + "' and JObpN='" + jobPN + "' and version='" + Version + "' and Slot='' ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable UpdateQSMS_Verify(string machine,string jobPN,string version,string compPN,string venderCode,string transDateTime,string type)
        {
            string strSQL = "";
            if (type == "1")
            {
                strSQL = "select CompPN,VendorCode,DateCode,LotCode from qsms_verify where machine='"+machine+"' and enddatetime=''";
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            else if(type == "2")
            {
                strSQL= "Select CompPN from QSMS_Feeder where Machine='"+machine+"' and JobPN='"+jobPN+"' and Version='"+version+"' and CompPN='"+compPN+"' and VendorCode='"+venderCode.Trim(' ')+"'";
                return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            strSQL = "Update QSMS_Feeder set TransDateTime = '"+transDateTime+"' where Machine = '"+machine+"' and TransDateTime = '' and CompPN = '"+compPN+"' and VendorCode = '"+compPN+"' ";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            return null;
        }
        public DataTable GetCodesByFeeder(string feeder)
        {
            string strSQL = "select CompPN,VendorCode,DateCode,LotCode from QSMS_Feeder where Feeder='"+feeder+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }


    }
}
