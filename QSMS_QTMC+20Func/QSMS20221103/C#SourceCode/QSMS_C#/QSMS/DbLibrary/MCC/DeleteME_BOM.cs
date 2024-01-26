using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace QSMS.DbLibrary.MCC
{
    public class DeleteME_BOM
    {
        public DataTable GetLine()
        {
            string strSQL = "select distinct Line from QSMS_woGroup order by line" ;
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
      
        
        public DataTable GetGroupID(string JobPN, string BU, string BeginDate, string EndDate, string Line, bool OptRelease)
        {
            string TempJobPn = "";

            string strSQL = "";
            if (JobPN == "")
            {
                if (BU == "NB5")
                {
                    if (OptRelease == true)
                    {
                        strSQL = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" + BeginDate + "' and '" + EndDate + "' and line='" + Line + "' and closedflag='N' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )";

                    }
                    else
                    {
                        strSQL = "select distinct GroupID from QSMS_WOGroup  where substring(GroupID,4,8) between '" + BeginDate + "' and '" + EndDate + "' and line='" + Line + "' and closedflag='N' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )";
                    }
                }
                else
                {
                    if (OptRelease == true)
                    {
                        strSQL = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" + BeginDate + "' and '" + EndDate + "' and line='" + Line + "' and closedflag='N'";

                    }
                    else
                    {
                        if (BU == "PO")
                        {
                            strSQL = "select distinct GroupID from QSMS_WOGroup  where substring(GroupID,2,8) between '" + BeginDate + "' and '" + EndDate + "' and line='" + Line + "' and closedflag='N' ";

                        }
                        else
                        {
                            strSQL = "select distinct GroupID from QSMS_WOGroup  where substring(GroupID,4,8) between '" + BeginDate + "' and '" + EndDate + "' and line='" + Line + "' and closedflag='N' ";

                        }
                    }
                }

            }
            else
            {
                if (JobPN.IndexOf("-") > 0)
                {
                    TempJobPn = JobPN.Substring(1, 11);

                }
                else
                {
                    TempJobPn = JobPN;
                }
                if (OptRelease == true)
                {
                    strSQL = "select distinct GroupID from QSMS_WOGroup a,QSMS_JobBOM b   where a.WO_TransDateTime between  '" + BeginDate + "' and '" + EndDate + "' and a.line='" + Line + "' and a.work_order=b.work_order and b.jobpn='" + TempJobPn + "'and closedflag='N' ";

                }
                else
                {
                    strSQL = "select distinct GroupID from QSMS_WOGroup a,QSMS_JobBOM b   where a.WO_TransDateTime between  '" + BeginDate + "' and '" + EndDate + "' and a.line='" + Line + "' and a.work_order=b.work_order and b.jobpn='" + TempJobPn + "'and closedflag='N' ";
                }
            }

            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetjobPn(string BU, string Line, string BeginDate, string EndDate, bool optRelease)
        {
            string strSQL = "";
            if (optRelease == true)
            {
                strSQL = "select distinct JobPN from QSMS_WOGroup a, QSMS_JobBOM b where a.WO_TransDateTime between  '" + BeginDate + "' and '" + EndDate + "' and a.line='" + Line + "' and A.work_order=b.work_order";

            }
            else
            {
              
                strSQL = "select distinct JobPN from QSMS_WOGroup a, QSMS_JobBOM b where a.WO_TransDateTime between  '" + BeginDate + "' and '" + EndDate + "' and a.line='" + Line + "' and A.work_order=b.work_order";


            }
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetGroupWO(string GroupID, string TempJobPn)
        {
            string strSQL = "select distinct a.Work_Order from QSMS_WOGroup a,QSMS_JobBOM b  where a.GroupID = '" + GroupID + "' and a.work_Order = b.work_order and b.jobpn like '" + TempJobPn + "%'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetWO(string WO)
        {

            string strSQL = "Select WO from sap_wo_list where WO='" + WO + "' And status >= 10";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetJobGroupByJobRev(string Machine, string MBPN, string Rev)
        {
            string strSQL = "";
            if (Machine.Trim().Length > 6)
            {
                strSQL = "Select distinct jobgroup from qsms_mebom where (jobpn='" + MBPN + "' or jobpn in (select distinct jobpn from QSMS_JObBom where MBPN='" + MBPN + "')) and version='" + Rev + "'" + " and Machine like '" + Machine + "%' order by jobgroup ";
            }
            else
            {
                strSQL = "Select distinct jobgroup from qsms_mebom where (jobpn='" + MBPN + "' or jobpn in (select distinct jobpn from QSMS_JObBom where MBPN='" + MBPN + "')) and version='" + Rev + "' order by jobgroup";
            }
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetJobpn(string BeginDate, string EndDate, string line, bool radio)
        {
            string strSQL = "";
            if (radio == true)
            {
                strSQL = "select distinct c.jobpn,b.Mb_Rev from QSMS_WOGroup a ,sap_wo_list b,qsms_JobBOM c  where " + "a.WO_TransDateTime between  '" + BeginDate + "' and '" + EndDate + "' and a.line='" + line + "'" +
                          " and a.work_Order=b.wo and a.work_order=c.work_order order by c.jobpn,b.mb_rev";
            }
            else
            {
                strSQL = "select distinct c.jobpn,b.Mb_Rev  from QSMS_WOGroup a ,sap_wo_list b,qsms_JobBOM c where" + " substring(a.GroupID,2,8) between '" + BeginDate + "' and '" + EndDate + "' and a.line='" + line + "'" +
                        " and a.work_Order=b.wo and a.work_order=c.work_order  order by c.jobpn,b.mb_rev";
            }

            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetMachine(string Line)
        {
            string strSQL = "Select distinct machine from qsms_mebom where line like '"+ Line + "%'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetMachineWo(string WO)
        {
            string strSQL = "select [group] from sap_wo_list where wo='" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetComp(string Machine, string MBPN, string WO, string Line)
        {
            string strSQL = "";
            if (Machine.Trim().ToUpper() != "ALL")
            {
                strSQL = "select CompPN from QSMS_WO  where Work_Order='" + WO + "'  and Machine='" + Machine + "'";
            }
            else
            {
                strSQL = "select CompPN from QSMS_WO  where Work_Order='" + WO + "'";
            }
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetSlot(string Machine,string Line)
        {
            string strSQL = "select distinct slot from qsms_mebom where machine='" + Machine + "'  and line = '" + Line + "'order by slot";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetWoinfo(string WO)
        {

            string strSQL = "select PN, Qty ,MB_Rev,Line,BuildType from sap_wo_list where WO='" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetCustomer(string TxtMBPN)
        {

            string strSQL = "select Customer from ModelName where PN='" + TxtMBPN+ "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetjobPn(string WO)
        {

            string strSQL = "select jobPn from QSMS_JobBOM where work_Order='" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_GetEMSFlagMB(string Group)
        {
            string strSQL = "Select a.JobPN,b.MB_Rev from QSMS_JobBom a,Sap_Wo_List b where b.[Group]='" + Group + "'  and a.Work_Order=b.WO and b.PN like '%MB%'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_GetEMSFlagSB(string Group)
        {
            string strSQL = "Select a.JobPN,b.MB_Rev from QSMS_JobBom a,Sap_Wo_List b where b.[Group]='" + Group + "'  and a.Work_Order=b.WO ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_GetEMSFlag(string Group)
        {
            string strSQL = "exec QSMS_GetEMSFlag '" + Group + "' ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_GetEMSFlagInitAOIFlag(string Group)
        {
            string strSQL = "Select a.JobPN,b.MB_Rev from QSMS_JobBOM a,Sap_Wo_List B where b.[Group]='" + Group+ "' and a.work_order=b.wo and b.InitAOIFlag='Y'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetSite()
        {
            string strSQL = "select site from site" ;
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        
        public DataTable GetMachine(string WO, string joppn)
        {
            string strSQL = "Select distinct machine from qsms_mebom a,qsms_jobbom b,sap_wo_list c where b.work_order='" + WO + "' and " +
            "b.work_order=c.wo and a.line=c.line and c.mb_rev=a.version and b.jobpn=a.jobpn  and jobgroup in " + joppn;
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void QSMSDeleteME_BOMdel(string JobPN, string Line, string StrBU, string UID, string jobgroup, string WO, string Rev, string Machine,
            string Side, string CboComp, string slot, string MBPN)
        {
            string strSQL = "delete  FROM QSMS_MEBOM where JobGroup ='" + jobgroup.Trim()+ "'" ;
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Delete ME_BOM','" + Line + "+" + JobPN + "+" + jobgroup + "+" + Rev + "+" + Machine + "+" + Side + "+" + CboComp + "+" + slot + "','" + UID + "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public void QSMSDeleteME_BOMall(string JobPN, string Line, string StrBU, string UID, string jobgroup, string WO, string Rev, string Machine, 
            string Side, string CboComp, string slot, string MBPN)
        
        {
            string strSQL = "delete  FROM QSMS_MEBOM where (jobpn ='" + MBPN + "' or jobpn in (select jobpn from qsms_jobbom where mbpn='" + MBPN + "')) and JobGroup in "+ jobgroup + "  and version='" + Rev + "' and Machine like '" + Machine + "' and line like '" +Line+ "%' ";
             SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
       
        public DataTable QSMS_MEBOM(string MBPN, string jobgroup, string Rev, string Machine, string Line)
        {
            string strSQL = " select* FROM QSMS_MEBOM where (jobpn = '" + MBPN + "' or jobpn in (select jobpn from qsms_jobbom where mbpn = '" + MBPN + "'))  and JobGroup in  " + jobgroup + " and version = '" + Rev + "' and Machine like '" + Machine + "' and line like '" +Line+ "%'";
             return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable WO_MultiLine(string WO)
        {
            string strSQL = "select a.Line from WO_MultiLine a, Sap_Wo_List b where a.WO=b.WO and B.[Group] in(select [Group] from Sap_Wo_List where BuildType='4' and WO='" +WO+"')";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void insertQSMS_Log(string MBPN, string Rev, string Machine, string OPID)
        {
            string strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Delete ME_BOM','" +MBPN+ "+" +Rev+ "+" + Machine+ "','" +OPID+"',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void insertQSMS_Logbyline(string Line, string OPID)
        {
            string strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Delete ME_BOM By Line','" +Line+ "','" + OPID + "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void delQSMS_MEBOMBYLine(string Line, string site)
        {
            string strSQL = "";
            if(site=="NB3")
            {
                strSQL = "delete  FROM QSMS_MEBOM where line='" + Line + "' AND Machine not LIKE '%Others%'";
            }
            else
            {
                strSQL = "delete  FROM QSMS_MEBOM where line='" + Line + "'";
            }
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void delQSMS_MEBOMBYLinejob(string Line,string JobPN, string site)
        {
            string strSQL = "";
            if (site == "NB3")
            {
                strSQL = "delete  FROM QSMS_MEBOM where line='" + Line + "' and JobPN='"+ JobPN + "' AND Machine not LIKE '%Others%'";
            }
            else
            {
                strSQL = "delete  FROM QSMS_MEBOM where line='" + Line + "' and JobPN='" + JobPN + "'";
            }
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
    }
}
