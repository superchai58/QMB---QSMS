using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data;


namespace QSMS.DbLibrary.Report
{
   public  class ReportProcess : QMSSDK.Db.WinForm
    {
        public DataTable B_ToolTip_Config()
        {
            string strSQL = "Select [Key],Value from B_ToolTip_Config order by [Key]";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable Program_DefineItem()
        {
            string strSQL = "select distinct value from Program_DefineItem where AppName='QSMS' and FuncType='Report' and item='ReportType'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnSMT);
        }

        public DataTable Line()
        {
            string strSQL = "select distinct Line from SAP_WO_List order by line";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable ListWoall(string GroupID,string TempJobPn)
        {
            string strSQL = "select distinct a.Work_Order from QSMS_WOGroup a,QSMS_JobBOM b  where a.GroupID= '" + GroupID + "' and a.work_Order=b.work_order and b.jobpn like '" + TempJobPn + "%'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetWO(string WO)
        {

            string strSQL = "Select WO from sap_wo_list where WO='"+ WO + "' And status >= 10";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetJobpn(string BeginDate, string EndDate, string line, bool radio)
        {
            string strSQL = "";
            if (radio==true)
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
        public DataTable GetComp(string Machine, string MBPN, string WO, string Line)
        {
            string strSQL = "";
            if (Machine.Trim().ToUpper()!="ALL")
            {
                strSQL = "select CompPN from QSMS_WO  where Work_Order='" +WO+"'  and Machine='"+Machine+ "'";
            }
            else
            {
                strSQL = "select CompPN from QSMS_WO  where Work_Order='" +WO+ "'";
            }
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetJobGroupByJobRev(string Machine, string MBPN, string Rev)
        {
            string strSQL = "";
            if (Machine.Trim().Length >6)
            {
                strSQL = "Select distinct jobgroup from qsms_mebom where (jobpn='" + MBPN + "' or jobpn in (select distinct jobpn from QSMS_JObBom where MBPN='"+ MBPN+ "')) and version='" + Rev+ "'" + " and Machine like '" + Machine + "%' ";
            }
            else
            {
                strSQL = "Select distinct jobgroup from qsms_mebom where (jobpn='" + MBPN + "' or jobpn in (select distinct jobpn from QSMS_JObBom where MBPN='" + MBPN + "')) and version='" + Rev + "'";
            }
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetWoinfo(string WO)
        {

            string strSQL = "select PN, Qty ,MB_Rev,Line,BuildType from sap_wo_list where WO='" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetCustomer(string MBPN,string Rev)
        {

            string strSQL = "select Customer from ModelName where modelname='" + MBPN + "-"+ Rev + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetjobPn(string WO)
        { 

            string strSQL = "select jobPn from QSMS_JobBOM where work_Order='" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
      
        public DataTable GetMachine(string WO)
        {

            string strSQL = "select distinct Machine,MachinefinishedFlag from QSMS_WO where Work_Order= '" + WO.Trim() + "' ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetGroupID(string JobPN,string BU,string BeginDate,string  EndDate, string Line, bool OptRelease)
        {
            string TempJobPn = "";
            
            string strSQL="";
            if(JobPN=="")
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
                if(JobPN.IndexOf("-")>0)
                {
                    TempJobPn = JobPN.Substring(1, 11);
                    
                }
                else
                {
                    TempJobPn = JobPN;
                }
                if(OptRelease==true)
                {
                    strSQL = "select distinct GroupID from QSMS_WOGroup a,QSMS_JobBOM b   where a.WO_TransDateTime between  '" + BeginDate + "' and '" + EndDate + "' and a.line='" + Line + "' and a.work_order=b.work_order and b.jobpn='" + TempJobPn +"'and closedflag='N' ";

                }
                else
                {
                    strSQL = "select distinct GroupID from QSMS_WOGroup a,QSMS_JobBOM b   where a.WO_TransDateTime between  '" + BeginDate + "' and '" + EndDate + "' and a.line='" + Line + "' and a.work_order=b.work_order and b.jobpn='" + TempJobPn + "'and closedflag='N' ";
                }
            }

            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetjobPn(string BU,string Line, string BeginDate, string EndDate,bool optRelease)
        {
           string strSQL = "";
           if (optRelease==true)
            {
                strSQL = "select distinct JobPN from QSMS_WOGroup a, QSMS_JobBOM b where a.WO_TransDateTime between  '" + BeginDate + "' and '" + EndDate + "' and a.line='" + Line + "' and A.work_order=b.work_order";

            }
            else
            {
                if(BU=="PO")
                {
                    strSQL = " select distinct JobPN from QSMS_WOGroup a,QSMS_JobBOM b where substring(GroupID,2,8) between '" + BeginDate + "' and '" + EndDate + "'  and a.line='" + Line + "' and a.work_Order=b.work_Order";

                }
                else
                {
                    strSQL = " select distinct JobPN from QSMS_WOGroup a,QSMS_JobBOM b where substring(GroupID,4,8) between '" + BeginDate + "' and '" + EndDate + "'  and a.line='" + Line + "' and a.work_Order=b.work_Order";

                }
            }
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMSRptPrepareMaterial(string WO, string Machine)
        {
            string strSQL = "exec QSMSRptPrepareMaterial '"+ WO + "','" + Machine+ "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetWork_Order(string GroupID)
        {

            string strSQL = "select Work_Order from QSMS_WoGroup where GroupID='" + GroupID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);

        }
        public DataTable XL_WOPlanSeq(string Shift, string line, string sDateTime, string eDateTime)
        {
            string strSQL = "select distinct Wo from  XL_WOPlanSeq where shift = '" + Shift + "' and Line = '" + line + "' and BeginDateTime between '" + sDateTime + "' and '" + eDateTime + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataSet QSMSRptPrepareMaterialByWO(string WO)
        {
            string strSQL = "exec QSMSRptPrepareMaterialByWO '" + WO + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
        public DataSet QSMSRptPrepareMaterialByLineShift(string WO)
        {
            string strSQL = "exec QSMSRptPrepareMaterialByLineShift '" + WO + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
        public DataSet QSMSRptPrepareMaterialByGroup(string WO)
        {
            string strSQL = "exec QSMSRptPrepareMaterialByGroup '" + WO + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
        public DataSet QSMSRptPrepareMaterialByJobPN(string WO, string Jobpn)
        {
            string strSQL = "exec QSMSRptPrepareMaterialByJobPN '" + WO + "','" + Jobpn + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }

        public void DelQSMS_CompPNcheck_Temp()
        {
            string strSQL = "Truncate table QSMS_CompPNcheck_Temp";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
       
        public DataTable QSMS_CheckCompPN(string strJobPN, string strCompPN, string strUID,string flag)
        {
            string strSQL = "EXEC QSMS_CheckCompPN '" + strJobPN + "','" + strCompPN + "','" + strUID + "','"+ flag+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable Getsapbom( string strCompPN)
        {
            string strSQL = "select top 1 * from sapbom where MBPN='" + strCompPN + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void QSMS_CompPNcheck_Temp(string strJobPN, string strCompPN, string UserID)
        {
            string strSQL = "insert into QSMS_CompPNcheck_Temp(JobPN, CompPN, UserID, TransDateTime) select '"+ strJobPN + "','" + strCompPN + "','"+ UserID + "',''";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);

        }
        public DataTable QSMS_QueryReplacePN()
        {
            string strSQL = "EXEC QSMS_QueryReplacePN";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GenChangeLineReport2(string dtpSDate, string dtpEDate)
        {
            string strSQL = "Exec GenChangeLineReport2 '"+ dtpSDate + "','"+ dtpEDate + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable DIDDeleteRecords(string dtpSDate, string dtpEDate)
        {
            string strSQL = "Exec QueryDIDUsed '" + dtpSDate + "','" + dtpEDate + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable LineChangeStatistics(string dtpSDate, string dtpEDate)
        {
            string strSQL = "Exec GenChangeLineReport1 '" + dtpSDate + "','" + dtpEDate + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable XL_MonitorReport(string dtpSDate)
        {
            string strSQL = "Exec XL_MonitorReport '" + dtpSDate+ "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_DispatchQTYByWO(string strWO)
        {
            string strSQL = " EXEC QSMS_DispatchQTYByWO '','','" + strWO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_DID_ToWH(string dtpSDate, string dtpEDate)
        {
            string strSQL = "select substring(TransDateTime,5,2)+'/'+substring(TransDateTime,7,2)as [Date],count(*) as TotalNumber from QSMS_DID_ToWH where ToWHType='Return' and isGood='Y' and WareHouseID<>''  "
                + "and TransDateTime Between '" + dtpSDate + "' and '" + dtpEDate + "' group by substring(TransDateTime,5,2)+'/'+substring(TransDateTime,7,2)" +
                "order by substring(TransDateTime,5,2)+'/'+substring(TransDateTime,7,2) ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMSRpt_QSMS_WO(string strWO,string flag)
        {
            string strSQL = "exec QSMSRpt_QSMS_WO '" + strWO + "','"+flag+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataSet CopyToExcelWipByMaterial(string CompPN)
        {
            string strSQL = "exec QSMSRptWipByMaterial '" + CompPN + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
        public DataSet QSMSRptWipBydate()
        {
            string strSQL = "exec QSMSRptWipBydate ";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
        public DataSet QSMSRptWipByGroup(string GroupID)
        {
            string strSQL = "exec QSMSRptWipByGroup '" + GroupID + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
        public DataSet QSMSRptLackCompByWo(string Work_Order)
        {
            string strSQL = "exec QSMSRptLackCompByWo '" + Work_Order + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMSRptWipDifferentMaterial(string Work_Order)
        {
            string strSQL = "exec QSMSRptWipDifferentMaterial '" + Work_Order + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable qsms_verify(string line,string CompPn)
        {
            string strSQL = "select * from qsms_verify where machine like '" + line + "%' and comppn='" + CompPn + "' Order by Machine,Slot,LR,BegindateTime Desc";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_QuerySapBom(string WO)
        {
            string strSQL = "exec QSMS_QuerySapBom '" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable sap_bom(string WO)
        {
            string strSQL = "select * from sap_bom where work_Order='"+ WO + "' order by CompPN,Item,CompLevel";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetSapGroupByWo(string WO)
        {
            string strSQL = "select * from sap_wo_list  where [group] in (select [group] from sap_wo_list where wo='" + WO + "') order by wo";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable sap_wo_list(string WO)
        {
            string strSQL = "Select [Group],Line,BuildType from sap_wo_list where wo='" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable sapwolist(string WO)
        {
            string strSQL = "select * from Sap_Wo_List where BuildType='1' and WO='" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_GetEMSFlag(string Group)
        {
            string strSQL = "exec QSMS_GetEMSFlag '"+ Group + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_GetQSMS_JobBomMB(string Group)
        {
            string strSQL = "Select a.JobPN,b.MB_Rev from QSMS_JobBom a,Sap_Wo_List b where b.[Group]='" + Group + "' and a.Work_Order=b.WO and b.PN like '%MB%'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_GetQSMS_JobBom(string Group)
        {
            string strSQL = "Select a.JobPN,b.MB_Rev from QSMS_JobBom a,Sap_Wo_List b where b.[Group]='" + Group + "' and a.Work_Order=b.WO";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_GetQSMS_JobBomEMS(string Group)
        {
            string strSQL = "Select a.JobPN,b.MB_Rev from QSMS_JobBOM a,Sap_Wo_List B where b.[Group]='" + Group + "' and a.work_order=b.wo and b.InitAOIFlag='Y'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable WO_MultiLine(string sInputStr)
        {
            string strSQL = "select rtrim(Line)+ltrim(Side) as Line from WO_MultiLine where WO='"+ sInputStr+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable WO_MultiLineWO(string WO)
        {
            string strSQL = "select a.Line from WO_MultiLine a, Sap_Wo_List b where a.WO=b.WO and B.[Group] in(select [Group] from Sap_Wo_List where BuildType='4' and WO='"+WO+ "')";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        
        public DataTable GetMEBom(string line1,string line2,string Rev,string WO,string TempJObGroup)
        {
            string strSQL = "";
            if (WO != "")
            {
                if (line2 != "")
                {
                    strSQL = "select a.Machine,a.JobGroup,a.JObpN,a.Version,A.CompPN,A.lR,''''+ a.Slot as Slot,a.Qty,a.BuildType,a.Side,a.Factory,a.UID,a.TransDateTime,a.Location from QSMS_MeBom a where  a.version like '" + Rev + "%' " +
                " and (a.line='" + line1 + "'or a.line='" + line2 + "') and a.jobpn in (select distinct jobpn from qsms_Jobbom where Work_Order='" + WO + "') and jobgroup in " + TempJObGroup + " order by a.JobpN,a.machine,CompPN,Slot";
                }
                else

                {
                    strSQL = "select a.Machine,a.JobGroup,a.JObpN,a.Version,A.CompPN,A.lR,''''+ a.Slot as Slot,a.Qty,a.BuildType,a.Side,a.Factory,a.UID,a.TransDateTime,a.Location from QSMS_MeBom a where  a.version like '" + Rev + "%' " +
                " and a.line='" + line1 + "' and a.jobpn in (select distinct jobpn from qsms_Jobbom where Work_Order='" + WO + "') and jobgroup in " + TempJObGroup + " order by a.JobpN,a.machine,CompPN,Slot";
                }
            }
            else
            {
                strSQL = "select a.Machine,a.JobGroup,a.JObpN,a.Version,A.CompPN,A.lR,''''+ a.Slot as Slot,a.Qty,a.BuildType,a.Side,a.Factory,a.UID,a.TransDateTime,a.Location from QSMS_MeBom a where a.version='" + Rev + "' " +
        " and a.line='" + line1 + "'  and jobgroup in " + TempJObGroup + " order by a.JobpN,a.machine,CompPN,Slot";

            }
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable CompPN_DataCurrent(string Work_Order)
        {
            string strSQL = "Select b.* from CompPN_Data a,QSMS_WO b with(nolock) where b.work_order='" +Work_Order+ "'and a.CompPN=b.CompPN  and a.type='WastagePN' order by jobpn,machine,slot,lr,comppn";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable CompPN_DataHis(string Work_Order)
        {
            string strSQL = "Select b.* from CompPN_Data a,QSMS_History..QSMS_WO b with(nolock) where b.work_order='" +Work_Order+ "' and a.CompPN=b.CompPN and a.type='WastagePN' order by jobpn,machine,slot,lr,comppn";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
 
        public DataSet GetGroupIDDataByCompPN(string GroupID, string compPN)
        {
            string strSQL = "exec GetGroupIDDataByCompPN '" + GroupID + "', '" + compPN + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_Log(string DID)
        {
            string strSQL = "Select * from QSMS_Log where DID like '%" +DID+ "%' AND Event_No='Delete Me_BOM' order by Trans_Date desc";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_WOCurrent(string Work_Order)
        {
            string strSQL = "Select * from QSMS_WO with(nolock) where work_order='" +Work_Order+ "' order by jobpn,machine,slot,lr,comppn";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable QSMS_WOHis(string Work_Order)
        {
            string strSQL = "Select * from QSMS_History..QSMS_WO with(nolock) where work_order='" + Work_Order + "' order by jobpn,machine,slot,lr,comppn";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_GetReplacePNByWOList(string WO)
        {
            string strSQL = "exec QSMS_GetReplacePNByWOList '" + WO .Trim()+ "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_CheckBOM_CheckCycleTime(string WO)
        {
            string strSQL = "exec QSMS_CheckBOM_CheckCycleTime '" + WO.Trim() + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable qsms_wogroup(string WO)
        {
            string strSQL = "select * from qsms_wogroup where work_order='" +WO+ "' and ClosedFlag='Y'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataSet GetCheckBomData(string Work_Order,string g_userName,string DualModel,string flag,string StartDate,string EndDate)
        {
            string strSQL = "exec GetCheckBomData  '" + Work_Order.Trim() + "','"+ g_userName + "','"+ DualModel + "','"+ flag + "','"+ StartDate + "','"+ EndDate +"'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }

        public DataSet CheckQSMSWO(string Work_Order)
        {
            string strSQL = "Select * from QSMS_WO with(nolock) where work_order='" + Work_Order + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }

        public DataSet GetSAPBOMFailInfo(string Work_Order)
        {
            string strSQL = "select *  from Sap_BOM_Fail with(nolock) where work_order='" + Work_Order + "'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }

        public DataSet CheckBom(string Work_Order)
        {
            string strSQL = "exec QSMS_CheckBomSP '" + Work_Order + "','Y'";
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }

        public DataTable QSMS_Wo_Diff(string WO)
        {
            string strSQL = "Select * from QSMS_Wo_Diff where Work_Order='" +WO+ "' ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable Sap_Return(string Work_Order, string CboComp, string CboGroupID, string Report_Type, string sDateTime, string eDateTime)
        {
            string strSQL = "exec Sap_Return  '" + Work_Order.Trim() + "','" + CboComp + "','" + CboGroupID + "','" + Report_Type + "','" + sDateTime + "','" + eDateTime + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable ReturnDID(string CboReportType,string CboGroupID)
        {
            string strSQL = "";
            if(CboReportType== "ReturnDIDByGroupID")
            {
                strSQL = "exec XL_ReturnDIDByGroupID '" + CboGroupID + "'";
            }
            else if (CboReportType == "ReturnDIDByWO")
            {
                strSQL = "exec XL_ReturnDIDByWO  '" + CboGroupID + "'";
            }
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable AOIQtySummary(string Work_Order)
        {
            string strSQL = "exec QSMSGetAOISummary '" + Work_Order + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable AOIDetail(string Work_Order)
        {
            string strSQL = "select * from QSMS_AOI where wo= '"+ Work_Order+ "' order by station,transdatetime";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable MachineType(string flag,string compPN)
        {
            string strSQL = "";
            if(flag== "MachineType")
            {
                strSQL = " select Vendor, Line,SeqIDByLine,Side , Factory,Machine,Unit,Qty, MaxSlotNum, LR, MappingID,FujiData,DIOCircuit from Machine";

            }
            else if(flag== "TraySlot")
            {
                strSQL = "select * from TraySlot";
            }
            else if (flag == "SpliceReplacePN")
            {
                strSQL = "Select * from QSMS_Log where Event_No like '%SpliceReplacePN%' order by Trans_Date desc";
            }
            else if (flag == "SplicePN")
            {
                strSQL = "Select * from QSMS_Log with(nolock) where Event_No like '%SplicePN%' order by Trans_Date desc";
            }
            else if (flag == "MaintainFeeder")
            {
                strSQL = "Select * from QSMS_Log with(nolock) where Event_No='MaintainFeeder' order by Trans_Date desc";
            }
            else if (flag == "XL_ReelBaseQty")
            {
                strSQL = "select A.Plant,A.CompPN,A.BaseReelQty,B.Location,B.StockQty,B.Status,B.WorkDate,B.Shift,B.TransDateTime from XL_ReelBaseQty A, XL_StockQtyByLocation B where A.comppn=B.comppn AND A.comppn='" +compPN+ "' order by B.transdatetime desc ";
            }
            else if (flag == "NonAVL")
            {
                strSQL = "Select * from QSMS_NonAVL where CompPN like '" + compPN + "%'";
            }
            else if (flag == "FUJI_AVLList")
            {
                strSQL = "Select * from QSMS_FujiAVL Where [group] in(select [Group] from sap_wo_list where wo='"+ compPN + "') order by TransDateTime Desc";
            }
            else if (flag == "CheckBom_Log")
            {
                strSQL = "select DID as Wo, User_Name,Trans_date from qsms_log where did='"+ compPN+ "' and system_name='SMT_QSMS' and Event_No='CheckBom' order by trans_date desc";
            }
            else if (flag == "CheckBom_Result")
            {
                strSQL = "select DID as Wo,user_Name,case when ReturnQty='0' then 'Y' else 'N' end as Result, Trans_Date from qsms_log where did='"+compPN+"' and system_name='SMT_QSMS' and Event_No='CheckBOMResult' order by trans_date desc";
            }
            
            else if (flag == "AllDispatchByGroupID")
            {
                strSQL = "exec QSMS_GetAllDispatchInforByGroupID '"+ compPN + "'";
            }
            
            else if (flag == "CastRate")
            {
                strSQL = "select * from QSMS_CastRate";
            }
            else if (flag == "OneByOne")
            {
                strSQL = "select * from QSMS_OneByOne";
            }
            else if (flag == "SameGroupWO")
            {
                strSQL = "select WO,PN,MB_Rev,Line,QTY,CombineQty,Trans_Date,WO_Type from sap_wo_list where [group] in (select [group] from sap_wo_list where wo= '" + compPN + "')";
            }
            else if (flag == "CompPN_DIDData")
            {
                strSQL = "exec GetCompPNDIDData '"+ compPN + "'";
            }
            else if (flag == "CompPNQty")
            {
                strSQL = "exec GetCompPNQty '" + compPN + "'";
            }
            else if (flag == "UnCloseGroupID")
            {
                strSQL = "select distinct a.GroupID,substring(GroupID, 4, 4) + '/' + substring(GroupID, 8, 2) + '/' + substring(GroupID, 10, 2) as Create_date from qsms_wogroup a,sap_wo_list b where a.work_order = b.wo and a.closedflag <> 'Y' order by create_date desc";
            }
            else if (flag == "ME_BOM_GroupID")
            {
                strSQL = "exec QSMS_MEBOM_ByGroupID '" + compPN + "'";
            }
            else if (flag == "MEBom_EQProgram")
            {
                strSQL = "exec QSMS_MEBOM_ByEQProgram '" + compPN + "'";
            }
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable machine()
        {
            string strSQL = "";
          
            strSQL = "select distinct LEFT(machine, 1) AS Expr1 from machine WHERE machine<>'' order by LEFT(machine, 1)";
           
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GenVerifyReportToExcel(string Line,string Type)
        {
            string strSQL = "";
            if (Type == "VerifyReport")
            {
                strSQL = "Exec GenVerifyReportToExcel  '" + Line + "'";

            }
            else if(Type == "VerifyReportWOChged")
            {
                strSQL = "Exec SP8_VerificationReportWOChged  '" + Line + "'";
            }
            
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GenVerifyFailReport(string SDate,string Edate)
        {
            string strSQL = "";

            strSQL = "Exec GenVerifyFailReport '"+ SDate + "','"+ Edate + "'";

            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetUnDispatchList(string Work_Order, string Type)
        {
            string strSQL = "";
            if (Type == "N")
            {
                strSQL = "Exec QSMS_WOUnDispatch  '" + Work_Order + "'";

            }
            else if (Type == "Y")
            {
                strSQL = "SELECT COUNT(*) FROM QSMS_WO WHERE Work_Order='" + Work_Order + "'";
            }

            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void QSMS_ReCountDispatchQty(string WO,string Type)
        {
            string strSQL = "";
            strSQL = "Exec QSMS_ReCountDispatchQty '" + WO + "','"+ Type + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        
        public DataSet  XL_MaterialDemand(string Sdate,string Edate,string CboComp, string Flag)
        {
            string strSQL = "";
            if(Flag == "XL_MaterialDemand")
            {
                strSQL = "Exec XL_RptMaterialDemand '" + Sdate + "','" + Edate + "'";
            }
            else if (Flag == "DIDIntegration")
            {
                strSQL = "Exec QSMS_DIDIntegration  'Report','"+ CboComp + "','" + Sdate + "','" + Edate + "'";
            }
            else if (Flag == "Glue_DataByDay")
            {
                strSQL = "Exec Glue_DataByDay '" + Sdate + "','" + Edate + "'";
            }
            else if (Flag == "MaterialReturn")
            {
                strSQL = "select BU,  ReferenceID, Status, DID, CompPN, Qty, VendorCode, DateCode, LotCode, OldDID, OldDIDDateTime, ToWHType,SAPClient, InPlant, OutPlant, OutLineMC, WHType, BatchNo, Material_Cost_Center, IsGood, WareHouseID, UID, TransDateTime, GenRefIDDateTime, WHTransDateTime from dbo.QSMS_DID_ToWH where CompPN like '" +CboComp+ "%' and TransDateTime between '" + Sdate + "' and '" + Edate + "' order by ReferenceID,TransDateTime";
            }
            else if (Flag == "PanalnterLock")
            {
                strSQL = "select Machine,FeederID,DID,Slot,LR,JobPN,BeginDatetime as TransDateTime from QSMS_Verify where BeginDatetime between '" + Sdate + "' and '" + Edate + "' and EndDatetime='' order by Machine,Slot,LR";
            }
            else if (Flag == "Glue_CallOff")
            {
                strSQL = "Exec RPTGlue_CallOff '" + Sdate + "','" + Edate + "'";
            }
            else if (Flag == "Glue_Consumption")
            {
                strSQL = "Exec RPTGlue_Consumption '" + Sdate + "','" + Edate + "'";
            }
            else if (Flag == "XL_DemandDetail")
            {
                strSQL = "select w.GroupID, x.* from xl_qsms_wo x join qsms_wogroup w on x.work_order = w.work_order where x.workdate between '" + Sdate + "' and '" + Edate + "' order by x.workdate,x.shift,x.line,x.work_order,x.comppn";
            }
            else if (Flag == "DIDCompare")
            {
                strSQL = "SELECT Line  as Line,case when substring(transdatetime,9, 6)> '080000' and substring(transdatetime,9, 6)< '20000' " +
                         "then 'D' else 'N' end as Shift,Machine,slot+'-'+cast(lr as char(1)) as Slot,DID,NewDID,ScanDID,TransDateTime," +
                         "case CheckResult when 'Y' then N'相同' else N'不同' end as CheckResult,OPID "+
                         "FROM qsms_checkcomplog WHERE Slot<>'' and transdatetime between '" + Sdate + "' and '" + Edate + "'";
            }
            else if (Flag == "CheckSpliceReplacePN")
            {
                strSQL = "Exec QSMS_rptChkSpliceReplacePN '" + Sdate + "','" + Edate + "'";
            }
            else if (Flag == "ForbiddenPN")
            {
                strSQL = "select top 60000 RefID,ModelName,PN,VendorCode,DateCode,LotCode,Status,[User],TransDateTime,LastUpdateTime,Trans_Flag,''as DelDateTime " +
                         "from forbiddenpn where PN like '" +  CboComp + "%' and transdatetime between '" + Sdate + "' and '" + Edate + "' " +
                         "Union All " +
                         "select top 60000 RefID,ModelName,PN,VendorCode,DateCode,LotCode,Status,[User],TransDateTime,LastUpdateTime,Trans_Flag,DelDateTime" +
                         " from forbiddenpn_trace WHERE PN like '" + CboComp + "%' and transdatetime between '" + Sdate + "' and '" + Edate + "' ";
            }
            else if(Flag== "PDA_DistributeDIDLog")
            {
                strSQL = "Select * from PDA_DistributeDIDLog with(nolock) where line like '" +CboComp+ "%' and TransDateTime between '" + Sdate + "' and '" + Edate + "' order by TransDateTime desc";
            }
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
        public DataSet Rpt_XL_GetWoInputPlan(string Sdate, string Edate,string WO,string ReportType)
        {
            string strSQL = "";
            if(ReportType== "WoInputPlan")
            {
                strSQL = "Exec Rpt_XL_GetWoInputPlan '" + Sdate + "','" + Edate + "','" + WO + "'";
            }
            else if(ReportType== "WoInputPlanBySide")
            {
                strSQL = "Exec XL_GetWoInputPlanBySide '" + Sdate + "','" + Edate + "','" + WO + "'";
            }
           
           
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
        public DataTable XL_DispatchStatus(string Line,string strSdate, string strEDate, string strShift )
        {
            string strSQL = "";
            strSQL = "exec XL_DispatchStatus '" + Line + "','" + strSdate + "','"+ strEDate + "','" + strShift + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable MEBom_Model(string tmpJobPN, string tmpRev)
        {
            string strSQL = "";
            strSQL = "select * from qsms_mebom where JobPN='"+tmpJobPN+"' and Version='"+tmpRev+ "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QueryWOData(string WO)
        {
            string strSQL = "select TOP 1 * from sap_wo_list where wo='" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetDiffCompPNInfo(string Group1, string Group2)
        {
            string strSQL = "Exec XL_GetDiffCompPNInfo '" + Group1 + "','" + Group2 + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetSNByComp(string CompPN, string VendorCode, string DateCode, string LotCode, string BeginDateTime, string EndDateTime, string Model, string Type)
        {
            string strSQL = "Exec TraceReport_GetSNByComp '" + CompPN + "','" + VendorCode + "','" + DateCode + "','" + LotCode + "','" + BeginDateTime + "','" + EndDateTime + "','" + Model + "','" + Type + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable TraceReport_GetCompBySN(string One, string SN, string CompPN, string Type)
        {
            string strSQL = "Exec TraceReport_GetCompBySN '" + One + "','" + SN + "','" + CompPN + "','" + Type + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable TraceReport_GetSNByDID(string One, string SN, string Type)
        {
            string strSQL = "Exec TraceReport_GetSNByDID '" + One + "','" + SN + "','" + Type + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable TraceReport_GetCompByWO(string SN, string CompPN, string Type)
        {
            string strSQL = "Exec TraceReport_GetCompByWO '" + SN + "','" + CompPN + "','" + Type + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable TraceReport_DeleteFile(string DeleteFile)
        {
            string strSQL = "Exec TraceReport_DeleteFile '" + DeleteFile + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable TraceReport_TempSN(string SN, string CompPN)
        {
            string strSQL = "insert into TraceReport_TempSN(SN,CompPN,HostName) values ('" + SN + "','" + CompPN + "',host_name())";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QueryTempSN()
        {
            string strSQL = "select distinct SN,CompPN from TraceReport_TempSN where HostName=host_name() order by SN";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable DeleteTempSN()
        {
            string strSQL = "delete from TraceReport_TempSN where HostName=host_name()";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_RegisterCheckBOM(string Work_Order,string Type)
        {
            string strSQL = "Exec QSMS_RegisterCheckBOM  '"+ Work_Order + "','"+ Type + "',''";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void delSap_BOM_Fail(string Work_Order)
        {
            string strSQL = "delete from Sap_BOM_Fail  where Work_Order ='" + Work_Order + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void InsQSMS_LOG(string Work_Order,string Type,string g_userName,string Step)
        {
            string strSQL = "Insert into QSMS_LOG(system_name, event_no, DID, user_name, ReturnQty, trans_date) values('SMT_QSMS', '" + Type + "', '" + Work_Order + "', '" + g_userName + "', '"+ Step + "', dbo.FormatDate(getdate(), 'yyyymmddhhnnss'))";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_CheckBOM(string Work_Order)
        {
            string strSQL = "select 0 from QSMS_CheckBOM where WorkOrder='" + Work_Order + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GETBuildType(string Work_Order)
        {
            string strSQL = "select BuildType from sap_wo_list where Wo='" + Work_Order + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_CheckBomSP(string Work_Order,string BuildType, string DualModel)
        {
            string strSQL = "";
            if (DualModel == "N")
            {
                strSQL = "Exec QSMS_CheckBomSP '" + Work_Order + "','N','" + BuildType + "'";
            }
            else
            {
                strSQL = "Exec QSMS_CheckBomSP_Dual '" + Work_Order + "','N','" + BuildType + "'";
            }
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable Negative(string MBPN)
        {
            string strSQL = "select MBPN from QSMS_NegativeBrd where MBPN='" + MBPN + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable Pilot(string Work_Order)
        {
            string strSQL = "select PN,Pilot from Sap_WO_List where WO='" + Work_Order + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void Updateqsms_error_log(string Work_Order)
        {
            string strSQL = "update qsms_error_log set col1=dbo.formatdate(getdate(),'yyyymmddhhnnss') where subid='" + Work_Order.Trim() + "' and appname='SMT_QSMS' and subfunction='ReplacePN' and col1<>''";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GETSite()
        {
            string strSQL = "Select Site from Site" ;
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void Updateqsms_error_logPASS(string Work_Order)
        {
            string strSQL = "update qsms_error_log set col1=dbo.formatdate(getdate(),'yyyymmddhhnnss') where subid='" + Work_Order.Trim() + "' and appname='SMT_QSMS' and subfunction='ReplacePN' and col1=''";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetSap_BOM_Fail(string Work_Order)
        {
            string strSQL = "select *  from Sap_BOM_Fail  where Work_Order ='"+ Work_Order+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        
    }

    


}
