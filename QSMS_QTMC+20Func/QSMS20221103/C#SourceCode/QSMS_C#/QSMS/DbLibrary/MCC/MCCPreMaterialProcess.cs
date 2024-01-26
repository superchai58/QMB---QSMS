using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace QSMS.DbLibrary.MCC
{
    public class MCCPreMaterialProcess
    {
        public DataTable GetWOListByGroupID(string groupID,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + groupID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable ChkWo(string wo,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '"+type+"','','','"+wo+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetWoFinishedFlag(string wo,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '"+type+"','','','"+wo+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetGroupID(string sDate,string eDate,string line,string type1,string type2)
        {
            string strSQL = "";
            if (type1 == "release")
            {
                if (type2 == "NB5")
                {
                    strSQL = "Select distinct GroupID from QSMS_WOGroup where WO_TransDateTime between  '"+sDate+"' and '"+eDate+"' and line='"+line+"' and closedflag='N' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )";
                }
                else
                {
                    strSQL = "Select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '"+sDate+"' and '"+eDate+"' and line='"+line+"' and closedflag='N'";
                }
            }
            else
            {
                if (type2 == "NB5")
                {
                    strSQL = "Select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '"+sDate+"' and '"+eDate+"' and line='"+line+"' and closedflag='N' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )";
                }
                else
                {
                    strSQL = "Select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '"+sDate+"' and '"+eDate+"' and line='"+line+"' and closedflag='N'";
                }
            }
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetWoArray(string wo,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + wo + "'";
            return SqlHelper.ExecuteTable(strSQL,Parameter.ConnQSMS);
        }

        public DataTable GetCompPN(string group,string customer,string model,string woList,string type)
        {
            string strSql = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + group + "','"+woList+"','"+customer+"','"+model+"'";
            return SqlHelper.ExecuteTable(strSql, Parameter.ConnQSMS);
        }

        public DataTable GetDispatch(string group,string woList,string machine,string jobGroup,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + group + "','"+woList+"','"+machine+"','"+jobGroup+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public void UpdateDispatchFlag(string machine,string woStr)
        {
            string strSQL = "Exec UpdateDispatchFlag '"+machine+"','"+woStr+"'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetWoByGroup(string group,string type)
        {
            string strSql = "exec QSMS_PD_QueryDataByType '" + type + "','','','"+group+"'";
            return SqlHelper.ExecuteTable(strSql, Parameter.ConnQSMS);
        }

        public DataTable GetMachineFlag(string woList,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + woList + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetGroupIdByWO(string wo,string type)
        {
            string strSql = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + wo + "'";
            return SqlHelper.ExecuteTable(strSql, Parameter.ConnQSMS);
        }

        public DataTable GetJobGroupByWoAndMachine(string wo, string machine,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + wo + "','"+machine+"'";
            return SqlHelper.ExecuteTable(strSQL,Parameter.ConnQSMS);
        }

        public DataTable GetGroupByWO(string wo, string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + wo + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetWOByGroupAndWO(string wo, string group, string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + wo + "','"+group+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetWOInfoByWO(string wo,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + wo + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetTotalQtyByWO(string wo, string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + wo + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetCustomerByPN(string pn, string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + pn + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetWithOutQty(string group,string machine,string compPN,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + group + "','"+machine+"','"+compPN+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable ChkDIDCompPN(string DID, string compPN,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + DID + "','" + compPN + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetCodeByDID(string DID, string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + DID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetNonAVLCode(string customer,string compPN,string model,string vendorCode,string dateCode,string lotCode,string wo,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + customer + "','"+compPN+"','"+model+"','"+vendorCode+"','"+dateCode+"','"+lotCode+"','"+wo+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetAVLCustomerByCustomer(string customer, string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + customer + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetAVLInfo(string AVLCustomer, string model, string compPN, string VC, string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + AVLCustomer + "','"+model+"','"+compPN+"','"+VC+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetVCByModelAndCompPN(string model, string compPN, string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + model + "','"+compPN+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetUsedFlagByDID(string DID, string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + DID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetWOInfoByGroupIDAndWO(string groupID, string wo, string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + groupID + "','"+wo+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetIPQCFlag(string DID, string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + DID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetXLInfo(string compPN, string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + compPN + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable PrepairMaterial(string woGroup, string woArray, string compPN, string jobPN, string machine, string VC)
        {
            string strSQL = "exec QSMS_PrepairMaterial '"+woGroup+"','"+woArray+"','"+compPN+"','"+jobPN+"','"+machine+"','"+VC+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetDIDDispatchInfoByDIDAndWO(string WO, string DID,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + WO + "','"+DID+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable InsertDispatch(string WO, string groupID, string line, string woQty, string jobPN, string machine, string compPN, string slot, string lR, string baseQty, string totalNeedQty, string DID, string DIDTotalQty, int tempDIDQty, string VC, string DC, string LC, string g_userName, string transDateTime)
        {
            string strSQL = "exec QSMSInsertDispatch '" + WO + "','"+groupID+"','"+line+"','" + woQty + "','" + jobPN + "','"+machine+"','"+compPN+"'"+slot+"','"+lR+"'"+baseQty+"','"+totalNeedQty+"'"+DID+"','"+DIDTotalQty+"'"+tempDIDQty+"','"+VC+"'"+DC+"','"+LC+"'"+g_userName+"','"+transDateTime+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public void RecordDispatchFDT(string WO)
        {
            string strSQL = "exec RecordDispatchFDT '"+WO+"'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetWOByWO(string WO,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + WO + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public void UpdateDispatchFlagByWO1(string WO)
        {
            string strSQL = "Update QSMS_WOGroup set DispatchFlag='N' Where Work_Order='"+WO+"'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public void UpdateDispatchFlagByWO2(string WO,string transDate)
        {
            string strSQL = "Update QSMS_WOGroup set DispatchFlag='1' Where Work_Order='" + WO + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            strSQL = "Update Sap_WO_List set DispatchOKDateTime='"+transDate+"' where WO='"+WO+"'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetLine(string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        
        public DataTable GetUseFlagByDID(string DID,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + DID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetDispatchInfo(string DID,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + DID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetDIDLog(string DID,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + DID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetDIDQtyByDIDMachine(string sapWOGroup, string woArray, string machine, string DID,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + sapWOGroup + "','"+woArray+"','"+machine+"','"+DID+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public void GetDIDFromSourceBU(string DID)
        {
            string strSQL = "EXEC QSMS_GetDIDFromSourceBU @DID='"+DID+"'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetDIDInfo(string DID,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','"+DID+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetQty(string group, string woArray, string jobPN, string machine,string compPN, string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + group + "','"+woArray+"','"+jobPN+"','"+machine+"','"+compPN+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetDispatch1(string WO, string woArray, string machine, string compPN, string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + WO + "','"+woArray+"','"+machine+"','"+compPN+"'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetUsedFlag(string group,string woArray,string machine,string DID,string type)
        {
            string strSQL = "exec QSMS_PD_QueryDataByType '" + type + "','','','" + group + "','" + woArray + "','" + machine + "','" + DID + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetDispatchInfoAll(string DID, string WO, string type)
        {
            string strSQL = "";
            if (type != "Checked" && type != "MCC_GetDIDDispatchInfo")   //Aris增加Type
            {
                if (WO != "")
                {
                    strSQL = "Select * from QSMS_Dispatch where CompPN='" + DID + "'  and work_order='" + WO + "'";
                }
                else
                {
                    strSQL = "Select * from QSMS_Dispatch where CompPN='" + DID + "'";
                }
            }
            else if (type == "MCC_GetDIDDispatchInfo")    //Aris增加Type
            {
                strSQL = "select Work_Order,Line,JobPN,Machine,Slot,LR,NeedQty,DIDQty,Side,DeletedFlag from QSMS_Dispatch where DID='" + DID + "' and work_order='" + WO + "'";
            }                                             //
            else
            {
                strSQL = "Select a.Work_Order,a.GroupID,a.Line,a.WoQty,a.JobPN,a.Machine,a.CompPN,a.Slot,a.LR,a.BaseQTY,a.NeedQty,a.DID,a.TotalQty,a.DIDQty,a.VendorCode,a.DateCode,a.LotCode,a.UID,a.TransDateTime,a.DIDDateTime,a.DeletedFlag,a.Inherit_wo,a.JobGroup,a.Side,b.ReturnFlag,b.UID,b.TransDateTime from QSMS_Dispatch a left join QSMS_GroupDID b on a.did=b.did and a.diddatetime=b.diddatetime where a.did='" + DID + "'";
            }
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public DataTable GetDispatchInfoAll1(string WO,string machine,string type)
        {
            string strSQL =  "exec QSMS_PD_QueryDataByType '" + type + "','','','" + WO + "','" + machine + "'";   
            //if (type == "MCC_GetDispatchInfoAll")                
            //{
            //    strSQL = "Select * from QSMS_Dispatch where Work_Order='" + WO + "' and machine like '" + machine + "%' union Select * from [QSMS_History].dbo.QSMS_Dispatch where Work_Order='" + WO + "' and machine like '" + machine + "%'";
            //} 
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetSlotInfo(string WO, string Machine)              // Aris 增加
        {
            string strSQL = "select distinct Slot from QSMS_WO where work_order='" + WO + "'  and Machine = '" + Machine + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable TransferDispatchDIDInfo(string Machine, string Slot, string LR, string WO, string CompPN, string JobPN, string DID, string DispatchQty, string NewMachine, string NewSlot, string NewLR, string Version)
        {
            string strSQL = "exec QSMS_TransferDispatchDID '" + Machine + "','" + Slot + "','" + LR + "','" + WO + "','" + CompPN + "','" + JobPN + "','" + DID + "','" + DispatchQty + "','" + NewMachine + "','" + NewSlot + "','" + NewLR + "','" + Version + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable UpdateMachineFlagByWOInfo(string WO,string Machine, string type)
        {
            string strSQL = "";
            if (type == "MachineFlag")
            {
                strSQL = "Select Distinct Machine from QSMS_Wo where Work_Order='" + WO + "'order by machine";
            }
            else if (type == "MachineFlag1")
            {
                strSQL = "Select distinct Machine  from QSMS_WO where work_order='" + WO + "' and Machine='" + Machine + "' and BalanceQty<0";
            }
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            
        }
        public void UpdateMachineFlagByWOInfo1(string WO, string Machine)
        {
            string strSQL = "Update QSMS_WO set MachineFinishedFlag='Y' where work_order='" + WO + "' and Machine='" +Machine+"'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void UpdateMachineFlagByWOInfo2(string WO, string Machine)
        {
            string strSQL = "Update QSMS_WO set MachineFinishedFlag='N' where work_order='" + WO + "' and Machine='" +Machine+"'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable UpdateMachineFlagByWOInfoAll(string WO)
        {
            string strSQL = "Select distinct Machine from QSMS_Wo where work_Order='" + WO + "' and MachineFinishedFlag='N'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void UpdateMachineFlagByWOInfoAll1(string WO)
        {
            string strSQL = "Update QSMS_Wo set WoFinishedFlag='Y' where work_Order='" + WO + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void UpdateMachineFlagByWOInfoAll2(string WO)
        {
            string strSQL = "Update QSMS_Wo set WoFinishedFlag='N' where work_Order='" + WO + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

    }
}
