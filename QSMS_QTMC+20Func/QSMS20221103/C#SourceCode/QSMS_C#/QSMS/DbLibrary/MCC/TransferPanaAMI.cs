using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace QSMS.DbLibrary.MCC
{
    public class TransferPanaAMI
    {
        public DataTable QSMS_GetPCBWO(string strJobPN,string strRev, string strLine, string strBuildType)
        {
           
            string strSQL = "Exec QSMS_GetPCBWO '"+strJobPN+ "','" +strRev+ "','" +strLine+ "','" +strBuildType+ "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }       
        public DataSet GetCheckBomData(string Work_Order, string g_userName, string DualModel,string flag)
        {
            string strSQL = "Exec GetCheckBomData '" + Work_Order + "','" + g_userName + "','" + DualModel + "','"+ flag +"'" ;
            return SqlHelper.ExecuteDataSet(strSQL, Parameter.ConnQSMS);
        }
        public DataTable CheckFormat(string PartNumber)
        {
            string strSQL = "Exec CheckFormat 'PARTNUMBER','" + PartNumber + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void insertQSMS_Log(string BomFile, string g_userName)
        {
            string strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_PanaAMI Start','" + BomFile +"','" + g_userName +"',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))";
           SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable QSMS_JobSide(string JobPN)
        {
            string strSQL = "select * from QSMS_JobSide where JobPN='" + JobPN + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable  Getmachine(string machine, string line, string strside)
        {
            string strSQL = "select machine from machine where machine='" + machine + "' and line ='" + line + "' and side='" + strside + "'";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void deleteMIx(string ChkEQProgram,string strFullJobGroup,string jobgroup, string Machine, string BrdPN, string Rev, string BuildType, string Factory, string Slot, string Line)
        {
            string strSQL = "delete from QSMS_MEBom where JobGroup='"+jobgroup + "' and Machine='"+Machine+"' and JobPN='" +BrdPN + "' and version='"+Rev+ "' And BuildType = '"+BuildType+ "' and Factory='" +Factory+ "'and Slot like '"+Slot+"%' and Line='" + Line+ "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            if(ChkEQProgram=="Y")
            {
                strSQL = "delete from QSMS_MEBom_EQProgram where FullJobGroup='" + strFullJobGroup + "' and Machine='" + Machine + "' and JobPN='" + BrdPN + "' and version='" + Rev + "' And BuildType = '" + BuildType + "' and Factory='" + Factory + "' and Slot like '" + Slot + "%' and Line='" + Line + "'";
                SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
            
        }
        public void deleteData(string ChkEQProgram, string strFullJobGroup, string jobgroup, string Machine, string BrdPN, string Rev, string BuildType, string Factory, string Slot, string Line)
        {
            string strSQL = "delete from QSMS_MEBom where JobGroup='" + jobgroup+ "' and Machine='" +Machine +"' and JobPN='"+BrdPN + "' and version='"+Rev +"' And BuildType = '"+ BuildType+ "' and Factory='" +Factory+ "' and Line='" +Line+"'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            if (ChkEQProgram == "Y")
            {
                strSQL = "delete from QSMS_MEBom_EQProgram where FullJobGroup='" + strFullJobGroup + "' and Machine='" + Machine + "' and JobPN='" + BrdPN + "' and version='" + Rev + "' And BuildType = '" + BuildType + "' and Factory='" + Factory + "' and Line='" + Line + "'";
                SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }


        }
        public void insertQSMS_MEBom(string ChkEQProgram, string Machine, string BrdPN,string jobgroup,
                       string Rev, string compPN,int Side, string Slot, int Qty, string BuildType,string strSide,
                        string UID, string Factory, string Line, string ReelWidth, string location, string DualLaneMode, string strFullJobGroup)
        {
            string strSQL ="Insert into QSMS_MEBom(Machine, JobPN,JobGroup,Version,CompPN,LR,Slot,Qty,BuildType,Side,UID,Factory,Line,ReelWidth,Location,DualLaneMode) values('"
                +Machine.Trim()+ "','" +BrdPN.Trim() + "','" + jobgroup.Trim() + "','" +Rev.Trim() + "','" +compPN.Trim() + "','" +Side + "','"+Slot.Trim() + "','" +Qty 
                + "','" +BuildType.Trim() + "','" +strSide.Trim() + "','" +UID.Trim() + "','" + Factory.Trim() + "','" +Line.Trim() + "','" + ReelWidth.Trim() + "','" +location.Trim() + "','" +DualLaneMode.Trim() + "')" ;
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            if (ChkEQProgram == "Y")
            {
                strSQL = "Insert into QSMS_MEBom_EQProgram(Machine, JobPN,JobGroup,FullJobGroup,Version,CompPN,LR,Slot,Qty,BuildType,Side,UID,Factory,Line,ReelWidth,Location,DualLaneMode) values('"
                 + Machine.Trim() + "','" + BrdPN.Trim() + "','" + jobgroup.Trim() + "','"+ strFullJobGroup + "','" + Rev.Trim() + "','" + compPN.Trim() + "','" + Side + "','" + Slot.Trim() + "','" + Qty
                 + "','" + BuildType.Trim() + "','" + strSide.Trim() + "','" + UID.Trim() + "','" + Factory.Trim() + "','" + Line.Trim() + "','" + ReelWidth.Trim() + "','" + location.Trim() + "','" + DualLaneMode.Trim() + "')";
                SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
            }
        }
        public void DelNozzleLocation(string Factory, string Line, string strSide, string BuildType, string jobgroup)
        {
            string strSQL = "Delete from NozzleLocation Where Factory='"+ Factory + "' and Line='" +Line+ "' and Side='"+strSide+ "' and BuildType='" +BuildType+ "' and JobGroup='" +jobgroup+ "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void InsertNozzleLocation(string Factory, string Line, string strSide, string BuildType, string jobgroup,string Machine,string location,string NozzleType,string g_userName)
        {
            string strSQL = "Insert into NozzleLocation(Factory,Line,Side,BuildType,JobGroup,Machine,Location,NozzleType,UID,TransDateTime)Values('"+Factory+"','" +Line+"','" +strSide+"','"+BuildType+ "','"+jobgroup+"','"+Machine+"','"+location+ "','"+NozzleType+ "','" +g_userName+"',dbo.formatdate(Getdate(),'YYYYMMDDHHNNSS'))";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }

        public void insertQSMS_LOG(string Type, string  jobgroup, string PCBSize, string g_userName)
        {
            string strSQL = "Insert into QSMS_LOG(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) Values('"+ Type + "','" + jobgroup + "','" +PCBSize+ "','" +g_userName+"',0,dbo.formatdate(Getdate(),'YYYYMMDDHHNNSS'))";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
    }
}
