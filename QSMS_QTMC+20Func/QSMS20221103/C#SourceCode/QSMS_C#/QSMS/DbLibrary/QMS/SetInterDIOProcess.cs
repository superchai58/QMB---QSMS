using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace QSMS.DbLibrary.QMS
{
   public  class SetInterDIOProcess
    {
        public DataTable GetMachine()
        {
            string strSQL = "select * from Machine where DisableInterlock='1' ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetMachine1()
        {
            string strSQL = "select * from Machine ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetSet(string Line)
        {
            string strSQL = "select machine,DisableInterlock from Machine where substring(machine,1,3)='" +Line+ "' ";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public DataTable GetLine()
        {
            string strSQL = "select distinct Line from QSMS_woGroup order by Line";
            return SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);
        }
        public void SetInterDIO(int ClickFlg, string chkAllLineData, string OPID, string chkMachineChecked, string chkMachineText)
        {
            string strSQL = "Exec SetInterDIO "+ ClickFlg + ",'"+ chkAllLineData + "','"+ OPID + "','"+ chkMachineChecked + "','"+ chkMachineText + "'";
            SqlHelper.ExecuteTable(strSQL, Parameter.ConnQSMS);

        }
    }
}
