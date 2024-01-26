using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;

namespace PrinterLib
{
    /// <summary>
    /// 增加替换动作 .Replace("^","_5E") 20220712 Rain
    /// </summary>
    public class PrintLabel 
    {
        public int QtyPerLabel = 0;
        public string printertype = string.Empty;
        public List<LabelSetting> Settings { get; set; }

        public PrintLabel()
        {
            Settings = new List<LabelSetting>();
        }

        public bool LabelSetting(string commsetting, string port, int qty, ref string msg)
        {
            Settings = new List<LabelSetting>();
            string args = string.Empty;
            if (port.ToString().ToUpper().StartsWith("COM"))
            {
                printertype = "COM";
                args = port.ToString() + ";" + commsetting.ToString();
            }
            else if (port.ToString().ToUpper().StartsWith("LPT"))
            {
                printertype = "LPT";
                args = port.ToString();
            }
            else if (port.ToUpper().StartsWith("NET") && port.ToUpper() != "NETWORK")
            {
                printertype = "NETWORK";
                args = commsetting + ";" + port.Substring(3, (port.Length - 3));
            }
            else
            {
                printertype = "Default";
                args = port.ToString();
            }

            Settings.Add(new LabelSetting()
            {
                Printer = new PrinterSetting()
                {
                    PrinterType = printertype,
                    Setting = args
                },
                LabelQty = qty,
            });
            return true;
        }

        public bool Print(string strContent, DataTable dt, ref string msg)
        {
            PrintBase printer = Printer.GenPrinter(Settings[0].Printer);
            printer.LabelQty = Settings[0].LabelQty;
            printer.Content = GetPrintOut(strContent, dt);
            if (printer.Print() == false)
            {
                msg = "打印失败,请检查打印机或联系QMS人员";
                return false;
            }
            return true;
        }

        private string GetPrintOut(string template, DataTable dtPrintData)
        {
            int count;
            string output = string.Empty;

            for (int i = 0; i < dtPrintData.Columns.Count; i++)
            {
                template = template.Replace("<" + dtPrintData.Columns[i].ColumnName + ">", dtPrintData.Rows[0][i].ToString().Replace("^","_5E"));
                template = template.Replace("<" + dtPrintData.Columns[i].ColumnName.ToUpper() + ">", dtPrintData.Rows[0][i].ToString().Replace("^", "_5E"));
                template = template.Replace("<" + dtPrintData.Columns[i].ColumnName.ToLower() + ">", dtPrintData.Rows[0][i].ToString().Replace("^", "_5E"));
            }

            count = 1;
            foreach (DataRow dr in dtPrintData.Rows)
            {
                for (int iCol = 0; iCol < dtPrintData.Columns.Count; iCol++)
                {
                    //SN1~SNn
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToString() + count.ToString() + ">", dr[iCol].ToString().Replace("^", "_5E"));
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToUpper() + count.ToString() + ">", dr[iCol].ToString().Replace("^", "_5E"));
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToLower() + count.ToString() + ">", dr[iCol].ToString().Replace("^", "_5E"));

                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToString() + ">", dr[iCol].ToString().Replace("^", "_5E"));
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToUpper() + ">", dr[iCol].ToString().Replace("^", "_5E"));
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToLower() + ">", dr[iCol].ToString().Replace("^", "_5E"));
                }
                count++;
            }

            for (int iSeq = count; iSeq < 30; iSeq++)
            {
                for (int iCol = 0; iCol < dtPrintData.Columns.Count; iCol++)
                {
                    //SN1~SNn
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToString() + iSeq.ToString() + ">", "<DEL_LINE>");
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToUpper() + iSeq.ToString() + ">", "<DEL_LINE>");
                    template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToLower() + iSeq.ToString() + ">", "<DEL_LINE>");
                }
            }

            for (int iCol = 0; iCol < dtPrintData.Columns.Count; iCol++)
            {
                //SN1~SNn
                template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToString() + ">", "<DEL_LINE>");
                template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToUpper() + ">", "<DEL_LINE>");
                template = template.Replace("<" + dtPrintData.Columns[iCol].ColumnName.ToLower() + ">", "<DEL_LINE>");
            }

            string[] lines = template.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Contains("<DEL_LINE>") == false)
                {
                    output += lines[i].ToString() + "\n";
                }
            }

            return output;
        }

        public bool PrintReturn(string strContent, DataTable dt, string BU, ref string msg)
        {

            PrintBase printer = Printer.GenPrinter(Settings[0].Printer);
            printer.LabelQty = Settings[0].LabelQty;
            printer.Content = GetDIDPrintOut(strContent, BU, dt);
            if (printer.Print() == false)
            {
                msg = "打印失败,请检查打印机或联系QMS人员";
                return false;
            }
            return true;
        }

        private string GetDIDPrintOut(string template, string BU, DataTable dtPrintData)
        {
            string output = string.Empty;
            string strDID, strQty, strVendorCode="";

            if (dtPrintData.Rows[0]["IsGood"].ToString() == "Y")
            {
                strDID = dtPrintData.Rows[0]["DID"].ToString().Replace("^", "_5E");
            }
            else
            {
                strDID = dtPrintData.Rows[0]["CompPN"].ToString().Replace("^", "_5E");
            }
            strVendorCode = dtPrintData.Rows[0]["VENDORCODE"].ToString().Trim();
            if (Convert.ToInt32(dtPrintData.Rows[0]["Qty"].ToString().Trim()) <= -10000)
                strQty = "RefID";
            else
                strQty = dtPrintData.Rows[0]["Qty"].ToString();
            template = template.ToUpper();
            template = template.Replace("<DID_CODE>", strDID);
            template = template.Replace("<DID_TEXT>", strDID);
            template = template.Replace("<LINE>", BU);
            template = template.Replace("<QTY>", strQty);
            //template = template.Replace("<UID>", dtPrintData.Rows[0]["UID"].ToString());
            template = template.Replace("<DATE>", DateTime.Now.ToString("yyMMddHHmmss"));
            template = template.Replace("<VENDORCODE1>", strVendorCode);
            for (int i = 0; i < dtPrintData.Columns.Count; i++)
            {
                template = template.Replace("<" + dtPrintData.Columns[i].ColumnName + ">", dtPrintData.Rows[0][i].ToString().Replace("^", "_5E"));
                template = template.Replace("<" + dtPrintData.Columns[i].ColumnName.ToUpper() + ">", dtPrintData.Rows[0][i].ToString().Replace("^", "_5E"));
                template = template.Replace("<" + dtPrintData.Columns[i].ColumnName.ToLower() + ">", dtPrintData.Rows[0][i].ToString().Replace("^", "_5E"));
            }
            if (template.IndexOf("<WHID>") > 0)//003
            {
                template = template.Replace("<WHID>", dtPrintData.Rows[0]["WareHouseID"].ToString());
            }
            string[] lines = template.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Contains("<DEL_LINE>") == false)
                {
                    output += lines[i].ToString() + "\n";
                }
            }

            return output;
        }

        public bool PrintReturnDID(string strContent, string BU, string PrintedVenderCode, string PrintedSeqID, DataTable dtPrintData, DataSet PrintData, ref string msg)
        {
            PrintBase printer = Printer.GenPrinter(Settings[0].Printer);
            printer.LabelQty = Settings[0].LabelQty;
            printer.Content = GetNewDIDPrintOut(strContent, BU, PrintedVenderCode, PrintedSeqID, dtPrintData, PrintData);
            if (printer.Print() == false)
            {
                msg = "打印失败,请检查打印机或联系QMS人员";
                return false;
            }
            return true;
        }
        private string GetNewDIDPrintOut(string template, string BU, string PrintedVenderCode, string PrintedSeqID, DataTable dtPrintData, DataSet PrintData)
        {
            string output = string.Empty;
            int count = 1;
            string strDID, T1 = "", strWO, strDay;
            DataTable dt = PrintData.Tables[1];
            strDID = dt.Rows[0]["DID"].ToString();  //002
            if (BU == "NB4")
            {
                strDay = DateTime.Now.ToString("yyyyMMdd");
            }
            else
            {
                strDay = DateTime.Now.ToString("yyMMddHHmmss");
            }
            template = template.Replace("<DID_CODE>", strDID);
            template = template.Replace("<DID_TEXT>", strDID);
            template = template.Replace("<LINE>", dt.Rows[0]["Line"].ToString());
            template = template.Replace("<BU>", dt.Rows[0]["Line"].ToString());
            template = template.Replace("<QTY>", dtPrintData.Rows[0]["Qty"].ToString());
            template = template.Replace("<UID>", dtPrintData.Rows[0]["UID"].ToString());
            template = template.Replace("<DATE>", strDay);
            template = template.Replace("<WOTYPE>", dt.Rows[0]["WOType"].ToString());
            template = template.Replace("<WHID>", dtPrintData.Rows[0]["WareHouseID"].ToString());
            template = template.Replace("<DIDWOGROUP>", dt.Rows[0]["WoGroup"].ToString());
            foreach (DataRow dr in dt.Rows)
            {
                for (int iCol = 0; iCol < dt.Columns.Count; iCol++)
                {
                    template = template.Replace("<" + dt.Columns[iCol].ColumnName.ToString() + ">", dr[iCol].ToString());
                    template = template.Replace("<" + dt.Columns[iCol].ColumnName.ToUpper() + ">", dr[iCol].ToString());
                    template = template.Replace("<" + dt.Columns[iCol].ColumnName.ToLower() + ">", dr[iCol].ToString());
                }
                count++;
            }
            dt = PrintData.Tables[2];

            for (int i = 1; i <= 5 && i <= dt.Rows.Count; i++)
            {
                template = template.Replace("<WO" + i.ToString() + ">", dt.Rows[i - 1]["Machine"].ToString().Trim() + " " + dt.Rows[i - 1]["Slot"].ToString().Trim() + dt.Rows[i - 1]["LR"].ToString().Trim());
                template = template.Replace("<MACHINE" + i.ToString() + ">", dt.Rows[i - 1]["Machine"].ToString().Trim().Substring(1, 1) + "-" + dt.Rows[i - 1]["Slot"].ToString().Trim().Substring(i - 1, 1) + "-" + dt.Rows[i - 1]["Machine"].ToString().Trim().Substring(5, 1));
                template = template.Replace("<SLOT" + i.ToString() + ">", "");
                if (PrintedVenderCode == "Y")
                {
                    template = template.Replace("<VENDORCODE" + i.ToString() + ">", dt.Rows[i - 1]["VenderCode"].ToString().Trim());
                    template = template.Replace("<LR" + i.ToString() + ">", dt.Rows[i - 1]["SLR"].ToString().Trim());
                }
                else
                {
                    template = template.Replace("<VENDORCODE" + i.ToString() + ">", "");
                    template = template.Replace("<LR" + i.ToString() + ">", "");
                }
                if (PrintedSeqID == "Y")
                {
                    template = template.Replace("<COUNT" + i.ToString() + ">", dt.Rows[i - 1]["SeqID"].ToString().Trim());
                }
                else
                {
                    template = template.Replace("<COUNT" + i.ToString() + ">", "");
                }
                template = template.Replace("<MACHINETYPE>", dt.Rows[i - 1]["Machine"].ToString().Trim().Substring(dt.Rows[0]["Machine"].ToString().Trim().Length - 3, 3));
                template = template.Replace("<MACHINECODE>", dt.Rows[i - 1]["Machine"].ToString().Trim().Substring(dt.Rows[0]["Machine"].ToString().Trim().Length - 1, 1));
            }
            count = 1;
            template = template.ToUpper();
            foreach (DataRow dr in dt.Rows)
            {
                for (int iCol = 0; iCol < dt.Columns.Count; iCol++)
                {
                    //SN1~SNn
                    T1 = dt.Columns[iCol].ColumnName;
                    T1 = dr[iCol].ToString();
                    template = template.Replace("<" + dt.Columns[iCol].ColumnName.ToString() + count.ToString() + ">", dr[iCol].ToString());
                    template = template.Replace("<" + dt.Columns[iCol].ColumnName.ToUpper() + count.ToString() + ">", dr[iCol].ToString());
                    template = template.Replace("<" + dt.Columns[iCol].ColumnName.ToLower() + count.ToString() + ">", dr[iCol].ToString());

                    template = template.Replace("<" + dt.Columns[iCol].ColumnName.ToString() + ">", dr[iCol].ToString());
                    template = template.Replace("<" + dt.Columns[iCol].ColumnName.ToUpper() + ">", dr[iCol].ToString());
                    template = template.Replace("<" + dt.Columns[iCol].ColumnName.ToLower() + ">", dr[iCol].ToString());
                }
                count++;
            }

            if (template.IndexOf("<SlotCH1>".ToUpper()) > 0) //003
            {
                template = template.Replace("<SlotCH1>".ToUpper(), dt.Rows[0]["Slot"].ToString());
            }
            strWO = dt.Rows[0]["Machine"].ToString().Trim() + " " + dt.Rows[0]["Slot"].ToString().Trim() + dt.Rows[0]["LR"].ToString().Trim();
            if (strWO.IndexOf(" ") > 1)
            {
                template = template.Replace("<SLOT>", strWO.Substring(strWO.IndexOf(" "), strWO.Length - strWO.IndexOf(" ")));
            }

            string[] lines = template.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Contains("<DEL_LINE>") == false)
                {
                    output += lines[i].ToString() + "\n";
                }
            }

            return output;
        }


    }


    public class PrinterSetting
    {
        public string PrinterType { get; set; }

        public string Setting { get; set; }
    }

    public class LabelSetting
    {
        public PrinterSetting Printer { get; set; }

        public int LabelQty { get; set; }
    }

}
