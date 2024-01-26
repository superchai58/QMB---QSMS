using System;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.SqlClient;
using System.Threading;
using System.Diagnostics;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;





namespace QSMS.QSMS.Report
{
    public partial class frmReport : Form
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        static extern System.IntPtr FindWindow(string lpClassName, string lpWindowName);

        Excel.Application oExcel = null;

        private string exportFormat;
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.Report.ReportProcess Report = new DbLibrary.Report.ReportProcess();
        string strAddress = "";
        DataTable arryTipData = null;
        public frmReport()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void FrmReport_Load(object sender, EventArgs e)
        {
            //if(Parameter.StrBU == "NB5")
            //{
            //    this.Width = 745;
            //    this.Height = 480;               
            //}
            //int k = 1;
            arryTipData = Report.B_ToolTip_Config();
            dtpSDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
            dtpEDate.Text = DateTime.Now.AddDays(1).ToString("yyyy/MM/dd");
            CboShift.Items.Add("Day_Shift");
            CboShift.Items.Add("Night_Shift");
            GetReportType();
            GetLine();
            if (pubFunction.ConfigListGetValue("chkDual") == "Y")             
            {
                chkDual.Visible = true;
            }
           

        }
        public void GetReportType()
        {
            DataTable dt = new DataTable();
            dt = Report.Program_DefineItem();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    CboReportType.Items.Add(dt.Rows[i]["value"].ToString());
                }
            }
        }
        public void GetLine()
        {
            DataTable dt = new DataTable();
            dt = Report.Line();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    CboLine.Items.Add(dt.Rows[i]["line"].ToString());
                }
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void CboGroupID_Click(object sender, EventArgs e)
        {
          
        }

        private void CboGroupID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar == 13 || e.KeyChar == 9) && CboGroupID.Text != "")
            {
                GetGroupWO(CboGroupID.Text);
            }
        }

        private void CboGroupID_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetGroupWO(CboGroupID.Text);
        }
        public void GetGroupWO(string groupid)
        {
            string TempJobPn = "";
            if (CboJobPN.Text != "")
            {
                TempJobPn = CboJobPN.Text.Substring(0, 11);
            }
            DataTable dt = new DataTable();
            dt = Report.ListWoall(groupid, TempJobPn);
            //ListWoall.Items.Clear();
            CboWo.Items.Clear();
            CboWo.Text = "";
            CboNotChkBOM.Items.Clear();
            CboNotChkBOM.Text = "";
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                   if (ChkQSMS_WO(dt.Rows[i]["Work_Order"].ToString())==false)
                    {
                        CboNotChkBOM.Items.Add(dt.Rows[i]["Work_Order"].ToString());
                    }
                   else
                    {
                        //ListWoall.Items.Add(dt.Rows[i]["Work_Order"].ToString());
                        CboWo.Items.Add(dt.Rows[i]["Work_Order"].ToString());

                    }
                }
            }

        }
        public bool ChkQSMS_WO(string WO)
        {
            DataTable dt = new DataTable();
            dt = Report.GetWO(WO);
            if (dt.Rows.Count==0)
            {
                return false;
            }
            return true;
        }

        private void CboLine_SelectedIndexChanged(object sender, EventArgs e)
        {
            string BeginDate, EndDate;            
            BeginDate = dtpSDate.Value.ToString("yyyy/MM/dd");
            BeginDate = BeginDate.Replace("-", "");
            BeginDate = BeginDate.Replace("/", "");
            EndDate = dtpEDate.Value.ToString("yyyy/MM/dd");
            EndDate = EndDate.Replace("-", "");
            EndDate = EndDate.Replace("/", "");
            DataTable dt = new DataTable();                     
            dt = Report.GetJobpn(BeginDate, EndDate,CboLine.Text, OptRelease.Checked);
            CboJobPN.Items.Clear();
            CboJobPN.Text = "";
            if (dt != null && dt.Rows.Count>0)
            {
                for(int i=0;i<dt.Rows.Count;i++)
                {
                    CboJobPN.Items.Add(dt.Rows[i]["Jobpn"].ToString()+"-"+ dt.Rows[i]["MB_Rev"].ToString());
                }
                
            }
           
        }

        private void CboLine_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar==13 && e.ToString()!="")
            {
                string BeginDate, EndDate;
                BeginDate = dtpSDate.Value.ToString("yyyy/MM/dd");
                BeginDate = BeginDate.Replace("-", "");
                BeginDate = BeginDate.Replace("/", "");
                EndDate = dtpEDate.Value.ToString("yyyy/MM/dd");
                EndDate = EndDate.Replace("-", "");
                EndDate = EndDate.Replace("/", "");
                DataTable dt = new DataTable();
                dt = Report.GetJobpn(BeginDate, EndDate, CboLine.Text, OptRelease.Checked);
                CboJobPN.Items.Clear();
                CboJobPN.Text = "";
                if (dt != null && dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        CboJobPN.Items.Add(dt.Rows[i]["Jobpn"].ToString() + "-" + dt.Rows[i]["MB_Rev"].ToString());
                    }

                }
            }
        }

        private void CboLine_Click(object sender, EventArgs e)
        {
            //string BeginDate, EndDate;            
            //BeginDate = dtpSDate.Value.ToString("yyyy/MM/dd");
            //BeginDate = BeginDate.Replace("-", "");
            //BeginDate = BeginDate.Replace("/", "");
            //EndDate = dtpEDate.Value.ToString("yyyy/MM/dd");;
            //EndDate = EndDate.Replace("-", "");
            //EndDate = EndDate.Replace("/", "");
            //DataTable dt = new DataTable();
            //dt = Report.GetJobpn(BeginDate, EndDate, CboLine.Text, OptRelease.Checked);
            //CboJobPN.Text = "";
            //CboJobPN.Items.Clear();
            //if (dt != null && dt.Rows.Count > 0)
            //{
            //    for (int i = 0; i < dt.Rows.Count; i++)
            //    {
            //        CboJobPN.Items.Add(dt.Rows[i]["Jobpn"].ToString().Trim() + "-" + dt.Rows[i]["MB_Rev"].ToString().Trim());
            //    }

            //}
        }

        private void CboMachine_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && e.ToString() != "")
            {
                GetComp(CboMachine.Text.Trim(), TxtMBPN.Text.Trim(), TxtWO.Text.Trim(), CboLine.Text.Trim());
                GetJobGroupByJobRev(CboMachine.Text.Trim(), TxtMBPN.Text.Trim(), TxtRev.Text.Trim());
            }
        }

        private void CboMachine_Click(object sender, EventArgs e)
        {
            //GetComp(CboMachine.Text.Trim(), TxtMBPN.Text.Trim(), TxtWO.Text.Trim(), CboLine.Text.Trim());
            //GetJobGroupByJobRev(CboMachine.Text.Trim(), TxtMBPN.Text.Trim(),TxtRev.Text.Trim());
        }
        public void GetComp(string Machine,string MBPN,string WO,string Line)
        {
            DataTable dt = new DataTable();
            dt = Report.GetComp(Machine, MBPN, WO, Line);
            CboComp.Items.Clear();
            CboComp.Text = "";
            if(dt != null && dt.Rows.Count>0)
            {
                for(int i=0;i<dt.Rows.Count;i++)
                {
                    CboComp.Items.Add(dt.Rows[i]["compPN"].ToString());
                }
            }
        }
        public void GetJobGroupByJobRev(string Machine,string MBPN,string Rev)
        {
            DataTable dt = new DataTable();
            dt = Report.GetJobGroupByJobRev(Machine, MBPN, Rev);
            ListAllJobGroup.Items.Clear();
            ListselectingJobGroup.Items.Clear();
            if (dt != null && dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ListAllJobGroup.Items.Add(dt.Rows[i]["jobgroup"].ToString());
                }
            }
        }

        private void CboNotChkBOM_Click(object sender, EventArgs e)
        {
            

        }

        private void CboNotChkBOM_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar==13 && e.ToString()!="")
            {
                TxtWO.Text = CboNotChkBOM.Text.Trim();
                GetWOInfo(TxtWO.Text);
                GetMachine(TxtWO.Text);
            }
        }
        public  void GetWOInfo(string WO)
        {
            DataTable dt = new DataTable();
            dt = Report.GetWoinfo(WO);
            if(dt != null && dt.Rows.Count>0)
            {
                TxtMBPN.Text = dt.Rows[0]["PN"].ToString().Trim();
                TxtWOQty.Text= dt.Rows[0]["QTY"].ToString().Trim();
                TxtRev.Text = dt.Rows[0]["MB_REV"].ToString().Trim();
                CboLine.Text = dt.Rows[0]["Line"].ToString().Trim();
            }
            dt = Report.GetCustomer(TxtMBPN.Text, TxtRev.Text);
            if(dt != null && dt.Rows.Count>0)
            {
                TxtCustomer.Text = dt.Rows[0]["Customer"].ToString();
            }
            CboJobPN.Items.Clear();
            CboJobPN.Text = "";
            dt = Report.GetjobPn(WO);
            if(dt != null && dt.Rows.Count>0)
            {
                for(int i=0;i<dt.Rows.Count;i++)
                {
                    CboJobPN.Items.Add(dt.Rows[i]["jobpn"].ToString());
                }
            }
        }
        public void GetMachine(string WO)
        {
            DataTable dt = new DataTable();
            dt = Report.GetMachine(WO);
            CboMachine.Items.Clear();
            CboMachine.Text = "";
            CboMachine.Items.Add("ALL");
            if(dt != null && dt.Rows.Count>0)
            {
                for(int i=0;i<dt.Rows.Count;i++)
                {
                    CboMachine.Items.Add(dt.Rows[i]["Machine"].ToString().Trim());
                }
            }
        }

        private void CboReportType_Click(object sender, EventArgs e)
        {
            if(CboReportType.Text.ToUpper()=="DIDCALLBACK")
            {
                QSMS.Report.frmQueryCheckBOM frm = new QSMS.Report.frmQueryCheckBOM();
                pubFunction.HaveOpened(frm, "frmQueryDIDCallBack");
            }
            else if (CboReportType.Text.ToUpper() == "DIDCALLBACK")
            {
                return;
            }
        }

        private void CboWo_Click(object sender, EventArgs e)
        {
            //TxtWO.Text = CboWo.Text.Trim();
            //GetWOInfo(TxtWO.Text);
            //GetMachine(TxtWO.Text);
        }

        private void CboWo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && e.ToString() != "")
            {
                //CboWo_Click(null, null);
                TxtWO.Text = CboWo.Text.Trim();
                GetWOInfo(TxtWO.Text);
                GetMachine(TxtWO.Text);
            }
        }

        private void cmdADD_Click(object sender, EventArgs e)
        {
            int Pointer;
            if (ListWoall.Items.Count <= 0)
            {
                return;
            }
            if (ListWoall.SelectedIndex < 0)
            {
                return;
            }
            Pointer = ListWoall.SelectedIndex;
            ListWoSelecting.Items.Add(ListWoall.SelectedItem.ToString());
            ListWoall.Items.RemoveAt(Pointer);
            if (ListWoall.Items.Count != Pointer)
            {
                ListWoall.SelectedIndex = Pointer;
            }

        }

        private void cmdADDALL_Click(object sender, EventArgs e)
        {
            if (ListWoall.Items.Count <= 0)
            {
                return;
            }
            for (int i = ListWoall.Items.Count; i > 0; i--)
            {
                ListWoall.SelectedIndex = 0;
                ListWoSelecting.Items.Add(ListWoall.SelectedItem.ToString());
                ListWoall.Items.RemoveAt(0);
            }
        }

        private void cmdDEL_Click(object sender, EventArgs e)
        {
            int Pointer;
            if (ListWoSelecting.Items.Count <= 0)
            {
                return;
            }
            if (ListWoSelecting.SelectedIndex < 0)
            {
                return;
            }
            Pointer = ListWoSelecting.SelectedIndex;
            ListWoall.Items.Add(ListWoSelecting.SelectedItem.ToString());
            ListWoSelecting.Items.RemoveAt(Pointer);
            if (ListWoSelecting.Items.Count != Pointer)
            {
                ListWoSelecting.SelectedIndex = Pointer;
            }
        }

        private void cmdDELALL_Click(object sender, EventArgs e)
        {
            if (ListWoSelecting.Items.Count <= 0)
            {
                return;
            }
            for (int i = ListWoSelecting.Items.Count; i > 0; i--)
            {
                ListWoSelecting.SelectedIndex = 0;
                ListWoall.Items.Add(ListWoSelecting.SelectedItem.ToString());
                ListWoSelecting.Items.RemoveAt(0);
            }
        }

        private void CmdAddGroup_Click(object sender, EventArgs e)
        {
            int Pointer;
            if (ListAllJobGroup.Items.Count <= 0)
            {
                return;
            }
            if (ListAllJobGroup.SelectedIndex < 0)
            {
                return;
            }
            Pointer = ListAllJobGroup.SelectedIndex;
            ListselectingJobGroup.Items.Add(ListAllJobGroup.SelectedItem.ToString());
            ListAllJobGroup.Items.RemoveAt(Pointer);
            if (ListAllJobGroup.Items.Count != Pointer)
            {
                ListAllJobGroup.SelectedIndex = Pointer;
            }
        }

        private void cmdADDALLGroup_Click(object sender, EventArgs e)
        {
            if (ListAllJobGroup.Items.Count <= 0)
            {
                return;
            }
            for (int i = ListAllJobGroup.Items.Count; i > 0; i--)
            {
                ListAllJobGroup.SelectedIndex = 0;
                ListselectingJobGroup.Items.Add(ListAllJobGroup.SelectedItem.ToString());
                ListAllJobGroup.Items.RemoveAt(0);
            }
        }

        private void cmdDELGroup_Click(object sender, EventArgs e)
        {
            int Pointer;
            if (ListselectingJobGroup.Items.Count <= 0)
            {
                return;
            }
            if (ListselectingJobGroup.SelectedIndex < 0)
            {
                return;
            }
            Pointer = ListselectingJobGroup.SelectedIndex;
            ListAllJobGroup.Items.Add(ListselectingJobGroup.SelectedItem.ToString());
            ListselectingJobGroup.Items.RemoveAt(Pointer);
            if (ListselectingJobGroup.Items.Count != Pointer)
            {
                ListselectingJobGroup.SelectedIndex = Pointer;
            }
        }

        private void cmdDELALLGroup_Click(object sender, EventArgs e)
        {
            if (ListselectingJobGroup.Items.Count <= 0)
            {
                return;
            }
            for (int i = ListselectingJobGroup.Items.Count; i > 0; i--)
            {
                ListselectingJobGroup.SelectedIndex = 0;
                ListAllJobGroup.Items.Add(ListselectingJobGroup.SelectedItem.ToString());
                ListselectingJobGroup.Items.RemoveAt(0);
            }
        }



        private void CmdQuery_Click(object sender, EventArgs e)
        {
            if(CboLine.Text=="")
            {
                MessageBox.Show("Please input line");
                return;
            }
            GetGroupID("");
            GetJobPN();
        }
        public void GetGroupID(string JobPN)
        {
            string BeginDate, EndDate;           
            BeginDate = dtpSDate.Value.ToString("yyyy/MM/dd");
            BeginDate = BeginDate.Replace("-", "");
            BeginDate = BeginDate.Replace("/", "");
            EndDate = dtpEDate.Value.ToString("yyyy/MM/dd"); ;
            EndDate = EndDate.Replace("-", "");
            EndDate = EndDate.Replace("/", "");
            DataTable dt = new DataTable();
            CboGroupID.Items.Clear();
            CboGroupID.Text = "";
            if (JobPN!="" && JobPN.IndexOf("-")>0)
            {
                TxtMBPN.Text = JobPN.Substring(1, 11);
                TxtRev.Text = JobPN.Substring(JobPN.Length-3);
             }  
            dt = Report.GetGroupID(JobPN,Parameter.BU, BeginDate, EndDate, CboLine.Text, OptRelease.Checked);
            
            if(dt != null && dt.Rows.Count>0)
            {
                for(int i=0;i<dt.Rows.Count;i++)
                {
                    CboGroupID.Items.Add(dt.Rows[i]["GroupID"].ToString().Trim());
                }
            }
            else
            {
                MessageBox.Show("No data");
                return;
            }

        }
        public void GetJobPN()
        {
            string BeginDate, EndDate;            
            BeginDate = dtpSDate.Value.ToString("yyyy/MM/dd");
            BeginDate = BeginDate.Replace("-", "");
            BeginDate = BeginDate.Replace("/", "");
            EndDate = dtpEDate.Value.ToString("yyyy/MM/dd"); ;
            EndDate = EndDate.Replace("-", "");
            EndDate = EndDate.Replace("/", "");
            DataTable dt = new DataTable();
            dt = Report.GetjobPn(Parameter.BU,CboLine.Text,BeginDate,EndDate,OptRelease.Checked);
            CboJobPN.Items.Clear();
            CboJobPN.Text = "";
            if(dt != null && dt.Rows.Count>0)
            {
                for(int i=0;i<dt.Rows.Count;i++)
                {
                    CboJobPN.Items.Add(dt.Rows[i]["JobPN"].ToString().Trim());
                }
            }
            else
            {
                MessageBox.Show("No data");
            }
        }

        private void CboJobPN_Click(object sender, EventArgs e)
        {
            //GetGroupID(CboJobPN.Text.ToString().Trim());
            //GetJobGroupByJobRev("", TxtMBPN.Text.Trim(), TxtRev.Text.Trim());
        }

        private void CboJobPN_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar==13 && e.ToString()!="")
            {
                //CboJobPN_Click(null, null);
                GetGroupID(CboJobPN.Text.ToString().Trim());
                GetJobGroupByJobRev("", TxtMBPN.Text.Trim(), TxtRev.Text.Trim());
            }
        }

        private void CboJobPN_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetGroupID(CboJobPN.Text.ToString().Trim());
            GetJobGroupByJobRev("", TxtMBPN.Text.Trim(), TxtRev.Text.Trim());
        }

        private void CboWo_SelectedIndexChanged(object sender, EventArgs e)
        {
            TxtWO.Text = CboWo.Text.Trim();
            GetWOInfo(TxtWO.Text);
            GetMachine(TxtWO.Text);
        }

        private void CboReportType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CboReportType.Text.ToUpper() == "DIDCALLBACK")
            {
                QSMS.Report.frmQueryCheckBOM frm = new QSMS.Report.frmQueryCheckBOM();
                pubFunction.HaveOpened(frm, "frmQueryDIDCallBack");
            }
            else if (CboReportType.Text.ToUpper() == "DIDCALLBACK")
            {
                return;
            }
        }

        private void CboNotChkBOM_SelectedIndexChanged(object sender, EventArgs e)
        {
            TxtWO.Text = CboNotChkBOM.Text.Trim();
            GetWOInfo(TxtWO.Text);
            GetMachine(TxtWO.Text);
        }
        private void cmdExcel_Click(object sender, EventArgs e)
        {

            string WO = "";
            if (CboReportType.Text.Trim() == "PrepariMaterialList")
            {
                if (CboMachine.Text == "")
                {
                    MessageBox.Show("Please select Machine!");
                    return;
                }
                if (ListWoSelecting.Items.Count <= 0)
                {
                    MessageBox.Show("Please select the wo to listbox---wo selecting");
                    return;
                }
                for (int i = 0; i < ListWoSelecting.Items.Count; i++)
                {
                    ListWoSelecting.SelectedIndex = i;
                    WO = WO + ListWoSelecting.SelectedItem.ToString() + ",";
                }
                WO = WO.Substring(0, WO.Length - 1);
                CopyToExcelPrepareMaterialList("PrepariMaterialList", WO.Trim(), CboMachine.Text.Trim(), CboLine.Text.Trim());
            }
            else if (CboReportType.Text.Trim() == "PrepariMaterialByJobPN")
            {
                if (ListWoSelecting.Items.Count <= 0)
                {
                    MessageBox.Show("Please select the wo to listbox---wo selecting");
                    return;
                }
                PrepareMaterialByWONew("By_JobPN", CboLine.Text.Trim());
            }
            else if (CboReportType.Text.Trim() == "QSMS_CheckCompPN")
            {
                Load_QSMS_CheckCompPN("1");
            }
            else if(CboReportType.Text.Trim() == "CheckReplacePNBySAPBOM")
            {
                if(txtFilePath.Text.Trim()=="")
                {
                    MessageBox.Show("请先选择上传的文档，谢谢!");
                    return;
                }
                Load_CheckReplacePNBySAPBOM("1");
            }
            else if(CboReportType.Text.Trim() == "PrepariMaterialByLineShift")
            {
                if (CboShift.Text == "")
                {
                    MessageBox.Show("Please select the shift!");
                    return;
                }
                if (CboLine.Text == "")
                {
                    MessageBox.Show("Please select the Line!");
                    return;
                }
                if (dtpSDate.Value > dtpEDate.Value)
                {
                    MessageBox.Show("Please Select BeginDate and EndDate Again!");
                    return;
                }
                PrepareMaterialByWONew("By_Shift", CboLine.Text.Trim());
            }
            else if(CboReportType.Text.Trim() == "PrepariMaterialByGroup")
            {
                PrepareMaterialByWONew("By_Group", CboLine.Text.Trim());
            }
            else if(CboReportType.Text.Trim() == "PrepariMaterialByWos")
            {
                PrepareMaterialByWONew("By_WorkOrders", CboLine.Text.Trim());
            }
            else if(CboReportType.Text.Trim() == "PrepariMaterialByWo")
            {
                PrepareMaterialByWONew("By_WorkOrder", CboLine.Text.Trim());
            }

            else if(CboReportType.Text.Trim() == "LineChangeStatisticsByall")
            {
                LineChangeStatisticsByall();
            }
            else if(CboReportType.Text.Trim() == "DIDDeleteRecords")
            {
                DIDDeleteRecords();
            }
            else if(CboReportType.Text.Trim() == "LineChangeStatistics")
            {
                LineChangeStatistics();//需要改excel格式
            }
            else if(CboReportType.Text.Trim() == "PrepariMaterialMonitor")
            {
                XL_MonitorReport();
            }
            else if(CboReportType.Text.Trim() == "DispatchQTYByWO")
            {
                DispatchQTYByWO();
            }
            else if(CboReportType.Text.Trim() == "QSMS_DID_ToWH")
            {
                QSMS_DID_ToWH();
            }
            else if(CboReportType.Text.Trim() == "QSMS_WO")
            {
                QSMS_WO();
            }
            else if(CboReportType.Text.Trim() == "WipByMaterial")
            {
                if (CboComp.Text == "")
                {
                    MessageBox.Show("CompPN can not be empty,Please check");
                    return;
                }
                CopyToExcelWipByMaterial("WipByMaterial", CboComp.Text.Trim());
            }
            else if(CboReportType.Text.Trim() == "WipByDate")
            {

                CopyToExcelWipByDate("WipByDate");
            }
            else if(CboReportType.Text.Trim() == "WipByGroup")
            {
                if (CboGroupID.Text == "")
                {
                    MessageBox.Show("Please check the GroupID");
                    return;
                }
                CopyToExcelWipByGroup("WipByGroup", CboGroupID.Text);
            }
            else if(CboReportType.Text.Trim() == "WipLackbyWo")
            {
                if (CboWo.Text == "")
                {
                    MessageBox.Show("Please check the WO");
                    return;
                }

                CopyToExcelWipLackByWo("WipLackbyWo", TxtWO.Text.Trim());
            }
            else if(CboReportType.Text.Trim() == "MaterialDifferentList")
            {
                if (CboWo.Text == "")
                {
                    MessageBox.Show("Please check the WO");
                    return;
                }

                CopyToExcelWipDifferentMaterial("MaterialDifferentList", TxtWO.Text.Trim());
            }
            else if (CboReportType.Text.Trim() == "PDUsedByCompLine")
            {
                PDUsedByCompLine(CboLine.Text.Trim(), CboComp.Text.Trim());
            }
            else if(CboReportType.Text.Trim() == "SAP_BOM")
            {

                GetSapBom(TxtWO.Text.Trim());
            }
            else if(CboReportType.Text.Trim() == "SAP_GroupByWo")
            {
                GetSapGroupByWo(TxtWO.Text.Trim());
            }
            else if(CboReportType.Text.Trim() == "ME_BOM")
            {
                GetMEBom(TxtWO.Text.Trim());
            }
            else if(CboReportType.Text.Trim() == "CheckWO_WastagePN")
            {
                CheckWO_WastagePN(TxtWO.Text.Trim());
            }
            else if(CboReportType.Text.Trim() == "GetGroupIDDataByCompPN") //需要测试
            {
                if (CboGroupID.Text == "")
                {
                    MessageBox.Show("Please input the GroupID");
                    CboGroupID.Focus();
                    return;
                }
                if (CboComp.Text == "")
                {
                    MessageBox.Show("Please input the CompPN");
                    CboComp.Focus();
                    return;
                }
                GETGROUPIDDATABYCOMPPN(CboGroupID.Text, CboComp.Text);
            }
            else if(CboReportType.Text.Trim() == "MEBOM_Delete_Log")
            {
                MEBOM_Delete_Log(TxtMBPN.Text.Trim());
            }
            else if(CboReportType.Text.Trim() == "ME_BOM_WO")
            {
                GetMEBom_WO(TxtWO.Text.Trim());
            }
            else if (CboReportType.Text.Trim() == "ReplacePN")
            {
                if (TxtWO.Text != "")
                {
                    GetReplacePN(TxtWO.Text);
                }
                else if (ListWoSelecting.Items.Count <= 0)
                {
                    MessageBox.Show("Please select the wo to listbox---wo selecting,or upload the wolist");
                }
                else
                {
                    for (int i = 0; i < ListWoSelecting.Items.Count - 1; i++)
                    {
                        ListWoSelecting.SelectedIndex = i;
                        WO = WO + ListWoSelecting.SelectedItem + ",";
                    }
                    WO = WO.Substring(0, WO.Length - 1);
                    GetReplacePN(WO);
                }
            }
            else if(CboReportType.Text.Trim() == "CheckBOM")
            {
                CheckBOM();

            }
            else if(CboReportType.Text.Trim() == "CheckBOMDiff")
            {
                GetChkBOMDiff(TxtWO.Text);

            }
            else if(CboReportType.Text.Trim() == "RefreshBOM")
            {
                RefreshBOM(TxtWO.Text);

            }
            else if(CboReportType.Text.Trim() == "CheckBOM_Rate")
            {
                CheckBOM_Rate();

            }
            else if(CboReportType.Text.Trim() == "SAPCostSum" || CboReportType.Text.Trim() == "SAP1" ||
                CboReportType.Text.Trim() == "SAP1His" || CboReportType.Text.Trim() == "SAP2" ||
                CboReportType.Text.Trim() == "ReturnDID" || CboReportType.Text.Trim() == "ReturnDID_ByDate" ||
                CboReportType.Text.Trim() == "DispatchDID" || CboReportType.Text.Trim() == "Return_Dispatch" ||
                CboReportType.Text.Trim() == "DIDCallBack" || CboReportType.Text.Trim() == "SAPFileChk" ||
                CboReportType.Text.Trim() == "CastQty" || CboReportType.Text.Trim() == "WO_SingleCompPNData" ||
                CboReportType.Text.Trim() == "GroupIDCostQty")
            {
                Sap_Return(CboReportType.Text.Trim());

            }
            else if(CboReportType.Text.Trim() == "ReturnDIDByGroupID" || CboReportType.Text.Trim() == "ReturnDIDByWO")
            {
                ReturnDID(CboReportType.Text.Trim(), CboGroupID.Text.Trim());
            }
            else if (CboReportType.Text.Trim() == "AOIQtySummary")
            {
                AOIQtySummary();
            }
            else if (CboReportType.Text.Trim() == "AOIDetail")
            {
                AOIDetail(TxtWO.Text);
            }
            else if(CboReportType.Text.Trim() == "MachineType")
            {
                MachineType(CboReportType.Text.Trim(), "");
            }
            else if (CboReportType.Text.Trim() == "TraySlot" || CboReportType.Text.Trim() == "CastRate"||
                CboReportType.Text.Trim() == "OneByOne"|| CboReportType.Text.Trim()== "UnCloseGroupID"||
                CboReportType.Text.Trim()== "SpliceReplacePN"|| CboReportType.Text.Trim() == "SplicePN"||
                CboReportType.Text.Trim() == "MaintainFeeder")
            {
                MachineType(CboReportType.Text.Trim(), "");
            }
            else if(CboReportType.Text.Trim() == "VerifyReport" || CboReportType.Text.Trim() == "VerifyReportWOChged")
            {
                ToExcelVerifyReport(CboReportType.Text.Trim());//需要测试
            }
            else if(CboReportType.Text.Trim() == "VerifyJobFailLog")
            {
                VerifyJobFailLog();//需要测试
            }
            else if(CboReportType.Text.Trim() == "UnDispatchList")
            {
                GetUnDispatchList(TxtWO.Text.ToString());//需要测试
            }
            else if(CboReportType.Text.Trim() == "XL_ReelBaseQty"||CboReportType.Text.Trim() == "NonAVL" || 
                CboReportType.Text.Trim() == "CompPN_DIDData" || CboReportType.Text.Trim() == "CompPNQty")
            {
                MachineType(CboReportType.Text.Trim(), CboComp.Text);//需要测试
            }
            else if(CboReportType.Text.Trim() == "SameGroupWO"|| CboReportType.Text.Trim() == "FUJI_AVLList"||
                CboReportType.Text== "CheckBom_Log" || CboReportType.Text == "CheckBom_Result")
            {
                MachineType(CboReportType.Text.Trim(), TxtWO.Text);//需要测试
            }
            //if (CboReportType.Text.Trim() == "ME_BOM_GroupID")
            //{
            //    MachineType(CboReportType.Text.Trim(), CboGroupID.Text);//需要测试
            //}
            else if(CboReportType.Text.Trim() == "AllDispatchByGroupID"|| CboReportType.Text.Trim() == "ME_BOM_GroupID")
            {
                MachineType(CboReportType.Text.Trim(), CboGroupID.Text);//需要测试
            }
            else if(CboReportType.Text.Trim() == "MEBom_EQProgram")
            {
                MachineType("MEBom_EQProgram", TxtJobGroup.Text);//需要测试
            }
            else if (CboReportType.Text.Trim() == "CheckDispatchQty")
            {
                CheckDispatchQty( TxtWO.Text);//需要测试
            }
            else if (CboReportType.Text.Trim() == "DIDIntegration" || CboReportType.Text.Trim() == "Glue_CallOff" ||
                CboReportType.Text.Trim() == "XL_MaterialDemand"||CboReportType.Text== "XL_DemandDetail" || 
                CboReportType.Text == "Glue_Consumption")
            {
                XL_MaterialDemand(dtpSDate.Value.ToString("yyyyMMdd"), dtpEDate.Value.ToString("yyyyMMdd"),CboReportType.Text);//需要测试
            }
            else if(CboReportType.Text.Trim() == "WoInputPlan"||CboReportType.Text== "WoInputPlanBySide")
            {
                WoInputPlan(CboReportType.Text.Trim());//需要测试
            }
            else if(CboReportType.Text.Trim() == "XL_DispatchStatus")
            {
                XL_DispatchStatus();//需要测试
            }

            else if(CboReportType.Text.Trim() == "MEBom_Model")
            {
                MEBom_Model();//需要测试
            }
            else if(CboReportType.Text.Trim() == "DIDCompare" )
            {
                XL_MaterialDemand(dtpSDate.Value.ToString("yyyyMMdd")+ "080000", dtpEDate.Value.AddDays(1).ToString("yyyyMMdd") + "080000", CboReportType.Text);//需要测试
            }
            else if(CboReportType.Text.Trim() == "CheckSpliceReplacePN"||CboReportType.Text== "ForbiddenPN")
            {
                XL_MaterialDemand(dtpSDate.Value.ToString("yyyyMMdd") + "074000", dtpEDate.Value.AddDays(1).ToString("yyyyMMdd") + "074000", CboReportType.Text);//需要测试
            }
            else if(CboReportType.Text.Trim() == "Glue_DataByDay")
            {
                XL_MaterialDemand(dtpSDate.Value.ToString("yyyy/MM/dd") , dtpEDate.Value.AddDays(1).ToString("yyyy/MM/dd"), CboReportType.Text);//需要测试
            }
            else if(CboReportType.Text.Trim() == "MaterialReturn" || CboReportType.Text.Trim() == "PanalnterLock")
            {
                XL_MaterialDemand(dtpSDate.Value.ToString("yyyyMMdd") + "000000", dtpEDate.Value.AddDays(1).ToString("yyyyMMdd") + "240000", CboReportType.Text);//需要测试
            }
            else if(CboReportType.Text.Trim() == "PDA_DistributeDIDLog" )
            {
                XL_MaterialDemand(dtpSDate.Value.ToString("yyyyMMdd") + "000000", dtpEDate.Value.ToString("yyyyMMdd") + "240000", CboReportType.Text);//需要测试
            }
            else
            {
                MessageBox.Show("Please select the function type");
            }
        }
        public void MEBom_Model()
        {
            DataTable dt = new DataTable();
            if(TxtMBPN.Text=="" || TxtRev.Text=="")
            {
                MessageBox.Show("Please input JobPN and Revision !");
                TxtMBPN.Focus();
                return;
            }
            else
            {
                dt = Report.MEBom_Model(TxtMBPN.Text, TxtRev.Text);
                if(dt.Rows.Count>0)
                {
                    OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                    //CopyToExcel("", "", "", "", "", "", dt, null);
                }
                else
                {
                    MessageBox.Show("No Data");
                    return;
                }
            }
        }
        //public void XL_ReelBaseQty()
        //{
        //    if(CboComp.Text=="")
        //    {
        //        MessageBox.Show("CompPN can not be empty,Please check");
        //        return;
        //    }
        //}
        public void XL_DispatchStatus()
        {
            string strSdate = "", strEDate="", strShift = "";
            DataTable dt = new DataTable();
            if ((dtpEDate.Value-dtpSDate.Value).Days>3)
            {
                MessageBox.Show("Date range over 3 days!");
                return;
            }
            if(CboShift.Text=="")
            {
                MessageBox.Show("Please select the shift!");
                return;
            }
            if(CboLine.Text=="")
            {
                MessageBox.Show("Please select the line!");
                return;
            }
            strSdate = dtpSDate.Value.ToString("yyyyMMdd");
            strEDate = dtpEDate.Value.ToString("yyyyMMdd");
            strShift = CboShift.Text.Substring(0, 1);
            dt = Report.XL_DispatchStatus(CboLine.Text, strSdate, strEDate, strShift);
            if(dt.Rows.Count>0)
            {
                //CopyToExcel("", "", "", "", "", "", dt, null);
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
            }
            else
            {
                MessageBox.Show("No data found!");
                return;
            }

        }
        public void WoInputPlan(string ReportType)
        {
            DataSet ds = new DataSet();
            string SDate, EDate;
            SDate = dtpSDate.Value.ToString("yyyyMMdd");
            EDate = dtpEDate.Value.AddDays(1).ToString("yyyyMMdd");
            if(dtpSDate.Value>dtpEDate.Value)
            {
                MessageBox.Show("The StartDate must be smaller than EndDate !");
                return;
            }
            if((dtpEDate.Value-dtpSDate.Value).Days>31)
            {
                MessageBox.Show("The day must less than 31 days !");
                return;
            }
            ds = Report.Rpt_XL_GetWoInputPlan(SDate,EDate,TxtWO.Text, ReportType);
            if(ds.Tables.Count>0)
            {
                if (ReportType == "WoInputPlan")
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        //CopyToExcel("", "", "", "", "", "", ds.Tables[0], null);
                        OfficeExcel(ds.Tables[0], ds.Tables[0].Rows.Count, ds.Tables[0].Columns.Count);
                    }
                    if (ds.Tables[1].Rows.Count > 0)
                    {
                        //CopyToExcel("", "", "", "", "", "", ds.Tables[0], null);
                        OfficeExcel(ds.Tables[1], ds.Tables[1].Rows.Count, ds.Tables[1].Columns.Count);
                    }
                    else
                    {
                        MessageBox.Show("No Detail data");
                        return;
                    }
                }
                else if(ReportType== "WoInputPlanBySide")
                {
                    ToExcel(ds);
                }
            }
           

        }
        public void XL_MaterialDemand(string Sdate,string Edate,string ReportType)
        {
            DataSet ds = new DataSet();
            if(ReportType== "DIDIntegration")
            {
                ds = Report.XL_MaterialDemand(Sdate, Edate, CboGroupID.Text, ReportType);
            }
            else if(ReportType == "PDA_DistributeDIDLog")
            {
                ds = Report.XL_MaterialDemand(Sdate, Edate, CboLine.Text, ReportType);
            }
            else
            {
                ds = Report.XL_MaterialDemand(Sdate, Edate, CboComp.Text, ReportType);
            }            
            
            if (ds.Tables[0].Rows.Count > 0 && ReportType== "XL_MaterialDemand") 
            {
                ToExcel(ds);
            }
            else if (ds.Tables[0].Rows.Count > 0 && ReportType == "Glue_DataByDay")
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    //CopyToExcel("", "", "", "", "", "", ds.Tables[0], null);
                    OfficeExcel(ds.Tables[1], ds.Tables[1].Rows.Count, ds.Tables[1].Columns.Count);
                }
                if (ds.Tables[1].Rows.Count > 0)
                {
                    //CopyToExcel("", "", "", "", "", "", ds.Tables[1], null);
                    OfficeExcel(ds.Tables[1], ds.Tables[1].Rows.Count, ds.Tables[1].Columns.Count);
                }
                else
                {
                    MessageBox.Show("No Detail data!");
                    return;
                }
            }
            else
            {
                if(ds.Tables[0].Rows.Count>0)
                {
                    //CopyToExcel("", "", "", "", "", "", ds.Tables[0], null);
                    OfficeExcel(ds.Tables[0], ds.Tables[0].Rows.Count, ds.Tables[0].Columns.Count);
                }
                else
                {
                    MessageBox.Show("No data found!");
                    return;
                }
               
            }
        }
        public void CheckDispatchQty(string WO)
        {
            DataTable dt = new DataTable();
            dt = Report.sapwolist(WO);
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("Can't find the WO!Please check!");
                return;
            }
            Report.QSMS_ReCountDispatchQty(WO,"2");
            MessageBox.Show("That's OK!");
        }
        public void GetUnDispatchList(string Work_Order)
        {
            DataTable dt = new DataTable();
            dt = Report.GetUnDispatchList(Work_Order,"N");
            if (dt.Rows.Count > 0)
            {
                //CopyToExcel("", "", "", "", "", "", dt, null);
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
            }
            else
            {
                dt = Report.GetUnDispatchList(Work_Order, "Y");
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("The WO has complete to dispatch!");
                }
                else
                {
                    MessageBox.Show("The WO hasn't begined to dispatch!");
                }
            }
        }
        public void VerifyJobFailLog()
        {
            DataTable dt = new DataTable();
            string Sdate = "", Edate = "";
            if ((dtpEDate.Value- dtpSDate.Value).Days>7)
            {
                MessageBox.Show("The date range must be <= 7 days!");
                return;
            }
            Sdate = dtpSDate.Value.ToString("yyyyMMdd");
            Edate = dtpEDate.Value.ToString("yyyyMMdd");
            dt = Report.GenVerifyFailReport(Sdate, Edate);
            if (dt.Rows.Count > 0)
            {
                //CopyToExcel("", "", "", "", "", "", dt, null);

                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
            }
            else
            {
                MessageBox.Show("No data found");
            }
        }
        public void ToExcelVerifyReport(string Type)
        {
            string strLine = "";
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            dt = Report.machine();
            if(dt.Rows.Count>0)
            {
                string col1;
                int row, col;               
                object missing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xBk;
                xBk = appExcel.Workbooks.Add(true);                
                Microsoft.Office.Interop.Excel.Worksheet xSt;
                //xBk = appExcel.Workbooks.Add(true);               
                appExcel.Visible = true;         

                for (int i=0;i<dt.Rows.Count;i++)
                {
                    xBk = appExcel.Workbooks.Add(true);
                    xSt = (Excel.Worksheet)xBk.Worksheets.get_Item(i);
                    xSt = (Excel.Worksheet)xBk.ActiveSheet;
                    xSt.Name= dt.Rows[i]["Expr1"].ToString();
                    strLine = dt.Rows[i]["Expr1"].ToString();
                    dt1 = Report.GenVerifyReportToExcel(strLine,Type);
                    row = dt1.Rows.Count;
                    col = dt1.Columns.Count - 1;
                    col1 = GetColumnChar(col);
                    if (dt.Rows.Count>0)
                    {
                        for (int m = 0; m < dt1.Columns.Count; m++)
                        {
                            xSt.Cells[1, m + 1] = dt1.Columns[m].ColumnName.ToString();
                        }
                        for (int n = 0; n < dt1.Rows.Count; n++)
                        {
                            for (int m = 0; m < dt1.Columns.Count; m++)
                            {
                                xSt.Cells[n + 2, m + 1] = dt1.Rows[i][m].ToString();
                            }

                        }
                    }
                    string col2 = col1 + Convert.ToString(row + 1);
                    Microsoft.Office.Interop.Excel.Range range1 = xSt.Range["A1", col2];
                    range1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    range1.EntireColumn.AutoFit();
                    range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//水平居中  
                    range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//垂直居中

                }
            }
        }
        public void MachineType(string flag,string compPN)
        {

            DataTable dt = new DataTable();
            if(flag== "XL_ReelBaseQty")
            {
                if(compPN=="")
                {
                    MessageBox.Show("CompPN can not be empty,Please check");
                    return;
                }
            }
            if (flag == "ME_BOM_GroupID" || flag== "MEBom_EQProgram")
            {
                if (compPN == "")
                {
                    MessageBox.Show("Please Input GroupID !!!");
                    return;
                }
            }
            if (flag == "AllDispatchByGroupID")
            {
                if (compPN == "")
                {
                    MessageBox.Show("Please select one groupid");
                    return;
                }
            }
           

            dt = Report.MachineType(flag, compPN);
            if (dt != null && dt.Rows.Count > 0)
            {
                //CopyToExcel("", "", "", "", "", "", dt, null);
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
            }
            else
            {
                MessageBox.Show("No data found");
            }

        }
        public void AOIDetail(string Work_Order)
        {
          
            DataTable dt = new DataTable();
            dt = Report.AOIDetail(Work_Order);
            if (dt != null &&  dt.Rows.Count > 0)
            {
                //CopyToExcel("", "", "", "", "", "", dt, null);
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
            }
            else
            {
                MessageBox.Show("No data found");
            }

        }
        public void AOIQtySummary()
        {
            string WO = "";
            DataTable dt = new DataTable();
            if(TxtWO.Text!="")
            {
                dt=Report.AOIQtySummary(TxtWO.Text.Trim());
            }
            else if(ListWoSelecting.Items.Count<=0)
            {
                MessageBox.Show("Please select the wo to listbox---wo selecting,or upload the wolist");
                return;
            }
            else
            {
                for(int i=0;i<ListWoSelecting.Items.Count;i++)
                {
                    ListWoSelecting.SelectedIndex = i;
                    WO = WO + ListWoSelecting.SelectedItem.ToString() + ",";
                }
                WO = WO.Substring(0,WO.Length-1);
                dt = Report.AOIQtySummary(WO);
            }
            if(dt != null &&  dt.Rows.Count>0)
            {
                //CopyToExcel("", "", "", "", "", "", dt, null);
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
            }
            else
            {
                MessageBox.Show("No data found");
            }

        }
        public void ReturnDID(string CboReportType,string CboGroupID)
        {
            DataTable dt = new DataTable();
            dt = Report.ReturnDID(CboReportType, CboGroupID);
            if(dt != null &&  dt.Rows.Count>0)
            {
                //CopyToExcel("", "", "", "", "", "", dt, null);
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
            }
            else
            {
                MessageBox.Show("NO data");
            }
        }
        public void Sap_Return(string Report_Type)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            string sDateTime="", eDateTime = "";
            if (Report_Type== "GroupIDCostQty")
            {
                if(CboGroupID.Text=="")
                {
                    MessageBox.Show("Please choose GroupID first!");
                    return;
                }
            }
            if (Report_Type == "DispatchDID")
            {
                if (dtpSDate.Value.ToString() == dtpEDate.Value.ToString())
                {
                    if(CboShift.Text=="")
                    {
                        MessageBox.Show("Please choose shift first!");
                        return;
                    }
                    if(CboShift.Text == "Day_Shift")
                    {
                        sDateTime = dtpSDate.Value.ToString("yyyyMMdd")+ "0740";
                        eDateTime= dtpEDate.Value.ToString("yyyyMMdd") + "1940";
                    }
                    else
                    {
                        sDateTime = dtpSDate.Value.ToString("yyyyMMdd") + "1940";
                        eDateTime = dtpEDate.Value.AddDays(1).ToString("yyyyMMdd") + "0740";
                    }                   
                }
                else
                {
                    if((dtpSDate.Value-dtpEDate.Value).TotalDays>1 && CboGroupID.Text=="")
                    {
                        MessageBox.Show("Date Range can not over 1 days when did not choose one wo group!");
                        return;
                    }
                    sDateTime = dtpSDate.Value.ToString("yyyyMMdd") + "0740";
                    eDateTime = dtpEDate.Value.ToString("yyyyMMdd") + "0740";
                }
            }
            if(Report_Type == "ReturnDID_ByDate")
            {
                sDateTime = dtpSDate.Value.ToString("yyyyMMdd") + "000000";
                eDateTime = dtpEDate.Value.ToString("yyyyMMdd") + "240000";
            }
            dt = Report.Sap_Return(CboWo.Text, CboComp.Text, CboGroupID.Text, Report_Type, sDateTime, eDateTime);
            if (dt != null &&  dt.Rows.Count > 0)
            {

                // CopyToExcel("", "", "", "", "", "", dt, null);
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);

            }
            else
            {
                MessageBox.Show("No data");
            }
        }
        public void CheckBOM_Rate()
        {
            string BeginDate, EndDate;            
            BeginDate = dtpSDate.Value.ToString("yyyyMMdd") + "000000";
            EndDate = dtpEDate.Value.ToString("yyyyMMdd")+ "240000";
            DataSet ds = new DataSet();
            ds = Report.GetCheckBomData("", "", "", "CheckBOM_Rate", BeginDate, EndDate);
            if (ds != null &&  ds.Tables[0].Rows.Count > 0)
            {
                //OfficeExcel(ds.Tables[0], ds.Tables[0].Rows.Count, ds.Tables[0].Columns.Count);
                CopyToExcel("CheckBOM_Rate", "", "", "", "", "", null, ds);
                
            }
        }
        public void RefreshBOM(string wo)
        {
            //DataSet ds = new DataSet();
            //ds = Report.GetCheckBomData(wo, "", "", "RefreshBOM","","");
            //if (ds.Tables[0].Rows.Count > 0)
            //{
            //    MessageBox.Show(ds.Tables[0].Rows[0]["Msg"].ToString());
            //    if (ds != null && ds.Tables[0].Rows[0]["Result"].ToString() == "0")
            //    {
            //        //CopyToExcel("", "", "", "", "", "", ds.Tables[1], null);
            //        OfficeExcel(ds.Tables[1], ds.Tables[1].Rows.Count, ds.Tables[1].Columns.Count);
            //    }
            //}
            DataSet ds01 = new DataSet();
            ds01 = Report.CheckQSMSWO(wo);
            if (ds01 == null)
            {
                MessageBox.Show("Please check bom first");
                return;
            }

            DataSet ds02 = new DataSet();
            ds02 = Report.CheckBom(wo);
            if (ds02 != null)
            {
                MessageBox.Show("Check bom fail");
            }

            DataSet ds03 = new DataSet();
            ds03 = Report.GetSAPBOMFailInfo(wo);
            if (ds03 != null)
            {
                OfficeExcel(ds03.Tables[0], ds03.Tables[0].Rows.Count, ds03.Tables[0].Columns.Count);
            }
            else
            {
                MessageBox.Show("refresh BOM OK");
            }

        }
        public void GetChkBOMDiff(string WO)
        {
            DataTable dt = new DataTable();
            if(WO=="")
            {
                MessageBox.Show("Please select WO");
                return;
            }
            dt = Report.QSMS_Wo_Diff(WO);
            if(dt != null &&  dt.Rows.Count>0)
            {
                //CopyToExcel("", "", "", "", "", "", dt, null);
                OfficeExcel(dt, dt.Rows.Count,dt.Columns.Count);
            }
            else
            {
                MessageBox.Show("No data");
            }
        }
        public void CheckBOM()
        {
            DataTable dt = new DataTable();
            string DualModel="", BomTest="";
            if (chkDual.Checked == true)
            {
                DualModel = "Y";
            }
            else
            {
                DualModel = "N";
            }
            if (pubFunction.ConfigListGetValue("CheckCycleTimeBU") != "Y")
            {
                if (pubFunction.ConfigListGetValue("CheckCycleTime") == "Y")
                {
                    dt = Report.QSMS_CheckBOM_CheckCycleTime(TxtWO.Text);
                    if (dt.Rows.Count > 0)
                    {
                        if (int.Parse(dt.Rows[0]["Result"].ToString()) != 0)
                        {
                            MessageBox.Show(dt.Rows[0]["Descr"].ToString());
                        }
                    }
                }
            }
                dt = Report.qsms_wogroup(TxtWO.Text);
                if(dt.Rows.Count>0)
                {
                    MessageBox.Show("This wo had been closed !");
                    return;
                }
                BomTest = TxtRev.Text.Trim();
                if(pubFunction.ConfigListGetValue("CheckBomLogon")=="Y")
                    {
                     if (pubFunction.ConfigListGetValue("CheckBomRight") == "Y")
                    //if (Parameter.CheckBomRight == false)
                    {
                        MessageBox.Show("You have no right to check bom!");
                        return;
                    }
                    else
                    {
                        GetCheckBomData(TxtWO.Text, Parameter.UID, DualModel, "GetCheckBomData");
                    }
                }
                else
                {
                    GetCheckBomData(TxtWO.Text, Parameter.UID, DualModel, "GetCheckBomData");
                }
            
        }
        public void GetCheckBomData(string Work_Order, string g_userName,string  DualModel,string flag)
        {
            string BuildType = "", chkBOMPass="", Pilot="";
            bool Negative = false;
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            //ds = Report.GetCheckBomData(Work_Order, g_userName, DualModel, flag, "", "");
            //if(ds.Tables[0].Rows.Count>0)
            //{
            //    MessageBox.Show(ds.Tables[0].Rows[0]["Msg"].ToString());
            //    if(ds != null && ds.Tables[0].Rows[0]["Result"].ToString()=="0")
            //    {
            //        CopyToExcel("", "", "", "", "", "", ds.Tables[1], null);
            //    }
            //}
            if(Work_Order.Trim()=="")
            {
                MessageBox.Show("Please check the WO");
                return;
            }
            else
            {
                dt = Report.QSMS_RegisterCheckBOM(Work_Order, "0");
                if(dt.Rows.Count>0)
                {
                    if(dt.Rows[0]["rtnCode"].ToString()=="0")
                    {
                        MessageBox.Show("Now,SomeBody is doing CheckBOM in computer "+ dt.Rows[0]["hostname"].ToString() + ", The system don't allow more than one person to do CheckBOM");
                        return;
                    }
                }
            }
            Report.delSap_BOM_Fail(Work_Order);
            Report.InsQSMS_LOG(Work_Order, "CheckBOM", Parameter.g_userName,"1");
            dt = Report.QSMS_CheckBOM(Work_Order);
            if(dt.Rows.Count==0)
            {
                dt = Report.QSMS_RegisterCheckBOM(Work_Order,"0");
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["rtnCode"].ToString() == "0")
                    {
                        MessageBox.Show("Now,SomeBody is doing CheckBOM in computer " + dt.Rows[0]["hostname"].ToString() + ", The system don't allow more than one person to do CheckBOM");
                        return;
                    }
                }
            }
            dt = Report.GETBuildType(Work_Order);
            if(dt.Rows.Count==0)
            {
                MessageBox.Show("PMC didn't release the WO,Please Check");
                Report.InsQSMS_LOG(Work_Order, "CheckBOM", Parameter.g_userName,"4");
                return;
            }
            else
            {
                BuildType = dt.Rows[0]["BuildType"].ToString();
                if(BuildType!="1" && BuildType!="2" && BuildType!="3" && BuildType!="4")
                {
                    MessageBox.Show("BuildType Error,Please call QMS");
                    Report.InsQSMS_LOG(Work_Order, "CheckBOM", Parameter.g_userName, "4");
                    return;
                }
            }
            dt = Report.QSMS_CheckBomSP(Work_Order, BuildType, DualModel);
            if(dt.Rows.Count>0)
            {
                dt = Report.Pilot(Work_Order);
                if(dt.Rows.Count>0)
                {
                    string MBPN = "";
                    MBPN = dt.Rows[0]["PN"].ToString();
                    Pilot= dt.Rows[0]["Pilot"].ToString();
                    dt = Report.Negative(MBPN);
                    if(dt.Rows.Count>0)
                    {
                        Negative = true;
                    }
                    else
                    {
                        Negative = false;
                    }
                }
                if(Negative==true && Pilot.ToUpper()=="NEW")
                {
                    dt = Report.QSMS_RegisterCheckBOM(Work_Order,"1");
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["rtnCode"].ToString() == "0")
                        {
                            MessageBox.Show("Clear the WorkOder from  XL_CheckBOM fail,Please check XL_CheckBOM");
                            
                        }
                    }
                    Report.InsQSMS_LOG(Work_Order, "CheckBOM", Parameter.g_userName, "3");
                    Report.Updateqsms_error_log(Work_Order);
                    Report.InsQSMS_LOG(Work_Order, "CheckBOMResult", Parameter.g_userName, "0");
                }
                MessageBox.Show("Check bom fail");
                chkBOMPass = "N";
            }
            else
            {
                chkBOMPass = "Y";
            }

           if(chkBOMPass=="Y")
            {
                Report.Updateqsms_error_logPASS(Work_Order);
                Report.InsQSMS_LOG(Work_Order, "CheckBOMResult", Parameter.g_userName, "0");
            }
           else
            {
                Report.InsQSMS_LOG(Work_Order, "CheckBOMResult", Parameter.g_userName, "-1");
            }
            dt = Report.GetSap_BOM_Fail(Work_Order);
            if(dt.Rows.Count>0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                //CopyToExcel("", "", "", "", "", "", dt, null);
            }
            else
            {
                Report.QSMS_ReCountDispatchQty(Work_Order,"1");
                MessageBox.Show("Check Bom OK");
            }
            dt = Report.QSMS_RegisterCheckBOM(Work_Order, "1");
            if(dt.Rows.Count>0)
            {
                if (dt.Rows[0]["rtnCode"].ToString() == "0")
                {
                    MessageBox.Show("Clear the WorkOder from  XL_CheckBOM fail,Please check XL_CheckBOM");

                }
            }
            Report.delSap_BOM_Fail(Work_Order);
            Report.InsQSMS_LOG(Work_Order, "CheckBOM", Parameter.g_userName, "2");
        }

        public  void GetReplacePN(string WO)
        {
            DataTable dt = new DataTable();
            if(WO=="")
            {
                MessageBox.Show("Please Input WO !!!");
                return;
            }
            dt = Report.QSMS_GetReplacePNByWOList(WO);
            if(dt != null &&  dt.Rows.Count>0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                //CopyToExcel("", "", "", "", "", "", dt, null);
            }
            else
            {
                MessageBox.Show("no data ");
            }
        }
        public void GetMEBom_WO(string Work_Order)
        {
            DataTable dt = new DataTable();
            if(TxtWO.Text=="")
            {
                MessageBox.Show("Please Input WO !!!");
                return;
            }
            if (pubFunction.ConfigListGetValue("ME_BOM_WO") == "Y")//需要修改
                {
                dt = Report.QSMSRpt_QSMS_WO(Work_Order, "ME_BOM_WO");
                if(dt.Rows.Count>0)
                {
                    OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                    //CopyToExcel("", "", "", "", "", "", dt, null);
                   
                }
                else
                {
                    MessageBox.Show("No Data in current and history !!");
                    return;
                }
            }
            else
            {
                dt = Report.QSMS_WOCurrent(Work_Order);
                if(dt != null && dt.Rows.Count>0)
                {
                    OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                    //CopyToExcel("", "", "", "", "", "", dt, null);
                }
                else
                {
                    dt = Report.QSMS_WOHis(Work_Order);
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                        //CopyToExcel("", "", "", "", "", "", dt, null);
                    }
                    else
                    {
                        MessageBox.Show("No Data in current and history !!");
                        return;
                    }

                }
            }
        }
        public void MEBOM_Delete_Log(string DID)
        {
            DataTable dt = new DataTable();
            if(TxtMBPN.Text=="")
            {
                MessageBox.Show("Please Input 'MB/Job PN' !!!");
                return;
            }
            dt = Report.QSMS_Log(DID);
            if(dt != null && dt.Rows.Count>0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                //CopyToExcel("", "", "", "", "", "", dt, null);
                return;
            }
            else
            {
                MessageBox.Show("No data found!!");
                return;
            }
        }
        public void GETGROUPIDDATABYCOMPPN(string GroupID, string compPN )
        {
            DataSet Ds = new DataSet();
            Ds =Report.GetGroupIDDataByCompPN(GroupID, compPN);
            //if(Parameter.chkDomain=="N")
            //{

            //}
            if (Ds != null)
            {
                OfficeExcel(Ds.Tables[0], Ds.Tables[0].Rows.Count, Ds.Tables[0].Columns.Count);
                OfficeExcel(Ds.Tables[1], Ds.Tables[1].Rows.Count, Ds.Tables[1].Columns.Count);
                OfficeExcel(Ds.Tables[2], Ds.Tables[2].Rows.Count, Ds.Tables[2].Columns.Count);
                //CopyToExcel("", "", "", "", "", "", Ds.Tables[0], null);
                //CopyToExcel("", "", "", "", "", "", Ds.Tables[1], null);
                //CopyToExcel("", "", "", "", "", "", Ds.Tables[2], null);
            }
           
            if(Ds != null && Ds.Tables[3] !=null)
            {
                OfficeExcel(Ds.Tables[3], Ds.Tables[3].Rows.Count, Ds.Tables[3].Columns.Count);
                //CopyToExcel("", "", "", "", "", "", Ds.Tables[3], null); //需要测试
            }
           
        }
        public void CheckWO_WastagePN(string Work_Order)
        {
            DataTable dt = new DataTable();
            if(TxtWO.Text=="")
            {
                MessageBox.Show("Please Input WO !!!");
                return;
            }
            dt = Report.CompPN_DataCurrent(Work_Order);
            if(dt != null && dt.Rows.Count>0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                ///CopyToExcel("", "", "", "", "", "", dt, null);
            }
            else
            {
                dt = Report.CompPN_DataHis(Work_Order);
                if (dt != null && dt.Rows.Count > 0)
                {
                    OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                    // CopyToExcel("", "", "", "", "", "", dt, null);
                }
                else
                {
                    MessageBox.Show("No Data in current and history !!");
                }
            }
        }
        public void GetMEBom(string Work_Order)
        {
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            string machine="",  TempGroup="",  TempJObGroup="",line="";
            string[] MultiLineStr;
            if (TxtWO.Text!="")
            {
                dt = Report.sap_wo_list(Work_Order);
                if(dt.Rows.Count>0)
                {
                    TempGroup = dt.Rows[0]["Group"].ToString();
                    TempJObGroup = GetJobGroup(TempGroup);
                    if(int.Parse(dt.Rows[0]["BuildType"].ToString())== 4)
                    {
                        machine = GetMultiLine(Work_Order, "WO");
                    }
                    else
                    {
                        machine = dt.Rows[0]["Line"].ToString()+"%";
                    }
                }
            }
            else
            {
                TempJObGroup = GetSelectingJobGroup();
                if(CboMachine.Text=="ALL")
                {
                    machine = CboLine.Text + "%";
                }
                else
                {
                    if(CboMachine.Text.Length>0)
                    {
                        CboLine.Text = CboMachine.Text.Substring(0, 1);
                    }
                    machine = CboMachine.Text + "%";
                }
                if(CboLine.Text=="")
                {
                    MessageBox.Show("Please Check line");
                    return;
                }
               
            }
            if(TempJObGroup == "")
            {
                if(CboMachine.Text=="")
                {
                    MessageBox.Show("请选择Machine");
                    return;
                }
                MessageBox.Show("请选择 JobGroup or work order");
                return;
            }
            if(machine.IndexOf(",")>0)//需要测试
            {
                MultiLineStr = machine.Split(',');
                foreach(string i in MultiLineStr)
                {
                    if(i.Substring((i.Length-1),1)=="Q")
                    {
                        machine = machine + "((a.machine like '" + i.Substring(0, 1) + "Q%' OR a.machine like '" + i.Substring(0, 1) + "W%') and (a.Side like 'Q%' or a.Side like 'W%' )) OR ";
                    }
                    else
                    {
                        machine = machine + "(a.machine like '" + i.Substring(0, 1) + "S%' and a.Side like '" + i.Substring((i.Length - 1), 1) + "%') OR ";
                    }
                }
                machine = machine.Substring(0, machine.Length - 3);
                machine = "(" + machine + ")";
            }
            else
            {
                machine = "a.machine like '" + machine + "'";
            }
            if(TxtWO.Text!="")
            {
                DataTable dts = new DataTable();
                dts = Report.WO_MultiLine(TxtWO.Text);
                if (dts.Rows.Count > 0)
                {
                    line = dts.Rows[0]["Line"].ToString();
                    dt1 = Report.GetMEBom(dt.Rows[0]["Line"].ToString(),line, TxtRev.Text, TxtWO.Text, TempJObGroup);
                }
                else
                {
                    dt1 = Report.GetMEBom(dt.Rows[0]["Line"].ToString(), "", TxtRev.Text, TxtWO.Text, TempJObGroup);
                }
            }
            else
            {
                dt1 = Report.GetMEBom(CboLine.Text, "", TxtRev.Text, "", TempJObGroup);
            }
            if(dt1 != null && dt1.Rows.Count>0)
            {
                OfficeExcel(dt1, dt1.Rows.Count, dt1.Columns.Count);
               // CopyToExcel("", "", "", "", "", "", dt1, null);
            }
            else
            {
                MessageBox.Show("No data found,Please check");
                return;
            }
        }
        public string GetSelectingJobGroup()
        {
            string SelectingJobGroup = "", GetSelectingJobGroup="";
            if (ListselectingJobGroup.Items.Count<=0)
            {
                if(TxtJobGroup.Text.Length>0 && TxtJobGroup.Text.IndexOf("-")>0)
                {
                    GetSelectingJobGroup = "(" + TxtJobGroup.Text + ")";
                }
                else
                {
                    SelectingJobGroup = "";
                }
            }
            for(int i=1;i<ListselectingJobGroup.Items.Count;i++)
            {
                ListselectingJobGroup.SelectedIndex = i - 1;
                SelectingJobGroup = SelectingJobGroup + "'" + ListselectingJobGroup.SelectedItem.ToString() + "'" + ",";
            }
            GetSelectingJobGroup = "(" + SelectingJobGroup.Substring(0, SelectingJobGroup.Length - 1) + ")";
            return GetSelectingJobGroup;
        }
        public string GetMultiLine(string sInputStr, string sType )
        {
            string sLineList = "", GetMultiLine="";
            DataTable dt = new DataTable();
            if (sType=="WO")
            {
                dt = Report.WO_MultiLine(sInputStr);
            }
            if(dt.Rows.Count>0)
            {
                for(int i=0;i<dt.Rows.Count;i++)
                {
                    sLineList = sLineList + dt.Rows[i]["Line"].ToString()+",";
                }
            }
            if(sLineList!="")
            {
                sLineList = sLineList.Substring(0, sLineList.Length - 1);
            }
            GetMultiLine = sLineList;
            return GetMultiLine;
        }
        public string  GetJobGroup(string Group)
        {
            string Jobpn="", jobgroup="", GetJobGroup="";
            DataTable dt = new DataTable();
            dt = Report.QSMS_GetEMSFlag(Group);
            if(dt.Rows.Count==0 || dt.Rows[0]["EMSFlag"].ToString().ToUpper()== "NONE")
            {
                GetJobGroup = "";
                return GetJobGroup;                
            }
            if(dt.Rows[0]["EMSFlag"].ToString().ToUpper()!="Y")
            {
                dt = Report.QSMS_GetQSMS_JobBomMB(Group);
                if(dt.Rows.Count==0)
                {
                    dt = Report.QSMS_GetQSMS_JobBom(Group);
                    if(dt.Rows.Count==0)
                    {
                        GetJobGroup = "";
                        return GetJobGroup;
                    }
                    else
                    {
                        for(int i=0;i<dt.Rows.Count;i++)
                        {
                            jobgroup = dt.Rows[i]["Jobpn"].ToString().Trim()+'-' + dt.Rows[i]["MB_Rev"].ToString().Trim();
                            Jobpn = Jobpn + "'" + jobgroup + "'" + ",";
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        jobgroup = dt.Rows[i]["Jobpn"].ToString().Trim()+ "-" + dt.Rows[i]["MB_Rev"].ToString().Trim();
                        Jobpn = Jobpn + "'" + jobgroup + "'" + ",";
                    }
                }
            }
            else
            {
                dt = Report.QSMS_GetQSMS_JobBomEMS(Group);
                if(dt.Rows.Count==0)
                {
                    GetJobGroup = "";
                    return GetJobGroup;
                }
                else
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        jobgroup = dt.Rows[i]["Jobpn"].ToString().Trim() + "-" + dt.Rows[i]["MB_Rev"].ToString().Trim();
                        Jobpn = Jobpn + "'" + jobgroup + "'" + ",";
                    }
                }
            }
            Jobpn = Jobpn.Substring(0, Jobpn.Length - 1);
            Jobpn = "(" + Jobpn + ")";
            GetJobGroup = Jobpn;
            return GetJobGroup;
        }
        public void GetSapGroupByWo(string WO)
        {
            DataTable dt = new DataTable();
            if(TxtWO.Text=="")
            {
                MessageBox.Show("Please check the WO");
                return;
            }
            dt = Report.GetSapGroupByWo(WO);
            if (dt.Rows.Count > 0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                //CopyToExcel("", "", "", "", "", "", dt, null);
            }
            else
            {
                MessageBox.Show("No data found");
            }
        }
        public void GetSapBom(string txtWo)
        {
            string strWO = "";
            DataTable dt = new DataTable();
            if(txtFilePath.Text!="" && ListWoSelecting.Items.Count>0)
            {
                strWO = GetWO("BY_WORKORDERS", "N");
                dt = Report.QSMS_QuerySapBom(strWO);
            }
            else
            {
                dt = Report.sap_bom(txtWo);
            }
            if(dt != null && dt.Rows.Count>0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                //CopyToExcel("", "", "", "", "", "", dt, null);
            }
            else
            {
                MessageBox.Show("No data found <GetSapBom> !");
            }
            txtFilePath.Text = "";
            ListWoSelecting.Items.Clear();
        }
        public void PDUsedByCompLine(string CboLine, string CboComp)
        {
            DataTable dt = new DataTable();
            dt = Report.qsms_verify(CboLine, CboComp);
            if(dt != null && dt.Rows.Count>0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                //CopyToExcel("", "", "", "", "", "", dt, null);
            }
            else
            {
                MessageBox.Show("NO data");
                return;
            }
        }
        public void CopyToExcelWipDifferentMaterial(string Sheetname, string Work_Order)
        {
            string LocalPath = "D:\\QSMS_Report\\";
            string TransDate = DateTime.Now.ToString("yyyyMMdd");
            string strFileName = LocalPath + "SMTMaterialReport_" + Work_Order + ".xls";
            DataTable  dt = new DataTable();
            if (!System.IO.Directory.Exists(LocalPath))
            {
                System.IO.Directory.CreateDirectory(LocalPath);
            }
            lBLmESSAGE.Text = "";
            dt = Report.QSMSRptWipDifferentMaterial(Work_Order);
            if (dt.Rows.Count <= 0)
            {
                MessageBox.Show("NO data");
                return;
            }
            else
            {
                System.IO.File.Copy(AppDomain.CurrentDomain + "\\SMTMaterialReport.xls", strFileName);
                CopyToExcel("MaterialDifferentList", strFileName, Sheetname, TransDate, "", "", dt, null);
                lBLmESSAGE.Text = "Report OK" + strAddress;
            }
        }
        public void CopyToExcelWipLackByWo(string Sheetname,string Work_Order)
        {
            string LocalPath = "D:\\QSMS_Report\\";
            string TransDate = DateTime.Now.ToString("yyyyMMdd");
            string strFileName = LocalPath + "SMTMaterialReport_" + Work_Order + ".xls";
            DataSet dS = new DataSet();
            if (!System.IO.Directory.Exists(LocalPath))
            {
                System.IO.Directory.CreateDirectory(LocalPath);
            }
            lBLmESSAGE.Text = "";
            dS = Report.QSMSRptLackCompByWo(Work_Order);
            if (dS != null && dS.Tables[1].Rows.Count > 0)
            {
                System.IO.File.Copy(AppDomain.CurrentDomain + "\\SMTMaterialReport.xls", strFileName);
                CopyToExcel("CopyToExcelWipLackByWo", strFileName, Sheetname, "", "", "", dS.Tables[1], dS);
                lBLmESSAGE.Text = "Report OK" + strAddress;
                
            }
            else
            {
                MessageBox.Show("NO data");
                return;
            }
        }
        public void CopyToExcelWipByGroup(string Sheetname, string GroupID)
        {
            string LocalPath = "D:\\QSMS_Report\\";
            string TransDate = DateTime.Now.ToString("yyyyMMdd");
            string strFileName = LocalPath + "SMTMaterialReport_" + GroupID + ".xls";
            DataSet dS = new DataSet();
            if (!System.IO.Directory.Exists(LocalPath))
            {
                System.IO.Directory.CreateDirectory(LocalPath);
            }
            lBLmESSAGE.Text = "";
            dS = Report.QSMSRptWipByGroup(GroupID);
            if (dS != null && dS.Tables[1].Rows.Count > 0)
            {
                System.IO.File.Copy(AppDomain.CurrentDomain + "\\SMTMaterialReport.xls", strFileName);
                CopyToExcel("CopyToExcelWipByGroup", strFileName, Sheetname, GroupID, "", "", dS.Tables[1], dS);
                lBLmESSAGE.Text = "Report OK" + strAddress;

            }
            else
            {

                MessageBox.Show("NO data");
                return;
            }
        }
        public void CopyToExcelWipByDate(string Sheetname)
        {
            string LocalPath = "D:\\QSMS_Report\\";
            string TransDate = DateTime.Now.ToString("yyyyMMdd");
            string strFileName = LocalPath + "SMTMaterialReport_" + TransDate + ".xls";
            DataSet dS = new DataSet();
            if (!System.IO.Directory.Exists(LocalPath))
            {
                System.IO.Directory.CreateDirectory(LocalPath);
            }
            lBLmESSAGE.Text = "";
            dS = Report.QSMSRptWipBydate();
            if (dS != null &&  dS.Tables[1].Rows.Count> 0)
            {
                System.IO.File.Copy(AppDomain.CurrentDomain + "\\SMTMaterialReport.xls", strFileName);
                CopyToExcel("CopyToExcelWipByDate", strFileName, Sheetname, TransDate, "", "", dS.Tables[1], dS);
                lBLmESSAGE.Text = "Report OK" + strAddress;
                
            }
            else
            {
                MessageBox.Show("NO data");
                return;
            }
        }
        public void CopyToExcelWipByMaterial(string Sheetname, string CompPN)
        {
            string LocalPath = "D:\\QSMS_Report\\";       
            string strFileName = LocalPath + "SMTMaterialReport_" + CompPN + ".xls";            
            DataSet dS = new DataSet();
            if (!System.IO.Directory.Exists(LocalPath))
            {
                System.IO.Directory.CreateDirectory(LocalPath);
            }
            lBLmESSAGE.Text = "";
            dS = Report.CopyToExcelWipByMaterial(CompPN);
            if (dS != null && dS.Tables[1].Rows.Count >0)
            {
                System.IO.File.Copy(AppDomain.CurrentDomain + "\\SMTMaterialReport.xls", strFileName);
                CopyToExcel("CopyToExcelWipByMaterial", strFileName, Sheetname, CompPN, "", "", dS.Tables[1], dS);
                lBLmESSAGE.Text = "Report OK" + strAddress;
                
            }
            else
            {
                MessageBox.Show("NO data");
                return;
            }

        }
        public void QSMS_WO()
        {
            DataTable dt = new DataTable();
            string strWO = "";
            if(txtFilePath.Text!="" && ListWoSelecting.Items.Count>0)
            {
                
                strWO = GetWO("BY_WorkOrders","N");
                dt = Report.QSMSRpt_QSMS_WO(strWO,"");
            }
            if (dt != null && dt.Rows.Count > 0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
               // CopyToExcel("", "", "", "", "", "", dt,null);//需要改excel格式
            }
            else
            {
                MessageBox.Show("No data");
            }
        }
      
        public void QSMS_DID_ToWH()
        {
            DataTable dt = new DataTable();
            string BeginDate, EndDate;
            BeginDate = dtpSDate.Value.ToString("yyyyMMdd");
            EndDate = dtpEDate.Value.AddDays(1).ToString("yyyyMMdd");//这里需要看时间是否一致
            if (dtpSDate.Value > dtpEDate.Value.AddDays(1))//这里需要看时间是否一致
            {
                MessageBox.Show("The StartDate must be smaller than Today !");
                return;
            }
            dt = Report.QSMS_DID_ToWH(BeginDate, EndDate);
            if (dt != null && dt.Rows.Count > 0)
            {
                CopyToExcel("QSMS_DID_ToWH", AppDomain.CurrentDomain+"\\QSMS_ReturnDID_Summary.XLS", "QSMS_ReturnDID_Report", "", "", "", dt,null);//需要改excel格式
            }
            else
            {
                MessageBox.Show("No data found");
            }
        }
        public void DispatchQTYByWO()
        {
            string strWO = "";
            DataTable dt = new DataTable();
            strWO = TxtWO.Text;
            if (strWO=="")
            {
                MessageBox.Show("WO Can not be empty !");
                return;
            }
            dt = Report.QSMS_DispatchQTYByWO(strWO.Trim());
            if (dt != null && dt.Rows.Count > 0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                //CopyToExcel("", "", "", "", "", "", dt, null);//需要改excel格式
            }
            else
            {
                MessageBox.Show("No data found");
            }

        }
        public void XL_MonitorReport()
        {
            DataTable dt = new DataTable();
            string BeginDate, EndDate;
            BeginDate = dtpSDate.Value.ToString("yyyyMMdd");
            EndDate = dtpEDate.Value.AddDays(1).ToString("yyyyMMdd");//这里需要看时间是否一致
            if(dtpSDate.Value > dtpEDate.Value.AddDays(1))//这里需要看时间是否一致
            {
                MessageBox.Show("The StartDate must be smaller than Today !");
                return;
            }
            dt = Report.XL_MonitorReport(BeginDate);
            if (dt != null && dt.Rows.Count > 0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                //CopyToExcel("", "", "", "", "", "", dt, null);//需要改excel格式
            }
            else
            {
                MessageBox.Show("No data found");
            }
        }
        public void LineChangeStatistics()
        {
            DataTable dt = new DataTable();
            string BeginDate, EndDate;          
            BeginDate = dtpSDate.Value.ToString("yyyyMMdd");
            EndDate = dtpEDate.Value.ToString("yyyyMMdd");
            dt = Report.LineChangeStatistics(BeginDate, EndDate);
            if (dt != null && dt.Rows.Count > 0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                //CopyToExcel("", "", "", "", "", "", dt, null);//需要改excel格式
            }
            else
            {
                MessageBox.Show("No data found");
            }
        }
        public void DIDDeleteRecords()
        {
            DataTable dt = new DataTable();
            string BeginDate, EndDate;            
            BeginDate = dtpSDate.Value.ToString("yyyyMMdd");
            EndDate = dtpEDate.Value.ToString("yyyyMMdd");
            dt = Report.DIDDeleteRecords(BeginDate, EndDate);
            if (dt != null && dt.Rows.Count > 0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                //CopyToExcel("", "", "", "", "", "", dt, null);
            }
            else
            {
                MessageBox.Show("No data found");
            }
        }
        public void LineChangeStatisticsByall()
        {
            DataTable dt = new DataTable();
            string BeginDate, EndDate;          
            BeginDate = dtpSDate.Value.ToString("yyyyMMdd");         
            EndDate = dtpEDate.Value.ToString("yyyyMMdd");
            dt = Report.GenChangeLineReport2(BeginDate, EndDate);
            if(dt != null && dt.Rows.Count>0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                // CopyToExcel("", "", "", "", "", "", dt, null);
            }
            else
            {
                MessageBox.Show("No data found");
            }
          
        }
        public void Load_CheckReplacePNBySAPBOM(string  Shift_Item)
        {
            string  strCompPN = "", strUID;
            int tmpRow;
            DataTable dt = new DataTable();
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xBk = appExcel.Workbooks.Open(txtFilePath.Text.Trim()); ;
            Microsoft.Office.Interop.Excel.Worksheet xSt;
            //xBk = appExcel.Workbooks.Add(true);
            xSt = (Microsoft.Office.Interop.Excel.Worksheet)xBk.ActiveSheet;
            appExcel.Visible = true;
            xSt = (Microsoft.Office.Interop.Excel.Worksheet)xBk.Worksheets.get_Item(1);
            Report.DelQSMS_CompPNcheck_Temp();
            tmpRow = 2;
            strUID = Parameter.UID;
            while (xSt.Cells[tmpRow, 1].ToString() != "")
            {
                strCompPN = xSt.Cells[tmpRow, 1].ToString();                
                dt = Report.Getsapbom( strCompPN);
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("this MBPN '" + strCompPN + "' have not exist in SAPBOM,please chceck it;");
                    return;
                }
               Report.QSMS_CompPNcheck_Temp("", strCompPN,"");
               tmpRow = tmpRow + 1;
            }
            dt = Report.QSMS_QueryReplacePN();
            if (dt != null && dt.Rows.Count > 0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                //CopyToExcel("", "", "", "", "", "", dt, null);
            }
            else
            {
                MessageBox.Show("Those PN have no different ReplacePN!");
            }
        }
        public void Load_QSMS_CheckCompPN(string Shift_Item)
        {
            string strJobPN="", strCompPN = "", strUID;
            int tmpRow;
            DataTable dt = new DataTable();       
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xBk= appExcel.Workbooks.Open(txtFilePath.Text.Trim()); ;
            Microsoft.Office.Interop.Excel.Worksheet xSt;
            //xBk = appExcel.Workbooks.Add(true);
            xSt = (Microsoft.Office.Interop.Excel.Worksheet)xBk.ActiveSheet;
            appExcel.Visible = true;
            xSt = (Microsoft.Office.Interop.Excel.Worksheet)xBk.Worksheets.get_Item(1);
            Report.DelQSMS_CompPNcheck_Temp();
            tmpRow = 2;
            strUID = Parameter.UID;
            while(xSt.Cells[tmpRow,1].ToString()!="" && xSt.Cells[tmpRow, 2].ToString() != "")
            {
                strJobPN = xSt.Cells[tmpRow, 1].ToString();
                strCompPN = xSt.Cells[tmpRow, 2].ToString();
                dt = Report.QSMS_CheckCompPN(strJobPN, strCompPN, strUID,"W");
                if(dt != null && dt.Rows.Count>0)
                {
                    if(dt.Rows[0]["result"].ToString()!="0")
                    {
                        MessageBox.Show("Err:"+ dt.Rows[0]["desc1"].ToString());
                        MessageBox.Show("*** Load  Finish ! ***");
                        return;
                    }
                }
                tmpRow = tmpRow + 1;
            }
            dt = Report.QSMS_CheckCompPN(strJobPN, strCompPN, strUID,"");
            if(dt != null && dt.Rows.Count>0)
            {
                OfficeExcel(dt, dt.Rows.Count, dt.Columns.Count);
                //CopyToExcel("", "", "", "", "", "", dt, null);
            }
            else
            {
                MessageBox.Show("No data");
            }
        }
        public void PrepareMaterialByWONew( string SheetName, string Line)
        {
            DataSet ds = new DataSet();
            string WO = "";            
            if (SheetName == "By_WorkOrder")
            {
                WO = GetWO("BY_WorkOrder", "N");
                if(WO=="")
                {
                    MessageBox.Show("Please select the work order");
                    return;
                }
                ds = Report.QSMSRptPrepareMaterialByWO(WO);
            }
            else if(SheetName == "By_Shift")
            {
                if(CboShift.Text=="")
                {
                    MessageBox.Show("Please select Shift");
                    return;
                }
                string shift = CboShift.Text.Substring(0, 1);
                if(shift!="D" && shift!="N")
                {
                    MessageBox.Show("Wrong Shift for you select");
                    return;
                }
                WO=GetWO("BY_Shift", "N");
                if (WO == "")
                {
                    
                    return;
                }
                ds = Report.QSMSRptPrepareMaterialByLineShift(WO);
            }
            else if(SheetName== "By_WorkOrders")
            {
                WO = GetWO("BY_WorkOrders", "N");
                if (WO == "")
                {

                    return;
                }
                ds = Report.QSMSRptPrepareMaterialByGroup(WO);
              
            }
            else if (SheetName == "By_Group")
            {
                WO = GetWO("By_Group", "N");
                if (WO == "")
                {

                    return;
                }
                ds = Report.QSMSRptPrepareMaterialByGroup(WO);

            }
            else if (SheetName == "By_JobPN")
            {
                WO = GetWO("BY_WorkOrders", "N");
                if (CboJobPN.Text == "")
                {
                    MessageBox.Show("Please select jobpn");
                    return;
                }
                else
                {
                    string jobpn = CboJobPN.Text.Substring(0, 11);
                    ds = Report.QSMSRptPrepareMaterialByJobPN(WO, jobpn);
                }

                
            }
            if (ds != null )
            {
                ToExcel(ds);
            }
            


        }
        public string GetWO(string Ctype,string cOutPut)
        {
            string Shift = "";
            string line="";
            string GetWO = "";
            string WO = "";
            string WoOutPut = "";
            string sDateTime, eDateTime;
            DataTable dt = new DataTable();
            sDateTime = dtpSDate.Value.ToString("yyyy/MM/dd");
            sDateTime = sDateTime.Replace("-", "");
            sDateTime = sDateTime.Replace("/", "");
            sDateTime = sDateTime + "0000";
            eDateTime = dtpEDate.Value.ToString("yyyy/MM/dd"); ;
            eDateTime = eDateTime.Replace("-", "");
            eDateTime = eDateTime.Replace("/", "");
            eDateTime = eDateTime + "2400";
            if(Ctype.Trim().ToUpper()== "BY_WORKORDERS")
            {
                for(int i=0;i<ListWoSelecting.Items.Count;i++)
                {
                    ListWoSelecting.SelectedIndex = i;
                    WO = WO + ListWoSelecting.SelectedItem.ToString() + ",";
                    WoOutPut = WoOutPut + ListWoSelecting.Text + "\n" + "\r";
                }
                if(WO!="")
                {
                    WO = WO.Substring(0, WO.Length - 1);
                }
                else
                {
                    MessageBox.Show("Please select the work order list");
                    return "";
                }
            }
            else if(Ctype.Trim().ToUpper()== "BY_GROUP")
            {
                if(CboGroupID.Text=="")
                {
                    MessageBox.Show("Please select GroupID");
                    return "";
                }
                dt = Report.GetWork_Order(CboGroupID.Text.Trim());
                for(int i=0;i<dt.Rows.Count;i++)
                {
                    WO = WO+dt.Rows[i]["Work_Order"].ToString();
                    WoOutPut=WoOutPut+ dt.Rows[i]["Work_Order"].ToString() + "\n" + "\r"; ;
                }
                WO = WO.Substring(0, WO.Length - 1);
            }
            else if(Ctype.Trim().ToUpper()== "BY_WORKORDER")
            {
                WO = TxtWO.Text;
                WoOutPut = WO;
            }
            else if(Ctype.Trim().ToUpper()== "BY_SHIFT")
            {
                Shift = CboShift.Text.Substring(0, 1);
                line = CboLine.Text.Substring(0, 1);
                dt = Report.XL_WOPlanSeq(Shift,line,sDateTime,eDateTime);
                if(dt.Rows.Count<=0)
                {
                    MessageBox.Show("No data");
                    return "";
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    WO = WO + dt.Rows[i]["WO"].ToString();
                    WoOutPut = WoOutPut + dt.Rows[i]["WO"].ToString() + "\n" + "\r"; ;
                }
                WO = WO.Substring(0, WO.Length - 1);
            }
            if(cOutPut=="Y")
            {
                GetWO = WoOutPut;
            }
            else
            {
                GetWO = WO;
            }
             return GetWO;

        }

        public void CopyToExcelPrepareMaterialList(string Sheetname, string Wo, string Machine, string Line)
        {
            string LocalPath= "D:\\QSMS_Report\\";           
            string strFileName = LocalPath+ "SMTMaterialReport_"+Wo+ ".xls";           
            DataTable dt = new DataTable();
            if (!System.IO.Directory.Exists(LocalPath))
            {
                System.IO.Directory.CreateDirectory(LocalPath);
            }
            lBLmESSAGE.Text = "";
            dt = Report.QSMSRptPrepareMaterial(Wo, Machine);
            if(dt != null && dt.Rows.Count<=0)
            {

                MessageBox.Show("NO data");
                return;
            }
            else
            {
                System.IO.File.Copy(AppDomain.CurrentDomain + "\\SMTMaterialReport.xls", strFileName);
                CopyToExcel("PrepariMaterialList",strFileName, Sheetname, Line,Wo, Machine, dt, null);
            }

        }
        public void ToExcel(DataSet Ds)
        {
            string col1;
            int row, col,sheet=1;
            row = Ds.Tables[0].Rows.Count;
            col = Ds.Tables[0].Columns.Count - 1;
            col1 = GetColumnChar(col);
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xBk; //= appExcel.Workbooks.Open(FileName); ;
            Microsoft.Office.Interop.Excel.Worksheet xSt;
            xBk = appExcel.Workbooks.Add(true);
            xSt = (Microsoft.Office.Interop.Excel.Worksheet)xBk.ActiveSheet;
            appExcel.Visible = true;
            xSt = (Microsoft.Office.Interop.Excel.Worksheet)xBk.Worksheets.get_Item("Sheet" + Convert.ToString(sheet));
            if(Ds.Tables[1].Rows.Count>0)
            {
                sheet = sheet + 1;
                for(int i=0;i<Ds.Tables[1].Columns.Count;i++)
                {
                    xSt.Cells[1, i + 1] = Ds.Tables[0].Columns[i].ColumnName.ToString();
                    xSt.Cells[1, i + 1].Interior.Color = System.Drawing.Color.Yellow;
                    if (Ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "SLOT" || Ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "DESTSLOT" || Ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "COMPLEVEL")
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        range2.NumberFormat = "@";
                    }
                    else if (Ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper().IndexOf("DATE") > 0)
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        //range2.NumberFormat = "@";

                        range2.NumberFormat = "0";
                    }
                }
                for (int i = 0; i < Ds.Tables[0].Rows.Count; i++)
                {
                    for (int m = 0; m < Ds.Tables[0].Columns.Count; m++)
                    {
                        xSt.Cells[i+2, m + 1] = Ds.Tables[0].Rows[i][m].ToString();
                    }
                      
                }
                string col2 = col1 + Convert.ToString(row + 1);
                Microsoft.Office.Interop.Excel.Range range1 = xSt.Range["A1", col2];
                range1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range1.EntireColumn.AutoFit();
                range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//水平居中  
                range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//垂直居中

            }

            if (Ds.Tables[1].Rows.Count > 0)
            {
                row = Ds.Tables[1].Rows.Count;
                col = Ds.Tables[1].Columns.Count - 1;
                col1 = GetColumnChar(col);
                xSt = xBk.Sheets.Add();
                xSt = (Microsoft.Office.Interop.Excel.Worksheet)xBk.Worksheets.get_Item("Sheet" + Convert.ToString(sheet));
                xSt = (Microsoft.Office.Interop.Excel.Worksheet)xBk.ActiveSheet;
                appExcel.Visible = true;
                for (int i = 0; i < Ds.Tables[1].Columns.Count; i++)
                {
                    xSt.Cells[1, i + 1] = Ds.Tables[1].Columns[i].ColumnName.ToString();
                    xSt.Cells[1, i + 1].Interior.Color = System.Drawing.Color.Yellow;
                    if (Ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "SLOT" || Ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "DESTSLOT" || Ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "COMPLEVEL")
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        range2.NumberFormat = "@";
                    }
                    else if (Ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper().IndexOf("DATE") > 0)
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        //range2.NumberFormat = "@";

                        range2.NumberFormat = "0";
                    }
                }
                for (int i = 0; i < Ds.Tables[1].Rows.Count; i++)
                {
                    for (int m = 0; m < Ds.Tables[1].Columns.Count; m++)
                    {
                        xSt.Cells[i+2, m + 1] = Ds.Tables[1].Rows[i][m].ToString();
                    }

                }
                string col2 = col1 + Convert.ToString(row + 1);
                Microsoft.Office.Interop.Excel.Range range1 = xSt.Range["A1", col2];
                range1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range1.EntireColumn.AutoFit();
                range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;//水平居中  
                //range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//垂直居中
                //MessageBox.Show("导出Excle成功！");
                appExcel.DisplayAlerts = false;
                IntPtr handler = FindWindow(null, appExcel.Caption);
                SetForegroundWindow(handler);

            }
        }
        public void CopyToExcel(string Type,string FileName,string Sheetname,string Line,string WO,string Machine, DataTable dt,DataSet ds)
        {
            string col1;
            int row, col;
            if(dt==null)
            {
                dt = ds.Tables[0];
                row = ds.Tables[0].Rows.Count + ds.Tables[1].Rows.Count;
                col = ds.Tables[1].Columns.Count - 1;
                col1 = GetColumnChar(col);
            }
            else
            {
                row = dt.Rows.Count;
                col = dt.Columns.Count - 1;
                col1 = GetColumnChar(col);

            }
           
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xBk;
            if (FileName!="")
            {
                if (Sheetname == "")
                {
                    Sheetname = "1";
                }

                xBk = appExcel.Workbooks.Open(FileName);
              
            }
            else
            {
                if(Sheetname=="")
                {
                    Sheetname = "1";
                }
               
                xBk = appExcel.Workbooks.Add(true);
            }

            Microsoft.Office.Interop.Excel.Worksheet xSt;
            //xBk = appExcel.Workbooks.Add(true);
            xSt = (Microsoft.Office.Interop.Excel.Worksheet)xBk.ActiveSheet;
            appExcel.Visible = true;
            xSt = (Microsoft.Office.Interop.Excel.Worksheet)xBk.Worksheets.get_Item(Convert.ToInt16(Sheetname));
            if(Type== "CheckBOM_Rate")
            {
                if(ds.Tables[0].Rows.Count>0)
                {
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        xSt.Cells[1, i + 1] = ds.Tables[0].Columns[i].ColumnName.ToString();
                        if (ds.Tables[0].Columns[i].ColumnName.ToString().ToUpper() == "SLOT" || ds.Tables[0].Columns[i].ColumnName.ToString().ToUpper() == "DESTSLOT" || ds.Tables[0].Columns[i].ColumnName.ToString().ToUpper() == "COMPLEVEL")
                        {
                            string cellName1 = GetColumnChar(i) + "1";
                            string cellName2 = GetColumnChar(i) + row.ToString();
                            Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                            range2.NumberFormat = "@";
                        }
                        else if (ds.Tables[0].Columns[i].ColumnName.ToString().ToUpper().IndexOf("DATE") > 0)
                        {
                            string cellName1 = GetColumnChar(i) + "1";
                            string cellName2 = GetColumnChar(i) + row.ToString();
                            Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                            //range2.NumberFormat = "@";

                            range2.NumberFormat = "0";
                        }
                    }
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        for (int m = 0; m < ds.Tables[0].Columns.Count; m++)
                        {
                            xSt.Cells[i + 2, m + 1] = ds.Tables[0].Rows[i][m].ToString();
                        }

                    }
                }
                if (ds.Tables[1].Rows.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[1].Columns.Count; i++)
                    {
                        xSt.Cells[4, i + 1] = ds.Tables[1].Columns[i].ColumnName.ToString();
                        if (ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "SLOT" || ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "DESTSLOT" || ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "COMPLEVEL")
                        {
                            string cellName1 = GetColumnChar(i) + "1";
                            string cellName2 = GetColumnChar(i) + row.ToString();
                            Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                            range2.NumberFormat = "@";
                        }
                        else if (ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper().IndexOf("DATE") > 0)
                        {
                            string cellName1 = GetColumnChar(i) + "1";
                            string cellName2 = GetColumnChar(i) + row.ToString();
                            Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                            //range2.NumberFormat = "@";

                            range2.NumberFormat = "0";
                        }
                    }
                    for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                    {
                        for (int m = 0; m < ds.Tables[1].Columns.Count; m++)
                        {
                            xSt.Cells[i + 5, m +1] = ds.Tables[1].Rows[i][m].ToString();
                        }

                    }
                }
            }
            if(Type== "MaterialDifferentList")
            {
                xSt.Cells[2, 2] = Line;
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    xSt.Cells[3, i + 1] = dt.Columns[i].ColumnName.ToString();
                    if (dt.Columns[i].ColumnName.ToString().ToUpper() == "SLOT" || dt.Columns[i].ColumnName.ToString().ToUpper() == "DESTSLOT" || dt.Columns[i].ColumnName.ToString().ToUpper() == "COMPLEVEL")
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        range2.NumberFormat = "@";
                    }
                    else if (dt.Columns[i].ColumnName.ToString().ToUpper().IndexOf("DATE") > 0)
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        //range2.NumberFormat = "@";

                        range2.NumberFormat = "0";
                    }
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int m = 0; m < dt.Columns.Count; m++)
                    {
                        xSt.Cells[i + 4, m + 1] = dt.Rows[i][m].ToString();
                    }

                }
            }
            if(Type== "CopyToExcelWipLackByWo")
            {            
                xSt.Cells[2, 5] = ds.Tables[0].Rows[0][0].ToString();
                for (int i = 0; i < ds.Tables[1].Columns.Count; i++)
                {
                    xSt.Cells[3, i + 1] = ds.Tables[1].Columns[i].ColumnName.ToString();
                    if (ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "SLOT" || ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "DESTSLOT" || ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "COMPLEVEL")
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        range2.NumberFormat = "@";
                    }
                    else if (ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper().IndexOf("DATE") > 0)
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        //range2.NumberFormat = "@";

                        range2.NumberFormat = "0";
                    }
                }
                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                {
                    for (int m = 0; m < ds.Tables[1].Columns.Count; m++)
                    {
                        xSt.Cells[i + 4, m + 1] = ds.Tables[1].Rows[i][m].ToString();
                    }

                }
            }
            if(Type== "CopyToExcelWipByGroup")
            {
                xSt.Cells[2, 2] = Line;
                xSt.Cells[2, 6] = ds.Tables[0].Rows[0][0].ToString();
                for (int i = 0; i < ds.Tables[1].Columns.Count; i++)
                {
                    xSt.Cells[3, i + 1] = ds.Tables[1].Columns[i].ColumnName.ToString();
                    if (ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "SLOT" || ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "DESTSLOT" || ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "COMPLEVEL")
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        range2.NumberFormat = "@";
                    }
                    else if (ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper().IndexOf("DATE") > 0)
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        //range2.NumberFormat = "@";

                        range2.NumberFormat = "0";
                    }
                }
                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                {
                    for (int m = 0; m < ds.Tables[1].Columns.Count; m++)
                    {
                        xSt.Cells[i + 4, m + 1] = ds.Tables[1].Rows[i][m].ToString();
                    }

                }
            }
            if(Type== "CopyToExcelWipByDate")
            {
                xSt.Cells[2, 2] = Line;
                xSt.Cells[2, 5] = ds.Tables[0].Rows[0][0].ToString();
                for (int i = 0; i < ds.Tables[1].Columns.Count; i++)
                {
                    xSt.Cells[3, i + 1] = ds.Tables[1].Columns[i].ColumnName.ToString();
                    if (ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "SLOT" || ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "DESTSLOT" || ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "COMPLEVEL")
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        range2.NumberFormat = "@";
                    }
                    else if (ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper().IndexOf("DATE") > 0)
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        //range2.NumberFormat = "@";

                        range2.NumberFormat = "0";
                    }
                }
                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                {
                    for (int m = 0; m < ds.Tables[1].Columns.Count; m++)
                    {
                        xSt.Cells[i + 4, m + 1] = ds.Tables[1].Rows[i][m].ToString();
                    }

                }
            }
            if(Type== "CopyToExcelWipByMaterial")
            {
                xSt.Cells[2, 4] = Line;
                xSt.Cells[2, 6] = ds.Tables[0].Rows[0][0].ToString();
                for (int i = 0; i < ds.Tables[1].Columns.Count; i++)
                {
                    xSt.Cells[3, i + 1] = ds.Tables[1].Columns[i].ColumnName.ToString();
                    if (ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "SLOT" || ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "DESTSLOT" || ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper() == "COMPLEVEL")
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        range2.NumberFormat = "@";
                    }
                    else if (ds.Tables[1].Columns[i].ColumnName.ToString().ToUpper().IndexOf("DATE") > 0)
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        //range2.NumberFormat = "@";

                        range2.NumberFormat = "0";
                    }
                }
                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                {
                    for (int m = 0; m < ds.Tables[1].Columns.Count; m++)
                    {
                        xSt.Cells[i + 4, m + 1] = ds.Tables[1].Rows[i][m].ToString();
                    }

                }

            }
            if (Type== "PrepariMaterialList")
            {
                xSt.Cells[2, 3] = Line;
                xSt.Cells[2, 5] = WO;
                xSt.Cells[2, 7] = Machine;
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    xSt.Cells[3, i + 1] = dt.Columns[i].ColumnName.ToString();
                    if (dt.Columns[i].ColumnName.ToString().ToUpper() == "SLOT" || dt.Columns[i].ColumnName.ToString().ToUpper() == "DESTSLOT" || dt.Columns[i].ColumnName.ToString().ToUpper() == "COMPLEVEL")
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        range2.NumberFormat = "@";
                    }
                    else if (dt.Columns[i].ColumnName.ToString().ToUpper().IndexOf("DATE") > 0)
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        //range2.NumberFormat = "@";

                        range2.NumberFormat = "0";
                    }
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int m = 0; m < dt.Columns.Count; m++)
                    {
                        xSt.Cells[i + 4, m + 1] = dt.Rows[i][m].ToString();
                    }

                }
            }
            if(Type== "QSMS_DID_ToWH")
            {
                xSt.Cells[1, 1] = "Today:" + DateTime.Now.ToString("MM/DD");// Format(Now, "MM/DD");
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    xSt.Cells[2, i + 1] = dt.Columns[i].ColumnName.ToString();                 
                    if (dt.Columns[i].ColumnName.ToString().ToUpper() == "SLOT" || dt.Columns[i].ColumnName.ToString().ToUpper() == "DESTSLOT" || dt.Columns[i].ColumnName.ToString().ToUpper() == "COMPLEVEL")
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        range2.NumberFormat = "@";
                    }
                    else if (dt.Columns[i].ColumnName.ToString().ToUpper().IndexOf("DATE") > 0)
                    {
                        string cellName1 = GetColumnChar(i) + "1";
                        string cellName2 = GetColumnChar(i) + row.ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
                        //range2.NumberFormat = "@";

                        range2.NumberFormat = "0";
                    }
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int m = 0; m < dt.Columns.Count; m++)
                    {
                        xSt.Cells[i+3, m + 1] = dt.Rows[i][m].ToString();
                    }

                }
            }
            //if(Type=="")
            //{
            //    for (int i = 0; i < dt.Columns.Count; i++)
            //    {
            //        xSt.Cells[1, i + 1] = dt.Columns[i].ColumnName.ToString();
            //        xSt.Cells[1, i + 1].Interior.Color = System.Drawing.Color.Yellow;
            //        if (dt.Columns[i].ColumnName.ToString().ToUpper() == "SLOT" || dt.Columns[i].ColumnName.ToString().ToUpper() == "DESTSLOT" || dt.Columns[i].ColumnName.ToString().ToUpper() == "COMPLEVEL")
            //        {
            //            string cellName1 = GetColumnChar(i) + "1";
            //            string cellName2 = GetColumnChar(i) + row.ToString();
            //            Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
            //            range2.NumberFormat = "@";
            //        }
            //        else if (dt.Columns[i].ColumnName.ToString().ToUpper().IndexOf("DATE") > 0)
            //        {
            //            string cellName1 = GetColumnChar(i) + "1";
            //            string cellName2 = GetColumnChar(i) + row.ToString();
            //            Microsoft.Office.Interop.Excel.Range range2 = xSt.Range[cellName1, cellName2];
            //            //range2.NumberFormat = "@";

            //            range2.NumberFormat = "0";
            //        }
            //    }
            //    for (int i = 0; i < dt.Rows.Count; i++)
            //    {
            //        for (int m = 0; m < dt.Columns.Count; m++)
            //        {
            //            xSt.Cells[i+2, m + 1] = dt.Rows[i][m].ToString();
            //        }

            //    }
            //}
            
            string col2 = col1 + Convert.ToString(row + 1);
            Microsoft.Office.Interop.Excel.Range range1 = xSt.Range["A1", col2];
            range1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.EntireColumn.AutoFit();
            range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;//水平居中  
            //range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//垂直居中
            appExcel.DisplayAlerts = false;
            ////MessageBox.Show("导出Excle成功！");
            IntPtr handler = FindWindow(null, appExcel.Caption);
            SetForegroundWindow(handler);


        }
        private static string GetColumnChar(int col)
        {
            var a = col / 26;
            var b = col % 26;

            if (a > 0) return GetColumnChar(a - 1) + (char)(b + 65);

            return ((char)(b + 65)).ToString();
        }
        public static void OfficeExcel(DataTable  DT, int rowsStr, int colsStr)
        {
            dynamic _app = new Microsoft.Office.Interop.Excel.Application();
            dynamic _workbook;
            _workbook = _app.Workbooks.Add(true);
            Microsoft.Office.Interop.Excel._Worksheet objSheet;
            objSheet = _workbook.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range range;
            _app.Visible = true;
            try
            {
                range = objSheet.get_Range("A1", Missing.Value);
                range = range.get_Resize(rowsStr+1, colsStr);
                object[,] saRet = new object[rowsStr+1, colsStr];
                for(int iCol = 0; iCol<colsStr;iCol++)
                {
                    saRet[0, iCol] = DT.Columns[iCol].ColumnName.ToString();
                    if (DT.Columns[iCol].ColumnName.ToString().ToUpper() == "SLOT" || DT.Columns[iCol].ColumnName.ToString().ToUpper() == "DESTSLOT" || DT.Columns[iCol].ColumnName.ToString().ToUpper() == "COMPLEVEL")
                    {
                        string cellName1 = GetColumnChar(iCol) + "1";
                        string cellName2 = GetColumnChar(iCol) + (rowsStr + 1).ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = objSheet.Range[cellName1, cellName2];
                        range2.NumberFormat = "@";
                    }
                    else if (DT.Columns[iCol].ColumnName.ToString().ToUpper().IndexOf("DATE") > 0)
                    {
                        string cellName1 = GetColumnChar(iCol) + "1";
                        string cellName2 = GetColumnChar(iCol) + (rowsStr + 1).ToString();
                        Microsoft.Office.Interop.Excel.Range range2 = objSheet.Range[cellName1, cellName2];
                        //range2.NumberFormat = "@";

                        range2.NumberFormat = "0";
                    }

                   
                }
                for (int iRow = 0; iRow < rowsStr; iRow++)
                {
                    int row = iRow;
                    for (int iCol = 0; iCol < colsStr; iCol++)
                    {
                        int col = iCol;
                        string DD= DT.Rows[row][col].ToString();
                        saRet[iRow+1, iCol] = DT.Rows[row][col].ToString();
                    }
                }
                range.set_Value(Missing.Value, saRet);
              
                _app.UserControl = true;
                string col2 = colsStr + Convert.ToString(rowsStr + 1);
                //Microsoft.Office.Interop.Excel.Range range1 = _app.Range["A1", col2];
                range.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                range.EntireColumn.AutoFit();
                range.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;//水平居中  
                //range.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;//垂直居中
                Microsoft.Office.Interop.Excel.Range range1 = objSheet.get_Range("A1", Missing.Value);
                range1 = range.get_Resize(1, colsStr);
                range1.Interior.Color = System.Drawing.Color.Yellow;
                _app.DisplayAlerts = false;
                //MessageBox.Show("导出Excle成功！");
                IntPtr handler = FindWindow(null, _app.Caption);
                SetForegroundWindow(handler);

            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                MessageBox.Show(errorMessage, "Error");
            }
        }

        public void ExportDataToExcel(DataTable TableName, string FileName)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            //设置文件标题
            saveFileDialog.Title = "导出Excel文件";
            //设置文件类型
            saveFileDialog.Filter = "Microsoft Office Excel 工作簿(*.xls)|*.xls";
            //设置默认文件类型显示顺序  
            saveFileDialog.FilterIndex = 1;
            //是否自动在文件名中添加扩展名
            saveFileDialog.AddExtension = true;
            //是否记忆上次打开的目录
            saveFileDialog.RestoreDirectory = true;
            //设置默认文件名
            saveFileDialog.FileName = FileName;
            //按下确定选择的按钮  
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //获得文件路径 
                string localFilePath = saveFileDialog.FileName.ToString();

                //数据初始化
                int TotalCount;     //总行数
                int RowRead = 0;    //已读行数
                int Percent = 0;    //百分比

                TotalCount = TableName.Rows.Count;
                //lblStatus.Text = "共有" + TotalCount + "条数据";
                //lblStatus.Visible = true;
                //barStatus.Visible = true;

                //数据流
                Stream myStream = saveFileDialog.OpenFile();
                StreamWriter sw = new StreamWriter(myStream, Encoding.GetEncoding("gb2312"));
                string strHeader = "";

                //秒钟
                Stopwatch timer = new Stopwatch();
                timer.Start();

                try
                {
                    //写入标题
                    for (int i = 0; i < TableName.Columns.Count; i++)
                    {
                        if (i > 0)
                        {
                            strHeader += "\t";
                        }
                        strHeader += TableName.Columns[i].ColumnName.ToString();
                    }
                    sw.WriteLine(strHeader);

                    //写入数据
                    //string strData;
                    for (int i = 0; i < TableName.Rows.Count; i++)
                    {
                        RowRead++;
                        Percent = (int)(100 * RowRead / TotalCount);
                        //barStatus.Maximum = TotalCount;
                        //barStatus.Value = RowRead;
                        //lblStatus.Text = "共有" + TotalCount + "条数据，已写入" + Percent.ToString() + "%的数据，共耗时" + timer.ElapsedMilliseconds + "毫秒。";
                        Application.DoEvents();

                        string strData = "";
                        for (int j = 0; j < TableName.Columns.Count; j++)
                        {
                            if (j > 0)
                            {
                                strData += "\t";
                            }
                            strData += TableName.Rows[i][j].ToString();
                        }
                        sw.WriteLine(strData);
                    }
                    //关闭数据流
                    sw.Close();
                    myStream.Close();
                    //关闭秒钟
                    timer.Reset();
                    timer.Stop();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                finally
                {
                    //关闭数据流
                    sw.Close();
                    myStream.Close();
                    //关闭秒钟
                    timer.Stop();
                }

                //成功提示
                if (MessageBox.Show("导出成功，是否立即打开？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(localFilePath);
                }

                //赋初始值
                //lblStatus.Visible = false;
                //barStatus.Visible = false;
            }
        }

        private void CmdSFile_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.FileName = "";
            this.openFileDialog1.Filter = "电子表格(*.xls;*.xlsx;*.ods)|*.xls;*.xlsx;*.ods";
            this.openFileDialog1.Title = "请选择要上传的Excel文件!";
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = this.openFileDialog1.FileName;
                //cboSheetName.DataSource = GetSheetList(txtFilePath.Text);
                if (txtFilePath.Text.EndsWith(".ods"))
                {
                    exportFormat = "LibreOffice";
                }
            }
        }

        private void cmdUpload_Click(object sender, EventArgs e)
        {
            int rCount = 2;
            if (txtFilePath.Text=="")
            {
                MessageBox.Show("请先选择上传文档，谢谢！");
                return;
            }
            if(CboReportType.Text== "QSMS_CheckCompPN" || CboReportType.Text== "CheckReplacePNBySAPBOM")
            {
                cmdExcel_Click(null, null);
            }
            else
            {
                Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xBk;
                xBk = appExcel.Workbooks.Open(txtFilePath.Text);
                Microsoft.Office.Interop.Excel.Worksheet xSt;
                //xBk = appExcel.Workbooks.Add(true);
                xSt = (Microsoft.Office.Interop.Excel.Worksheet)xBk.ActiveSheet;
                appExcel.Visible = true;
                xSt = (Microsoft.Office.Interop.Excel.Worksheet)xBk.Worksheets.get_Item(1);
                ListWoSelecting.Items.Clear();
                for(rCount = 2; rCount <xSt.Cells.Count;rCount++)
                {
                    if(xSt.Cells[rCount,1].ToString()!="")
                    {
                        ListWoSelecting.Items.Add(xSt.Cells[rCount, 1].ToString());
                    }
                }
                xBk.Close();
            }
        }

        private void frmReport_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmReport");
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void TxtWOQty_TextChanged(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void TxtRev_TextChanged(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void TxtJobGroup_TextChanged(object sender, EventArgs e)
        {

        }

        private void CboJobPN_SelectedValueChanged(object sender, EventArgs e)
        {
           
        }

        private void CboMachine_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetComp(CboMachine.Text.Trim(), TxtMBPN.Text.Trim(), TxtWO.Text.Trim(), CboLine.Text.Trim());
            GetJobGroupByJobRev(CboMachine.Text.Trim(), TxtMBPN.Text.Trim(), TxtRev.Text.Trim());
        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void CboComp_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void TxtMBPN_TextChanged(object sender, EventArgs e)
        {

        }

      

   
    }
}
