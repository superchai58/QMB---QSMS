using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.MCC
{
    public partial class FrmDeleteME_BOM : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.MCC.DeleteME_BOM DeleteME_BOM = new DbLibrary.MCC.DeleteME_BOM();
        public FrmDeleteME_BOM()
        {
            InitializeComponent();
        }

        private void FrmDeleteME_BOM_Load(object sender, EventArgs e)
        {
            dtpSDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
            dtpEDate.Text = DateTime.Now.AddDays(1).ToString("yyyy/MM/dd");
            CmdDeleteByLine.Visible = false;
            if (Parameter.DeleteMeBomByLine == true)
            {
                CmdDeleteByLine.Visible = true;
            }
            if(Parameter.StrBU=="NB5")
            {
                lblslot.Visible = true;
                cboslot.Visible = true;
                lblSide.Visible = true;
                txtside.Visible = true;
            }
            GetLine();
        }
        public void GetLine()
        {
            DataTable dt = new DataTable();
            dt = DeleteME_BOM.GetLine();
            if(dt.Rows.Count>0)
            {
                for(int i=0;i<dt.Rows.Count;i++)
                {
                    CboLine.Items.Add(dt.Rows[i]["Line"].ToString());
                }
            }
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
            if (JobPN != "" && JobPN.IndexOf("-") > 0)
            {
                TxtMBPN.Text = JobPN.Substring(1, 11);
                TxtRev.Text = JobPN.Substring(JobPN.Length - 3);
            }
            dt = DeleteME_BOM.GetGroupID(JobPN, Parameter.BU, BeginDate, EndDate, CboLine.Text, OptRelease.Checked);

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
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
            dt = DeleteME_BOM.GetjobPn(Parameter.BU, CboLine.Text, BeginDate, EndDate, OptRelease.Checked);
            CboJobPN.Items.Clear();
            CboJobPN.Text = "";
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    CboJobPN.Items.Add(dt.Rows[i]["JobPN"].ToString().Trim());
                }
            }
            else
            {
                MessageBox.Show("No data");
            }
        }

        public void GetGroupWO(string groupid)
        {
            string TempJobPn = "";
            if (CboJobPN.Text != "")
            {
                TempJobPn = CboJobPN.Text.Substring(0, 11);
            }
            DataTable dt = new DataTable();
            dt = DeleteME_BOM.GetGroupWO(groupid, TempJobPn);
            ListWoall.Items.Clear();
            CboWo.Items.Clear();
            CboWo.Text = "";
            CboNotChkBOM.Items.Clear();
            CboNotChkBOM.Text = "";
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (ChkQSMS_WO(dt.Rows[i]["Work_Order"].ToString()) == false)
                    {
                        CboNotChkBOM.Items.Add(dt.Rows[i]["Work_Order"].ToString());
                    }
                    else
                    {
                        ListWoall.Items.Add(dt.Rows[i]["Work_Order"].ToString());
                        CboWo.Items.Add(dt.Rows[i]["Work_Order"].ToString());

                    }
                }
            }
        }
        public bool ChkQSMS_WO(string WO)
        {
            DataTable dt = new DataTable();
            dt = DeleteME_BOM.GetWO(WO);
            if (dt.Rows.Count == 0)
            {
                return false;
            }
            return true;
        }

        private void CboGroupID_Click(object sender, EventArgs e)
        {
            
        }

        private void CboGroupID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar==13 && e.ToString()!="")
            {
                //CboGroupID_SelectedIndexChanged(null, null);
                GetGroupWO(CboGroupID.Text.Trim());
            }
        }

        private void CboJobPN_Click(object sender, EventArgs e)
        {
           
        }
        public void GetJobGroupByJobRev(string Machine, string MBPN, string Rev)
        {
            DataTable dt = new DataTable();
            dt = DeleteME_BOM.GetJobGroupByJobRev(Machine, MBPN, Rev);
            ListAllJobGroup.Items.Clear();
            ListselectingJobGroup.Items.Clear();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ListAllJobGroup.Items.Add(dt.Rows[i]["jobgroup"].ToString());
                }
            }
        }

        private void CboJobPN_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && e.ToString() != "")
            {
                GetGroupID(CboJobPN.Text);
                GetJobGroupByJobRev("", TxtMBPN.Text, TxtRev.Text);

            }
        }

        private void CboLine_Click(object sender, EventArgs e)
        {
            string BeginDate, EndDate;
            BeginDate = dtpSDate.Value.ToString("yyyy/MM/dd");
            BeginDate = BeginDate.Replace("-", "");
            BeginDate = BeginDate.Replace("/", "");
            EndDate = dtpEDate.Value.ToString("yyyy/MM/dd"); ;
            EndDate = EndDate.Replace("-", "");
            EndDate = EndDate.Replace("/", "");
            DataTable dt = new DataTable();
            dt = DeleteME_BOM.GetJobpn(BeginDate, EndDate, CboLine.Text, OptRelease.Checked);
            CboJobPN.Text = "";
            CboJobPN.Items.Clear();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    CboJobPN.Items.Add(dt.Rows[i]["Jobpn"].ToString().Trim() + "-" + dt.Rows[i]["MB_Rev"].ToString().Trim());
                }

            }
            if(Parameter.StrBU=="NB5")
            {
                dt = DeleteME_BOM.GetMachine(CboLine.Text);
                CboMachine.Items.Clear();
                if(dt.Rows.Count>0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        CboMachine.Items.Add(dt.Rows[i]["Machine"].ToString());
                    }
                    
                }
            }
        }

        private void CboLine_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (e.KeyChar == 13 && e.ToString() != "")
            {
                CboLine_Click(null, null);
            }
        }

        private void CboMachine_Click(object sender, EventArgs e)
        {
            //GetComp(CboMachine.Text.Trim(), TxtMBPN.Text.Trim(), TxtWO.Text.Trim(), CboLine.Text.Trim());
            //GetJobGroupByJobRev(CboMachine.Text.Trim(), TxtMBPN.Text.Trim(), TxtRev.Text.Trim());
            //if(Parameter.StrBU=="NB5" && CboMachine.Text.Trim()!="All")
            //{
            //    GetSlot(CboMachine.Text.Trim());
            //}
        }
        public void GetSlot(string Machine)
        {
            DataTable dt = new DataTable();
            dt = DeleteME_BOM.GetSlot(Machine,CboLine.Text);
            cboslot.Items.Clear();
            if(dt.Rows.Count>0)
            {

                for(int i=0;i<dt.Rows.Count;i++)
                {
                    cboslot.Items.Add(dt.Rows[i]["Slot"].ToString());
                }
            }
        }
        public void GetComp(string Machine, string MBPN, string WO, string Line)
        {
            DataTable dt = new DataTable();
            dt = DeleteME_BOM.GetComp(Machine, MBPN, WO, Line);
            CboComp.Items.Clear();
            CboComp.Text = "";
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    CboComp.Items.Add(dt.Rows[i]["compPN"].ToString());
                }
            }
        }

        private void CboMachine_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && e.ToString() != "")
            {
                GetComp(CboMachine.Text.Trim(), TxtMBPN.Text.Trim(), TxtWO.Text.Trim(), CboLine.Text.Trim());
                GetJobGroupByJobRev(CboMachine.Text.Trim(), TxtMBPN.Text.Trim(), TxtRev.Text.Trim());
                if (Parameter.StrBU == "NB5" && CboMachine.Text.Trim() != "All")
                {
                    GetSlot(CboMachine.Text.Trim());
                }
            }
        }

        private void CboNotChkBOM_Click(object sender, EventArgs e)
        {
           
        }
        public void GetWOInfo(string WO)
        {
            DataTable dt = new DataTable();
            dt = DeleteME_BOM.GetWoinfo(WO);
            if (dt.Rows.Count > 0)
            {
                TxtMBPN.Text = dt.Rows[0]["PN"].ToString().Trim();
                TxtWOQty.Text = dt.Rows[0]["QTY"].ToString().Trim();
                TxtRev.Text = dt.Rows[0]["MB_REV"].ToString().Trim();
                CboLine.Text = dt.Rows[0]["Line"].ToString().Trim();
            }
            dt = DeleteME_BOM.GetCustomer(TxtMBPN.Text);
            if (dt.Rows.Count > 0)
            {
                TxtCustomer.Text = dt.Rows[0]["Customer"].ToString();
            }
            CboJobPN.Items.Clear();
            CboJobPN.Text = "";
            dt = DeleteME_BOM.GetjobPn(WO);
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    CboJobPN.Items.Add(dt.Rows[i]["jobpn"].ToString());
                }
            }
        }
        public void GetMachine(string WO)
        {
            string woGroup, joppn;
            DataTable dt = new DataTable();
            dt = DeleteME_BOM.GetMachineWo(WO);
            if(dt.Rows.Count==0)
            {
                MessageBox.Show("This WO is not exist! Please check!");
                return;
            }
            woGroup = dt.Rows[0]["Group"].ToString();
            joppn = GetJobGroup(woGroup);
            dt = DeleteME_BOM.GetMachine(WO, joppn);
            CboMachine.Items.Clear();
            CboMachine.Text = "";
            CboMachine.Items.Add("ALL");
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    CboMachine.Items.Add(dt.Rows[i]["Machine"].ToString().Trim());
                }
            }
        }
        public string GetJobGroup(string Group)
        {
            string jobgroup="", Jobpn = "", GetJobGroup="";
            DataTable dt = new DataTable();
            dt = DeleteME_BOM.QSMS_GetEMSFlag(Group);
            if(dt.Rows.Count>0)
            {
                if(dt.Rows[0]["EMSFlag"].ToString()=="NONE")
                {
                    GetJobGroup = "";
                    return GetJobGroup;
                }
                if (dt.Rows[0]["EMSFlag"].ToString() != "Y")
                {
                    dt = DeleteME_BOM.QSMS_GetEMSFlagMB(Group);
                    if(dt.Rows.Count==0)
                    {
                        dt = DeleteME_BOM.QSMS_GetEMSFlagSB(Group);
                        if(dt.Rows.Count==0)
                        {
                            GetJobGroup = "";
                            return GetJobGroup;
                        }
                        else
                        {
                            for(int i=0;i<dt.Rows.Count;i++)
                            {
                                jobgroup = dt.Rows[i]["jobpn"].ToString().Trim() + "-" + dt.Rows[i]["Mb_Rev"].ToString().Trim();
                                Jobpn = Jobpn + "'" + jobgroup + "'" + ",";
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            jobgroup = dt.Rows[i]["jobpn"].ToString().Trim() + "-" + dt.Rows[i]["Mb_Rev"].ToString().Trim();
                            Jobpn = Jobpn + "'" + jobgroup + "'" + ",";
                        }
                    }
                }
                else
                {
                    dt = DeleteME_BOM.QSMS_GetEMSFlagInitAOIFlag(Group);
                    if(dt.Rows.Count==0)
                    {
                        GetJobGroup = "";
                        return GetJobGroup;
                    }
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        jobgroup = dt.Rows[i]["jobpn"].ToString().Trim() + "-" + dt.Rows[i]["Mb_Rev"].ToString().Trim();
                        Jobpn = Jobpn + "'" + jobgroup + "'" + ",";
                    }
                }
            }
            Jobpn = Jobpn.Substring(0, Jobpn.Length - 1);
            Jobpn = "(" + Jobpn + ")";
            GetJobGroup = Jobpn;
            return GetJobGroup;

        }

        private void CboNotChkBOM_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && e.ToString() != "")
            {
                //CboNotChkBOM_SelectedValueChanged(null, null);

                TxtWO.Text = CboNotChkBOM.Text;
                GetWOInfo(TxtWO.Text);
                GetMachine(TxtWO.Text);
            }
        }

        private void CboWo_Click(object sender, EventArgs e)
        {
           
        }

        private void CboWo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && e.ToString() != "")
            {
                CboWo_SelectedValueChanged(null, null);
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
            if (CboLine.Text == "")
            {
                MessageBox.Show("Please input line");
                return;
            }
            GetGroupID("");
            GetJobPN();
        }

        private void CmdDeleteByLine_Click(object sender, EventArgs e)
        {
            string site = "";
             DataTable dt = new DataTable();
            dt = DeleteME_BOM.GetSite();
            site = dt.Rows[0]["Site"].ToString();
            if (MessageBox.Show("Are you sure to delete this ME_BOM by line " + CboLine.Text+ " ?\r\n", "确认", MessageBoxButtons.YesNo).ToString().ToUpper() == "NO")
            {
                return;
            }
            if(CboJobPN.Text=="")
            {
                DeleteME_BOM.delQSMS_MEBOMBYLine(CboLine.Text, site);
              
            }
            else
            {
                DeleteME_BOM.delQSMS_MEBOMBYLinejob(CboLine.Text,CboJobPN.Text, site);

            }
            //dt = DeleteME_BOM.QSMSDeleteME_BOM("DeleteByLine", CboJobPN.Text, CboLine.Text,Parameter.StrBU,Parameter.UID,"","","","","","","","");
            DeleteME_BOM.insertQSMS_Logbyline(CboLine.Text, Parameter.UID);
            MessageBox.Show("Delete ME bom OK");                             
            
        }

        private void CmdDelete_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            string site = "",Machine="",Line="";
            string jobgroup = TxtJobGroup.Text;
            dt = DeleteME_BOM.GetSite();
            site = dt.Rows[0]["Site"].ToString();
            if (TxtWO.Text=="" && (site == "NB5" || site == "NB3"))
            {
                if (TxtJobGroup.Text != "")
                {
                    DeleteME_BOM.QSMSDeleteME_BOMdel(CboJobPN.Text, CboLine.Text, site, Parameter.UID, jobgroup, "", TxtRev.Text, CboMachine.Text, txtside.Text, CboComp.Text, cboslot.Text, TxtMBPN.Text); ;
                    //dt = DeleteME_BOM.QSMSDeleteME_BOM("DeleteMEBOM", CboJobPN.Text, CboLine.Text, site, Parameter.UID, jobgroup, "", "", "", "", "", "", "");
                    MessageBox.Show("Delete ME bom OK");
                }
            }
            else
            {
                if(TxtMBPN.Text=="")
                {
                    MessageBox.Show("Please input MBPN");
                    return;
                }
                if(CboMachine.Text=="")
                {
                    MessageBox.Show("Please input Machine (set Machine=All to delete all)");
                    return;
                }
                jobgroup = GetSelectingJobGroup();
                if(jobgroup=="")
                {
                    MessageBox.Show("请选择 JobGroup");
                    return;
                }
                if(CboMachine.Text.ToUpper()=="ALL")
                {
                    Machine = "%";
                }
                else
                {
                    Machine = CboMachine.Text.ToUpper()+"%";
                }
                dt = DeleteME_BOM.QSMS_MEBOM(TxtMBPN.Text,jobgroup,TxtRev.Text,CboMachine.Text,CboLine.Text);
               
                if (dt.Rows.Count==0)
                {
                    MessageBox.Show("can not find ME BOM ,Please check the MBPN or Rev");
                    return;
                }
                DeleteME_BOM.QSMSDeleteME_BOMall( CboJobPN.Text, CboLine.Text, site, Parameter.UID, jobgroup, "", TxtRev.Text, CboMachine.Text, "", "", "", TxtMBPN.Text);
                dt = DeleteME_BOM.WO_MultiLine(TxtWO.Text.Trim());
                if(dt.Rows.Count>0)
                {
                    Line = dt.Rows[0]["Line"].ToString();
                    DeleteME_BOM.QSMSDeleteME_BOMall(CboJobPN.Text, Line, site, Parameter.UID, jobgroup, "", TxtRev.Text, CboMachine.Text, "", "", "", TxtMBPN.Text);
                 }
                DeleteME_BOM.insertQSMS_Log(TxtMBPN.Text,TxtRev.Text,CboMachine.Text,Parameter.UID);
                MessageBox.Show("Delete ME bom OK");
            }
           
        }
        public string GetSelectingJobGroup()
        {
            string SelectingJobGroup = "", GetSelectingJobGroup = "";
            if (ListselectingJobGroup.Items.Count <= 0)
            {
                if (TxtJobGroup.Text.Length > 0 && TxtJobGroup.Text.IndexOf("-") > 0)
                {
                    GetSelectingJobGroup = "(" + TxtJobGroup.Text + ")";
                }
                else
                {
                    SelectingJobGroup = "";
                }
            }
            for (int i =0; i < ListselectingJobGroup.Items.Count; i++)
            {
                ListselectingJobGroup.SelectedIndex = i;
                SelectingJobGroup = SelectingJobGroup + "'" + ListselectingJobGroup.SelectedItem.ToString() + "'" + ",";
            }
            GetSelectingJobGroup = "("+ SelectingJobGroup.Substring(0, SelectingJobGroup.Length - 1)+")" ;
            return GetSelectingJobGroup;
        }

        private void TxtMBPN_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar==13 && e.ToString()!="")
            {
                TxtRev.Enabled = true;
                TxtRev.Focus();
            }
        }

        private void TxtRev_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && e.ToString() != "")
            {
                GetJobGroupByJobRev(CboMachine.Text.Trim(), TxtMBPN.Text.Trim(), TxtRev.Text.Trim());
            }
        }

        private void FrmDeleteME_BOM_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("FrmDeleteME_BOM");
        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void ListAllJobGroup_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void ListselectingJobGroup_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void CboLine_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void TxtCustomer_TextChanged(object sender, EventArgs e)
        {

        }

        private void CboJobPN_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetGroupID(CboJobPN.Text);
            GetJobGroupByJobRev("", TxtMBPN.Text, TxtRev.Text);
        }

        private void CboNotChkBOM_SelectedIndexChanged(object sender, EventArgs e)
        {
          
        }

        private void CboWo_SelectedIndexChanged(object sender, EventArgs e)
        {
            //TxtWO.Text = CboNotChkBOM.Text;
            //GetWOInfo(TxtWO.Text);
            //GetMachine(TxtWO.Text);
        }

        private void CboGroupID_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetGroupWO(CboGroupID.Text.Trim());
        }

        private void CboMachine_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetComp(CboMachine.Text.Trim(), TxtMBPN.Text.Trim(), TxtWO.Text.Trim(), CboLine.Text.Trim());
            GetJobGroupByJobRev(CboMachine.Text.Trim(), TxtMBPN.Text.Trim(), TxtRev.Text.Trim());
            if (Parameter.StrBU == "NB5" && CboMachine.Text.Trim() != "All")
            {
                GetSlot(CboMachine.Text.Trim());
            }
        }

        private void CboWo_SelectedValueChanged(object sender, EventArgs e)
        {
            TxtWO.Text = CboWo.Text;
            GetWOInfo(TxtWO.Text);
            GetMachine(TxtWO.Text);
        }

        private void CboNotChkBOM_SelectedValueChanged(object sender, EventArgs e)
        {
            TxtWO.Text = CboNotChkBOM.Text;
            GetWOInfo(TxtWO.Text);
            GetMachine(TxtWO.Text);
        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void TxtWO_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
