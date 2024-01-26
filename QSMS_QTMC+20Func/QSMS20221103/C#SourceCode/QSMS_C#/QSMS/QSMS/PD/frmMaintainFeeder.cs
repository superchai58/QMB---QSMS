using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.PD
{
    public partial class frmMaintainFeeder : Form
    {
        BrLibrary.PublicFunction PublicFunction = new BrLibrary.PublicFunction();
        DbLibrary.PD.PDProcess PD = new DbLibrary.PD.PDProcess();
        DataTable dt = new DataTable();
        DataTable dtt = new DataTable();
        DataSet ds = new DataSet();
        string Type = "", CompPN = "", VendorCode = "", DateCode = "", LotCode = "", Slot = "", LR = "", AVLCustomer = "", ModelFlag = "N";
        string[] NewComp;
        public frmMaintainFeeder()
        {
            InitializeComponent();
        }

        private void frmMaintainFeeder_Load(object sender, EventArgs e)
        {
            CboMachine.Text = "";
            CboLine.Text = "";
            CboMachine.Items.Clear();
            CboLine.Items.Clear();
            Type = "GeiMachineData";
            dt = PD.GeiMachineData(Type);
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                { CboMachine.Items.Add(dt.Rows[i]["Machine"].ToString().ToUpper()); }

            }
            else
            {
                MessageBox.Show("找不到Machine数据！！");
                return;
            }
            Type = "GeiLine";
            dt = PD.GeiLine(Type);
            if (dt.Rows.Count != 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                { CboLine.Items.Add(dt.Rows[i]["Line"].ToString().ToUpper()); }
            }
            else
            {
                MessageBox.Show("找不到Line数据！！");
                return;
            }
            OptPanal.TabIndex = 0;
        }

        private void CboGroupID_SelectedIndexChanged(object sender, EventArgs e)
        {
            dt = null;
            dtt = null;
            cboWO.Text = "";
            cboWO.Items.Clear();
            ListNoChkBOM.Text = "";
            ListNoChkBOM.Items.Clear();
            ListNotDispatch.Text = "";
            ListNotDispatch.Items.Clear();
            ListClosed.Text = "";
            ListClosed.Items.Clear();
            Type = "GetWoByGroupID";
            dt = PD.GetWoByGroupID(CboGroupID.Text.Trim().ToString(), Type);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i]["InitAOIFlag"].ToString().ToUpper().Trim() == "Y")
                {
                    cboWO.Items.Add(dt.Rows[i]["WO"].ToString());
                    if (int.Parse(dt.Rows[i]["Status"].ToString()) < 10)
                    {
                        ListNoChkBOM.Items.Add(dt.Rows[i]["Status"].ToString());
                    }
                    if (dt.Rows[i]["Sap1Flag"].ToString().ToUpper().Trim() == "N")
                    {
                        ListNotDispatch.Items.Add(dt.Rows[i]["WO"].ToString());
                    }
                    if (dt.Rows[i]["ClosedFlag"].ToString().ToUpper().Trim() == "Y")
                    {
                        ListClosed.Items.Add(dt.Rows[i]["WO"].ToString());
                    }
                }
            }
            if (ListNoChkBOM.Items.Count != 0)
            { ListNoChkBOM.SelectedIndex = 0; }
            else
            { ListNoChkBOM.Items.Clear(); }
            if (ListNotDispatch.Items.Count != 0)
            { ListNotDispatch.SelectedIndex = 0; }
            else
            { ListNotDispatch.Items.Clear(); }
            if (ListClosed.Items.Count != 0)
            { ListClosed.SelectedIndex = 0; }
            else
            { ListClosed.Items.Clear(); }
        }

        private void CmdQuery_Click(object sender, EventArgs e)
        {
            dt = null;
            CboGroupID.Text = "";
            CboGroupID.Items.Clear();
            string BeginDate = Convert.ToDateTime(dtSDate.Text).ToString("yyyyMMdd");
            string EndDate = Convert.ToDateTime(dtpEDate.Text).ToString("yyyyMMdd");
            if (CboLine.Text.Trim().ToString() == "")
            {
                MessageBox.Show("线别不能为空！！");
                return;
            }
            Type = "GetGroupIDByLine";
            dt = PD.GetGroupIDByLine(BeginDate, EndDate, CboLine.Text.Trim().ToString(), Type);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                CboGroupID.Items.Add(dt.Rows[i]["GroupID"].ToString());
            }
        }

        private void cboWO_SelectedIndexChanged(object sender, EventArgs e)
        {
            dt = null;
            CboSBWO.Text = "";
            TxtMBPN.Text = "";
            TxtCustomer.Text = "";
            TxtModel.Text = "";
            TxtWOQty.Text = "";
            TxtGroup.Text = "";
            CboSBWO.Items.Clear();
            Type = "GetGroup";
            dt = PD.GetGroup(cboWO.Text.ToString().Trim(), Type);
            TxtGroup.Text = dt.Rows[0]["Group"].ToString().Trim();
            TxtMBPN.Text = dt.Rows[0]["PN"].ToString().Trim();
            TxtCustomer.Text = dt.Rows[0]["Customer"].ToString().Trim();
            TxtModel.Text = dt.Rows[0]["Model"].ToString().Trim();
            TxtWOQty.Text = dt.Rows[0]["Qty"].ToString().Trim();
            Type = "GetSBWO";
            dt = PD.GetSBWO(cboWO.Text.ToString().Trim(), TxtGroup.Text.ToString().Trim(), Type);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                CboSBWO.Items.Add(dt.Rows[i]["WO"].ToString().Trim());
            }
            Type = "GetMachine";
            dt = PD.GetMachine(TxtGroup.Text.ToString().Trim(), Type);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                CboMachine.Items.Add(dt.Rows[i]["Machine"].ToString().Trim());
            }
        }

        private void cmdCheck_Click(object sender, EventArgs e)
        {
            if (cboWO.Text.ToString().Trim() == null)
            {
                MessageBox.Show("WO不能为空，请检查！！");
                return;
            }
        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            dt = null;
            Type = "CopyToExcel";
            if (CboMachine.Text.ToString().Trim() != "" && cboJobGroup.Text.ToString().Trim() != "")
            {
                dt = PD.CopyToExcel(cboJobGroup.Text.ToString().Trim(), CboMachine.Text.ToString().Trim(), Type);
            }
            else
            {
                MessageBox.Show("JobGroup 和 Machine不能为空，请检查！！");
                return;
            }
            PublicFunction.doExport(dt);

        }

        private void CboMachine_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboJobGroup.Text = "";
            cboJobGroup.Items.Clear();
            Type = "GetJobByMachine";
            if (CboMachine.Text.ToString().Trim() != null && TxtGroup.Text.ToString().Trim() != null)
            {
                dt = PD.GetJobByMachine(CboMachine.Text.ToString().Trim(), TxtGroup.Text.ToString().Trim(), Type);
            }
            else
            {
                MessageBox.Show("Machine/Group不能为空，请检查！！");
                return;
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cboJobGroup.Items.Add(dt.Rows[i]["JobGroup"].ToString().Trim());
            }
            RefreshMachineFeeder();

        }
        private void RefreshMachineFeeder()
        {
            DGMachine.DataSource = null;
            DGNeed.DataSource = null;
            DGDID.DataSource = null;
            DgDIDSlot.DataSource = null;
            Type = "MachineFeeder";
            ds = PD.MachineFeeder(cboJobGroup.Text.ToString().Trim(), CboMachine.Text.ToString().Trim(), cboWO.Text.ToString().Trim(), Type);
            GetRefreshMachineFeederData(ds.Tables[0], ds.Tables[1], ds.Tables[2], ds.Tables[3]);
        }
        private void GetRefreshMachineFeederData(DataTable dt, DataTable dt1, DataTable dt2, DataTable dt3)
        {
            if (dt.Rows[0]["Machine"].ToString().Trim() != "0")
            {
                DGMachine.DataSource = dt;
            }
            else
            {
                DGMachine.DataSource = null;
            }
            if (dt1.Rows[0]["Machine"].ToString().Trim() != "0")
            {
                DGNeed.DataSource = dt1;
            }
            else
            {
                DGNeed.DataSource = null;
            }
            if (dt2.Rows[0]["Machine"].ToString().Trim() != "0")
            {
                DGDID.DataSource = dt2;
            }
            else
            {
                DGDID.DataSource = null;
            }
            if (dt3.Rows[0]["Machine"].ToString().Trim() != "0")
            {
                DgDIDSlot.DataSource = dt3;
            }
            else
            {
                DgDIDSlot.DataSource = null;
            }
            TxtDID.Focus();
        }

        private void cboJobGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboJobGroup.Text.ToString().Trim() != null)
            {
                RefreshMachineFeeder();
            }
        }

        private void cmdReset_Click(object sender, EventArgs e)
        {
            TxtDID.Enabled = true;
            TxtCompPN.Enabled = true;
            TxtFeeder.Enabled = true;
            TxtLR.Enabled = true;
            TxtDID.Text = "";
            TxtCompPN.Text = "";
            TxtFeeder.Text = "";
            TxtLR.Text = "";
            CompPN = "";
            VendorCode = "";
            DateCode = "";
            LotCode = "";
            Slot = "";
            LR = "";
            TxtDID.Focus();
        }

        private void OptPanal_Click(object sender, EventArgs e)
        {
            TxtLR.Enabled = true;
            TxtLR.BackColor = Color.White;
        }
        private void OptFuJi_Click(object sender, EventArgs e)
        {
            TxtLR.Enabled = false;
            TxtLR.BackColor = Color.Gray;
        }

        private void CmdSave_Click(object sender, EventArgs e)
        {
            try
            {
                string Line = "";
                dt = null;
                dt = PD.GenUNID(TxtDID.Text.ToString().Trim(), "");
                if (dt.Rows[0]["Result"].ToString().Trim() != "OK")
                {
                    MessageBox.Show(dt.Rows[0]["Msg"].ToString());
                    TxtDID.Text = "";
                    TxtDID.Focus();
                    return;
                }
                else
                {
                    TxtDID.Text = dt.Rows[0]["UNID"].ToString().Trim();
                }
                Type = "QueryDID";
                dt = PD.QueryDID(TxtDID.Text.ToString().Trim(), Type);
                if (dt.Rows[0]["Result"].ToString().Trim() != "OK")
                {
                    MessageBox.Show(dt.Rows[0]["Msg"].ToString());
                    TxtDID.Text = "";
                    TxtDID.Focus();
                    return;
                }
                else
                {
                    if (dt.Rows[0]["Msg"].ToString().Trim() != null)
                    {
                        MessageBox.Show(dt.Rows[0]["Msg"].ToString());
                        TxtDID.Text = dt.Rows[0]["DID"].ToString().Trim();
                        TxtCompPN.Text = dt.Rows[0]["CompPN"].ToString().Trim();
                    }
                    TxtDID.Text = dt.Rows[0]["DID"].ToString().Trim();
                    TxtCompPN.Text = dt.Rows[0]["CompPN"].ToString().Trim();
                }
                if (ChkValid() == false)
                {
                    return;
                }
                if (PublicFunction.ConfigListGetValue("CheckFeeder") == "Y")   //将！=改为==
                {
                    dt = null;
                    Type = "CheckFeeder";
                    dt = PD.CheckFeeder(TxtDID.Text.ToString().Trim(), Type);
                    if (dt.Rows.Count > 0)
                    {
                        Line = dt.Rows[0]["line"].ToString().Trim();
                    }
                    else
                    {
                        MessageBox.Show("This DID not have dispatch record");
                        return;
                    }
                    dt = null;
                    Type = "CheckFeederLine";
                    dt = PD.CheckFeederLine(TxtFeeder.Text.ToString().Trim(), Type);
                    if (dt.Rows.Count > 0)
                    {
                        if (PublicFunction.ConfigListGetValue("ChkFeederLine") == "Y")//将！=改为==
                        {
                            if (Line != dt.Rows[0]["line"].ToString().Trim())
                            {
                                MessageBox.Show("This Feeder must be used in line: " + dt.Rows[0]["line"].ToString().Trim() + " !");
                                return;
                            }
                            dt = null;
                            Type = "CheckMaintain";
                            dt = PD.CheckFeederLine(TxtFeeder.Text.ToString().Trim(), Type);
                            if (dt.Rows.Count > 0)
                            {
                                MessageBox.Show("This feeder must be repair first!");
                                return;
                            }
                        }
                    }
                    dt = null;
                    Type = "CheckFeederData";
                    dt = PD.CheckFeederLine(TxtFeeder.Text.ToString().Trim(), Type);
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["ErrorCode"].ToString().Trim() != "0")
                        {
                            if (dt.Rows[0]["ErrorCode"].ToString().Trim() == "100")
                            {
                                MessageBox.Show(dt.Rows[0]["Result"].ToString());
                                TxtFeeder.Text = "";
                                TxtFeeder.Focus();
                                return;
                            }
                            else
                            {
                                MessageBox.Show(dt.Rows[0]["Result"].ToString());
                                return;
                            }
                        }
                    }
                }
                LR = GetLRMapping();
                if (PublicFunction.ConfigListGetValue("MaintainFeederDID") != "Y" || OptFuJi.Checked != true)
                {
                    dt = null;
                    Type = "QueryWOInfo";
                    dt = PD.QueryWOInfo(cboJobGroup.Text.ToString().Trim(), TxtCompPN.Text.ToString().Trim(), CboMachine.Text.ToString().Trim(), TxtGroup.Text.ToString().Trim(), 1, Type);
                    if (chkByjobGroup.Checked == true)
                    {
                        dt = PD.QueryWOInfo(cboJobGroup.Text.ToString().Trim(), TxtCompPN.Text.ToString().Trim(), CboMachine.Text.ToString().Trim(), TxtGroup.Text.ToString().Trim(), 0, Type);
                    }
                    if (dt.Rows.Count > 0)
                    {
                        if (CboMachine.Text.ToString().Trim().IndexOf("MSF") > 0)
                        {
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                Slot += dt.Rows[i]["Slot"].ToString().Trim() + ",";
                            }
                        }
                        else
                        {
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                if (LR == dt.Rows[i]["LR"].ToString().Trim())
                                    Slot += dt.Rows[i]["Slot"].ToString().Trim() + ",";
                            }
                            MessageBox.Show("LR is error, Please check the LR!!");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Component is error,Please check the Machine Name!!");
                        return;
                    }
                    if (Slot != "")
                    {
                        Slot = Slot.Substring(0, Slot.Length - 1);
                    }
                    else
                    {
                        MessageBox.Show("Component is error,Please check the Machine Name!!");
                        return;
                    }
                }
                else if (PublicFunction.ConfigListGetValue("MaintainFeederDID") == "Y" && OptFuJi.Checked == true)
                {
                    dt = null;
                    Type = "QueryWO";
                    dt = PD.QueryWO(TxtCompPN.Text.ToString().Trim(), Type);
                    if (dt.Rows.Count <= 0)
                    {
                        MessageBox.Show("Component is error,Please check the Machine Name!!");
                        return;
                    }
                    dt = null;
                    Type = "QueryProConfig";
                    dt = PD.QueryProConfig(TxtFeeder.Text.ToString().Trim(), Line, Type);
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0]["result"].ToString().Trim() != "0")
                        {
                            MessageBox.Show(dt.Rows[0]["msg"].ToString());
                            return;
                        }
                    }
                }
                dt = null;
                Type = "SevenFeederData";
                dt = PD.SevenFeederData(CboMachine.Text.ToString().Trim(), cboJobGroup.Text.ToString().Trim(), TxtDID.Text.ToString().Trim(), TxtCompPN.Text.ToString().Trim(), VendorCode, DateCode, LotCode, TxtFeeder.Text.ToString().Trim(), Slot.Substring(0, 20), LR, Parameter.g_userName, OptPanal.Checked, Line, Type);
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["Result"].ToString().Trim() == "1")
                    {
                        MessageBox.Show("Component is error,Please check the Machine Name!!");
                        TxtDID.Text = "";
                        TxtDID.Focus();
                        return;
                    }
                }
                dt = null;
                Type = "SevenLOG";
                PD.SevenLOG(CboLine.Text.ToString().Trim(), TxtDID.Text.ToString().Trim(), CboMachine.Text.ToString().Trim(), TxtFeeder.Text.ToString().Trim(), Parameter.g_userName, Type);
                RefreshMachineFeeder();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                TxtDID.Text = "";
                TxtDID.Focus();
                return;
            }
        }
        private bool ChkValid()
        {
            DgDIDSlot.DataSource = null;
            bool ChkValid = true;
            dt = PD.GenUNID(TxtDID.Text.ToString().Trim(), "");
            if (dt.Rows[0]["Result"].ToString().Trim() != "OK")
            {
                MessageBox.Show(dt.Rows[0]["Msg"].ToString());
                ChkValid = false;
                return ChkValid;
            }
            else
            {
                TxtDID.Text = dt.Rows[0]["UNID"].ToString().Trim();
            }
            if (ChkDID(TxtDID.Text.ToString().Trim()) == false)
            {
                ChkValid = false;
            }
            if (PublicFunction.ConfigListGetValue("MaintainFeederDID") != "Y" || OptFuJi.Checked != true)
            {
                if (OptFuJi.Checked == true)
                {
                    if (TxtLR.Text.ToString().Trim() == "L" || TxtLR.Text.ToString().Trim() == "R" || TxtLR.Text.ToString().Trim() == "0")
                    { }
                    else
                    {
                        MessageBox.Show("LR Is invalid,Pleaes check");
                        TxtLR.Enabled = true;
                        TxtLR.Focus();
                        ChkValid = false;
                    }
                }
                if (cboJobGroup.Text.ToString().Trim() == "" || CboMachine.Text.ToString().Trim() == "")
                {
                    MessageBox.Show("JObPN or Machine or Version can not be empty,Please check");
                    ChkValid = false;
                }
            }
            Type = "ChkIfInCurretnFeeder";
            dt = PD.ChkIfInCurretnFeeder(TxtFeeder.Text.ToString().Trim(), TxtDID.Text.ToString().Trim(), Type);
            if (dt.Rows.Count > 0)
            {
                DgDIDSlot.DataSource = dt;
            }
            else
            {
                MessageBox.Show("The DID or Feeder In machine use,Please clear the link relationship on DID & Slot Link");
                ChkValid = false;
            }
            return ChkValid;
        }
        private bool ChkDID(string DID)
        {
            dt = null;
            bool ChkDID = true;
            if (ChkNonAVL(DID, TxtCustomer.Text.ToString().Trim(), TxtModel.Text.ToString().Trim(), TxtMBPN.Text.ToString().Trim(), cboWO.Text.ToString().Trim()) == false)
            {
                ChkDID = false;
                return ChkDID;
            }
            Type = "CheckDIDValidity";
            dt = PD.CheckDIDValidity(DID, CboMachine.Text.ToString().Trim(), CboGroupID.Text.ToString().Trim(), cboWO.Text.ToString().Trim(), CboLine.Text.ToString().Trim(), Type);
            if (dt.Rows[0]["Result"].ToString().Trim().Substring(0, 4) == "PASS")
            {
                CompPN = dt.Rows[0]["DIDCompPN"].ToString().Trim();
                VendorCode = dt.Rows[0]["VendorCode"].ToString().Trim();
                DateCode = dt.Rows[0]["DateCode"].ToString().Trim();
                LotCode = dt.Rows[0]["LotCode"].ToString().Trim();
            }
            else
            {
                MessageBox.Show(dt.Rows[0]["Result"].ToString());
                ChkDID = false;
                return ChkDID;
            }
            return ChkDID;
        }
        private bool ChkNonAVL(string DID, string Customer, string Model, string MBPN, string WO)
        {
            dt = null;
            bool ChkNonAVL = true;
            Type = "ChkNonAVL";
            dt = PD.ChkNonAVL(DID, Type);
            if (dt.Rows.Count > 0)
            {
                CompPN = dt.Rows[0]["CompPN"].ToString().Trim();
                VendorCode = dt.Rows[0]["VendorCode"].ToString().Trim();
                DateCode = dt.Rows[0]["DateCode"].ToString().Trim();
                LotCode = dt.Rows[0]["LotCode"].ToString().Trim();
            }
            else
            {
                MessageBox.Show("Can not find the DID,Please check");
                ChkNonAVL = false;
                return ChkNonAVL;
            }
            dt = null;
            Type = "ChkNonAVLData";
            dt = PD.ChkNonAVLData(TxtCustomer.Text.ToString().Trim(), TxtCompPN.Text.ToString().Trim(), MBPN, Model, VendorCode, DateCode, LotCode, WO, Type);
            if (dt.Rows.Count > 0)
            {
                ChkNonAVL = true;
                return ChkNonAVL;
            }
            else
            {
                ChkNonAVL = false;
            }
            if (PublicFunction.ConfigListGetValue("Check_NonAVL") != "Y")
            {
                ChkNonAVL = true;
            }
            if (ChkNonAVL == false)
            {
                MessageBox.Show("Check NonAVL failed");
            }
            return ChkNonAVL;
        }
        private string GetLRMapping()
        {
            string GetLRMapping = "";
            if (OptFuJi.Checked == true)
            {
                GetLRMapping = "0";
                return GetLRMapping;
            }
            if (OptPanal.Checked == true)
            {
                if (TxtLR.Text.ToString().Trim() == "0")
                {
                    GetLRMapping = "0";
                    return GetLRMapping;
                }
                else if (TxtLR.Text.ToString().Trim() == "L")
                {
                    GetLRMapping = "1";
                    return GetLRMapping;
                }
                else if (TxtLR.Text.ToString().Trim() == "R")
                {
                    GetLRMapping = "2";
                    return GetLRMapping;
                }
            }

            return GetLRMapping;
        }
        private bool ChkAVL(string CompPN, string VendorCode, string Customer, string Model)
        {
            bool ChkAVL = true;
            bool ControlPart = false;
            dt = null;
            Type = "AVL_Vendor";
            dt = PD.AVL_Vendor(Customer, Type);
            if (dt.Rows.Count > 0)
            {
                AVLCustomer = dt.Rows[0]["avl_customer"].ToString().Trim();
                ModelFlag = dt.Rows[0]["ModelFlag"].ToString().Trim();
            }
            else
            {
                AVLCustomer = "QUANTA";
            }
            if (AVLCustomer.ToUpper() != "QUANTA")
            {
                Type = "QueryAVL";
                if (PublicFunction.ConfigListGetValue("ModelFlag") != "Y")
                {
                    dt = PD.QueryAVL(CompPN, VendorCode, Customer, Model, 1, Type);
                }
                else
                {
                    dt = PD.QueryAVL(CompPN, VendorCode, Customer, Model, 0, Type);
                }
                if (dt.Rows.Count <= 0)
                {
                    ChkAVL = false;
                }
            }
            dt = null;
            Type = "QueryControlPart";
            dt = PD.QueryControlPart(CompPN, Model, Type);
            if (dt.Rows.Count <= 0)
            {
                ChkAVL = true;
                return ChkAVL;
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i]["VendorCode"].ToString().Trim().ToUpper() == VendorCode.ToUpper())
                {
                    ControlPart = true;
                }
            }
            if (ControlPart == true)
            {
                ChkAVL = true;
            }
            else
            {
                ChkAVL = false;
            }
            ChkAVL = true;
            return ChkAVL;
        }

        private void frmMaintainFeeder_FormClosed(object sender, FormClosedEventArgs e)
        {
            PublicFunction.RemoveForm("frmMaintainFeeder");
        }
    }
}
