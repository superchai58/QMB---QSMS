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
    public partial class frmTransferDispatchDID : Form
    {

        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.MCC.MCCPreMaterialProcess mccProcess = new DbLibrary.MCC.MCCPreMaterialProcess();
        string strWO = "",msg = "", strGetDIDFromSourceBU = "";

        public frmTransferDispatchDID()
        {
            InitializeComponent();
        }
        private void frmTransferDispatchDID_Load(object sender, EventArgs e)
        {
            dtpSDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
            dtpEDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
            strGetDIDFromSourceBU = pubFunction.ReadIniFile("QSMS", "GetDIDFromSourceBU", AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "set.ini");
            GetLine();
        }

        private void GetLine()
        {
            DataTable dt = mccProcess.GetLine("MCC_GetLine");
            cboLine.Items.Clear();
            foreach (DataRow dr in dt.Rows)
            {
                cboLine.Items.Add(dr["Line"].ToString());
            }
        }

 #region Find

        private void cmdQuery_Click(object sender, EventArgs e)
        {
            if (cboLine.SelectedItem == null)
            {
                MessageBox.Show("Please input line!");
                return;
            }
            string SDate = dtpSDate.Value.ToString("yyyyMMdd");
            string Edate = dtpEDate.Value.ToString("yyyyMMdd");
            DataTable dt = null;
            if (Parameter.BU == "NB5")
            {
                if (optRelease.Checked == true)
                {
                    dt = mccProcess.GetGroupID(SDate, Edate, cboLine.SelectedItem.ToString(), "release", "NB5");
                }
                else
                {
                    dt = mccProcess.GetGroupID(SDate, Edate, cboLine.SelectedItem.ToString(), "", "NB5");
                }
            }
            else
            {
                if (optRelease.Checked == true)
                {
                    dt = mccProcess.GetGroupID(SDate, Edate, cboLine.SelectedItem.ToString(), "release", "");
                }
                else
                {
                    dt = mccProcess.GetGroupID(SDate, Edate, cboLine.SelectedItem.ToString(), "", "");
                }
            }
            cboGroupID.Items.Clear();
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    cboGroupID.Items.Add(dr["GroupID"].ToString().Trim());
                }
            }
            else
            {
                MessageBox.Show("No data");
            }
        }
#endregion


#region  GroupID

        private void cboGroupID_Click(object sender, EventArgs e)
        {
            if (cboGroupID.SelectedItem != null)
            {
                GetWoGroup(cboGroupID.SelectedItem.ToString());
            }
        }

        private void cboGroupID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 || e.KeyChar == 9)            
            {
                cboGroupID_Click(sender, e);
            }
        }

        private void cboGroupID_TextChanged(object sender, EventArgs e)
        {
            if (cboGroupID.Text.Length > 12)
            {
                cboGroupID_Click(sender, e);
            }
        }

        private void GetWoGroup(string GroupID)
        {            
            cboClosed.Items.Clear();
            DataTable dt = mccProcess.GetWOListByGroupID(GroupID, "MCC_GetWOListByGroupID");
            cboWO.Items.Clear();
            cboNotFinishedWO.Items.Clear();
            cboNotChkBom.Items.Clear();
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["ClosedFlag"].ToString() == "Y")
                    {
                        cboClosed.Items.Add(dr["Work_Order"].ToString().Trim());     //"ClosedFlag"
                    }
                    else
                    {                       
                            if (ChkWo(dr["Work_Order"].ToString(), "MCC_ChkQSMS_WO") == false)
                            {
                                cboNotChkBom.Items.Add(dr["Work_Order"].ToString().Trim());
                            }
                            else
                            {
                                if (ChkWo(dr["Work_Order"].ToString(), "MCC_ChkWoFinished") == true)
                                {
                                    cboWO.Items.Add(dr["Work_Order"].ToString().Trim());
                                }
                                else
                                {
                                    cboNotFinishedWO.Items.Add(dr["Work_Order"].ToString().Trim());
                                }
                            }                       
                    }
                }
            }
        }

        private bool ChkWo(string wo, string type)
        {
            if (type == "MCC_ChkMBWo")
            {
                if (mccProcess.ChkWo(wo, type).Rows.Count > 0)
                {
                    return true;
                }
            }
            else if (type == "MCC_ChkQSMS_WO")
            {
                if (mccProcess.ChkWo(wo, type).Rows.Count > 0)
                {
                    return true;
                }
            }
            else if (type == "MCC_ChkWoFinished")
            {
                DataTable dt = mccProcess.GetWoFinishedFlag(wo, "MCC_GetWoFinishedFlag");
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["WofinishedFlag"].ToString() == "Y")
                    {
                        if (mccProcess.GetWoFinishedFlag(wo, "MCC_GetWoGroupFinishedFlag").Rows.Count <= 0)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

#endregion


        private void cboNotChkBom_Click(object sender, EventArgs e)
        {
            if (cboNotChkBom.SelectedItem != null)
            {
                GetSBWO(cboNotChkBom.SelectedItem.ToString());
                //GetWOInfo(cboNotChkBom.SelectedItem.ToString());      //待定模块
            }
        }
        private void GetSBWO(string wo)
        {
            string Group = "";
            //cboSBWO.Items.Clear();
            panel2.Visible = false;
            DataTable dt = mccProcess.GetGroupByWO(wo, "MCC_GetGroupByWO");
            if (dt.Rows.Count > 0)
            {
                txtVersion.Text = dt.Rows[0]["MB_Rev"].ToString().Trim();
                Group = dt.Rows[0]["Group"].ToString().Trim();                
            }
            dt = null;
            dt = mccProcess.GetWOByGroupAndWO(wo, Group, "MCC_GetWOByGroupAndWO");
            if (dt.Rows.Count > 0)
            {
                //foreach (DataRow dr in dt.Rows)
                //{
                //    cboSBWO.Items.Add(dr["WO"].ToString().Trim());
                //}
                //if (cboSBWO.Items.Count > 0)
                //{
                //    panel2.Visible = true;                //待添加
                //}
            }
        }

        private void cboWO_Click(object sender, EventArgs e)
        {
            if (txtWO.Text != "")
            {
                GetSBWO(txtWO.Text);
                //GetWOInfo(txtWO.Text);
                string woStr = GetWOArray(txtWO.Text);
                GetMachine(woStr);
            }
        }

        private void cboWO_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                cboWO_Click(sender, e);
            }
        }
        private string GetWOArray(string wo)
        {
            string woArray = "";
            DataTable dtWO = mccProcess.GetWoArray(wo, "MCC_GetWoArray");
            if (dtWO.Rows.Count > 0)
            {
                foreach (DataRow dr in dtWO.Rows)
                {
                    woArray = woArray + "'" + dr["WO"].ToString().Trim() + "',";
                }
            }
            return woArray.Substring(0, woArray.Length - 1);
        }

        private void GetMachine(string woStr)
        {
            DataTable dtWO = null;
            dtWO = null;
            dtWO = mccProcess.GetMachineFlag(woStr, "MCC_GetMachineFlag");
            if (dtWO.Rows.Count > 0)
            {
                foreach (DataRow dr in dtWO.Rows)
                {                    
                        if (dr["MachinefinishedFlag"].ToString().ToUpper() == "N")
                        {
                            cboNewMachine.Items.Add(dr["Machine"].ToString().Trim());
                        }                                    
                }
            }
        }

        private void cboNotFinishedWO_Click(object sender, EventArgs e)
        {
            txtWO.Text = cboNotFinishedWO.Text.Trim();
            GetSBWO(txtWO.Text.Trim());
            //GetWOInfo(txtWO.Text.Trim());
            string wostr = GetWOArray(txtWO.Text.Trim());
            GetMachine(wostr);
        }

        private void cboNotFinishedWO_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                cboNotFinishedWO_Click(sender, e);
            }
        }

        private void txtDID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 || e.KeyChar == 9)
            {
                if (txtDID.Text == "") return;
                txtDID.Text = txtDID.Text.Trim();
                if (GetDIDInfo(txtDID.Text, txtWO.Text) == false)         
                {
                    txtDID.Text = "";
                    txtDID.Focus();
                }
                else
                {
                    cboNewMachine.Focus();
                }
            }
        }


        private bool GetDIDInfo(string DID,string WO)
        {
            txtDispatchQty.Text = "";
            txtLine.Text = "";
            txtSide.Text = "";
            txtJobPN.Text = "";
            txtCompPN.Text = "";
            txtDIDDateTime.Text = "";
            txtDIDTotalQty.Text = "";
            txtDispatchQty.Text = "";
            txtVendorCode.Text = "";
            txtDateCode.Text = "";
            txtLotCode.Text = "";
            txtMachine.Text="";
            txtSlot.Text = "";
            txtLR.Text = "";
            return true;                 //return位置对照VB：GetDIDInfo = True
            DataTable dt = mccProcess.GetDIDInfo(DID, "MCC_GetDIDInfo1");    //Str = "select CompPN,VendorCode,DateCode,LotCode...
            if (dt.Rows.Count > 0)
            {
                txtCompPN.Text = dt.Rows[0]["CompPN"].ToString();
                txtVendorCode.Text = dt.Rows[0]["VendorCode"].ToString();
                txtDateCode.Text = dt.Rows[0]["DateCode"].ToString();
                txtLotCode.Text = dt.Rows[0]["LotCode"].ToString();
                txtDIDTotalQty.Text = dt.Rows[0]["Qty"].ToString();
                txtDIDDateTime.Text = dt.Rows[0]["TransDateTime"].ToString();
            }
            else
            {
                DataTable dt1 = mccProcess.GetDIDInfo(DID, "MCC_GetDIDInfo1");
                if (dt.Rows.Count > 0)
                {
                    txtCompPN.Text = dt1.Rows[0]["CompPN"].ToString();
                    txtVendorCode.Text = dt1.Rows[0]["VendorCode"].ToString();
                    txtDateCode.Text = dt1.Rows[0]["DateCode"].ToString();
                    txtLotCode.Text = dt1.Rows[0]["LotCode"].ToString();
                    txtDIDTotalQty.Text = dt1.Rows[0]["Qty"].ToString();
                    txtDIDDateTime.Text = dt1.Rows[0]["TransDateTime"].ToString();
                }
                else
                {
                    MessageBox.Show("Can't find this DID,please check!");
                    return false;                                           //return位置对照VB：GetDIDInfo = False
                }              
            }
            DataTable dt2 = mccProcess.GetDispatchInfoAll(DID, WO, "MCC_GetDIDDispatchInfo");    //select Work_Order,Line,JobPN,Machine...
            if (dt.Rows.Count > 0)
            {
                    txtDispatchQty.Text = dt2.Rows[0]["DIDQty"].ToString();
                    txtLine.Text = dt2.Rows[0]["Line"].ToString();
                    txtSide.Text = dt2.Rows[0]["Side"].ToString();
                    txtJobPN.Text = dt2.Rows[0]["JobPN"].ToString();
                    txtMachine.Text =dt2.Rows[0]["Machine"].ToString();
                    txtSlot.Text = dt2.Rows[0]["Slot"].ToString();
                    txtLR.Text = dt2.Rows[0]["LR"].ToString();
            }
            else
            {
                MessageBox.Show("Can't find this DID,please check!");
                return false;
            }
        }


        private void cboNewMachine_Click(object sender, EventArgs e)
        {
            string TransDate = "";
            string WO = txtWO.Text;
            string machine = cboNewMachine.Text;
            cboNewSlot.Items.Clear();
            cboNewLR.Items.Clear();
            if (cboNewMachine.SelectedItem != null)
            {
                GetSlot(WO, machine);  // 函数wo machine 对照VB  
            }
            
         }

        private void GetSlot(string WO, string Machine)     
        {
            DataTable dt = mccProcess.GetSlotInfo(WO,Machine);
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    cboNewSlot.Items.Add(dr["Slot"].ToString().Trim());
                }
            }
            cboNewSlot.Items.Add("0");
            cboNewSlot.Items.Add("1");
            cboNewSlot.Items.Add("2");
        }

        private void cmdOK_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtDID.Text == "")
                {
                    MessageBox.Show("Please input DID!");
                    return;
                }
                if (txtWO.Text == "" || txtLine.Text == "" || cboNewMachine.Text == "" || cboNewSlot.Text == "" || cboNewLR.Text == "" || txtCompPN.Text == "" || txtMachine.Text == "" || txtSlot.Text == "" || txtDispatchQty.Text == "")
                {
                    MessageBox.Show("Please input the machine & slot infomation!");
                    return;
                }
                TransferDispatchDID(txtMachine.Text, txtSlot.Text, txtLR.Text, txtWO.Text, txtCompPN.Text, txtJobPN.Text, txtDID.Text, txtDispatchQty.Text, cboNewMachine.SelectedItem.ToString(),cboNewSlot.SelectedItem.ToString(),cboNewLR.SelectedItem.ToString(),txtVersion.Text);
                UpdateMachineFlagByWO(txtWO.Text);
                
                MessageBox.Show("OK!");
                ClearDataCmdOK();
                return;            //Exit Sub

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ",Please contact QSMS SMT Staff");
            }
        }

        private void TransferDispatchDID(string Machine, string Slot, string LR, string WO, string CompPN, string JobPN, string DID, string DispatchQty,string NewMachine,string NewSlot,string NewLR ,string Version)
        {
            DataTable dt = mccProcess.TransferDispatchDIDInfo(Machine, Slot, LR, WO, CompPN, JobPN, DID, DispatchQty, NewMachine, NewSlot, NewLR, Version);
        }
        private void UpdateMachineFlagByWO(string WO)
        {
            string Machine = "";
            DataTable dtflag = mccProcess.UpdateMachineFlagByWOInfo(WO,"", "MachineFlag");
            if (dtflag.Rows.Count > 0)
            {
                Machine = dtflag.Rows[0]["Machine"].ToString();
                DataTable dtflag1 = mccProcess.UpdateMachineFlagByWOInfo(WO,Machine,"MachineFlag1");
                if (dtflag1.Rows.Count > 0)
                {
                    mccProcess.UpdateMachineFlagByWOInfo1(WO, Machine);
                }
                else
                {
                    mccProcess.UpdateMachineFlagByWOInfo2(WO, Machine);
                }
            }
            DataTable dt = mccProcess.UpdateMachineFlagByWOInfoAll(WO);
            if (dt.Rows.Count > 0)
            {
                mccProcess.UpdateMachineFlagByWOInfoAll1(WO);
            }
            else 
            {
                mccProcess.UpdateMachineFlagByWOInfoAll2(WO);
            }          
        }

        public void ClearDataCmdOK()
        {
            txtDID.Text = "";
            txtDispatchQty.Text = "";
            txtCompPN.Text = "";
            txtVendorCode.Text = "";
            txtDateCode.Text = "";
            txtLotCode.Text = "";
            txtDIDDateTime.Text = "";
            txtDIDTotalQty.Text = "";
            txtLine.Text = "";
            txtSide.Text = "";
            txtJobPN.Text = "";
            txtMachine.Text = "";
            txtSlot.Text = "";
            txtLR.Text = "";
        }


    }
}
