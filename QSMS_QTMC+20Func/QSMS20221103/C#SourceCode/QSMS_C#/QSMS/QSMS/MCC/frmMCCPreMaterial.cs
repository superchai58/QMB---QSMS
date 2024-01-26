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
    public partial class frmMCCPreMaterial : Form
    {
        DbLibrary.MCC.MCCPreMaterialProcess process = new DbLibrary.MCC.MCCPreMaterialProcess();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        string msg = "", strGetDIDFromSourceBU="";
        public static string PN="";
        public frmMCCPreMaterial()
        {
            InitializeComponent();
        }

        private void frmMCCPreMaterial_Load(object sender, EventArgs e)
        {
            dtpSDate.Text= DateTime.Now.ToString("yyyy/MM/dd");
            dtpEDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
            strGetDIDFromSourceBU = pubFunction.ReadIniFile("QSMS", "GetDIDFromSourceBU", AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "set.ini");
            GetLine();
        }

        private void frmMCCPreMaterial_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmMCCPreMaterial");
        }

        private void cboGroupID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                cboGroupID_Click(sender, e);
            }
        }

        private void cboGroupID_Click(object sender, EventArgs e)
        {
            if (cboGroupID.SelectedItem != null)
            {
                GetWoGroup(cboGroupID.SelectedItem.ToString());
            }
        }

        private void cboJob_Click(object sender, EventArgs e)
        {
            if (cboJob.SelectedItem != null)
            {
                if (!string.IsNullOrEmpty(txtMachine.Text) && !(string.IsNullOrEmpty(txtLine.Text))&&!(string.IsNullOrEmpty(txtWO.Text)))
                {
                    if (GetDID(txtMachine.Text, cboJob.SelectedItem.ToString(), txtWO.Text, txtLine.Text)==0)
                    {
                        cmdRefresh_Click(sender, e);
                    }
                }
            }
        }

        private void cboJob_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                cboJob_KeyPress(sender, e);
            }
        }

        private void cmdRefresh_Click(object sender, EventArgs e)
        {
            string woStr = "",woStr1="";
            if (txtWO.Text != "")
            {
                woStr = GetWOArray(txtWO.Text.Trim());
            }
            woStr1 = woStr.Replace("'","");
            if (txtMachine.Text != "")
            {
                process.UpdateDispatchFlag(txtMachine.Text.Trim(), woStr1);
            }
            GetMachine(woStr);
            if (cboGroupID.SelectedItem!= null)
            {
                GetGroupWO(cboGroupID.SelectedItem.ToString());
            }
        }

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
                    dt = process.GetGroupID(SDate, Edate, cboLine.SelectedItem.ToString(), "release", "NB5");
                }
                else
                {
                    dt = process.GetGroupID(SDate, Edate, cboLine.SelectedItem.ToString(), "", "NB5");
                }
            }
            else
            {
                if (optRelease.Checked == true)
                {
                    dt = process.GetGroupID(SDate, Edate, cboLine.SelectedItem.ToString(), "release", "");
                }
                else
                {
                    dt = process.GetGroupID(SDate, Edate, cboLine.SelectedItem.ToString(), "", "");
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

        private void cboMachineNOK_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                cboMachineNOK_Click(sender, e);
            }
        }

        private void cboMachineNOK_Click(object sender, EventArgs e)
        {
            if (cboMachineNOK.SelectedItem != null)
            {
                txtMachine.Text = cboMachineNOK.SelectedItem.ToString().Trim();
                if (txtWO.Text != "")
                {
                    if (ChkWOSeq(txtWO.Text, txtMachine.Text) == false)
                    {
                        MessageBox.Show("The Work order has No group ID, please call PMC! ");
                        txtDID.Enabled = false;
                        return;
                    }
                    else
                    {
                        txtDID.Enabled = true;
                    }
                    GetJobForBuildType(txtWO.Text, txtMachine.Text, txtBuildType.Text);
                    GetDID(txtMachine.Text, cboJob.SelectedItem.ToString(), txtWO.Text, txtLine.Text);
                    txtDID.Text = "";
                    txtDID.Focus();
                    cboMachineOK.Text = "";
                }

            }
        }

        private void cboMachineOK_Click(object sender, EventArgs e)
        {
            if (cboMachineOK.SelectedItem != null)
            {
                if (ChkWOSeq(txtWO.Text, txtMachine.Text) == false)
                {
                    txtDID.Enabled = false;
                }
                else
                {
                    txtDID.Enabled = true;
                }
                GetJobForBuildType(txtWO.Text, txtMachine.Text, txtBuildType.Text);
                GetDID(txtMachine.Text, cboJob.SelectedItem.ToString(), txtWO.Text, cboLine.SelectedItem.ToString());
                cboWithout.Items.Clear();
                txtDID.Text = "";
                txtDID.Focus();
                cboMachineNOK.Items.Clear();
            }
        }

        private void cboMachineOK_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                cboMachineOK_Click(sender, e);
            }
        }

        private void cboNotChkBom_Click(object sender, EventArgs e)
        {
            if (cboNotChkBom.SelectedItem != null)
            {
                GetSBWO(cboNotChkBom.SelectedItem.ToString());
                GetWOInfo(cboNotChkBom.SelectedItem.ToString());
            }
        }

        private void cboNotFinishedWO_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                cboNotFinishedWO_Click(sender, e);
            }
        }

        private void cboNotFinishedWO_Click(object sender, EventArgs e)
        {
            if (txtWO.Text != "")
            {
                GetSBWO(txtWO.Text.Trim());
                GetWOInfo(txtWO.Text.Trim());
                string woList = GetWOArray(txtWO.Text.Trim());
                GetMachine(woList);
            }
        }

        private void cboWithout_Click(object sender, EventArgs e)
        {
            if (txtGroup.Text != "" && txtMachine.Text != "" && cboWithout.SelectedItem != null)
            {
                DataTable dt = process.GetWithOutQty(txtGroup.Text, txtMachine.Text, cboWithout.SelectedItem.ToString(), "MCC_GetWithOutQty");
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        txtWTotal.Text = dr["NeedQty"].ToString().Trim();
                        txtWBalance.Text = dr["BalanceQty"].ToString().Trim();
                    }
                }
            }
        }

        private void cboWO_Click(object sender, EventArgs e)
        {
            if (txtWO.Text != "")
            {
                GetSBWO(txtWO.Text);
                GetWOInfo(txtWO.Text);
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

        private void cmdAdd_Click(object sender, EventArgs e)
        {
            if (listWONotFinished.Items.Count <= 0)
            {
                return;
            }
            if (listWONotFinished.SelectedItem == null)
            {
                return;
            }
            listWoDispatching.Items.Add(listWONotFinished.SelectedItem.ToString());
            listWONotFinished.Items.Remove(listWONotFinished.SelectedItem.ToString());
        }

        private void cmdAddAll_Click(object sender, EventArgs e)
        {
            if (listWONotFinished.Items.Count <= 0)
            {
                return;
            }
            do
            {
                listWONotFinished.SelectedIndex = 0;
                listWoDispatching.Items.Add(listWONotFinished.SelectedItem);
                listWONotFinished.Items.RemoveAt(0);
            } while (listWONotFinished.Items.Count > 0);

        }

        private void cmdDel_Click(object sender, EventArgs e)
        {
            if (listWoDispatching.Items.Count <= 0)
            {
                return;
            }
            if (listWoDispatching.SelectedItem == null)
            {
                return;
            }
            listWONotFinished.Items.Add(listWoDispatching.SelectedItem.ToString());
            listWoDispatching.Items.Remove(listWoDispatching.SelectedItem.ToString());
        }

        private void cmdDelAll_Click(object sender, EventArgs e)
        {
            if (listWoDispatching.Items.Count <= 0)
            {
                return;
            }
            do
            {
                listWoDispatching.SelectedIndex = 0;
                listWONotFinished.Items.Add(listWoDispatching.SelectedItem);
                listWoDispatching.Items.RemoveAt(0);
            } while (listWoDispatching.Items.Count > 0);
        }

        private void cmdConfirm_Click(object sender, EventArgs e)
        {
            try
            {
                if (ChkErr(ref msg) == true)
                {
                    if (Insert_QSMS_Out() == true)
                    {
                        pubFunction.Sound("OK");
                        LblMessage.Text = "insert OK";
                    }
                    else
                    {
                        pubFunction.Sound("ERROR");
                        LblMessage.Text = "didn't dispatch the DID";
                    }
                }
                else
                {
                    MessageBox.Show(msg);
                    pubFunction.Sound("ERROR");
                }
                cboMachineNOK.Text = txtMachine.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ",Please contact QSMS SMT Staff");
                return;
            }
        }

        private void txtDID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                if (txtDID.Text == "") return;
                string DID = txtDID.Text.Trim();
                if (strGetDIDFromSourceBU == "Y")
                {
                    GetDIDFromSourceBU(DID);
                }

                if (ChkDIDBelongMachine(txtWO.Text, txtMachine.Text.Trim(), DID, txtGroup.Text) == false)
                {
                    txtDID.Text = "";
                    txtDID.Focus();
                    pubFunction.Sound("ERROR");
                    return;
                }
                GetDIDInfo(DID, txtWO.Text);
                if (Parameter.Check_DID == "Y")
                {
                    pubFunction.HaveOpened(new QSMS.MCC.frmInputBox(), "frmInputBox");
                    if (PN.ToUpper() != txtCompPN.Text.ToUpper())
                    {
                        pubFunction.Sound("ERROR");
                        MessageBox.Show("DID and CompPN aren't matched,Please check DID and CompPN");
                        txtDID.Text = "";
                        txtDID.Focus();
                        return;
                    }
                }
                if (ChkAVL(txtCompPN.Text.Trim(), txtVendorCode.Text.Trim(), txtCustomer.Text.Trim(), txtModel.Text.Trim()) == false)
                {
                    pubFunction.Sound("ERROR");
                    MessageBox.Show("Check AVL failed,please check");
                    return;
                }
                GetCompDispInfo(txtWO.Text.Trim(), cboJob.SelectedItem.ToString().Trim(), txtMachine.Text.Trim(), txtCompPN.Text.Trim());
                txtDispatchQty.Focus();
            }
        }

        private void txtChkDID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                listDIDDtatus.Items.Clear();
            }
            string woArray = GetWOArray(txtWO.Text.Trim());
            DataTable dt = process.GetUsedFlag(txtGroup.Text.Trim(), woArray, txtMachine.Text.Trim(), txtDID.Text.Trim(), "MCC_GetUsedFlag");
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["UsedFlag"].ToString().ToUpper() == "Y")
                {
                    listDIDDtatus.Items.Add("Has been Dispatched");
                }
                else
                {
                    listDIDDtatus.Items.Add("Not Dispatched");
                }
            }
            else
            {
                listDIDDtatus.Items.Add("Not belong to the Machine");
            }
        }

        private void cmdExcel_Click(object sender, EventArgs e)
        {
            DataTable dt = null;
            if (txtQryDID.Text == "")
            {
                MessageBox.Show("请输入DID!");
                return;
            }
            if (chkAll.Checked == false)
            {
                if (txtWO.Text != "")
                {
                    dt = process.GetDispatchInfoAll(txtQryDID.Text, txtWO.Text, "");
                }
                else
                {
                    dt = process.GetDispatchInfoAll(txtQryDID.Text, "", "");
                }
            }
            else
            {
                dt = process.GetDispatchInfoAll(txtQryDID.Text, "", "Checked");
            }
            try
            {
                if (dt.Rows.Count > 0)
                {
                    pubFunction.doExport(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ExportToExcel Error:" + ex.Message);
            }
        }

        private void txtQryWO_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                cmdExcel_Click(sender, e);
            }
        }

        private void cmdExcelDID_Click(object sender, EventArgs e)
        {
            if (txtQryWO.Text != "" || txtQryMachine.Text != "")      //Aris 修改条件&&， ||
            {
                DataTable dt = process.GetDispatchInfoAll1(txtQryWO.Text, txtQryMachine.Text, "MCC_GetDispatchInfoAll");
                if (dt.Rows.Count > 0)
                {
                    try
                    {
                        pubFunction.doExport(dt);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("ExportToExcel Error:" + ex.Message);
                    }
                }
            }
        }

        private void txtQryMachine_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                cmdExcelDID_Click(sender, e);
            }
        }

        private void DGCompNotOK_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == -1 || e.RowIndex == -1) return;
                DataGridViewCell cell = this.DGCompNotOK.Rows[e.RowIndex].Cells[0];
                if (cell == null || cell.Value.ToString() == "") return;
                txtCompPN.Text = cell.Value.ToString();
                GetDispatchQtyBySlot(txtWO.Text, txtMachine.Text, txtCompPN.Text);
            }
            catch (Exception ex)
            {
                txtDID.Text = "";
                MessageBox.Show(ex.Message);
            }
        }

        private void DGDIDNotOK_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == -1 || e.RowIndex == -1) return;
                DataGridViewCell cell = this.DGDIDNotOK.Rows[e.RowIndex].Cells[0];
                if (cell == null || cell.Value.ToString() == "") return;
                txtDID.Text = cell.Value.ToString();
                GetDIDInfo(txtDID.Text, txtWO.Text);
                GetCompDispInfo(txtWO.Text, cboJob.SelectedItem.ToString(), txtMachine.Text, txtCompPN.Text);
            }
            catch (Exception ex)
            {
                txtDID.Text = "";
                MessageBox.Show(ex.Message);
            }
        }

        private void DGDIDOK_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.ColumnIndex == -1 || e.RowIndex == -1) return;
                DataGridViewCell cell = this.DGDIDOK.Rows[e.RowIndex].Cells[0];
                if (cell == null || cell.Value.ToString() == "") return;
                txtDID.Text = cell.Value.ToString();
                GetDIDInfo(txtDID.Text, txtWO.Text);
                GetCompDispInfo(txtWO.Text, cboJob.SelectedItem.ToString(), txtMachine.Text, txtCompPN.Text);
            }
            catch (Exception ex)
            {
                txtDID.Text = "";
                MessageBox.Show(ex.Message);
            }
        }
        
        private void GetGroupWO(string groupID)
        {
            listWONotFinished.Items.Clear();
            cboClosed.Items.Clear();
            DataTable dt = process.GetWOListByGroupID(groupID, "MCC_GetWOListByGroupID");
            cboWO.Items.Clear();
            cboNotFinishedWO.Items.Clear();
            cboNotChkBom.Items.Clear();
            if (dt.Rows.Count > 0)
            {
                foreach(DataRow dr in dt.Rows)
                {
                    if (dr["ClosedFlag"].ToString() == "Y")
                    {
                        cboClosed.Items.Add(dr["ClosedFlag"].ToString().Trim());
                    }
                    else
                    {
                        if (ChkWo(dr["Work_Order"].ToString(), "MCC_ChkMBWo") == true)
                        {
                            if (ChkWo(dr["Work_Order"].ToString(), "MCC_ChkQSMS_WO") == false)
                            {
                                cboNotChkBom.Items.Add(dr["Work_Order"].ToString().Trim());
                            }
                            else
                            {
                                if(ChkWo(dr["Work_Order"].ToString(), "MCC_ChkWoFinished") == true)
                                {
                                    cboWO.Items.Add(dr["Work_Order"].ToString().Trim());
                                }
                                else
                                {
                                    listWONotFinished.Items.Add(dr["Work_Order"].ToString().Trim());
                                    cboNotFinishedWO.Items.Add(dr["Work_Order"].ToString().Trim());
                                }
                            }
                        }
                    }
                }
            }
        }
        
        private void GetWoGroup(string GroupID)
        {
            listWONotFinished.Items.Clear();
            cboClosed.Items.Clear();
            DataTable dt = process.GetWOListByGroupID(GroupID, "MCC_GetWOListByGroupID");
            cboWO.Items.Clear();
            cboNotFinishedWO.Items.Clear();
            cboNotChkBom.Items.Clear();
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["ClosedFlag"].ToString() == "Y")
                    {
                        cboClosed.Items.Add(dr["ClosedFlag"].ToString().Trim());
                    }
                    else
                    {
                        if (ChkWo(dr["Work_Order"].ToString(), "MCC_ChkMBWo") == true)
                        {
                            if (ChkWo(dr["Work_Order"].ToString(), "MCC_ChkQSMS_WO") == false)
                            {
                                cboNotChkBom.Items.Add(dr["Work_Order"].ToString().Trim());
                            }
                            else
                            {
                                if(ChkWo(dr["Work_Order"].ToString(), "MCC_ChkWoFinished") == true)
                                {
                                    cboWO.Items.Add(dr["Work_Order"].ToString().Trim());
                                }
                                else
                                {
                                    listWONotFinished.Items.Add(dr["Work_Order"].ToString().Trim());
                                    cboNotFinishedWO.Items.Add(dr["Work_Order"].ToString().Trim());
                                }
                            }
                        }
                    }
                }
            }
        }

        private bool ChkWo(string wo,string type)
        {
            if (type == "MCC_ChkMBWo")
            {
                if (process.ChkWo(wo,type).Rows.Count > 0)
                {
                    return true;
                }
            }else if (type == "MCC_ChkQSMS_WO")
            {
                if (process.ChkWo(wo, type).Rows.Count > 0)
                {
                    return true;
                }
            }else if (type == "MCC_ChkWoFinished")
            {
                DataTable dt = process.GetWoFinishedFlag(wo, "MCC_GetWoFinishedFlag");
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["WofinishedFlag"].ToString() == "Y")
                    {
                        if(process.GetWoFinishedFlag(wo, "MCC_GetWoGroupFinishedFlag").Rows.Count <= 0)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }
        
        private int GetDID(string machine,string jobGroup,string wo,string line)
        {
            label54.Text = "";
            string woStr = GetWOArray(wo);
            DataTable dt = null;
            woStr = woStr.Replace("(", "").Replace(")","");
            if (!string.IsNullOrEmpty(txtGroup.Text) && !string.IsNullOrEmpty(txtModel.Text) && !string.IsNullOrEmpty(txtCustomer.Text))
            {
                dt=process.GetCompPN(txtGroup.Text.ToString(), txtCustomer.Text.ToString(), txtModel.Text.ToString(), woStr, "MCC_GetCompPN");
            }
            DGAVL.DataSource = dt;
            DGAVL.Refresh();
            DataTable dtCompNotOK=process.GetDispatch(txtGroup.Text, woStr, txtMachine.Text, jobGroup, "MCC_GetDispatch");
            label54.Text = "Comp didn't dispatch:"+dtCompNotOK.Rows.Count.ToString();
            DGCompNotOK.DataSource = dtCompNotOK;
            if (dtCompNotOK.Rows.Count == 0)
            {
                this.cmdRefresh.Click += new System.EventHandler(this.cmdRefresh_Click);
            }
            DGCompNotOK.Refresh();
            return dtCompNotOK.Rows.Count;
        }
        
        private string GetWOArray(string wo)
        {
            string woArray = "";
            DataTable dtWO = process.GetWoArray(wo, "MCC_GetWoArray");
            if (dtWO.Rows.Count > 0)
            {
                foreach (DataRow dr in dtWO.Rows)
                {
                    woArray = woArray + "'" + dr["WO"].ToString().Trim() + "',";
                }
            }
            if (listWoDispatching.Items.Count > 0)
            {
                DataTable dt = null;
                for (int i = 0; i < listWoDispatching.Items.Count; i++)
                {
                    dt = process.GetWoArray(listWoDispatching.Items[i].ToString(), "MCC_GetWoArray");
                    woArray = woArray + "'" + dt.Rows[0]["WO"].ToString() + "',";
                    dt = null;
                }
            }
            return woArray.Substring(0, woArray.Length - 1);
        }

        private void GetMachine( string woStr)
        {
            DataTable dtWO = null;
            string tempMachine = "";
            if (txtGroup.Text != "")
            {
                dtWO = process.GetWoByGroup(txtGroup.Text.Trim(), "MCC_GetWoByGroup");
            }
            if (dtWO.Rows.Count > 0)
            {
                woStr = woStr + ",";
                foreach(DataRow dr in dtWO.Rows)
                {
                    woStr=woStr+"'"+dr["Wo"].ToString()+"',";
                }
            }
            woStr = woStr.Substring(0, woStr.Length - 1).ToString();
            dtWO = null;
            dtWO=process.GetMachineFlag(woStr, "MCC_GetMachineFlag");
            cboMachineOK.Items.Clear();
            cboMachineNOK.Items.Clear();
            if (dtWO.Rows.Count > 0)
            {
                foreach(DataRow dr in dtWO.Rows)
                {
                    if(tempMachine=="" || tempMachine != dr["Machine"].ToString().ToUpper())
                    {
                        if (dr["MachinefinishedFlag"].ToString().ToUpper() == "N")
                        {
                            cboMachineNOK.Items.Add(dr["Machine"].ToString().Trim());
                        }
                        else
                        {
                            cboMachineOK.Items.Add(dr["Machine"].ToString().Trim());
                        }
                    }
                    tempMachine = dr["Machine"].ToString().ToUpper();
                }
            }
        }

        private void GetJobForBuildType(string wo, string machine, string buildType)
        {
            cboJob.Items.Clear();
            switch (buildType)
            {
                case "1":
                    label51.Visible = false;
                    cboSlot.Visible = false;
                    break;
                case "2":
                    label51.Visible = true;
                    cboSlot.Visible = true;
                    DataTable dt = process.GetJobGroupByWoAndMachine(wo,machine, "MCC_GetJobGroupByWoAndMachine");
                    if (dt.Rows.Count > 0){
                        foreach(DataRow dr in dt.Rows)
                        {
                            cboJob.Items.Add(dr["JobGroup"].ToString().Trim());
                        }
                    }
                    break;
            }
        }

        private bool ChkWOSeq(string wo, string machine)
        {
            string TempGroupID = "", Seq_No = "";
            DataTable dt = process.GetGroupIdByWO(wo, "MCC_GetGroupIDByWO");
            if (dt.Rows.Count > 0)
            {
                TempGroupID = dt.Rows[0]["GroupID"].ToString();
                Seq_No = dt.Rows[0]["Seq_No"].ToString();
            }
            else
            {
                return false;
            }
            return true;
        }
        
        private void GetWOInfo(string wo)
        {
            DataTable dt=process.GetWOInfoByWO(wo, "MCC_GetWOInfoByWO");
            if (dt.Rows.Count > 0)
            {
                txtMBPN.Text = dt.Rows[0]["PN"].ToString().Trim();
                txtModel.Text = dt.Rows[0]["PN"].ToString().Substring(2,3).Trim();
                txtWoQty.Text = dt.Rows[0]["Qty"].ToString().Trim();
                txtGroup.Text = dt.Rows[0]["Group"].ToString().Trim();
                txtBuildType.Text = dt.Rows[0]["BuildType"].ToString().Trim();
                txtLine.Text = dt.Rows[0]["Line"].ToString().Trim();
            }
            dt = null;
            dt = process.GetTotalQtyByWO(wo, "MCC_GetTotalQtyByWO");
            if (dt.Rows.Count > 0)
            {
                txtPlanQty.Text = dt.Rows[0]["TotalQty"].ToString().Trim();
            }
            else
            {
                txtPlanQty.Text = "";
            }
            dt = null;
            dt = process.GetCustomerByPN(txtMBPN.Text, "MCC_GetCustomerByPN");
            if (dt.Rows.Count > 0)
            {
                txtCustomer.Text = dt.Rows[0]["Customer"].ToString().Trim();
            }
        }

        private void GetSBWO(string wo)
        {
            cboSBWO.Items.Clear();
            panel2.Visible = false;
            DataTable dt = process.GetGroupByWO(wo, "MCC_GetGroupByWO");
            if (dt.Rows.Count > 0)
            {
                txtGroup.Text = dt.Rows[0]["Group"].ToString().Trim();
            }
            dt = null;
            dt = process.GetWOByGroupAndWO(wo, txtGroup.Text, "MCC_GetWOByGroupAndWO");
            if (dt.Rows.Count > 0)
            {
                foreach(DataRow dr in dt.Rows)
                {
                    cboSBWO.Items.Add(dr["WO"].ToString().Trim());
                }
                if (cboSBWO.Items.Count > 0)
                {
                    panel2.Visible = true;
                }
            }
        }
        
        private bool Insert_QSMS_Out()
        {
            bool flag = true;
            int tempDisQty = 0, consumedQty = 0, tempBalanceQty = 0, tempDIDQty=0, count = 0, firstTimeOfDispatch=1, dispQtyByWO=0, balanceQtyByWO=0;
            string transDateTime= DateTime.Now.ToString("yyyyMMddhhmmss");
            string woArray = GetWoArrayForsp();
            string tempWo = "", Item = "",Slot="",LR="",baseQty="";
            DataTable dt = null;
            if (Parameter.BU != "NB5")
            {
                dt = process.GetXLInfo(txtCompPN.Text, "MCC_GetXLInfo");
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show("此材料为XL材料，不能使用此界面发料");
                    return false;
                }
            }
            tempDisQty = Int32.Parse(txtDispatchQty.Text);
            consumedQty = Int32.Parse(txtConsumedQty.Text) + tempDisQty;
            txtConsumedQty.Text = consumedQty.ToString();
            tempBalanceQty = consumedQty - Int32.Parse(txtNeedQty.Text);
            txtBalanceQty.Text = tempBalanceQty.ToString();
            txtDIDRemainQty.Text = (Int32.Parse(txtDIDRemainQty.Text) - tempDisQty).ToString();
            dt = null;
            dt = process.PrepairMaterial(txtGroup.Text,woArray,txtCompPN.Text,cboJob.SelectedItem.ToString(),txtMachine.Text,txtVendorCode.Text);
            if (dt.Rows.Count > 0 && tempDisQty>0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if (ChkDIDDispatchedToWo(dr["Work_Order"].ToString(), txtDID.Text, txtMachine.Text, dr["Slot"].ToString(), dr["LR"].ToString()) == false)
                    {
                        MessageBox.Show("The DID has dispathcd to The WO, Please check:" + dr["Work_Order"].ToString());
                    }
                    else
                    {
                        cboSlot.Visible = false;
                        label51.Visible = false;
                        Item = dr["Item"].ToString();
                        Slot = dr["Slot"].ToString();
                        LR = dr["LR"].ToString();
                        baseQty = dr["BaseQty"].ToString();
                        if (tempDisQty + Int32.Parse(dr["BalanceQty"].ToString()) > 0)
                        {
                            tempDIDQty = -Int32.Parse(dr["BalanceQty"].ToString());
                        }
                        else
                        {
                            tempDIDQty = tempDisQty;
                        }
                        dispQtyByWO = Int32.Parse(dr["DispatchQty"].ToString()) + tempDIDQty;
                        balanceQtyByWO = dispQtyByWO - Int32.Parse(dr["NeedQty"].ToString());
                        tempDisQty = tempDisQty - tempDIDQty;
                        DataTable dtDispatch = process.InsertDispatch(dr["Work_Order"].ToString(), cboGroupID.SelectedItem.ToString(),
                            cboLine.SelectedItem.ToString(), dr["WoQty"].ToString(), dr["JobPN"].ToString(), txtMachine.Text,
                            txtCompPN.Text, Slot, LR, baseQty, dr["TotalNeedQty"].ToString(), txtDID.Text, txtDIDTotalQty.Text, tempDIDQty,
                            txtVendorCode.Text, txtDateCode.Text, txtLotCode.Text, Parameter.g_userName, transDateTime);
                        if (dtDispatch.Rows.Count == 0)
                        {
                            MessageBox.Show("Insert into QSMS_Dispatch Error,please retry again");
                            return false;
                        }
                        else
                        {
                            if (dtDispatch.Rows[0][0].ToString().ToUpper() == "PASS")
                            {
                                if (firstTimeOfDispatch == 1)
                                {
                                    process.RecordDispatchFDT(dr["Work_Order"].ToString());
                                    firstTimeOfDispatch = firstTimeOfDispatch + 1;
                                }
                            }
                            else
                            {
                                DataTable dtDispatch1 = process.InsertDispatch(dr["Work_Order"].ToString(), cboGroupID.SelectedItem.ToString(),
                                    cboLine.SelectedItem.ToString(), dr["WoQty"].ToString(), dr["JobPN"].ToString(), txtMachine.Text,
                                    txtCompPN.Text, Slot, LR, baseQty, dr["TotalNeedQty"].ToString(), txtDID.Text, txtDIDTotalQty.Text, tempDIDQty, txtVendorCode.Text,
                                    txtDateCode.Text, txtLotCode.Text, Parameter.g_userName, transDateTime);
                                if (dtDispatch1.Rows.Count == 0)
                                {
                                    MessageBox.Show("Insert into QSMS_Dispatch Error,please retry again");
                                    return false;
                                }
                                else
                                {
                                    if (dtDispatch.Rows[0][0].ToString().ToUpper() == "PASS")
                                    {
                                        if (firstTimeOfDispatch == 1)
                                        {
                                            MessageBox.Show("Insert into QSMS_Dispatch Error,please retry again");
                                            return false;
                                        }
                                    }
                                }
                            }
                        }
                        tempWo = dr["Work_Order"].ToString();
                        count += 1;
                    }
                }
            }
            if (count == 0)
            {
                MessageBox.Show("didn't dispatch the  material ,please check");
                return false;
            }
            if(tempBalanceQty>=0 || Int32.Parse(txtDIDRemainQty.Text) == 0)
            {
                RefreshDID_Machine_WO("DID",txtCompPN.Text,txtMachine.Text,txtWO.Text,cboJob.SelectedItem.ToString(),cboLine.SelectedItem.ToString(),woArray);
            }
            ChkWOItemFinished(woArray);
            txtCompPN.Text = "";
            txtDID.Text = "";
            txtDID.Focus();
            return flag;
        }

        private bool ChkWOItemFinished(string woArray)
        {
            string transDateTime= DateTime.Now.ToString("yyyyMMddhhmmss");
            string[] woList = new string[100];
            int n = 0;
            bool flag = true;
            for(int i = 0; i < 100; i++)
            {
                woList[i] = "";
            }
            woArray = woArray.Replace("'", "");
            while (woArray.Length >= 9)
            {
                woList[n] = woArray.Substring(0, 9);
                if (woArray.IndexOf(",") > 0)
                {
                    woArray = woArray.Substring(10);
                }
                else
                {
                    woArray = "";
                }
                n = n + 1;
            }
            n = 0;
            DataTable dt = null;
            while (woList[n] != "")
            {
                dt = process.GetWOByWO(woList[n], "MCC_GetWOByWO");
                if (dt.Rows.Count > 0)
                {
                    process.UpdateDispatchFlagByWO1(woList[n]);
                    flag = false;
                }
                else
                {
                    process.UpdateDispatchFlagByWO2(woList[n],transDateTime);
                    flag = true;
                }
                dt = null;
                n = n + 1;
            }
            return flag;
        }

        private void RefreshDID_Machine_WO(string type, string compPN, string machine, string WO, string jobGroup, string line, string woArray)
        {
            switch (type.ToUpper())
            {
                case "DID":
                    GetDID(machine, jobGroup, WO, line);
                    break;
                case "MACHINE":
                    GetMachine(woArray);
                    break;
                case "WO":
                    GetWoGroup(jobGroup);
                    break;
            }
        }

        private bool ChkDIDDispatchedToWo(string WO, string DID, string machine, string Slot, string LR)
        {
            bool flag = false;
            DataTable dt = process.GetDIDDispatchInfoByDIDAndWO(WO, DID, "MCC_GetDIDDispatchInfoByDIDAndWO");
            if (dt.Rows.Count == 0)
            {
                return true;
            }
            foreach(DataRow dr in dt.Rows)
            {
                if(dr["DeletedFlag"].ToString().ToUpper()=="Y" || dr["Machine"].ToString().ToUpper()==machine.ToUpper() 
                    && dr["Slot"].ToString().ToUpper()==Slot.ToUpper() && dr["LR"].ToString().ToUpper() == LR.ToUpper())
                {
                    flag = true;
                }
                else
                {
                    flag = false;
                    return flag;
                }
            }
            return flag;
        }

        private string GetWoArrayForsp()
        {
            string WoArray = "";
            DataTable dt = process.GetWoArray(txtWO.Text, "MCC_GetWoArray");
            if (dt.Rows.Count > 0)
            {
                foreach(DataRow dr in dt.Rows)
                {
                    WoArray = WoArray + dr["WO"].ToString() + ",";
                }
            }
            for(int i = 0; i < listWoDispatching.Items.Count; i++)
            {
                dt = null;
                listWoDispatching.SelectedIndex = i;
                dt = process.GetWoArray(listWoDispatching.SelectedItem.ToString(), "MCC_GetWoArray");
                if (dt.Rows.Count > 0)
                {
                    foreach(DataRow dr in dt.Rows)
                    {
                        WoArray = WoArray + dr["WO"].ToString() + ",";
                    }
                }
            }
            return WoArray.Substring(0, WoArray.Length - 1);
        }

        private bool ChkErr(ref string msg)
        {
            if (cboLine.SelectedItem == null || cboLine.SelectedItem.ToString().Trim().ToUpper() != txtLine.Text.Trim().ToUpper())
            {
                msg = "Line does not match with GroupID,please check the line";
                return false;
            }
            if (txtWO.Text == "")
            {
                msg = "The WO is error,Please check";
                return false;
            }
            if (cboGroupID.SelectedItem == null || cboGroupID.SelectedItem.ToString() == "")
            {
                msg = "The GroupID is error,Please check";
                return false;
            }
            if (txtMachine.Text == null || txtMachine.Text == "")
            {
                msg = "Machine selected is error, Please check";
                return false;
            }
            if (txtDID.Text == null || txtMachine.Text == "")
            {
                msg = "DID selected  is error,please check";
                return false;
            }
            if (txtDispatchQty.Text == null || txtDispatchQty.Text == "" || txtDispatchQty.Text == "0")
            {
                msg = "Dispatch Qty is error,Please check";
                return false;
            }
            if (ChkWOSeq(txtWO.Text, txtMachine.Text) == false)
            {
                return false;
            }
            if (Int32.Parse(txtDIDTotalQty.Text) > Int32.Parse(txtDIDRemainQty.Text))
            {
                msg = "dispatch Qty can not larger than DID remainQty";
                return false;
            }
            if (txtCompPN.Text == "")
            {
                msg = "CompPN can not be empty,Please check";
                return false;
            }
            if (ChkDIDCompPN(txtDID.Text, txtCompPN.Text) == false)
            {
                msg = "The DID and CompPN doesn't match";
                return false;
            }
            if (ChkNonAVL(txtDID.Text, txtCustomer.Text, txtModel.Text, txtWO.Text, ref msg) == false)
            {
                return false;
            }
            if (ChkAVL(txtCompPN.Text,txtVendorCode.Text,txtCustomer.Text,txtModel.Text) == false)
            {
                msg = "Check AVL failed,please check";
                return false;
            }
            DataTable dt = null;
            dt = process.GetUsedFlagByDID(txtDID.Text, "MCC_GetUsedFlagByDID");
            if (dt.Rows.Count > 0)
            {
                msg = "The DID has been used,please check";
                return false;
            }
            dt = null;
            dt = process.GetWOInfoByGroupIDAndWO(cboGroupID.SelectedItem.ToString(),txtWO.Text, "MCC__GetWOInfoByGroupIDAndWO");
            if (dt.Rows.Count == 0)
            {
                msg = "The Work Order: "+txtWO.Text+" does not belong to the GroupID :" +cboGroupID.SelectedItem.ToString()+" Please check";
                return false;
            }
            else if(dt.Rows[0]["ClosedFlag"].ToString()=="Y")
            {
                msg = "The Work Order: " +txtWO.Text+" has closed, Please check !!";
                return false;
            }
            dt = null;
            dt = process.GetIPQCFlag(txtDID.Text, "MCC_GetIPQCFlag");
            if (dt.Rows.Count > 0)
            {
                if(dt.Rows[0]["IPQCFlag"].ToString()=="" && dt.Rows[0]["IPQCFlag"].ToString() == "N")
                {
                    msg = "IPQC test fail or not test";
                    return false;
                }
            }
            return true;
        }

        private bool ChkAVL(string compPN, string VC, string customer, string model)
        {
            bool flag = true, controlPart=false;
            string AVLCustomer = "", ModelFlag = "";
            DataTable dt = process.GetAVLCustomerByCustomer(customer, "MCC_GetAVLCustomerByCustomer");
            if (dt.Rows.Count > 0)
            {
                AVLCustomer = dt.Rows[0]["AVL_Customer"].ToString();
                ModelFlag = dt.Rows[0]["ModelFlag"].ToString();
            }
            else
            {
                AVLCustomer = "Quanta";
            }
            if (AVLCustomer != "Quanta")
            {
                DataTable dtAVL = null;
                if (ModelFlag == "Y")
                {
                     dtAVL= process.GetAVLInfo(AVLCustomer,model,compPN,VC,"MCC_GetAVLInfo");
                }
                else
                {
                    dtAVL= process.GetAVLInfo(AVLCustomer, "", compPN, VC, "MCC_GetAVLInfo");
                }
                if (dtAVL.Rows.Count == 0)
                {
                    flag = false;
                }
            }
            dt = null;
            dt = process.GetVCByModelAndCompPN(model,compPN,"MCC_GetVCByModelAndCompPN");
            if (dt.Rows.Count == 0)
            {
                return true;
            }
            else
            {
                if (dt.Rows[0]["VendorCode"].ToString() == VC)
                {
                    controlPart = true;
                }
            }
            if (controlPart == true)
            {
                flag = true;
            }
            else
            {
                flag = false;
            }
            return flag;
        }

        private bool ChkDIDCompPN(string DID, string compPN)
        {
            if(process.ChkDIDCompPN(DID,compPN, "MCC_ChkDIDCompPN").Rows.Count == 0)
            {
                return false;
            }
            return true;
        }

        private bool ChkNonAVL(string DID,string customer,string model,string wo,ref string msg)
        {
            bool flag = true;
            DataTable dt = process.GetCodeByDID(txtDID.Text, "MCC_GetCodeByDID");
            if (dt.Rows.Count == 0)
            {
                msg = "Can not find the DID,Please check";
                flag = false;
                return flag;
            }
            dt = null;
            dt = process.GetNonAVLCode(customer, dt.Rows[0]["CompPN"].ToString(), model, dt.Rows[0]["VendorCode"].ToString(), dt.Rows[0]["DateCode"].ToString(), dt.Rows[0]["LotCode"].ToString(),wo, "MCC_GetNonAVCodeByDID");
            if (dt.Rows.Count == 0)
            {
                return flag;
            }
            else
            {
                flag = false;
            }
            if (Parameter.Check_NonAVL != "Y")
            {
                flag = true;
            }
            if (flag == false)
            {
                msg = "Check NonAVL failed";
            }
            return flag;
        }

        private void GetLine()
        {
            DataTable dt = process.GetLine("MCC_GetLine");
            cboLine.Items.Clear();
            foreach(DataRow dr in dt.Rows)
            {
                cboLine.Items.Add(dr["Line"].ToString());
            }
        }
        
        private void GetCompDispInfo(string WO, string jobPN, string machine, string compPN)
        {
            txtConsumedQty.Text = "";
            txtCompBaseQty.Text = "";
            txtNeedQty.Text = "";
            string woArray = GetWOArray(WO);
            DataTable dt = process.GetQty(txtGroup.Text,woArray,jobPN,machine,compPN,"MCC_GetQty");
            if (dt.Rows.Count > 0)
            {
                txtCompBaseQty.Text = dt.Rows[0]["BaseQty"].ToString();
                txtConsumedQty.Text = dt.Rows[0]["DispatchQty"].ToString();
                txtBalanceQty.Text = dt.Rows[0]["BalanceQty"].ToString();
                txtNeedQty.Text = dt.Rows[0]["NeedQty"].ToString();
                txtTotalQty.Text = dt.Rows[0]["TotalNeedQty"].ToString();
            }
            if (txtNeedQty.Text == "")
            {
                return;
            }
            if (Int32.Parse(txtDIDRemainQty.Text) > -Int32.Parse(txtBalanceQty.Text))
            {
                txtDispatchQty.Text = "-" + txtBalanceQty.Text;
            }
            else
            {
                txtDispatchQty.Text = txtDIDRemainQty.Text;
            }
            GetDispatchQtyBySlot(WO,machine,compPN);
        }

        private void GetDispatchQtyBySlot(string WO, string machine, string compPN)
        {
            string woArray = GetWOArray(WO);
            DataTable dt = process.GetDispatch1(WO, woArray, machine, compPN, "MCC_GetDispatch1");
            DGSlot.DataSource = dt;
        }

        private void GetDIDInfo(string DID, string WO)
        {
            txtConsumedQty.Text = "";
            txtCompBaseQty.Text = "";
            txtNeedQty.Text = "";
            txtDispatchQty.Text = "";
            txtCompPN.Text = "";
            txtVendorCode.Text = "";
            txtDIDDateTime.Text = "";
            txtDateCode.Text = "";
            txtLotCode.Text = "";
            txtDIDRemainQty.Text = "";
            DataTable dt = process.GetDIDInfo(DID, "MCC_GetDIDInfo1");
            if (dt.Rows.Count > 0)
            {
                txtCompPN.Text = dt.Rows[0]["CompPN"].ToString();
                txtVendorCode.Text= dt.Rows[0]["VendorCode"].ToString();
                txtDateCode.Text= dt.Rows[0]["DateCode"].ToString();
                txtLotCode.Text= dt.Rows[0]["LotCode"].ToString();
                txtRackID.Text= dt.Rows[0]["DIDLoc"].ToString();
                txtDIDTotalQty.Text= dt.Rows[0]["Qty"].ToString();
                txtDIDRemainQty.Text= dt.Rows[0]["RemainQty"].ToString();
                txtDIDDateTime.Text= dt.Rows[0]["TransDateTime"].ToString();
            }
            if (dt.Rows[0]["UsedFlag"].ToString().ToUpper() == "Y")
            {
                txtDispatchQty.Enabled = false;
                txtDispatchQty.Text = "0";
            }
            else
            {
                txtDispatchQty.Enabled = true;
            }
        }

        private void GetDIDFromSourceBU(string dID)
        {
            try
            {
                process.GetDIDFromSourceBU(dID);
            }
            catch (Exception e)
            {
                MessageBox.Show("Exception:" + e.Message);
            }
        }

        private void cboGroupID_TextChanged(object sender, EventArgs e)
        {
            if (cboGroupID.Text.Length > 12)
            {
                cboGroupID_Click(sender, e);
            }
        }

        private bool ChkDIDBelongMachine(string WO, string machine, string DID, string sapWOGroup)
        {
            string message = "";
            DataTable dt = process.GetUseFlagByDID(DID, "MCC_GetUseFlagByDID");
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["UsedFlag"].ToString() == "Y")
                {
                    message = "";
                    DataTable dt1 = process.GetDispatchInfo(DID, "MCC_GetDispatchInfo");
                    if (dt1.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt1.Rows)
                        {
                            message = message + "WO:" + dr["Work_Order"].ToString() + " Machine:" + dr["Machine"].ToString() + " Slot:" + dr["Slot"].ToString() +"\r\n";
                        }
                        MessageBox.Show("The DID has been used at: \r\n"+message+ " ===PLease check=== !");
                    }
                    return false;
                }
                else
                {
                    message = "";
                    DataTable dt1 = process.GetDispatchInfo(DID, "MCC_GetDispatchInfo");
                    if (dt1.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dt1.Rows)
                        {
                            if (dr["Machine"].ToString().Trim().Substring(0, 2) != machine.Substring(0, 2))
                            {
                                message = message + "WO:" + dr["Work_Order"].ToString() + " Machine:" + dr["Machine"].ToString() + " Slot:" + dr["Slot"].ToString() + "\r\n";
                            }
                        }
                        if (message.Length > 0)
                        {
                            MessageBox.Show("DID can not be dispatched to different line and side,the DID has been dispatched to:\r\n" + message + " ===PLease check=== !");
                            return false;
                        }
                    }
                }
            }
            else
            {
                DataTable dt1 = process.GetDIDLog(DID, "MCC_GetDIDLog");
                if (dt1.Rows.Count > 0)
                {
                    MessageBox.Show("This DID had been deleted !");
                }
                else
                {
                    MessageBox.Show("找不到DID，请检查");
                }
                return false;
            }
            string woArray = GetWOArray(WO);
            if (process.GetDIDQtyByDIDMachine(sapWOGroup, woArray, machine, DID, "MCC_GetDIDQtyByDIDMachine").Rows.Count > 0)
            {
                return true;
            }
            else
            {
                MessageBox.Show("the DID does not belong to the machine :" + machine + ".  Please Check.");
                return false;
            }
        }
        
    }
}
