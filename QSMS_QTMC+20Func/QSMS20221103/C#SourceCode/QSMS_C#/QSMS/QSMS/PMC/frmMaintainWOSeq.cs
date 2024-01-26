using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.PMC
{
    public partial class frmMaintainWOSeq : Form
    {
        public frmMaintainWOSeq()
        {
            InitializeComponent();
        }
        string BeginDate, EndDate, Wo_TransDate, strLine;
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.PMC.PMCProcess PMC = new DbLibrary.PMC.PMCProcess();
        private void frmMaintainWOSeq_Load(object sender, EventArgs e)
        {
            DataTable dt = PMC.GetLinefromMachine();
            if(dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    CboLine.Items.Add(dt.Rows[i]["Line"].ToString().Trim());
                }
            }
        }

        private void CboGroupID_DropDown(object sender, EventArgs e)
        {

        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            //if (CboLine.Text.Trim() == "")
            if (string.IsNullOrEmpty(CboLine.Text.Trim()))
            {
                MessageBox.Show("Please input line");
                return;
            }
            else
            {
                BeginDate = string.Format("{0:yyyyMMdd}", dtpSDate.Value) + "000000";
                EndDate = string.Format("{0:yyyyMMdd}", dtpEDate.Value) + "240000";

                DataTable dt = PMC.GetListWO(Parameter.BU,BeginDate,EndDate,CboLine.Text.Trim());
                if(dt.Rows.Count >0)
                {
                    lstWO_LIST.Items.Clear();
                    for (int i = 0; i < dt.Rows.Count; i ++ )
                    {
                        if (GetGroupID(dt.Rows[i]["WO"].ToString().Trim()) == "")
                        {
                            lstWO_LIST.Items.Add(dt.Rows[i]["WO"].ToString());
                        }
                    }
                }
                strLine = CboLine.Text.Trim();
            }
        }

        private void btnQueryID_Click(object sender, EventArgs e)
        {
            DataTable dt;
            BeginDate = string.Format("{0:yyyyMMdd}", dtpSDate.Value);
            EndDate = string.Format("{0:yyyyMMdd}", dtpEDate.Value);
            if (rbtnRelease.Checked == true)
            {
                dt = PMC.GetGroupIDbyRelease(CboLine.Text.Trim(), BeginDate, EndDate);
            }
            else
            {
                dt = PMC.GetGroupID(CboLine.Text.Trim(), BeginDate, EndDate);
            }
            CboGroupID.Items.Clear();
            if(dt.Rows.Count >0)
            {
                for (int i = 0; i < dt.Rows.Count;i++ )
                {
                    CboGroupID.Items.Add(dt.Rows[i]["GroupID"].ToString());
                }
            }
        }

        private void CboGroupID_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt = PMC.GetGroupWO(CboGroupID.Text.Trim());
            if(dt.Rows.Count >0)
            {
                for (int i = 0; i < dt.Rows.Count; i++ )
                {
                    CboWo.Items.Add(dt.Rows[i]["Work_Order"].ToString());
                }
            }
        }

        private void CboWo_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable dt;
            txtWO.Text = CboWo.Text.Trim();
            dt = PMC.GetWoInfo(txtWO.Text);
            if(dt.Rows.Count > 0)
            {
                txtMBPN.Text = dt.Rows[0]["PN"].ToString().Trim();
                txtWOQty.Text = dt.Rows[0]["Qty"].ToString().Trim();
                Wo_TransDate = dt.Rows[0]["Trans_Date"].ToString().Trim().Substring(0,8);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DataTable dt= null;
            if (ChkDelete(CboWo.Text) == true)
            {
                dt = PMC.DeleteWOByGroup(CboGroupID.Text.Trim(),CboWo.Text.Trim(),Parameter.g_userName);
                if (dt.Rows.Count > 0)
                {
                    MessageBox.Show(dt.Rows[0]["Description"].ToString());
                }
                else
                {
                    MessageBox.Show("can not find the Wo :" + CboWo.Text);
                }
            }
        }
        private bool ChkDelete(string WO)
        {
            DataTable ds=null;
            ds = PMC.CheckWOGroup(WO);
            if(ds.Rows.Count >0)
            {
                if (ds.Rows[0]["DispatchFlag"].ToString().Trim() == "Y")
                {
                    MessageBox.Show("can not delete!!!!The work Order has been dispatched:" + WO);
                    return false;
                }
            }
            ds = PMC.CheckDispatch(WO);
            if(ds.Rows[0]["Qty"].ToString().Trim() == "0")
            {
                MessageBox.Show("can not delete!!!!! The word order is dispatching:" + WO);
                return false;
            }
            return true;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            int Pointer;
            if(lstWO_LIST.Items.Count <= 0 || lstWO_LIST.SelectedIndex < 0)
            {
                return;
            }
            Pointer = lstWO_LIST.SelectedIndex;
            lstWO_SELECT.Items.Add(lstWO_LIST.SelectedItem.ToString());
            lstWO_LIST.Items.RemoveAt(Pointer);
            if (lstWO_LIST.Items.Count != Pointer)
            {
                lstWO_LIST.SelectedIndex = Pointer;
            }
        }

        private void btnADDALL_Click(object sender, EventArgs e)
        {
            if(lstWO_LIST.Items.Count <= 0)
            {
                return;
            }
            for (int i = lstWO_LIST.Items.Count; i > 0; i--)
            {
                lstWO_LIST.SelectedIndex = 0;
                lstWO_SELECT.Items.Add(lstWO_LIST.SelectedItem.ToString());
                lstWO_LIST.Items.RemoveAt(0);
            }
        }

        private void btnDEL_Click(object sender, EventArgs e)
        {
            int Pointer;
            if (lstWO_SELECT.Items.Count <= 0 || lstWO_SELECT.SelectedIndex < 0)
            {
                return;
            }
            Pointer = lstWO_SELECT.SelectedIndex;
            lstWO_LIST.Items.Add(lstWO_SELECT.SelectedItem.ToString());
            lstWO_SELECT.Items.RemoveAt(Pointer);
            if (lstWO_SELECT.Items.Count != Pointer)
            {
                lstWO_SELECT.SelectedIndex = Pointer;
            }
        }

        private void btnDELALL_Click(object sender, EventArgs e)
        {
            if (lstWO_SELECT.Items.Count <= 0)
            {
                return;
            }
            for (int i = lstWO_SELECT.Items.Count; i > 0; i--)
            {
                lstWO_SELECT.SelectedIndex = 0;
                lstWO_LIST.Items.Add(lstWO_SELECT.SelectedItem.ToString());
                lstWO_SELECT.Items.RemoveAt(0);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string WOList = "", tempwo, TempGroupID, TempGroupDatetime, MBFlag;
            DataTable rs;DataSet ds;
            if(lstWO_SELECT.Items.Count <= 0)
            {
                return;
            }
            for (int i = 0; i < lstWO_SELECT.Items.Count; i++ )
            {
                lstWO_SELECT.SelectedIndex = i;
                tempwo = lstWO_SELECT.Text.Trim();
                WOList = WOList + tempwo + ",";
                rs = PMC.DblChkLine(tempwo,CboLine.Text.Trim());
                if(rs.Rows.Count <= 0)
                {
                    MessageBox.Show("Line doesn't match the wo,Please check");
                    return;
                }                
            }
            TempGroupID = GenGroupID();
            TempGroupDatetime = string.Format("{0:yyyyMMddHHmmss}", DateTime.Now);
            rs = PMC.CHKMaintainWO(WOList,CboLine.Text.Trim(),TempGroupID);
            if (rs.Rows.Count > 0)
            {
                if (rs.Rows[0]["Result"].ToString().Trim().ToUpper() == "FAIL")
                {
                    MessageBox.Show(rs.Rows[0]["Description"].ToString());
                    return;
                }
            }
            if(Parameter.CheckPNGroup == "Y")
            {
                WOList = WOList.Substring(0, WOList.Length - 1);
                ds = PMC.ChkPNGroup(WOList);
                if (ds.Tables[0].Rows[0]["Result"].ToString().Trim() != "0")
                {
                    MessageBox.Show(ds.Tables[0].Rows[0]["Description"].ToString().Trim() + ds.Tables[1].Rows[0]["ErrInfo"].ToString().Trim());
                    return;
                }
            }
            for (int i = 0; i < lstWO_SELECT.Items.Count; i++)
            {
                lstWO_SELECT.SelectedIndex = i;
                tempwo = lstWO_SELECT.Text.Trim();
                WOList = WOList + tempwo + ",";
                rs = PMC.CheckWOGroupID(tempwo, TempGroupID);
                if (rs.Rows.Count <= 0)
                {
                    if (rs.Rows[0]["Item"].ToString().Trim().ToUpper() == "N")
                    {
                        MessageBox.Show("Other work order which is in the same PCB has already in the system,GroupID is: " + rs.Rows[0]["GroupID"].ToString().Trim());
                        return;
                    }
                }
                if (pubFunction.ConfigListGetValue("ChkWOGroupID") == "Y")
                {
                    rs = PMC.QSMS_ChkWOGroupID(tempwo, TempGroupID);
                    if (rs.Rows[0]["Result"].ToString().Trim() == "1")
                    {
                        if (MessageBox.Show(rs.Rows[0]["Err"].ToString().Trim() + "  Do you want to continue?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                    }
                }
                rs = PMC.GetWoInfo(tempwo);
                if (rs.Rows.Count > 0)
                {
                    txtMBPN.Text = rs.Rows[0]["PN"].ToString().Trim();
                    txtWOQty.Text = rs.Rows[0]["Qty"].ToString().Trim();
                    Wo_TransDate = rs.Rows[0]["Trans_Date"].ToString().Trim().Substring(0, 8);
                }
                rs = PMC.ChkWOGroup_His(tempwo);
                if (rs.Rows.Count < 1)
                {
                    rs = PMC.ChkMBWo(tempwo);
                    if (rs.Rows.Count > 0)
                    {
                        MBFlag = "1";
                    }
                    else
                    {
                        MBFlag = "0";
                    }
                    PMC.Insert_QSMSWoGroup(tempwo, txtMBPN.Text.Trim(), MBFlag, strLine, (i + 1).ToString(), TempGroupID, Wo_TransDate, TempGroupDatetime, Parameter.g_userName);
                    rs = PMC.XL_CheckWOGroupID(tempwo,TempGroupID);
                    if (rs.Rows[0]["Item"].ToString() == "N")
                    {
                        MessageBox.Show("SP:XL_CheckWOGroupID Warnning");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("The Work Order already In DB, GroupID is:" + rs.Rows[0]["GroupID"].ToString());
                }
            }
            lstWO_SELECT.Items.Clear();
        }
        private string GenGroupID()
        {
            string  TransDate,TempGroupHead,seq;
            TransDate = string.Format("{0:yyyyMMdd}", DateTime.Now);
            if(pubFunction.ConfigListGetValue("NewGroupIDRule") =="Y")
            {
                TempGroupHead = CboLine.Text.Trim().ToUpper()+ TransDate.Substring(2,2);
            }
            else
            {
                TempGroupHead = CboLine.Text.Trim().ToUpper()+TransDate;
            }
            DataTable rs = PMC.GenGroupID(TempGroupHead);
            if(rs.Rows.Count >0)
            {
                seq =(int.Parse(rs.Rows[0]["GroupID"].ToString().Trim().Substring(rs.Rows[0]["GroupID"].ToString().Trim().Length -4))+1).ToString(); 
                if(seq.Length >5)
                {
                    seq = seq.Substring(seq.Length-3) + "1";
                }
                return TempGroupHead + seq;
            }
            else
            {
                return TempGroupHead + "0001";
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            lstWO_LIST.Items.Clear();
            lstWO_SELECT.Items.Clear();
            CboLine.Items.Clear();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            string TempGroupDatetime, Seq_No, tempwo;
            DataTable rs;
            TempGroupDatetime =string.Format("{0:yyyyMMddHHmmss}", DateTime.Now);
            rs = PMC.GetMaxSeq(CboGroupID.Text.Trim());
            Seq_No = rs.Rows[0]["Max"].ToString().Trim();
            rs = PMC.QSMS_ChkGroupID(CboGroupID.Text.Trim());
            if (rs.Rows[0]["Item"].ToString().Trim().ToUpper() == "N")
            {
                MessageBox.Show("The GroupID is over 3 week,can not insert");
                return;
            }
            for (int i = 0; i < lstWO_SELECT.Items.Count;i++ )
            {
                lstWO_SELECT.SelectedIndex = i;
                tempwo = lstWO_SELECT.Text;
                rs = PMC.CheckWOGroupID(tempwo,CboGroupID.Text.Trim());
                if (rs.Rows[0]["Item"].ToString().Trim().ToUpper() == "N")
                {
                    MessageBox.Show("Other work order which is in the same PCB has already in the system,GroupID is:" + rs.Rows[0]["GroupID"].ToString().Trim());
                    return;
                }
                rs = PMC.CHKMaintainWO(tempwo,CboLine.Text.Trim(),CboGroupID.Text.Trim());
                if (rs.Rows[0]["Result"].ToString().Trim().ToUpper() == "FAIL")
                {
                    MessageBox.Show(rs.Rows[0]["Description"].ToString().Trim());
                    return;
                }
                if (pubFunction.ConfigListGetValue("ChkWOGroupID") == "Y")
                {
                    rs = PMC.QSMS_ChkWOGroupID(tempwo, CboGroupID.Text.Trim());
                    if (MessageBox.Show(rs.Rows[0]["Err"].ToString().Trim() + "  Do you want to continue?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        return;
                    }
                }
                rs = PMC.GetWoInfo(tempwo);
                if (rs.Rows.Count > 0)
                {
                    txtMBPN.Text = rs.Rows[0]["PN"].ToString().Trim();
                    txtWOQty.Text = rs.Rows[0]["Qty"].ToString().Trim();
                    Wo_TransDate = rs.Rows[0]["Trans_Date"].ToString().Trim().Substring(0, 8);
                }
                Seq_No = (int.Parse(Seq_No) + 1).ToString();
                rs = PMC.CheckWOGroup(tempwo);
                if (rs.Rows.Count > 0)
                {
                    PMC.Update_QSMSWoGroup(tempwo, Seq_No);
                }
                else
                {
                    PMC.Insert_QSMSWoGroup(tempwo,"","",CboLine.Text.Trim(),Seq_No,CboGroupID.Text.Trim(),Wo_TransDate,TempGroupDatetime,Parameter.g_userName);
                }
                rs = PMC.XL_CheckWOGroupID(tempwo,CboGroupID.Text.Trim());
                if (rs.Rows[0]["Item"].ToString().ToUpper().Trim() == "N")
                {
                    MessageBox.Show("Run SP Fail:XL_CheckWOGroupID");
                    return;
                }
            }
            MessageBox.Show("Update Ok");
        }

        private void lstWO_LIST_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable rs = PMC.GetWoInfo(lstWO_LIST.Text);
            if (rs.Rows.Count > 0)
            {
                txtMBPN.Text = rs.Rows[0]["PN"].ToString().Trim();
                txtWOQty.Text = rs.Rows[0]["Qty"].ToString().Trim();
                Wo_TransDate = rs.Rows[0]["Trans_Date"].ToString().Trim().Substring(0, 8);
            }
        }

        private string GetGroupID(string WO)
        {
            DataTable ds;
            ds = PMC.Get_GroupID(WO);
            if (ds.Rows.Count > 0)
            {
                return ds.Rows[0]["GroupID"].ToString().Trim();
            }
            else
            {
                return "";
            }
        }

        private void frmMaintainWOSeq_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmMaintainWOSeq");
        }
    }
}
