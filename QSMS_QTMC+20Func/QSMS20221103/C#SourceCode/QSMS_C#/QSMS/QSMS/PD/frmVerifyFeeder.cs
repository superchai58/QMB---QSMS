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
    public partial class frmVerifyFeeder : Form
    {
        DbLibrary.PD.VerifyFeederProcess process = new DbLibrary.PD.VerifyFeederProcess();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();

        public frmVerifyFeeder()
        {
            InitializeComponent();
        }
        private void frmVerifyFeeder_Load(object sender, EventArgs e)
        {
            dtpSDate.Value = DateTime.Now;
        }

        private void cboGroupID_Click(object sender, EventArgs e)
        {
            if (cboGroupID.SelectedItem!=null)
            {
                cboWO.Items.Clear();
                cboWO.Text = "";
                listNotDispatch.Items.Clear();
                listClosed.Items.Clear();
                DataTable dt=process.GetWoByGroupID(cboGroupID.Text.ToString());
                if (dt.Rows.Count > 0)
                {
                    foreach(DataRow dr in dt.Rows)
                    {
                        if(dr["Sap1Flag"].ToString().Trim()=="Y" && dr["ClosedFlag"].ToString().Trim() == "N")
                        {
                            cboWO.Items.Add(dr["Work_Order"].ToString());
                        }
                        if(dr["Sap1Flag"].ToString().Trim() == "N")
                        {
                            listNotDispatch.Items.Add(dr["Work_Order"].ToString());
                        }
                        if(dr["ClosedFlag"].ToString().Trim() == "Y")
                        {
                            listClosed.Items.Add(dr["Work_Order"].ToString());
                        }
                    }
                }
            }
        }

        private void cboGroupID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                cboGroupID_Click(sender, e);
            }
        }

        private void cboMachine_Click(object sender, EventArgs e)
        {
            if (cboMachine.Text.ToString()!=null && cboWO.Text.ToString()!="")
            {
                DataTable dt= process.GetJobByMachine(cboMachine.Text.ToString(), cboWO.Text.ToString());
                if (dt.Rows.Count > 0)
                {
                    cboJobPN.Items.Clear();
                    cboJobPN.Text = "";
                    foreach(DataRow dr in dt.Rows)
                    {
                        cboJobPN.Items.Add(dr["JobPN"].ToString());
                    }
                }
            }
            else
            {
                lblMessage.Text="Please select Machine or WO!";
                return;
            }
        }

        private void cboMachine_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                cboMachine_Click(sender, e);
            }
        }

        private void cmdQuery_Click(object sender, EventArgs e)
        {
            string SDate = "",line="";
            if(txtLine.Text.ToString() == "")
            {
                MessageBox.Show("Please input line");
                return;
            }
            line = txtLine.Text.ToString();
            SDate = dtpSDate.Value.ToString("yyyyMMdd");
            if (!string.IsNullOrEmpty(SDate))
            {
                DataTable dt = process.GetGroupIDByLine(SDate, line);
                if (dt.Rows.Count > 0)
                {
                    cboGroupID.Items.Clear();
                    cboGroupID.Text = "";
                    foreach (DataRow dr in dt.Rows)
                    {
                        cboGroupID.Items.Add(dr["GroupID"].ToString());
                    }
                }
            }
        }

        private void cboWO_Click(object sender, EventArgs e)
        {
            if (cboWO.Text.ToString() != null && cboWO.Text.ToString() != "")
            {
                string rev = "";
                DataTable dt = process.GetMachineByWo(cboWO.Text.ToString(), "", "GetRev");
                if (dt.Rows.Count > 0)
                {
                    rev = dt.Rows[0]["Mb_Rev"].ToString();
                    cboVersion.Text = rev;
                }
                cboMachine.Items.Clear();
                cboMachine.Text = "";
                DataTable dt1 = process.GetMachineByWo(cboWO.Text.ToString(), rev, "");
                if (dt1.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt1.Rows)
                    {
                        cboMachine.Items.Add(dr["Machine"].ToString());
                    }
                }
            }
        }

        private void cboWO_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                cboWO_Click(sender, e);
            }
        }

        private void txtSlot_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                string tempCompPN = "",dateTime="";
                if(cboMachine.Text.ToString()=="" || cboJobPN.Text.ToString()=="" || cboVersion.Text.ToString()=="" || txtSlot.Text.ToString()=="" || txtLR.Text.ToString() == "")
                {
                    lblMessage.Text = "Machine or JobPN or Version or Slot or LR can not be null!";
                    return;
                }
                DataTable dt = process.GetCompPN(cboMachine.Text.ToString(),cboJobPN.Text.ToString(),cboVersion.Text.ToString(),txtSlot.Text.ToString(),txtLR.Text.ToString());
                if (dt.Rows.Count == 0)
                {
                    txtSlot.Text = "";
                    txtSlot.Focus();
                    pubFunction.Sound("ERROR");
                    lblMessage.Text = "Can not find the CompPN by the Slot,Please check!";
                    return;
                }
                tempCompPN = dt.Rows[0]["CompPN"].ToString();
                if (tempCompPN.ToUpper() == txtCompPN.Text.ToString().ToUpper())
                {
                    process.UpdateQSMS_Feeder(txtSlot.Text.ToString(), txtFeeder.Text.ToString());
                    DataTable dt1 = process.ChkMachineVerifyFinished(cboMachine.Text.ToString(), cboJobPN.Text.ToString(), cboVersion.Text.ToString());
                    if (dt1.Rows.Count == 0)
                    {
                        dateTime = DateTime.Now.ToString("yyyyMMddhhmmss");
                        DataTable dt2 = process.UpdateQSMS_Verify(cboMachine.Text.ToString(), cboJobPN.Text.ToString(), cboVersion.Text.ToString(), tempCompPN, "", "", "1");
                        if (dt2.Rows.Count > 0)
                        {
                            foreach (DataRow dr in dt2.Rows)
                            {
                                if (process.UpdateQSMS_Verify(cboMachine.Text.ToString(), cboJobPN.Text.ToString(), cboVersion.Text.ToString(), tempCompPN, "", "", "2").Rows.Count == 0)
                                {
                                    process.UpdateQSMS_Verify(cboMachine.Text.ToString(), cboJobPN.Text.ToString(), cboVersion.Text.ToString(), tempCompPN, "", dateTime, "");
                                }
                            }
                        }
                    }
                    txtFeeder.Enabled = true;
                    txtFeeder.Text = "";
                    txtSlot.Text = "";
                    txtFeeder.Focus();
                }
                else
                {
                    txtSlot.Text = "";
                    txtSlot.Focus();
                    pubFunction.Sound("ERROR");
                    lblMessage.Text = "CompPN is different with the Feeder ,Please check:" + tempCompPN;
                    return;
                }
            }
        }

        private void txtFeeder_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                if (txtFeeder.Text != null && txtFeeder.Text!="")
                {
                    DataTable dt = process.GetCodesByFeeder(txtFeeder.Text.ToString());
                    if (dt.Rows.Count > 0)
                    {

                    }
                }
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            txtFeeder.Enabled = true;
            txtSlot.Enabled = true;
            txtFeeder.Text = "";
            txtSlot.Text = "";
            txtFeeder.Focus();
        }

        private void frmVerifyFeeder_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmVerifyFeeder");
        }
    }
}
