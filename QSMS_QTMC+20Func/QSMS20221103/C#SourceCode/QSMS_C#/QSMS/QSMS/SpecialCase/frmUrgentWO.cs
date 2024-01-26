using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.SpecialCase
{
    public partial class frmUrgentWO : Form
    {
        public frmUrgentWO()
        {
            InitializeComponent();
        }
        DbLibrary.SpecialCase.UrgentWO UrgentWO = new DbLibrary.SpecialCase.UrgentWO();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();

        private void btnQuery_Click(object sender, EventArgs e)
        {
            if (txtWO.Text.ToString() != "")
            {
                DataSet ds = new DataSet();
                ds = UrgentWO.QueryWOSeq(txtWO.Text.ToString());

                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    if (ds.Tables[i].Rows.Count > 0)
                    {
                        if (i == 0)
                        {
                            dgWOPlanSeq.DataSource = ds.Tables[i];
                            continue;
                        }
                        if (i == 1)
                        {
                            dgCurWoSeq.DataSource = ds.Tables[i];
                            continue;
                        }
                        if (i == 2)
                        {
                            dgWoInputPlan.DataSource = ds.Tables[i];
                            continue;
                        }
                        if (i == 3)
                        {
                            dgQSMS_WO_XL.DataSource = ds.Tables[i];
                            continue;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("工单不能为空!");
                txtWO.Focus();
                return;
            }
        }

        private void txtWO_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && txtWO.Text != "")
            {
                btnQuery_Click(sender, e);
            }
        }

        private void frmUrgentWO_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmUrgentWO");
        }

        private void btnUrgent_Click(object sender, EventArgs e)
        {
            //if (txtWO.Text.ToString() == "")
            if (string.IsNullOrEmpty(txtWO.Text.Trim()))
            {
                MessageBox.Show("工单不能为空!");
                txtWO.Focus();
                return;
            }
            DataTable dt = new DataTable();
            //dt = UrgentWO.CheckUrgentWO(txtWO.Text.ToString());
            dt = UrgentWO.CheckUrgentWO(txtWO.Text.Trim());
            if (dt.Rows[0]["Result"].ToString() == "0")
            {
                DialogResult dr = MessageBox.Show("Do you really want to insert WO:" + txtWO.Text.ToString().Trim() + ",Date:" + dt.Rows[0]["Workdate"].ToString() + ",Shift:" + dt.Rows[0]["Shift"].ToString() + " ?", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dr == DialogResult.Yes)
                {
                    UrgentWO.InsertUrgentWO(Parameter.g_userName, txtWO.Text.ToString().Trim(), dt.Rows[0]["Workdate"].ToString(), dt.Rows[0]["Shift"].ToString());
                }
            }
            else if (dt.Rows[0]["Result"].ToString() == "1")
            {
                MessageBox.Show("PMC did not upload this WO information of Date= " + dt.Rows[0]["Workdate"].ToString() + " and Shit= " + dt.Rows[0]["Shift"].ToString() + "，please check it!");
                return;
            }
            else if (dt.Rows[0]["Result"].ToString() == "2")
            {
                MessageBox.Show(dt.Rows[0]["Prompt"].ToString());
                return;
            }
        }
    }
}
