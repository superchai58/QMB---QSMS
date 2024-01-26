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
    public partial class frmGenXLMD : Form
    {
        public DataTable dt = new DataTable();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.SpecialCase.GenXLMaterialDemandProcess process = new DbLibrary.SpecialCase.GenXLMaterialDemandProcess();
        public frmGenXLMD()
        {
            InitializeComponent();
        }
        private void btnGenXLMD_Click(object sender, EventArgs e)
        {
            txtmsg.Text = "";
            if (cobFac.Text == "")
            {
                txtmsg.Text = "Please select the Factory!";
                return;
            }
            btnGenXLMD.Enabled = false;
            //if (Parameter.BU == "NB4")
            //{
            //    if (cobType.Text == "")
            //    {
            //        txtmsg.Text = "Please select the Type!";
            //        return;
            //    }
            //    dt = process.XL_Job_12H_NB4(Parameter.UID, cobFac.Text.ToString(), cobType.Text.ToString());
            //    if (dt.Rows.Count > 0)
            //    {
            //        if (dt.Rows[0]["Result"].ToString() != "OK")
            //        {
            //            MessageBox.Show(dt.Rows[0]["Result"].ToString(), "Error Tips");
            //        }
            //        else
            //        {
            //            MessageBox.Show("生成需求成功", "Tips");
            //        }
            //    }
            //}
            if (Parameter.BU.ToUpper() == "NB4")
            {
                if (string.IsNullOrEmpty(cobType.Text.Trim()))
                {
                    txtmsg.Text = "Please select the Type!";
                    return;
                }
            }

            if (Parameter.BU.ToUpper() == "PO")
            {
                dt = process.XL_JOB_PO(Parameter.UID, cobFac.Text.Trim());
            }
            else if (Parameter.BU.ToUpper() == "NB4")
            {
                dt = process.XL_Job_12H_NB4(Parameter.UID, cobFac.Text.Trim(), cobType.Text.Trim());
            }
            else
            {
                dt = process.XL_JOB_Others(Parameter.UID, cobFac.Text.Trim());
            }
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["RESULT"].ToString().ToUpper() != "OK")
                {
                    MessageBox.Show(dt.Rows[0]["MSG"].ToString(),"ErrorMessage Tips",MessageBoxButtons.YesNo,MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("生成需求成功", "MessageTips", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                }
            }
            btnGenXLMD.Enabled = true;
        }

        private void frmGenXLMD_Load(object sender, EventArgs e)
        {
            label1.Visible = false;
            cobType.Visible = false;
            txtmsg.Text = "注意:\r\n  可以再次计算XL需求的时间是第一次XL跑过1H~5H之间\r\n例如:\r\n  XL时间为7:40 那么可以再次计算需求的时间段为8:40~12:40,\r\n如果超过这个时间点将不允许需手动跑,将在由系统自动计算.";
            dt = process.GetSite();
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    cobFac.Items.Add(dr["Factory"].ToString());
                }
            }
            dt = null;
            if (Parameter.BU == "NB4")
            {
                label1.Visible = true;
                cobType.Visible = true;
                dt = process.GetXL_Type();
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        cobType.Items.Add(dr["XL_Type"].ToString());
                    }
                }
            }
        }

        private void frmGenXLMD_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("GenXLMaterialDemand");
        }
    }
}
