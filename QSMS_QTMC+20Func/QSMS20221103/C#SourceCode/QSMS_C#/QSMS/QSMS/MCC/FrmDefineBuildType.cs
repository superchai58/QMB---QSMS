using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using QuantaSDK.Excel;

namespace QSMS.QSMS.MCC
{
    public partial class FrmDefineBuildType : Form
    {
        DbLibrary.MCC.MCCProcess mccProcess = new DbLibrary.MCC.MCCProcess();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        public FrmDefineBuildType()
        {
            InitializeComponent();
        }

        private void FrmDefineBuildType_Load(object sender, EventArgs e)
        {
            GetLine();
            CboBuildType.Items.Add("1");
            CboBuildType.Items.Add("2");
            CboBuildType.Items.Add("3");
            CboBuildType.Items.Add("4");
            GetData();
            CboStation.Items.Add("SP");
            CboStation.Items.Add("SP2");
        }

        private void GetData()
        {
            DataGDetail.DataSource = mccProcess.GetDataGDetail();
            DataWOMulti.DataSource = mccProcess.GetDataWOMulti();
        }

        private void GetLine()
        {
            DataTable dt = mccProcess.GetLine();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    CboLine.Items.Add(dt.Rows[i]["Line"].ToString());
                    cmbLine.Items.Add(dt.Rows[i]["Line"].ToString());
                }

            }
        }

        private void CboLine_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(CboLine.Text.Trim()))
            {
                GetWO(CboLine.Text.Trim());
            }
        }

        private void GetWO(string line)
        {
            cboWO.Items.Clear();
            TxtWO.Text = "";
            TxtWOType.Text = "";
            TxtModel.Text = "";
            TxtGroup.Text = "";
            TxtWOQty.Text = "";
            txtBuild.Text = "";
            DataTable dt = mccProcess.GetWOByLine(line);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cboWO.Items.Add(dt.Rows[i]["WO"].ToString());
            }
        }

        private void cboWO_SelectedIndexChanged(object sender, EventArgs e)
        {
            TxtWO.Text = cboWO.Text.Trim();
            GetWoinfoBasic(TxtWO.Text);
        }

        private void GetWoinfoBasic(string wo)
        {
            DataTable dt = mccProcess.GetWoinfoBasic(wo.Trim());
            if (dt.Rows.Count > 0)
            {
                TxtWOType.Text = dt.Rows[0]["PN"].ToString().Trim();
                TxtWOQty.Text = dt.Rows[0]["Qty"].ToString().Trim();
                TxtModel.Text = dt.Rows[0]["PN"].ToString().Trim() + "-" + dt.Rows[0]["MB_Rev"].ToString().Trim();
                TxtGroup.Text = dt.Rows[0]["Group"].ToString().Trim();
                txtBuild.Text = dt.Rows[0]["BuildType"].ToString().Trim();
            }
        }

        private void cmdOK_Click(object sender, EventArgs e)
        {
            if (!ChkErr(TxtWO.Text))
            {
                return;
            }
            DataTable dt = mccProcess.QSMS_SetBuildType(TxtWO.Text.Trim(), CboBuildType.Text.Trim(), cmbLine.Text.Trim(), cmbSide.Text.Trim(), CboStation.Text.Trim());
            if (dt.Rows[0]["Result"].ToString() == "1")
            {
                MessageBox.Show(dt.Rows[0]["Msg"].ToString());
                return;
            }
            GetData();
            if (!GetCheckBomFail(TxtWO.Text, CboBuildType.Text))
            {
                return;
            }

            MessageBox.Show("Set BuildType values is OK!");

        }

        private bool GetCheckBomFail(string Work_Order, string BuildType)
        {
            if (string.IsNullOrEmpty(Work_Order.Trim()))
            {
                MessageBox.Show("Please check the WO");
                return false;
            }
            mccProcess.DeleteQSMS_WO(TxtGroup.Text.Trim());

            DataTable dt = mccProcess.GetWoinfoByGroup(TxtGroup.Text.Trim());
            // Wo,BuildType
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (!mccProcess.CheckBom(dt.Rows[i]["WO"].ToString().Trim(), dt.Rows[i]["BuildType"].ToString().Trim()))
                {
                    MessageBox.Show("Check bom fail");
                }
            }
            dt = mccProcess.CheckSapBomFail(TxtGroup.Text.Trim());
            if (dt.Rows.Count > 0)
            {
                QMSSDK.Br.ExcelTool.ExcelExporter.Export(dt);
                return false;
            }
            return true;
        }

        private bool ChkErr(string WO)
        {

            switch (CboBuildType.Text)
            {
                case "1":
                case "2":
                case "3":
                    break;
                case "4":
                    if (string.IsNullOrEmpty(cmbLine.Text) || string.IsNullOrEmpty(cmbSide.Text) || string.IsNullOrEmpty(CboStation.Text))
                    {
                        MessageBox.Show("BuildType=4,Please select the line and side and Station!");
                        return false;
                    }
                    break;
                default:
                    MessageBox.Show("BuildType values can only is 1,2,3 or 4,Please check!");
                    return false;
            }

            DataTable dt = mccProcess.CheckWODispatch(TxtGroup.Text.Trim());

            if (dt.Rows.Count > 0)
            {
                MessageBox.Show("The PCB Work Order is dispatching ,can not be modify ,please check:" + dt.Rows[0]["Work_Order"]);
                return false;
            }
            return true;

        }

        private void FrmDefineBuildType_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("FrmDefineBuildType");
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }
    }
}
