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
    public partial class FrmModifyDIDTotalQty : Form
    {
        int CommandType = 0;
        string strErrMessage = "";
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.MCC.ModifyDIDTotalQty Modify = new DbLibrary.MCC.ModifyDIDTotalQty();
        public FrmModifyDIDTotalQty()
        {
            InitializeComponent();
        }

        private void FrmModifyDIDTotalQty_Load(object sender, EventArgs e)
        {
            RefreshDg("");
        }
        public void RefreshDg(string compPN)
        {
            DataTable dt = new DataTable();
            dt = Modify.RefreshDg(compPN);
            DG1.DataSource = dt;
        }

        private void CboCompPN_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && e.ToString() != "")
            {
                CboCompPN.Text = CboCompPN.Text.Replace(" ", "");
                strErrMessage = "";
                strErrMessage = FunPartNumberCheck(CboCompPN.Text);
                if(strErrMessage!="PASS")
                {
                    MessageBox.Show(strErrMessage);
                    CboCompPN.Focus();
                    return;
                }
                else
                {
                    CboCompPN_Click(null, null);
                }
            }
        }

        private void CboCompPN_Click(object sender, EventArgs e)
        {
            CboVendorCode.Focus();
        }
        private string FunPartNumberCheck(string PartNumber)
        {
            DataTable dtCheck = Modify.CheckFormat(PartNumber);
            if (dtCheck.Rows.Count > 0)
            {
                if (dtCheck.Rows[0]["ErrorCode"].ToString().ToUpper() == "0")
                {
                    return "PASS";
                }
                else
                {
                    return dtCheck.Rows[0]["Result"].ToString().ToUpper();
                }
            }
            else
            {
                return "FAIL";
            }
        }

        private void CboDateCode_Click(object sender, EventArgs e)
        {
            CboLotCode.Focus();
        }

        private void CboDateCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && e.ToString() != "")
            {
                CboDateCode_Click(null, null);
            }
        }

        private void CboDID_Click(object sender, EventArgs e)
        {
            cmdFind_Click(null, null);
        }        

        private void CboLotCode_Click(object sender, EventArgs e)
        {
            TxtQty.Focus();
        } 

        private void CboLotCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && e.ToString() != "")
            {
                CboLotCode_Click(null, null);
            }
        }

        private void CboVendorCode_Click(object sender, EventArgs e)
        {
            CboDateCode.Focus();
           
        }
        private void CboVendorCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && e.ToString() != "")
            {
                CboVendorCode_Click(null, null);
            }
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            CboCompPN.Text = "";
            CboVendorCode.Text = "";
            CboDateCode.Text = "";
            CboLotCode.Text = "";
            TxtQty.Text = "";
            CboDID.Text = "";
        }

        private void cmdExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void cmdFind_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = Modify.GetQSMS_DID(CboDID.Text);
            DG1.DataSource = dt;
            cmdUpdate.Enabled = true;
            cmdSave.Enabled = true;
        }

        private void CmdRefresh_Click(object sender, EventArgs e)
        {
            RefreshDg("");
        }

        private void cmdSave_Click(object sender, EventArgs e)
        {
            string TransDate = "";
            string TempDID = "";
            Int64 intDIDInitQty = 0;
            DataTable dt = new DataTable();
            cmdFind.Enabled = true;
            cmdUpdate.Enabled = true;
            cmdSave.Enabled = true;
            cmdCancel.Enabled = true;
            cmdExit.Enabled = true;
            if(TxtQty.Text=="")
            {
                MessageBox.Show("qty can't be empty!!");
                CboCompPN.Enabled = true;
                CboCompPN.Focus();
                return;
            }
            dt = Modify.GetDate();
            TransDate = dt.Rows[0][0].ToString();
            if(CommandType==2)
            {
                if(CboDID.Text=="")
                {
                    MessageBox.Show("DID can't be empty!!");
                    CboDID.Enabled = true;
                    CboDID.Focus();
                    return;
                }
                TempDID = CboDID.Text.Trim();
                dt = Modify.GetRemainQty(TempDID);
                if(dt.Rows.Count>0)
                {
                    intDIDInitQty = Convert.ToInt64(dt.Rows[0]["Qty"].ToString().Trim());
                }
                else
                {
                    MessageBox.Show("Without DID Total Qty, please contact QMS.");
                    return;
                }
                Modify.updateQSMS_DID(TxtQty.Text.Trim(),CboDID.Text.Trim(),Parameter.UID, intDIDInitQty);
                if(CboDID.Text=="")
                {
                    CboDID.Text = TempDID;
                }
                cmdFind_Click(null, null);
            }
            RefreshDg("");
            CommandType = 0;
            TxtGroupQty.Text = "1";
            cmdCancel_Click(null, null);

        }

        private void cmdUpdate_Click(object sender, EventArgs e)
        {
            cmdUpdate.Enabled = true;

            cmdSave.Enabled = true;
            cmdCancel.Enabled = true;
            cmdExit.Enabled = true;
            cmdFind.Enabled = true;

            CboCompPN.Enabled = true;
            CboVendorCode.Enabled = true;
            CboDateCode.Enabled = true;
            CboLotCode.Enabled = true;
            TxtQty.Enabled = true;
            CommandType = 2;
        }

        private void DG1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.RowIndex>-1)
            {
                CboDID.Text = DG1.Rows[e.RowIndex].Cells[0].Value.ToString();
                CboCompPN.Text = DG1.Rows[e.RowIndex].Cells[1].Value.ToString();
                CboVendorCode.Text = DG1.Rows[e.RowIndex].Cells[2].Value.ToString();
                CboDateCode.Text = DG1.Rows[e.RowIndex].Cells[3].Value.ToString();
                CboLotCode.Text = DG1.Rows[e.RowIndex].Cells[4].Value.ToString();
                TxtQty.Text = DG1.Rows[e.RowIndex].Cells[5].Value.ToString();
                cmdUpdate.Enabled = true;
                cmdCancel.Enabled = true;
            }
            
        }

        private void DG1_SelectionChanged(object sender, EventArgs e)
        {

           
        }

        private void TxtQty_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && e.ToString() != "")
            {
                TxtQty.Text = GetDIDQty(TxtQty.Text);
            }
        }
        public string GetDIDQty(string Qty)
        {
            string strQty = "";
            int a = 0;
            for(int i=0;i<Qty.Length;i++)
            {
                if (int.TryParse(Qty.Substring(i, 1),out a) == true)
                    strQty = strQty + Qty.Substring(i, 1);
            }
            return strQty;
        }

        private void FrmModifyDIDTotalQty_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("FrmModifyDIDTotalQty");
        }
    }
}
