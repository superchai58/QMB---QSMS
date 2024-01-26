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
    public partial class FrmPNCompare : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.PD.PDProcess PD = new DbLibrary.PD.PDProcess();
        DataTable dtResult = new DataTable();
        //private const int WM_RBUTTONDOWN = 0x0204;
        //private const int WM_GETTEXT = 0x000d;
        //private const int WM_COPY = 0x0301;
        bool IsPaste = false;
        string DIDLine = string.Empty;
        string strLine = Parameter.strLine;

        public FrmPNCompare()
        {
            InitializeComponent();
            txtStatus.Text = "";

            GetLine();

            if (CheckDIDCheckStatus() == false)
            {
                txtDID.ReadOnly = true;
                txtCompPN.ReadOnly = true;

                txtStatus.BackColor = Color.Red;//&HFF&
                //txtStatus.ForeColor = Color.Black;//&H8000000E
            }

            this.Text = this.Text + " Line：" + strLine;
            dlLine.Focus();

        }

        private void GetLine()
        {
            string strSQL = "select distinct Line from Sap_WO_List where Trans_Date>dbo.FormatDate(getdate()-32,'YYYYMMDDHHNNSS') order by Line asc";
            DataTable dt = PD.QSMS_EXE(strSQL);

            txtStatus.Text = "";
            dlLine.Items.Clear();
            //dlLine.Items.Add("");

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dlLine.Items.Add(dt.Rows[i]["Line"].ToString().Trim());
                }
            }

        }

        private bool CheckDIDCheckStatus()
        {
            string strSQL = "select * from QSMS_CompPNCheck(nolock) where Status= 'FAIL'  and Line ='" + strLine + "'";
            DataTable dt = PD.QSMS_EXE(strSQL);
            txtStatus.Text = "";

            if (dt.Rows.Count > 0)
            {
                txtDID.Text = dt.Rows[0]["DID"].ToString().Trim();
                txtCompPN.Text = dt.Rows[0]["CompPN"].ToString().Trim();
                txtStatus.Text = "This DID and CompPN is not match ,please unlock at first !";

                return true;
            }
            else
            {
                return false;
            }

        }

        private bool Func_CheckDIDExist()
        {
            string strSQL = "select  0 from QSMS_CompPNCheck(nolock) where DID= '" + txtDID.Text.Trim() + "'";            
            DataTable dt = PD.QSMS_EXE(strSQL);
            txtStatus.Text = "";

            if (dt.Rows.Count > 0)
            {
                txtDID.Text = "";
                txtCompPN.Text = "";

                return true;
            }
            else
            {
                return false;
            }

        }

        private string GetDIDLine()
        {
            string strSQL = "select  top 1  Line  from QSMS_Dispatch(nolock) where DID = '" + txtDID.Text.Trim() + "'";
            DataTable dt = PD.QSMS_EXE(strSQL);

            if (dt.Rows.Count > 0)
            {
                return dt.Rows[0]["Line"].ToString().Trim();
            }
            else
            {
                return "";
            }

        }

        private void SaveData(string Line, string Side, string WO, string GroupID, string Status)
        {
            string strSQL = "insert into QSMS_CompPNCheck ( GroupID , Line ,side,  WO , DID , CompPN , Status ,TransDateTime , UID) values('"
               + GroupID + "','" + Line + "','" + Side + "','" + WO + "','" + txtDID.Text.Trim() + "','" + txtCompPN.Text.Trim() + "','" + Status + "',dbo.FormatDate (GETDATE (),'YYYYMMDDHHNNDD'),'" + Parameter.UID + "'";
            DataTable dt = PD.QSMS_EXE(strSQL);

            reFreshData();
        }

        private void reFreshData()
        {
            string strSQL = "select  top 1000 *  from QSMS_CompPNCheck  order by transdatetime desc";
            DataTable dt = PD.QSMS_EXE(strSQL);

            if (dt.Rows.Count > 0)
            {
                GV_Data.DataSource = dt;
            }

        }

        private void btn_Excel_Click(object sender, EventArgs e)
        {
            DataTable dtResult = QueryData();

            if (dtResult.Rows.Count > 0)
            {
                GV_Data.DataSource = dtResult;
                pubFunction.doExport(dtResult);
            }
        }

        private DataTable QueryData()
        {
            DataTable dtResult = new DataTable();
            string BeginDate = Begin_Date.Text.ToString();
            string EndDate = End_Date.Text.ToString();

            BeginDate = BeginDate.Replace("/", "");
            BeginDate = BeginDate.Replace("-", "");
            EndDate = EndDate.Replace("/", "");
            EndDate = EndDate.Replace("-", "");

            string strSQL = "SELECT  GroupID ,   Line , WO , DID , CompPN , Status ,Side, UID ,TransDateTime   FROM  QSMS_CompPNCheck where  Line='" + dlLine.Text.Trim() + "'   and  TransDateTime>='" + BeginDate + "000000'  and  TransDateTime<='" + EndDate + "235900' order by TransDateTime desc";
            dtResult = PD.QSMS_EXE(strSQL);

            return dtResult;
        }

        private void btn_Find_Click(object sender, EventArgs e)
        {
            txtStatus.Text = "";

            if (dlLine.Text.Trim() == "")
            {
                txtStatus.Text = "Please choose Line!";
                return;
            }

            DataTable dtResult = QueryData();

            if (dtResult.Rows.Count > 0)
            {
                GV_Data.DataSource = dtResult;
            }
        }

        //protected override void WndProc(ref Message m)
        //{
        //    if (m.Msg == WM_RBUTTONDOWN || m.Msg == WM_GETTEXT || m.Msg == WM_COPY)
        //        return;//WM_RBUTTONDOWN是为了不让出现鼠标菜单
        //    base.WndProc(ref m);
        //}

        private void txtDID_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = IsPaste;
        }

        private void txtDID_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                //return;
                txtDID.ContextMenu = new ContextMenu();
            }
        }

        private void txtDID_KeyDown(object sender, KeyEventArgs e)
        {
            string Line;
            timer_DID.Enabled = true;
            txtStatus.Text = "";
            txtStatus.BackColor = Color.Pink;

            if (e.Control && e.KeyCode.ToString().ToUpper() == "V")
            {
                IsPaste = true;
            }
            else
            {
                IsPaste = false;
            }

            if (e.KeyCode == Keys.Enter && txtDID.Text.Trim() != "")
            {
                timer_DID.Enabled = false;

                if (CheckDIDCheckStatus() == false)
                {
                    txtDID.ReadOnly = true;
                    txtCompPN.ReadOnly = true;

                    txtStatus.BackColor = Color.Red;//&HFF&
                    //txtStatus.ForeColor = Color.Black;//&H8000000E
                }

                Line = GetDIDLine();
                if (Line != "")
                {
                    if (strLine.ToUpper() != dlLine.Text.Trim().ToUpper())
                    {
                        txtDID.Text = "";
                        txtStatus.BackColor = Color.Red;
                        txtStatus.Text = "此DID属于" + Line + " 线，请再次确认";
                        MessageBox.Show("此DID属于" + Line + " 线，请再次确认");

                        txtDID.Focus();
                        return;
                    }
                }
                else
                {
                    txtDID.Text = "";
                    txtStatus.BackColor = Color.Red;
                    txtStatus.Text = "此DID 在Dispatch 中不存在 ，请再次确认";
                    MessageBox.Show("此DID 在Dispatch 中不存在 ，请再次确认");

                    txtDID.Focus();
                    return;
                }

                if (!Func_CheckDIDExist())
                {
                    txtCompPN.Focus();
                }
                else
                {
                    txtDID.Text = "";
                    txtStatus.BackColor = Color.Red;
                    txtStatus.Text = txtDID.Text + "This DID  has been checked !";
                    MessageBox.Show(txtDID.Text + "This DID  has been checked !");

                    txtDID.Focus();
                    return;
                }

            }
        }

        private void txtCompPN_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = IsPaste;
        }

        private void txtCompPN_KeyDown(object sender, KeyEventArgs e)
        {
            txtStatus.Text = "";
            string SQL = "";
            string strstatus = "";
            DataTable dtResult = new DataTable();
            timer_CompPN.Enabled = true;
            txtStatus.BackColor = Color.Pink;
            

            if (e.Control && e.KeyCode.ToString().ToUpper() == "V")
            {
                IsPaste = true;
            }
            else
            {
                IsPaste = false;
            }

            if (e.KeyCode == Keys.Enter && txtCompPN.Text.Trim() != "")
            {
                timer_CompPN.Enabled = false;

                if (txtCompPN.Text.IndexOf(";") > 0)
                {
                    txtCompPN.Text = txtCompPN.Text.Substring(0, txtCompPN.Text.IndexOf(";"));
                }

                if (CheckDIDCheckStatus() == false)
                {
                    txtDID.ReadOnly = true;
                    txtCompPN.ReadOnly = true;

                    txtStatus.BackColor = Color.Red;//&HFF&
                    //txtStatus.ForeColor = Color.Black;//&H8000000E
                }

                if (Func_CheckDIDExist())
                {
                    txtStatus.BackColor = Color.Red;
                    txtStatus.Text = txtDID.Text + "This DID  has been checked !";
                    MessageBox.Show(txtDID.Text + "This DID  has been checked !");
                    return;
                }

                SQL = "select top 1  CompPN ,  Line , side ,Work_Order ,GroupID  from  QSMS_Dispatch  where DID ='" + txtDID.Text + "'";
                dtResult = PD.QSMS_EXE(SQL);

                if (dtResult.Rows.Count > 0)
                {
                    if (dtResult.Rows[0]["Line"].ToString().Trim() == strLine)
                    {
                        if (dtResult.Rows[0]["Line"].ToString().Trim() == strLine)
                        {
                            strstatus = "PASS";
                            txtStatus.BackColor = Color.Pink;
                            txtStatus.Text = strstatus;
                        }
                        else
                        {
                            strstatus = "Fail";
                            txtStatus.BackColor = Color.Red;
                            txtStatus.Text = txtDID.Text + " and " + txtCompPN.Text + "is not match, please unlock at first !";
                            txtDID.ReadOnly = true;
                            txtCompPN.ReadOnly = true;
                            return;
                        }
                    }
                    else
                    {
                        txtStatus.Text = txtDID.Text + " belong to " + dtResult.Rows[0]["Line"].ToString().Trim();
                        txtDID.Text = "";
                        txtCompPN.Text = "";
                        txtDID.Focus();
                        return;
                    }
                }
                else
                {
                    txtDID.Text = "";
                    txtCompPN.Text = "";
                    txtStatus.Text = txtDID.Text + " is not exist in Dispatch!";
                    txtDID.Focus();
                    return;
                }

                SaveData(dtResult.Rows[0]["Line"].ToString().Trim(), dtResult.Rows[0]["side"].ToString().Trim(), dtResult.Rows[0]["Work_Order"].ToString().Trim(), dtResult.Rows[0]["GroupID"].ToString().Trim(), strstatus);
                txtDID.Focus();

                if (strstatus != "Fail")
                {
                    txtDID.Text = "";
                    txtCompPN.Text = "";
                }
            }
        }       

        private void txtCompPN_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                txtCompPN.ContextMenu = new ContextMenu();
            }
        }

        private void timer_DID_Tick(object sender, EventArgs e)
        {
            txtDID.Text = "";
        }

        private void timer_CompPN_Tick(object sender, EventArgs e)
        {
            txtCompPN.Text = "";
        }

        private void FrmPNCompare_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("FrmPNCompare");
        }






    }
}
