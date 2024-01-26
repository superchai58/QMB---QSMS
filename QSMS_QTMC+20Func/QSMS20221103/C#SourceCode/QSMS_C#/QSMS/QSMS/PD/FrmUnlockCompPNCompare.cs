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
    public partial class FrmUnlockCompPNCompare : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.PD.PDProcess PD = new DbLibrary.PD.PDProcess();

        public FrmUnlockCompPNCompare()
        {
            string StrSQL = "";
            InitializeComponent();
            reFreshData();

            dtpSDate.Text = DateTime.Now.ToString();
            dtpSDate.Text = DateTime.Now.ToString();
            StrSQL = "select distinct Line from QSMS_woGroup(nolock)";

            CobLine.Items.Clear();
            DataTable dtResult = PD.QSMS_EXE(StrSQL);

            if (dtResult.Rows.Count > 0)
            {
                foreach (DataRow rw in dtResult.Rows)
                {
                    CobLine.Items.Add(rw["Line"].ToString());
                }
            }
        }

        private void reFreshData()
        {
            txtDID.Text = "";
            txtCompPN.Text = "";
            txtReason.Text = "";

            string strSQL = "select Line ,   DID ,CompPN , GroupID from  QSMS_CompPNCheck(nolock) ";
            DataTable dt = PD.QSMS_EXE(strSQL);

            if (dt.Rows.Count > 0)
            {
                GV_Data1.DataSource = dt;
            }

        }

        private void InsertLog(string sql)
        {
            string SQLLog = "";
            sql = sql.Replace("'", "''");

            SQLLog = "insert into QSMS_LOG(  [System_Name]  ,[Event_No]  ,[DID] ,[User_Name],[ReturnQty] ,[Trans_Date])" +
            "values('QSMS_UnlockCompPNCheck','QSMS','" + sql + "','" + Parameter.UID + "','0', dbo.FormatDate (GETDATE (),'YYYYMMDDHHNNDD'))";

            PD.QSMS_EXE(SQLLog);
        }

        private void txtCompPN_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && txtCompPN.Text.Trim() != "")
            {
                if (txtCompPN.Text.IndexOf(";") > 0)
                {
                    txtCompPN.Text = txtCompPN.Text.Substring(0, txtCompPN.Text.IndexOf(";"));
                }

                txtReason.Focus();
            }
        }

        private void GV_Data1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtDID.Text = "";
            txtCompPN.Text = "";
            txtReason.Text = "";

            txtDID.Text = GV_Data1.Rows[e.RowIndex].Cells[1].Value.ToString();
        }

        private void btnUnlock_Click(object sender, EventArgs e)
        {
            string strSQL = "";

            if (txtCompPN.Text == "" || txtReason.Text == "")
            {
                lblStatus.Text = "Reason or PN is not null";
                return;
            }

            strSQL = "insert into QSMS_UnlockCompPNCheck(GroupID , Line ,WO , DID , OLDCompPN ,NewCompPN ,Side ,Reason ,TransDateTime , UID) " + " select GroupID , Line ,WO , DID , CompPN , '"
                + txtCompPN.Text.Trim() + "'  ,Side , N'" + txtReason.Text.Trim() + "', dbo.formatdate(getdate(),'yyyymmddhhnnss'), '" + Parameter.UID + "'  from QSMS_CompPNCheck(nolock) where DID = '" + txtDID.Text.Trim() + "'";
            PD.QSMS_EXE(strSQL);

            strSQL = "delete from QSMS_CompPNCheck  where DID = '" + txtDID.Text.Trim() + "' ";
            PD.QSMS_EXE(strSQL);

            lblStatus.Text = "unlock is ok";
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            DataTable dt = GetData();

            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("No  Data !");
                return;
            }

            GV_Data2.DataSource = dt;
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            DataTable dt = GetData();

            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("No  Data !");
            }

            pubFunction.doExport(dt);
        }

        private DataTable GetData()
        {
            string BeginDate = dtpSDate.Text.ToString();
            string EndDate = dtpEDate.Text.ToString();

            BeginDate = BeginDate.Replace("/", "");
            BeginDate = BeginDate.Replace("-", "");
            EndDate = EndDate.Replace("/", "");
            EndDate = EndDate.Replace("-", "");

            if (CobLine.Text != "")
            {
                string strSQL = " select top 1000 * from  QSMS_UnlockCompPNCheck(nolock) where  Line='" + CobLine.Text.Trim() + "'   and  TransDateTime>='" + BeginDate + "000000'  and  TransDateTime<='" + EndDate + "235900' order by TransDateTime desc ";

                return PD.QSMS_EXE(strSQL);
            }
            else
            {
                MessageBox.Show("Please  choose Line!");
            }

            return null;
        }

        private void FrmUnlockCompPNCompare_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("FrmUnlockCompPNCompare");
        }
    }
}
