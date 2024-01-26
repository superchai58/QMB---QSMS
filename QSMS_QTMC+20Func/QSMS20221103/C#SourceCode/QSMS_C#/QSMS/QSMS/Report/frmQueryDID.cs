using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.Report
{
    public partial class frmQueryDID : Form
    {
        public frmQueryDID()
        {
            InitializeComponent();
        }
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.Report.QueryDID QueryDID = new DbLibrary.Report.QueryDID();
        DataTable rs;
        DataSet ds;
        private void frmQueryDID_Load(object sender, EventArgs e)
        {
            cboMachine.Items.Clear();
            rs = QueryDID.GetMachine();
            if(rs.Rows.Count > 0)
            {
                for (int i = 0; i < rs.Rows.Count; i++)
                {
                    cboMachine.Items.Add(rs.Rows[i]["Machine"].ToString().Trim()); 
                }
                cboMachine.SelectedIndex = 0;
            }
            cboLine.Items.Clear();
            rs = QueryDID.GetLine();
            if (rs.Rows.Count > 0)
            {
                for (int i = 0; i < rs.Rows.Count; i++)
                {
                    cboLine.Items.Add(rs.Rows[i]["line"].ToString().Trim());
                }
                cboLine.SelectedIndex = 0;
            }
            cboSlot.Items.Clear();
            rs = QueryDID.GetSlot();
            if (rs.Rows.Count > 0)
            {
                for (int i = 0; i < rs.Rows.Count; i++)
                {
                    cboSlot.Items.Add(rs.Rows[i]["Slot"].ToString().Trim());
                }
                cboSlot.SelectedIndex = 0;
            }
            cboCompPN.Items.Clear();
            rs = QueryDID.GetComPN();
            if (rs.Rows.Count > 0)
            {
                for (int i = 0; i < rs.Rows.Count; i++)
                {
                    cboCompPN.Items.Add(rs.Rows[i]["CompPN"].ToString().Trim());
                }
                cboCompPN.SelectedIndex = 0;
            }
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            string txtSQL=""; 
            if(ChkMachine.Checked == true)
            {
                txtSQL = " A.Machine= '" + cboMachine.Text.Trim() + "'";
            }
            if(ChkSlot.Checked == true)
            {
                if (txtSQL.Trim() == "")
                {
                    txtSQL = " A.Slot= '" + cboSlot.Text.Trim() + "'";
                }
                else
                {
                    txtSQL = txtSQL + "and A.Slot= '" + cboSlot.Text.Trim() + "'";
                }
            }
            if (ChkCompPN.Checked == true)
            {
                if (txtSQL.Trim() == "")
                {
                    txtSQL = " A.CompPN= '" + cboCompPN.Text.Trim() + "'";
                }
                else
                {
                    txtSQL = txtSQL + "and A.CompPN= '" + cboCompPN.Text.Trim() + "'";
                }
            }
            if (chkLine.Checked == true)
            {
                if (txtSQL.Trim() == "")
                {
                    txtSQL = " A.machine like '" + cboLine.Text.Trim() + "%'";
                }
                else
                {
                    txtSQL = txtSQL + "and A.machine like '" + cboLine.Text.Trim() + "%'";
                }
            }
            if (txtSQL.Trim() == "")
            {
                MessageBox.Show("Please Select Query Condition！");
                return;
            }
            else
            {
                txtSQL = txtSQL + " order by A.begindatetime,A.machine,A.slot,A.lr,A.did ";
                rs = QueryDID.Query(txtSQL);
                dgv1.DataSource = rs;
            }
        }

        private void btnToExcel_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource == null)
            {
                MessageBox.Show("No Data,Please Query again!");
                return;
            }
            else
            {
                pubFunction.CopyToExcel(dgv1, "use DID", true);
                return;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {            
            if(txtDID.Text.Trim() =="")
            {
                MessageBox.Show("please input DID");
            }
            else
            {
                ds = QueryDID.QueryDIDUse(txtDID.Text.Trim());
                
                dgv1.DataSource = ds.Tables[0];
                dgv1.Columns["ChkTransDateTime"].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";
                dgv2.DataSource = ds.Tables[1];
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] Name = {"use DID","DID Details"};
            if(ds.Tables.Count <= 0)
            {
                MessageBox.Show("No Data,Please Query again!");
                return;
            }
            pubFunction.ExportDataSetToExcel(ds,Name);
        }

        private void frmQueryDID_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmQueryDID");
        }
    }
}
