using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.IPQC
{
    public partial class frmInspection_Del : Form
    {
        public frmInspection_Del()
        {
            InitializeComponent();
        }
        DataTable rs = null;
        DbLibrary.IPQC.IPQCProcess Process = new DbLibrary.IPQC.IPQCProcess();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        private void btnQuery_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            string Sdatetime = string.Format("{0:yyyyMMddHHmmss}", dtpSDateTime.Value);
            string Edatetime = string.Format("{0:yyyyMMddHHmmss}",dtpEDateTime.Value);

            rs = Process.QueryInSpect(txtDID.Text.Trim(),Sdatetime,Edatetime);
            if (rs.Rows.Count > 0)
            {
                dataGridView1.DataSource = rs;
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if(txtCompPN.Text.Trim() == "" || txtTransdatetime.Text.Trim() == "")
            {
                MessageBox.Show("Please Select the Data you want to Delete");
                return;
            }
            if (MessageBox.Show("Are you sure to Delete?", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Process.DelInSpect(txtCompPN.Text.Trim(),txtTransdatetime.Text.Trim());
                MessageBox.Show("Delete OK");
                btnQuery_Click(null,null);
            }            
        }


        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                int i = e.RowIndex;
                if (i < 0)
                { return; }
                txtCompPN.Text = dataGridView1.Rows[i].Cells["Comppn"].Value.ToString().Trim();
                txtUpper.Text = dataGridView1.Rows[i].Cells["Upper"].Value.ToString().Trim();
                txtLower.Text = dataGridView1.Rows[i].Cells["Lower"].Value.ToString().Trim();
                txtValue.Text = dataGridView1.Rows[i].Cells["TestValue"].Value.ToString().Trim();
                txtResult.Text = dataGridView1.Rows[i].Cells["TestResult"].Value.ToString().Trim();
                txtTransdatetime.Text = dataGridView1.Rows[i].Cells["TransDateTime"].Value.ToString().Trim();
                this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please query the data again!" + ex.Message);
            }
        }

        private void dataGridView1_KeyUp(object sender, KeyEventArgs e)
        {
            try
            {

                DataGridViewRow dgvCurrentRow = dataGridView1.CurrentRow;
                if (dgvCurrentRow.Index >=0)
                {
                    int i = dgvCurrentRow.Index;
                    txtCompPN.Text = dataGridView1.Rows[i].Cells["Comppn"].Value.ToString().Trim();
                    txtUpper.Text = dataGridView1.Rows[i].Cells["Upper"].Value.ToString().Trim();
                    txtLower.Text = dataGridView1.Rows[i].Cells["Lower"].Value.ToString().Trim();
                    txtValue.Text = dataGridView1.Rows[i].Cells["TestValue"].Value.ToString().Trim();
                    txtResult.Text = dataGridView1.Rows[i].Cells["TestResult"].Value.ToString().Trim();
                    txtTransdatetime.Text = dataGridView1.Rows[i].Cells["TransDateTime"].Value.ToString().Trim();
                    this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please query the data again!" + ex.Message);
            }
        }

        private void frmInspection_Del_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmInspection_Del");
        }
    }
}
