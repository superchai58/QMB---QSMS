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
    public partial class frmUpdateUID : Form
    {
        public frmUpdateUID()
        {
            InitializeComponent();
        }
        DbLibrary.SpecialCase.UpdateUID QueryUID = new DbLibrary.SpecialCase.UpdateUID();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        private void btnQuery_Click(object sender, EventArgs e)
        {             
            DGUidinfo.DataSource = QueryUID.QueryUID();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            string OldUID = txtOldUID.Text.ToString().Trim();
            string NewUID = txtNewUID.Text.ToString().Trim();
            if (string.IsNullOrEmpty(OldUID)==true||string.IsNullOrEmpty(NewUID)==true)
            {
                MessageBox.Show("新旧工号不能为空");
                txtOldUID.Focus();
                return;
            }
            DbLibrary.SpecialCase.UpdateUID UpdateUID = new DbLibrary.SpecialCase.UpdateUID();
            UpdateUID.updateUID(NewUID,OldUID);
            MessageBox.Show("更新UID成功");
        }

        private void DGUidinfo_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                if (DGUidinfo.SelectedCells.Count != 0)
                {
                    this.txtOldUID.Text = DGUidinfo.Rows[e.RowIndex].Cells["UID"].Value.ToString();        
                }
            }
        }

        private void frmUpdateUID_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmUpdateUID");
        }

    }
}
