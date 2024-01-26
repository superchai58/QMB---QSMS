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
    public partial class FrmDIDBake : Form
    {
        public FrmDIDBake()
        {
            InitializeComponent();
        }
        BrLibrary.PublicFunction publicFunction = new BrLibrary.PublicFunction();
        DbLibrary.PD.PDProcess Process = new DbLibrary.PD.PDProcess();
        private void FrmDIDBake_Load(object sender, EventArgs e)
        {

        }

        private void cmdBakeQ_Click(object sender, EventArgs e)
        {
            DataTable dt = Process.QSMS_DIDBake(txtBakeDID.Text.Trim(), Parameter.g_userName, "Query");
            DataGridDIDBake.DataSource = dt.DefaultView;
        }

        private void cmdBakeOK_Click(object sender, EventArgs e)
        {
            if (txtBakeDID.Text != "")
            {
                DataTable dt = Process.QSMS_DIDBake(txtBakeDID.Text.Trim(), Parameter.g_userName, "Bake");
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["result"].ToString() != "OK")
                    {
                        MessageBox.Show(dt.Rows[0]["Desc"].ToString());
                        txtBakeDID.Text = "";
                        return;
                    }
                    reFreshBakeData();
                    txtBakeDID.Text = "";
                }
            }
        }

        private void cmdEndBake_Click(object sender, EventArgs e)
        {
            if (txtBakeDID.Text.Trim() != "")
            {
                DataTable dt = Process.QSMS_DIDBake(txtBakeDID.Text.Trim(), Parameter.g_userName, "EndBake");
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["result"].ToString() != "OK")
                    {
                        MessageBox.Show(dt.Rows[0]["Desc"].ToString());
                        txtBakeDID.Text = "";
                        return;
                    }
                    reFreshBakeData();
                    txtBakeDID.Text = "";
                }
            }
        }

        private void reFreshBakeData()
        {
            DataTable dt = Process.QSMS_DIDBake(txtBakeDID.Text.Trim(), Parameter.g_userName, "Query");
            if (dt.Rows.Count > 0)
            {
                DataGridDIDBake.DataSource = dt.DefaultView;
            }
        }

        private void FrmDIDBake_FormClosed(object sender, FormClosedEventArgs e)
        {
            publicFunction.RemoveForm("FrmDIDBake");
        }
    }
}
