using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.MCC
{

    public partial class IC_Burn : Form
    {
        DbLibrary.MCC.MCCProcess MCC = new DbLibrary.MCC.MCCProcess();

        private string PN = string.Empty;
        private string DID = string.Empty;
        private string CompPN = string.Empty;
        private string Model = string.Empty;
        private string UID = string.Empty;
        private dynamic xlApps = new object();

          
       
        public IC_Burn()
        {
            InitializeComponent();
            UID = Parameter.g_userName;
            ArrayList listPlant = new ArrayList();
            DataTable dt = MCC.Load();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                listPlant.Add(dt.Rows[i][0].ToString().Trim());
            }
            for (int i = 0; i < listPlant.Count; i++)
            {
                cboPN.Items.Add(listPlant[i]);
            }
        }

        private void txtDID_KeyPress(object sender, KeyPressEventArgs e)
        {
            DID = txtDID.Text.Trim();
            DataTable dt = MCC.DID(DID);
            CompPN = dt.Rows[0]["CompPN"].ToString().Trim();
            txtCompPN.Text = CompPN;
            if (DID == "")
            {
                MessageBox.Show("Please input DID");
                return;
            }
            if (PN == "")
            {
                MessageBox.Show("Please select PN ");
                return;
            }

            //txtCompPN.Text = CompPN;
        }
        private void cboPN_SelectedIndexChanged(object sender, EventArgs e)
        {
            PN = this.cboPN.SelectedItem.ToString();
            DataTable dt = MCC.Model(PN);
            Model = dt.Rows[0]["ModelName"].ToString().Trim();
            txtModelName.Text = Model;
        }

        private void btnLinkshearpin_Click(object sender, EventArgs e)
        {
            string msg = "";
            DataTable dt = MCC.ShearPinLinkDID(DID, Model, PN, CompPN, UID);
            msg = dt.Rows[0]["description"].ToString().Trim();
            MessageBox.Show(msg);
        }
  

        private void btnccshearpin_Click(object sender, EventArgs e)
        {

              
               //ExcelTool et = new ExcelTool();
               //et.OpenDecument(System.AppDomain.CurrentDomain.BaseDirectory + @"\IC_ShearPin.xlsx");
               // QuantaSDK.Excel.ExcelHelper s = new QuantaSDK.Excel.ExcelHelper();
            DataTable dt = MCC.ccshearpin(PN, Model);
            
            QMSSDK.Br.ExcelTool.ExcelExporter.Export(dt);
          
        }
     
          

               








    }
}
