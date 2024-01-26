using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using Microsoft.Win32;

namespace QSMS.QSMS
{
    public partial class frmPrinterSetting : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        private string SavePath = "HKEY_CURRENT_USER\\Software\\VB and VBA Program Settings\\SMT\\QSMS";
        private string ReadPath = "Software\\VB and VBA Program Settings\\SMT\\QSMS";
        private string strComm = string.Empty;
        private string[] arryComm;
        private string Printer = string.Empty;
        private string Port = string.Empty;
        private string dmp = string.Empty;
        private string strCommPort = string.Empty;

        public frmPrinterSetting()
        {
            InitializeComponent();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                strComm = txtport1.Text.Trim() + "," + txtport2.Text.Trim() + "," + txtport3.Text.Trim() + "," + txtport4.Text.Trim();
                strCommPort = txtCompPort.Text.Trim();

                if (SaveCheck() == false) return;

                if (rbtnZebra.Checked == true)
                {
                    Printer = "Zebra";
                }
                else
                {
                    Printer = "SATO";
                }

                if (rbtnCom.Checked == true)
                {
                    Port = "COM";
                }
                else if (rbtnLPT.Checked == true)
                {
                    Port = "LPT";
                }
                else
                {
                    Port = "Network";
                }

                if (rbtn300.Checked == true)
                {
                    dmp = "300";
                }
                else
                {
                    dmp = "200";
                }

                Registry.SetValue(SavePath, "Comm", strComm, RegistryValueKind.String);
                Registry.SetValue(SavePath, "CommPort", strCommPort, RegistryValueKind.String);
                Registry.SetValue(SavePath, "DPM", dmp, RegistryValueKind.String);
                Registry.SetValue(SavePath, "Port", Port, RegistryValueKind.String);
                Registry.SetValue(SavePath, "Printer", Printer, RegistryValueKind.String);

                MessageBox.Show("Registry Setting OK!(打印设置成功!)");
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("System Exception ErrMsg:{0},please contact QMS developer", ex.Message);

            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmPrinterSetting_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmPrinterSetting");
        }

        private void frmPrinterSetting_Load(object sender, EventArgs e)
        {
            RegistryKey rgst = Registry.CurrentUser.OpenSubKey(ReadPath);
            if (rgst == null)
            {
                RegistryKey rgstCreate = Registry.CurrentUser;
                RegistryKey software = rgstCreate.CreateSubKey(ReadPath);

                rgst = Registry.CurrentUser.OpenSubKey(ReadPath);
            }

            if (!string.IsNullOrEmpty(rgst.GetValue("CommPort", "").ToString().Trim()))
            {
                txtCompPort.Text = rgst.GetValue("CommPort", "").ToString().Trim();
            }
            else
            {
                txtCompPort.Text = "1";
            }

            if (!string.IsNullOrEmpty(rgst.GetValue("Comm", "").ToString().Trim()))
            {
                arryComm = rgst.GetValue("Comm", "").ToString().Trim().Split(',');
                txtport1.Text = arryComm[0];
                txtport2.Text = arryComm[1];
                txtport3.Text = arryComm[2];
                txtport4.Text = arryComm[3];
            }
            else
            {
                txtport1.Text = "9600";
                txtport2.Text = "N";
                txtport3.Text = "8";
                txtport4.Text = "1";
            }

            if (!string.IsNullOrEmpty(rgst.GetValue("Port", "").ToString().Trim()))
            {
                if (rgst.GetValue("Port", "").ToString().Trim() == "LPT")
                {
                    rbtnLPT.Checked = true;
                }
                else if (rgst.GetValue("Port", "").ToString().Trim() == "COM")
                {
                    rbtnCom.Checked = true;
                }
                else
                {
                    rbtnNetwork.Checked = true;
                }
            }

            if (!string.IsNullOrEmpty(rgst.GetValue("Printer", "").ToString().Trim()))
            {
                if (rgst.GetValue("Printer", "").ToString().Trim() == "Zebra")
                {
                    rbtnZebra.Checked = true;
                }
                else
                {
                    rbtnSATO.Checked = true;
                }
            }
            if (!string.IsNullOrEmpty(rgst.GetValue("DPM", "").ToString().Trim()))
            {
                if (rgst.GetValue("DPM", "").ToString().Trim() == "200")
                {
                    rbtn200.Checked = true;
                }
                else
                {
                    rbtn300.Checked = true;
                }
            }
        }

        private bool SaveCheck()
        {
            if (string.IsNullOrEmpty(txtCompPort.Text.Trim()))
            {
                MessageBox.Show("Please Input CompPort");
                return false;
            }

            if (string.IsNullOrEmpty(txtport1.Text.Trim()) || string.IsNullOrEmpty(txtport2.Text.Trim()) ||
                string.IsNullOrEmpty(txtport3.Text.Trim()) || string.IsNullOrEmpty(txtport4.Text.Trim()))
            {
                MessageBox.Show("Please Input Settings");
                return false;
            }
            return true;
        }
    }
}
