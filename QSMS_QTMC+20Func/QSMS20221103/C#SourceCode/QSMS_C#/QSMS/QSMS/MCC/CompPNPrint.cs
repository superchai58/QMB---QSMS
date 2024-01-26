using PrinterLib;
using QSMS.DbLibrary.MCC;
using System;
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
    public partial class CompPNPrint : Form
    {
       
     BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
     PrintLabel Print = new PrintLabel();
       private string strDIDPrintLabel = string.Empty;
         private string strPrintPort = string.Empty;
        private string strCommSetting = string.Empty;
        private string BR = string.Empty;
        DbLibrary.MCC.MCCProcess MCC = new DbLibrary.MCC.MCCProcess();
        string PortName = string.Empty;
        int BaudRate = 19200;
        public CompPNPrint()
        {
            InitializeComponent();
            txtUID.Text = Parameter.g_userName;
            txtUID.ForeColor = Color.Red;
           
            strPrintPort = pubFunction.ConfigListGetValue("PrintPort");
            strCommSetting = pubFunction.ConfigListGetValue("CommSetting");
            BR = pubFunction.ConfigListGetValue("PrintBR");

        }
        //private void frmCompPrint_Load(object sender, EventArgs e)
        //{
        //    reFreshData();
        //    txtUID.Text = Parameter.g_userName;
        //    txtUID.ForeColor = Color.Red;
           
        //    strPrintPort = pubFunction.ConfigListGetValue("PrintPort");
        //    strCommSetting = pubFunction.ConfigListGetValue("CommSetting");
            
        //}

          private void reFreshData()
        {
            DataTable dt = MCC.QSMS_MCC_QueryDataByType("MCC_GetCompPrintLog", "", "", "", "", "");
            dataGridView1.DataSource = dt;
        }


          private void btnPrint_Click(object sender, EventArgs e)
          {
              string CompPN = txtCompPN.Text.Trim();
              string NewCompPN = txtNewCompPN.Text.Trim();
              string UID = txtUID.Text.Trim();
              string Msg = string.Empty;
              
              BaudRate = 0;
              PortName = "";
              BaudRate = int.Parse(BR);
              PortName = pubFunction.ConfigListGetValue("PRINTERPORTTYPE");
              try
              {
                  if (UID == "")
                  {
                      lblMsg.Text = "UID为空!";
                      return;
                  }
                  if (CompPN == "")
                  {
                      lblMsg.Text = "CompPN is blank!!!";
                      return;
                  }
                  int len = CompPN.Length;
                  if (len < 11)
                  {
                      lblMsg.Text = "The CompPN's length must be >11 ";
                      return;
                  }
                  if (NewCompPN == "")
                  {
                      lblMsg.Text = "NewCompPN is blank!!";
                      return;
                  }

                  strDIDPrintLabel = Application.StartupPath + "\\" + pubFunction.ConfigListGetValue("CompPNLabelPrint");
                  if (File.Exists(strDIDPrintLabel) == false)
                  {
                      lblMsg.Text = "在路径[" + strDIDPrintLabel + "]没找到对应模板!";
                      return;
                  }



                  StreamReader reader = new StreamReader(strDIDPrintLabel, Encoding.Default);
                  string tmpPrintStr = reader.ReadToEnd();
                  reader.Close();
                  tmpPrintStr = tmpPrintStr.ToUpper();
                  if (Print.LabelSetting(strCommSetting, strPrintPort, 1, ref Msg) == false)
                  {
                      lblMsg.Text = Msg;
                      return;
                  }
                  DataTable dt = MCC.QSMS_CompPN(NewCompPN);

                  DataTable dtPrint = null;
                  dtPrint = dt.Clone();
                  dtPrint.Clear();
                  dtPrint.ImportRow(dt.Rows[0]);
                  tmpPrintStr = Extensions.TemplateReplace(tmpPrintStr, dt);
                  //if (Print.Print(tmpPrintStr, dtPrint, ref Msg) == false)
                  //{
                  //    lblMsg.Text = Msg;
                  //    return;
                  //}

                  Print.CommControl commcon = new Print.CommControl(strPrintPort, BaudRate);

                  if (commcon.Write(tmpPrintStr) == false)
                  {
                      MessageBox.Show("打印失败,请检查打印机或联系QMS人员!!");
                      commcon.Dispose();
                      return;
                  }

                  commcon.Dispose();

                  lblMsg.Text = "打印成功!";
                  lblMsg.ForeColor = Color.Green;

                  //reFreshData();
                 
                  txtCompPN.Text = "";
                  txtNewCompPN.Text = "";



              }

              catch (Exception ex)
              {
                  lblMsg.Text = ex.Message + ",请联系QMS!";
              }

          }
    }
}
