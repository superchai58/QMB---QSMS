using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace QSMS.QSMS.MCC
{
    public partial class frmTransferPanaMSF : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.MCC.MCCProcess MCCProcess = new DbLibrary.MCC.MCCProcess();
        DataTable rs;
        private struct MSFData
        {
            public string[] CompPN;
            public string[] Slot;
            public string[] Jobpn;
            public string[] Revision;
        }
        private struct MSFBomQty
        {
            public string[] CompPN;
            public string[] Slot;
            public string[] Jobpn;
            public string[] Revision;
            public string[] Qty;
        }
        public frmTransferPanaMSF()
        {
            InitializeComponent();
        }        
        private void btnSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "请选择要上传的文件";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtFile.Text = dialog.FileName;
            }

        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            if(txtFile.Text.Trim() =="")
            {
                MessageBox.Show("You must select a file!!");
                return;
            }
            if(LoadDataFile(txtFile.Text.Trim()) == false)
            {
                MessageBox.Show("Fail");
                return;
            }
            toolStripStatusLabel1.Text = txtFile + "  OK";
            toolStripStatusLabel3.Text = "Finished DateTime:" + DateTime.Now.ToString();
        }
        private bool LoadDataFile(string strFile)
        {
            try
            {
                int I = 0, J = 0, K = 0; ;
                bool CheckS = true;
                string[] arry, arryPN, temp;
                string BomFile, Factory, Line, Machine, Jobpn, Revision, BuildType, Side, jobgroup, strCurrent, strErrMessage;
                string StrSlot, strCompPN, strJobPN;
                MSFData MSFData = new MSFData();
                MSFBomQty MSFBomQty = new MSFBomQty();
                arry = strFile.Split('\\');
                BomFile = arry[arry.GetUpperBound(0)].ToString().Trim();
                temp = BomFile.Split('-');
                if (temp.GetUpperBound(0) != 6)
                {
                    MessageBox.Show("Filename format must be Factory-Line-Machine-PN-Rev-BuildType-Side !");
                    return false;
                }
                if (temp[5].Trim() != "1" && temp[5].Trim() != "2" && temp[5].Trim() != "3")
                {
                    MessageBox.Show("BuildType must be 1,2 or 3.");
                    return false;
                }
                if (temp[6].Trim().Substring(0, 1) != "S" && temp[6].Trim().Substring(0, 1) != "C" && temp[6].Trim().Substring(0, 1) != "Q")
                {
                    MessageBox.Show("Side must be S,C or Q.");
                    return false;
                }
                strErrMessage = "";
                strErrMessage = FunPartNumberCheck(temp[3].Trim());
                if (strErrMessage != "PASS")
                {
                    MessageBox.Show(strErrMessage);
                    return false;
                }
                if (temp[4].Trim().Length != 3 && temp[4].Trim().Length != 2)
                {
                    MessageBox.Show("The Version:" + temp[4].Trim() + ",length must be 2 or 3,Please check the Version!");
                    return false;
                }
                Factory = temp[0].Trim();
                Line = temp[1].Trim();
                Machine = temp[2].Trim();
                Jobpn = temp[3].Trim();
                Revision = temp[4].Trim();
                BuildType = temp[5].Trim();
                Side = temp[6].Trim().Substring(0, 1);

                if (CheckMachine(Line, Machine, Side) == false)
                {
                    return false;
                }

                if ((BuildType == "2" && Side == "S") || (BuildType == "3" && Side == "C"))
                {
                    MessageBox.Show("The side is " + Side + ",BuildType is " + BuildType + ",they are not match,side must be S side.");
                    return false;
                }
                jobgroup = Jobpn.Trim() + "-" + Revision.Trim();

                if (Factory.Trim() == "" || Machine.Trim() == "" || Jobpn == "" || Revision.Trim() == "")
                {
                    MessageBox.Show("Filename format must be Factory-line-Machine-PN-Rev !");
                    return false;
                }
                FileStream fs = new FileStream(strFile, FileMode.Open);
                toolStripStatusLabel1.Text = "GetMEBom_" + BomFile;
                toolStripStatusLabel2.Text = "Start DateTime:" + DateTime.Now.ToString();
                StreamReader rd = new StreamReader(fs);

                while (!rd.EndOfStream)
                {
                    I = I + 1;
                    if (I <= 1)
                    {
                        continue;
                    }
                    strCurrent = rd.ReadLine();
                    strCurrent = (strCurrent.Replace("\r\n", "")).Replace("\t", " ");
                    if (I > 1)
                    {
                        arry = strCurrent.Split(',');
                        StrSlot = arry[1].Trim();
                        strCompPN = arry[2].Trim();
                        strJobPN = arry[6].Trim();
                        if (StrSlot != "" && strCompPN != "")
                        {
                            MSFData.Slot[J] = StrSlot;
                            MSFData.CompPN[J] = strCompPN;
                            arryPN = strJobPN.Split('-');
                            if (arryPN.GetUpperBound(0) > 0)
                            {
                                if (arryPN.GetUpperBound(0) != 2)
                                {
                                    MessageBox.Show("BoardType wrong : " + strJobPN + ", must be PN-REV!");
                                    return false;
                                }
                                else
                                {
                                    rs = MCCProcess.CheckFormat("PARTNUMBER", arryPN[1].ToString());
                                    if (rs.Rows[0]["ErrorCode"].ToString().Trim() != "0")
                                    {
                                        MessageBox.Show(rs.Rows[0]["Result"].ToString().Trim());
                                        return false;
                                    }
                                    if (arryPN[1].Trim().Length != 3 && arryPN[1].Trim().Length != 2)
                                    {
                                        MessageBox.Show("The Version:" + arryPN[1].Trim() + ",length must be 2 or 3,Please check the Version!");
                                        return false;
                                    }
                                    MSFData.Jobpn[J] = arryPN[0].Trim();
                                    MSFData.Revision[J] = arryPN[1].Trim();
                                }
                            }
                            else
                            {
                                MSFData.Jobpn[J] = Jobpn.Trim();
                                MSFData.Revision[J] = Revision.Trim();
                            }
                            J = J + 1;
                        }
                    }
                }
                rd.Close();
                fs.Close();
                for (int i = 0; i <= MSFData.CompPN.GetUpperBound(0); i++)
                {
                    if (i == 0)
                    {
                        MSFBomQty.Slot[K] = MSFData.Slot[i];
                        MSFBomQty.CompPN[K] = MSFData.CompPN[i];
                        MSFBomQty.Jobpn[K] = MSFData.Jobpn[i];
                        MSFBomQty.Revision[K] = MSFData.Revision[i];
                        MSFBomQty.Qty[K] = "1";
                    }
                    else
                    {
                        for (int t = 0; t <= MSFBomQty.CompPN.GetUpperBound(0); t++)
                        {
                            if (MSFBomQty.CompPN[t] == MSFData.CompPN[i] && MSFBomQty.Slot[t] == MSFData.Slot[i] && MSFBomQty.Jobpn[t] == MSFData.Jobpn[i] && MSFBomQty.Revision[t] == MSFData.Revision[i])
                            {
                                MSFBomQty.Qty[t] = (int.Parse(MSFBomQty.Qty[t]) + 1).ToString();
                                CheckS = true;
                            }
                            if (t == MSFBomQty.CompPN.GetUpperBound(0) && CheckS == false)
                            {
                                K = K + 1;
                                MSFBomQty.Slot[K] = MSFData.Slot[i];
                                MSFBomQty.CompPN[K] = MSFData.CompPN[i];
                                MSFBomQty.Jobpn[K] = MSFData.Jobpn[i];
                                MSFBomQty.Revision[K] = MSFData.Revision[i];
                                MSFBomQty.Qty[K] = "1";
                            }
                        }
                    }
                    CheckS = false;
                }
                for (int i = 0; i <= MSFBomQty.Jobpn.GetUpperBound(0); i++)
                {
                    MCCProcess.Del_QSMSMEBom(Machine, Jobpn, jobgroup, MSFBomQty.Revision[i].Trim(), BuildType, Factory, Line);
                    MCCProcess.Insert_QSMSMEBom(Machine, MSFBomQty.Jobpn[i].Trim(), jobgroup, MSFBomQty.Revision[i].Trim(), MSFBomQty.CompPN[i].Trim(), "0", MSFBomQty.Slot[i].Trim(), MSFBomQty.Qty[i].Trim(), BuildType, Side, Parameter.g_userName, Factory, Line);
                }
                MCCProcess.Save_Log(Parameter.g_userName, BomFile.Substring(0, 50));
                MessageBox.Show("*** Load  finish ! ***   \n Total Counter : " + MSFBomQty.Qty.GetUpperBound(0));
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        private string FunPartNumberCheck(string PartNumber)
        {
            DataTable dt = MCCProcess.CheckFormat("PARTNUMBER", PartNumber);
            if (dt.Rows.Count > 0)
            {
                if (dt.Rows[0]["ErrorCode"].ToString().Trim() == "0")
                {
                    return "PASS";
                }
                else
                {
                    return dt.Rows[0]["Result"].ToString().Trim();
                }
            }
            else
            {
                return "Fail";
            }
            
        }

        private bool CheckMachine(string Line,string Machine,string Side)
        {
            if(Machine.Substring(0,1) != "*")
            {
                DataTable dt = MCCProcess.CheckMachine(Line, Machine, Side);
                if (dt.Rows.Count <= 0)
                {
                    MessageBox.Show("The Machine:" + Machine + " in line:" + Line + " and side: " + Side + " (you uploaded) was not defined in machine,please check it in machinetype");
                    return false;
                }
            } 
            return true;
        }


        private void frmTransferPanaMSF_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmMaintainDIDAutoDispatch");
        }
    }
}
