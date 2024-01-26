using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

namespace QSMS.QSMS.MCC
{
    public partial class frmTransferFujiXML : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.MCC.MCCProcess MCC = new DbLibrary.MCC.MCCProcess();
        XmlDocument xmlDoc = new XmlDocument();
        private string folderPath = string.Empty;

        public frmTransferFujiXML()
        {
            InitializeComponent();
        }

        private void frmTransferFujiXML_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmTransferFujiXML");
        }

        private void frmTransferFujiXML_Load(object sender, EventArgs e)
        {
            txtFile.Text = "C:\\";
            ScanFolder(txtFile.Text);
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            try
            {
                folderBrowserDialog1.SelectedPath = "C:\\";
                if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                {
                    txtFile.Text = folderBrowserDialog1.SelectedPath;
                    folderPath = txtFile.Text;
                    ScanFolder(txtFile.Text);
                }
                else
                {
                    return;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ScanFolder(string FilePath)
        {
            if (FilePath == "")
            {
                return;
            }
            FilePath = FilePath.Replace(@"/", @"\");

            if (FilePath.Substring(FilePath.Length - 1) != @"\")
            {
                FilePath = FilePath + @"\";
            }
            string[] datas = Directory.GetFiles(FilePath, "*.xml");

            foreach (string file in datas)
            {
                lstFile1.Items.Add(Path.GetFileName(file));
            }

            //for (int i = 0; i < datas.Length; i++)
            //{
            //    lstFile1.Items.Add(datas[i]);
            //}
        }

        private void cmdADD_Click(object sender, EventArgs e)
        {
            if (lstFile1.SelectedItems.Count <= 0)
            {
                return;
            }
            if (lstFile1.Items.Count <= 0)
            {
                return;
            }
            int index = lstFile1.SelectedItems[0].Index;
            if (index < 0)
            {
                return;
            }
            ListViewItem lst = new ListViewItem(lstFile1.Items[index].SubItems[0].Text);
            lstFile2.Items.Add(lst);
            lstFile1.Items.RemoveAt(index);
            if (lstFile1.Items.Count != index)
            {
                this.lstFile1.Focus();
                this.lstFile1.Items[0].Selected = true;
            }
        }

        private void cmdADDALL_Click(object sender, EventArgs e)
        {
            if (lstFile1.Items.Count <= 0)
            {
                return;
            }
            for (int i = 0; i < lstFile1.Items.Count; i++)
            {
                ListViewItem lst = new ListViewItem(lstFile1.Items[i].SubItems[0].Text);
                lstFile2.Items.Add(lst);
            }
            lstFile1.Items.Clear();
        }

        private void cmdDEL_Click(object sender, EventArgs e)
        {
            if (lstFile2.SelectedItems.Count <= 0)
            {
                return;
            }
            if (lstFile2.Items.Count <= 0)
            {
                return;
            }
            int index = lstFile2.SelectedItems[0].Index;
            if (index < 0)
            {
                return;
            }
            ListViewItem lst = new ListViewItem(lstFile2.Items[index].SubItems[0].Text);
            lstFile1.Items.Add(lst);
            lstFile2.Items.RemoveAt(index);
            if (lstFile2.Items.Count != index)
            {
                this.lstFile2.Focus();
                this.lstFile2.Items[index].Selected = true;
            }
        }

        private void cmdDELALL_Click(object sender, EventArgs e)
        {
            if (lstFile2.Items.Count <= 0)
            {
                return;
            }
            for (int i = 0; i < lstFile2.Items.Count; i++)
            {
                ListViewItem lst = new ListViewItem(lstFile2.Items[i].SubItems[0].Text);
                lstFile1.Items.Add(lst);
            }
            lstFile2.Items.Clear();
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            if (lstFile2.Items.Count <= 0)
            {
                return;
            }
            string Path = string.Empty;
            string FileName = string.Empty;
            string Side = string.Empty;
            string strErrMessage = string.Empty;
            int count = lstFile2.Items.Count - 1;
            Path = folderPath;
            for (int i = 0; i <= count; i++)
            {
                FileName = lstFile2.Items[i].SubItems[0].Text;
                txtFile.Text = Path + "\\" + FileName;
                lstFile2.Items[0].Selected = true;
                if (i == 0)
                {
                    string[] temp = FileName.Split('-');
                    if (temp.GetUpperBound(0) != 6)
                    {
                        lblMsg.Text = "FileName format must be Factory-Line-Machine-PN-REV-BuildType-Side.xml!";
                        return;
                    }

                    Side = temp[6].Substring(0, temp[6].Length - 4).Trim().ToUpper();
                    if (temp[5] != "1" && temp[5] != "2" && temp[5] != "3" && temp[5] != "4")
                    {
                        lblMsg.Text = "BuildType must be 1,2,3 or 4.";
                        return;
                    }

                    if (Side != "S" && Side != "C" && Side != "Q")
                    {
                        lblMsg.Text = "Side must be S,C or Q.";
                        return;
                    }

                    if (temp[5] == "2" && Side != "S")
                    {
                        lblMsg.Text = "The Side is " + Side + ",BuildType is " + temp[5] + ",they are not match,the Side must be S side.";
                        return;
                    }

                    if (temp[5] == "3" && Side != "C")
                    {
                        lblMsg.Text = "The Side is " + Side + ",BuildType is " + temp[5] + ",they are not match,the Side must be C side.";
                        return;
                    }

                    strErrMessage = FunPartNumberCheck(temp[3]);
                    if (strErrMessage != "PASS")
                    {
                        lblMsg.Text = strErrMessage;
                        return;
                    }

                    if (temp[4].Trim().Length != 3 && temp[4].Trim().Length != 2)
                    {
                        lblMsg.Text = "The Version:" + temp[4].Trim() + ",the length must be 2 or 3.Please check the Version!";
                        return;
                    }

                    if (MessageBox.Show("Do you want to delete all machine bom by line, side (except OTHERS)?", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        string Factory = temp[0].Trim();
                        string Line = temp[1].Trim();
                        string Machine = temp[2].Trim();
                        string MBPN = temp[3].Trim();
                        string Version = temp[4].Trim();
                        string jobgroup = MBPN + "-" + Version;
                        MCC.MCC_QueryDataByType("MCC_DelQSMS_MEBOM", "", "", jobgroup, Side, Factory, Line, "", "", "");
                    }
                }
                if (LoadDataFile(txtFile.Text) == false)
                {
                    lblMsg.Text = "Fail";
                }
                else
                {
                    lblMsg.Text = "Finish";
                }
            }
        }

        private string FunPartNumberCheck(string PartNumber)
        {
            DataTable dtCheck = MCC.CheckFormat(PartNumber);
            if (dtCheck.Rows.Count > 0)
            {
                if (dtCheck.Rows[0]["ErrorCode"].ToString().ToUpper() == "0")
                {
                    return "PASS";
                }
                else
                {
                    return dtCheck.Rows[0]["Result"].ToString().ToUpper();
                }
            }
            else
            {
                return "FAIL";
            }
        }

        private bool LoadDataFile(string FilePath)
        {
            string FileName = string.Empty;
            string[] temp = FilePath.Split('\\');

            FileName = temp[temp.GetUpperBound(0)];

            string[] Filetemp = FileName.Split('-');

            if (Filetemp.GetUpperBound(0) != 6)
            {
                lblMsg.Text = "FileName format must be Factory-Line-Machine-PN-REV-BuildType-Side.xml!";
                return false;
            }

            LoadArrayWithData(FileName, FilePath);

            return true;
        }

        private void LoadArrayWithData(string FileName, string FilePath)
        {
            try
            {
                bool NXT = false, AIMEX = false;
                DataTable dt = new DataTable();
                int TabQty = 0, SlotQty = 0, Num = 0;

                string[] BrdPN = new string[500];
                string[] BrdRev = new string[500];

                string[] temp = FileName.Split('-');
                string Machine = temp[2].Trim();
                string Line = temp[1].Trim();
                string Side = temp[6].Substring(0, temp[6].Length - 4).Trim().ToUpper();
                string strErrMessage = "", strLR = "", str = "", Factory = "", MBPN = "", Version = "", jobgroup = "", BuildType = "", strLocation = "", StrSlot = "";
                if (Machine.IndexOf("NXT") > -1)
                {
                    NXT = true;
                }
                else if (Machine.IndexOf("AIMEX") > -1)
                {
                    AIMEX = true;
                }
                else
                {
                    NXT = false;
                }

                if (NXT == false && AIMEX == false)
                {
                    dt = MCC.MCC_QueryDataByType("MCC_GetMachineInfo", "", "", Machine, Line, Side, "", "", "", "");
                    if (dt.Rows.Count > 0)
                    {
                        TabQty = Convert.ToInt32(dt.Rows[0]["Qty"].ToString());
                        SlotQty = Convert.ToInt32(dt.Rows[0]["MaxSlotNum"].ToString());
                    }
                    else
                    {
                        MessageBox.Show("Can't find The Machine:" + Machine + " in line:" + Line + " and side: " + Side + " ,please define it in machinetype or check its format!", "LoadArrayWithData");
                        return;
                    }
                }
                Factory = temp[0].Trim();
                MBPN = temp[3].Trim();
                Version = temp[4].Trim();
                jobgroup = MBPN + "-" + Version;
                BuildType = temp[5].Trim();

                dt = null;
                dt = MCC.MCC_QueryDataByType("MCC_GetFujiBrdSeqMapping", "", "", MBPN, Version, "", "", "", "", "");
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    BrdPN[Convert.ToInt32(dt.Rows[i]["BrdSeq"].ToString())] = dt.Rows[i]["BrdPN"].ToString();
                    BrdRev[Convert.ToInt32(dt.Rows[i]["BrdSeq"].ToString())] = dt.Rows[i]["BrdRev"].ToString();
                }
                List<StrData> list = new List<StrData>();
                xmlDoc.Load(FilePath);
                //XmlNode xn = xmlDoc.SelectSingleNode("Unit");
                //XmlNodeList xnl = xn.ChildNodes;
                XmlNodeList xnl = xmlDoc.SelectNodes("/PartReportUnit/Unit");
                for (int i = 1; i < xnl.Count; i++)
                {
                    Num = Num + 1;
                    StrData x = new StrData();
                    //XmlElement xe = (XmlElement)xn1;
                    x.seq = Num;

                    /*
                     * 0:seqInsOrder    1:seqInsOrder_L2    2:seqPartNum    3:seqPnl    4:seqBrdNum    5:seqRef    6:seqMemo    7:fsSetPos
                     */

                    //x.compPN = xe.GetAttribute("seqPartNum").ToString();
                    x.compPN = xnl[i].ChildNodes.Item(2).InnerText;

                    strErrMessage = "";
                    strErrMessage = FunPartNumberCheck(x.compPN);
                    if (strErrMessage != "PASS")
                    {
                        MessageBox.Show(strErrMessage);
                        return;
                    }

                    //str = xe.GetAttribute("seqBrdNum").ToString();
                    str = xnl[i].ChildNodes.Item(4).InnerText;

                    if (pubFunction.IsNumeric(str, "INT") == false)
                    {
                        MessageBox.Show("The SeqBoardNum:" + str + " must be numeric!");
                        return;
                    }

                    if (BrdPN[Convert.ToInt16(str)] != "")
                    {
                        x.Jobpn = BrdPN[Convert.ToInt16(str)];
                        x.Rev = BrdRev[Convert.ToInt16(str)];
                    }
                    else
                    {
                        MessageBox.Show("Can not find the SeqBoardNum mapping:" + str);
                        return;
                    }

                    if (x.Rev.Length > 5)
                    {
                        MessageBox.Show("Rev:" + x.Rev + " is too long!");
                        return;
                    }

                    //strLocation = xe.GetAttribute("seqRef").ToString();
                    strLocation = xnl[i].ChildNodes.Item(5).InnerText;

                    x.location = strLocation;

                    if (pubFunction.ConfigListGetValue("ChkMEBOM_Location") == "Y" && x.location == "")
                    {
                        MessageBox.Show("Location:" + x.location + " can not be empty!");
                        return;
                    }

                    //StrSlot = xe.GetAttribute("fsSetPos").ToString();
                    StrSlot = xnl[i].ChildNodes.Item(7).InnerText;

                    //if (pubFunction.IsNumeric(StrSlot.Replace("-", "").Replace(" ",""), "INT") == false)
                    //{
                    //    MessageBox.Show("The StrSlot:" + StrSlot + " must be numeric!");
                    //    return;
                    //}
                    string[] temp1 = StrSlot.Replace(" ", "").Split('-');
                    #region //switch
                    switch (temp1.GetUpperBound(0))
                    {
                        case 0:
                            if (NXT == true)
                            {
                                MessageBox.Show("The NXT Machine slot" + StrSlot + "format is wrong,please check!");
                                return;
                            }
                            strLR = "0";
                            if (Convert.ToInt32(temp1[0]) > 240 || Convert.ToInt32(temp1[0]) < 0)
                            {
                                MessageBox.Show("The Slot " + StrSlot + " foramt is wrong,please check!");
                                return;
                            }
                            else
                            {
                                x.Slot = temp1[0];
                                x.Machine = Machine;
                            }
                            break;
                        case 1:
                            if (Convert.ToInt32(temp1[1]) > 100 && NXT == false)
                            {
                                temp1[1] = (Convert.ToInt32(temp1[1]) - 100).ToString();
                            }
                            if (NXT == false)
                            {
                                if (Convert.ToInt32(temp1[0]) > TabQty || Convert.ToInt32(temp1[1]) > SlotQty)
                                {
                                    MessageBox.Show("The Slot " + StrSlot + " foramt is wrong,please check!");
                                    return;
                                }
                                else
                                {
                                    x.Slot = temp1[0] + "-" + temp1[1];
                                    x.Machine = Machine;
                                    strLR = "0";
                                }
                            }
                            else
                            {
                                if (temp1[0].Length == 1)
                                {
                                    x.Machine = Machine + "0" + temp1[0];
                                }
                                else
                                {
                                    x.Machine = Machine + temp1[0];
                                }
                                dt = MCC.MCC_QueryDataByType("MCC_GetMachineInfo", "", "", x.Machine, Line, Side, "", "", "", "");
                                if (dt.Rows.Count > 0)
                                {
                                    TabQty = Convert.ToInt32(dt.Rows[0]["Qty"].ToString());
                                    SlotQty = Convert.ToInt32(dt.Rows[0]["MaxSlotNum"].ToString());
                                }
                                else
                                {
                                    MessageBox.Show("Can't find The Machine:" + x.Machine + " in line:" + Line + " and side: " + Side + " ,please define it in machinetype or check its format!", "LoadArrayWithData");
                                    return;
                                }
                                if (TabQty != 1 || Convert.ToInt32(temp1[1]) > SlotQty)
                                {
                                    MessageBox.Show("The Slot " + StrSlot + " foramt is wrong,please check!", "LoadArrayWithData");
                                    return;
                                }
                                x.Slot = temp1[1];
                                strLR = "0";
                            }
                            break;
                        case 2:
                            if (Convert.ToInt32(temp1[1]) > 100 && NXT == false)
                            {
                                temp1[1] = (Convert.ToInt32(temp1[1]) - 100).ToString();
                            }
                            if (NXT == false && AIMEX == false)
                            {
                                if ((temp1[2] != "1" && temp1[2] != "2") || Convert.ToInt32(temp1[0]) > TabQty || Convert.ToInt32(temp1[1]) > SlotQty)
                                {
                                    MessageBox.Show("The Slot " + StrSlot + " foramt is wrong,please check!");
                                    return;
                                }
                                else
                                {
                                    strLR = temp1[2];
                                    x.Slot = temp1[0] + "-" + temp1[1];
                                    x.Machine = Machine;
                                }
                            }
                            else if (NXT == true || AIMEX == true)
                            {
                                if (temp1[0].Length == 1)
                                {
                                    x.Machine = Machine + "0" + temp1[0];
                                }
                                else
                                {
                                    x.Machine = Machine + temp1[0];
                                }
                                dt = MCC.MCC_QueryDataByType("MCC_GetMachineInfo", "", "", x.Machine, Line, Side, "", "", "", "");
                                if (dt.Rows.Count > 0)
                                {
                                    TabQty = Convert.ToInt32(dt.Rows[0]["Qty"].ToString());
                                    SlotQty = Convert.ToInt32(dt.Rows[0]["MaxSlotNum"].ToString());
                                }
                                else
                                {
                                    MessageBox.Show("Can't find The Machine:" + x.Machine + " in line:" + Line + " and side: " + Side + " ,please define it in machinetype or check its format!", "LoadArrayWithData");
                                    return;
                                }
                                if (NXT == true)
                                {
                                    if ((temp1[2] != "1" && temp1[2] != "2") || TabQty != 1 || Convert.ToInt32(temp1[1]) > SlotQty)
                                    {
                                        MessageBox.Show("The Slot " + StrSlot + " foramt is wrong,please check!", "LoadArrayWithData");
                                        return;
                                    }
                                    x.Slot = temp1[1];
                                    strLR = temp1[2];
                                }
                                else
                                {
                                    if (TabQty != 2 || Convert.ToInt32(temp1[1]) > SlotQty)
                                    {
                                        MessageBox.Show("The Slot " + StrSlot + " foramt is wrong,please check!", "LoadArrayWithData");
                                        return;
                                    }
                                    x.Slot = temp1[1] + "-" + temp1[2];
                                    strLR = "0";
                                }
                            }
                            break;
                        case 3:
                            if (Parameter.BU == "NB4" || Parameter.BU == "NB7" || Parameter.BU == "ESBU")
                            {
                                if (Convert.ToInt32(temp1[2]) > 100 && NXT == false)
                                {
                                    temp1[2] = (Convert.ToInt32(temp1[2]) - 100).ToString();
                                }
                                if (NXT == false)
                                {
                                    if ((temp1[3] != "1" && temp1[3] != "2") || Convert.ToInt32(temp1[0]) > TabQty || Convert.ToInt32(temp1[2]) > SlotQty)
                                    {
                                        MessageBox.Show("The Slot " + StrSlot + " foramt is wrong,please check!");
                                        return;
                                    }
                                    else
                                    {
                                        strLR = temp1[3];
                                        x.Slot = temp1[0] + "-" + temp1[2];
                                        x.Machine = Machine;

                                    }
                                }
                                else
                                {
                                    if (temp1[0].Length == 1)
                                    {
                                        x.Machine = Machine + "0" + temp1[0];
                                    }
                                    else
                                    {
                                        x.Machine = Machine + temp1[0];
                                    }
                                    dt = MCC.MCC_QueryDataByType("MCC_GetMachineInfo", "", "", x.Machine, Line, Side, "", "", "", "");
                                    if (dt.Rows.Count > 0)
                                    {
                                        TabQty = Convert.ToInt32(dt.Rows[0]["Qty"].ToString());
                                        SlotQty = Convert.ToInt32(dt.Rows[0]["MaxSlotNum"].ToString());
                                    }
                                    else
                                    {
                                        MessageBox.Show("Can't find The Machine:" + x.Machine + " in line:" + Line + " and side: " + Side + " ,please define it in machinetype or check its format!", "LoadArrayWithData");
                                        return;
                                    }
                                    if ((temp1[3] != "1" && temp1[3] != "2") || TabQty != 1 || Convert.ToInt32(temp1[2]) > SlotQty)
                                    {
                                        MessageBox.Show("The Slot " + StrSlot + " foramt is wrong,please check!", "LoadArrayWithData");
                                        return;
                                    }
                                    x.Slot = temp1[2];
                                    strLR = temp1[3];
                                }
                            }
                            break;
                        case 4:
                            if (AIMEX == true)
                            {
                                if (temp1[0].Length == 1)
                                {
                                    x.Machine = Machine + "0" + temp1[0];
                                }
                                else
                                {
                                    x.Machine = Machine + temp1[0];
                                }
                                dt = MCC.MCC_QueryDataByType("MCC_GetMachineInfo", "", "", x.Machine, Line, Side, "", "", "", "");
                                if (dt.Rows.Count > 0)
                                {
                                    TabQty = Convert.ToInt32(dt.Rows[0]["Qty"].ToString());
                                    SlotQty = Convert.ToInt32(dt.Rows[0]["MaxSlotNum"].ToString());
                                }
                                else
                                {
                                    MessageBox.Show("Can't find The Machine:" + x.Machine + " in line:" + Line + " and side: " + Side + " ,please define it in machinetype or check its format!", "LoadArrayWithData");
                                    return;
                                }
                                if (TabQty != 2)
                                {
                                    MessageBox.Show("The Slot " + StrSlot + " foramt is wrong,please check!", "LoadArrayWithData");
                                    return;
                                }
                                if (temp1[2] == "B")
                                {
                                    temp1[3] = (Convert.ToInt32(temp1[3]) + 12).ToString();
                                }
                                x.Slot = temp1[1] + "-" + temp1[3];
                                strLR = temp1[4];
                            }
                            break;
                        default:
                            MessageBox.Show("The Slot " + StrSlot + " foramt is wrong,please check!");
                            return;
                    }
                    #endregion

                    x.LR = strLR;
                    x.Qty = 1;
                    x.Enabled = true;
                    list.Add(x);
                }

                int m, n;
                for (m = 0; m < Num; m++)
                {
                    for (n = m + 1; n < Num; n++)
                    {
                        if (list[n].Enabled == true)
                        {
                            if (list[m].compPN == list[n].compPN && list[m].Jobpn == list[n].Jobpn && list[m].Slot == list[n].Slot && list[m].Machine == list[n].Machine)
                            {
                                list[m].Qty = list[m].Qty + 1;
                                list[m].location = list[m].location + ";" + list[n].location;
                                list[n].Enabled = false;
                            }
                        }
                    }
                }

                string preJobPN = "", strMachine = "";
                for (m = 0; m < Num; m++)
                {
                    if (list[m].Enabled == true)
                    {
                        if (list[m].Jobpn != preJobPN || list[m].Machine != strMachine)
                        {
                            MCC.MCC_QueryDataByType("MCC_DelQSMS_MEBOM_1", "", "", jobgroup, list[m].Machine, list[m].Jobpn, list[m].Rev, BuildType, Line, Factory);
                            strMachine = list[m].Machine;
                            preJobPN = list[m].Jobpn;
                        }
                    }
                }

                for (m = 0; m < Num; m++)
                {
                    if (list[m].Enabled == true)
                    {
                        MCC.InsertIntoQSMS_MEBom(list[m].Machine, list[m].Jobpn, jobgroup, list[m].Rev, list[m].compPN, list[m].LR, list[m].Slot, list[m].Qty.ToString(),
                            BuildType, Side, Parameter.g_userName, Factory, Line, list[m].location);
                    }
                }

                MCC.MCC_QueryDataByType("MCC_SaveQSMS_Log", "", "", FileName, Parameter.g_userName, "", "", "", "", "");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "LoadArrayWithData");
                return;
            }
        }

    }

    class StrData
    {
        public int seq;
        public string compPN;
        public string Jobpn;
        public string Rev;
        public string location;
        public string Slot;
        public string Machine;
        public string LR;
        public int Qty;
        public bool Enabled;
    }

}
