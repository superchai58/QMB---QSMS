using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using System.Collections;
using System.Threading;

namespace QSMS.QSMS.MCC
{
    public partial class frmTransferPanaAMI : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.MCC.TransferPanaAMI TransferPanaAMI = new DbLibrary.MCC.TransferPanaAMI();
        int I;
        int Interval;
        string strJobPN = "", strRev = "", strLine = "", strBuildType="";
        string[] MC_Header;
        string[] PD_Header;
        string[] PT_Header;
        string[] PL_Header;
        string[] NZ_Header;
        string[] SD_Header;
        string[] BA_Header;
        string[] BD_Header;
        public class MachineType
        {
            public string MCName = "";
            public string MCNo = "";
            public string HeadNo = "";
            public int Table=0;
        }

        public class NozzleType
        {
            public string Machine = "";
            public string HeadNo = "";
            public string tmpSlot = "";
            public string location = "";
            public string Nozzle = "";
        }


        public class PartData
        {
            public int IdNo=0;
            public string PN = "";
            public string location = "";
            public string Skip = "";
        }



        public class PositionData
        {
            public string Machine = "";
            public string BrdPN = "";
            public string Rev = "";
            public string PU = "";
            public string Table = "";
            public string TraySlot = "";
            public string Slot = "";
            public int Side=0 ;
            public long Head=0 ;
            public int Parts =0;
            public string  compPN = "";
            public int Qty=0;
            public bool Enabled=false ;
            public bool FstMachinePNRev=false;
            public string location = "";
            public string NPMReelWidth = "";
            public string DualLaneMode="" ;
            public string B = "";
            public string F = "";
        }



        public class StockData
        {
            public string PU = "";
            public string PA = "";
            public string PB = "";
            public string TA = "";
            public string TB = "";
        }
        public class BlockAttribute
        {
            public string IDNUM = "";
            public string B = "";
            
        }    
      
        public frmTransferPanaAMI()
        {
            InitializeComponent();
        }

        private void frmTransferPanaAMI_Load(object sender, EventArgs e)

        {
            toolStripStatusLabel1.Text = "QSMS will auto check bom if you upload the PanaMAI machine bom!";
            chkAutChkBom.Checked = true;
         }

        private void cmdSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "请选择要上传的文件";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtFile.Text = dialog.FileName;
            }
        }

        private void cmdGetMEBom_Click(object sender, EventArgs e)
        {
            if(txtFile.Text=="")
            {
                MessageBox.Show("You must select a file!!");
                return;
            }
            if (LoadDataFile(txtFile.Text.Trim()) == false)
            {
                MessageBox.Show("Fail");

                LabelRun.BackColor = Color.Red;
                LabelRun.Text = "Fail";
                toolStripStatusLabel1.Text = txtFile.Text + " Upload Fail";
                toolStripStatusLabel3.Text = "Finished DateTime:" + DateTime.Now.ToString();
                return;
            }
            else
            {
                if(chkAutChkBom.Checked==true)
                {
                    AutoCheckBom();
                }                
                    MessageBox.Show("Finish");               

            }
            LabelRun.BackColor = Color.Aqua;
            LabelRun.Text = "OK";
            toolStripStatusLabel1.Text = txtFile.Text + " Upload OK";
            toolStripStatusLabel3.Text = "Finished DateTime:" + DateTime.Now.ToString();
        }
        private bool LoadDataFile(string strFile)
        {
            int Total_Qty = 0,Insert_Qty=0, Update_Qty=0;
            string BomFile = "", strBomFileName="" , strFullJobGroup = "", strErrMessage="";
            string[] Arry, Arry2, tempBomFileName, temp;
            bool LoadDataFile = true;
            Arry = strFile.Split('\\');
            BomFile = Arry[Arry.GetUpperBound(0)];
            temp = BomFile.Split('-');
            if (pubFunction.ConfigListGetValue("ChkEQProgram")=="Y")
            {
                strBomFileName = BomFile.Substring(0,BomFile.IndexOf("."));
                tempBomFileName = strBomFileName.Split('-');
                strFullJobGroup = tempBomFileName[2].Trim() + "-" + tempBomFileName[3].Trim();
                if(tempBomFileName.GetUpperBound(0)>=6)
                {
                    for (int intBomFileName = 6; intBomFileName < tempBomFileName.GetUpperBound(0); intBomFileName++)
                    {
                        strFullJobGroup = strFullJobGroup + "-" + tempBomFileName[intBomFileName].Trim();
                    }
                }
                temp = strBomFileName.Substring(1, 26).Split('-');
            }
          if(temp.Count() != 5 && temp.Count() != 6)
            {
                MessageBox.Show("Filename format must be Factory-Line-PN-Rev-BuildType-Side !");
                return LoadDataFile;
            }
            if (temp[4]!= "1" && temp[4] != "2" && temp[4] != "3" && temp[4] != "4")
            {
                MessageBox.Show("BuildType must be 1,2,3 or 4.");
                return LoadDataFile;
            }
            if(temp[5].Substring(0,1)!="S" && temp[5].Substring(0, 1) != "C" && temp[5].Substring(0, 1) != "Q")
            {
                MessageBox.Show("Side must be S,C or Q.");
                return LoadDataFile;
            }
            if (temp[3].Length != 3 && temp[3].Length!=2 && temp[3].Length != 0)
            {
                MessageBox.Show("The version length must be 2 or 3 or 0!");
                return LoadDataFile;
            }
            strErrMessage = "";
            strErrMessage = FunPartNumberCheck(temp[2].Trim());
            if(strErrMessage!="PASS")
            {
                MessageBox.Show(strErrMessage);
                return LoadDataFile;
            }
            TransferPanaAMI.insertQSMS_Log(BomFile.Trim().Substring(0, BomFile.Length),Parameter.UID);
            int NFile = 0;
            bool IsTableHeader=false;
            string StartTime = "", strCurrent="", log="", Flg="", BoardType="", PCBSize="";
            string Factory = "", MBPN = "", Line = "", Revision = "", BuildType = "", strSide = "", jobgroup = "", strJobPn = "",  MCType = "";
            int idxMC = 0, idxPD = 0, idxPT = 0, idxPL = 0, idxNZ = 0, idxSD = 0, TraySlot = 0, idxBA = 0;
            //ArrayList<MachineType> MC = new ArrayList();
            List<MachineType> MC = new List<MachineType> { };            
            List<PositionData> PD = new List<PositionData> { };
            List<PartData> pt = new List<PartData> { };
            List<NozzleType> NZ = new List<NozzleType> { };
            List<StockData> SD = new List<StockData> { };
            List<BlockAttribute> BA = new List<BlockAttribute> { };               
            string[,] PL =new string[0,0];      
                       
            Factory = temp[0];
            MBPN= temp[2];
            Line = temp[1];
            Revision = temp[3];
            BuildType = temp[4];
            strSide = temp[5].Trim().Substring(0, 1);

            jobgroup = MBPN + "-" + Revision;
            strJobPN = MBPN;
            strRev = Revision;
            strBuildType = BuildType;

            MCType = "";
            if((BuildType=="2" || BuildType=="4") && strSide != "S")
            {
                MessageBox.Show("The side is " + strSide + ",BuildType is 2 or 4,they are not match,the side must be S side.");
                return LoadDataFile;
            }
            if (BuildType == "3" && strSide != "C")
            {
                MessageBox.Show("The side is " + strSide + ",BuildType is 3,they are not match,the side must be C side.");
                return LoadDataFile;
            }
            NFile = 1;
            StartTime = DateTime.Now.ToString();
            LabelRun.BackColor = Color.Red;
            LabelRun.Text = "Runing...";
            toolStripStatusLabel1.Text = "GetMEBom_" + BomFile;
            toolStripStatusLabel2.Text = "Start DateTime:" + StartTime;
            System.IO.StreamReader file = new System.IO.StreamReader(strFile);
            //string readall = file.ReadLine();
            int m = 1;
            while(m!=0)
            {
                strCurrent = file.ReadLine();
                if (strCurrent == null)
                {
                    m = 0;
                    break;
                }
                strCurrent = strCurrent.Replace("\r\n", "");
                strCurrent = strCurrent.Replace(Convert.ToChar(10).ToString(), "");
                strCurrent = strCurrent.Replace(Convert.ToChar(13).ToString(), "");
                strCurrent = strCurrent.Replace(Convert.ToChar(9).ToString(), " ");
                strCurrent = strCurrent.Replace("\"", "");

               
                if(strCurrent== "[Index]")
                {
                    log = "ID";
                }
                else if(strCurrent == "[Machines]")
                {
                    log = "MC";
                    IsTableHeader = true;
                }
                else if(strCurrent == "[PositionData]")
                {
                    log = "PD";
                    IsTableHeader = true;
                }
                else if(strCurrent == "[PartsData]")
                {
                    log = "PT";
                    IsTableHeader = true;
                }
                else if(strCurrent == "[PartsLIB]")
                {
                    log = "PL";
                    IsTableHeader = true;
                }
                else if(strCurrent == "[NozzleStock]")
                {
                    log = "NZ";
                    IsTableHeader = true;
                }
                else if(strCurrent == "[StockData]")
                {
                    log = "SD";
                    IsTableHeader = true;
                }
                else if(strCurrent == "[BlockAttribute]")
                {
                    log = "BA";
                    IsTableHeader = true;
                }
                else if (strCurrent == "[BoardData]")
                {
                    log = "BD";
                    IsTableHeader = true;
                }
                else
                {
                    if(IsTableHeader== true)
                    {
                        PhraseHeader(strCurrent, log);
                        IsTableHeader = false;
                    }
                    else
                    {
                        Arry = strCurrent.Split(' ');
                        if(strCurrent!="" &&(Arry.GetUpperBound(0) > 3 ||(Arry.GetUpperBound(0) == 2 && log== "NZ")|| (Arry.GetUpperBound(0) == 0 && (log == "ID" || log=="BD"))))
                        {
                            Flg = "OK";
                        }
                        else
                        {
                            log = "";
                            Flg = "Cancel";
                        }
                        if ((log + Flg) == "IDOK")
                        {
                            if (MCType == "")
                            {
                                MCType = GetKeyValueM(strCurrent, "Machine").Trim();
                            }
                        }
                        else if ((log + Flg) == "MCOK")
                        {
                            idxMC = idxMC + 1;
                            MachineType pn = new MachineType();
                            //MC = new MachineType[idxMC];
                            int mm = GetPosition(MC_Header, "MCNAME");
                            string aa = Arry[mm].Trim();
                            pn.MCName = Arry[mm].Trim();
                            if (pn.MCName.IndexOf("NPM") >= 0)
                            {
                                pn.MCName = pn.MCName.Substring(0, pn.MCName.IndexOf("NPM") + 3);
                            }
                            pn.MCNo = Arry[GetPosition(MC_Header, "MCNo")].Trim();
                            pn.HeadNo = Arry[GetPosition(MC_Header, "HeadNo")].Trim();
                            MC.Add(pn);
                        }
                        else if ((log + Flg) == "SDOK")
                        {
                            idxSD = idxSD + 1;
                            StockData SDda = new StockData();
                            SDda.PU = Arry[GetPosition(SD_Header, "N")].Trim();
                            SDda.PA = Arry[GetPosition(SD_Header, "PA")].Trim();
                            SDda.PB = Arry[GetPosition(SD_Header, "PB")].Trim();
                            SDda.TA = Arry[GetPosition(SD_Header, "TA")].Trim();
                            SDda.TB = Arry[GetPosition(SD_Header, "TB")].Trim();
                            SD.Add(SDda);
                        }
                        else if ((log + Flg) == "PDOK")
                        {
                            idxPD = idxPD + 1;
                            PositionData PdDa = new PositionData();

                            PdDa.Parts = Convert.ToInt16(Arry[GetPosition(PD_Header, "PARTS")].Trim());
                            PdDa.PU = Arry[GetPosition(PD_Header, "PU")].Trim();
                            PdDa.location = Arry[GetPosition(PD_Header, "C")].Trim();
                            PdDa.B = Arry[GetPosition(PD_Header, "B")].Trim();
                            PdDa.F = Arry[GetPosition(PD_Header, "F")].Trim();
                            if (Convert.ToInt64(Arry[GetPosition(PD_Header, "SIDE")]) > 2)
                            {
                                MessageBox.Show("Side (LR) wrong : " + Arry[GetPosition(PD_Header, "SIDE")] + ", it must be 0 or 1 or 2!");
                                return false;
                            }
                            PdDa.Side = Convert.ToInt16(Arry[GetPosition(PD_Header, "SIDE")].Trim());
                            PdDa.Head = Convert.ToInt16(Arry[GetPosition(PD_Header, "HEAD")].Trim());
                            PdDa.Qty = 1;
                            if (PdDa.PU.Length >= 4)
                            {
                                PdDa.Enabled = true;
                            }
                            BoardType = Arry[GetPosition(PD_Header, "C")].Trim();
                            Arry2 = BoardType.Split('-');
                            if (Arry2.GetUpperBound(0) > 0)
                            {
                                int aa = Arry2.GetUpperBound(0);
                                if (Arry2.GetUpperBound(0) != 2)
                                {
                                    MessageBox.Show("BoardType wrong : " + BoardType + ", must be PN-REV!");
                                    return false;
                                }
                                if (Arry2[2].Replace("\"\"", "").Length != 3 && Arry2[2].Replace("\"\"", "").Length != 2)
                                {
                                    MessageBox.Show("Version wrong : " + BoardType + ", the version " + Arry2[2].Replace("\"\"", "") + "length must be 2 or 3!");
                                    return false;
                                }
                                strErrMessage = "";
                                strErrMessage = FunPartNumberCheck(Arry2[1].Replace("\"\"", ""));
                                if (strErrMessage != "PASS")
                                {
                                    MessageBox.Show(strErrMessage);
                                    return false;
                                }
                                PdDa.BrdPN = Arry2[1];
                                PdDa.Rev = Arry2[2].Replace("\"\"", "");
                                
                            }
                            else
                            {
                                PdDa.BrdPN = MBPN;
                                PdDa.Rev = Revision;

                            }
                            PdDa.FstMachinePNRev = true;
                            PD.Add(PdDa);
                        }
                        else if ((log + Flg) == "PTOK")
                        {
                            idxPT = Convert.ToInt16(Arry[GetPosition(PT_Header, "IDNUM")]);
                            //pt[idxPT].IdNo = Convert.ToInt16(Arry[GetPosition(PT_Header, "IDNUM")].Trim());
                            //pt[idxPT].PN = Arry[GetPosition(PT_Header, "NAME")].Trim();
                            //pt[idxPT].Skip = Arry[GetPosition(PT_Header, "SKIP")].Trim();
                            PartData ptda = new PartData();
                            while (idxPT != pt.Count() + 1)
                            {
                                ptda = new PartData();
                                ptda.IdNo = 0;
                                ptda.PN = "";
                                ptda.Skip = "";

                                pt.Add(ptda);
                            }
                                ptda = new PartData();
                            //PartData ptda = new PartData();
                                ptda.IdNo = Convert.ToInt16(Arry[GetPosition(PT_Header, "IDNUM")].Trim());
                                ptda.PN = Arry[GetPosition(PT_Header, "NAME")].Trim();
                                ptda.Skip = Arry[GetPosition(PT_Header, "SKIP")].Trim();
                            
                            pt.Add(ptda);
                        }
                        else if ((log + Flg) == "PLOK")
                        {
                            idxPL = idxPL + 1;
                            PL = new string [2,idxPT];
                            PL[0, idxPL] = Arry[GetPosition(PL_Header, "PartsName")].Trim().Replace("\"\"", "");
                            PL[1, idxPL] = Arry[GetPosition(PL_Header, "ReelWidth")].Trim().Replace("\"\"", "");  
                        }
                        else if ((log + Flg) == "NZOK")
                        {
                            idxNZ =Convert.ToInt16 (Arry[GetPosition(NZ_Header,"IDNUM")]);
                            NozzleType NZda = new NozzleType();
                            NZda.HeadNo = Arry[GetPosition(NZ_Header, "N")].Trim().Substring(0, Arry[GetPosition(NZ_Header, "N")].Trim().Length-2);
                            NZda.tmpSlot = Arry[GetPosition(NZ_Header, "N")].Trim().Substring(Arry[GetPosition(NZ_Header, "N")].Trim().Length - 2, 2);
                            NZda.Nozzle = Arry[GetPosition(NZ_Header, "P")].Trim();
                            NZ.Add(NZda);
                        }
                        else if ((log + Flg) == "BDOK")
                        {
                           if(PCBSize=="")
                            {
                                PCBSize = GetKeyValueM(BD_Header[0], "X");
                            }
                        }
                        else if ((log + Flg) == "BAOK")
                        {
                            if (Arry[GetPosition(BA_Header, "B")].Trim()=="1")
                            {
                                idxBA = idxBA + 1;
                                BlockAttribute BADa=new BlockAttribute();
                                BADa.IDNUM = Arry[GetPosition(BA_Header, "IDNUM")].ToString();
                                BADa.B = Arry[GetPosition(BA_Header, "B")].ToString();
                                BA.Add(BADa);
                            }
                        }
                    }
                }
            }
            DataTable dt = new DataTable();
            string MCSeq = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string PreMCNo = "0", ReelWidth="";
            for(int i=0;i<MC.Count(); i++)
            {
                if(MC[i].MCNo==PreMCNo)
                {
                    MC[i].Table = MC[i - 1].Table + 1;
                }
                else
                {
                    MC[i].Table = 1;
                    PreMCNo = MC[i].MCNo;
                }
                for(int j=0;j<NZ.Count();j++)
                {
                    if(NZ[j].HeadNo==MC[i].HeadNo)
                    {
                        NZ[j].Machine = MC[i].MCName + MCSeq.Substring(Convert.ToInt16(MC[Convert.ToInt16(MC[i].HeadNo)-1].MCNo)-1, 1);
                        NZ[j].location = MC[i].Table + "-" + NZ[j].tmpSlot;
                    }
                }

            }
          
            if (MCType.ToUpper()== pubFunction.ConfigListGetValue("NPMMACHINETYPE"))
            {
                string aa = pubFunction.ConfigListGetValue("NPMMACHINETYPE");
                for(int i=0;i<SD.Count();i++)
                {
                    for(int j=0;j<PD.Count();j++)
                    {
                        if(SD[i].PA==PD[j].Parts.ToString() && SD[i].PU== PD[j].PU)
                        {
                            PD[j].NPMReelWidth = SD[i].TA;
                        }
                        else if(SD[i].PB==PD[j].Parts.ToString() && SD[i].PU==PD[j].PU)
                        {
                            PD[j].NPMReelWidth = SD[i].TB;
                        }
                    }
                }
            }
            for(int i=0;i<PD.Count();i++)
            {
                for(int j=0;j<BA.Count();j++)
                {
                    if(Convert.ToInt64(BA[j].B)==1 && BA[j].IDNUM==PD[i].B)
                    {
                        PD[i].Enabled = false;
                    }
                }
                for(int j=0;j<pt.Count();j++)
                {
                    if(pt[j].Skip=="")
                    {
                        pt[j].Skip = "0";
                    }
                    if(pt[j].Skip.ToString()=="1" && pt[j].IdNo==PD[i].Parts)
                    {
                        PD[i].Enabled = false;
                    }
                }
                if(Convert.ToInt64(PD[i].F)==2)
                {
                    PD[i].Enabled = false;
                }
                int a = (int)Conversion.Val(PD[i].PU.Substring(0,1));
                if(a>0)
                {
                    PD[i].compPN = pt[Convert.ToInt16(PD[i].Parts)-1].PN.Replace("\"\"", "");
                    if(PD[i].compPN.Trim().Length!=11 && PD[i].compPN.Trim().Length != 14)
                    {
                        int n = 0;
                        if(PD[i].compPN.Trim().IndexOf("-")>0)
                        {
                            n = PD[i].compPN.Trim().IndexOf("-")-1;
                        }
                        else
                        {
                            n = PD[i].compPN.Trim().Length;
                        }
                        PD[i].compPN = PD[i].compPN.Trim().Substring(0, n);
                    }
                    strErrMessage = "";
                    strErrMessage = FunPartNumberCheck(PD[i].compPN);
                    if(strErrMessage!="PASS")
                    {
                        MessageBox.Show(strErrMessage);
                        return false;
                    }
                    string qq = MC[Convert.ToInt16(PD[i].Head)-1].MCNo;
                    int mm = Convert.ToInt16(MC[Convert.ToInt16(PD[i].Head) - 1].MCNo) - 1;
                    PD[i].Machine=MC[Convert.ToInt16(PD[i].Head)-1].MCName+MCSeq.Substring(Convert.ToInt16(MC[Convert.ToInt16(PD[i].Head)-1].MCNo) -1,1);
                    if(PD[i].Machine.IndexOf("NPM")>=0 && temp.GetUpperBound(0)==6)
                    {
                        if(temp[6].Trim().Substring(0,1)=="F" || temp[6].Trim().Substring(0,1)=="R")
                        {
                            PD[i].DualLaneMode = "Mix";
                        }
                    }
                    int ii = Convert.ToInt16(Conversion.Val(PD[i].PU.Substring(0, PD[i].PU.Length - 4)));
                    PD[i].Table = MC[Convert.ToInt16(Conversion.Val(PD[i].PU.Substring(0, PD[i].PU.Length - 4)))-1].Table.ToString();
                    PD[i].Slot = PD[i].Table + "-" + Conversion.Val(PD[i].PU.Substring(PD[i].PU.Length - 2, 2)).ToString();
                    PD[i].TraySlot = Conversion.Val(PD[i].PU.Trim().Substring(0, PD[i].PU.Length - 2).Substring(PD[i].PU.Trim().Substring(0, PD[i].PU.Length - 2).Length - 2, 2)).ToString();
                    if(Parameter.Settings.UpdateJobSide.ToUpper()=="Y")
                    {
                        dt = TransferPanaAMI.QSMS_JobSide(PD[i].BrdPN);
                        if(dt.Rows.Count==0)
                        {
                            MessageBox.Show(PD[i].BrdPN.Trim()+ ":Can't find the job side by the JobPN,please check!");
                            return false;
                        }
                        else
                        {
                            if(dt.Rows[0]["Side"].ToString().ToUpper()!="S" && dt.Rows[0]["Side"].ToString().ToUpper() != "C")
                            {
                                MessageBox.Show(dt.Rows[0]["Side"].ToString().ToUpper()+ ":Job side's format is wrong ,the side must be S or C,please define it afresh!");
                                return false;
                            }
                        }
                    }
                }
            }
            for(int i=0;i<PD.Count();i++)
            {
                if (PD[i].Enabled == true)
                {
                    for (int j = i + 1; j < PD.Count(); j++)
                    {
                        if (PD[j].Enabled == true && PD[i].BrdPN == PD[j].BrdPN && PD[i].compPN == PD[j].compPN &&
                            PD[i].Machine == PD[j].Machine && PD[i].Slot == PD[j].Slot && PD[i].Side == PD[j].Side)
                        {
                            PD[j].Enabled = false;
                            PD[i].Qty = PD[i].Qty + 1;
                            PD[i].location = PD[i].location + ";" + PD[j].location;
                        }
                        if (PD[i].DualLaneMode.ToUpper() == "MIX")
                        {
                            if (PD[j].FstMachinePNRev == true && PD[j].Machine == PD[i].Machine && PD[i].BrdPN == PD[j].BrdPN && PD[i].Table == PD[j].Table)
                            {
                                PD[j].FstMachinePNRev = false;
                            }
                        }
                        else
                        {
                            if (PD[j].FstMachinePNRev == true && PD[j].Machine == PD[i].Machine && PD[i].BrdPN == PD[j].BrdPN)
                            {
                                PD[j].FstMachinePNRev = false;
                            }
                        }
                    }
                    PD[i].location = PD[i].location.Replace("\"\"", "");
                    if (PD[i].FstMachinePNRev == true)
                    {
                        if (CheckMachine(Line, PD[i].Machine, strSide) == false)
                        {
                            return false;
                        }
                        if (PD[i].DualLaneMode.ToUpper() == "MIX")
                        {
                            TransferPanaAMI.deleteMIx(pubFunction.ConfigListGetValue("ChkEQProgram"), strFullJobGroup, jobgroup, PD[i].Machine, PD[i].BrdPN, PD[i].Rev, BuildType, Factory, PD[i].Slot.Substring(0, 1), Line);
                        }
                        else
                        {
                            TransferPanaAMI.deleteData(pubFunction.ConfigListGetValue("ChkEQProgram"), strFullJobGroup, jobgroup, PD[i].Machine, PD[i].BrdPN, PD[i].Rev, BuildType, Factory, PD[i].Slot.Substring(0, 1), Line);
                        }
                    }

                    if (Parameter.Settings.UpdateJobSide.ToUpper() == "Y")
                    {
                        dt = TransferPanaAMI.QSMS_JobSide(PD[I].BrdPN);
                        if (dt.Rows.Count == 0)
                        {
                            MessageBox.Show(PD[i].BrdPN.Trim() + ":Can't find the job side by the JobPN,please check!");
                            return false;
                        }
                        else
                        {
                            if (dt.Rows[0]["Side"].ToString().ToUpper() != "S" && dt.Rows[0]["Side"].ToString().ToUpper() != "C")
                            {
                                MessageBox.Show(dt.Rows[0]["Side"].ToString().ToUpper() + ":Job side's format is wrong ,the side must be S or C,please define it afresh!");
                                return false;
                            }
                            else
                            {
                                strSide = dt.Rows[0]["Side"].ToString();
                            }
                        }
                    }
                    if (pubFunction.ConfigListGetValue("ChkMEBOM_Location") == "Y" && PD[i].location == "")
                    {
                        MessageBox.Show(PD[i].location + ":location can not be empty,please check");
                        return false;
                    }
                    strLine = Line;
                    if(idxPL==0)
                    {
                        if(MCType.ToUpper()== pubFunction.ConfigListGetValue("NPMMachineType"))
                        {
                            ReelWidth = PD[i].NPMReelWidth;
                        }
                        else
                        {
                            ReelWidth = "";
                        }
                    }
                    else
                    {
                        if (MCType.ToUpper() == pubFunction.ConfigListGetValue("NPMMachineType"))
                        {
                            ReelWidth = PD[i].NPMReelWidth;
                        }
                        else
                        {
                            if(MCType.ToUpper() == pubFunction.ConfigListGetValue("NPMMachineType"))
                            {
                                ReelWidth = PD[i].NPMReelWidth;
                            }
                            else
                            {
                                if(inArray2(PL,PD[i].compPN.Replace("_","%"),0)>0)
                                {
                                    ReelWidth = PL[1, inArray2(PL, PD[i].compPN.Replace("_", "%"), 0)];
                                }
                                else
                                {
                                    ReelWidth = "";
                                }
                            }
                        }
                    }
                    
                    TransferPanaAMI.insertQSMS_MEBom(pubFunction.ConfigListGetValue("ChkEQProgram"), PD[i].Machine,PD[i].BrdPN, jobgroup,
                        PD[i].Rev,PD[i].compPN.Replace("_","%"),PD[i].Side,PD[i].Slot,PD[i].Qty,BuildType,
                        strSide,Parameter.UID, Factory,Line,ReelWidth,PD[i].location,PD[i].DualLaneMode, strFullJobGroup);
                    Insert_Qty = Insert_Qty + 1;
                   
                }
                Total_Qty = Total_Qty + 1;
            }
            TransferPanaAMI.DelNozzleLocation(Factory, Line, strSide, BuildType, jobgroup);
            for(int i=0;i<NZ.Count;i++)
            {
                TransferPanaAMI.InsertNozzleLocation(Factory,Line,strSide,BuildType, jobgroup,NZ[i].Machine,NZ[i].location,NZ[i].Nozzle,Parameter.UID);
            }
            TransferPanaAMI.insertQSMS_LOG("SMT_QSMS_PCBSize",jobgroup,PCBSize,Parameter.UID);
            Thread.Sleep(1000);
            TransferPanaAMI.insertQSMS_LOG("SMT_QSMS", "Load_PanaAMI End", BomFile.Trim(), Parameter.UID);
            MessageBox.Show("*** Load  finish ! ***"+ "   " +"\r\n"+"Total Counter : " + Total_Qty+"Insert succeed : " + Insert_Qty +"\r\n"+ "Update succeed : " + Update_Qty );
            file.Close();

            return LoadDataFile;

        }
        public int inArray2(string [,] Arry,string str,int Compare)
        {
            int inArray2 = -1;
            for(int i= Arry.GetLowerBound(2); i<Arry.GetUpperBound(2);i++)
            {
                if(Arry[0,i].ToUpper()==str.ToUpper() || Arry[1,i].ToUpper()==str.ToUpper())
                {
                    inArray2 = i;
                    return inArray2;
                }
                else
                {
                    if(Arry[0,i].Trim()==str || Arry[1,i].Trim()==str)
                    {
                        inArray2 = i;
                        return inArray2;
                    }
                }
            }
            return inArray2;
        }
        public bool CheckMachine(string line,string machine,string strside)
        {
            DataTable dt = new DataTable();
            if(machine.Substring(0,1)!="*")
            {
                dt = TransferPanaAMI.Getmachine(machine, line, strside);
                if(dt.Rows.Count==0)
                {
                    MessageBox.Show("The Machine:" + machine + " in line:" + line + " and side: " + strside + " (you uploaded) was not defined in machine,please check it in machinetype");
                    return false;
                }
            }
            return true;
        }
        public int GetPosition(string[] src, string key)
        {
            int i = 0;
            for(i = 0; i < src.Count(); i++)
            {
                if(src[i].ToUpper()==key.ToUpper())
                {
                    return i;
                }
            }
            MessageBox.Show("GetPosition: can not find key:" + key);
            return i;
        }
        public string GetKeyValueM(string src,string key)
        {
            string  GetKeyValueM = "", oneLine = "";
            string[] aLine;
            int pos1 = 0;
            aLine = src.Split('\r');
            key = key.ToUpper();
            int i=0;
            for (i=0;i<aLine.Count();i++)
            {
                int m = aLine[i].ToUpper().IndexOf("A");
                string mm = aLine[i].ToUpper();
                if (aLine[i].ToUpper().IndexOf(key)>=0)
                {
                    oneLine = aLine[i];
                    break;
                }                
            }
            if(i>aLine.Count())
            {
                GetKeyValueM = "";
                return GetKeyValueM;
            }
            pos1 = oneLine.ToUpper().IndexOf(key.Trim() + "=");
            if(pos1>=0)
            {
                pos1 = pos1 + (key.Trim() + "=").Length;
                GetKeyValueM = oneLine.Substring(pos1, oneLine.Length-pos1);
            }
            else
            {
                GetKeyValueM = "";
            }

            return GetKeyValueM;
        }
        public void PhraseHeader(string src, string log)
        {
            if(log=="MC")
            {
                MC_Header= src.Split(' ');
            }
            if (log == "PD")
            {
                PD_Header = src.Split(' ');
            }
            if (log == "PT")
            {
                PT_Header = src.Split(' ');
            }
            if (log == "PL")
            {
                PL_Header = src.Split(' ');
            }
            if (log == "NZ")
            {
                NZ_Header = src.Split(' ');
            }
            if (log == "SD")
            {
                SD_Header = src.Split(' ');
            }
            if (log == "BA")
            {
                BA_Header = src.Split(' ');
            }
            if (log == "BD")
            {
                BD_Header = src.Split(' ');
            }
        }

        private string FunPartNumberCheck(string PartNumber)
        {
            DataTable dtCheck = TransferPanaAMI.CheckFormat(PartNumber);
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
        private void AutoCheckBom()
        {
            DataTable dt = new DataTable();
           if(strJobPN!="" || strRev!="")
            {
                dt = TransferPanaAMI.QSMS_GetPCBWO(strJobPN, strRev, strLine, strBuildType);
                if(dt.Rows.Count>0)
                {
                    for(int i=0;i<dt.Rows.Count;i++)
                    {
                        LabelRun.BackColor = Color.Red;
                        LabelRun.Text = "AutoCheckBom Running...";
                        toolStripStatusLabel1.Text = "AutoCheckBom:" +dt.Rows[i]["WO"].ToString();
                        GetCheckBomData(dt.Rows[i]["WO"].ToString(), Parameter.UID, "N");
                    }
                }
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        public void GetCheckBomData(string Work_Order,string g_userName,string  DualModel)        
        {
            DataSet ds = new DataSet();
            ds = TransferPanaAMI.GetCheckBomData(Work_Order, g_userName, DualModel, "GetCheckBomData");
            if (ds.Tables[0].Rows.Count > 0)
            {
                MessageBox.Show(ds.Tables[0].Rows[0]["Msg"].ToString());
                if (ds.Tables[0].Rows[0]["Result"].ToString() == "0" && ds.Tables.Count > 1)
                {

                    CopyToExcel("", "", "", "", "", "", ds.Tables[1], null);
                }
            }
        }
        private static string GetColumnChar(int col)
        {
            var a = col / 26;
            var b = col % 26;

            if (a > 0) return GetColumnChar(a - 1) + (char)(b + 65);

            return ((char)(b + 65)).ToString();
        }
        public void CopyToExcel(string Type, string FileName, string Sheetname, string Line, string WO, string Machine, DataTable dt, DataSet ds)
        {
            string col1;
            int row, col;
            row = dt.Rows.Count;
            col = dt.Columns.Count - 1;
            col1 = GetColumnChar(col);
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application appExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xBk;
            if (FileName != "")
            {
                if (Sheetname == "")
                {
                    Sheetname = "1";
                }

                xBk = appExcel.Workbooks.Open(FileName);

            }
            else
            {
                if (Sheetname == "")
                {
                    Sheetname = "1";
                }

                xBk = appExcel.Workbooks.Add(true);
            }

            Microsoft.Office.Interop.Excel.Worksheet xSt;
            //xBk = appExcel.Workbooks.Add(true);
            xSt = (Microsoft.Office.Interop.Excel.Worksheet)xBk.ActiveSheet;
            appExcel.Visible = true;
            xSt = (Microsoft.Office.Interop.Excel.Worksheet)xBk.Worksheets.get_Item(Sheetname);
            if (Type == "CheckBOM_Rate")
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        xSt.Cells[i + 1, i + 1] = ds.Tables[0].Columns[i].ColumnName.ToString();
                    }
                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                    {
                        for (int m = 0; m < ds.Tables[0].Columns.Count; m++)
                        {
                            xSt.Cells[i + 2, m + 1] = ds.Tables[0].Rows[i][m].ToString();
                        }

                    }
                }
                if (ds.Tables[1].Rows.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[1].Columns.Count; i++)
                    {
                        xSt.Cells[i + 4, i + 1] = ds.Tables[1].Columns[i].ColumnName.ToString();
                    }
                    for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                    {
                        for (int m = 0; m < ds.Tables[1].Columns.Count; m++)
                        {
                            xSt.Cells[i + 5, m + 1] = ds.Tables[1].Rows[i][m].ToString();
                        }

                    }
                }
            }
            if (Type == "MaterialDifferentList")
            {
                xSt.Cells[2, 2] = Line;
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    xSt.Cells[3, i + 1] = dt.Columns[i].ColumnName.ToString();
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int m = 0; m < dt.Columns.Count; m++)
                    {
                        xSt.Cells[i + 4, m + 1] = dt.Rows[i][m].ToString();
                    }

                }
            }
            if (Type == "CopyToExcelWipLackByWo")
            {
                xSt.Cells[2, 5] = ds.Tables[0].Rows[0][0].ToString();
                for (int i = 0; i < ds.Tables[1].Columns.Count; i++)
                {
                    xSt.Cells[3, i + 1] = ds.Tables[1].Columns[i].ColumnName.ToString();
                }
                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                {
                    for (int m = 0; m < ds.Tables[1].Columns.Count; m++)
                    {
                        xSt.Cells[i + 4, m + 1] = ds.Tables[1].Rows[i][m].ToString();
                    }

                }
            }
            if (Type == "CopyToExcelWipByGroup")
            {
                xSt.Cells[2, 2] = Line;
                xSt.Cells[2, 6] = ds.Tables[0].Rows[0][0].ToString();
                for (int i = 0; i < ds.Tables[1].Columns.Count; i++)
                {
                    xSt.Cells[3, i + 1] = ds.Tables[1].Columns[i].ColumnName.ToString();
                }
                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                {
                    for (int m = 0; m < ds.Tables[1].Columns.Count; m++)
                    {
                        xSt.Cells[i + 4, m + 1] = ds.Tables[1].Rows[i][m].ToString();
                    }

                }
            }
            if (Type == "CopyToExcelWipByDate")
            {
                xSt.Cells[2, 2] = Line;
                xSt.Cells[2, 5] = ds.Tables[0].Rows[0][0].ToString();
                for (int i = 0; i < ds.Tables[1].Columns.Count; i++)
                {
                    xSt.Cells[3, i + 1] = ds.Tables[1].Columns[i].ColumnName.ToString();
                }
                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                {
                    for (int m = 0; m < ds.Tables[1].Columns.Count; m++)
                    {
                        xSt.Cells[i + 4, m + 1] = ds.Tables[1].Rows[i][m].ToString();
                    }

                }
            }
            if (Type == "CopyToExcelWipByMaterial")
            {
                xSt.Cells[2, 4] = Line;
                xSt.Cells[2, 6] = ds.Tables[0].Rows[0][0].ToString();
                for (int i = 0; i < ds.Tables[1].Columns.Count; i++)
                {
                    xSt.Cells[3, i + 1] = ds.Tables[1].Columns[i].ColumnName.ToString();
                }
                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                {
                    for (int m = 0; m < ds.Tables[1].Columns.Count; m++)
                    {
                        xSt.Cells[i + 4, m + 1] = ds.Tables[1].Rows[i][m].ToString();
                    }

                }

            }
            if (Type == "PrepariMaterialList")
            {
                xSt.Cells[2, 3] = Line;
                xSt.Cells[2, 5] = WO;
                xSt.Cells[2, 7] = Machine;
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    xSt.Cells[3, i + 1] = dt.Columns[i].ColumnName.ToString();
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int m = 0; m < dt.Columns.Count; m++)
                    {
                        xSt.Cells[i + 4, m + 1] = dt.Rows[i][m].ToString();
                    }

                }
            }
            if (Type == "QSMS_DID_ToWH")
            {
                xSt.Cells[1, 1] = "Today:" + DateTime.Now.ToString("MM/DD");// Format(Now, "MM/DD");
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    xSt.Cells[2, i + 1] = dt.Columns[i].ColumnName.ToString();
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int m = 0; i < dt.Columns.Count; i++)
                    {
                        xSt.Cells[3, i + 1] = dt.Rows[i][m].ToString();
                    }

                }
            }
            if (Type == "")
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    xSt.Cells[1, i + 1] = dt.Columns[i].ColumnName.ToString();
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int m = 0; i < dt.Columns.Count; i++)
                    {
                        xSt.Cells[1, i + 1] = dt.Rows[i][m].ToString();
                    }

                }
            }

            string col2 = col1 + Convert.ToString(row + 1);
            Microsoft.Office.Interop.Excel.Range range1 = xSt.Range["A1", col2];
            range1.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range1.EntireColumn.AutoFit();
            range1.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;//水平居中  
            range1.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;//垂直居中


        }

        private void frmTransferPanaAMI_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmTransferPanaAMI");
        }

    }
}
