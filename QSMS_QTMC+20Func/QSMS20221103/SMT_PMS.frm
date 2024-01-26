VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form mdiMain 
   Caption         =   "SMT QSMS (2023/10/24)"
   ClientHeight    =   5670
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8070
   Icon            =   "SMT_PMS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   6120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label LabFac 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   7335
   End
   Begin VB.Menu MnuMCC 
      Caption         =   "MCC"
      Begin VB.Menu mnuMCCPreMaterial 
         Caption         =   "MCCPrepMaterial"
      End
      Begin VB.Menu mnuQueryByReturnedDID 
         Caption         =   "QueryByReturnedDID"
      End
      Begin VB.Menu mnuInheritDIDByWO 
         Caption         =   "InheritDIDByWO"
      End
      Begin VB.Menu mnuDispatchDIDAdditionnal 
         Caption         =   "DispatchDIDAdditionnal"
      End
      Begin VB.Menu mnuTransferDispatchDID 
         Caption         =   "TransferDispatchDID"
      End
      Begin VB.Menu MnuUpLoadBom 
         Caption         =   "UpLoadBom"
      End
      Begin VB.Menu MnuUpLoadMachineType 
         Caption         =   "UpLoadMachineType"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu menuAutoDispatch 
         Caption         =   "MaintainDIDAutoDispatch"
      End
      Begin VB.Menu mnuMaintainDID 
         Caption         =   "MaintainDID"
      End
      Begin VB.Menu mnuModfyDIDTotalQty 
         Caption         =   "ModifyDIDTotalQty"
      End
      Begin VB.Menu mnuReturnDID 
         Caption         =   "ReturnDID"
      End
      Begin VB.Menu mnuReturnComp 
         Caption         =   "ReturnComp"
      End
      Begin VB.Menu mnuReturnDIDALL 
         Caption         =   "ReturnDIDALL"
      End
      Begin VB.Menu mnuDIDCallBack 
         Caption         =   "DIDCallBack"
      End
      Begin VB.Menu mnuDIDChkStock 
         Caption         =   "DIDCheckStock"
      End
      Begin VB.Menu mnuTransferFujiXML 
         Caption         =   "TransferFujiXML"
      End
      Begin VB.Menu mnuTransferFujiNexim 
         Caption         =   "TransferFujiNexim"
      End
      Begin VB.Menu mnuTransferFujiNexim_MI 
         Caption         =   "TransferFujiNexim_MI"
      End
      Begin VB.Menu mnuTransferPhilips 
         Caption         =   "TransferPhilips"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTransferPanaAMI 
         Caption         =   "TransferPanaMAI"
      End
      Begin VB.Menu mnuTransferPanaMSF 
         Caption         =   "TransferPanaMSF"
      End
      Begin VB.Menu mnuSingleSideBrdConfirm 
         Caption         =   "SingleSideBrdConfirm"
      End
      Begin VB.Menu mnuDefineBuildType 
         Caption         =   "DefineBuildType"
      End
      Begin VB.Menu mnuSAP1DataPatch 
         Caption         =   "SAP1DataPatch"
      End
      Begin VB.Menu mnuCostCenter 
         Caption         =   "CostBU"
      End
      Begin VB.Menu mmuWOInputPlan 
         Caption         =   "QSMS_WOInputPlan"
      End
      Begin VB.Menu mmuQSMS_SapHis 
         Caption         =   "Update QSMS_SapHis"
      End
      Begin VB.Menu mnuDeleteME_BOM 
         Caption         =   "DeleteME_BOM"
      End
      Begin VB.Menu mnuCompPrint 
         Caption         =   "CompPrint"
      End
      Begin VB.Menu mnuCompPNPrint 
         Caption         =   "CompPNPrint"
      End
      Begin VB.Menu mnuTransferCompPrint 
         Caption         =   "TransferCompPrint"
      End
      Begin VB.Menu mnuIC_Burn 
         Caption         =   "IC_Burn"
      End
      Begin VB.Menu mmuDIDIntegration 
         Caption         =   "DIDIntegration"
      End
      Begin VB.Menu mnuDummyECN 
         Caption         =   "Dummy ECN"
      End
      Begin VB.Menu mnuFixDispatchData 
         Caption         =   "QWMS_FixDispatchData"
      End
   End
   Begin VB.Menu mnumaintainFeeder1 
      Caption         =   "PD"
      Begin VB.Menu mnuDIDBake 
         Caption         =   "DIDBake"
      End
      Begin VB.Menu mnumaintainFeeder 
         Caption         =   "MaintainFeeder"
      End
      Begin VB.Menu mnuupdRealqty 
         Caption         =   "UpdateRealQty"
      End
      Begin VB.Menu mnuVerifyFeederSlot 
         Caption         =   "VerifyFeederSlot"
      End
      Begin VB.Menu mnuDeleteFeeder 
         Caption         =   "UnlinkFeederDID"
      End
      Begin VB.Menu mnuCloseWO 
         Caption         =   "CloseWo"
      End
      Begin VB.Menu mnuTransferFujiAVL 
         Caption         =   "TransferFujiAVL"
      End
      Begin VB.Menu mmuCompPNCompare 
         Caption         =   "CompPNCompare"
      End
      Begin VB.Menu mmuUnlockCompPNCompare 
         Caption         =   "UnlockCompPNCompare"
      End
      Begin VB.Menu mmuDIDDistribution 
         Caption         =   "DIDDistribution"
      End
      Begin VB.Menu mmuCheckDID 
         Caption         =   "CheckDID"
      End
      Begin VB.Menu mmuPEMainTain_WO 
         Caption         =   "PEMainTain_WO"
      End
      Begin VB.Menu mmuQSMS_Record_DIDInfo 
         Caption         =   "QSMS_Record_DIDInfo"
      End
   End
   Begin VB.Menu mnuPMC 
      Caption         =   "PMC"
      Begin VB.Menu mnumaintainWOSeq 
         Caption         =   "MaintainWO_Seq"
      End
      Begin VB.Menu mnuQueryWOGroup 
         Caption         =   "QueryWoGroup"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Begin VB.Menu mnuWipReport 
         Caption         =   "Report"
      End
      Begin VB.Menu mnfrmBeforeCheckBom 
         Caption         =   "Beforehand_CheckBom"
      End
      Begin VB.Menu mnuTraceReport 
         Caption         =   "TraceReport"
      End
      Begin VB.Menu mnuQueryKP 
         Caption         =   "Trace Report_Old"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCompPNReport 
         Caption         =   "CompPN Report"
      End
      Begin VB.Menu mnuQueryCheckBOM 
         Caption         =   "QueryCheckBOMResult"
      End
      Begin VB.Menu mnuQueryDID 
         Caption         =   "QueryDIDUse"
      End
      Begin VB.Menu mnuDIDNoUsed 
         Caption         =   "QueryDIDNoUsed"
      End
      Begin VB.Menu mnuQDIDNeedCut 
         Caption         =   "QueryDIDNeedCut"
      End
      Begin VB.Menu munPanelDiff 
         Caption         =   "PanelDiff"
      End
      Begin VB.Menu munQueryReplacePN 
         Caption         =   "QueryReplacePN"
      End
   End
   Begin VB.Menu mnuIPQC 
      Caption         =   "IPQC"
      Begin VB.Menu mnuInSpection 
         Caption         =   "InSpection"
      End
      Begin VB.Menu mnuQuery 
         Caption         =   "Query"
      End
      Begin VB.Menu mnuIPQCRelieve 
         Caption         =   "Relieve"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuQMS 
      Caption         =   "QMS"
      Begin VB.Menu mnuSetDIOandInterlock 
         Caption         =   "SetDIO&Interlock"
      End
      Begin VB.Menu mnuCheckDispatchQty 
         Caption         =   "Check&DispatchQty"
      End
      Begin VB.Menu mnuSendXLRemainDemand 
         Caption         =   "SendXLRemainDemand"
      End
   End
   Begin VB.Menu mnuSC 
      Caption         =   "SpecialCase"
      Begin VB.Menu mnuUpdateUID 
         Caption         =   "UpdateUID"
      End
      Begin VB.Menu mnuGenXLPrior 
         Caption         =   "GenXLPrior"
      End
      Begin VB.Menu mnuGenXLMD 
         Caption         =   "GenXLMaterialDemand"
      End
      Begin VB.Menu mnuUrgentInsertWO 
         Caption         =   "Urgent WO"
      End
      Begin VB.Menu mnuUnChkWO 
         Caption         =   "CloseUnCheckWO"
      End
      Begin VB.Menu mnuUrgentDIDToWH 
         Caption         =   "Urgent_DID_ToWH"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuStartSplitLineMC 
         Caption         =   "StartSplitLineMC"
      End
   End
   Begin VB.Menu mnuPrinterSetting 
      Caption         =   "PrinterSetting"
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************************************************************
'**文 件 名: Smt_PMS.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: LynnSun
'**日    期: 2007.12.22
'**描    述: DID Header
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**LynnSun      2007.12.22     Modify program get DID Head from table instead of Set.ini --------(0001)
'**Jing         2008.02.26     Add a function for Special Case  (0002)
'**Denver       2008.03.28     it need close all forms when Close form Main  --(0003)
'**Jing         2008.04.15     update userright for (mnuDIDChkStock,mmuWOInputPlan,mmuQSMS_SapHis)  (0004)
'**Jing         2008.06.03     Modify, if groupID is over 3 week,can not insert  (0005)
'**Denver       2008.06.16     Get DIDHead and PrtCallBKandReturn by Site(double check) (0006)
'**Salon        2008.07.25     Check  GroupID (0007)
'**Salon        2008.08.01     Add QB and QC factory (0008)
'**Udall        2008.08.06     Auto delete the QSMS part of the WorkOrder which SMT part has been deleted when close WO. (0009)
'**Udall        2008.10.09     为使产线能够在工单尚未开出来之前CheckBom,方便调整SAP BOM 和ME BOM，增加form frmBeforehandCheckBom. (0010)
'**Salon        2008.11.06     If OP choice any wo when OP define the Group for the wo,all wo in the group for the wo are choiced.  (000011)
'**Kane         2008.12.19     Add authority check for update real qty (0012)
'**Udall        2009.04.22     Add a new function for 特殊工单不用核对QSMS_AOI数据 (0013)
'**Udall        2009.04.29     针对BU物料有在不同厂区时的情况，灵活处理OP所在的厂区 (0014)
'**Kane         2009.08.13     增加传送替代料资料到Fuji Avllist表的功能(0015)
'**Richie       2009.10.19     增加机种烧录的IC对应的料号(MCC=>IC Burn) （0016）
'**Austin       2009.11.06     QueryDIDUse Report 增加查询出DID详细资料  （0017）
'**Austin       2009.11.10     QueryDIDUse 当不存在记录时，提示先查询   (0018)
'**Denver       2010.03.19     Add IC comp check function  （0068）
'**Austin       2010.03.22     MaintainWO_Seq时，如果改变线别->Query WO 的时候，先前选择的工单清空.(0069)
'**Denver       2010.04.07     BU Name change (ESBU to CC,ASBU to LC)  （0070）
'**Denver       2010.04.23     DID Return后，如无需求，现不直接退库，打印时DID位置显示CompPN(0071)
'**Austin       2010.04.23     在建WOGroup的时候，抓取线别从Machine表中 (0072)
'**Richie       2010.06.04     在建WOGroup的时候，检查该工单所在PCB是否已经有GroupID (0073)
'**Feix         2017.06.19     添加LCR 4次限制，同时添加解除（1258）
'**Feix         2018.10.24     添加新的界面TransferFujiNexim_MI（1271）
'***********************************************************************************/

Option Explicit
Dim returnDIDflag As Boolean
Private Sub Command1_Click()

'Call GetMEBOMQty("500047659", "AL006648012", "00", "A")
'Call CheckBom("500047659")
'Call GetWoinfo("500050977")
'Call InsertToQSMS_WO("500050977")
'Dim Rs As ADODB.Recordset
'Dim Str As String
'Dim I As Integer
'Str = "select '500037483','A',120, '31KH2MB0007','Others',CompPN,'',0, Qty/120,Qty,0,-Qty,'N','N'" & _
'      "from sap_bom where compPn not in (select comppn from QSMS_mebom where jobpn in (select jobpn from QSMS_jobbom where work_order='500037483'))"
'Set Rs = Conn.Execute(Str)
'I = 0
'While Not Rs.EOF
'     Str = "Insert into QSMS_Wo values('" & Trim(Rs.Fields(0)) & "','" & Trim(Rs.Fields(1)) & "','" & Trim(Rs.Fields(2)) & "','" & Trim(Rs.Fields(3)) & "','" & Trim(Rs.Fields(4)) & "' " & _
'          ",'" & Trim(Rs.Fields(5)) & "','" & I & "','" & Trim(Rs.Fields(7)) & "','" & Trim(Rs.Fields(8)) & "','" & Trim(Rs.Fields(9)) & "','" & Trim(Rs.Fields(10)) & "','" & Trim(Rs.Fields(11)) & "','N','N')"
'     Conn.Execute Str
'      I = I + 1
'      Rs.MoveNext
'Wend

End Sub

Private Sub Command2_Click()
Dim str As String
Dim Rs As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim Item, Slot As String

str = "select a.machine,a.comppn ,a.slot, sum(b.didqty) as qty from QSMS_wo a,QSMS_dispatch b where a.work_order=b.work_order and a.machine=b.machine and a.comppn=b.comppn and a.work_order='500049445' " & _
 "     and a.slot=b.slot group by a.machine,a.comppn,a.slot order by a.machine ,a.comppn,a.slot"
 Set Rs = Conn.Execute(str)
 While Not Rs.EOF
       str = "select Item,Slot from QSMS_Wo where Work_Order='500049445' and compPN='" & Rs!COMPPN & "' and Machine='" & Trim(Rs!Machine) & "' and Slot='" & Trim(Rs!Slot) & "'"
        Set rsTemp = Conn.Execute(str)
        If Not rsTemp.EOF Then
           Item = Trim(rsTemp!Item)
           Slot = Trim(rsTemp!Slot)
        End If
        '##################'for JuJi system,one slot has two subslot(L,R),so maybe need update the record by LR.----mark by leimo 20060516####################
        If Item = "0" Then
            str = "Update QSMS_Wo set dispatchqty= " & Rs!Qty & ",BalanceQty= -needqty+ " & Rs!Qty & " where Work_Order='500049445' and  CompPN='" & Rs!COMPPN & "' and Slot='" & Slot & "' and Machine='" & Trim(Rs!Machine) & "'"
        Else
        str = "Update QSMS_Wo set dispatchqty= " & Rs!Qty & ",BalanceQty=  balanceQty+ " & Rs!Qty & " where Work_Order='500049445' and  " & _
              " Item='" & Item & "' and slot='" & Slot & "' and Machine='" & Trim(Rs!Machine) & "'"
        End If
        Conn.Execute str
        Rs.MoveNext
 Wend
 
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Program As String
Dim Version As String
Dim strSQL As String
Dim Rs As New ADODB.Recordset
On Error GoTo errhander

mdiMain.Left = (Screen.Width - mdiMain.Width) / 2
mdiMain.Top = (Screen.Height - mdiMain.Height) / 2
 Select Case Right(App.path, 1)
        Case "\"
            WorkDir = Left(App.path, Len(App.path) - 1)
        Case Else
            WorkDir = App.path
End Select
Profile = WorkDir & "\set.ini"
hSECTION = "COMMON"
GetSettings Profile, hSECTION
Program = App.ProductName
'Version = "V200803019"
'If ChkPrgVer(Program, Version) = False Then
'    MsgBox ("Sorry, the program version " & Version & " is not the lastest,Please call QMS !!"), vbCritical
'    End
'Else
'    Me.Caption = ProgramDescription + " IP: " & IP
'End If

'
'If App.Title <> App.EXEName Then  'If run source code mode can ignore ChkVersion
'   Call ChkVersion("ALL", "QSMS", App.EXEName & ".exe")
'End If

'pFlagTest = ReadIniFile("QSMS", "FlagTest", App.Path & "\set.ini")

Check_NonAVL = ReadIniFile("QSMS", "Check_NonAVL", App.path & "\set.ini")
Check_AVL = ReadIniFile("QSMS", "Check_AVL", App.path & "\set.ini")
Check_DID = ReadIniFile("QSMS", "Check_DID", App.path & "\set.ini")
'DIDTailFlag = ReadIniFile("QSMS", "DIDTailFlag", App.Path & "\set.ini")
imagePath = ReadIniFile("QSMS", "ImagePath", App.path & "\set.ini")
TestFilepath = ReadIniFile("QSMS", "TestFilepath", App.path & "\set.ini")

'''**RQ09102710  Denver      2009.10.27    Add 测试LCR 型号为4300 的仪器  （0063）
IPQC_ChkVendorPN = ReadIniFile("QSMS", "IPQC_ChkVendorPN", App.path & "\set.ini")

'''**  Denver      2010.03.19    Add IC comp check function  （0068）
IC_CompChk = ReadIniFile("QSMS", "IC_CompChk", App.path & "\set.ini")


''是否导入DID CallBack and Reutn 打印DID Label
PrtCallBKandReturn = ReadIniFile("QSMS", "PrtCallBKandReturn", App.path & "\set.ini")
BU = ReadIniFile("Common", "BU", App.path & "\set.ini")
'CheckBomLogon = ReadIniFile("QSMS", "CheckBomLogon", App.Path & "\set.ini")


'strSQL = "select DIDHead from site"    ' (0001)
'Set Rs = Conn.Execute(strSQL)
'If Not Rs.EOF Then
'    DIDHead = Trim(Rs!DIDHead)
'Else
'    MsgBox "Can't get DID Head from table [Site], please define it first!", vbCritical
'    Unload Me
'End If


'20080616  Denver  Get DIDHead and PrtCallBKandReturn by Site(double check)--(0006)
'20100407 Denver    BU Name change (ESBU to CC,ASBU to LC)  （0070）
'20100423   Denver    DID Return后，如无需求，现不直接退库，打印时DID位置显示CompPN(0071)
strSQL = "exec XL_SiteData " & sq(Factory) & "," & sq(PrtCallBKandReturn)
Set Rs = Conn.Execute(strSQL)
If Not Rs.EOF Then

    PrtCallBKandReturn = Trim(Rs!PrtCallBKandReturn)
    DIDHead = Trim(Rs!DIDHead)
    AutoDispatchForAnotherBU = Trim(Rs!AutoDispatchForAnotherBU)
    CheckPNGroup = Trim(Rs!CheckPNGroup)
    BUDIDShow = Trim(Rs!BUDIDShow)
    DIDnotToQWMS = Trim(Rs!DIDnotToQWMS)
    If DIDHead = "" Then
        MsgBox "Can't get DID Head from table [Site], please define it first!", vbCritical
        Unload Me
    End If
Else
    MsgBox "Can't get DID Head from table [Site], please define it first!", vbCritical
    Unload Me
End If



Call SetMenu
LocalIP = Winsock.LocalIP
If Factory = "" Then   ''''(1084)
    If CheckFacIP = False Then
       End
    End If
End If

If Trim(Factory) <> "" Then
    LabFac.Caption = "Your QSMS program Factory: " & Factory & ", IP:" & LocalIP
Else
    LabFac.Caption = "Your IP:" & LocalIP & ", it is not defined into set.ini"
End If


If CreateDIDFlag = "N" Then
    menuAutoDispatch.Enabled = False
    mnuReturnDID.Enabled = False
    mnuReturnDIDALL = False
    mnuReturnComp = False
End If

Me.Caption = Me.Caption & Version & ProgramDescription & " IP: " & IP & "; Factory:" & Factory

UID = g_userName
Exit Sub
errhander:
MsgBox Err.Description
End
End Sub
Private Function CheckFacIP() As Boolean
Dim strIP() As String
Dim Rs As New ADODB.Recordset
Dim strSQL As String
Dim i As Integer, j As Integer
    LocalIP = Winsock.LocalIP
    CheckFacIP = False
    Factory = ""
    CreateDIDFlag = "N"
  ''''''(0014)  Start
    strSQL = "select distinct Factory from Site"
    If Rs.State = 1 Then Rs.Close
    Rs.CursorLocation = adUseClient
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = False Then
        ReDim FactoryID(Rs.RecordCount, 2)
    Else
        MsgBox "The Factory is empty,please connect with QMS for set the Factory in the Site table!"
        Exit Function
    End If
    If Rs.RecordCount > 1 Then      ''(0001)
        While Rs.EOF = False
             FactoryID(i, 0) = Rs.Fields("Factory")
             FactoryID(i, 1) = ReadIniFile("QSMS", Trim(Rs.Fields("Factory")), App.path & "\set.ini")
             i = i + 1
             Rs.MoveNext
        Wend
        For i = 0 To UBound(FactoryID)
            If Trim(FactoryID(i, 0) <> "" And Trim(FactoryID(i, 1) = "")) Then
                MsgBox "Your BU produce in " & FactoryID(i, 0) & " factories,please connect with QMS for set the " & FactoryID(i, 0) & " IP!"
                Exit Function
            End If
            strIP = Split(FactoryID(i, 1), ";")
            For j = 0 To UBound(strIP)
                If strIP(j) = Left(LocalIP, Len(strIP(j))) And Trim(strIP(j)) <> "" Then
                    If Trim(Factory) <> "" Then
                        MsgBox "Your IP " & LocalIP & " is exist in different factory,please connect with QMS check!"
                        Exit Function
                    Else
                        Factory = Trim(FactoryID(i, 0))
                        CreateDIDFlag = "Y"
                    End If
                End If
            Next j
        Next i
    Else
        Factory = Trim(Rs.Fields("Factory"))
        CreateDIDFlag = "Y"
    End If
    CheckFacIP = True
    ''''(0014)---------
End Function
Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    '**Denver       2008.03.28     it need close all forms when Close form Main  --(0003)
    For Each frm In Forms
        If frm.Name <> "Main" Then
            Unload frm
            
        End If
    Next frm
    
End Sub

Private Sub frmUnChkWO_Click()
frmUnChkWO.Show
End Sub

Private Sub menuAutoDispatch_Click()
frmMaintainDIDAutoDispatch.Show
End Sub

Private Sub mmuCheckDID_Click()
FrmChecKDID.Show  ''1257
End Sub

Private Sub mmuCompPNCompare_Click()
FrmPNCompare.Show
End Sub

Private Sub mmuDIDBake_Click()
FrmDIDBake.Show
End Sub

Private Sub mmuDIDDistribution_Click()
FrmDIDDistribution.Show
End Sub

Private Sub mmuDIDIntegration_Click()
FrmDIDInteGration.Show
End Sub

Private Sub mmuPEMainTain_WO_Click()
FrmPEMainTain_WO.Show  ''1259
End Sub

Private Sub mmuQSMS_Record_DIDInfo_Click()
FrmQSMS_Record_DIDInfo.Show          ''1259
End Sub

Private Sub mmuQSMS_SapHis_Click()
FrmUpdSapHis.Show 1
End Sub

Private Sub mmuUnlockCompPNCompare_Click()
FrmUnlockCompPNCompare.Show
End Sub

Private Sub mmuWOInputPlan_Click()
FrmInputPlanQty.Show
End Sub
Private Sub mnfrmBeforeCheckBom_Click()
frmBeforehandCheckBom.Show
End Sub
Private Sub mnuCheckDispatchQty_Click()
FrmCompDiff.Show
End Sub

Private Sub mnuCloseWO_Click()
FrmCloseWO.Show
End Sub
Private Sub mnuCompPNReport_Click()
frmCompPNReport.Show
End Sub
'20100617 Maggie Add CompPrint
Private Sub mnuCompPrint_Click()
frmCompPrint.Show
End Sub

'20101014 Maggie Add CompPNPrint   '(1013)
Private Sub mnuCompPNPrint_Click()
frmCompPNPrint.Show
End Sub

Private Sub mnuCostCenter_Click()
FrmcostBU.Show
End Sub

Private Sub mnuDefineBuildType_Click()
FrmDefineBuildType.Show
End Sub

Private Sub mnuDelete_Click()
    FrmInspection_Del.Show
End Sub

Private Sub mnuDeleteFeeder_Click()
    FrmUnlinkDIDFeeder.Show
End Sub

Private Sub mnuDeleteME_BOM_Click()
    StrDeleteLog = "Y"
    FrmDeleteME_BOM.Show
End Sub

Private Sub mnuDIDBake_Click()
  FrmDIDBake.Show
End Sub

Private Sub mnuDIDCallBack_Click()
    frmDIDCallBack_New.Show
End Sub

Private Sub mnuDIDChkStock_Click()
    frmDIDChkStock.FuncType = "ManualChk"
    frmDIDChkStock.Show
End Sub

Private Sub mnuDIDNoUsed_Click()
frmDIDNoUsed.Show
End Sub

Private Sub mnuDispatchDIDAdditionnal_Click()
FrmDispatchDIDAdditional.Show
End Sub

 

Private Sub mnuFixDispatchData_Click()
frmDispatchDataToQWMS.Show
End Sub

Private Sub mnuGenXLMD_Click()
    frmGenXLMD.Show
End Sub
'1288
Private Sub mnuGenXLPrior_Click()
    FrmGenXLPrior.Show
End Sub

Private Sub mnuIC_Burn_Click()
    frmICBurn.Show
End Sub


Private Sub mnuInheritDIDByWO_Click()
FrmInheritDIDByWO.Show
End Sub

Private Sub mnuInSpection_Click()
FrmInSpection.Visible = False
FrmInSpection.Show
End Sub

Private Sub mnuMaintainDID_Click()
'frmMaintainDIDAutoDispatch.Show
FrmMaintainDID.cmdForceDel.Visible = False
FrmMaintainDID.Show
End Sub

Private Sub mnumaintainFeeder_Click()
FrmMaintainFeeder.Show
End Sub

Private Sub mnumaintainWOSeq_Click()
FrmMaintainWOSeq.Show
End Sub

Private Sub mnuPDPreMaterial_Click()
'FrmPDPrepairMaterial.Show
End Sub

Private Sub MnuPrematerial_Click()
'FrmMCCPrepairMaterial.Show
End Sub

Private Sub mnuMCCPreMaterial_Click()
FrmPDPrepairMaterial.Show
End Sub

Private Sub mnuModfyDIDTotalQty_Click()
FrmModifyDIDTotalQty.Show
End Sub

Private Sub mnuPrinterSetting_Click()
    frmPrinterSetting.Show
End Sub

Private Sub mnuQDIDNeedCut_Click()
frmQueryDIDNeedCut.Show
End Sub

Private Sub mnuQuery_Click()
    FrmQueryInspect.Show
End Sub
Private Sub mnuIPQCRelieve_Click()
    FrmInRelieve.Show
End Sub

Private Sub mnuQueryByReturnedDID_Click()
    FrmQryReturnedDID.Show
End Sub

Private Sub mnuQueryCheckBOM_Click()
    FrmQueryCheckBOM.Show
End Sub

Private Sub mnuQueryDID_Click()
FrmQueryDID.Show 1
End Sub

Private Sub mnuQueryKP_Click()
FrmKPTS.Show
End Sub

Private Sub mnuQueryWOGroup_Click()
FrmQueryWoGroup.Show
End Sub

Private Sub mnuReturnComp_Click()
    FrmReturnComp.Show
End Sub

Private Sub mnuReturnDID_Click()
FrmReturnDID.returnDIDflag = returnDIDflag
FrmReturnDID.FraReturnDID.Visible = True
FrmReturnDID.FraReturnDIDALL.Visible = False

FrmReturnDID.Show
End Sub

Private Sub mnuReturnDIDALL_Click()
FrmReturnDID.FraReturnDID.Visible = False
FrmReturnDID.FraReturnDIDALL.Visible = True
FrmReturnDID.FraReturnDIDALL.Left = FrmReturnDID.FraReturnDID.Left
FrmReturnDID.FraReturnDIDALL.Top = FrmReturnDID.FraReturnDID.Top
FrmReturnDID.Show
End Sub


Private Sub mnuReturnMaterial_Click()
frmReturnMaterial.Show
End Sub

Private Sub mnuSAP1DataPatch_Click()
FrmSAP1Patch.Show
End Sub

Private Sub mnuSendXLRemainDemand_Click()
    Dim path As String
    path = App.path & "\C#EXE\QSMS.exe"
    
    'Shell "D:\superchai\TmpProgramForDev\PU9\QSMS_C#\QSMS\bin\Debug\QSMS.exe", vbNormalFocus        'superchai add 20231004
    Shell path, vbNormalFocus        'superchai add 20231004
End Sub

Private Sub mnuSetDIOandInterlock_Click()
frmSetInterDIO.Show
End Sub

Private Sub mnuSingleSideBrdConfirm_Click()
FrmSingleSideBrdConfirm.Show
End Sub

Private Sub mnuStartSplitLineMC_Click() '1181
FrmStartSplitLineMC.Show
End Sub

Private Sub mnuTraceReport_Click()
frmTraceReport.Show           ''''(1010)
End Sub

Private Sub mnuTransferCompPrint_Click()
    frmSelectCustomer.Show
End Sub

Private Sub mnuTransferDispatchDID_Click()
FrmTransferDispatchedDID_New.Show
End Sub

Private Sub mnuTransferFujiAVL_Click()
frmTransferFujiAVL.Show '0015
End Sub

Private Sub mnuTransferFujiNexim_Click()
frmTransferFujiNexim.Show
End Sub
''1271
Private Sub mnuTransferFujiNexim_MI_Click()
frmTransferFujiNexim_MI.Show
End Sub
''1271

Private Sub mnuTransferFujiXML_Click()
FrmTransferFujiXML.Show
End Sub

Private Sub mnuTransferPanaAMI_Click()
frmTransferPanaAMI.Show
End Sub

Private Sub mnuTransferPanaMSF_Click()
FrmTransferPanaMSF.Show
End Sub

Private Sub mnuTransferPhilips_Click()
FrmTransferPhilips.Show
End Sub


Private Sub mnuUnChkWO_Click()
frmUnChkWO.Show
End Sub

Private Sub mnuUpdateUID_Click()
FrmUpdateUID.Show
End Sub

'Private Sub mnuUniReport_Click()
'FrmUniReport.Show
'End Sub

Private Sub mnuupdRealqty_Click()
FrmUpdRealQty.Show
End Sub

Private Sub MnuUpLoadBom_Click()

FrmUpLoadData.Show
End Sub
Private Function SetMenu()
Dim i As Long


'MCC
mnuMCCPreMaterial.Enabled = False         '(1)
mnuInheritDIDByWO.Enabled = False         '(2)
mnuDispatchDIDAdditionnal.Enabled = False '(3)
mnuTransferDispatchDID.Enabled = False    '(4)
MnuUpLoadBom.Enabled = False              '(5)
mnuMaintainDID.Enabled = False            '(6)
mnuModfyDIDTotalQty.Enabled = False       '(7)
mnuReturnDID.Enabled = False              '(9)
mnuReturnDIDALL.Enabled = False           '(10)
mnuDIDCallBack.Enabled = False            '(11)
mnuDIDChkStock.Enabled = False               '(22)

mnuTransferFujiXML.Enabled = True        '(12)
mnuTransferFujiNexim.Enabled = False
mnuTransferPanaAMI.Enabled = False        '(13)
mnuTransferPanaMSF.Enabled = False        '(14)
     '(15)
mnuSingleSideBrdConfirm.Enabled = False   '(16)
mnuSAP1DataPatch.Enabled = False          '(18)
mnuDefineBuildType.Enabled = False        '(19)
MnuUpLoadMachineType.Enabled = False      '(20)
menuAutoDispatch.Enabled = False          '(21)

mmuWOInputPlan.Enabled = False
mmuQSMS_SapHis.Enabled = False

mnuReturnComp.Enabled = False
mmuDIDIntegration.Enabled = False   '(1074)
'20100626  Maggie Add mnuCompPrint
mnuCompPrint.Enabled = False
'20101014 Maggie Add mnuCompPNPrint
mnuCompPNPrint.Enabled = False            '(1013)
mnuFixDispatchData.Enabled = False          '1014

mnuDummyECN.Enabled = False

'PMC
mnumaintainWOSeq.Enabled = False          '(1)
mnuQueryWOGroup.Enabled = False           '(2)
'PD
mnuupdRealqty.Enabled = False  '0012
mnumaintainFeeder.Enabled = False         '(1)
mnuVerifyFeederSlot.Enabled = False       '(2)
mnuDeleteFeeder.Enabled = False           '(4)
mnuCloseWO.Enabled = False                '(5)
mmuCompPNCompare.Enabled = False '1064
mnuDIDBake.Enabled = False ''1275
mmuUnlockCompPNCompare.Enabled = False '1064
'Report
mnuWipReport.Enabled = False              '(1)
mnuQueryKP.Enabled = False                '(2)
mnuQueryDID.Enabled = False
'mnuUniReport.Enabled = False

'IPQC
mnuInSpection.Enabled = False

'QMS
mnuSetDIOandInterlock.Enabled = False     '(1)
mnuCheckDispatchQty.Enabled = False       '(2)
mnuDeleteME_BOM.Enabled = False
mnuSendXLRemainDemand.Enabled = True        'superchai add 20231004


'PrinterSetting
'20101115 Maggie Add PrinterSetting
mnuPrinterSetting.Enabled = True

''''''Added by Jing (0002)''''''
'Special Case

mnuUrgentInsertWO.Enabled = False
'mnuUnChkWO.Enabled = False              '(0013)
mnuGenXLMD.Enabled = False
mnuGenXLPrior.Enabled = False            '(1288)
mnuTransferFujiAVL.Enabled = False

'20141114 Sarah Add mnuStartSplitLineMC
mnuStartSplitLineMC.Enabled = False
mnuUpdateUID.Enabled = False

mmuPEMainTain_WO.Enabled = False

mmuQSMS_Record_DIDInfo.Enabled = False

strKeyInPNByManual = False
 For i = 0 To UBound(g_userRight)
   'MCC
     If g_userRight(i) = "mnuMCCPreMaterial" Then             '(1)
       mnuMCCPreMaterial.Enabled = True
    End If
    If g_userRight(i) = "mnuInheritDIDByWO" Then              '(2)
       mnuInheritDIDByWO.Enabled = True
    End If
    If g_userRight(i) = "mnuDispatchDIDAdditionnal" Then      '(3)
       mnuDispatchDIDAdditionnal.Enabled = True
    End If
    If g_userRight(i) = "mnuTransferDispatchDID" Then         '(4)
       mnuTransferDispatchDID.Enabled = True
    End If
    If g_userRight(i) = "MnuUpLoadBom" Then                   '(5)
       MnuUpLoadBom.Enabled = True
       mnuDeleteME_BOM.Enabled = True
    End If
    If g_userRight(i) = "mnuMaintainDID" Then                 '(6)
       mnuMaintainDID.Enabled = True
    End If
    If g_userRight(i) = "mnuModfyDIDTotalQty" Then            '(7)
       mnuModfyDIDTotalQty.Enabled = True
    End If
   
    If g_userRight(i) = "mnuReturnDID" Then                   '(9)
        mnuReturnDID.Enabled = True
        mnuReturnComp.Enabled = True
        'mnuDIDChkStock.Enabled = True
    End If
    
    If g_userRight(i) = "mnuDIDChkStock" Then   '''(0004)
        mnuDIDChkStock.Enabled = True
    End If
    
    If g_userRight(i) = "mnuReturnDIDALL" Then                '(10)
       mnuReturnDIDALL.Enabled = True
    End If
    
    If g_userRight(i) = "mnuDIDCallBack" Then                 '(11)
       mnuDIDCallBack.Enabled = True
       'mnuDIDChkStock.Enabled = True
    End If
    
    If g_userRight(i) = "mnuTransferFujiXML" Then             '(12)
       mnuTransferFujiXML.Enabled = True
    End If
    
    If g_userRight(i) = "mnuTransferFujiNexim" Then             '(12)
       mnuTransferFujiNexim.Enabled = True
    End If
    
    If g_userRight(i) = "mnuTransferFujiNexim_MI" Then             '(1271)
       mnuTransferFujiNexim.Enabled = True
    End If
    
    If g_userRight(i) = "mnuTransferPanaAMI" Then             '(13)
       mnuTransferPanaAMI.Enabled = True
    End If
    If g_userRight(i) = "mnuTransferPanaMSF" Then             '(14)
       mnuTransferPanaMSF.Enabled = True
    End If

    If g_userRight(i) = "mnuSingleSideBrdConfirm" Then        '(16)
       mnuSingleSideBrdConfirm.Enabled = True
    End If
    If g_userRight(i) = "mnuSAP1DataPatch" Then              '(18)
       mnuSAP1DataPatch.Enabled = True
    End If
    If g_userRight(i) = "mnuDefineBuildType" Then            '(19)
       mnuDefineBuildType.Enabled = True
    End If
    If g_userRight(i) = "MnuUpLoadMachineType" Then          '(20)
       MnuUpLoadMachineType.Enabled = True
    End If
'    If g_userRight(i) = "MnuAutoDispatch" Then
'        menuAutoDispatch.Enabled = True
'    End If
    If g_userRight(i) = "mnuMaintainDIDAutoDispatch" Then  '(1126)
        menuAutoDispatch.Enabled = True
    End If
    
    If g_userRight(i) = "mnuDIDBake" Then  '(1275)
        mnuDIDBake.Enabled = True
    End If
    
    If g_userRight(i) = "mmuWOInputPlan" Then   '''(0004)
        mmuWOInputPlan.Enabled = True
    End If
    
    If g_userRight(i) = "mmuQSMS_SapHis" Then   '''(0004)
        mmuQSMS_SapHis.Enabled = True
    End If
    
    '20100817  Maggie Add mnuCompPrint
    If g_userRight(i) = "mnuCompPrint" Then
        mnuCompPrint.Enabled = True
        mnuCompPNPrint.Enabled = True   '(1013)
    End If
    
    '20110621  Denver Add mnudummyecn
    If g_userRight(i) = "mnuDummyECN" Then
        mnuDummyECN.Enabled = True
        mnuDummyECN.Enabled = True   '(1013)
    End If
    
    'PMC
     If g_userRight(i) = "mnumaintainWOSeq" Then              '(1)
       mnumaintainWOSeq.Enabled = True
     End If
     If g_userRight(i) = "mnuQueryWOGroup" Then               '(2)
       mnuQueryWOGroup.Enabled = True
     End If
     
    'PD
    If g_userRight(i) = "mnumaintainFeeder" Then              '(1)
       mnumaintainFeeder.Enabled = True
    End If
      If g_userRight(i) = "mnuVerifyFeederSlot" Then          '(2)
       mnuVerifyFeederSlot.Enabled = False
    End If

    If g_userRight(i) = "mnuDeleteFeeder" Then                '(4)
       mnuDeleteFeeder.Enabled = True
    End If
    If g_userRight(i) = "mnuClearDIDSplicing" Then                '(4)
       mnuDeleteFeeder.Enabled = False
    End If
    If g_userRight(i) = "mnuCloseWO" Then                     '(5)
       mnuCloseWO = True
    End If
    'Report
     If g_userRight(i) = "mnuWipReport" Then                  '(1)
       mnuWipReport.Enabled = True
'       mnuUniReport.Enabled = True
     End If
     If g_userRight(i) = "mnuQueryKP" Then                    '(2)
       mnuQueryKP.Enabled = True
     End If
     If g_userRight(i) = "mnuQueryDID" Then                   '(3)
       mnuQueryDID.Enabled = True
     End If
     If g_userRight(i) = "returnDIDflag" Then
        returnDIDflag = True
     End If
    ' IPQC
     If UCase(g_userRight(i)) = "MNUINSPECTION" Then
        mnuInSpection.Enabled = True
     End If
     If UCase(g_userRight(i)) = "MNUDELETE" Then
        mnuDelete.Enabled = True
     End If
     If UCase(g_userRight(i)) = "MNURELIEVE" Then     '00001
        mnuIPQCRelieve.Enabled = True                 '00001
     End If
    'QMS
     If g_userRight(i) = "mnuSetDIOandInterlock" Then                  '(1)
       mnuSetDIOandInterlock.Enabled = True
     End If

     If g_userRight(i) = "mnuCheckDispatchQty" Then                  '(2)
       mnuCheckDispatchQty.Enabled = True
     End If
     If g_userRight(i) = "mmuCompPNCompare" Then  '1064
        mmuCompPNCompare.Enabled = True
     End If
     If g_userRight(i) = "mmuUnlockCompPNCompare" Then  '1064
        mmuUnlockCompPNCompare.Enabled = True
     End If
 
     ''''''Added by Jing 2008.02.26 (0002)''''''
     'Special Case
     If g_userRight(i) = "mnuUrgentInsertWO" Then
        mnuUrgentInsertWO.Enabled = True
     End If
     '1288
     If g_userRight(i) = "mnuGenXLPrior" Then
        mnuGenXLPrior.Enabled = True
     End If
     
     If g_userRight(i) = "mnuGenXLMD" Then
        mnuGenXLMD.Enabled = True
     End If

     If g_userRight(i) = "mnuUrgentDIDToWH" Then
        mnuUrgentDIDToWH.Enabled = True
     End If
     If g_userRight(i) = "mnuUpdRealQty" Then   '0012
        mnuupdRealqty.Enabled = True
     End If
     If g_userRight(i) = "mnuTransferFujiAVL" Then
        mnuTransferFujiAVL.Enabled = True
     End If
     'If UCase(g_userRight(I)) = "MNUUNCHKWO" Then       '''(0013)
        'mnuUnChkWO.Enabled = True
     'End If
     If UCase(g_userRight(i)) = "CHECKBOM" Then       '''(0013)
        CheckBomRight = True
     End If
     If UCase(g_userRight(i)) = "MNUFIXDISPATCHDATA" Then       '''(1014)
        mnuFixDispatchData.Enabled = True
     End If
     If UCase(g_userRight(i)) = UCase("KeyInPNByManual") Then
        strKeyInPNByManual = True
     End If
     If UCase(g_userRight(i)) = UCase("mmuDIDintegration") Then '' 1074
        mmuDIDIntegration.Enabled = True
     End If
     If UCase(g_userRight(i)) = UCase("DeleteMeBomByLine") Then '1131
        DeleteMeBomByLine = True
     End If
     If UCase(g_userRight(i)) = UCase("mnuStartSplitLineMC") Then
        mnuStartSplitLineMC.Enabled = True
     End If
     If UCase(g_userRight(i)) = UCase("mnuUpdateUID") Then '''''添加UpdateUID权限 1207
        mnuUpdateUID.Enabled = True
     End If
     If UCase(g_userRight(i)) = UCase("mmuPEMainTain_WO") Then             '1259
        mmuPEMainTain_WO.Enabled = True
     End If
     If UCase(g_userRight(i)) = UCase("mmuQSMS_Record_DIDInfo") Then             '1259
        mmuQSMS_Record_DIDInfo.Enabled = True
     End If
 Next i
End Function


Private Sub mnuUrgentDIDToWH_Click()
frmUrgentDIDToWH.Show
End Sub

Private Sub mnuUrgentInsertWO_Click()
FrmUrgentInsertWO.Show
End Sub
Private Sub mnuVerifyFeederSlot_Click()
FrmVerifyFeeder.Show
End Sub

Private Sub mnuWipReport_Click()
FrmReport.Show
End Sub

Private Sub ChkVersion(strLine As String, strStation As String, EXEName As String)
Dim Rs As New ADODB.Recordset
Dim Sqlstr As String
    Sqlstr = "select * from  Application_List  where AppEXE= '" & EXEName & "'"
    If strLine <> "" Or UCase(strLine) <> "ALL" Then
       Sqlstr = Sqlstr & " and Line = '" & Trim(strLine) & "' and StationName = '" & strStation & "' "
    End If
    Rs.Open Sqlstr, Conn, adOpenForwardOnly, adLockReadOnly
    If Rs.EOF = True Then
       MsgBox "The Program Version is Wrong,pls Access through MainMenu or Contact QMS!!", vbCritical
       End
    End If
End Sub

''20110620  Denver   Add Dummy ECN function
Private Sub mnuDummyECN_Click()
    frmDummyECN.Show
    
End Sub

Private Sub munPanelDiff_Click()
frmPanelDiff.Show
End Sub

Private Sub munQueryReplacePN_Click()
FrmQueryReplacePN.Show  ''1260
End Sub

