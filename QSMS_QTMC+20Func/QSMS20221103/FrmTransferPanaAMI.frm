VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTransferPanaAMI 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TransPanaMAI[20170413]"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10890
   Icon            =   "FrmTransferPanaAMI.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkNewVersion 
      Caption         =   "New Version"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
   Begin VB.CheckBox chkAutChkBom 
      BackColor       =   &H00C0C0FF&
      Caption         =   "AutoCheckBom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtFile 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   5295
   End
   Begin VB.CommandButton cmdGetMEBom 
      Caption         =   "Upload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9720
      TabIndex        =   0
      Top             =   270
      Width           =   1035
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1050
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10584
            MinWidth        =   10584
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5645
            MinWidth        =   5645
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LabelRun 
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   6000
      Width           =   10935
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmTransferPanaAMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************************************************************
'**文 件 名: frmTransferPanaAMI.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人:
'**日    期:
'**描    述: QSMS upload Pana machine bom
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**EQMS         Udall        2008.04.07     When Upload the machine bom,program aotu check bom   (00001)
'**             Jeanson      2009.02.03     Get the Parts Data by IDNUM instead of Row ID   (00002)
'                                           [PartsData]
'                                           IDNUM NAME CHIP SKIP A PACK SSIZE SHEIGHT CHIPCON RETRY
'                                           107 "CH41002KB93" 107 0 -0.000 90 0 0 0 3 (1-->107)
'                                           108 "CS31002FB26" 108 0 -0.000 90 0 0 0 3 (2-->108)
'**             Kane         2009.06.22     ME新版程序产生的文件位置有变化，之前16位是JobGroup，现在改为17位 '(00003)
'**RQ09081407   Kane         2009.09.22     增加一列ReelWidth保存到数据库'(00004)
'**QMS          Archer       2009.10.20     Add New column for Save Nozzle data (0005)
'**QMS          Austin       2010.05.06     读取Location到QSMS_MEBom中..(0006)
'**QMS          Austin       2010.08.07     Modify PositionData(Header) integer->Long (0006)
'-----------------------------------------------------------------------------
Option Explicit
Dim I As Integer
Dim Interval As Integer
Dim strJobPN As String, strRev As String, strLine As String, strBuildType As String                ''(00001)

Private Type MachineType
    MCName As String
    MCNo As String
    HeadNo As String
    Table As Integer
End Type

Private Type NozzleType     '''(0005)
    Machine As String
    HeadNo As String
    tmpSlot As String
    location As String
    NozzleType As String
End Type

Private Type PartData
    IdNo As Integer
    PN As String
    location As String
    Skip As String  ''1172
End Type

Private Type PositionData
    Machine As String
    BrdPN As String
    Rev As String
    PU As String
    Table As String
    TraySlot As String
    Slot As String
    Side As Integer
    Head As Long   ''''0006
    Parts As Integer
    compPN As String
    Qty As Integer
    Enabled As Boolean
    FstMachinePNRev As Boolean
    location As String   ''Add Location
    NPMReelWidth As String   ''Add NPM ReelWidth    (1069)
    DualLaneMode As String
    B As String ''1172
    F As String ''1172
End Type

Private Type StockData
    PU As String
    PA As String
    PB As String
    TA As String
    TB As String
End Type

Private Type BlockAttribute ''1172
    IDNUM As String
    B As String
End Type

Const MCSeq = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

'1178 begin
Private MC_Header() As String
Private PD_Header() As String
Private PT_Header() As String
Private PL_Header() As String
Private NZ_Header() As String
Private SD_Header() As String
Private BA_Header() As String
Private BD_Header() As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'1178 end

Private Sub cmdGetMEBom_Click()

    If Trim(txtFile) = "" Then
        MsgBox "You must select a file!!", vbInformation
        Exit Sub
    End If
        
    If LoadDataFile(Trim(txtFile)) = False Then
        MsgBox ("Fail")
    Else
        If chkAutChkBom.Value = 1 Then
            Call AutoCheckBom        ''(00001)
        End If
        MsgBox ("Finish")
    End If
    
    LabelRun.BackColor = &HFFFFC0
    LabelRun.Caption = "OK"
    StatusBar1.Panels(1) = txtFile & " OK"
    StatusBar1.Panels(3) = "Finished DateTime:" & Now
    I = 0
End Sub
    
Function LoadDataFile(strFile As String) As Boolean
Dim NFile As Integer, I As Integer, j As Integer, t As Integer, intBomFileName As Integer
Dim Arry() As String, Arry2() As String, temp() As String, log As String, Flg As String, tempBomFileName() As String, strFullJobGroup As String
Dim Factory As String, Line As String, MBPN As String, Revision As String, BuildType As String, strSide As String, Machine As String, MCNo As String, Head As String, Parts As String, PU As String, Side As String, IDNUM As String, PartsName As String
Dim BomFile As String, strBomFileName As String, strCurrent As String
Dim StartTime As String, EndTime As String, BoardType As String
Dim strSql As String, BrdPN As String, BrdRev As String, jobgroup As String, MCType As String, strSqlBomFileName As String
Dim idxMC As Integer, idxPD As Integer, idxPT As Integer, idxPL As Integer, idxNZ As Integer, idxSD As Integer, TraySlot As String, idxBA As String ''1172
Dim ReelWidth As String
Dim MC() As MachineType, PD() As PositionData, pt() As PartData, PL() As String, NZ() As NozzleType, SD() As StockData, BA() As BlockAttribute ''1172
Dim rs As ADODB.Recordset
Dim Total_Qty, Insert_Qty, Update_Qty As Integer
Dim ErrorDesc As String, PCBSize As String
'On Error GoTo errHandler
ReDim BA(0) ''1172

'1178 begin
Dim IsTableHeader As Boolean  '标记下一行是不是表头
IsTableHeader = False
'1178 end

LoadDataFile = False
Arry = Split(strFile, "\")
BomFile = Arry(UBound(Arry))

temp = Split(Trim(BomFile), "-")

If ChkEQProgram = "Y" Then   ''(1219)
    strBomFileName = Left(BomFile, InStr(BomFile, ".") - 1)
    tempBomFileName = Split(Trim(strBomFileName), "-")
    strFullJobGroup = Trim(tempBomFileName(2)) + "-" + Trim(tempBomFileName(3))
    If UBound(tempBomFileName) >= 6 Then
        For intBomFileName = 6 To UBound(tempBomFileName)
            strFullJobGroup = strFullJobGroup + "-" + Trim(tempBomFileName(intBomFileName))
        Next intBomFileName
    End If
    temp = Split(Trim(Left(strBomFileName, 26)), "-")
End If
        
If UBound(temp) <> 5 And UBound(temp) <> 6 Then
    MsgBox ("Filename format must be Factory-Line-PN-Rev-BuildType-Side !")  '(0007) '(1026)
    Exit Function
End If
If Trim(temp(4)) <> "1" And Trim(temp(4)) <> "2" And Trim(temp(4)) <> "3" And Trim(temp(4)) <> "4" Then
   MsgBox ("BuildType must be 1,2,3 or 4.")
   Exit Function
End If
        
If Left(Trim(temp(5)), 1) <> "S" And Left(Trim(temp(5)), 1) <> "C" And Left(Trim(temp(5)), 1) <> "Q" Then
   MsgBox ("Side must be S,C or Q.")
   Exit Function
End If
If Len(Trim(temp(3))) <> 3 And Len(Trim(temp(3))) <> 2 And Len(Trim(temp(3))) <> 0 Then   '(1150)
   MsgBox ("The version length must be 2 or 3 or 0!")
   Exit Function
End If
'******************************
'****add by jeanson 2007/09/03
strErrMessage = ""
strErrMessage = FunPartNumberCheck(Trim(temp(2)))
If strErrMessage <> "PASS" Then
    MsgBox strErrMessage
    Exit Function
End If


strSql = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_PanaAMI Start','" & Left(Trim(BomFile), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSql)
'******************************

    Factory = Trim(temp(0)) 'add by giant 2008/06/27 (0007)
    MBPN = Trim(temp(2))
    Line = Trim(temp(1))
    Revision = Trim(temp(3))
    BuildType = Trim(temp(4))
    strSide = Left(Trim(temp(5)), 1)
    
    jobgroup = MBPN & "-" & Revision
    
    strJobPN = MBPN         ''(00001)
    strRev = Revision       ''(00001)
    strBuildType = BuildType        ''(00001)
    
    MCType = ""
   If (BuildType = "2" Or BuildType = "4") And strSide <> "S" Then
         MsgBox ("The side is " & strSide & ",BuildType is 2 or 4,they are not match,the side must be S side.")
         Exit Function
   End If
   If BuildType = "3" And strSide <> "C" Then
         MsgBox ("The side is " & strSide & ",BuildType is 3,they are not match,the side must be C side.")
         Exit Function
   End If
    
    
    NFile = FreeFile
    StartTime = Now
    LabelRun.BackColor = &HFF&
    LabelRun.Caption = "Running..."
    Open strFile For Input As #NFile
    StatusBar1.Panels(1) = "GetMEBom_" & BomFile
    StatusBar1.Panels(2) = "Start DateTime:" & StartTime
    idxMC = 0: idxPD = 0: idxPT = 0: idxPL = 0: idxNZ = 0: idxSD = 0: idxBA = 0 ''1172
    While Not EOF(NFile)
        I = I + 1
        Line Input #NFile, strCurrent
        strCurrent = Trim(Replace(Replace(strCurrent, vbCrLf, ""), Chr(9), " "))

        Select Case strCurrent
            Case "[Index]"
                 log = "ID"
                 Line Input #NFile, strCurrent
                 I = I + 1
            Case "[Machines]"
                 log = "MC"
                 'Line Input #NFile, strCurrent  '1178 不要跳过表头所在的行
                 I = I + 1
                 IsTableHeader = True
            Case "[PositionData]"
                 log = "PD"
                 'Line Input #NFile, strCurrent
                 I = I + 1
                 IsTableHeader = True
            Case "[PartsData]"
                 log = "PT"
                 'Line Input #NFile, strCurrent
                 I = I + 1
                 IsTableHeader = True
            Case "[PartsLIB]" '(00004)
                 log = "PL"
                 'Line Input #NFile, strCurrent
                 I = I + 1
                 IsTableHeader = True
            Case "[NozzleStock]" '(0005)
                 log = "NZ"
                 'Line Input #NFile, strCurrent
                 I = I + 1
                 IsTableHeader = True
            Case "[StockData]"   ''0006
                log = "SD"
                'Line Input #NFile, strCurrent
                I = I + 1
                IsTableHeader = True
            Case "[BlockAttribute]" ''1172
                log = "BA"
                'Line Input #NFile, strCurrent
                I = I + 1
                IsTableHeader = True
            Case "[BoardData]"      ''1203
                log = "BD"
                'Line Input #NFile, strCurrent
                I = I + 1
                IsTableHeader = True
            Case Else
                If IsTableHeader = True Then
                    Call PhraseHeader(strCurrent, log)
                    IsTableHeader = False
                Else
                    Arry = Split(strCurrent, " ")
                    If UBound(Arry()) > 3 Or (UBound(Arry()) = 2 And log = "NZ") Or (UBound(Arry()) = 0 And (log = "ID" Or log = "BD")) Then
                          Flg = "OK"
                    Else
                       log = ""
                       Flg = "Canel"
                    End If
                    Select Case log + Flg
                        Case "IDOK"
                            If MCType = "" Then
                                MCType = Trim(GetKeyValueM(strCurrent, "Machine"))
                            End If
                        Case "MCOK"
                            idxMC = idxMC + 1
                            ReDim Preserve MC(idxMC)
                            
                            MC(idxMC).MCName = Trim(Arry(GetPosition(MC_Header, "MCNAME")))
                            If MC(idxMC).MCName Like "*NPM*" Then   ''(1182)
                                MC(idxMC).MCName = Left(MC(idxMC).MCName, InStr(MC(idxMC).MCName, "NPM") + 2)
                            End If
                            'MC(idxMC).MCName = Trim(Arry(1))
    '                        If Len(MC(idxMC).MCName) <> 5 Then ''(1083) marked
    '                            MsgBox "MCName : " & MC(idxMC).MCName & " prefix length must be 5!"
    '                            Exit Function
    '                        End If
                            MC(idxMC).MCNo = Trim(Arry(GetPosition(MC_Header, "MCNo"))) 'MC(idxMC).MCNo = Trim(Arry(2))
                            MC(idxMC).HeadNo = Trim(Arry(GetPosition(MC_Header, "HeadNo"))) 'MC(idxMC).HeadNo = Trim(Arry(3))
                        Case "SDOK"
                            idxSD = idxSD + 1
                            ReDim Preserve SD(idxSD)
                            SD(idxSD).PU = Trim(Arry(GetPosition(SD_Header, "N"))) 'SD(idxSD).PU = Trim(Arry(1))
                            SD(idxSD).PA = Trim(Arry(GetPosition(SD_Header, "PA"))) 'SD(idxSD).PA = Trim(Arry(2))
                            SD(idxSD).PB = Trim(Arry(GetPosition(SD_Header, "PB"))) 'SD(idxSD).PB = Trim(Arry(3))   ''1184
                            SD(idxSD).TA = Trim(Arry(GetPosition(SD_Header, "TA"))) 'SD(idxSD).TA = Trim(Arry(12))  ''1184
                            SD(idxSD).TB = Trim(Arry(GetPosition(SD_Header, "TB"))) 'SD(idxSD).TB = Trim(Arry(13))  ''1184
                        Case "PDOK"
                            idxPD = idxPD + 1
                            ReDim Preserve PD(idxPD)
'                            If UBound(Arry()) = 23 And StrBU = "NB3" Then          ''（1174）
'                                For I = 14 To 21
'                                    Arry(I) = Arry(I + 1)
'                                Next I
'                            End If
                            
                            PD(idxPD).Parts = Trim(Arry(GetPosition(PD_Header, "PARTS"))) 'PD(idxPD).Parts = Trim(Arry(5))
                            PD(idxPD).PU = Trim(Arry(GetPosition(PD_Header, "PU"))) 'PD(idxPD).PU = Trim(Arry(11))
                            PD(idxPD).location = Trim(Arry(GetPosition(PD_Header, "C"))) 'PD(idxPD).location = Trim(Arry(16))
                            PD(idxPD).B = Trim(Arry(GetPosition(PD_Header, "B"))) 'PD(idxPD).B = Trim(Arry(14))    ''1172
                            PD(idxPD).F = Trim(Arry(GetPosition(PD_Header, "F"))) 'PD(idxPD).F = Trim(Arry(18))  ''1172
                            'Trim(Arry(12)) represents LR;  0:no LR, 1:L, 2:R
'                            If Trim(Arry(12)) > 2 Then
'                                MsgBox "Side (LR) wrong : " & Arry(12) & ", it must be 0 or 1 or 2!"
'                                Exit Function
'                            End If
                            If Trim(Arry(GetPosition(PD_Header, "SIDE"))) > 2 Then
                                MsgBox "Side (LR) wrong : " & Arry(GetPosition(PD_Header, "SIDE")) & ", it must be 0 or 1 or 2!"
                                Exit Function
                            End If
                            PD(idxPD).Side = Trim(Arry(GetPosition(PD_Header, "SIDE"))) 'PD(idxPD).Side = Trim(Arry(12))
                            PD(idxPD).Head = Trim(Arry(GetPosition(PD_Header, "HEAD"))) 'PD(idxPD).Head = Trim(Arry(13))
                            PD(idxPD).Qty = 1
                            If Len(PD(idxPD).PU) >= 4 Then
                                PD(idxPD).Enabled = True
                            End If
                            
                            'BoardType = IIf(chkNewVersion.Value = 1, Trim(Arry(17)), Trim(Arry(16))) '(00003)
                            BoardType = Trim(Arry(GetPosition(PD_Header, "C")))
                            Arry2 = Split(BoardType, "-")
                            If UBound(Arry2) > 0 Then
                                If UBound(Arry2) <> 2 Then
                                    MsgBox "BoardType wrong : " & BoardType & ", must be PN-REV!"
                                    Exit Function
                                End If
                                If Len(Replace(Trim(Arry2(2)), """", "")) <> 3 And Len(Replace(Trim(Arry2(2)), """", "")) <> 2 Then
                                    MsgBox "Version wrong : " & BoardType & ", the version " & Replace(Trim(Arry2(2)), """", "") & "length must be 2 or 3!"
                                    Exit Function
                                End If
                                '******************************
                                '****add by jeanson 2007/09/03
                                strErrMessage = ""
                                strErrMessage = FunPartNumberCheck(Replace(Trim(Arry2(1)), """", ""))
                                If strErrMessage <> "PASS" Then
                                    MsgBox strErrMessage
                                Exit Function
                                End If
                                '******************************
                                    
                                PD(idxPD).BrdPN = Trim(Arry2(1))
                                PD(idxPD).Rev = Replace(Trim(Arry2(2)), """", "")
                            Else
                                PD(idxPD).BrdPN = MBPN
                                PD(idxPD).Rev = Revision
                            End If
                            PD(idxPD).FstMachinePNRev = True
                        Case "PTOK"
                            'Get the Parts Data by IDNUM instead of Row ID   (00002)
                            'idxPT = idxPT + 1
                            idxPT = Trim(Arry(GetPosition(PT_Header, "IDNUM"))) 'idxPT = Trim(Arry(0))
                            ReDim Preserve pt(idxPT)
                            pt(idxPT).IdNo = Trim(Arry(GetPosition(PT_Header, "IDNUM"))) 'pt(idxPT).IdNo = Trim(Arry(0))
                            pt(idxPT).PN = Trim(Arry(GetPosition(PT_Header, "NAME"))) 'pt(idxPT).PN = Trim(Arry(1))
                            pt(idxPT).Skip = Trim(Arry(GetPosition(PT_Header, "SKIP"))) 'pt(idxPT).Skip = Trim(Arry(3))  ''1172
                        Case "PLOK" '(00004)
                            idxPL = idxPL + 1
                            ReDim Preserve PL(2, idxPL) As String
                            PL(0, idxPL) = Replace(Trim(Arry(GetPosition(PL_Header, "PartsName"))), """", "") 'PL(0, idxPL) = Replace(Trim(Arry(1)), """", "")
                            PL(1, idxPL) = Replace(Trim(Arry(GetPosition(PL_Header, "ReelWidth"))), """", "") 'PL(1, idxPL) = Replace(Trim(Arry(45)), """", "")
                        Case "NZOK" '(0005)
                            idxNZ = Trim(Arry(GetPosition(NZ_Header, "IDNUM"))) 'idxNZ = Trim(Arry(0))
                            ReDim Preserve NZ(idxNZ)
                            NZ(idxNZ).HeadNo = Left(Trim(Arry(GetPosition(NZ_Header, "N"))), Len(Trim(Arry(GetPosition(NZ_Header, "N")))) - 2) 'NZ(idxNZ).HeadNo = Left(Trim(Arry(1)), Len(Trim(Arry(1))) - 2)
                            NZ(idxNZ).tmpSlot = Right(Trim(Arry(GetPosition(NZ_Header, "N"))), 2) 'NZ(idxNZ).tmpSlot = Right(Trim(Arry(1)), 2)
                            NZ(idxNZ).NozzleType = Trim(Arry(GetPosition(NZ_Header, "P"))) 'NZ(idxNZ).NozzleType = Trim(Arry(2))
                        Case "BDOK"
                            If PCBSize = "" Then
                                PCBSize = Trim(GetKeyValueM(BD_Header(0), "X"))    ''1203
                            End If
                        Case "BAOK" ''1172
'                            If Trim(Arry(4)) = 1 Then
'                                idxBA = idxBA + 1
'                                ReDim Preserve BA(idxBA)
'                                BA(idxBA).IDNUM = Trim(Arry(0))
'                                BA(idxBA).B = Trim(Arry(4))
'                            End If
                            If Trim(Arry(GetPosition(BA_Header, "B"))) = 1 Then
                                idxBA = idxBA + 1
                                ReDim Preserve BA(idxBA)
                                BA(idxBA).IDNUM = Trim(Arry(GetPosition(BA_Header, "IDNUM")))
                                BA(idxBA).B = Trim(Arry(GetPosition(BA_Header, "B")))
                            End If
                    End Select
                End If
        End Select
    Wend
    Close #NFile

    Dim PreMCNo As String
    PreMCNo = 0
    For I = 1 To UBound(MC)
        If MC(I).MCNo = PreMCNo Then
            MC(I).Table = MC(I - 1).Table + 1
        Else
            MC(I).Table = 1
            PreMCNo = MC(I).MCNo
        End If
        For j = 1 To UBound(NZ)         ''''(0005)
            If NZ(j).HeadNo = MC(I).HeadNo Then
                NZ(j).Machine = MC(I).MCName & Mid(MCSeq, MC(MC(I).HeadNo).MCNo, 1)
                NZ(j).location = MC(I).Table & "-" & NZ(j).tmpSlot
'                Debug.Print NZ(j).machine & " " & NZ(j).Location & " " & NZ(j).NozzleType
            End If
        Next j
'        Debug.Print MC(i).MCName & " " & MC(i).MCNo & " " & MC(i).HeadNo & " " & MC(i).Table
    Next I
    I = 0
    'Get CompPN from [PartsData] and set machine from [Machine]
    
    
    'Get Location from [PositionData]   '相同CompPN，Slot，LR的Location放一起
'    For I = 1 To UBound(pt) '''- 1  (1232)
'        For j = 1 To UBound(PD) ''''- 1 (1232)
'            If pt(I).IdNo = PD(j).Parts Then
'                pt(I).location = pt(I).location & ";" & PD(j).location
'            End If
'        Next j
'    Next I
    
    'Get NPM ReelWidth from [StockData] (1069)
    If MCType = NPMMachineType Then    ''(1079)
        For I = 1 To UBound(SD)
            For j = 1 To UBound(PD)
                If SD(I).PA = PD(j).Parts And SD(I).PU = PD(j).PU Then
                    PD(j).NPMReelWidth = SD(I).TA
                ElseIf SD(I).PB = PD(j).Parts And SD(I).PU = PD(j).PU Then
                    PD(j).NPMReelWidth = SD(I).TB
                End If
            Next j
        Next I
    End If
    
    For I = 1 To UBound(PD)
    
        For j = 1 To UBound(BA) ''1172 skip 区块
            If BA(j).B = 1 And BA(j).IDNUM = PD(I).B Then
                PD(I).Enabled = False
            End If
        Next j
        
        For j = 1 To UBound(pt) ''1172 skip 元件
            If IIf(pt(j).Skip = "", 0, pt(j).Skip) = 1 And pt(j).IdNo = PD(I).Parts Then    ''1217
                PD(I).Enabled = False
            End If
        Next j
        
        If PD(I).F = 2 Then ''1172 skip 站位
            PD(I).Enabled = False
        End If
        
        'If PD(i).Head > 0 Then
        If Val(Left(Trim(PD(I).PU), 1)) > 0 Then
'            PD(I).compPN = Left(Replace(pt(PD(I).Parts).PN, """", ""), 11)   'remove quote  ''(1183)
            PD(I).compPN = Replace(pt(PD(I).Parts).PN, """", "")  ''(1183)
            ''''''''''''''''1193'''''''''''
            If Len(Trim(PD(I).compPN)) <> 11 And Len(Trim(PD(I).compPN)) <> 14 Then
                PD(I).compPN = Mid(PD(I).compPN, 1, IIf(InStr(PD(I).compPN, "-") <> 0, InStr(PD(I).compPN, "-") - 1, Len(PD(I).compPN)))
            End If
            '''''''''''''1193''''''''''''
'            PD(I).location = Replace(pt(PD(I).Parts).location, """", "")
            'C11-41VF3SS03N0-A3A
            '******************************
            '****add by jeanson 2007/09/03
            strErrMessage = ""
            strErrMessage = FunPartNumberCheck(Trim(PD(I).compPN))
            If strErrMessage <> "PASS" Then
                MsgBox strErrMessage
                Exit Function
            End If
            '******************************
            
            PD(I).Machine = MC(PD(I).Head).MCName & Mid(MCSeq, MC(PD(I).Head).MCNo, 1)
            
            '20111017 Maggie 上传AMI文件时检查文件名称,区分是否为NPM DualLaneMode (1075)
            'If PD(i).machine Like "*NPM*" And (Left(Trim(temp(6)), 1) = "F" Or Left(Trim(temp(6)), 1) = "R") Then
                'PD(i).DualLaneMode = "Mix"
            'End If
            If PD(I).Machine Like "*NPM*" And UBound(temp) = 6 Then
                If (Left(Trim(temp(6)), 1) = "F" Or Left(Trim(temp(6)), 1) = "R") Then
                    PD(I).DualLaneMode = "Mix"
                End If
            End If

            'PD(i).Table = Trim(CStr(PD(i).Head - 4 * (MC(PD(i).Head).MCNo - 1)))
            'Left(Trim(PD(i).PU), 1) should be equal to PD(i).Head
            'PD(I).Table = Trim(CStr(Val(Left(Trim(PD(I).PU), Len(Trim(PD(I).PU)) - 4)) - 4 * (MC(Val(Left(Trim(PD(I).PU), Len(Trim(PD(I).PU)) - 4))).MCNo - 1)))

            
            PD(I).Table = MC(Val(Left(Trim(PD(I).PU), Len(Trim(PD(I).PU)) - 4))).Table
            PD(I).Slot = PD(I).Table & "-" & Trim(CStr(Val(Right(Trim(PD(I).PU), 2))))
            PD(I).TraySlot = Trim(CStr(Val(Right(Left(Trim(PD(I).PU), Len(PD(I).PU) - 2), 2))))
            
            If Settings.UpdateJobSide = "Y" Then
               strSql = "select * from QSMS_JobSide where JobPN='" & Trim(PD(I).BrdPN) & "'"
               Set rs = Conn.Execute(strSql)
               If rs.EOF Then
                  MsgBox Trim(PD(I).BrdPN) & ":Can't find the job side by the JobPN,please check!", vbCritical, "ErrMessage"
                  Exit Function
               Else
                  If UCase(Trim(rs("side"))) <> "S" And UCase(Trim(rs("side"))) <> "C" Then
                     MsgBox Trim(rs("Side")) & ":Job side's format is wrong ,the side must be S or C,please define it afresh!", vbCritical, "ErrMessage"
                     Exit Function
                  End If
               End If
            End If
            
        End If
    Next I
    
    'Count Comp Qty group by same machine, slot, side, Job, Rev
    For I = 1 To UBound(PD)
        If PD(I).Enabled Then
            For j = I + 1 To UBound(PD)
                If PD(j).Enabled And PD(I).BrdPN = PD(j).BrdPN And PD(I).compPN = PD(j).compPN And _
                    PD(I).Machine = PD(j).Machine And PD(I).Slot = PD(j).Slot And _
                    PD(I).Side = PD(j).Side Then
                    
                    PD(j).Enabled = False
                    PD(I).Qty = PD(I).Qty + 1
                    PD(I).location = PD(I).location & ";" & PD(j).location  '相同CompPN，Slot，LR的Location放一起
                    
                End If
                
                'CHECK first machine/brdpn/rev
                
                If UCase(PD(I).DualLaneMode) = "MIX" Then
                    If PD(j).FstMachinePNRev = True And PD(j).Machine = PD(I).Machine And PD(I).BrdPN = PD(j).BrdPN And PD(I).Table = PD(j).Table Then
                        PD(j).FstMachinePNRev = False
                    End If
                Else
                    If PD(j).FstMachinePNRev = True And PD(j).Machine = PD(I).Machine And PD(I).BrdPN = PD(j).BrdPN Then
                        PD(j).FstMachinePNRev = False
                    End If
                End If
            Next j
            
            PD(I).location = Replace(PD(I).location, """", "")

            'delete by machine, jobpn, rev at first time
            If PD(I).FstMachinePNRev Then
                If CheckMachine(Line, PD(I).Machine, strSide) = False Then '(1032)
                    Exit Function
                End If
                
                If UCase(PD(I).DualLaneMode) = "MIX" Then
                    strSql = "delete from QSMS_MEBom where JobGroup='" & jobgroup & "' and Machine=" & sq(PD(I).Machine) & " and JobPN=" & sq(PD(I).BrdPN) & " and version=" & sq(PD(I).Rev) & " And BuildType = " & sq(BuildType) & " and Factory=" & sq(Factory) & "and Slot like '" & Left(Trim(PD(I).Slot), 1) & "%' and Line=" & sq(Line) & "" '(0007)
                    strSqlBomFileName = "delete from QSMS_MEBom_EQProgram where FullJobGroup='" & strFullJobGroup & "' and Machine=" & sq(PD(I).Machine) & " and JobPN=" & sq(PD(I).BrdPN) & " and version=" & sq(PD(I).Rev) & " And BuildType = " & sq(BuildType) & " and Factory=" & sq(Factory) & "and Slot like '" & Left(Trim(PD(I).Slot), 1) & "%' and Line=" & sq(Line) & "" '(1219)
                Else
                    strSql = "delete from QSMS_MEBom where JobGroup='" & jobgroup & "' and Machine=" & sq(PD(I).Machine) & " and JobPN=" & sq(PD(I).BrdPN) & " and version=" & sq(PD(I).Rev) & " And BuildType = " & sq(BuildType) & " and Factory=" & sq(Factory) & " and Line=" & sq(Line) & ""     '(0007)
                    strSqlBomFileName = "delete from QSMS_MEBom_EQProgram where FullJobGroup='" & strFullJobGroup & "' and Machine=" & sq(PD(I).Machine) & " and JobPN=" & sq(PD(I).BrdPN) & " and version=" & sq(PD(I).Rev) & " And BuildType = " & sq(BuildType) & " and Factory=" & sq(Factory) & " and Line=" & sq(Line) & ""     '(1219)
                End If
                
                Conn.Execute (strSql)
                If ChkEQProgram = "Y" Then   ''(1219)
                    Conn.Execute (strSqlBomFileName)
                End If
            End If
            
            With PD(I)
                If Settings.UpdateJobSide = "Y" Then
                   strSql = "select * from QSMS_JobSide where JobPN='" & Trim(PD(I).BrdPN) & "'"
                   Set rs = Conn.Execute(strSql)
                   If rs.EOF Then
                      MsgBox Trim(PD(I).BrdPN) & ":Can't find the job side by the JobPN,please check!", vbCritical, "ErrMessage"
                      Exit Function
                   Else
                      If UCase(Trim(rs("side"))) <> "S" And UCase(Trim(rs("side"))) <> "C" Then
                         MsgBox Trim(rs("Side")) & ":Job side's format is wrong ,the side must be S or C,please define it afresh!", vbCritical, "ErrMessage"
                         Exit Function
                      Else
                         strSide = Trim(rs("Side"))
                      End If
                   End If
                 End If
                 
                If ChkMEBOM_Location = "Y" And Trim(PD(I).location) = "" Then  ''(1250)
                    MsgBox Trim(PD(I).location) & ":location can not be empty,please check", vbCritical, "ErrMessage"
                    Exit Function
                End If
                                 
                strLine = Line           ''(00001) '(1032)
                If idxPL = 0 Then
                    If MCType = NPMMachineType Then    ''(1069)''(1079)
                        ReelWidth = .NPMReelWidth
                    Else
                        ReelWidth = ""
                    End If
                Else
                    If MCType = NPMMachineType Then  ''1108
                        ReelWidth = .NPMReelWidth
                    Else
                        If inArray2(PL, Replace(.compPN, "_", "%")) > 0 Then
                            ReelWidth = PL(1, inArray2(PL, Replace(.compPN, "_", "%")))
                        Else
                            ReelWidth = ""
                        End If
                    End If
                End If
                    strSql = "Insert into QSMS_MEBom(Machine, JobPN,JobGroup,Version,CompPN,LR,Slot,Qty,BuildType," & _
                            "Side,UID,Factory,Line,ReelWidth,Location,DualLaneMode) values('" & .Machine & "','" & .BrdPN & "','" & jobgroup & "'," & _
                            "'" & .Rev & "','" & Replace(.compPN, "_", "%") & "','" & .Side & "','" & .Slot & "'," & _
                            "'" & .Qty & "','" & Trim(BuildType) & "','" & Trim(strSide) & "','" & Trim(g_userName) & "'," & _
                            "'" & Trim(Factory) & "','" & Trim(Line) & "','" & ReelWidth & "','" & Trim(.location) & "','" & Trim(.DualLaneMode) & "') "  '(00004)'(1026)
                  
                    strSqlBomFileName = "Insert into QSMS_MEBom_EQProgram(Machine, JobPN,JobGroup,FullJobGroup,Version,CompPN,LR,Slot,Qty,BuildType," & _
                            "Side,UID,Factory,Line,ReelWidth,Location,DualLaneMode) values('" & .Machine & "','" & .BrdPN & "','" & jobgroup & "','" & strFullJobGroup & "'," & _
                            "'" & .Rev & "','" & Replace(.compPN, "_", "%") & "','" & .Side & "','" & .Slot & "'," & _
                            "'" & .Qty & "','" & Trim(BuildType) & "','" & Trim(strSide) & "','" & Trim(g_userName) & "'," & _
                            "'" & Trim(Factory) & "','" & Trim(Line) & "','" & ReelWidth & "','" & Trim(.location) & "','" & Trim(.DualLaneMode) & "') "  '(1219)
                   Insert_Qty = Insert_Qty + 1
            End With
            Conn.Execute (strSql)
            If ChkEQProgram = "Y" Then   ''(1219)
                Conn.Execute (strSqlBomFileName)
            End If
            Total_Qty = Total_Qty + 1
        End If
    Next I
    
    '''(0005)1214
    strSql = "Delete from NozzleLocation Where Factory='" & Trim(Factory) & "' and Line='" & Trim(Line) & "' and Side='" & Trim(strSide) & "' and BuildType='" & Trim(BuildType) & "' and " & _
            "JobGroup='" & Trim(jobgroup) & "'"
    Conn.Execute (strSql)
    For I = 1 To UBound(NZ)
        strSql = "Insert into NozzleLocation(Factory,Line,Side,BuildType,JobGroup,Machine,Location,NozzleType,UID,TransDateTime)" & _
                "Values('" & Trim(Factory) & "','" & Trim(Line) & "','" & Trim(strSide) & "','" & Trim(BuildType) & "','" & Trim(jobgroup) & "','" & Trim(NZ(I).Machine) & "','" & _
                Trim(NZ(I).location) & "','" & Trim(NZ(I).NozzleType) & "','" & Trim(g_userName) & "',dbo.formatdate(Getdate(),'YYYYMMDDHHNNSS'))"
        Conn.Execute (strSql)
    Next I
    
    strSql = "Insert into QSMS_LOG(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date)" & _
             "Values('SMT_QSMS_PCBSize','" & jobgroup & "','" & Trim(PCBSize) & "','" & Trim(g_userName) & "','0',dbo.formatdate(Getdate(),'YYYYMMDDHHNNSS'))"
    Conn.Execute (strSql)    ''1203
    
    EndTime = Now
'''''1094
Sleep (1000)
'''''1094
strSql = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_PanaAMI End','" & Left(Trim(BomFile), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSql)

LoadDataFile = True
 MsgBox "*** Load  finish ! ***" & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Update succeed : " & Update_Qty & vbCrLf
Exit Function
errHandler:
    MsgBox (Err.Description & vbCrLf & "Item : " & CStr(I) & vbCrLf & vbCrLf & "SQL : " & strSql)
    ErrorDesc = Err.Description & " Item : " & CStr(I) & " SQL : " & strSql
    strSql = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_PanaAMI Error','" & ErrorDesc & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
    Conn.Execute (strSql)
    LoadDataFile = False
End Function

Private Sub Form_Initialize()
  If App.PrevInstance Then
    MsgBox "The program has been running in this machine, this instance will close !"
    End
  End If
End Sub

Private Sub Form_Load()
    StatusBar1.Panels(1) = "QSMS will auto check bom if you upload the PanaMAI machine bom!"
    chkAutChkBom.Value = 1
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub
Private Sub cmdSelect_Click()
  CommonDialog1.ShowOpen
  txtFile = CommonDialog1.FileName
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub SaveLog(ByVal MBPN As String, ByVal Message As String)
Dim FileNumber As Integer
Dim TempString As String
    
    FileNumber = FreeFile
    TempString = String(300, " ")

     Mid(TempString, 1, 25) = Now
     Mid(TempString, 26, 5) = "FAIL"
     Mid(TempString, 31, 35) = "PartNumber" & MBPN
     Mid(TempString, 66, 130) = Message
     Open MEBomErrPath & Format(Date, "YYMMDD") & "ErrorLog.txt" For Append As #FileNumber
     Print #FileNumber, TempString
     Close #FileNumber
    
End Sub

Private Sub AutoCheckBom()               ''(00001)
Dim Sqlstr As String
Dim rs As ADODB.Recordset
If Trim(strJobPN) <> "" And Trim(strRev) <> "" And Trim(strLine) <> "" And Trim(strBuildType) <> "" Then
    Sqlstr = "Exec QSMS_GetPCBWO '" & Trim(strJobPN) & "','" & Trim(strRev) & "','" & Trim(strLine) & "','" & Trim(strBuildType) & "'"
    Set rs = Conn.Execute(Sqlstr)
    If rs.EOF = False Then
        While Not rs.EOF
            LabelRun.BackColor = &HFF&
             LabelRun.Caption = "AutoCheckBom Running..."
            StatusBar1.Panels(1) = "AutoCheckBom:" & Trim(rs("WO"))
            Call GetCheckBomData(Trim(rs("WO")), g_userName, "N")
            rs.MoveNext
        Wend
    End If
End If
End Sub

Private Sub PhraseHeader(ByVal src As String, ByVal log As String)  '1178
''根据log类型初始化相应的数组
Select Case log
    Case "MC"
        MC_Header() = Split(src, " ")
    Case "PD"
        PD_Header() = Split(src, " ")
    Case "PT"
        PT_Header() = Split(src, " ")
    Case "PL"
        PL_Header() = Split(src, " ")
    Case "NZ"
        NZ_Header() = Split(src, " ")
    Case "SD"
        SD_Header() = Split(src, " ")
    Case "BA"
        BA_Header() = Split(src, " ")
    Case "BD"
        BD_Header() = Split(src, " ")
End Select
End Sub

Private Function GetPosition(src() As String, ByVal key As String) As Integer  '1178
''从数组中获取key所在的下标
For I = LBound(src) To UBound(src)
    If UCase(src(I)) = UCase(key) Then
        GetPosition = I
        Exit Function
    End If
Next I
MsgBox ("GetPosition: can not find key:" & key)
End Function
