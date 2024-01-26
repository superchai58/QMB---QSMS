VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGetMEBom 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TransPanaMAI 20060702"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10935
   Icon            =   "GetMEBom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtFile 
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton ComdS 
      Caption         =   "Search"
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
      Left            =   3120
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox CombMBPN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton cmdGetMEBom 
      Caption         =   "GetMEBom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8760
      TabIndex        =   1
      Top             =   120
      Width           =   1395
   End
   Begin VB.Timer Timer 
      Interval        =   60000
      Left            =   6000
      Top             =   3000
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5235
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   9234
      _Version        =   393216
      BackColor       =   8454143
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6645
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6271
            MinWidth        =   5292
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6271
            MinWidth        =   5292
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   6624
            MinWidth        =   5645
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
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LabelRun 
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   6000
      Width           =   10935
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmGetMEBom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim Interval As Integer

Private Type Machine
    MCName As String
    MCNo As String
    HeadNo As String
End Type

Private Type PartData
    IdNo As Integer
    PN As String
End Type

Private Type PositionData
    Machine As String
    BrdPN As String
    Rev As String
    PU As String
    Table As String
    Slot As String
    Side As Integer
    Head As Integer
    Parts As Integer
    CompPN As String
    Qty As Integer
    Enabled As Boolean
End Type


Private Sub cmdGetMEBom_Click()
    If Trim(txtFile) = "" Then
        MsgBox "You must select a file!!", vbInformation
        Exit Sub
    End If
        
    If LoadDataFile(Trim(txtFile)) = False Then
        MsgBox ("Fail")
    Else
        MsgBox ("Finish")
    End If
    
    LabelRun.BackColor = &HFFFFC0
    LabelRun.Caption = "OK"
    StatusBar1.Panels(1) = txtFile & " OK"
    StatusBar1.Panels(3) = "Finished DateTime:" & Now
    'Call GetMEBom(MBPN)
    Call GetMBPN
    i = 0
End Sub
    
Function LoadDataFile(strFile As String) As Boolean
Dim NFile As Integer, i As Integer, j As Integer
Dim Arry() As String, arry2() As String, Log As String, Flg As String
Dim MBPN As String, Revision As String, Machine As String, MCNo As String, Head As String, Parts As String, PU As String, Side As String, IDNUM As String, PartsName As String
Dim BomFile As String, strCurrent As String
Dim StartTime As String, EndTime As String, BoardType As String
Dim strSql As String, BrdPN As String, BrdRev As String
Dim idxMC As Integer, idxPD As Integer, idxPT As Integer
Dim MC() As Machine, PD() As PositionData, PT() As PartData
'On Error GoTo errHandler


LoadDataFile = False
Arry = Split(strFile, "\")
BomFile = Arry(UBound(Arry))

If InStr(BomFile, "-") = 0 Then
    MsgBox ("Filename format must be PN-Rev !")
    Exit Function
End If

If BomFile <> "" Then
    MBPN = Left(BomFile, InStr(BomFile, "-") - 1)
    Revision = Mid(BomFile, InStr(BomFile, "-") + 1, 3)
    
'    strSql = "delete from QSMS_MEBom_Machines where MBPN=" & sq(MBPN) & " and Rev=" & sq(Revision)
'    cnn.Execute (strSql)
    
    NFile = FreeFile
    StartTime = Now
    LabelRun.BackColor = &HFF&
    LabelRun.Caption = "Running..."
    Open strFile For Input As #NFile
    StatusBar1.Panels(1) = "GetMEBom_" & BomFile
    StatusBar1.Panels(2) = "Start DateTime:" & StartTime
    idxMC = 0: idxPD = 0: idxPT = 0
    While Not EOF(NFile)
        i = i + 1
        Line Input #NFile, strCurrent
        Select Case Replace(strCurrent, vbCrLf, "")
            Case "[Machines]"
                 Log = "MC"
                 Line Input #NFile, strCurrent
                 i = i + 1
            Case "[PositionData]"
                 Log = "PD"
                 Line Input #NFile, strCurrent
                 i = i + 1
            Case "[PartsData]"
                 Log = "PT"
                 Line Input #NFile, strCurrent
                 i = i + 1
            Case Else
                Arry = Split(strCurrent, " ")
                If UBound(Arry()) > 3 Then
                      Flg = "OK"
                Else
                   Log = ""
                   Flg = "Canel"
                End If
                Select Case Log + Flg
                    Case "MCOK"
                        idxMC = idxMC + 1
                        ReDim Preserve MC(idxMC)
                        
                        MC(idxMC).MCName = Trim(Arry(1))
                        MC(idxMC).MCNo = Trim(Arry(2))
                        MC(idxMC).HeadNo = Trim(Arry(3))
                        strSql = "delete QSMS_MEBom where Machine=" & sq(MC(idxMC).MCName & MC(idxMC).MCNo)
                        cnn.Execute (strSql)
                    Case "PDOK"
                        idxPD = idxPD + 1
                        ReDim Preserve PD(idxPD)
                        
                        PD(idxPD).Parts = Trim(Arry(5))
                        PD(idxPD).PU = Trim(Arry(11))
                        PD(idxPD).Side = Trim(Arry(12))
                        PD(idxPD).Head = Trim(Arry(13))
                        PD(idxPD).Qty = 1
                        If Len(PD(idxPD).PU) >= 4 Then
                            PD(idxPD).Enabled = True
                        End If
                        BoardType = Trim(Arry(16))
                        arry2 = Split(BoardType, "-")
                        If UBound(arry2) > 0 Then
                            If UBound(arry2) <> 2 Then
                                MsgBox "BoardType wrong : " & BoardType & ", must be PN-REV!"
                                Exit Function
                            End If
                                
                            PD(idxPD).BrdPN = Trim(arry2(1))
                            PD(idxPD).Rev = Replace(Trim(arry2(2)), """", "")
                        Else
                            PD(idxPD).BrdPN = MBPN
                            PD(idxPD).Rev = Revision
                        End If
'                        strSql = "Insert into QSMS_MEBom_PositionData(MBPN,Rev,PU,Side,Parts) values('" & BrdPN & "','" & BrdRev & "','" & PU & "','" & Side & "','" & Parts & "')"
'                        cnn.Execute (strSql)
                    Case "PTOK"
                        idxPT = idxPT + 1
                        ReDim Preserve PT(idxPT)
                        PT(idxPT).IdNo = Trim(Arry(0))
                        PT(idxPT).PN = Trim(Arry(1))
                End Select
        End Select
    Wend
    Close #NFile

    
    i = 0
    'Get CompPN from [PartsData] and set machine from [Machine]
    Do While i < UBound(PD)
        If PD(i).Head > 0 Then
            PD(i).CompPN = Replace(PT(PD(i).Parts).PN, """", "")    'remove quote
            PD(i).Machine = MC(PD(i).Head).MCName & MC(PD(i).Head).MCNo
            PD(i).Table = Trim(CStr(PD(i).Head - 4 * (MC(PD(i).Head).MCNo - 1)))
            PD(i).Slot = PD(i).Table & "-" & Replace(Right(Trim(PD(i).PU), 4), "0", "")
        End If
        i = i + 1
    Loop
    
    'Count Comp Qty group by same machine, slot, side, Job, Rev
    For i = 1 To UBound(PD)
        If PD(i).Enabled Then
            For j = i + 1 To UBound(PD)
                If PD(j).Enabled And PD(i).BrdPN = PD(j).BrdPN And PD(i).CompPN = PD(j).CompPN And _
                    PD(i).Machine = PD(j).Machine And PD(i).Slot = PD(j).Slot And _
                    PD(i).Side = PD(j).Side Then
                    
                    PD(j).Enabled = False
                    PD(i).Qty = PD(i).Qty + 1
                    
                End If
            Next j
            With PD(i)
                strSql = "Insert into QSMS_MEBom(Machine, JobPN,Version,CompPN,LR,Slot,Qty) values('" & .Machine & _
                    "','" & .BrdPN & "','" & .Rev & "','" & .CompPN & "','" & .Side & "','" & .Slot & "','" & .Qty & "')"
            End With
            cnn.Execute (strSql)
        End If
    Next i
    
    If Dir(strFile, vbNormal) <> "" Then
       FileCopy strFile, MEBomBKPath & BomFile
       Kill strFile
    End If
    EndTime = Now
End If
LoadDataFile = True
Exit Function
errHandler:
    MsgBox (Err.Description)
    LoadDataFile = False
End Function

Private Sub GetMEBom(MBPN As String)
Dim strSql As String
Dim rsTemp As New ADODB.Recordset
cnn.CursorLocation = adUseClient
   strSql = "Select * from QSMS_MEBom where JobPN like '" & MBPN & "%' order by JobPN,Version,Machine"
   If rsTemp.State Then rsTemp.Close
   rsTemp.CursorLocation = adUseClient
   Set rsTemp = cnn.Execute(strSql)
   Set DataGrid1.DataSource = rsTemp
End Sub



Private Sub Label1_Click()

End Sub

Private Sub ComdS_Click()
     Call GetMEBom(Trim(CombMBPN))
End Sub

Private Sub Timer_Timer()
On Error GoTo errHandler
i = i + 1
  If i > Interval - 1 Then
     Call cmdGetMEBom_Click
  End If
Exit Sub
errHandler:
    LabelRun.BackColor = &HFF&
    LabelRun.Caption = Err.Description
End Sub

Private Sub Form_Load()
    Interval = ReadIniFile("System", "Interval", App.Path & "\set.ini")
    MEBomPath = Trim(ReadIniFile("System", "MEBomPath", App.Path & "\set.ini"))
    If MEBomPath = "" Then
         MsgBox ("Can't get [SYSTEM] MEBomPath!")
         End
    End If
    connStr = ReadIniFile("database", "connection", App.Path & "\set.ini")
    cnn.Open connStr
    
    MEBomBKPath = MEBomPath + "\Bak\"
    MEBomErrPath = MEBomPath + "\Error\"
    
    If Dir(MEBomPath, vbDirectory) = "" Then MkDir MEBomPath
    If Dir(MEBomBKPath, vbDirectory) = "" Then MkDir MEBomBKPath
    If Dir(MEBomErrPath, vbDirectory) = "" Then MkDir MEBomErrPath
    Call GetMBPN
End Sub

Private Sub GetMBPN()
Dim strSql As String
Dim rs As New ADODB.Recordset
    strSql = "Select distinct JobPN from QSMS_MEBom order by JobPN"
    If rs.State Then rs.Close
    rs.CursorLocation = adUseClient
    rs.Open strSql, cnn, adOpenForwardOnly, adLockReadOnly
    CombMBPN.Clear
    While Not rs.EOF
        CombMBPN.AddItem rs("JobPN")
        rs.MoveNext
    Wend
End Sub

Private Sub Form_Initialize()
  If App.PrevInstance Then
    MsgBox "The program has been running in this machine, this instance will close !"
    End
  End If
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

Private Sub Form_Unload(Cancel As Integer)
    
Dim ans
ans = MsgBox("If close program, SMT QSMS_MEBom will fail! Are you sure?", vbYesNo)

If ans = vbNo Then
    Cancel = True
Else
    Close
    End
End If

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
