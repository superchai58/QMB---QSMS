VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmPNCompare 
   BackColor       =   &H80000000&
   Caption         =   "CompPNCompaer(20120712)"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2895
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.CommandButton CmdQuery 
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3720
         Picture         =   "FrmPNCompare1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox CboLine 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtCompPN 
         Height          =   375
         Left            =   6000
         TabIndex        =   3
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtDID 
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   2280
         Width           =   4215
      End
      Begin VB.CommandButton cmdexcel 
         Caption         =   "&Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4920
         Picture         =   "FrmPNCompare1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139460611
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   139460611
         CurrentDate     =   36482
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Line"
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
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Begin Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label LblDID 
         BackColor       =   &H0000FF00&
         Caption         =   "DID"
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
         Left            =   0
         TabIndex        =   9
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label LblCompPN 
         BackColor       =   &H0000FF00&
         Caption         =   "CompPN"
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
         Left            =   5040
         TabIndex        =   8
         Top             =   2280
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   3135
      Left            =   480
      TabIndex        =   13
      Top             =   3120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5530
      _Version        =   393216
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
            LCID            =   2052
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
            LCID            =   2052
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
   Begin VB.Label LblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   14
      Top             =   6480
      Width           =   9735
   End
End
Attribute VB_Name = "FrmPNCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

 Dim strDoubleCheck As Boolean
 Dim strDelaytime As Long
Private Sub CmdExcel_Click()
Dim Rs As ADODB.Recordset
If Trim(CboLine.text) = "" Then
    MsgBox "Please choose Line!"
    Exit Sub
End If
Set Rs = QueryData
If Rs.EOF = False Then
    Set DataGrid2.DataSource = Nothing
    Set DataGrid2.DataSource = Rs
    DataGrid2.Refresh
    Call CopyToExcel(Rs)
Else
    MsgBox "No  Data !"
End If
End Sub

Private Sub CmdQuery_Click()
Dim Rs As ADODB.Recordset
If Trim(CboLine.text) = "" Then
    MsgBox "Please choose Line!"
    Exit Sub
End If
Set Rs = QueryData
If Rs.EOF = False Then
    Set DataGrid2.DataSource = Nothing
    Set DataGrid2.DataSource = Rs
    DataGrid2.Refresh
End If
End Sub
Private Function QueryData() As ADODB.Recordset
Dim str As String
Dim Rs As ADODB.Recordset
Dim BeginDate, EndDate As String
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    Set DataGrid2.DataSource = Nothing
    str = "SELECT  GroupID ,   Line , WO , DID , CompPN , Status ,Side, UID ,TransDateTime   FROM  QSMS_CompPNCheck where  Line='" & Trim(CboLine.text) & "'   and  TransDateTime>='" & Trim(BeginDate) & "000000'  and  TransDateTime<='" & Trim(EndDate) & "235900' order by TransDateTime desc "
    Set Rs = Conn.Execute(str)
    Set QueryData = Rs
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
Dim strSQL As String
Dim Rs As ADODB.Recordset
    FrmPNCompare.Caption = FrmPNCompare.Caption + "   Line : " + ProgLine
    dtpSDate.Value = Now
    dtpEDate.Value = Now
    Call GetLine
    If CheckDIDCheckStatus = False Then
        TxtCompPN.Locked = True
        TxtDID.Locked = True
        lblstatus.BackColor = &HFF&
        lblstatus.ForeColor = &H8000000E
    End If
    Call Hook(TxtCompPN.hWnd)  ''''1102
    Call Hook(TxtDID.hWnd)
End Sub
Private Function GetLine()
Dim str As String
Dim Rs As ADODB.Recordset
    Set DataGrid2.DataSource = Nothing
    str = "select distinct Line from Sap_WO_List where Trans_Date>dbo.FormatDate(getdate()-32,'YYYYMMDDHHNNSS') order by Line asc "
    Set Rs = Conn.Execute(str)
    CboLine.Clear
    While Not Rs.EOF
        CboLine.AddItem Rs!Line
        Rs.MoveNext
    Wend
End Function
Private Function CheckDIDCheckStatus() As Boolean
    strSQL = "select  * from QSMS_CompPNCheck where Status= 'FAIL'  and Line ='" & ProgLine & "'"
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = False Then
        lblstatus.Caption = " This DID and CompPN is not match ,please unlock at first ! "
        TxtDID.text = Rs("DID")
        TxtCompPN.text = Rs("CompPN")
        CheckDIDCheckStatus = False
        Exit Function
    Else
        CheckDIDCheckStatus = True
        Exit Function
    End If
End Function
Private Function CheckDIDExist() As Boolean
    strSQL = "select  0 from QSMS_CompPNCheck where DID= '" & Trim(TxtDID.text) & "'"
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = False Then
        TxtDID.text = ""
        TxtCompPN.text = ""
        CheckDIDExist = True
        Exit Function
    Else
        CheckDIDExist = False
        Exit Function
    End If
End Function

Private Sub txtCompPN_Click()
Sendkeys "{HOME}+{END}"
End Sub


Private Sub txtCompPN_KeyPress(KeyAscii As Integer)
Dim position As Integer
Dim strSQL As String
Dim strstatus  As String
Dim Rs As ADODB.Recordset
lblstatus.Caption = ""
'''''''''''''''''''1102
        If strDelaytime <> 0 Then
            If GetTickCount - strDelaytime > 100 Then
                MsgBox "Please use scaner!"
                TxtCompPN.text = ""
                strDelaytime = 0
                Call txtCompPN_Click
                Exit Sub
            End If
        End If
    strDelaytime = GetTickCount
    If KeyAscii = 13 Or KeyAscii = 9 Then
         strDelaytime = 0
    End If
''''''''''''''''''''''1102
    If KeyAscii = 13 And Trim(TxtDID.text) <> "" And Trim(TxtCompPN.text) <> "" Then

        position = InStr(Trim(TxtCompPN.text), ";") ' 处理2D的ComopPN
        If (position > 1) Then
            TxtCompPN.text = Mid(Trim(TxtCompPN.text), 1, position - 1)
        End If
        If CheckDIDCheckStatus = False Then   '---check status is fail
            TxtCompPN.Locked = True
            TxtDID.Locked = True
            lblstatus.BackColor = &HFF&
            lblstatus.ForeColor = &H8000000E
            
            Exit Sub
        End If
        If CheckDIDExist = True Then
            lblstatus.Caption = Trim(TxtDID.text) & " This DID  has been checked ! "
            lblstatus.BackColor = &HFF&
            lblstatus.ForeColor = &H8000000E
            MsgBox Trim(TxtDID.text) & " This DID  has been checked ! "
            Exit Sub
        End If
       strSQL = "select top 1  CompPN ,  Line , side ,Work_Order ,GroupID  from  QSMS_Dispatch  where DID ='" & Trim(TxtDID.text) & "'"
       Set Rs = Conn.Execute(strSQL)
       If Rs.EOF = False Then
           If Rs("Line") = ProgLine Then
                If UCase(Trim(Rs("compPN"))) = UCase(Trim(TxtCompPN.text)) Then
                    strstatus = "PASS"
                    lblstatus.Caption = strstatus
                    lblstatus.BackColor = &HFF00&
                Else
'                      txtCompPN.Text = InputBox("Please Input CompPN again ")
'                      position = InStr(Trim(txtCompPN.Text), ";") ' 处理2D的ComopPN
'                      If (position > 1) Then
'                          txtCompPN.Text = Mid(Trim(txtCompPN.Text), 1, position - 1)
'                      End If
'
'                      If UCase(Trim(rs("compPN"))) = UCase(Trim(txtCompPN.Text)) Then
'                            strstatus = "PASS"
'                            LblStatus.Caption = strstatus
'                      Else
                            strstatus = "FAIL"
                            lblstatus.Caption = TxtDID.text & " and  " & TxtCompPN.text & " is not match, please unlock at first ! "
                            TxtDID.Locked = True
                            TxtCompPN.Locked = True
                            lblstatus.BackColor = &HFF&
                            lblstatus.ForeColor = &H8000000E
'                     End If
                End If
           Else
                lblstatus.Caption = Trim(TxtDID.text) & " belong to " & Trim(Rs("Line"))
                TxtDID.SetFocus
                TxtDID.text = ""
                TxtCompPN.text = ""
                Exit Sub
           End If
       Else
            lblstatus.Caption = Trim(TxtDID.text) & "  is not exist in Dispatch   "
            TxtDID.SetFocus
            TxtDID.text = ""
            TxtCompPN.text = ""
            Exit Sub
       End If
       Call SaveData(ProgLine, UCase(Trim(Rs("Side"))), UCase(Trim(Rs("Work_Order"))), UCase(Trim(Rs("GroupID"))), strstatus)
        TxtDID.SetFocus
        If strstatus <> "FAIL" Then
            TxtDID.text = ""
            TxtCompPN.text = ""
        End If
    End If
End Sub
Private Sub txtDID_Click()
Sendkeys "{HOME}+{END}"
End Sub
Private Sub txtDID_KeyPress(KeyAscii As Integer)
Dim strSQL As String
Dim Rs As ADODB.Recordset
lblstatus.Caption = ""
lblstatus.BackColor = &HC0C0FF     '&H00C0C0FF&
lblstatus.ForeColor = &H80000012
Dim strLine As String
'''''''''''''''''1102
    If strDelaytime <> 0 Then
        If GetTickCount - strDelaytime > 100 Then
            MsgBox "Please use scaner!"
            TxtDID.text = ""
            strDelaytime = 0
            Call txtDID_Click
            Exit Sub
        End If
    End If
 strDelaytime = GetTickCount
If KeyAscii = 13 Or KeyAscii = 9 Then
     strDelaytime = 0
End If
'''''''''''''''''1102
If KeyAscii = 13 And Trim(TxtDID.text) <> "" Then
    If CheckDIDCheckStatus = False Then
        TxtCompPN.Locked = True
        TxtDID.Locked = True
        lblstatus.BackColor = &HFF&
        lblstatus.ForeColor = &H8000000E
        Exit Sub
    End If
    strLine = GetDIDLine
    If strLine <> "" Then
        If strLine <> ProgLine Then
            TxtDID.text = ""
            TxtDID.SetFocus
            
            lblstatus.Caption = " The DID belongs to " & strLine & " Line, please comfirm again."
            lblstatus.BackColor = &HFF&
            lblstatus.ForeColor = &H8000000E
            MsgBox " The DID belongs to " & strLine & " Line, please comfirm again."
            TxtDID.text = ""
            Exit Sub
        End If
    Else
            
            lblstatus.Caption = " The DID does not exist in Dispatch Table, please comfirm again."
            lblstatus.BackColor = &HFF&
            lblstatus.ForeColor = &H8000000E
            MsgBox " The DID does not exist in Dispatch Table, please comfirm again."
            TxtDID.text = ""
            
            Exit Sub
    End If
    If CheckDIDExist = False Then
        TxtCompPN.SetFocus
    Else
        lblstatus.Caption = Trim(TxtDID.text) & " This DID  has been checked ! "
        lblstatus.BackColor = &HFF&
        lblstatus.ForeColor = &H8000000E
        MsgBox Trim(TxtDID.text) & " This DID  has been checked ! "
         '        LblStatus.Caption = Trim(txtDID.Text) & " This DID  has been checked ! "
        TxtDID.SetFocus
    End If
End If
End Sub

Private Function GetDIDLine() As String
Dim strSQL As String
Dim Rs As ADODB.Recordset
strSQL = "select  top 1  Line  from QSMS_Dispatch where DID = '" & Trim(TxtDID.text) & "'"
Set Rs = Conn.Execute(strSQL)
If Rs.EOF = False Then
     GetDIDLine = UCase(Trim(Rs("Line")))   ''''1137
Else
    GetDIDLine = ""
End If
End Function
 
Private Sub SaveData(Line As String, Side As String, WO As String, GroupID As String, status As String)
Dim strSQL As String
Dim Rs As ADODB.Recordset
strSQL = "insert into QSMS_CompPNCheck ( GroupID , Line ,side,  WO , DID , CompPN , Status ,TransDateTime , UID) values('" & UCase(GroupID) & "','" & UCase(Line) & "','" & UCase(Side) & "','" & WO & "','" & UCase(Trim(TxtDID.text)) & "','" & UCase(Trim(TxtCompPN.text)) & "','" & status & "'," & " dbo.FormatDate (GETDATE (),'YYYYMMDDHHNNDD'),'" & g_userName & "')"
Conn.Execute (strSQL)
strSQL = "select  top 1000 *  from QSMS_CompPNCheck order by transdatetime desc"
Set Rs = Conn.Execute(strSQL)

If Rs.EOF = False Then
    Set DataGrid2.DataSource = Nothing
    Set DataGrid2.DataSource = Rs
    DataGrid2.Refresh
End If
 
End Sub


Private Sub reFreshData()
Dim strSQL As String
Dim Rs As ADODB.Recordset
    strSQL = "select  top 1000 *  from QSMS_CompPNCheck  order by transdatetime desc"
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = False Then
        
        Set DataGrid2.DataSource = Rs
    End If
End Sub
