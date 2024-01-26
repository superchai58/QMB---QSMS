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
         Format          =   143327235
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
         Format          =   143327235
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
Private Sub cmdexcel_Click()
Dim RS As ADODB.Recordset
If Trim(CboLine.Text) = "" Then
    MsgBox "Please choose Line!"
    Exit Sub
End If
Set RS = QueryData
If RS.EOF = False Then
    Set DataGrid2.DataSource = Nothing
    Set DataGrid2.DataSource = RS
    DataGrid2.Refresh
    Call CopyToExcel(RS)
Else
    MsgBox "No  Data !"
End If
End Sub

Private Sub CmdQuery_Click()
Dim RS As ADODB.Recordset
If Trim(CboLine.Text) = "" Then
    MsgBox "Please choose Line!"
    Exit Sub
End If
Set RS = QueryData
If RS.EOF = False Then
    Set DataGrid2.DataSource = Nothing
    Set DataGrid2.DataSource = RS
    DataGrid2.Refresh
End If
End Sub
Private Function QueryData() As ADODB.Recordset
Dim str As String
Dim RS As ADODB.Recordset
Dim BeginDate, EndDate As String
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    Set DataGrid2.DataSource = Nothing
    str = "SELECT  GroupID ,   Line , WO , DID , CompPN , Status ,Side, UID ,TransDateTime   FROM  QSMS_CompPNCheck where  Line='" & Trim(CboLine.Text) & "'   and  TransDateTime>='" & Trim(BeginDate) & "000000'  and  TransDateTime<='" & Trim(EndDate) & "235900' order by TransDateTime desc "
    Set RS = Conn.Execute(str)
    Set QueryData = RS
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
Dim strsql As String
Dim RS As ADODB.Recordset
    FrmPNCompare.Caption = FrmPNCompare.Caption + "   Line : " + ProgLine
    dtpSDate.Value = Now
    dtpEDate.Value = Now
    Call GetLine
    If CheckDIDCheckStatus = False Then
        txtCompPN.Locked = True
        txtDID.Locked = True
        LblStatus.BackColor = &HFF&
        LblStatus.ForeColor = &H8000000E
    End If
    Call Hook(txtCompPN.hWnd)  ''''1102
    Call Hook(txtDID.hWnd)
End Sub
Private Function GetLine()
Dim str As String
Dim RS As ADODB.Recordset
    Set DataGrid2.DataSource = Nothing
    str = "select distinct Line from Sap_WO_List where Trans_Date>dbo.FormatDate(getdate()-32,'YYYYMMDDHHNNSS') order by Line asc "
    Set RS = Conn.Execute(str)
    CboLine.Clear
    While Not RS.EOF
        CboLine.AddItem RS!Line
        RS.MoveNext
    Wend
End Function
Private Function CheckDIDCheckStatus() As Boolean
    strsql = "select  * from QSMS_CompPNCheck where Status= 'FAIL'  and Line ='" & ProgLine & "'"
    Set RS = Conn.Execute(strsql)
    If RS.EOF = False Then
        LblStatus.Caption = " This DID and CompPN is not match ,please unlock at first ! "
        txtDID.Text = RS("DID")
        txtCompPN.Text = RS("CompPN")
        CheckDIDCheckStatus = False
        Exit Function
    Else
        CheckDIDCheckStatus = True
        Exit Function
    End If
End Function
Private Function CheckDIDExist() As Boolean
    strsql = "select  0 from QSMS_CompPNCheck where DID= '" & Trim(txtDID.Text) & "'"
    Set RS = Conn.Execute(strsql)
    If RS.EOF = False Then
        txtDID.Text = ""
        txtCompPN.Text = ""
        CheckDIDExist = True
        Exit Function
    Else
        CheckDIDExist = False
        Exit Function
    End If
End Function

Private Sub txtCompPN_Click()
SendKeys "{HOME}+{END}"
End Sub


Private Sub txtCompPN_KeyPress(KeyAscii As Integer)
Dim position As Integer
Dim strsql As String
Dim strstatus  As String
Dim RS As ADODB.Recordset
LblStatus.Caption = ""
'''''''''''''''''''1102
        If strDelaytime <> 0 Then
            If GetTickCount - strDelaytime > 100 Then
                MsgBox "Please use scaner!"
                txtCompPN.Text = ""
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
    If KeyAscii = 13 And Trim(txtDID.Text) <> "" And Trim(txtCompPN.Text) <> "" Then

        position = InStr(Trim(txtCompPN.Text), ";") ' 处理2D的ComopPN
        If (position > 1) Then
            txtCompPN.Text = Mid(Trim(txtCompPN.Text), 1, position - 1)
        End If
        If CheckDIDCheckStatus = False Then   '---check status is fail
            txtCompPN.Locked = True
            txtDID.Locked = True
            LblStatus.BackColor = &HFF&
            LblStatus.ForeColor = &H8000000E
            
            Exit Sub
        End If
        If CheckDIDExist = True Then
            LblStatus.Caption = Trim(txtDID.Text) & " This DID  has been checked ! "
            LblStatus.BackColor = &HFF&
            LblStatus.ForeColor = &H8000000E
            MsgBox Trim(txtDID.Text) & " This DID  has been checked ! "
            Exit Sub
        End If
       strsql = "select top 1  CompPN ,  Line , side ,Work_Order ,GroupID  from  QSMS_Dispatch  where DID ='" & Trim(txtDID.Text) & "'"
       Set RS = Conn.Execute(strsql)
       If RS.EOF = False Then
           If RS("Line") = ProgLine Then
                If UCase(Trim(RS("compPN"))) = UCase(Trim(txtCompPN.Text)) Then
                    strstatus = "PASS"
                    LblStatus.Caption = strstatus
                    LblStatus.BackColor = &HFF00&
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
                            LblStatus.Caption = txtDID.Text & " and  " & txtCompPN.Text & " is not match, please unlock at first ! "
                            txtDID.Locked = True
                            txtCompPN.Locked = True
                            LblStatus.BackColor = &HFF&
                            LblStatus.ForeColor = &H8000000E
'                     End If
                End If
           Else
                LblStatus.Caption = Trim(txtDID.Text) & " belong to " & Trim(RS("Line"))
                txtDID.SetFocus
                txtDID.Text = ""
                txtCompPN.Text = ""
                Exit Sub
           End If
       Else
            LblStatus.Caption = Trim(txtDID.Text) & "  is not exist in Dispatch   "
            txtDID.SetFocus
            txtDID.Text = ""
            txtCompPN.Text = ""
            Exit Sub
       End If
       Call SaveData(ProgLine, UCase(Trim(RS("Side"))), UCase(Trim(RS("Work_Order"))), UCase(Trim(RS("GroupID"))), strstatus)
        txtDID.SetFocus
        If strstatus <> "FAIL" Then
            txtDID.Text = ""
            txtCompPN.Text = ""
        End If
    End If
End Sub
Private Sub txtDID_Click()
SendKeys "{HOME}+{END}"
End Sub
Private Sub txtDID_KeyPress(KeyAscii As Integer)
Dim strsql As String
Dim RS As ADODB.Recordset
LblStatus.Caption = ""
LblStatus.BackColor = &HC0C0FF     '&H00C0C0FF&
LblStatus.ForeColor = &H80000012
Dim strLine As String
'''''''''''''''''1102
    If strDelaytime <> 0 Then
        If GetTickCount - strDelaytime > 100 Then
            MsgBox "Please use scaner!"
            txtDID.Text = ""
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
If KeyAscii = 13 And Trim(txtDID.Text) <> "" Then
    If CheckDIDCheckStatus = False Then
        txtCompPN.Locked = True
        txtDID.Locked = True
        LblStatus.BackColor = &HFF&
        LblStatus.ForeColor = &H8000000E
        Exit Sub
    End If
    strLine = GetDIDLine
    If strLine <> "" Then
        If strLine <> ProgLine Then
            txtDID.Text = ""
            txtDID.SetFocus
            
            LblStatus.Caption = " 此DID属于" & strLine & " 线，请再次确认"
            LblStatus.BackColor = &HFF&
            LblStatus.ForeColor = &H8000000E
            MsgBox " 此DID属于" & strLine & "，请再次确认"
            txtDID.Text = ""
            Exit Sub
        End If
    Else
            
            LblStatus.Caption = " 此DID 在Dispatch 中不存在 ，请再次确认"
            LblStatus.BackColor = &HFF&
            LblStatus.ForeColor = &H8000000E
            MsgBox " 此DID 在Dispatch 中不存在 ，请再次确认"
            txtDID.Text = ""
            
            Exit Sub
    End If
    If CheckDIDExist = False Then
        txtCompPN.SetFocus
    Else
        LblStatus.Caption = Trim(txtDID.Text) & " This DID  has been checked ! "
        LblStatus.BackColor = &HFF&
        LblStatus.ForeColor = &H8000000E
        MsgBox Trim(txtDID.Text) & " This DID  has been checked ! "
         '        LblStatus.Caption = Trim(txtDID.Text) & " This DID  has been checked ! "
        txtDID.SetFocus
    End If
End If
End Sub

Private Function GetDIDLine() As String
Dim strsql As String
Dim RS As ADODB.Recordset
strsql = "select  top 1  Line  from QSMS_Dispatch where DID = '" & Trim(txtDID.Text) & "'"
Set RS = Conn.Execute(strsql)
If RS.EOF = False Then
     GetDIDLine = UCase(Trim(RS("Line")))   ''''1137
Else
    GetDIDLine = ""
End If
End Function
 
Private Sub SaveData(Line As String, Side As String, WO As String, GroupID As String, status As String)
Dim strsql As String
Dim RS As ADODB.Recordset
strsql = "  insert into QSMS_CompPNCheck ( GroupID , Line ,side,  WO , DID , CompPN , Status ,TransDateTime , UID) values('" & UCase(GroupID) & "','" & UCase(Line) & "','" & UCase(Side) & "','" & WO & "','" & UCase(Trim(txtDID.Text)) & "','" & UCase(Trim(txtCompPN.Text)) & "','" & status & "'," & " dbo.FormatDate (GETDATE (),'YYYYMMDDHHNNDD'),'" & g_userName & "')"
Conn.Execute (strsql)
strsql = " select  top 1000 *  from QSMS_CompPNCheck  order by transdatetime desc"
Set RS = Conn.Execute(strsql)

If RS.EOF = False Then
    Set DataGrid2.DataSource = Nothing
    Set DataGrid2.DataSource = RS
    DataGrid2.Refresh
End If
 
End Sub


Private Sub RefreshData()
Dim strsql As String
Dim RS As ADODB.Recordset
    strsql = "select  top 1000 *  from QSMS_CompPNCheck  order by transdatetime desc"
    Set RS = Conn.Execute(strsql)
    If RS.EOF = False Then
        
        Set DataGrid2.DataSource = RS
    End If
End Sub
