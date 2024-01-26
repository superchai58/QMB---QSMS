VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCompPNCompare 
   Caption         =   "CompPNCompare [2011/07/12]"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2895
      Left            =   480
      TabIndex        =   2
      Top             =   0
      Width           =   9735
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
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton OptFail 
         Caption         =   "Fail"
         Height          =   375
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptPass 
         Caption         =   "Pass"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDID 
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   2280
         Width           =   4215
      End
      Begin VB.TextBox txtCompPN 
         Height          =   375
         Left            =   6000
         TabIndex        =   10
         Top             =   2280
         Width           =   3735
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
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
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
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
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
         Format          =   98172931
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
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
         Format          =   98172931
         CurrentDate     =   36482
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
         TabIndex        =   13
         Top             =   2280
         Width           =   975
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
         TabIndex        =   11
         Top             =   2280
         Width           =   855
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
         TabIndex        =   9
         Top             =   1320
         Width           =   1455
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
         TabIndex        =   8
         Top             =   1800
         Width           =   1455
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
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   4215
      Left            =   480
      TabIndex        =   0
      Top             =   3000
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7435
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   7320
      Width           =   9735
   End
End
Attribute VB_Name = "FrmCompPNCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

 Dim CheckTimes As Integer

Private Sub cmdexcel_Click()
Dim rs As ADODB.Recordset
If Trim(CboLine.Text) = "" Then
    MsgBox "Please choose Line!"
    Exit Sub
End If
rs = QueryData
If rs.EOF = False Then
    Set DataGrid2.DataSource = Nothing
    Set DataGrid2.DataSource = rs
    DataGrid2.Refresh
    Call CopyToExcel(rs)
Else
    MsgBox "No  Data !"
End If
End Sub

Private Sub CmdQuery_Click()
Dim rs As ADODB.Recordset
If Trim(CboLine.Text) = "" Then
    MsgBox "Please choose Line!"
    Exit Sub
End If
rs = QueryData
If rs.EOF = False Then
    Set DataGrid2.DataSource = Nothing
    DataGrid2.DataSource = rs
    DataGrid2.Refresh
End If
End Sub
Private Function QueryData() As ADODB.Recordset
Dim str As String
Dim rs As ADODB.Recordset
Dim BeginDate, EndDate As String
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    Set DataGrid2.DataSource = Nothing
    str = "SELECT   Line , WO , DID , CompPN , Status ,Side, isnull(Desc1,'') as Desc1 ,UID ,TransDateTime  FROM  QSMS_CompPNCheck order by TransDateTime desc "
    Set rs = Conn.Execute(str)
    QueryData = rs
End Function
Private Sub Form_Load()
    Call GetLine
    CheckTimes = 0
End Sub
Private Function GetLine()
Dim str As String
Dim rs As ADODB.Recordset
    Set DataGrid2.DataSource = Nothing
    str = "select distinct Line from Sap_WO_List where Trans_Date>dbo.FormatDate(getdate()-32,'YYYYMMDDHHNNSS') order by Line asc "
    Set rs = Conn.Execute(str)
    CboLine.Clear
    While Not rs.EOF
        CboLine.AddItem rs!Line
        rs.MoveNext
    Wend
End Function
Private Sub txtCompPN_Click()
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtCompPN_KeyPress(KeyAscii As Integer)
Dim position As Integer
LblStatus.Caption = ""
If KeyAscii = 13 And Trim(txtDID.Text) <> "" And Trim(txtCompPN.Text) <> "" Then
    position = InStr(Trim(txtCompPN.Text), ";") ' ´¦Àí2DµÄComopPN
    If (position > 1) Then
         txtCompPN.Text = Mid(Trim(txtCompPN.Text), 1, position - 1)
    End If
    Call CheckDID
End If
End Sub
Private Sub txtDID_Click()
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtDID_KeyPress(KeyAscii As Integer)
Dim strsql As String
Dim rs As ADODB.Recordset
Dim strsql As String
Dim rs As ADODB.Recordset
LblStatus.Caption = ""
    strsql = "select  DID , CompPN from  QSMS_CompPNCheck where  status<>'FAIL'"
    Set rs = Conn.Execute(strqsql)
    If rs.EOF = False Then
        txtDID.Text = rs("DID")
        txtCompPN.Text = rs("CompPN")
    End If
    If KeyAscii = 13 And txtCompPN.Text <> "" Then
        txtDID.SetFocus
    End If
End Sub

Private Sub CheckDID()
Dim strsql As String
Dim rs As ADODB.Recordset
    strsql = "select  top 1   Line,Side ,CompPN , Work_Order   from QSMS_Dispatch where DID = '" & Trim(txtCompPN.Text) & "'"
    Set rs = Conn.Execute(strsql)
    If rs.EOF = False Then
        If (UCase(Trim(rs("CompPN"))) = UCase(Trim(txtCompPN.Text))) Then
             LblStatus.Caption = "PASS"
        Else
             If CheckTimes = 1 Then
                  LblStatus.Caption = "DID and CompPN is not match , Please Input Again"
                  txtDID.Text = ""
                  txtCompPN.Text = ""
                  CheckTimes = 2
                  Exit Sub
             End If
             LblStatus.Caption = "FAIL"
        End If
        Call SaveData(UCase(Trim(rs("Line"))), UCase(Trim(rs("Side"))), UCase(Trim(rs("Work_Order"))), UCase(Trim(LblStatus.Caption)))
    Else
        LblStatus.Caption = "DID is not exist"
    End If
txtDID.Text = ""
txtCompPN.Text = ""
CheckTimes = 1
End Sub
Private Sub SaveData(Line As String, side As String, WO As String, status As String)
Dim strsql As String
Dim rs As ADODB.Recordset
    strsql = "select  from  QSMS_CompPNCheck where DID='" & Trim(txtDID.Text) & "'"
    Set rs = Conn.Execute(strqsql)
    If rs.EOF = False Then
        LblStatus.Caption = "this DID has been checked !"
    Else
        strsql = "  insert into QSMS_CompPNCheck (Line ,side,  WO , DID , CompPN , Status ,TransDateTime , UID) values('" & Line & "','" & side & "','" & WO & "','" & Trim(txtDID.Text) & "','" & Trim(txtCompPN.Text) & "','" & status & "'," & " dbo.FormatDate (GETDATE (),'YYYYMMDDHHNNDD'),'" & g_userName & "'"
        Conn.Execute (strsql)
    End If
End Sub


Private Sub RefreshData()
Dim strsql As String
Dim rs As ADODB.Recordset
    strsql = "select  top 1000 *  from QSMS_CompPNCheck  order by transdatetime desc"
    Set rs = Conn.Execute(strsql)
    If rs.EOF = False Then
        Set DataGrid2.DataSource = rs
    End If
End Sub
