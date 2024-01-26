VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmUnlockCompPNCompare 
   Caption         =   "Unlock"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Query"
      Height          =   3975
      Left            =   600
      TabIndex        =   9
      Top             =   3600
      Width           =   11175
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2895
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5106
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
      Begin VB.CommandButton CmdQuery 
         Caption         =   "Query"
         Height          =   735
         Left            =   9960
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   615
         Left            =   9960
         TabIndex        =   16
         Top             =   1920
         Width           =   975
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   1965
         _ExtentX        =   3466
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
         Format          =   51511299
         CurrentDate     =   39404
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   7560
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   1965
         _ExtentX        =   3466
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
         Format          =   51511299
         CurrentDate     =   39404
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
         TabIndex        =   15
         Top             =   240
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
         Left            =   6000
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "BeginDate"
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
         Index           =   2
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Unlock"
      Height          =   3375
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3015
         Left            =   6360
         TabIndex        =   18
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   5318
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
      Begin VB.CommandButton CmdUnlock 
         Caption         =   "Unlock"
         Height          =   735
         Left            =   5400
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtDID 
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox txtCompPN 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox TxtReason 
         Height          =   975
         Left            =   1080
         TabIndex        =   1
         Top             =   1680
         Width           =   4215
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
         Left            =   120
         TabIndex        =   7
         Top             =   1200
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
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "Reason"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Lblstatus 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   4
         Top             =   2760
         Width           =   4215
      End
   End
End
Attribute VB_Name = "FrmUnlockCompPNCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexcel_Click()


Dim BeginDate, EndDate As String, str As String
Dim rs As ADODB.Recordset
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    Set DataGrid2.DataSource = Nothing
    If Trim(CboLine.Text) <> "" Then
        str = " select top 1000 * from  QSMS_UnlockCompPNCheck where  Line='" & Trim(CboLine.Text) & "'   and  TransDateTime>='" & Trim(BeginDate) & "000000'  and  TransDateTime<='" & Trim(EndDate) & "235900' order by TransDateTime desc "
        Set rs = Conn.Execute(str)
        If rs.EOF = False Then
            Call CopyToExcel(rs)
        Else
            MsgBox "No  Data !"
        End If
    Else
        MsgBox "No  Data !"
    End If
 
End Sub
 

Private Sub CmdQuery_Click()
Dim BeginDate, EndDate As String, str As String
Dim rs As ADODB.Recordset
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    Set DataGrid2.DataSource = Nothing
    If Trim(CboLine.Text) <> "" Then
        str = " select top 1000 * from  QSMS_UnlockCompPNCheck where  Line='" & Trim(CboLine.Text) & "'   and  TransDateTime>='" & Trim(BeginDate) & "000000'  and  TransDateTime<='" & Trim(EndDate) & "235900' order by TransDateTime desc "
        Set rs = Conn.Execute(str)
        If rs.EOF = False Then
            Set DataGrid2.DataSource = rs
            DataGrid2.Refresh
        End If
    Else
        MsgBox "Please  choose Line:"
    End If
End Sub

Private Sub CmdUnlock_Click()
Dim strsql As String
Dim rs As ADODB.Recordset
Dim transdatetime As String
If TxtReason.Text = "" Or txtCompPN.Text = "" Then
    Lblstatus.Caption = "Reason or PN is not null"
    Exit Sub
Else
    strsql = "insert into QSMS_UnlockCompPNCheck(GroupID , Line ,WO , DID , OLDCompPN ,NewCompPN ,Side ,Reason ,TransDateTime , UID) " & _
    " select GroupID , Line ,WO , DID , CompPN , '" & Trim(txtCompPN.Text) & "'  ,Side , N'" & TxtReason.Text & "', dbo.formatdate(getdate(),'yyyymmddhhnnss'), '" & g_userName & "'  from QSMS_CompPNCheck where DID = '" & Trim(txtDID.Text) & "'"
    Conn.Execute (strsql)
    strsql = "delete from QSMS_CompPNCheck  where DID = '" & txtDID.Text & "' "
    Conn.Execute (strsql)
    Lblstatus.Caption = "unlock is ok"
End If
Call RefreshData
End Sub



Private Sub Command1_Click()

End Sub

Private Sub DataGrid1_Click()
txtDID.Text = ""
txtCompPN.Text = ""
TxtReason.Text = ""
 On Error Resume Next
    With DataGrid1
        txtDID.Text = .Columns(1).Value
    End With
End Sub

Private Sub Form_Load()
    Call RefreshData
    dtpSDate.Value = Now
    dtpEDate.Value = Now
    Dim rs As ADODB.Recordset, str As String
    str = "select distinct Line from QSMS_woGroup"
    Set rs = Conn.Execute(str)
    CboLine.Clear
    While Not rs.EOF
        CboLine.AddItem rs!Line
        rs.MoveNext
    Wend
End Sub
Private Sub RefreshData()
Dim strsql As String
Dim rs As ADODB.Recordset
    txtDID.Text = ""
    txtCompPN.Text = ""
    TxtReason.Text = ""
    strsql = "select Line ,   DID ,CompPN , GroupID from  QSMS_CompPNCheck where   status='FAIL'  "
    Set rs = Conn.Execute(strsql)
    Set DataGrid1.DataSource = Nothing
    If rs.EOF = False Then
        Set DataGrid1.DataSource = rs
        DataGrid1.Refresh
    End If
End Sub
Private Sub InsertLog(sql As String)
Dim SQLlog As String
    
sql = Replace(sql, "'", "''")
ProgramName = UCase(VB.App.EXEName) & " ProgrameName " & Me.Caption & " Form "
SQLlog = "insert into QSMS_LOG(  [System_Name]  ,[Event_No]  ,[DID] ,[User_Name],[ReturnQty] ,[Trans_Date])" & _
        "values('QSMS_UnlockCompPNCheck','QSMS','" & sql & "','" & g_userName & "','0', " & " dbo.FormatDate (GETDATE (),'YYYYMMDDHHNNDD'))"
Conn.Execute (SQLlog)

End Sub
Private Sub txtCompPN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCompPN.Text <> "" Then
            position = InStr(Trim(txtCompPN.Text), ";") ' ´¦Àí2DµÄComopPN
            If (position > 1) Then
                txtCompPN.Text = Mid(Trim(txtCompPN.Text), 1, position - 1)
            End If
            TxtReason.SetFocus
    Else
            txtCompPN.SetFocus
    End If
End Sub
