VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmQueryCheckBOM 
   Caption         =   "Query wo check bom information[2009-11-18]"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   10800
   StartUpPosition =   3  '窗口缺省
   Begin MSDataGridLib.DataGrid Dg1 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5953
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "All work order are check bom fail"
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
   Begin VB.Frame Frame1 
      Caption         =   "maintain"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2880
         TabIndex        =   11
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton CmdQuery 
         Caption         =   "Query"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox CboLine 
         Height          =   300
         Left            =   5400
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TxtWO 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   360
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
         Format          =   25755651
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   5400
         TabIndex        =   5
         Top             =   360
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
         Format          =   25755651
         CurrentDate     =   36482
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   9
         Top             =   960
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
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
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
         Left            =   3840
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Work_Order"
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
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmQueryCheckBOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdQuery_Click()
Dim str As String
Dim BeginDate As String
Dim EndDate As String
Dim rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
BeginDate = Format(dtpSDate, "YYYY/MM/DD")
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
EndDate = Format(dtpEDate, "YYYY/MM/DD")
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")
str = "EXEC QSMS_QueryCheckBOM '" & Trim(TxtWO) & "','" & Trim(CboLine) & "','" & Trim(BeginDate) & "','" & Trim(EndDate) & "'"
Set rs = Conn.Execute(str)
Set Dg1.DataSource = rs
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim str As String
Dim rs As ADODB.Recordset
Dim i As Long
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
str = "select getdate()"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
    'Date = Rs(0)
    'Time = Rs(0)
End If
dtpSDate = Date
dtpEDate = Date + 1
CboLine.Clear
str = "select distinct line from QSMS_WoGroup  order by line"
Set rs = Conn.Execute(str)
While Not rs.EOF
    CboLine.AddItem Trim(rs!Line)
    rs.MoveNext
Wend
str = "EXEC QSMS_QueryCheckBOM "
Set rs = Conn.Execute(str)
Set Dg1.DataSource = rs
End Sub
