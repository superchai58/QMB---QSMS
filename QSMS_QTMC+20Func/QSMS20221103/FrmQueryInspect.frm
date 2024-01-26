VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmQueryInspect 
   Caption         =   "Query Inspect[2016.08.11]"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13920
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   13920
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5295
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   9340
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   12975
      Begin MSComCtl2.DTPicker dtETime 
         Height          =   375
         Left            =   8760
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Format          =   88932354
         CurrentDate     =   .999305555555556
      End
      Begin MSComCtl2.DTPicker dtEDate 
         Height          =   375
         Left            =   7200
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   88932353
         CurrentDate     =   40289
      End
      Begin MSComCtl2.DTPicker dtSDate 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   88932353
         CurrentDate     =   40289
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   10080
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "Query"
         Height          =   375
         Left            =   10080
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtDID 
         Height          =   360
         Left            =   2760
         TabIndex        =   2
         Top             =   480
         Width           =   6975
      End
      Begin MSComCtl2.DTPicker dtSTime 
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   88932354
         CurrentDate     =   40289
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "EndTime:"
         Height          =   375
         Left            =   5880
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "StartTime:"
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "DID\CompPN:"
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmQueryInspect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim Num As Integer
Private Sub cmdExcel_Click()
    If Num = 0 Then
        Exit Sub
    End If
    rs.MoveFirst
    If rs.EOF = False Then
        Call CopyToExcel(rs)
    Else
        MsgBox ("No Data"), vbCritical
    End If
End Sub

Private Sub cmdQuery_Click()

    Dim str As String
    Dim sDate As String
    Dim eDate As String
    
    dtETime = Now
    sDate = Format(dtSDate, "YYYYMMDD") & Format(dtSTime, "HHNNSS")
    eDate = Format(dtEDate, "YYYYMMDD") & Format(dtETime, "HHNNSS")
    
    Set DataGrid1.DataSource = Nothing
    
    str = "select * from QSMS_DID_InSpect where (DID like '%" & txtDID & "%' or CompPN like '%" & txtDID & "%') and transDatetime between '" & sDate & "' and '" & eDate & "' order by TransDateTime desc"
    
    Set rs = Conn.Execute(str)
    Num = rs.RecordCount
    If Not rs.EOF Then
        Set DataGrid1.DataSource = rs
    Else
       MsgBox ("No Data"), vbCritical
    End If
End Sub

Private Sub Form_Load()
    Me.dtSDate = Now - 1 ''(1234)
    Me.dtEDate = Now ''(1234)
End Sub
