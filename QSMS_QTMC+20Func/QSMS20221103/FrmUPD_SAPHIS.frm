VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmUpdSapHis 
   BackColor       =   &H0000FF00&
   Caption         =   "Frm Update QSMS_Sap_His[2007-08-09]"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraConnection 
      BackColor       =   &H80000013&
      Caption         =   "Update QSMS_SapHis"
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      Begin VB.ComboBox CboBegin 
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
         Left            =   5640
         TabIndex        =   14
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton Cmdfind 
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
         Height          =   495
         Left            =   8040
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cboWO 
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
         Left            =   1680
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmdExcel 
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
         Height          =   855
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1320
         Width           =   735
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
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
         Format          =   120848387
         CurrentDate     =   39404
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   5520
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
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
         Format          =   120848387
         CurrentDate     =   39404
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "BeginDateTime"
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
         Left            =   3960
         TabIndex        =   16
         Top             =   840
         Width           =   1695
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
         Left            =   3960
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Work Order"
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
         Top             =   840
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
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid DG1 
      Height          =   2535
      Left            =   -120
      TabIndex        =   13
      Top             =   2520
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
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
   Begin MSDataGridLib.DataGrid DG2 
      Height          =   3015
      Left            =   -120
      TabIndex        =   17
      Top             =   5160
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "Work Order"
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
      Left            =   7440
      TabIndex        =   15
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "FrmUpdSapHis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BeginDate As String, EndDate As String
Dim CommandType As Long

Private Sub CboBegin_Click()
    Dim Strsql As String
    Dim Rs As ADODB.Recordset
    Strsql = "select  * from qsms_sapHis where work_order= '" & cboWO & "' and beginDateTime ='" & CboBegin & " ' Order by beginDateTime desc"
    Set Rs = Conn.Execute(Strsql)
    If Not Rs.EOF Then
        Set DG2.DataSource = Rs
    Else
        MsgBox ("QSMS_SapHis NOT DATA")
    End If
End Sub

Private Sub cboWO_Click()
    Dim Strsql As String
    Dim Rs As ADODB.Recordset
    Dim BeginDate As String, EndDate As String
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    Strsql = "select  * from qsms_sapHis where work_order= '" & cboWO & "' and beginDateTime>='" & BeginDate & " ' Order by beginDateTime desc"
    Set Rs = Conn.Execute(Strsql)
    If Not Rs.EOF Then
        Set DG2.DataSource = Rs
    Else
        MsgBox ("QSMS_SapHis NOT DATA")
    End If
    Strsql = "select  * from qsms_chksapfile where work_order= '" & cboWO & "' and beginDateTime>='" & BeginDate & " ' and (logflag='N' or fileflag='N') Order by beginDateTime desc"
    Set Rs = Conn.Execute(Strsql)
    If Not Rs.EOF Then
        Set DG1.DataSource = Rs
    Else
        MsgBox ("QSMS_Chksapfile NOT DATA")
    End If
    Call GetBeginDateTime
End Sub


Private Sub cmdCancel_Click()
    cboWO.Text = ""
End Sub
Private Sub cmdExcel_Click()
    Dim Str As String
    Dim Rs As ADODB.Recordset
    Dim BeginDate, EndDate As String
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    Str = "select  * from qsms_sapHis where work_order= '" & cboWO & "' and beginDateTime>='" & BeginDate & " ' Order by beginDateTime desc"
    Set Rs = Conn.Execute(Str)
     If Not Rs.EOF Then
           Call CopyToExcel(Rs)
        Else
           MsgBox ("No Data"), vbCritical
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim Rs As ADODB.Recordset
    Dim BeginDate, EndDate As String
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    Strsql = "select TOP 10  * from qsms_chksapfile Where (logflag='N' or fileflag='N') and beginDateTime between  '" & BeginDate & "' and '" & EndDate & "'   "
    Set Rs = Conn.Execute(Strsql)
    Set DG1.DataSource = Rs
    Strsql = "select top 10 * from qsms_sapHis  Where beginDateTime between  '" & BeginDate & "' and '" & EndDate & "'   "
    Set Rs = Conn.Execute(Strsql)
    Set DG2.DataSource = Rs
    Call GetWo
    
End Sub

Private Sub CmdRefresh_Click()
    Call cboWO_Click
End Sub


Private Function GetWo()
    Dim Str As String
    
    Dim GroupIDHead As String
    Dim i As Long
    Dim Rs As ADODB.Recordset
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    Str = "select distinct work_order from qsms_sapHis where work_order in (select work_order from qsms_chksapfile Where (logflag='N' or fileflag='N') and begindatetime between  '" & BeginDate & "' and '" & EndDate & "')  Order by work_order "
    'str = "select work_order from qsms_chksapfile Where (logflag='N' or fileflag='N') and begindatetime between  '" & BeginDate & "' and '" & EndDate & "'  Order by work_order "
    Set Rs = Conn.Execute(Str)
    i = 0
    cboWO.Clear
    While Not Rs.EOF
          cboWO.AddItem Trim(Rs!Work_Order)
          Rs.MoveNext
          i = i + 1
    Wend
End Function


Private Function GetBeginDateTime()
    Dim Str As String
    Dim BeginDate, EndDate As String
    Dim GroupIDHead As String
    Dim i As Long
    Dim Rs As ADODB.Recordset
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    Str = "select distinct begindatetime from qsms_sapHis where  work_order= '" & cboWO & "'  Order by begindatetime "
    'str = "select work_order from qsms_chksapfile Where (logflag='N' or fileflag='N') and begindatetime between  '" & BeginDate & "' and '" & EndDate & "'  Order by work_order "
    Set Rs = Conn.Execute(Str)
    i = 0
    CboBegin.Clear
    While Not Rs.EOF
          CboBegin.AddItem Trim(Rs!begindatetime)
          Rs.MoveNext
          i = i + 1
    Wend
End Function

Private Sub cmdUpdate_Click()
    Dim Strsql As String
    Dim Rs As ADODB.Recordset
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    Strsql = "Update qsms_sapHis set sendflag='N' where work_order= '" & cboWO & "' and beginDateTime='" & CboBegin & "' "
    Set Rs = Conn.Execute(Strsql)
'    If Rs.EOF = True Then
'        MsgBox ("update fail,please check BeginDateTime & WorkOrder.")
'    End If
    Strsql = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','update QSMS_SapHis','" & cboWO & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
    Conn.Execute (Strsql)
    Call cboWO_Click
End Sub
'
'
'
'Private Sub DG2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    Dim stemp As String
'    With DG2
'        cboWO = .Columns(0).Value
'        CboBegin = .Columns(8).Value
''        stemp = .Columns(8).Value
''        dtpSDate.Value = Left(stemp, 4) + "/" + Mid(stemp, 5, 2) + "/" + Mid(stemp, 7, 2)
''        stemp = .Columns(9).Value
''        dtpEDate.Value = Left(stemp, 4) + "/" + Mid(stemp, 5, 2) + "/" + Mid(stemp, 7, 2)
'    End With
''    cmdUpdate.Enabled = True
''    CmdDelete.Enabled = True
''    cmdCancel.Enabled = True
'End Sub

Private Sub Form_Load()
    Dim Rs As ADODB.Recordset
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call RefreshDg("")
    dtpSDate = Date
    dtpEDate = Date
End Sub


Private Function RefreshDg(ByVal Str1 As String)
    Dim Str As String
    Dim Rs As ADODB.Recordset
    Str = "select top 10 * from qsms_sapHis "
    Set Rs = Conn.Execute(Str)
    Set DG2.DataSource = Rs
    DG2.Refresh
    Str = "select top 10 * from qsms_chksapfile "
    Set Rs = Conn.Execute(Str)
    Set DG1.DataSource = Rs
    DG1.Refresh
End Function

