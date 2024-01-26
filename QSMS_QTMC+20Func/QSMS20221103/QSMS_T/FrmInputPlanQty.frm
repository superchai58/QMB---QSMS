VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmInputPlanQty 
   BackColor       =   &H0000FF00&
   Caption         =   "Frm Plan Qty"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   12555
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraConnection 
      BackColor       =   &H80000013&
      Caption         =   "PlanQty maintain "
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   14415
      Begin VB.CommandButton CmmDelete 
         Caption         =   "Delete"
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
         Left            =   2760
         TabIndex        =   22
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Cmdfind 
         Caption         =   "Find"
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
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Add"
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
         Left            =   1080
         TabIndex        =   19
         Top             =   1920
         Width           =   735
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
         Left            =   5520
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxtTalQty 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox CboWoGroup 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxtPlanQty 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   7
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh"
         Enabled         =   0   'False
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
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "EXIT"
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
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "&Excel"
         Enabled         =   0   'False
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
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1920
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
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1920
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   13
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
         Format          =   70254595
         CurrentDate     =   39404
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   5520
         TabIndex        =   14
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
         Format          =   70254595
         CurrentDate     =   39404
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PlanQty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3960
         TabIndex        =   18
         Top             =   1200
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
         Left            =   3960
         TabIndex        =   17
         Top             =   240
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
         Left            =   3960
         TabIndex        =   16
         Top             =   720
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
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "TotalQty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "WoGroup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid DG1 
      Height          =   4095
      Left            =   0
      TabIndex        =   21
      Top             =   3480
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   7223
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
End
Attribute VB_Name = "FrmInputPlanQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CommandType As Long
Private Sub cboWO_Click()
    Dim strsql As String
    Dim Rs As ADODB.Recordset
    Dim BeginDate  As String, EndDate As String
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    
'    Str = "select qty from sap_wo_list where wo= '" & Trim(cboWO) & "' "
'    Set Rs = Conn.Execute(Str)
    strsql = "select * from QSMS_WOInputPlan where work_order= '" & cboWO & "' and begindatetime= '" & BeginDate & "' "
    Set Rs = Conn.Execute(strsql)
    If Rs.EOF = False Then
        TxtTalQty.Text = Trim(Rs!TotalQty)
    Else
         TxtTalQty.Text = TxtPlanQty.Text
    End If
     
End Sub

Private Sub CmdAdd_Click()
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    TxtPlanQty.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    CboWoGroup.Text = ""
    cboWO.Text = ""
    TxtTalQty.Text = ""
    TxtPlanQty.Text = ""
End Sub
Private Sub cmdExcel_Click()
    Dim str As String
    Dim Rs As ADODB.Recordset
    Dim BeginDate, EndDate As String
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    str = "Select * From QSMS_WOInputPlan Where transDateTime between  '" & BeginDate & "' and '" & EndDate & "'  Order by transDateTime desc"
    Set Rs = Conn.Execute(str)
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
    strsql = "Select * From QSMS_WOInputPlan Where transDateTime between  '" & BeginDate & "' and '" & EndDate & "'  Order by transDateTime desc"
    Set Rs = Conn.Execute(strsql)
    Set DG1.DataSource = Rs
    Call GetWoGroup
    Call GetWo
    CmdRefresh.Enabled = True
    cmdExcel.Enabled = True
    cmdCancel.Enabled = True
    
End Sub

Private Sub CmdRefresh_Click()
    Call RefreshDg("")
End Sub

Private Function GetWoGroup()
    Dim str As String
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
    'Str = "select distinct Group from sap_wo_list where Trans_Date between  '" & BeginDate & "' and '" & EndDate & "' "
    'Set Rs = Conn.Execute(Str)
    'i = 0
    'CboWoGroup.Clear
    'While Not Rs.EOF
    '      CboWoGroup.AddItem Trim(Rs!Group)
    '      Rs.MoveNext
    '      i = i + 1
    'Wend
    str = "select distinct GroupID from QSMS_WOGroup  "
    Set Rs = Conn.Execute(str)
    i = 0
    CboWoGroup.Clear
    While Not Rs.EOF
          CboWoGroup.AddItem Trim(Rs!GroupID)
          Rs.MoveNext
          i = i + 1
    Wend
    If i = 0 Then
       MsgBox "No data"
    End If
End Function
Private Function GetWo()
    Dim str As String
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
    'Str = "select distinct wo from sap_wo_list where group like '" & CboWoGroup & "%'"
    str = "select distinct wo from sap_wo_list where trans_Date between  '" & BeginDate & "' and '" & EndDate & "' order by wo "
    Set Rs = Conn.Execute(str)
    i = 0
    cboWO.Clear
    While Not Rs.EOF
          cboWO.AddItem Trim(Rs!WO)
          Rs.MoveNext
          i = i + 1
    Wend
End Function

Private Sub cmdSave_Click()
   Dim strsql, str As String
    Dim Rs As ADODB.Recordset
    Dim TempDID As String
    Dim TransDate As String
    Dim i As Long
    Dim GroupIDHead As String
    Cmdfind.Enabled = True
    cmdCancel.Enabled = True
    cmdExit.Enabled = True
    Dim BeginDate, EndDate As String
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    If Trim(TxtPlanQty) = "" Then
        MsgBox (" PlanQty can't be empty!!"), vbCritical
        TxtPlanQty.SetFocus
        Exit Sub
    End If
    str = "select qty,combineqty from sap_wo_list where wo= '" & Trim(cboWO) & "' "
    Set Rs = Conn.Execute(str)
    If Not Rs.EOF Then
        If CLng(Trim(TxtPlanQty)) Mod Trim(Rs!CombineQty) <> 0 Then
            MsgBox (" this Work Order' CombineQty is " & Rs!CombineQty & " ,please add PlanQty is " & Rs!CombineQty & " Times")
            Exit Sub
        End If
        If Trim(Rs!Qty) < CLng(Trim(TxtTalQty)) Then
            MsgBox (" TatolQty can't big WO TatolQty in SAP_wo_LIST Tabel!!"), vbCritical
            TxtPlanQty.SetFocus
            Exit Sub
        End If
    End If
    strsql = "select getdate()"
    Set Rs = Conn.Execute(strsql)
    TransDate = Format(Rs(0), "YYYYMMDD")
    strsql = "select * from QSMS_WOInputPlan where work_order= '" & cboWO & "' and begindatetime= '" & BeginDate & "' "
    Set Rs = Conn.Execute(strsql)
    If Rs.EOF = True Then
        strsql = "insert into QSMS_WOInputPlan(work_order,wogroup,inputqty,totalqty,begindatetime,enddatetime,transdatetime,uid) " & _
        " values ('" & cboWO & "' , '" & CboWoGroup & "' , '" & Trim(TxtPlanQty) & "' , '" & TxtTalQty & "' , '" & BeginDate & "','" & EndDate & "' , '" & TransDate & "' , '" & g_userName & "')"
        Conn.Execute strsql
    Else
        strsql = "update QSMS_WOInputPlan set inputqty=" & Trim(TxtPlanQty) & " , work_order='" & Trim(cboWO) & "', wogroup='" & CboWoGroup & "', totalqty=" & TxtTalQty & " , begindatetime='" & BeginDate & "', enddatetime='" & EndDate & "' , transdatetime='" & TransDate & "',  uid='" & g_userName & "'  where work_order= '" & cboWO & "' and begindatetime= '" & BeginDate & "'"
        Conn.Execute strsql
    End If
    Call UpdatePlanDispatch(cboWO)
    Call cmdFind_Click
    Call RefreshDg("")
    Call cmdCancel_Click
End Sub

Private Sub cmdUpdate_Click()
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdCancel.Enabled = True
    TxtPlanQty.Enabled = True
End Sub

Private Sub CmmDelete_Click()
    Dim str, strsql As String
    Dim Rs As ADODB.Recordset
    Dim BeginDate, EndDate As String
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    strsql = "delete from QSMS_WOInputPlan where work_order= '" & cboWO & "' and begindatetime= '" & BeginDate & "' "
    Set Rs = Conn.Execute(strsql)
    strsql = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Delete Plan Qty','" & cboWO & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
    Conn.Execute (strsql)
    Call RefreshDg("")
    Call cmdCancel_Click
End Sub

Private Sub DG1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim stemp As String
    On Error Resume Next
    With DG1
        cboWO = .Columns(0).Value
        CboWoGroup = .Columns(1).Value
'        TxtPlanQty = .Columns(2).Value
        TxtTalQty = .Columns(3).Value
        stemp = .Columns(4).Value
        dtpSDate.Value = Left(stemp, 4) + "/" + Mid(stemp, 5, 2) + "/" + Mid(stemp, 7, 2)
        stemp = .Columns(5).Value
        dtpEDate.Value = Left(stemp, 4) + "/" + Mid(stemp, 5, 2) + "/" + Mid(stemp, 7, 2)
    End With
    cmdUpdate.Enabled = True
    'cmdDelete.Enabled = True
    cmdCancel.Enabled = True
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub
'Private Sub dtpEDate_Click()
'    Call GetWoGroup
'    Call GetWo
'End Sub

Private Sub Form_Load()
    Dim Rs As ADODB.Recordset
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call RefreshDg("")
    dtpSDate = Date
    dtpEDate = Date
End Sub

Private Function RefreshDg(ByVal CompPN As String)
    Dim str As String
    Dim Rs As ADODB.Recordset
    str = "Select * From QSMS_WOInputPlan "
    Set Rs = Conn.Execute(str)
    Set DG1.DataSource = Rs
    DG1.Refresh
End Function
Private Sub TxtPlanQty_Change()
    Dim strsql As String
    Dim Rs As ADODB.Recordset
    Dim BeginDate, EndDate As String
    BeginDate = Format(dtpSDate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(dtpEDate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
'    Str = "select qty from sap_wo_list where wo= '" & Trim(cboWO) & "' "
'    Set Rs = Conn.Execute(Str)
    strsql = "select * from QSMS_WOInputPlan where work_order= '" & cboWO & "' and begindatetime= '" & BeginDate & "' "
    Set Rs = Conn.Execute(strsql)
    If Not IsNumeric(TxtPlanQty.Text) And TxtPlanQty <> "" Then
        MsgBox "Please check Plan Qty is Numeric"
        TxtTalQty.Text = ""
        TxtPlanQty.Text = ""
        Exit Sub
    End If
    If TxtPlanQty = "" Then Exit Sub
    If Rs.EOF = False Then
        TxtTalQty = CLng(TxtPlanQty) + CLng(Trim(Rs!TotalQty))
    Else
         TxtTalQty = CLng(TxtPlanQty)
    End If
End Sub


Private Function UpdatePlanDispatch(ByVal WO As String)
    Dim str As String
    Dim Rs As ADODB.Recordset
    If WO <> "" Then
        str = "exec QSMS_UpdatePlanDispatchQty '" & WO & "'"
        Conn.Execute (str)
    Else
        MsgBox "WO is empty !"
    End If
    
End Function
