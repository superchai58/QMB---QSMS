VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCompDiff 
   Caption         =   "FrmTransferDispatchedDID"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13965
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   13965
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Dispatch Qty - QSMS_WO No Match Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   240
      TabIndex        =   22
      Top             =   3120
      Width           =   13215
      Begin MSDataGridLib.DataGrid DataGrid 
         Height          =   4335
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   7646
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
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin VB.TextBox txtRev 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11640
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdCheck 
         BackColor       =   &H0000FF00&
         Caption         =   "Check"
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
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1440
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPedate 
         Height          =   375
         Left            =   1680
         TabIndex        =   24
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   69795841
         CurrentDate     =   39091
      End
      Begin MSComCtl2.DTPicker DTPsdate 
         Height          =   375
         Left            =   1680
         TabIndex        =   23
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   69795841
         CurrentDate     =   39091
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
         Left            =   6600
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox TxtMBPN 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox TxtWOQty 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13200
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2280
         Width           =   735
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "Group"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton OptRelease 
         Caption         =   "Release"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
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
         Height          =   975
         Left            =   3360
         Picture         =   "FrmCompDiff.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox CboGroupID 
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
         Left            =   6600
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox TxtWO 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox TxtModel 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Frame FraSB 
         Caption         =   "Small Board WO"
         Height          =   615
         Left            =   6600
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   2655
         Begin VB.ComboBox CboSBWO 
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
            Left            =   240
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rev"
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
         Index           =   3
         Left            =   10920
         TabIndex        =   28
         Top             =   2280
         Width           =   735
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
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "OK Work Order"
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
         Left            =   4440
         TabIndex        =   20
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MB PN"
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
         Left            =   7560
         TabIndex        =   19
         Top             =   2280
         Width           =   1095
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
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Qty"
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
         Left            =   12480
         TabIndex        =   17
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "GroupID"
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
         Left            =   4440
         TabIndex        =   16
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "WO"
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
         Index           =   13
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Model"
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
         Index           =   21
         Left            =   3960
         TabIndex        =   14
         Top             =   2280
         Width           =   735
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
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmCompDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub CboGroupID_Click()
Call GetGroupWO(CboGroupID)
End Sub






Private Sub CboWo_Click()
TxtWO = Trim(cboWO)
Call GetWoinfo(TxtWO)
'Call GetMachine(TxtWO)
Set DataGrid.DataSource = Nothing
End Sub



Private Sub cmdCheck_Click()
Dim Str As String
Dim Rs As New ADODB.Recordset

Str = "exec QsmsComp_Diff '" & Trim(cboWO) & "'"
Set Rs = Conn.Execute(Str)

If Rs.EOF = True Then
    MsgBox "Work_order " & Trim(cboWO) & " check OK ! ", vbInformation
Else
    Set DataGrid.DataSource = Nothing
    Set DataGrid.DataSource = Rs
    DataGrid.Refresh
End If

End Sub

Private Sub CmdQuery_Click()
If Trim(CboLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If
Call GetGroupID("")
End Sub



Private Sub Form_Load()
Dim Str As String
Dim Rs As ADODB.Recordset

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
Str = "select getdate()"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
    'Date = Rs(0)
    'Time = Rs(0)
End If
DTPsdate = Date
DTPedate = Date
Call GetLine
End Sub

Private Function GetLine()
Dim Str As String
Dim Rs As ADODB.Recordset
Str = "select distinct Line from QSMS_woGroup"
Set Rs = Conn.Execute(Str)
CboLine.Clear
While Not Rs.EOF
    CboLine.AddItem Rs!Line
    Rs.MoveNext
Wend
End Function

Private Function GetGroupID(ByVal Jobpn As String)
Dim Str As String
Dim BeginDate, EndDate As String
Dim GroupIDHead As String
Dim i As Long
Dim Rs As ADODB.Recordset
Dim TempJobPn As String
BeginDate = Format(DTPsdate, "YYYY/MM/DD")
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
EndDate = Format(DTPedate, "YYYY/MM/DD")
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")

Str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "'"
Set Rs = Conn.Execute(Str)

CboGroupID.Clear
If Rs.EOF Then MsgBox "No data"
While Not Rs.EOF
      CboGroupID.AddItem Trim(Rs!GroupID)
      Rs.MoveNext
Wend
End Function

Private Function GetWoinfo(ByVal WO As String)
Dim Str As String
Dim Rs As ADODB.Recordset
Str = "select PN, Qty ,MB_Rev,Line from Sap_Wo_List where WO='" & Trim(WO) & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   TxtMBPN = Rs!PN
   TxtWOQty = Rs!Qty
   txtRev = Trim(Rs!Mb_Rev)
   TxtModel = Rs!PN + "-" + Trim(Rs!Mb_Rev)
End If
'Str = "select Customer from ModelName where PN='" & TxtMBPN & "'"
'Set Rs = Conn.Execute(Str)
'If Not Rs.EOF Then
'   TxtCustomer = Trim(Rs!Customer)
'End If

End Function



Private Function GetGroupWO(ByVal GroupID As String)
Dim Str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset
cboWO.Clear
Str = "select distinct work_order from qsms_wogroup where groupid='" & CboGroupID & "'"
Set Rs = Conn.Execute(Str)
While Not Rs.EOF
          
        cboWO.AddItem Trim(Rs!Work_Order)
        Rs.MoveNext
Wend
End Function



