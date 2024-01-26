VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmSAP1Patch 
   Caption         =   "FrmSAP1Patch"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CboType 
      Height          =   315
      Left            =   12240
      TabIndex        =   35
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton cmdExcelSAP 
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
      Left            =   15000
      Picture         =   "FrmSAP1Patch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4320
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DGSAP1 
      Height          =   3855
      Left            =   120
      TabIndex        =   30
      Top             =   3240
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   6800
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
      Caption         =   "SAP1 Data"
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
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   13200
         Picture         =   "FrmSAP1Patch.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Get SAP1 Patch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6600
         Picture         =   "FrmSAP1Patch.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox TxtPackingQty 
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
         Left            =   10680
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox TxtGroup 
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
         Left            =   10680
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Frame FraSB 
         Caption         =   "Small Board WO"
         Height          =   615
         Left            =   6600
         TabIndex        =   11
         Top             =   1200
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
            Style           =   2  'Dropdown List
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
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
         Left            =   14040
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TxtCustomer 
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
         Left            =   14040
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
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
         Left            =   10680
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
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
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   2655
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
         Picture         =   "FrmSAP1Patch.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptRelease 
         Caption         =   "Release"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "Group"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
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
         Left            =   14040
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1200
         Width           =   975
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
         Left            =   10680
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   2295
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
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
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
         Format          =   69861379
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1080
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
         Format          =   69861379
         CurrentDate     =   36482
      End
      Begin VB.Label LblMessage 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   2040
         Width           =   5895
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000FF&
         Caption         =   "PackingQty"
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
         Left            =   9240
         TabIndex        =   29
         Top             =   1680
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
         TabIndex        =   26
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Group(M/S)"
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
         Index           =   22
         Left            =   9240
         TabIndex        =   25
         Top             =   1200
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
         Left            =   12960
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Customer"
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
         Index           =   16
         Left            =   12960
         TabIndex        =   23
         Top             =   240
         Width           =   1095
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
         Left            =   9240
         TabIndex        =   22
         Top             =   240
         Width           =   1455
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
         TabIndex        =   21
         Top             =   120
         Width           =   2175
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
         Left            =   12960
         TabIndex        =   20
         Top             =   1200
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
         TabIndex        =   19
         Top             =   1560
         Width           =   1455
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
         Left            =   9240
         TabIndex        =   18
         Top             =   720
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
         TabIndex        =   17
         Top             =   720
         Width           =   2295
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
         TabIndex        =   16
         Top             =   600
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid DGSAP1More 
      Height          =   3855
      Left            =   120
      TabIndex        =   31
      Top             =   7320
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   6800
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
      Caption         =   "SAP1 more Data"
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
Attribute VB_Name = "FrmSAP1Patch"
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
Call GetSBWO(TxtWO)
Call GetWoinfo(TxtWO)
TxtPackingQty = GetPackingQty(TxtWO)
Call GetSAP1Data(Trim(TxtGroup))
End Sub

Private Sub cboWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CboWo_Click
End If
End Sub

Private Sub cmdExcelSAP_Click()
Dim Str As String
Dim Rs As ADODB.Recordset

If Trim(TxtGroup) = "" Or Trim(TxtWO) = "" Then
    MsgBox "Group or WO can not be empty,please check"
    Exit Sub
End If
CboType.AddItem "SAP1Lost"
CboType.AddItem "SAP1More"
CboType.AddItem "SAP1ALL"
Select Case Trim(CboType)
       Case "SAP1Lost"
                    Str = "select * from  qsms_sap_Balance where work_order in (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "') and lostqty>0 and ctype='packingqty' order by work_order,Upcomppn,item"
                 
       Case "SAP1More"
                      Str = "select * from  qsms_sap_Balance where work_order in (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "') and lostqty<0 and ctype='packingqty' order by work_order,upcomppn,item"
                 
       Case "SAP1ALL"
                     Str = "select * from  qsms_sap_Balance where work_order in (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "')  and ctype='packingqty' order by work_order,upcomppn,item"
End Select
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
    Call CopyToExcel(Rs)
Else
    MsgBox "No data found"
End If

End Sub



Private Sub cmdFind_Click()
Dim Str As String
Dim RsLost As ADODB.Recordset
Dim RsMore As ADODB.Recordset
If Trim(TxtGroup) = "" Or Trim(TxtWO) = "" Then
    MsgBox "Group or WO can not be empty,please check"
    Exit Sub
End If
LblMessage.BackColor = &HFF&
LblMessage.Caption = "Get SAP1 Patch,please wait"
DGSAP1.Caption = "SAP1 Lost"
Str = "exec QSMSGetSap1BalanceRpt '" & Trim(TxtWO) & "','PackingQty','',''," & CLng(Trim(TxtPackingQty)) & ""
Conn.Execute Str
Str = "select * from  qsms_sap_Balance where work_order in (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "') and lostqty>0 and ctype='packingqty'"
Set RsLost = Conn.Execute(Str)
Set DGSAP1.DataSource = RsLost

Str = "select * from  qsms_sap_Balance where work_order in (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "') and lostqty<0 and ctype='packingqty'"
Set RsLost = Conn.Execute(Str)
Set DGSAP1More.DataSource = RsLost
LblMessage.BackColor = &HFF00&
LblMessage.Caption = "Get SAP1 Patch finished"
End Sub

Private Sub CmdQuery_Click()
If Trim(CboLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If
Call GetGroupID
End Sub

Private Sub cmdSave_Click()
Dim Str As String
Dim Rs As ADODB.Recordset
Dim TransDateTime As String
Str = "select getdate()"
Set Rs = Conn.Execute(Str)
TransDateTime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
'Str = "delete from qsms_saphis where  work_order in   (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "') and EndDateTime='packingqty'"
'Conn.Execute Str
'insert into qsms_saphis
Str = "insert into qsms_saphis(work_order,item,comppn,upcomppn,Qty,status,sendflag,Transdatetime,BeginDateTime,EndDateTime,AOIQty) " & _
     " select work_order,item,comppn,upcomppn,LostQty,'open','N','" & TransDateTime & "' ,'" & TransDateTime & "','packingqty',0 " & _
     " from qsms_sap_Balance where work_order in  (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "') and Ctype='packingqty' and LostQty>0 "
Conn.Execute Str
'insert into qsms_sap
Str = "delete from qsms_sap  where  work_order in  (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "') and Status='open'"
Conn.Execute Str
Str = " insert into  qsms_sap (work_order,item,comppn,upcomppn,Qty,Status,TransDateTime) " & _
    " select work_order,item,comppn,upcomppn,sum(qty),'open', '" & TransDateTime & "' from qsms_saphis where " & _
    " work_order in  (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "')  group by work_order,item,comppn,upcomppn "
Conn.Execute Str
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
dtpSDate = Date
dtpEDate = Date
Call GetLine
CboType.AddItem "SAP1Lost"
CboType.AddItem "SAP1More"
CboType.AddItem "SAP1ALL"
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
Private Function GetPackingQty(ByVal WO As String) As Long
Dim Str As String
Dim Rs As ADODB.Recordset
GetPackingQty = 0
Str = "Select count(*) from SMT_Packing where workorder='" & WO & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   GetPackingQty = Rs.Fields(0)
End If

End Function
Private Function GetSAP1Data(ByVal PCBGroup As String)
Dim Str As String
Dim Rs As ADODB.Recordset
DGSAP1.Caption = "SAP1 send"
Str = "select work_order,upcomppn,item,comppn,qty from qsms_sap where work_order in (select wo from sap_wo_list where [group]='" & PCBGroup & "') and status='open' order by work_order,UpCompPN,item"
Set Rs = Conn.Execute(Str)
Set DGSAP1.DataSource = Rs
End Function

Private Function GetGroupID()
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
'GroupIDHead = Trim(CboLine) & TransDate
If OptRelease.Value = True Then
   Str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "'"
Else
    Str = "select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "'"
End If

Set Rs = Conn.Execute(Str)
i = 0
CboGroupID.Clear
While Not Rs.EOF
      CboGroupID.AddItem Trim(Rs!GroupID)
      Rs.MoveNext
      i = i + 1
Wend
If i = 0 Then
   MsgBox "No data"
   
End If
End Function

Private Function GetGroupWO(ByVal GroupID As String)
Dim Str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset

Str = "select Work_Order from QSMS_WOGroup  where GroupID= '" & GroupID & "' order by Seq_NO"

Set Rs = Conn.Execute(Str)
cboWO.Clear

While Not Rs.EOF
      If ChkMBWo(Rs!Work_Order) = True Then
              cboWO.AddItem Trim(Rs!Work_Order)
'            If ChkQSMS_WO(Trim(Rs!Work_Order)) = False Then
'                CboNotChkBOM.AddItem Trim(Rs!Work_Order)
'            Else
'
'
'                If ChkWoFinished(Rs!Work_Order) = True Then
'
'                    cboWO.AddItem Trim(Rs!Work_Order)
'                Else
'
'                     CboNotFinishedWO.AddItem Trim(Rs!Work_Order)
'                End If
'            End If
      End If
      Rs.MoveNext
Wend
End Function

Private Function GetSBWO(ByVal WO As String)
Dim Str As String
Dim Rs As ADODB.Recordset
Dim i As Long
Dim Group As String
i = 0
CboSBWO.Clear
FraSB.Visible = False
Str = "Select [Group] from Sap_Wo_List where wo='" & WO & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   Group = Trim(Rs!Group)
   TxtGroup = Group
End If
Str = "select Wo from Sap_Wo_list where [Group] ='" & Group & "' and wo<>'" & WO & "' order by wo"
Set Rs = Conn.Execute(Str)
While Not Rs.EOF
     CboSBWO.AddItem Trim(Rs!WO)
     Rs.MoveNext
     i = i + 1
Wend
If i > 0 Then
    FraSB.Visible = True

End If
End Function

Private Function GetWoinfo(ByVal WO As String)
Dim Str As String
Dim Rs As ADODB.Recordset
Str = "select PN, Qty,[Group] from Sap_Wo_List where WO='" & Trim(WO) & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   TxtMBPN = Rs!PN
   TxtModel = Mid(TxtMBPN, 3, 3)
   TxtWOQty = Rs!Qty
   TxtGroup = Trim(Rs![Group])
End If
Str = "select Customer from ModelName where PN='" & TxtMBPN & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   TxtCustomer = Trim(Rs!Customer)
End If
End Function
