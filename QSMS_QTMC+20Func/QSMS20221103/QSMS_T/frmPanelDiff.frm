VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPanelDiff 
   Caption         =   "Panel Different"
   ClientHeight    =   10695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   ScaleHeight     =   10695
   ScaleWidth      =   12480
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dgInfo2 
      Height          =   4095
      Left            =   120
      TabIndex        =   9
      Top             =   6480
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   7223
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin MSDataGridLib.DataGrid dgInfo1 
      Height          =   4095
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   7223
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
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
            LCID            =   1033
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
            LCID            =   1033
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
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12255
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
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
         Left            =   10800
         TabIndex        =   8
         Top             =   405
         Width           =   1215
      End
      Begin VB.ComboBox cboJobG2 
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
         Left            =   7800
         TabIndex        =   6
         Top             =   480
         Width           =   2775
      End
      Begin VB.ComboBox cboJobG1 
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
         Left            =   3360
         TabIndex        =   5
         Top             =   480
         Width           =   2775
      End
      Begin VB.ComboBox cboLine 
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
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "JobGroup2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   4
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "JobGroup1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2040
         TabIndex        =   3
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Line:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   510
      End
   End
   Begin VB.Label lblDiffQty2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   13
      Top             =   6120
      Width           =   60
   End
   Begin VB.Label lblDiffQty1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   12
      Top             =   1440
      Width           =   60
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Differences with JobGroup1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Differences with JobGroup2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   2895
   End
End
Attribute VB_Name = "frmPanelDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************************************************************
'**文 件 名: FrmPanelDiff.frm
'**Copyright (C) 2008-2010 QMS
'**文件编号:
'**创 建 人: Jing.Chen
'**日    期: 2008.09.26
'**描    述: Differences in two Panels
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'***********************************************************************************/

Option Explicit

Private Sub cboLine_LostFocus()
Dim strSQL As String
Dim tmpRS As New ADODB.Recordset

strSQL = "select distinct pn+'-'+mb_rev from sap_wo_list with(nolock) where Status='20' and CheckBomPassDateTime<>'' and charindex('MB',pn)>0 and Line='" & Trim(cboLine.Text) & "' "
Set tmpRS = Conn.Execute(strSQL)

While Not tmpRS.EOF
    cboJobG1.AddItem tmpRS(0)
    cboJobG2.AddItem tmpRS(0)
    tmpRS.MoveNext
Wend

End Sub

Private Sub cmdOK_Click()
Dim strSQL As String
Dim tmpRS As New ADODB.Recordset

On Error GoTo Handler:

If Trim(cboLine.Text) = "" Or Trim(cboJobG1.Text) = "" Or Trim(cboJobG2.Text) = "" Then
    MsgBox ("Please input line/JobGroup1/JobGroup2 !"), vbCritical
    cboLine.SetFocus
    Exit Sub
End If

If Trim(cboJobG1.Text) = Trim(cboJobG2.Text) Then
    MsgBox ("Please input two different JobGroups !"), vbCritical
    cboJobG2.SetFocus
    Exit Sub
End If

strSQL = "Exec QSMS_PanelDiff '" & Trim(cboLine.Text) & "','" & Trim(cboJobG1.Text) & "','" & Trim(cboJobG2.Text) & "'"
Set tmpRS = Conn.Execute(strSQL)

If tmpRS.EOF = False Then
    Set dgInfo1.DataSource = tmpRS.DataSource
    lblDiffQty1.Caption = tmpRS.RecordCount
Else
    Set dgInfo1.DataSource = Null
    lblDiffQty1.Caption = "0"
End If

Set tmpRS = tmpRS.NextRecordset
If tmpRS.EOF = False Then
    Set dgInfo2.DataSource = tmpRS.DataSource
    lblDiffQty2.Caption = tmpRS.RecordCount
Else
    Set dgInfo2.DataSource = Null
    lblDiffQty2.Caption = "0"
End If

Exit Sub
Handler:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
Dim strSQL As String
Dim tmpRS As New ADODB.Recordset

strSQL = "select distinct Line from sap_wo_list with(nolock) where Status='20' and CheckBomPassDateTime<>'' order by Line "
Set tmpRS = Conn.Execute(strSQL)

While Not tmpRS.EOF
    cboLine.AddItem tmpRS(0)
    tmpRS.MoveNext
Wend

End Sub
