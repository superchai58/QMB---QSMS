VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmUrgentInsertWO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Urgent WO (2010.07.21)"
   ClientHeight    =   10395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   14970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtWO 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   12960
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "QSMS_WO_XL"
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   7320
      Width           =   12495
      Begin MSDataGridLib.DataGrid dgQSMS_WO_XL 
         Height          =   2655
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   4683
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "QSMS_WoInputPlan"
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   12495
      Begin MSDataGridLib.DataGrid dgWoInputPlan 
         Height          =   1935
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3413
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
   End
   Begin VB.Frame Frame2 
      Caption         =   "XL_CurWOSeq"
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   12495
      Begin MSDataGridLib.DataGrid dgCurWoSeq 
         Height          =   2055
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3625
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
   End
   Begin VB.Frame Frame1 
      Caption         =   "XL_WOPlanSeq"
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12495
      Begin MSDataGridLib.DataGrid dgWOPlanSeq 
         Height          =   1815
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3201
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
   End
   Begin VB.CommandButton cmdUrgentInsertWO 
      Caption         =   "紧急插单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12960
      TabIndex        =   0
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "紧急插单只需要输入大板的工单号码，小板会随着大板一起插单。"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   12840
      TabIndex        =   13
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "输入工单，点击Query按钮，确定工单的信息正后,点击紧急插单按钮。"
      Height          =   735
      Left            =   12840
      TabIndex        =   12
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "WorkOrder:"
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
      Left            =   13080
      TabIndex        =   5
      Top             =   240
      Width           =   1425
   End
End
Attribute VB_Name = "FrmUrgentInsertWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************************************************************
'**文 件 名: FrmUrgentInsertWO.frm
'**Copyright (C) 2008-2010 QMS
'**文件编号:
'**创 建 人: Jing
'**日    期: 2008.02.27
'**描    述: Urgent Insert WO
'
'EQMS_ID             **修 改 人     修改日期        描    述
'-------------------------------------------------------------------------------------------------
'QMS                 **Lynn         2010.03.22     Modify Urgent insert By WO (0001)
'QMS                 **Lynn         2010.05.07     Modify Query Urgent WO also use SP:XL_SpecialCaseByWO(0002)
'QMS                 **Lynn         2010.07.21     Add Workdate & Shift for Urgent WO, the new method only need input Big board WO (0003)
'***********************************************************************************/

Option Explicit

Private Sub cmdQuery_Click()
Dim strSQL As String
Dim tmpRS As New Recordset

On Err GoTo errhandle:
If Trim(txtWO) = "" Then
    MsgBox ("Input WorkOrder please !"), vbInformation
    txtWO.SetFocus
    Exit Sub
Else
    ''(0002)
    strSQL = "exec XL_SpecialCaseByWO_New '','" & Trim(txtWO) & "'"  ''(1140)
    Set tmpRS = Conn.Execute(strSQL)
    
    If tmpRS.EOF = False Then
        Set dgWOPlanSeq.DataSource = tmpRS
    End If
    
    Set tmpRS = tmpRS.NextRecordset
    If tmpRS.EOF = False Then
        Set dgCurWoSeq.DataSource = tmpRS
    End If
    
    Set tmpRS = tmpRS.NextRecordset
    If tmpRS.EOF = False Then
        Set dgWoInputPlan.DataSource = tmpRS
    End If
    
    Set tmpRS = tmpRS.NextRecordset
    If tmpRS.EOF = False Then
        Set dgQSMS_WO_XL.DataSource = tmpRS
    End If
    
End If
Exit Sub

errhandle:
    MsgBox Err.Description
End Sub

Private Sub cmdUrgentInsertWO_Click()
Dim strSQL As String
Dim rs As New adodb.Recordset
On Err GoTo errhandle:

If Trim(txtWO) = "" Then
    MsgBox ("Input WorkOrder please !"), vbInformation
    txtWO.SetFocus
    Exit Sub
Else
    strSQL = "exec ChkUrgentWO '" & Trim(txtWO) & "'"
    Set rs = Conn.Execute(strSQL)
    If rs!result = 1 Then
        MsgBox "PMC did not upload this WO information of Date= " + rs!WorkDate + " and Shit= " + rs!Shift + "，please check it!"
        Exit Sub
    ElseIf rs!result = 2 Then
        MsgBox "Can not find XLTime according to this WO, please check the wo!"
        Exit Sub
    ElseIf rs!result = 0 Then
        If MsgBox("Do you really want to insert WO:" + Trim(txtWO.Text) + ",Date:" + rs!WorkDate + ",Shift:" + rs!Shift + " ?", vbOKCancel, "Tip") = vbOK Then
            strSQL = "exec XL_SpecialCaseByWO_new '" & g_userName & "','" & txtWO & "','" & rs!WorkDate & "','" & rs!Shift & "'" ''(0003)
            Conn.Execute (strSQL)
            MsgBox ("OK!"), vbInformation
        End If
    End If
End If

Exit Sub
errhandle:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
     ''20100903  Kyle added to solve the encoding problem of UI
     If StrBU = "PO" Then
        Label2.Caption = "块虫腹絏翴阑Query絋﹚虫癟タ絋翴阑候础虫秙"
        Label3.Caption = "候础虫惠璶块狾虫腹絏狾穦繦帝狾癬础虫"
        cmdUrgentInsertWO.Caption = "候础虫"
     End If
End Sub
