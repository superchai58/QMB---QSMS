VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUnChkWO 
   Caption         =   "CloseUnCheckWO(090422)"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
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
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
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
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   6360
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dgInfo 
      Height          =   4935
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8705
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmUnChkWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************************************************************
'**文 件 名: FrmUrgentDIDToWH.frm
'**Copyright (C) 2009-0422 QMS
'**文件编号:
'**创 建 人: Udall
'**日    期: 2009.04.22
'**描    述:
'
'**EQMS_ID                 修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'QMS                        Sandy       2009.06.24  check the wo if existence!(0001)
'***********************************************************************************/

Option Explicit

Private Sub CmdADD_Click()
Dim tmpSQL As String
Dim tmpRS As ADODB.Recordset

If Trim(TxtWO) = "" Then
    MsgBox ("Please input WO !"), vbCritical
    TxtWO.SetFocus
    Exit Sub
End If
    tmpSQL = "select * from CloseWO_UnCheck where WO='" & Trim(TxtWO) & "'"
    Set tmpRS = Conn.Execute(tmpSQL)
    If tmpRS.EOF = False Then
        MsgBox ("This workorder already exists.")
     Else
        tmpSQL = "Insert into CloseWO_UnCheck(WO,UID,TransDateTime) values('" & Trim(TxtWO) & "','" & Trim(g_userName) & "',dbo.formatdate(getdate(),'YYYYMMDDHHNNSS'))"
        Conn.Execute (tmpSQL)
        Call reFreshData
        
    End If

End Sub

Private Sub cmdDelete_Click()
Dim tmpSQL As String
Dim tmpRS As ADODB.Recordset

If Trim(TxtWO) = "" Then
    MsgBox ("Please input WO !"), vbCritical
    TxtWO.SetFocus
    Exit Sub
End If
    tmpSQL = "Delete CloseWO_UnCheck where WO='" & Trim(TxtWO) & "'"
    Conn.Execute (tmpSQL)
    Call reFreshData
End Sub

Private Sub txtWO_KeyPress(KeyAscii As Integer)
If Trim(TxtWO) <> "" And KeyAscii = 13 Then
    CmdADD.SetFocus
End If
End Sub

Private Sub Form_Load()
    Call reFreshData
End Sub

Private Sub reFreshData()
Dim tmpSQL As String
Dim tmpRS As ADODB.Recordset
    tmpSQL = "select * from CloseWO_UnCheck order by TransDateTime desc"
    Set tmpRS = Conn.Execute(tmpSQL)
    Set dgInfo.DataSource = tmpRS
End Sub

