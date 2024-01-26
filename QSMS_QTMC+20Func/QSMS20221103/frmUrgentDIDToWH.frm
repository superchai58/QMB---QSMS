VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmUrgentDIDToWH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Urgent_DID_ToWH (2008.03.25)"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dgInfo 
      Height          =   4935
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   5655
      _ExtentX        =   9975
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
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.CommandButton cmdQuery 
         Caption         =   "Query"
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
         Left            =   1680
         TabIndex        =   6
         Top             =   1200
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
         Left            =   3000
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
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
         Left            =   4320
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtRefID 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ReferenceID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   1
         Top             =   517
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmUrgentDIDToWH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************************************************************
'**文 件 名: FrmUrgentDIDToWH.frm
'**Copyright (C) 2008-2010 QMS
'**文件编号:
'**创 建 人: Jing
'**日    期: 2008.03.25
'**描    述: Urgent QSMS_DID_ToWH
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'***********************************************************************************/

Option Explicit

Private Sub cmdDelete_Click()
Dim tmpSQL As String
Dim tmpRS As ADODB.Recordset

If Trim(txtRefID) = "" Then
    MsgBox ("Please input ReferenceID !"), vbCritical
    txtRefID.SetFocus
    Exit Sub
End If

If MsgBox("Do you delete it really ?", vbOKCancel, "Tip") = vbOK Then
    tmpSQL = "Select * from QSMS_DID_ToWH where ReferenceID='" & Trim(txtRefID) & "'"
    Set tmpRS = Conn.Execute(tmpSQL)
    If tmpRS.EOF Then
        MsgBox ("Can not find this ReferenceID in QSMS_DID_ToWH !"), vbCritical
        txtRefID.SetFocus
        Exit Sub
    Else
        tmpSQL = "Exec XL_UrgentToWH  '" & Trim(txtRefID) & "','" & Trim(g_userName) & "','Delete'"
        Set tmpRS = Conn.Execute(tmpSQL)
        If tmpRS("result") = 1 Then
            Set dgInfo.DataSource = Nothing
            MsgBox ("Delete OK !")
        Else
            MsgBox ("Delete Fail !")
        End If
        'Call cmdQuery_Click
    End If
End If
End Sub

Private Sub cmdQuery_Click()
Dim tmpSQL As String
Dim tmpRS As ADODB.Recordset

If Trim(txtRefID) = "" Then
    MsgBox ("Please input ReferenceID !"), vbCritical
    txtRefID.SetFocus
    Exit Sub
Else
    tmpSQL = "Select * from QSMS_DID_ToWH where ReferenceID='" & Trim(txtRefID) & "'"
    Set tmpRS = Conn.Execute(tmpSQL)
    Conn.CursorLocation = adUseClient
    If tmpRS.EOF Then
        Set dgInfo.DataSource = Nothing
        MsgBox ("Can not find this ReferenceID in QSMS_DID_ToWH !"), vbCritical
        txtRefID.SetFocus
        Exit Sub
    Else
        Set dgInfo.DataSource = tmpRS
    End If
End If
End Sub

Private Sub cmdUpdate_Click()
Dim tmpSQL As String
Dim tmpRS As ADODB.Recordset

If Trim(txtRefID) = "" Then
    MsgBox ("Please input ReferenceID !"), vbCritical
    txtRefID.SetFocus
    Exit Sub
End If

If MsgBox("Do you update it really ?", vbOKCancel, "Tip") = vbOK Then
    tmpSQL = "Select * from QSMS_DID_ToWH where ReferenceID='" & Trim(txtRefID) & "'"
    Set tmpRS = Conn.Execute(tmpSQL)
    If tmpRS.EOF Then
        MsgBox ("Can not find this ReferenceID in QSMS_DID_ToWH !"), vbCritical
        txtRefID.SetFocus
        Exit Sub
    Else
        tmpSQL = "Exec XL_UrgentToWH  '" & Trim(txtRefID) & "','" & Trim(g_userName) & "','Update'"
        Set tmpRS = Conn.Execute(tmpSQL)
        If tmpRS("result") = 1 Then
            MsgBox ("Update OK !")
        Else
            MsgBox ("Update Fail !")
        End If
        Call cmdQuery_Click
    End If
End If
End Sub

Private Sub txtRefID_KeyPress(KeyAscii As Integer)
If Trim(txtRefID) <> "" And KeyAscii = 13 Then
    Call cmdQuery_Click
End If
End Sub
