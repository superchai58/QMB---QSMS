VERSION 5.00
Begin VB.Form frmCompPNReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CompPN Report (20080108)"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboType 
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
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "EXCEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4680
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtWO2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1320
      TabIndex        =   2
      Top             =   1170
      Width           =   2895
   End
   Begin VB.TextBox txtWO1 
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
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
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
      Left            =   480
      TabIndex        =   6
      Top             =   2160
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "WO2:"
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
      Left            =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "WO1:"
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
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   660
   End
End
Attribute VB_Name = "frmCompPNReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/*********************************************************************************
'**文 件 名: FrmCompPNReport.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: Jing
'**日    期: 2007.12.17
'**描    述: Get different CompPN from Two PCB
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------------
'Jing          2008.01.08     check two wo if they are in one pcb group       (0001)
'-----------------------------------------------------------------------------------

Option Explicit

Private Sub cmdExcel_Click()
Dim strSQL As String
Dim strType As String, strWO1 As String, strWO2 As String
Dim rsTmp As New ADODB.Recordset

strWO1 = Trim(txtWO1.Text)
strWO2 = Trim(txtWO2.Text)
strType = Trim(CboType.Text)

If strWO1 = "" Or strWO2 = "" Then
    MsgBox ("Input WO1 and WO2 please !")
    txtWO1.SetFocus
    Exit Sub
End If

If strType = "" Then
    MsgBox ("Select a report type please !")
    CboType.SetFocus
    Exit Sub
End If

Select Case UCase(strType)
    Case "DIFFERENTCOMPPNINFO"
        Call GetDifferentCompPNInfo
    Case Else
        MsgBox ("Report Type is err !"), vbCritical
End Select

End Sub

Private Sub Form_Load()
CboType.AddItem "DifferentCompPNInfo"
End Sub

Private Sub txtWO1_Click()
SendKeys "{HOME}+{END}"
End Sub

Private Sub txtWO1_KeyPress(KeyAscii As Integer)
Dim strSQL As String
Dim rsTmp As New ADODB.Recordset

If (KeyAscii = 13 Or KeyAscii = 9) And txtWO1 <> "" Then
    'strSQL = "select * from qsms_wo where work_order='" & Trim(txtWO1.Text) & "'"
    
    ''''''''''''Updated by Jing 2008.01.08  (0001)''''''''''''''''
    strSQL = "select * from sap_wo_list where wo='" & Trim(txtWO1.Text) & "'"
    Set rsTmp = Conn.Execute(strSQL)
    If rsTmp.EOF Then
        MsgBox ("This WO is not in SAP_WO_LIST !"), vbCritical
        Call txtWO1_Click
        Exit Sub
    Else
        txtWO2.SetFocus
    End If
End If
End Sub

Private Sub txtWO2_Click()
SendKeys "{HOME}+{End}"
End Sub

Private Sub txtWO2_KeyPress(KeyAscii As Integer)
Dim strSQL As String
Dim rsTmp As New ADODB.Recordset

If (KeyAscii = 13 Or KeyAscii = 9) And txtWO2 <> "" Then
    'strSQL = "select * from qsms_wo where work_order='" & Trim(txtWO2.Text) & "'"
    
    '''''''''''updated by Jing 2008.01.08   (0001)'''''''''
    strSQL = "select * from sap_wo_list where wo='" & Trim(txtWO2.Text) & "'"
    Set rsTmp = Conn.Execute(strSQL)
    If rsTmp.EOF Then
        MsgBox ("This WO is not in SAP_WO_LIST !")
        Call txtWO2_Click
        Exit Sub
    Else
        CboType.SetFocus
    End If
End If
End Sub

Public Sub GetDifferentCompPNInfo()
Dim strSQL As String
Dim strType As String, strWO1 As String, strWO2 As String, strGroup1 As String, strGroup2 As String
Dim rsTmp As New ADODB.Recordset

strWO1 = Trim(txtWO1.Text)
strWO2 = Trim(txtWO2.Text)
strType = Trim(CboType.Text)

'strSQL = "select GroupID from qsms_wogroup where work_order='" & strWO1 & "'"

''''''''''updated by Jing 2008.01.08    (0001)'''''''''''
strSQL = "select [group] from SAP_WO_LIST where wo='" & strWO1 & "'"

Set rsTmp = Conn.Execute(strSQL)
If rsTmp.EOF Then
    MsgBox ("The WO1 is not in SAP_WO_LIST !"), vbCritical
    Exit Sub
Else
    strGroup1 = rsTmp("group")
    Set rsTmp = Nothing
End If

'strSQL = "select GroupID from qsms_wogroup where work_order='" & strWO2 & "'"

''''''''''Updated by Jing 2008.01.08    (0001)''''''''''''
strSQL = "select [group] from SAP_WO_LIST where wo='" & strWO2 & "'"

Set rsTmp = Conn.Execute(strSQL)
If rsTmp.EOF Then
    MsgBox ("The WO2 is not in SAP_WO_LIST !"), vbCritical
    Exit Sub
Else
    strGroup2 = rsTmp("group")
    Set rsTmp = Nothing
End If

If (strWO1 = strWO2) Or (strGroup1 = strGroup2) Then
    MsgBox ("They are in one PCB Group!"), vbInformation        ''''Updated by Jing 2008.01.08 (0001)
    Exit Sub
Else
    strSQL = "Exec XL_GetDiffCompPNInfo '" & strGroup1 & "','" & strGroup2 & "'"
    Set rsTmp = Conn.Execute(strSQL)
    If rsTmp.EOF Then
        MsgBox ("NO DATA !"), vbInformation
    Else
        Call CopyToExcel(rsTmp)
    End If
End If

    
End Sub
