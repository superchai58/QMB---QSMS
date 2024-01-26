VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmKPTS 
   Caption         =   "Trace Report[20090306]"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptListSN 
      BackColor       =   &H0000FF00&
      Caption         =   "Trace By List SN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   29
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdExcel1 
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
      Height          =   735
      Left            =   8400
      Picture         =   "FrmKPTS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton OptDID 
      BackColor       =   &H0000FF00&
      Caption         =   "Trace By DID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   25
      Top             =   120
      Width           =   1815
   End
   Begin VB.OptionButton optWO 
      BackColor       =   &H0000FF00&
      Caption         =   "Trace By WO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   22
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame FraComp 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Begin Time"
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   9615
      Begin VB.ComboBox cboModel 
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
         Left            =   1800
         TabIndex        =   31
         Top             =   3240
         Width           =   4455
      End
      Begin VB.ComboBox CboLotCode 
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
         Left            =   1800
         TabIndex        =   12
         Top             =   2640
         Width           =   4455
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
         Height          =   735
         Left            =   6840
         Picture         =   "FrmKPTS.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox CboDateCode 
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
         Left            =   1800
         TabIndex        =   10
         Top             =   2160
         Width           =   4455
      End
      Begin VB.ComboBox CboVendorCode 
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
         Left            =   1800
         TabIndex        =   9
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox TxtCompPN 
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
         Left            =   1800
         TabIndex        =   5
         Top             =   1200
         Width           =   4455
      End
      Begin MSComCtl2.DTPicker DTPBeginTime 
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   21626883
         UpDown          =   -1  'True
         CurrentDate     =   37678
      End
      Begin MSComCtl2.DTPicker DTPBeginDate 
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   240
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
         Format          =   21626883
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker DTPEndTime 
         Height          =   375
         Left            =   4560
         TabIndex        =   16
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   21626883
         UpDown          =   -1  'True
         CurrentDate     =   37678
      End
      Begin MSComCtl2.DTPicker DTPEndDate 
         Height          =   375
         Left            =   4560
         TabIndex        =   17
         Top             =   240
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
         Format          =   21626883
         CurrentDate     =   36482
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
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
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         Caption         =   "End Time"
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
         Left            =   3240
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "Begin Time"
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
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Begin Date"
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
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Lot Code"
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
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Date Code"
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
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "VendorCode"
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
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "CompPN"
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
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame FraSN 
      BackColor       =   &H00FFFFC0&
      Caption         =   "KPTS BY SN"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   9615
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5520
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Cmdsure 
         Caption         =   "Choice Excel"
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
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdExcel2 
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
         Height          =   615
         Left            =   6600
         Picture         =   "FrmKPTS.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "outout Feeder& DID "
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtComp 
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
         Left            =   7320
         TabIndex        =   24
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox TxtSN 
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
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Lab 
         BackColor       =   &H0000FF00&
         Caption         =   "Comp PN"
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
         Left            =   6120
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblSNWO 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "SN/WO"
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
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.OptionButton OptComp 
      BackColor       =   &H0000FF00&
      Caption         =   "Trace By Comp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.OptionButton OptSN 
      BackColor       =   &H0000FF00&
      Caption         =   "Trace By SN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FrmKPTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************************************************************
'**文 件 名: Smt_PMS.frm
'**Copyright (C) 2007-2015 QMS
'**文件编号:
'**创 建 人: Salon
'**日    期: 2008.11.07
'**描    述: DID Header
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**Salon      2008.11.07     Unify the program interface
'***********************************************************************************/


Option Explicit
Public FileName As String


Private Sub cmdExcel_Click()
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Object
Dim fldCount, iCol As Long
Dim Str As String
Dim strSDateTime As String
Dim strEDateTime As String
Dim rs As ADODB.Recordset

strSDateTime = Trim(Format(DTPBeginDate & " " & DTPBeginTime.Value, "YYYYMMDDHHNNSS"))
strEDateTime = Trim(Format(DTPEndDate & " " & DTPEndTime.Value, "YYYYMMDDHHNNSS"))

Str = "Exec QSMSGetSNByComp '" & Trim(TxtCompPN) & "','" & Trim(CboVendorCode) & "' ,'" & Trim(CboDateCode) & "','" & Trim(CboLotCode) & "','" & strSDateTime & "','" & strEDateTime & "','" & Trim(cboModel) & "'"
Set rs = Conn.Execute(Str)
If rs.EOF Then
   MsgBox "No record,Please check the CompPN & VendorCode & DateCode "
Else
    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add
   
    xlApp.DisplayAlerts = False
    Set xlWs = xlApp.Worksheets(1)
    xlApp.UserControl = True
    fldCount = rs.Fields.Count
    xlWs.Cells(1, 1).Value = "Comp:" & TxtCompPN
    xlWs.Cells(1, 2).Value = "VendorCode:" & Trim(CboVendorCode)
    xlWs.Cells(1, 3).Value = "DateCode:" & Trim(CboDateCode)
    xlWs.Cells(1, 4).Value = "DateCode:" & Trim(CboLotCode)
    xlWs.Cells(1, 5).Value = "Model:" & Trim(cboModel)
    For iCol = 1 To fldCount
        xlWs.Cells(2, iCol).Value = rs.Fields(iCol - 1).Name
    Next iCol
    
    xlWs.Cells(3, 1).CopyFromRecordset rs
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlApp.Visible = True
  
    rs.Close
    Set rs = Nothing
  
    Set xlApp = Nothing
    Set xlsBook = Nothing
End If

End Sub

Private Sub cmdExcel1_Click()
Dim Str As String
Dim rs As ADODB.Recordset
   ''On Local Error GoTo EcmdSave_Click
    If OptSN.Value = True Then
        Str = "EXEC QSMSGetCompBySN '" & Trim(TxtSN) & "'"
    ElseIf OptDID.Value = True Then
'        Str = "SELECT * FROM [QSMS_Verify] WHERE [DID]='" & Trim(TxtSN) & "'"
'        Set rs = Conn.Execute(Str)
'        If rs.EOF Then
'            MsgBox "Can not find the record in [QSMS_Verify] Table,Please check the Data"
'            Exit Sub
'        End If
        Str = "EXEC QSMSFoundByDID '" & Trim(TxtSN) & "'"
    ElseIf OptListSN.Value = True Then
        Call GetCompPN(FileName)
        Exit Sub
    ElseIf optWO.Value = True Then
        txtComp.SetFocus
        Exit Sub
    End If

    Set rs = Conn.Execute(Str)

    If rs.EOF Then
       MsgBox "Can not find the record,Please check the Date"
    Else
       Call CopyToExcel(rs)
       TxtSN = ""
    End If
    Exit Sub
EcmdSave_Click:
    MsgBox Err.Description + "there isn't data in QSMS_Verify table"
End Sub


Private Sub Cmdsure_Click()
CommonDialog1.Action = 1
FileName = CommonDialog1.FileName
TxtSN.Text = FileName
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

DTPBeginDate = Date
DTPEndDate = Date
DTPBeginTime = Time
DTPEndTime = Time

End Sub


Private Sub OptComp_Click()
FraComp.Top = FraSN.Top
cmdExcel1.Visible = False
FraComp.Left = FraSN.Left
FraSN.Visible = False
FraComp.Visible = True
End Sub

Private Sub OptDID_Click()
FraSN.Visible = True
cmdExcel1.Visible = True
FraComp.Visible = False
Lab.Visible = False
txtComp.Visible = False
lblSNWO = "DID"
TxtSN = ""
End Sub

Private Sub OptListSN_Click()
FraSN.Visible = True
Cmdsure.Visible = True
cmdExcel1.Visible = True
FraComp.Visible = False
Lab.Visible = False
txtComp.Visible = False
lblSNWO.Visible = False
TxtSN = ""
End Sub

Private Sub OptSN_Click()
FraSN.Visible = True
cmdExcel1.Visible = False
FraComp.Visible = False
Lab.Visible = True
txtComp.Visible = True
Cmdsure.Visible = False
lblSNWO = "Serial Number"
TxtSN = ""
lblSNWO.Visible = True
End Sub

Private Sub optWO_Click()
FraSN.Visible = True
FraComp.Visible = False
cmdExcel1.Visible = False
Lab.Visible = True
txtComp.Visible = True
lblSNWO = "Work Order"
TxtSN = ""
End Sub



Private Sub TxtCompPN_KeyPress(KeyAscii As Integer)
Dim Str As String
Dim rs As ADODB.Recordset
If KeyAscii = 13 Or KeyAscii = 9 Then
'   Str = "select distinct VendorCode from QSMS_Dispatch where CompPN='" & Trim(TxtCompPN) & "' "
   Str = "    SELECT distinct DateCode,VendorCode,LotCode from QSMS_dispatch where CompPN='" & Trim(TxtCompPN) & "' " & _
    " Union " & _
    "SELECT DateCode,VendorCode,LotCode from  QSMS_History.dbo.QSMS_dispatch where CompPN='" & Trim(TxtCompPN) & "'"
   Set rs = Conn.Execute(Str)
   CboVendorCode.Clear
   While Not rs.EOF
        CboVendorCode.AddItem Trim(rs!VendorCode)
        CboDateCode.AddItem Trim(rs!DateCode)
        CboLotCode.AddItem Trim(rs!LotCode)
        rs.MoveNext
   Wend
   CboVendorCode.SetFocus
End If
End Sub

Private Sub TxtSN_KeyPress(KeyAscii As Integer)
Dim Str As String
Dim rs As ADODB.Recordset
If KeyAscii = 13 Or KeyAscii = 9 Then
    If OptSN.Value = True Then
        Str = "EXEC QSMSGetCompBySN '" & Trim(TxtSN) & "'"
    ElseIf OptDID.Value = True Then
        If BU <> "NB3" And BU <> "AS" Then
            Str = "SELECT * FROM [QSMS_Verify] WHERE [DID]='" & Trim(TxtSN) & "'"
            Set rs = Conn.Execute(Str)
            While rs.EOF
                MsgBox "Can not find the record in [QSMS_Verify] Table,Please check the Data"
                Exit Sub
            Wend
        End If
        Str = "EXEC QSMSFoundByDID '" & Trim(TxtSN) & "'"
    ElseIf OptListSN.Value = True Then
        Call GetCompPN(FileName)
        Exit Sub
    ElseIf optWO.Value = True Then
        txtComp.SetFocus
        Exit Sub
    End If
    Set rs = Conn.Execute(Str)
'    If Rs.EOF Then
'       MsgBox "Can not find the record,Please check the Data"
'    Else
       Call CopyToExcel(rs)
       TxtSN = ""
'    End If
End If
End Sub
Private Sub txtComp_KeyPress(KeyAscii As Integer)
Dim Str As String
Dim rs As New ADODB.Recordset
If KeyAscii <> 13 Or (OptSN.Value = False And optWO.Value = False) Then Exit Sub
    If Len(Trim(txtComp)) <= 0 Then
        MsgBox "Comp PN is empty!"
        Exit Sub
    End If
    If Len(Trim(TxtSN)) <= 0 Then
        MsgBox "SN or Work order is empty!"
        Exit Sub
    End If
    
    If optWO.Value = True Then
        Str = "exec Qsms_GetDIDbyWo '" & Trim(TxtSN) & "','" & Trim(txtComp) & "'"
        Set rs = Conn.Execute(Str)
    '    Rs.Open Str, Conn
        If Not rs.EOF Then
            If rs("result") = 0 Then
                MsgBox rs("ErrDesc")
                Exit Sub
            Else
                If Trim(rs("errdesc")) <> "EXISTS" Then
                    MsgBox rs("ErrDesc")
                    Exit Sub
                End If
                Set rs = rs.NextRecordset
                Call CopyToExcel(rs)
            End If
        Else
            MsgBox "Can not get any data!"
        End If
        
    Else
        Str = "exec QSMSGetCompBySN '" & Trim(TxtSN) & "','" & Trim(txtComp) & "'"
        Set rs = Conn.Execute(Str)
        If Not rs.EOF Then
            Call CopyToExcel(rs)
        End If
    End If
    
End Sub
Public Function GetCompPN(strFileName As String)
Dim xlApp As Excel.Application
Dim xlApp1 As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlsBook As Excel.Workbook
Dim i As Integer
Dim j As Integer
Dim length As Integer
Dim SN As String
Dim Str As String
Dim rs As New ADODB.Recordset
Dim xlWs As Object
Dim recArray As Variant
Dim fldCount As Long
Dim recCount As Long
Dim iCol As Long
Dim iRow As Long
Dim LngRScount As Integer
Dim SNPNItem As Integer
Dim SNItem As Integer
Dim Sheet As Integer
Dim strSheetName As String

If strFileName = "" Then
    MsgBox "NO Choice EXCEL"
    Exit Function
End If

''On Error GoTo Handler

Set xlApp = CreateObject("Excel.Application")        ''' 导入的Excel文档
Set xlBook = xlApp.Workbooks.Open(strFileName)

Sheet = 1
Set xlApp1 = CreateObject("Excel.Application")       ''' 导出的Excel文档
Set xlsBook = xlApp1.Workbooks.Add
'Set xlWs = xlApp1.Worksheets(Sheet)
Set xlWs = xlApp1.Worksheets.Add
            strSheetName = "Report" & Sheet
xlWs.Name = strSheetName
xlApp1.ActiveWorkbook.Sheets(strSheetName).Activate
xlApp1.Visible = False

i = 2
j = 2
SNPNItem = 0
SNItem = 0

Do While (xlApp.Worksheets(1).Cells(i, 1) <> "")    '''''判断多个SN
    If xlApp.Worksheets(1).Cells(i, 2) = "" Then
        SNItem = SNItem + 1
        Str = "EXEC QSMSGetCompBySN '" & Trim(xlApp.Worksheets(1).Cells(i, 1)) & "'"
    Else
        SNPNItem = SNPNItem + 1
        Str = "exec QSMSGetCompBySN '" & Trim(xlApp.Worksheets(1).Cells(i, 1)) & "','" & Trim(xlApp.Worksheets(1).Cells(i, 2)) & "'"
    End If
    'If i < 52 Then                                       '''''控制一次只能导入查询50
    '
    '    SN = Trim(xlApp.Worksheets(1).Cells(i, 1))
    '    Str = "EXEC QSMSGetCompBySN '" & Trim(SN) & "'"
    If SNItem < 51 And SNPNItem < 15001 Then
        If rs.State = 1 Then rs.Close
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenStatic         ''''''Rs.CursorType = adOpenStatic 能保证Rs.RecordCount获得准确的行数
        Set rs = Conn.Execute(Str)
        
        If Not rs.EOF Then
            ''xlApp1.Visible = False
            xlApp1.UserControl = True
            xlApp1.DisplayAlerts = False
            
            If rs.RecordCount + j > 20000 And j <> 2 Then       ''''''如果Sheet数目多于20000个，并且不是第一次放Sheet，那么就增加一个新的Sheet。
                j = 2
                Sheet = Sheet + 1
                Set xlWs = xlApp1.Worksheets.Add
                            strSheetName = "Report" & Sheet
                xlWs.Name = strSheetName
                xlApp1.ActiveWorkbook.Sheets(strSheetName).Activate
                
                fldCount = rs.Fields.Count
                For iCol = 1 To fldCount
                    xlWs.Cells(1, iCol).Select
                    xlApp1.Selection.Interior.ColorIndex = 6
                    xlWs.Cells(1, iCol).Value = rs.Fields(iCol - 1).Name
                    xlApp1.Selection.HorizontalAlignment = xlCenter
                    xlApp1.Selection.VerticalAlignment = xlCenter
                Next
            End If

            ' Copy field names to the fiRst row of the worksheet
            fldCount = rs.Fields.Count
            If i = 2 Then             '''''''''''''''第一次才对Excel表进行Format，当从第二个SN开始，就没必要对表进行Format，I为2表示第一次
                For iCol = 1 To fldCount
                    xlWs.Cells(1, iCol).Select
                    xlApp1.Selection.Interior.ColorIndex = 6
                    xlWs.Cells(1, iCol).Value = rs.Fields(iCol - 1).Name
                    xlApp1.Selection.HorizontalAlignment = xlCenter
                    xlApp1.Selection.VerticalAlignment = xlCenter
                Next
            End If
            xlWs.Cells(j, 1).CopyFromRecordset rs
            j = rs.RecordCount + j
            
''            xlApp1.Rows("2:2").Select
''            xlApp1.ActiveWindow.FreezePanes = True
''            xlApp1.ActiveWindow.SmallScroll Down:=0
''            xlApp1.Selection.CurrentRegion.Columns.AutoFit
''            xlApp1.Selection.CurrentRegion.Rows.AutoFit
            rs.Close
            Set rs = Nothing
        End If
             i = i + 1
    Else
    
        Exit Do
    End If
Loop

    xlApp1.Visible = True
    ''''Delete the Sheet1 and sheet2 and sheet3
    Set xlWs = xlApp1.Worksheets("sheet1")
    xlWs.Delete
    Set xlWs = xlApp1.Worksheets("sheet2")
    xlWs.Delete
    Set xlWs = xlApp1.Worksheets("sheet3")
    xlWs.Delete

'    Set xlApp = Nothing
'    Set xlsBook = Nothing
    xlApp.Quit
Exit Function
Handler:
    If Err.Number <> 0 Then MsgBox "Input Error or Data Lost" & Err.Description
    xlApp.Quit
    xlApp1.Quit
End Function
