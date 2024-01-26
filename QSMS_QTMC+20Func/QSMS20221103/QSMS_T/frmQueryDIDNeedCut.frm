VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmQueryDIDNeedCut 
   Caption         =   "Query did need cut 080813"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   14130
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6015
      Left            =   5040
      TabIndex        =   9
      Top             =   1320
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   10610
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
   Begin VB.Frame Frame2 
      Caption         =   "Component PN list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   4575
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   460
         Left            =   3480
         TabIndex        =   18
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox cboSheetName 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   480
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton cmdGetFile 
         Caption         =   ".."
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtFileName 
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtPN 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   2775
      End
      Begin VB.ListBox lstPN 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   4860
         ItemData        =   "frmQueryDIDNeedCut.frx":0000
         Left            =   240
         List            =   "frmQueryDIDNeedCut.frx":0002
         TabIndex        =   6
         Top             =   2040
         Width           =   4095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sheet"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COMP PN"
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
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   5040
      TabIndex        =   0
      Top             =   300
      Width           =   8895
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
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
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Excel"
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
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdQuery 
         Caption         =   "Query"
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
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
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
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtTO 
         Height          =   495
         Left            =   6840
         TabIndex        =   10
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   98238465
         CurrentDate     =   39673
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   495
         Left            =   4440
         TabIndex        =   11
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   98238465
         CurrentDate     =   39673
      End
      Begin VB.Label Label3 
         Caption         =   "TO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmQueryDIDNeedCut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Private Sub cmdClear_Click()
lstPN.Clear
End Sub

Private Sub cmdExcel_Click()
Dim rs As New ADODB.Recordset
Dim strPNList As String
Dim i As Integer
Dim sDate As String
Dim eDate As String
If lstPN.ListCount < 1 Then
    MsgBox "Please input component PN", vbCritical
    Exit Sub
End If
For i = 0 To lstPN.ListCount - 1
    lstPN.ListIndex = i
    strPNList = strPNList & "'" & Trim(lstPN.Text) & "',"
Next i
strPNList = "(" & Left(strPNList, Len(strPNList) - 1) & ")"
sDate = Format(dtFrom, "yyyymmdd")
eDate = Format(dtTO, "yyyymmdd")
If Len(strPNList) > 0 Then
    strSQL = "select top 5000 * from qsms_did where remainqty>0 and transdatetime between '" & Trim(sDate) & "' and '" & Trim(eDate) & "' " & _
            " and did not like '%-A%' AND QTY<>Remainqty and realqty>0 and comppn in " & strPNList & " order by comppn,line,transdatetime"
    Set rs = Conn.Execute(strSQL)
    If rs.EOF Then
        MsgBox "No data!"
        Exit Sub
    End If
    Call CopyToExcel(rs)
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdGetFile_Click()
CommonDialog1.ShowOpen
txtFileName = CommonDialog1.FileName
cboSheetName.Clear
Call ReadAllSheetName(txtFileName)
cboSheetName.Enabled = True
End Sub
Private Sub ReadAllSheetName(filePath As String)
    On Error GoTo ERRHEAR
    Dim TempStr As String
    Dim i As Long
    Workbooks.Open filePath
    Worksheets(1).Activate
    i = 0
    Do
       cboSheetName.AddItem ActiveSheet.Name
       ' TempDim(I) = TempStr
        ActiveSheet.Next.Select
        i = i + 1
    Loop
'    cboSheetName.AddItem "ALL"
No_Data:
    'AllNum = I
    Workbooks.Close
    GoTo PASS
ERRHEAR:
    If Err.Number = 91 Then
        Resume No_Data
    End If
PASS:
End Sub

Private Sub cmdLoad_Click()
Dim xlApp As New Excel.Application
Dim xlwk As New Excel.Workbook
Dim xlWs As New Excel.Worksheet
Dim COMPPN As String
Dim i As Integer
Dim rs As New ADODB.Recordset

Set xlApp = CreateObject("Excel.application")
xlApp.Visible = False
Set xlwk = xlApp.Workbooks.Open(txtFileName)
Set xlWs = xlwk.Worksheets(cboSheetName.Text)
xlWs.Activate
i = 2
While Trim(xlWs.Cells(i, 1)) <> ""
    COMPPN = Trim(xlWs.Cells(i, 1))
    If CheckExistsInPNList(COMPPN) = False Then
        strSQL = "select top 1 * from qsms_did where comppn='" & Trim(COMPPN) & "'"
        Set rs = Conn.Execute(strSQL)
        If rs.EOF = False Then
            lstPN.AddItem COMPPN
        Else
            MsgBox "Can not find this PN:" & COMPPN
        End If
    End If
    i = i + 1
Wend
    
End Sub

Private Sub CmdQuery_Click()
Dim rs As New ADODB.Recordset
Dim strPNList As String
Dim i As Integer
Dim sDate As String
Dim eDate As String
If lstPN.ListCount < 1 Then
    MsgBox "Please input component PN", vbCritical
    Exit Sub
End If
For i = 0 To lstPN.ListCount - 1
    lstPN.ListIndex = i
    strPNList = strPNList & "'" & Trim(lstPN.Text) & "',"
Next i
strPNList = "(" & Left(strPNList, Len(strPNList) - 1) & ")"
sDate = Format(dtFrom, "yyyymmdd")
eDate = Format(dtTO, "yyyymmdd")
If Len(strPNList) > 0 Then
    strSQL = "select top 5000 * from qsms_did where remainqty>0 and transdatetime between '" & Trim(sDate) & "' and '" & Trim(eDate) & "' " & _
            " and did not like '%-A%' AND QTY<>Remainqty and realqty>0 and comppn in " & strPNList & " and wogroup in (select distinct groupid from QSMS_WOGroup where ClosedFlag<>'Y') order by comppn,line,transdatetime"
    Set rs = Conn.Execute(strSQL)
    Set DataGrid1.DataSource = rs
End If
End Sub

Private Sub Form_Load()
dtFrom = Now - 1
dtTO = Now + 1
End Sub

Private Sub txtPN_Click()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtPN_KeyPress(KeyAscii As Integer)
Dim rs As New ADODB.Recordset
If KeyAscii <> 13 Or Trim(txtPN) = "" Then Exit Sub
strSQL = "select top 1 * from qsms_did where comppn='" & Trim(txtPN) & "'"
Set rs = Conn.Execute(strSQL)
If rs.EOF Then
    txtPN = ""
    txtPN.SetFocus
    MsgBox "Can not find this componet PN in system", vbCritical
    Exit Sub
End If
If CheckExistsInPNList(Trim(txtPN)) = False Then
    lstPN.AddItem Trim(txtPN)
End If
txtPN = ""
txtPN.SetFocus
End Sub
Private Function CheckExistsInPNList(PN As String)
Dim i As Integer
For i = 0 To lstPN.ListCount - 1
    lstPN.ListIndex = i
    If Trim(PN) = Trim(lstPN.Text) Then
        CheckExistsInPNList = True
        Exit Function
    End If
Next i
End Function
