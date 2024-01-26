VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmQueryDID 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Query DID Use[20100203]"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   11955
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1575
      Left            =   240
      TabIndex        =   17
      Top             =   4920
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   2778
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
      Caption         =   "Please Input DID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   11535
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FFFF&
         Caption         =   "&QUERY"
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
         Left            =   6000
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FFFF&
         Caption         =   "&ToExcel"
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
         Left            =   7800
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtDID 
         BackColor       =   &H8000000A&
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label lblDID 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "     DID"
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
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Please Select Query  Condition"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   11535
      Begin VB.CommandButton cmdQuert 
         Caption         =   "&QUERY"
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
         Left            =   7800
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdToExcel 
         Caption         =   "&ToExcel"
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
         Left            =   7800
         TabIndex        =   13
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox chkLine 
         BackColor       =   &H0000FF00&
         Caption         =   "By Line"
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
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cboLine 
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Text            =   "Line"
         Top             =   510
         Width           =   1095
      End
      Begin VB.ComboBox cboCompPN 
         Height          =   315
         Left            =   5640
         TabIndex        =   7
         Text            =   "CompPN"
         Top             =   1110
         Width           =   1575
      End
      Begin VB.ComboBox cboSlot 
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Text            =   "Slot"
         Top             =   1110
         Width           =   1095
      End
      Begin VB.ComboBox cboMachine 
         Height          =   315
         Left            =   5640
         TabIndex        =   5
         Text            =   "Machine"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox ChkMachine 
         BackColor       =   &H0000FF00&
         Caption         =   "By Machine"
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
         Left            =   3600
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox ChkSlot 
         BackColor       =   &H0000FF00&
         Caption         =   "By Slot"
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
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox ChkCompPN 
         BackColor       =   &H0000FF00&
         Caption         =   "By CompPN"
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
         Left            =   3600
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
   End
   Begin MSDataGridLib.DataGrid DG1 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   3240
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   2778
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
Attribute VB_Name = "FrmQueryDID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim RsDetail As New ADODB.Recordset     '--0017
Dim strSQL As String

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdQuert_Click()
Dim txtSQL As String
    ' judge query mode
    If chkMachine.Value = vbChecked Then
        'combine query sentence
        txtSQL = " A.Machine= '" & Trim(CboMachine.Text & " ") & "'"
    End If
      
    If ChkSlot.Value = vbChecked Then
        If Trim(txtSQL & " ") = "" Then
            txtSQL = " A.Slot =  '" & Trim(cboslot.Text & " ") & "'"
        Else
            txtSQL = txtSQL & " and A.Slot = '" & Trim(cboslot.Text & " ") & "' "
        End If
    End If

    If ChkCompPN.Value = vbChecked Then
        If Trim(txtSQL & " ") = "" Then
            txtSQL = " A.CompPN =  '" & Trim(CboCompPN.Text & " ") & "'"
        Else
            txtSQL = txtSQL & " and A.CompPN = '" & Trim(CboCompPN.Text & " ") & "' "
        End If
    End If
    If chkLine.Value = vbChecked Then
        If Trim(txtSQL & " ") = "" Then
            txtSQL = " A.machine like '" & Trim(CboLine.Text) & "%'"
        Else
            txtSQL = txtSQL & " and A.machine like '" & Trim(CboLine.Text) & "%' "
        End If
    End If
    'txtSQL = txtSQL & " order by begindatetime,machine,slot,lr,did "   'Modify by jeanson 20070821
    If Trim(txtSQL) = "" Then
        MsgBox "Please Select Query Condition£¡", vbOKOnly + vbExclamation, "¾¯¸æ"
        Exit Sub
    Else
        'show data
         txtSQL = txtSQL & " order by A.begindatetime,A.machine,A.slot,A.lr,A.did "   'Modify by jeanson 20070821
        If rs.State Then rs.Close
        rs.CursorLocation = adUseClient
        rs.Open "select A.*,Isnull(B.SplicingDT,Isnull(C.SplicingDT,'')) as SplicingDT,IsNull(B.Qty,Isnull(C.Qty,-1)) as Qty from QSMS_Verify A left Join QSMS_DID B ON A.DID=B.DID Left join QSMS_DID_Log C on A.DID=C.DID where" & txtSQL, Conn, adOpenForwardOnly, adLockReadOnly
        Set Dg1.DataSource = rs
    End If

End Sub
Private Function ToExcel()
 Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object
 Dim fldCount As Integer, iCol As Integer
 Dim str As String
 Dim Rs1 As ADODB.Recordset

 Dim strFileName, Trans_Date As String
'If ChkMachine.Value = vbUnchecked And ChkSlot.Value = vbUnchecked And ChkCompPN.Value = vbUnchecked Then
If rs.EOF Then  '0018
    MsgBox "No Data,Please Query again!", vbOKOnly + vbExclamation, "Wrong"
    Exit Function
End If
    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add

    xlApp.DisplayAlerts = False
    Set xlWs = xlApp.Worksheets(1)
    xlWs.Name = "use DID"
    xlApp.UserControl = True
'    Str = "Select * from QSMS_Verify where " & txtSQL
'    Set Rs = Conn.Execute(Str)
    
    fldCount = rs.Fields.Count
  
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = rs.Fields(iCol - 1).Name
    Next
        xlWs.Cells(2, 1).CopyFromRecordset rs
    
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlApp.Visible = True
  
    rs.Close
    Set rs = Nothing
    Trans_Date = Format(Now, "YYYYMMDD")

    Set xlApp = Nothing
    Set xlsBook = Nothing
End Function
'Austin Add
Private Function DetailToExcel()
 Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object
 Dim fldCount As Integer, iCol As Integer
 Dim str As String
 Dim Rs1 As ADODB.Recordset

 Dim strFileName, Trans_Date As String
'If ChkMachine.Value = vbUnchecked And ChkSlot.Value = vbUnchecked And ChkCompPN.Value = vbUnchecked Then
If rs.EOF And RsDetail.EOF Then  '0018
    MsgBox "No Data,Please Query again!", vbOKOnly + vbExclamation, "Wrong"
    Exit Function
End If
    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add

    xlApp.DisplayAlerts = False
    Set xlWs = xlApp.Worksheets(1)
    xlWs.Name = "use DID"
    xlApp.UserControl = True
'    Str = "Select * from QSMS_Verify where " & txtSQL
'    Set Rs = Conn.Execute(Str)
    
    fldCount = rs.Fields.Count
  
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = rs.Fields(iCol - 1).Name
    Next
        xlWs.Cells(2, 1).CopyFromRecordset rs
    
    '--0017
    Set xlWs = xlApp.Worksheets(2)
    xlWs.Name = "DID Details"
    
    fldCount = RsDetail.Fields.Count
    
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = RsDetail.Fields(iCol - 1).Name
    Next
        xlWs.Cells(2, 1).CopyFromRecordset RsDetail
    
    
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlApp.Visible = True
  
    rs.Close
    Set rs = Nothing
    Trans_Date = Format(Now, "YYYYMMDD")

    Set xlApp = Nothing
    Set xlsBook = Nothing
End Function

Private Sub cmdToExcel_Click()
Call ToExcel
End Sub

Private Sub Command1_Click()
Call DetailToExcel
End Sub

Private Sub Command2_Click()
Dim str As String
Dim Rs1 As ADODB.Recordset
If txtDID.Text <> "" Then
    str = "EXEC QueryDIDUse '" & Trim(txtDID.Text) & "'"
    Set rs = Conn.Execute(str)
    Set Dg1.DataSource = rs
    
    Set RsDetail = rs.NextRecordset  '--0017
    
    Set DataGrid1.DataSource = RsDetail
    
    
Else
    MsgBox ("please input DID"), vbCritical
End If
End Sub

Private Sub Form_Load()
    CboMachine.Clear
    strSQL = " SELECT distinct Machine FROM QSMS_Verify where Machine>'' order by Machine"
    Set rs = Conn.Execute(strSQL)
    If Not rs.EOF Then
        While Not rs.EOF
              CboMachine.AddItem Trim(rs("Machine"))
              rs.MoveNext
        Wend
        CboMachine.ListIndex = 0
    End If
    CboLine.Clear
    strSQL = "select distinct left(machine,1) as line from qsms_verify where machine>'' order by 1"
    Set rs = Conn.Execute(strSQL)
    If Not rs.EOF Then
        While Not rs.EOF
              CboLine.AddItem Trim(rs("line"))
              rs.MoveNext
        Wend
        CboLine.ListIndex = 0
    End If
    cboslot.Clear
    strSQL = "select distinct Slot from QSMS_Verify order by Slot"
    Set rs = Conn.Execute(strSQL)
    If Not rs.EOF Then
        While Not rs.EOF
              cboslot.AddItem Trim(rs("Slot"))
              rs.MoveNext
        Wend
        cboslot.ListIndex = 0
    End If

    CboCompPN.Clear
    strSQL = "select distinct CompPN from QSMS_Verify order by CompPN"
    Set rs = Conn.Execute(strSQL)
    If Not rs.EOF Then
        While Not rs.EOF
             CboCompPN.AddItem Trim(rs("CompPN"))
             rs.MoveNext
        Wend
        CboCompPN.ListIndex = 0
    End If
End Sub
