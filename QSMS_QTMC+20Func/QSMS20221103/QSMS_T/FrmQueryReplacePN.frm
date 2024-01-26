VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmQueryReplacePN 
   Caption         =   "QueryReplacePN"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4048
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
   Begin VB.Frame Frame1 
      Caption         =   "SAP_BOM"
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6855
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         TabIndex        =   7
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton CmdQuery 
         Caption         =   "Query"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox TxtModel 
         Height          =   375
         Left            =   4680
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox TxtCompPN 
         Height          =   405
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblSMT 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   0
         Left            =   3480
         TabIndex        =   4
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label lblSMT 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CompPN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1185
      End
   End
End
Attribute VB_Name = "FrmQueryReplacePN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim strSQL As String
Dim rs As ADODB.Recordset
Private Sub CmdExcel_Click()
Call CopyToExcel(rs)
End Sub

Private Sub CmdQuery_Click()
    strSQL = "EXEC QSMS_QuerySAP_BOM @PN = '" & Trim(TxtCompPN) & "' ,@Model='" & Trim(TxtModel) & "'"
    Set rs = Conn.Execute(strSQL)
    If Not rs.EOF Then
        If rs!result = "Fail" Then
              MsgBox Trim(rs.Fields("Desc")), vbOKOnly Or vbInformation, "系统提示"
                TxtCompPN = ""
                TxtCompPN.SetFocus
                Exit Sub
         Else
         Set DataGrid1.DataSource = rs
         End If
     End If
End Sub


Private Sub CopyToExcel(ByVal Rst As ADODB.Recordset)

 Dim xlWs As Object
 Dim recArray As Variant
 Dim strDB As String
 Dim fldCount As Long
 Dim recCount As Long
 Dim iCol As Long
 Dim iRow As Long
 Dim strFileName, Trans_Date As String
 ''On Error GoTo ErrHandler
 
Set xlWs = Nothing
Set xlsBook = Nothing
Set xlApp = Nothing
    
Set xlApp = CreateObject("Excel.application")
xlApp.DisplayAlerts = False
xlApp.UserControl = True
xlApp.Visible = True

    Set xlsBook = xlApp.Workbooks.Add
    Set xlWs = xlApp.Worksheets(1)
    xlApp.UserControl = True
    fldCount = Rst.Fields.Count

    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Select
        xlApp.Selection.Interior.ColorIndex = 6
        xlWs.Cells(1, iCol).Value = Rst.Fields(iCol - 1).Name
        xlApp.Selection.HorizontalAlignment = xlCenter
        xlApp.Selection.VerticalAlignment = xlCenter
    Next
    
    xlWs.Cells(2, 1).CopyFromRecordset Rst
    xlApp.Rows("2:2").Select
    xlApp.ActiveWindow.FreezePanes = True
    xlApp.ActiveWindow.SmallScroll Down:=0

    ' Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    
    Rst.Close
    Set Rst = Nothing
    Set xlWs = Nothing
    Set xlsBook = Nothing
    Set xlApp = Nothing
    
    Exit Sub

''ErrHandler:
    ''MsgBox ("CopyToExcel, " & Err.Description & "; please contact QMS!")
End Sub



