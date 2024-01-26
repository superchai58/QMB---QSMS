VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FrmChecKDID 
   BackColor       =   &H8000000B&
   Caption         =   "ChecKDID"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CboLine 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      TabIndex        =   14
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton ComExcel 
      Caption         =   "Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   13
      Top             =   3360
      Width           =   1695
   End
   Begin MCI.MMControl MM1 
      Height          =   330
      Left            =   240
      TabIndex        =   12
      Top             =   6480
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3625
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.CommandButton CmdQuery 
      Caption         =   "Query"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox TxtBarcode 
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox TxtDID 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox TxtEnd 
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Text            =   "0800"
      Top             =   960
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpEDate 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   159449089
      CurrentDate     =   42877
   End
   Begin VB.TextBox TxtBegin 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Text            =   "0800"
      Top             =   240
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpSDate 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   159449089
      CurrentDate     =   42877
   End
   Begin VB.Label LabLine 
      Caption         =   "Line:"
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
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "BarCode:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label LabDID 
      Caption         =   "DID:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "End Date："
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
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "BeginDate："
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "FrmChecKDID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim str As String
Dim strGroupID As String  ''1276
Dim rs As ADODB.Recordset

Private Sub CboLine_Click()
  strGroupID = ""
End Sub

Private Sub CmdQuery_Click()
Dim BeginDate, EndDate As String
BeginDate = Format(dtpSDate, "YYYY/MM/DD") & Trim(TxtBegin)
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
EndDate = Format(dtpEDate, "YYYY/MM/DD") & Trim(TxtEnd)
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")

If BeginDate <> "" And EndDate <> "" Or Trim(TxtDID) <> "" Then
    str = "EXEC QSMS_CHeckDID @DID='" & Trim(TxtDID) & "', @BeginDate='" & BeginDate & "',@EndDate='" & EndDate & "',@Type='Query'"
    Set rs = Conn.Execute(str)
    If Not rs.EOF Then
        Set DataGrid1.DataSource = rs
   End If
Else
    MsgBox "查询条件不能全为空", vbOKOnly Or vbInformation, "系统提示"
    Exit Sub
End If
End Sub
Private Sub Err_Sound()
    MM1.FileName = App.Path & "\OO.wav"
    MM1.Command = "open"
    MM1.Command = "play"
     Do While MM1.Mode = mciModePlay
     Loop
     MM1.Command = "close"
End Sub
Private Sub TxtDID_KeyPress(KeyAscii As Integer)
If Trim(TxtDID) <> "" And KeyAscii = 13 And Len(TxtDID) <= "30" Then
    If strChkDIDByLine = "Y" Then  ''1276
       If CboLine.Text = "" Then
          MsgBox "请选择线别", vbOKOnly Or vbInformation, "系统提示"
          Exit Sub
       End If
    End If
    TxtBarcode.SetFocus
End If
If Len(TxtDID) > "30" And KeyAscii = 13 Then
    MsgBox "请输入正确的DID", vbOKOnly Or vbInformation, "系统提示"
    TxtDID = ""
    TxtBarcode = ""
    TxtDID.SetFocus
End If
End Sub

Private Sub TxtBarcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Trim(TxtBarcode) <> "" Then
     If strChkDIDByLine = "Y" Then  ''1276 then
        str = "EXEC QSMS_CHeckDID @DID='" & Trim(TxtDID) & "',@BarCode='" & Trim(TxtBarcode) & "',@Type='Conf',@Line='" & Trim(CboLine.Text) & "'"
     Else
        str = "EXEC QSMS_CHeckDID @DID='" & Trim(TxtDID) & "',@BarCode='" & Trim(TxtBarcode) & "',@Type='Conf'"
     End If
     Set rs = Conn.Execute(str)
     If rs!result = 0 Then
          If strChkDIDByLine = "Y" And CboLine.Text <> "All" Then ''1276 then
            If strGroupID = "" Then
               strGroupID = Trim(rs.Fields("GroupID"))
            Else
               If strGroupID <> Trim(rs.Fields("GroupID")) Then
                  Call Err_Sound
                  MsgBox "相邻DID的GroupID不一致！", vbOKOnly Or vbInformation, "系统提示"
                  TxtDID = ""
                  TxtBarcode = ""
                  TxtDID.SetFocus
                  Exit Sub
               End If
            End If
          End If
          TxtDID = ""
          TxtBarcode = ""
          Call OK_Sound
          TxtDID.SetFocus
     Else
        Call Err_Sound
        MsgBox Trim(rs.Fields("Desc")), vbOKOnly Or vbInformation, "系统提示"
         TxtDID = ""
         TxtBarcode = ""
         TxtDID.SetFocus
     End If
End If
End Sub
Private Sub ComExcel_Click()
Call CopyToExcel(rs)
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
Private Function GetLine()  ''1276
str = "select distinct Line from QSMS_woGroup"
Set rs = Conn.Execute(str)
CboLine.Clear
CboLine.AddItem "All"
While Not rs.EOF
    CboLine.AddItem rs!Line
    rs.MoveNext
Wend
End Function
Private Sub OK_Sound()
    MM1.FileName = App.Path & "\OK.wav"
    MM1.Command = "open"
    MM1.Command = "play"
    Do While MM1.Mode = mciModePlay
    Loop
    MM1.Command = "close"
End Sub
Private Sub Form_Load()  ''1276
   CboLine.Visible = False
   LabLine.Visible = False
   If strChkDIDByLine = "Y" Then
        CboLine.Visible = True
        LabLine.Visible = True
        Call GetLine
   End If
End Sub

