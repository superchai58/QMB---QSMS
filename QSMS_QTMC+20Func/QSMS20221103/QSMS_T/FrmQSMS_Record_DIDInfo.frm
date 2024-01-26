VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmQSMS_Record_DIDInfo 
   Caption         =   "QSMS_Record_DIDInfo[2017/07/17]"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   360
      TabIndex        =   24
      Top             =   5160
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4471
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
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3960
         TabIndex        =   27
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton CmdQuery 
         Caption         =   "Query"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2160
         TabIndex        =   26
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton CmdADD 
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   25
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6000
         TabIndex        =   23
         Top             =   3840
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPDate 
         Height          =   375
         Left            =   960
         TabIndex        =   22
         Top             =   3240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   160563201
         CurrentDate     =   42933
      End
      Begin VB.TextBox TxtMachine 
         Height          =   375
         Left            =   6120
         TabIndex        =   21
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox TxtTime 
         Height          =   375
         Left            =   3960
         TabIndex        =   20
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox TxtQty 
         Height          =   405
         Left            =   6120
         TabIndex        =   19
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox TxtModel 
         Height          =   375
         Left            =   3600
         TabIndex        =   18
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox TxtLine 
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox TxtDID 
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox TxtEndTime 
         Height          =   375
         Left            =   4320
         TabIndex        =   8
         Text            =   "0000"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TxtBeginTime 
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Text            =   "0000"
         Top             =   360
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPEndTime 
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   160563201
         CurrentDate     =   42933
      End
      Begin MSComCtl2.DTPicker DTPBeginTime 
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   160563201
         CurrentDate     =   42933
      End
      Begin VB.Label LabLine 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "线别:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label LabDID 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "机种:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   2640
         TabIndex        =   15
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label LabDID 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "数量:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   5280
         TabIndex        =   14
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label LabDID 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "日期:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label LabDID 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "烧录时间:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   12
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label LabDID 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "机台:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5280
         TabIndex        =   11
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label LabDID 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label LabBegin 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BeginTime:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LabBegin 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "End   Time:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label LabEnd 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "End   Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label LabBegin 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BeginDate:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmQSMS_Record_DIDInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim DateTime As String
Dim BeginTime As String
Dim EndTime As String
Dim RS  As New ADODB.Recordset

Private Sub DataGrid1_Click()
   TxtDID = DataGrid1.Columns("DID").Value
   TxtLine = DataGrid1.Columns("Line").Value
   TxtModel = DataGrid1.Columns("Model").Value
   TxtQty = DataGrid1.Columns("Qty").Value
   TxtTime = DataGrid1.Columns("BurnTime").Value
   TxtMachine = DataGrid1.Columns("Machine").Value
End Sub
Private Sub CmdDelete_Click()
If TxtDID <> "" Then
    strSQL = "EXEC QSMS_Record_BIOSOrEC @DID ='" & Trim(TxtDID) & "',@Model ='" & Trim(TxtModel) & "',@Line ='" & Trim(TxtLine) & "',@Machine ='" & Trim(TxtMachine) & "',@Qty ='" & Trim(TxtQty) & "'," & _
            " @BurnTime ='" & Trim(TxtTime) & "', @DateTime ='" & Trim(DateTime) & "' ,@Type='Delete'"
       Conn.Execute (strSQL)
       ''Call reFreshData
       MsgBox "删除成功", vbOKOnly Or vbInformation, "系统提示"
       TxtDID = ""
       TxtModel = ""
       TxtMachine = ""
       TxtQty = ""
       TxtLine = ""
       TxtTime = ""
       
       strSQL = "SELECT * FROM QSMS_Record_DIDInfo"
       Set RS = Conn.Execute(strSQL)
       Set DataGrid1.DataSource = RS
End If
End Sub

Private Sub TxtDID_KeyPress(KeyAscii As Integer)
Dim strSQL As String
If KeyAscii = vbKeyReturn And TxtDID.Text <> "" Then
    strSQL = "SELECT Line,SUBSTRING(JOBPN,3,3) AS Model,TotalQty FROM QSMS_Dispatch  WHERE DID= '" & Trim(TxtDID) & "'"
    Set RS = Conn.Execute(strSQL)
        If Not RS.EOF Then
            TxtModel = Trim(RS!Model)
            TxtLine = Trim(RS!Line)
            TxtQty = Trim(RS!TotalQty)
        Else
            MsgBox "The DID Not Find", vbCritical
            TxtDID.Text = ""
            TxtDID.SetFocus
        End If
End If
End Sub

Private Sub CmdADD_Click()
Dim strSQL As String
DateTime = Format(DTPDate, "yyyymmdd")
If TxtDID = "" Or TxtTime = "" Or TxtMachine = "" Then
       MsgBox "添加信息不能有空，请确认！", vbOKOnly Or vbInformation, "系统提示"
       TxtMachine = ""
       TxtTime = ""
       TxtDID = ""
       TxtDID.SetFocus
       Exit Sub
Else
       strSQL = "SELECT TOP 1 0 FROM QSMS_Record_DIDInfo WHERE DID='" & Trim(TxtDID) & "'"
       Set RS = Conn.Execute(strSQL)
       If RS.EOF Then
            strSQL = "EXEC QSMS_Record_BIOSOrEC @DID ='" & Trim(TxtDID) & "',@Model ='" & Trim(TxtModel) & "',@Line ='" & Trim(TxtLine) & "',@Machine ='" & Trim(TxtMachine) & "',@Qty ='" & Trim(TxtQty) & "'," & _
            " @BurnTime ='" & Trim(TxtTime) & "' , @DateTime ='" & Trim(DateTime) & "' , @Type='ADD'"
            Conn.Execute (strSQL)
            Call reFreshData
            MsgBox "OK", vbOKOnly Or vbInformation, "系统提示"
            TxtDID = ""
            TxtModel = ""
            TxtMachine = ""
            TxtQty = ""
            TxtLine = ""
            TxtTime = ""
      Else
          MsgBox "该DID已存在，请确认！", vbOKOnly Or vbInformation, "系统提示"
      End If
End If
End Sub
Private Sub reFreshData()
Dim tmpSQL As String
    tmpSQL = "Select * from QSMS_Record_DIDInfo where DID ='" & Trim(TxtDID) & "'AND Model ='" & Trim(TxtModel) & "'AND Line ='" & Trim(TxtLine) & "' AND Machine ='" & Trim(TxtMachine) & "' AND Qty ='" & Trim(TxtQty) & "'" & _
       " AND BurnTime ='" & Trim(TxtTime) & "' AND Date_Time ='" & Trim(DateTime) & "'"
    Set RS = Conn.Execute(tmpSQL)
    Set DataGrid1.DataSource = RS
End Sub

Private Sub CmdQuery_Click()
Dim strSQL As String

''If TxtBeginTime = "" Or TxtEndTime = "" Then
    ''TxtBeginTime = "0000"
    ''TxtEndTime = "0000"
''End If

BeginTime = Format(DTPBeginTime, "yyyymmdd") & TxtBeginTime
EndTime = Format(DTPEndTime, "yyyymmdd") & TxtEndTime

If BeginTime <> "" Or EndTime <> "" Or TxtDID <> "" Or TxtLine <> "" Or TxtModel <> "" Or TxtMachine <> "" Then
    strSQL = "EXEC QSMS_Record_BIOSOrEC @DID ='" & Trim(TxtDID) & "',@Model ='" & Trim(TxtModel) & "',@Line ='" & Trim(TxtLine) & "',@Machine ='" & Trim(TxtMachine) & "',@BeginDate ='" & Trim(BeginTime) & "'," & _
       " @EndDate ='" & Trim(EndTime) & "',@Type='Query'"
Else
    MsgBox "输入不能为空", vbOKOnly Or vbInformation, "系统提示"
    TxtDID = ""
    TxtModel = ""
    TxtMachine = ""
    TxtLine = ""
    Exit Sub
End If
    Set RS = Conn.Execute(strSQL)
    Set DataGrid1.DataSource = RS
End Sub
Private Sub CmdExcel_Click()
Call CopyToExcel(RS)
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
