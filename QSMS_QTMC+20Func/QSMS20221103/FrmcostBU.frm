VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmcostBU 
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8340
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CombNB 
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
      Left            =   5760
      TabIndex        =   11
      Text            =   "NB"
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5318
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
   Begin VB.TextBox TxtCostBU 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox TxtCostCenter 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox TxtDescription 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1320
      Width           =   3855
   End
   Begin VB.CommandButton CmdOK 
      BackColor       =   &H8000000E&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Cmdrefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdoutdata 
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "NB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5040
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LblCostBU 
      Caption         =   "CostBU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label LblCostCenter 
      Caption         =   "Costcenter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label LblDescription 
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   8280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   30
      X2              =   30
      Y1              =   0
      Y2              =   5760
   End
   Begin VB.Line Line3 
      X1              =   8280
      X2              =   8280
      Y1              =   0
      Y2              =   5760
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   8280
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   8160
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "FrmcostBU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim StrConn As String
 Dim pwd As String
 Dim strSQL As String
 Dim strServer As String
 Dim strUser As String
 Dim strPwd As String
 Dim strDB As String
 Dim StrConnDB As String
 Dim StrBU As String
 Dim Strsql1 As String
 Dim Strsql2 As String
 Dim ConnDB As New ADODB.Connection
 Dim RSstemp As New ADODB.Recordset

Private Sub CmdDelete_Click()
'Call CombNB_LostFocus
'If CombNB.Text = "" Then
'   MsgBox "Choice NB first"
'   Exit Sub
'End If

    If TxtCostBU.Text = "" Or TxtCostCenter.Text = "" Then
       MsgBox "Please enter CostBU data ", vbOKOnly
       Exit Sub
    End If
    


    If MsgBox("Are you Sure to Delete??", vbYesNo) = vbYes Then
        If RSstemp.State = 1 Then
            RSstemp.Close
        End If
        strSQL = "select * from QSMS_CostCenter where bu='" & Trim(TxtCostBU.Text) & "'and costcenter='" & Trim(TxtCostCenter.Text) & "'"
        RSstemp.Open strSQL, Conn, adOpenStatic, adLockReadOnly
        If RSstemp.EOF Then
            MsgBox "NO this CostCenter Data! !"
            RSstemp.Close
            Exit Sub
        End If

        If RSstemp.State = 1 Then
            RSstemp.Close
        End If
 
        strSQL = "DELETE FROM QSMS_CostCenter WHERE BU='" & Trim(TxtCostBU.Text) & "' And CostCenter ='" & Trim(TxtCostCenter.Text) & "'"
        Conn.Execute (strSQL)
        MsgBox "Delete ok"
    End If

    TxtCostBU.Text = ""
    TxtCostCenter.Text = ""
    TxtDescription.Text = ""
 
    strSQL = "select * from QSMS_CostCenter"
    Call Data(strSQL)
End Sub
Private Sub cmdOK_Click()
Dim RS As New ADODB.Recordset
'Call CombNB_LostFocus
'If CombNB.Text = "" Then
'   MsgBox "Choice NB first"
'   Exit Sub
'End If

    If TxtCostBU.Text = "" Or TxtCostCenter.Text = "" Or TxtDescription.Text = "" Then
       MsgBox "Please enter CostBU data ", vbOKOnly
       Exit Sub
    End If

 
    If MsgBox("Are you Sure to Add ??", vbYesNo) = vbYes Then
        If RS.State = 1 Then
             RSstemp.Close
        End If
        strSQL = "select * from QSMS_CostCenter where costcenter='" & Trim(TxtCostCenter.Text) & "'"
        RS.Open strSQL, Conn, adOpenStatic, adLockReadOnly
        If Not RS.EOF Then
            MsgBox "Duplicate CostCenter ! !"
            RS.Close
            Exit Sub
        End If
'        If RSstemp.State = 1 Then
'             RSstemp.Close
'        End If
        
'        RSstemp.Open strSql, Conn, adOpenStatic, adLockReadOnly
        strSQL = "insert QSMS_CostCenter (BU,CostCenter,Description) VALUES ('" & Trim(TxtCostBU.Text) & "','" & Trim(TxtCostCenter.Text) & "',N'" & Trim(TxtDescription.Text) & "')"
        Conn.Execute (strSQL)
        MsgBox "Add ok"
    End If
    TxtCostBU.Text = ""
    TxtCostCenter.Text = ""
    TxtDescription.Text = ""
    
    strSQL = "select * from QSMS_CostCenter"
    Call Data(strSQL)
End Sub


Private Sub cmdoutdata_Click()
If CombNB.Text = "" Then
   MsgBox "Choice NB first"
   Exit Sub
End If

 strSQL = "select * from QSMS_CostCenter"
 If RSstemp.State = 1 Then
    RSstemp.Close
 End If
 RSstemp.Open strSQL, Conn, adOpenStatic, adLockReadOnly
 Call CopyToExcel(RSstemp)
End Sub

Private Sub CmdRefresh_Click()
strSQL = "select * from QSMS_CostCenter"
  Call Data(strSQL)
End Sub

Private Sub CombNB_LostFocus()
'If CombNB.Text = "" Then
'   MsgBox "Choice NB first"
'   Exit Sub
'End If
'
'If RSstemp.State = 1 Then
'    RSstemp.Close
' End If
'
' strSql = "select * from QSMS_SMT_DB where bu='" & CombNB.Text & "'"
' RSstemp.Open strSql, Conn, adOpenStatic, adLockReadOnly
'
'strServer = Trim(RSstemp!SMT_Server)
'strUser = Trim(RSstemp!QSMS_DB_User)
'strPwd = Trim(RSstemp!QSMS_DB_Pwd)
'strDB = Trim(RSstemp!SMT_DB)
'RSstemp.Close
' If ConnDB.State = 1 Then
'    ConnDB.Close
' End If
'StrConnDB = "Provider=sqloledb;UID=" & Trim(strUser) & ";Server=" & Trim(strServer) & ";database=" & Trim(strDB) & ";pwd=" & Trim(strPwd) & ""
'
'
'If ConnDB.State = 1 Then
' ConnDB.Close
' End If
' ConnDB.ConnectionString = StrConnDB
' ConnDB.Open
' strSql = "select * from QSMS_CostCenter"
' Call Data(strSql)
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
With DataGrid1
        TxtCostBU.Text = .Columns(0).Value
        TxtCostCenter.Text = .Columns(1).Value
        TxtDescription = .Columns(2).Value
    End With
End Sub

Private Sub Form_Load()
'CombNB.AddItem "NB1", 0
'CombNB.AddItem "NB25", 1
'CombNB.AddItem "NB3", 2
'CombNB.AddItem "NB4", 3
'CombNB.AddItem "NB6", 4
'CombNB.AddItem "AS", 5
'CombNB.AddItem "WBU", 6
'CombNB.AddItem "ES", 7
'CombNB.AddItem "MBU", 8
'
'StrBU = ReadIniFile("COMMON", "BU", App.Path & "\set.ini")
'StrConn = ReadIniFile("Database", "Connection", App.Path & "\set.ini")
' oEncrypt.key = "Quanta"
' pwd = ReadIniFile("DataBase", "pwd", App.Path & "\set.ini")
' StrConn = StrConn & ";pwd=" & oEncrypt.Decrypt(pwd)
'
' CombNB.Text = StrBU
' If Conn.State = 1 Then
'    Conn.Close
' End If
'
' Conn.ConnectionString = StrConn
' Conn.Open
strSQL = "select * from QSMS_CostCenter order by CostCenter"
If RSstemp.State = 1 Then
    RSstemp.Close
 End If
 RSstemp.Open strSQL, Conn, adOpenStatic, adLockReadOnly
 Set DataGrid1.DataSource = RSstemp
End Sub
Private Sub Data(str As String)
If RSstemp.State = 1 Then
    RSstemp.Close
 End If
 RSstemp.Open str, Conn, adOpenStatic, adLockReadOnly
 Set DataGrid1.DataSource = RSstemp

End Sub
Public Sub CopyToExcel(ByVal Rst As ADODB.Recordset)
 Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object
 Dim recArray As Variant
 Dim strDB As String
 Dim fldCount As Integer
 Dim recCount As Long
 Dim iCol As Integer
 Dim iRow As Integer
 Dim i As Integer
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add
    'important for disabled alerts
    xlApp.DisplayAlerts = False
    ''''''“Application.DisplayAlerts = False”这行代码的主要作用是不让Excel给出提示。
    '''''''一般的，当你作一些动作时，Excel会给你一个类似于“你确认吗？”这样的对话框让你来手工判断。
    '''''''如果没有这一行，Excel会在代码准备删除一个工作表时会给我一个警告对话框。
    '''''''将Application对象的DisplayAlerts属性设置为False可以关闭类似的消息。
    Set xlWs = xlApp.Worksheets(1)
  
    xlApp.UserControl = True
    
    ' Copy field names to the first row of the worksheet
    fldCount = Rst.Fields.Count
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = Rst.Fields(iCol - 1).Name
    Next
    
    ' Check version of Excel
    If Val(Left$(xlApp.Version, 1)) > 8 Then
        xlWs.Cells(2, 1).CopyFromRecordset Rst
    Else
   
             recArray = Rst.GetRows
      recCount = UBound(recArray, 2) + 1
       For iCol = 0 To fldCount - 1
            For iRow = 0 To recCount - 1
                If IsDate(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = Format(recArray(iCol, iRow))
                ElseIf IsArray(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = "Array Field"
                End If
            Next iRow 'next record
        Next iCol 'next field
            
        xlWs.Cells(2, 1).Resize(recCount, fldCount).Value = _
            TransposeDim1(recArray)
    End If
    
    ' Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlApp.Visible = True
    ' Close ADO objects
    Rst.Close
    Set Rst = Nothing
    
    ' Release Excel references
'    Set xlWs = Nothing
    Set xlApp = Nothing
    Set xlsBook = Nothing
End Sub

Function TransposeDim1(v As Variant) As Variant
' Custom Function to Transpose a 0-based array (v)
    
    Dim X As Long, Y As Long, Xupper As Long, Yupper As Long
    Dim tempArray As Variant
    
    Xupper = UBound(v, 2)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = v(Y, X)
        Next Y
    Next X
    
    TransposeDim1 = tempArray

End Function




