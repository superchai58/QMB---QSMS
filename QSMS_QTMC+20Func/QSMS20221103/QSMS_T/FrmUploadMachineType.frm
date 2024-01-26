VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmUploadMachineType 
   BackColor       =   &H00FFC0C0&
   Caption         =   "UpLoadMachineType[2009-07-18]"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   1095
   ScaleWidth      =   9075
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6240
      TabIndex        =   2
      Top             =   360
      Width           =   1155
   End
   Begin VB.TextBox txtFile 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   5655
   End
   Begin VB.CommandButton cmdGetMEType 
      Caption         =   "Upload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7680
      TabIndex        =   0
      Top             =   360
      Width           =   1155
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmUploadMachineType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''/**********************************************************************************
''**文 件 名: FrmMaintainDID.frm
''**Copyright (C) 2007-2010 QMS
''**文件编号:
''**创 建 人: Jeanson
''**日    期: 2007.10.01
''**描    述: QSMS Maintain DID
''
''**EQMS_ID      修 改 人     修改日期        描    述
''-----------------------------------------------------------------------------
''**             Sandy        2008.04.03      add SeqIDByLine cloumn in Machine table(0001)
''**             Sandy        2009.01.19      NB5需求MappingID长度扩大到10位(0002)
''**             Sandy        2009.03.31      Modify the function of uploading Machine data to check the Line,Vendor whether is correct or not according to Machine Name;(0003)
''**QMS          Archer       2009/07/18      Insert Machine Table Add Line Column(0004)
''***********************************************************************************1/
'Option Explicit
'Dim xlsSheetName As String
'Private Sub cmdGetMEType_Click()
'    If Trim(txtFile) = "" Then
'        MsgBox "You must select a file!", vbInformation
'        Exit Sub
'    End If
'    GetSheetName (Trim(txtFile))
'    If Load_MachineType(Trim(txtFile)) = False Then
'        MsgBox ("Fail")
'    Else
'        MsgBox ("Finish")
'    End If
'End Sub
'
'Private Sub cmdSelect_Click()
'    CommonDialog1.ShowOpen
'    txtFile = CommonDialog1.FileName
'End Sub
'Private Sub GetSheetName(filePath As String)
'    On Error GoTo ERRHEAR
'    Dim TempStr As String
'    Dim i As Long
'    Workbooks.Open filePath
'    Worksheets(1).Activate
'    xlsSheetName = ActiveSheet.Name
'No_Data:
'    'AllNum = I
'    Workbooks.Close
'    GoTo PASS
'ERRHEAR:
'    If Err.Number = 91 Then
'        Resume No_Data
'    End If
'PASS:
'End Sub
'Function Load_MachineType(strFile As String) As Boolean
'Dim xlApp As Excel.Application
'Dim xlsBook As Excel.Workbook
'Dim xlWs As Excel.Worksheets
'Dim rCount, Row_Count As Long
'Dim Deleted_Qty As Integer
'Dim Vendor  As String, Factory  As String, Machine As String, Unit As String, QTY As String, MaxSlotNum, LR As String, FujiData As String, Line As String, DeletedFlag As String
'Dim TempJobPn, tempVersion, tempLine As String
'Dim Total_Qty, Update_Qty, Insert_Qty As Long
'Dim i As Integer, SeqIDByLine As Integer
'Dim SearchChar1, MyPos1
'Dim MappingID, DIOCircuit As String
'Dim M1 As String, M2 As String, N1 As String
'Dim N2 As Integer, N3 As Integer, N4 As Integer
'Dim StrSQL As String
'Dim PreLine As String
'Dim LineArray As String
'Dim rs As ADODB.Recordset
'PreLine = ""
'If UCase(xlsSheetName) <> "MACHINETYPE" Then
'    Exit Function
'End If
'Set xlApp = CreateObject("Excel.Application")
'Let xlApp.Visible = False
'Set xlsBook = xlApp.Workbooks.Open(txtFile)
'xlApp.DisplayAlerts = False
'Load_MachineType = False
'
'    rCount = 2
'    Total_Qty = 0
'    Insert_Qty = 0
'    Update_Qty = 0
'    TempJobPn = ""
'    SearchChar1 = "-"
'
'    With xlsBook.Worksheets(Trim(xlsSheetName))
'
'          While Trim(.Cells(rCount, 1)) <> ""
'               Vendor = Trim(.Cells(rCount, 1) & vbNullString)
'               Line = Trim(.Cells(rCount, 2) & vbNullString)
'               Factory = Trim(.Cells(rCount, 3) & vbNullString)
'               Machine = Replace(Trim(.Cells(rCount, 4) & vbNullString), "'", " ")
'               Unit = Trim(.Cells(rCount, 5) & vbNullString)
'               QTY = Trim(.Cells(rCount, 6) & vbNullString)
'               MaxSlotNum = Trim(.Cells(rCount, 7) & vbNullString)
'               LR = Trim(.Cells(rCount, 8) & vbNullString)
'               MappingID = Trim(.Cells(rCount, 9) & vbNullString)
'               FujiData = Trim(.Cells(rCount, 10) & vbNullString)
'               DIOCircuit = Trim(.Cells(rCount, 11) & vbNullString)
'               DeletedFlag = Trim(.Cells(rCount, 12) & vbNullString)
'               SeqIDByLine = Trim(.Cells(rCount, 13) & vbNullString) '----(0001)
'               If Left(UCase(Machine), 1) <> UCase(Line) Then '---(0003)
'                    MsgBox "The Machine= " & Machine & "  or line= " & Line & " is define wrong! Row:" & rCount
'                    Exit Function
'               End If
'
'                '*********************JUDGE machine name Begin add by giant 061110*************************************************
'                If InStr(UCase(Machine), "NXT") = 0 Then
'                    If InStr(UCase(Machine), "OTHERS") = 0 Then
'                        'Non-NXT machine name is like ASCP7A, ASCP7B, ...
'                        If Not UCase(Machine) Like "[A-Z][S,C]???[A-Z]" Then
'                            MsgBox "The Machine name is not correct:" & Machine & vbCrLf & "Format should be [A-Z][S,C]???[A-Z]" & vbCrLf & "Row:" & rCount
'                            Exit Function
'                        End If
'                    ElseIf InStr(UCase(Machine), "OTHERS") > 0 Then
'                        'Non-NXT machine name is like ASOthers, ACOthers, ...
'                        If Not UCase(Machine) Like "[A-Z][S,C,Q,W]OTHERS*" Then
'                            MsgBox "The Machine name is not correct:" & Machine & vbCrLf & "Format should be [A-Z][S,C,Q,W]OTHERS*" & vbCrLf & "Row:" & rCount
'                            Exit Function
'                        End If
'                    End If
'                ElseIf InStr(UCase(Machine), "NXT") > 0 Then
'                    'NXT machine name is like ASNXTA01, ASNXTA02, ...
'                    If Not UCase(Machine) Like "[A-Z][S,C]NXT[A-Z][0-9][0-9]" Then
'                        MsgBox "The Machine name is not correct:" & Machine & vbCrLf & "NXT format should be [A-Z][S,C]NXT[A-Z][0-9][0-9]" & vbCrLf & "Row:" & rCount
'                        Exit Function
'                    End If
'                End If
'                '*********************JUDGE format Begin *********************************************************
'               If Vendor = "Fuji" Then
'                    MyPos1 = InStr(MappingID, SearchChar1)
'                    If MyPos1 = 0 Then
'                        MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
'                        Exit Function
'                    Else
'                        N3 = Mid(MappingID, MyPos1 + 1)
'                        N4 = Mid(MappingID, 1, MyPos1 - 1)
'                        If N3 > 20 Then
'                            MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
'                            Exit Function
'                        End If
'                        If N4 < 1 Or N4 > 5 Then
'                            MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
'                            Exit Function
'                        End If
'                    End If
'               Else
'                    M1 = Left(Machine, 2)
'                    M2 = Left(MappingID, 2)
'                    N1 = Mid(MappingID, 3, 2)
'                    If IsNumeric(Right(MappingID, 5)) = True Then
'                        N2 = CInt(Right(MappingID, 5))
'                    End If
'
'                    If N2 < 1 Then  '0002
'                        MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
'                        Exit Function
'                    End If
'
'                    If M1 <> M2 Or N1 <> "MC" Then
'                        MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
'                        Exit Function
'                    End If
'               End If
'
'                    If IsNumeric(Mid(FujiData, 1)) = True Then
'                        N2 = CInt(Mid(FujiData, 1))
'                        If N2 > 9 Then
'                            MsgBox ("Wrong FujiData :" & FujiData & ", Row:" & rCount)
'                            Exit Function
'                        End If
'                    Else
'                            MsgBox ("Wrong FujiData :" & FujiData & ", Row:" & rCount)
'                            Exit Function
'                    End If
'
'                    If IsNumeric(Mid(DIOCircuit, 1)) = True Then
'                        N2 = CInt(Mid(DIOCircuit, 1))
'                        If N2 > 9 Then
'                            MsgBox ("Wrong DIOCircuit :" & DIOCircuit & ", Row:" & rCount)
'                            Exit Function
'                        End If
'                    Else
'                            MsgBox ("Wrong DIOCircuit :" & DIOCircuit & ", Row:" & rCount)
'                            Exit Function
'                    End If
'                '*********************Delecte machine information by Line ***************************************************
'                If UCase(DeletedFlag) = "Y" Then
'                   StrSQL = "delete from Machine where Machine='" & Trim(Machine) & "' "
'                   Conn.Execute StrSQL
'                   Call InsertIntoQSMSLog("SMT_QSMS", "Delete Machine Type", "machine=" & Machine & " and factory=" & Factory & "")
'                   Deleted_Qty = Deleted_Qty + 1
'                Else
'                '*********************insert or update machine information ***************************************************
'                   StrSQL = "select * from Machine where Machine='" & Trim(Machine) & "'"
'                   Set rs = Conn.Execute(StrSQL)
'                   If rs.EOF Then
'                       StrSQL = "Insert into Machine(Vendor,Factory,Line,Machine,Unit,SeqIDByLine,Qty,MaxSlotNum,LR,MappingID,FujiData,DIOCircuit,OPID) " & _
'                                   " values('" & Trim(Vendor) & "','" & Factory & "','" & Trim(Line) & "','" & Trim(Machine) & "','" & Trim(Unit) & "','" & Trim(SeqIDByLine) & "','" & Trim(QTY) & "','" & Trim(MaxSlotNum) & "','" & Trim(LR) & "','" & Trim(MappingID) & "','" & Trim(FujiData) & "','" & Trim(DIOCircuit) & "','" & Trim(g_userName) & "')" '----(0001)
'                       Conn.Execute StrSQL
'                       Insert_Qty = Insert_Qty + 1
'                   Else
'                       StrSQL = "Update Machine set Vendor='" & Trim(Vendor) & "',Factory='" & Factory & "',Unit='" & Trim(Unit) & "',SeqIDByLine='" & Trim(SeqIDByLine) & "',Qty='" & Trim(QTY) & "',MaxSlotNum='" & Trim(MaxSlotNum) & "',LR='" & Trim(LR) & "',MappingID='" & Trim(MappingID) & "',FujiData='" & Trim(FujiData) & "',DIOCircuit='" & Trim(DIOCircuit) & "',OPID='" & Trim(g_userName) & "' where Machine='" & Trim(Machine) & "'" '----(0001)
'                       Conn.Execute StrSQL
'                       Update_Qty = Update_Qty + 1
'                   End If
'                End If
'                If PreLine <> Line Then '---(0003)
'                    StrSQL = "select line,count(distinct vendor) from machine where line='" & Trim(Line) & "' and vendor<>'DIP' group by line  having count(distinct vendor)>1"
'                    Set rs = Conn.Execute(StrSQL)
'                    If rs.EOF = False Then
'                        StrSQL = "select machine,vendor from machine where line='" & Trim(Line) & "' and vendor<>'DIP'"
'                        If rs.State Then rs.Close
'                        Set rs = Conn.Execute(StrSQL)
'                        While Not rs.EOF
'                            LineArray = LineArray + Trim(rs!Machine) + "-->" + Trim(rs!Vendor) + vbCrLf
'                            rs.MoveNext
'                        Wend
'                        MsgBox ("同一条线有多个MachineVendor，请帮助确认是否正确！" & vbCrLf & "" & LineArray & "如果不正确，请稍后把错误的资料删除或更新，点击确定可继续上传资料。")
'                    End If
'                    PreLine = Line
'                End If
'                rCount = rCount + 1
'                Total_Qty = Total_Qty + 1
'
'        Wend
'    End With
'Load_MachineType = True
'StrSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_MachineType','" & Left(Trim(strFile), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
'Conn.Execute (StrSQL)
'
'xlsBook.Close
'xlApp.Quit
'Set xlApp = Nothing
'Set xlsBook = Nothing
'Conn.Execute ("exec GenUpdateMachineType")     '*******************add by jeanson 10/10*******
'
'MsgBox "*** Load  finish ! ***" & "   " & vbCrLf & _
'             "Total Counter : " & Total_Qty & vbCrLf & _
'             "Insert succeed : " & Insert_Qty & vbCrLf & _
'             "Update succeed : " & Update_Qty & vbCrLf & _
'             "Delete Qty : " & Deleted_Qty
'
'End Function
'
'
