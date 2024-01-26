Attribute VB_Name = "mdlFunction"
Option Explicit
'需要写钩子函数，看下面：''  为窗体添加一个模块，在模块中编写钩子函数：''  首先声明使用的API函数及常量，
    
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function PlaySound Lib "Winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Const GWL_WNDPROC = -4
Public Const WM_RBUTTONUP = &H205
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public lpPrevWndProc     As Long
Private lngHWnd     As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long         '调用api实现延迟200ms
Public Function Delay(MSceond As Long)

  Dim I As Long

  If MSceond < 2 Then Exit Function

  I = GetTickCount

  Do While GetTickCount - I < MSceond

    DoEvents

  Loop

End Function

'钩子函数编写:
  
Public Sub Hook(hWnd As Long)
    lngHWnd = hWnd
    lpPrevWndProc = SetWindowLong(lngHWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
  
'钩子函数撤消:
  
Public Sub UnHook()
Dim lngReturnValue     As Long
lngReturnValue = SetWindowLong(lngHWnd, GWL_WNDPROC, lpPrevWndProc)
End Sub
  
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case uMsg
    '检测鼠标击键消息，如果是单击右键
    '      Case WM_RBUTTONUP
    Case WM_COPY, WM_PASTE  '如果是拷备或者粘贴就跳对话框出来不允许使用
        MsgBox "Can not use copy or paste function"
        WindowProc = 1
        Exit Function
    Case Else
        WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Select
End Function

''''''Added by Kyle 2010.08.10  (0077)''''''

Function FileReadAll(str As String, fname As String) As Long
    Dim fnum As Integer
    On Error Resume Next
    
    fnum = OpenInputFile(fname)
    If fnum = 0 Then
        FileReadAll = -1
    Else
        str = Input$(LOF(fnum), #fnum)
        FileReadAll = Len(str)
    End If
    Close #fnum
End Function

Function OpenInputFile(ByVal fname As String) As Integer
  Dim Fnumber As Integer
    
  On Error GoTo ErrorProcedure
  OpenInputFile = 0
  If Dir(fname) > "" Then
    Fnumber = FreeFile
    OpenInputFile = Fnumber
    Open fname For Input As #Fnumber
  End If
  Exit Function
ErrorProcedure:
  OpenInputFile = 0
End Function
''''''Added by Kyle 2010.08.10  (0077)''''''

Public Function CheckDataCode(InspectionNo As String, Vendor As String, DateCode As String) As Boolean
Dim strsql As String, DateFormat As String
Dim RS As ADODB.Recordset
Dim Dateregexp As RegExp
Dim DateMatches As MatchCollection
Dim DateMatch As Match
Set Dateregexp = New RegExp
Dateregexp.IgnoreCase = True ' 设置是否区分大小写
Dateregexp.Global = True     ' 搜索全部匹配

CheckDataCode = True
    If Trim(InspectionNo) <> "" Then
        Dateregexp.Pattern = ""
        strsql = "select * from VendorDateFormat where Vendor='" & Trim(Vendor) & "' and len(DateFormat)=len('" & Trim(DateCode) & "') and RegExp<>'' and Flag<>'Y'"
        Set RS = Conn.Execute(strsql)
        If RS.EOF = False Then
            Dateregexp.Pattern = Trim(RS.Fields("RegExp")) '设置模式
            DateFormat = Trim(RS.Fields("DateFormat"))
        Else
            If Len(Trim(DateCode)) = 8 Then
                Dateregexp.Pattern = "\b[12][0-9][0-1][0-9](0\d|1[0-2])([0-2]\d|3[0-1])\b" '设置模式(YYYYMMDD)
                DateFormat = "YYYYMMDD"
            End If
            If Len(Trim(DateCode)) = 6 Then
                Dateregexp.Pattern = "\b[0-1][0-9](0\d|1[0-2])([0-2]\d|3[0-1])\b" '设置模式(YYMMDD)
                DateFormat = "YYMMDD"
            End If
            If Len(Trim(DateCode)) = 4 Then
                Dateregexp.Pattern = "\b[0-1][0-9]([0-4]\d|5[0-6])\b" '设置模式(YYWW)
                DateFormat = "YYWW"
            End If
        End If
        If Dateregexp.Pattern <> "" Then
            Set DateMatches = Dateregexp.Execute(DateCode)  ' 执行搜索
            For Each DateMatch In DateMatches
                If DateCode <> DateMatch.Value Then
                    ''CheckDataCode = False   ''''1039
                    Exit Function
                Else
                    strsql = "Exec QSMS_ChkDateCode '" & Trim(InspectionNo) & "','" & Trim(DateCode) & "','" & Trim(DateFormat) & "'"
                    Set RS = Conn.Execute(strsql)
                    If UCase(Trim(RS.Fields("Result"))) <> "PASS" Then
                        MsgBox RS.Fields("iMessage"), vbCritical, "ErrMessage"
                        CheckDataCode = True
                    End If
                End If
            Next
        End If
    End If
End Function

Public Function ChkDateCodeSpecial(Vendor As String, compPN As String, DateCode As String) As Boolean
Dim strsql As String
Dim RS As ADODB.Recordset
    ChkDateCodeSpecial = True
    If Trim(Vendor) <> "" And Trim(DateCode) <> "" And Trim(compPN) <> "" Then
        strsql = "Exec QSMS_ChkDateCodeSpecial '" & Trim(Vendor) & "','" & Trim(compPN) & "','" & Trim(DateCode) & "'"
        Set RS = Conn.Execute(strsql)
        If UCase(Trim(RS.Fields("Result"))) <> "PASS" Then
            MsgBox RS.Fields("iMessage"), vbCritical, "ErrMessage"
             ChkDateCodeSpecial = False
        End If
    End If
End Function

Public Function IsNeedMSD(ByVal compPN As String) As Boolean
    Dim strsql As String
    Dim RS As New ADODB.Recordset
    
    strsql = "SELECT * FROM MSD_Data WHERE COMPPN='" & compPN & "'"
    Set RS = Conn.Execute(strsql)
    
    If RS.EOF = False Then
        IsNeedMSD = True
    Else
        IsNeedMSD = False
    End If
    
End Function


Public Function CheckMSD(MSD As String, compPN As String) As Boolean
    Dim strsql As String
    Dim RS As New ADODB.Recordset
    
    ''0063
    NeedMSD = IsNeedMSD(compPN)
    
    If NeedMSD = True Then
        If Trim(MSD) = "" Then
            MsgBox ("You must input MSD on CompPN=" & compPN)
            CheckMSD = False
            Exit Function
        Else
            strsql = "SELECT * FROM MSD_Current WHERE COMPPN='" & compPN & "' and CompSN='" & MSD & "'"
            Set RS = Conn.Execute(strsql)
        
            If RS.EOF Then
                MsgBox "MSD is not right for CompPN=" & compPN
                CheckMSD = False
                Exit Function
            End If
        End If
    End If
    
    CheckMSD = True
End Function

Public Sub WriteToListview(lvw As ListView, RS As ADODB.Recordset)
Dim rsTemp As New ADODB.Recordset
Dim nCH As Integer
Dim lst As ListItem
    
    lvw.ColumnHeaders.Clear
    lvw.ListItems.Clear
    For nCH = 0 To RS.Fields.Count - 1
        lvw.ColumnHeaders.Add , RS.Fields(nCH).Name, RS.Fields(nCH).Name
    Next
    While Not RS.EOF
        Set lst = lvw.ListItems.Add(, , Trim(IIf(IsNull(RS(0)), "", RS(0))))
        For nCH = 1 To lvw.ColumnHeaders.Count - 1
            lst.SubItems(nCH) = Trim(IIf(IsNull(RS(nCH)), " ", RS(nCH)))
        Next
        RS.MoveNext
    Wend
End Sub

Public Sub CopylvwToExcel(lvw As ListView, Desc As String)
    Dim xlApp As New Excel.Application
    Dim xlwk As New Excel.Workbook
    Dim xlWs As New Excel.Worksheet
    Dim I As Long, j As Long, t As Integer
    Dim rangeWidth As String
    
    Set xlApp = CreateObject("Excel.application")
    xlApp.Visible = True
    Set xlwk = xlApp.Workbooks.Add
    Set xlWs = xlwk.Worksheets(1)
    xlWs.Activate
    
    xlWs.Cells.NumberFormatLocal = "@"
    xlWs.Cells(1, 1) = Desc
    
    With xlWs
        For I = 1 To lvw.ColumnHeaders.Count
            .Cells(2, I) = lvw.ColumnHeaders(I).Text
        Next I
        
        For I = 1 To lvw.ListItems.Count
            .Cells(I + 2, 1) = lvw.ListItems(I).Text
            For j = 1 To lvw.ColumnHeaders.Count - 1
                .Cells(I + 2, j + 1) = lvw.ListItems(I).SubItems(j)
            Next j
        Next I
    End With
    
    '''设置表格的格式
    With xlApp
        ''''合并单元格，设置第一行为标题
        t = lvw.ColumnHeaders.Count
        '''''00007
        If t Mod 26 = 0 And t / 26 > 1 Then
            rangeWidth = Chr(t / 26 - 1 + 64) & Chr(26 + 64)
        ElseIf t Mod 26 > 1 And t / 26 > 1 Then
            rangeWidth = Chr(t / 26 + 64) + Chr(t Mod 26 + 64)
        Else
            rangeWidth = Chr(t + 64)
        End If

        .Range("A1:" & rangeWidth & "1").Select
        .Selection.Merge
        .ActiveCell.FormulaR1C1 = Desc
        .Selection.HorizontalAlignment = xlGeneral  '''对齐方式
        
        '''设置字体为粗体，橙黄色
        .Range("A2:" & rangeWidth & "2").Select
        .Selection.Font.Bold = True
        .Selection.Font.ColorIndex = 45
         '''加边框
        .Range("A2:" & rangeWidth & "" & lvw.ListItems.Count + 2 & "").Select
        .Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        .Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With .Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With .Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End With
    Set xlWs = Nothing
    Set xlwk = Nothing
    Set xlApp = Nothing
End Sub

Public Sub CreateFolder(sFolderPath As String)
    Dim fso As New FileSystemObject
    Dim sParentFolder As String
    Dim iPos As Integer
    iPos = InStr(1, sFolderPath, "\")
    While iPos <> 0
        If fso.FolderExists(Mid(sFolderPath, 1, iPos)) = False Then
            fso.CreateFolder Mid(sFolderPath, 1, iPos)
        End If
        iPos = InStr(iPos + 1, sFolderPath, "\")
    Wend
    
    If fso.FolderExists(sFolderPath) = False Then
        fso.CreateFolder (sFolderPath)
    End If
End Sub

'20101115 Maggie Save Printer setting in local Registry (1019)
Public Function GetPrinterSetting(frm As Form)
On Error GoTo errhandle:
    
    If GetSetting("SMT", "QSMS", "Printer") = "Zebra" Then
        frm.OptZebra.Value = True
    Else
        frm.OptSATO.Value = True
    End If
    
    If GetSetting("SMT", "QSMS", "Port") = "COM" Then
        frm.OptComp.Value = True
    ElseIf GetSetting("SMT", "QSMS", "Port") = "LPT" Then
        frm.OptPrint.Value = True
    Else
        frm.OptNetwork.Value = True
    End If
        
    If GetSetting("SMT", "QSMS", "CommPort") <> "" Then
        frm.TxtCompPort.Text = GetSetting("SMT", "QSMS", "CommPort")
    Else
        frm.TxtCompPort.Text = "1"
    End If
    
    If GetSetting("SMT", "QSMS", "Comm") <> "" Then
        frm.TxtComm.Text = GetSetting("SMT", "QSMS", "Comm")
    Else
        frm.TxtComm.Text = "9600,N,8,1"
    End If
    
    frm.OptZebra.Enabled = False
    frm.OptSATO.Enabled = False
    frm.OptComp.Enabled = False
    frm.OptPrint.Enabled = False
    frm.OptNetwork.Enabled = False
    frm.TxtCompPort.Enabled = False
    frm.TxtComm.Enabled = False
    frm.CmdCommSave.Visible = False
    
Exit Function

errhandle:
    MsgBox Err.Description
End Function

Public Function GetDIDLabelFile(frm As Form, Optional LabelType As String = "") As String ''(1080)
Dim strDPM As String
'path
GetDIDLabelFile = Settings.DIDLabelPath
If Right(Trim(GetDIDLabelFile), 1) <> "\" Then
    GetDIDLabelFile = GetDIDLabelFile & "\"
End If
'+printer  ---为防止重新设置没有重新打开界面，导致iszebra不一致的问题，还是调用frm的zebra比较好
'If GetSetting("SMT", "QSMS", "Printer") = "Zebra" Then
'    GetDIDLabelFile = GetDIDLabelFile & "Zebra_"
'Else
'    GetDIDLabelFile = GetDIDLabelFile & "SATO_"
'End If
If frm.OptZebra.Value = True Then
    GetDIDLabelFile = GetDIDLabelFile & "Zebra_"
    PrinterType = "Zebra"
Else
    GetDIDLabelFile = GetDIDLabelFile & "SATO_"
    PrinterType = "SATO"
End If

'+dpm
If GetSetting("SMT", "QSMS", "DPM") = "300" Then  ''(1080)
    strDPM = "300"
    PrintDpm = "300"
Else
    strDPM = "200"
    PrintDpm = "200"
End If
GetDIDLabelFile = GetDIDLabelFile & strDPM
''+labeltype(old/new/good/bad)
If LabelType <> "" Then
    GetDIDLabelFile = GetDIDLabelFile & "_" & LabelType
End If
''+file type
If frm.OptZebra.Value = True Then
    GetDIDLabelFile = GetDIDLabelFile & ".txt"
Else
    GetDIDLabelFile = GetDIDLabelFile & ".prn"
End If

End Function

Public Function CheckMachine(ByVal Line As String, ByVal machine As String, ByVal side As String) As Boolean
Dim strsql As String
Dim RS As New ADODB.Recordset

If Left(machine, 1) <> "*" Then ''(1047)
    'QMS             Denver         2011/01/04     can not Upload Fuji XML data                                   (1042)
    strsql = "select machine from machine where machine=" & sq(Trim(machine)) & " and line =" & sq(Trim(Line)) & " and side=" & sq(Trim(side)) '(1032)'(1043)
    'strSQL = "select machine from machine where machine like " & sq(Trim(machine) & "%") & " and line =" & sq(Trim(Line)) & " and side=" & sq(Trim(side)) '(1032)
    Set RS = Conn.Execute(strsql)
    If RS.EOF = True Then
        MsgBox ("The Machine:" & machine & " in line:" & Line & " and side: " & side & " (you uploaded) was not defined in machine,please check it in machinetype")
        CheckMachine = False
        Exit Function
    End If
End If
CheckMachine = True
End Function


Public Function EffectSound(Effect As String) As Long '1092
    Dim sEffect As String
    sEffect = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & Effect
    If Dir(sEffect) = vbNullString Then
        MsgBox sEffect & " not exist ! Please call QMS", vbCritical
        Exit Function
    End If
    PlaySound sEffect, 0, 1
End Function



Public Function MessageLabel(Status As String, msgLabel As Label, msgString As String)
    Select Case UCase(Status)
        Case "ML_PASS"
            With msgLabel
                 .BackColor = &H8000&
                 .ForeColor = &HFFFFFF
                 .FontSize = 12
                 .FontBold = True
                 .FontItalic = False
                 .Caption = msgString
            End With
            Call EffectSound("OK")
        Case "ML_ERROR"
            With msgLabel
                 .BackColor = &HFF&
                 .ForeColor = &HFFFFFF
                 .FontSize = 12
                 .FontBold = True
                 .FontItalic = False
                 .Caption = msgString
            End With
            Call EffectSound("OO")
        Case "ML_WARNING"
            With msgLabel
                 .BackColor = &HFFFF&
                 .ForeColor = &HFF
                 .FontSize = 12
                 .FontBold = True
                 .FontItalic = False
                 .Caption = msgString
            End With
            Call EffectSound("OO")
        Case "ML_INIT"
            With msgLabel
                .BackColor = &H7000&
                .ForeColor = &HFF00&
                .FontSize = 12
                .Caption = ""
            End With
                
    End Select
End Function


