Attribute VB_Name = "mdlhtml"

Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Sub Sleep Lib "kernel32.DLL" (ByVal dwMilliseconds As Long) '1167
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long '1167
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long '1167
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long '1167
Dim j As Integer

Public Sub ExportToHtml(rs As ADODB.Recordset)
Dim row As Integer, i As Integer, Name As String, html As String
Dim htmlBuilder As New StringBuilder
Dim Source As String, TransDate As String
Dim lngFolderSize As Long
Dim strFolder As String
Dim lngLength As Long
Dim pid As Long ''1167
Dim Value As String ''1170


On Error GoTo errHandler

 '''获取user的临时文件夹
lngFolderSize = 255
strFolder = String(lngFolderSize + 1, 0)
lngLength = GetTempPath(lngFolderSize, strFolder)

If lngLength > 1 Then
    strFolder = Left(strFolder, lngLength)
Else
    strFolder = vbNullString
End If

j = 1
TransDate = Format(Now, "YYYYMMDD") & Format(Now, "HHNNSS")
Source = strFolder & TransDate & "-log" & j & ".html"


If Not rs.EOF Then
'''获取结果集对应的html格式的字符串

            htmlBuilder.Append ("<html>")
            ''1171
            htmlBuilder.Append ("<head><meta http-equiv=" & """" & "content-type" & """" & " content=" & """" & " text/html; charset=utf-8 " & """" & ">")
            htmlBuilder.Append ("<title>报表</title>")
            htmlBuilder.Append ("</head>")
            htmlBuilder.Append ("<body>")
            htmlBuilder.Append ("<table  border=1 ><tr align=center bgcolor=#FFFF00>")
            
            For i = 0 To rs.Fields.Count - 1
                        
                            htmlBuilder.Append ("<th>")
                            htmlBuilder.Append (rs.Fields(i).Name)
                            htmlBuilder.Append ("</th>")
            Next i
            
            While Not rs.EOF
                htmlBuilder.Append ("<tr>")
                For i = 0 To rs.Fields.Count - 1
                    htmlBuilder.Append ("<td>")
                    '
                        If IsNull(rs.Fields(i).Value) Then  ''1170
                            Value = ""
                        Else
                            If Left(Trim(rs.Fields(i).Value), 1) = "=" Then
                                Value = "[" & Trim(rs.Fields(i).Value) & "]"
                            Else
                                Value = Trim(rs.Fields(i).Value)
                            End If
                        End If
                    htmlBuilder.Append (Value)
                    htmlBuilder.Append ("</td>")
                Next
                htmlBuilder.Append ("</tr>")
                DoEvents
                rs.MoveNext
            Wend
            
    Call SaveFile(Source, htmlBuilder.ToString(), "utf-8") '1171
    
    'Open Source For Output As #1
    'Print #1, Replace(htmlBuilder.ToString(), Chr(13), vbCrLf)
    'Close #1

    ''1167 由于shell是属于异步执行,所以增加Wait shell执行结束的代码
    pid = Shell("explorer.exe " & Source, vbNormalFocus) '获取该Process的PID
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
    ExitEvent = WaitForSingleObject(hProcess, INFINITE) '等待hProcess状态的改变,INFINITE代表无限等待
    Call CloseHandle(hProcess)
    Set rs = rs.NextRecordset
Else
    html = "<html><head><title>此报表没有数据</title></head><body bgcolor=#FFFFFF text=#000000 leftmargin=5 topmargin=5><font color=blue size=3 face=Calibri >There is no data!</font></body></html>"
    'Open Source For Output As #1
    'Print #1, Replace(html, Chr(13), vbCrLf)
    'Close #1
    
    Call SaveFile(Source, html, "utf-8") '1171
    Shell "explorer.exe " & Source
    Set rs = rs.NextRecordset
End If
    
    ''1167
    Do Until rs Is Nothing
        Sleep (1500)
        If Not rs.EOF Then
            Call ExportToHtmlSingle(rs)
            Set rs = rs.NextRecordset
        Else
            j = j + 1
            Source = strFolder & TransDate & "-log" & j & ".html"
            html = "<html><head><title>此报表没有数据</title></head><body bgcolor=#FFFFFF text=#000000 leftmargin=5 topmargin=5><font color=blue size=3 face=Calibri >There is no data!</font></body></html>"
            
          
            'Open Source For Output As #1
            'Print #1, Replace(html, Chr(13), vbCrLf)
            'Close #1
            Call SaveFile(Source, html, "utf-8") '1171
            Shell "explorer.exe " & Source
            Set rs = rs.NextRecordset
        End If
    Loop


Exit Sub


errHandler:
    'If Err.Description = "对象变量或 With 块变量未设置" Then Exit Sub
    MsgBox ("ExportToHtml, " & Err.Description & "; please contact QMS!")
End Sub
Public Sub ExportToHtmlSingle(rs As ADODB.Recordset) '1167
Dim row As Integer, i As Integer, Name As String, html As String
Dim htmlBuilder As New StringBuilder
Dim pid As Long
Dim Source As String, TransDate As String
Dim lngFolderSize As Long
Dim strFolder As String
Dim lngLength As Long
On Error GoTo errHandler

 '''获取user的临时文件夹
lngFolderSize = 255
strFolder = String(lngFolderSize + 1, 0)
lngLength = GetTempPath(lngFolderSize, strFolder)

If lngLength > 1 Then
    strFolder = Left(strFolder, lngLength)
Else
    strFolder = vbNullString
End If

j = j + 1
TransDate = Format(Now, "YYYYMMDD") & Format(Now, "HHNNSS")
Source = strFolder & TransDate & "-log" & j & ".html"

'''获取结果集对应的html格式的字符串

            htmlBuilder.Append ("<html>")
            htmlBuilder.Append ("<head>")
            ''1171
            htmlBuilder.Append ("<head><meta http-equiv=" & """" & "content-type" & """" & " content=" & """" & " text/html; charset=utf-8 " & """" & ">")
            htmlBuilder.Append ("<title>报表</title>")
            htmlBuilder.Append ("</head>")
            htmlBuilder.Append ("<body>")
            htmlBuilder.Append ("<table  border=1 ><tr align=center bgcolor=#FFFF00>")
            
            For i = 0 To rs.Fields.Count - 1
                        
                            htmlBuilder.Append ("<th>")
                            htmlBuilder.Append (rs.Fields(i).Name)
                            htmlBuilder.Append ("</th>")
            Next i
            
            While Not rs.EOF
                htmlBuilder.Append ("<tr>")
                For i = 0 To rs.Fields.Count - 1
                    htmlBuilder.Append ("<td>")
                    'htmlBuilder.Append (rs.Fields(i).Value)
                    
                     If IsNull(rs.Fields(i).Value) Then  ''1170
                            Value = ""
                        Else
                            If Left(Trim(rs.Fields(i).Value), 1) = "=" Then
                                Value = "[" & Trim(rs.Fields(i).Value) & "]"
                            Else
                                Value = Trim(rs.Fields(i).Value)
                            End If
                        End If
                    htmlBuilder.Append (Value)
                    htmlBuilder.Append ("</td>")
                Next
                htmlBuilder.Append ("</tr>")
                DoEvents
                rs.MoveNext
            Wend
            
 Call SaveFile(Source, htmlBuilder.ToString(), "utf-8") '1171
 
'Open Source For Output As #1
'Print #1, Replace(htmlBuilder.ToString(), Chr(13), vbCrLf)
'Close #1
'Shell "explorer.exe " & Source

''1167 由于shell是属于异步执行,所以增加Wait shell执行结束的代码
pid = Shell("explorer.exe " & Source, vbNormalFocus) '获取该Process的PID
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
ExitEvent = WaitForSingleObject(hProcess, INFINITE) '等待hProcess状态的改变,INFINITE代表无限等待
Call CloseHandle(hProcess)

Exit Sub


errHandler:
    MsgBox ("ExportToHtmlSingle, " & Err.Description & "; please contact QMS!")
End Sub


Private Sub SaveFile(FilePath As String, strText As String, Optional Charset As String = "uft-8")
        Dim Obj As Object
        Set Obj = CreateObject("ADODB.Stream")
        With Obj
            .Mode = 3
            .Charset = Charset
            .Open
            .WriteText strText
            .SaveToFile FilePath, 2
        End With
        Set Obj = Nothing
End Sub


