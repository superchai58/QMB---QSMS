Attribute VB_Name = "mdlStringProcess"
Option Explicit
Const BASE16 = "0123456789ABCDEF"
Const BASE34 = "0123456789ABCDEFGHJKLMNPQRSTUVWXYZ"  '0~9;A~Z except I, O
Const BASE32 = "0123456789ABCDEFGHJKLMNPRSTVWXYZ"    '0~9;A~Z except I, O, Q,U

Public Function StrBetween(src As String, gap1 As String, gap2 As String, Optional start As Long = 1) As String
  Dim pos1 As Long, pos2 As Long

  pos1 = InStr(start, src, gap1)
  If pos1 = 0 Then
    StrBetween = ""
    Exit Function
  End If
  pos2 = InStr(pos1 + 1, src, gap2)
  If Len(Trim(src)) >= pos2 And pos2 >= pos1 And (pos2 - pos1 - Len(gap1)) > 0 Then
    StrBetween = Mid(src, pos1 + Len(gap1), pos2 - pos1 - Len(gap1))
  Else
    StrBetween = ""
  End If
End Function

Public Function squote(ByVal Field As String) As String
   squote = "'" & Field & "'"
End Function

Public Function sq(ByVal Field As String) As String
   sq = "'" & Field & "'"
End Function

Public Function GetKeyValue(src, key) As String
Dim pos1 As Long, pos2 As Long
    pos1 = InStr(UCase(src), UCase(Trim(key) & "="))
    If pos1 > 0 Then
        pos1 = pos1 + Len(Trim(key) & "=")
        pos2 = InStr(pos1, src, ";")
        If pos2 > 0 Then
            GetKeyValue = Mid(src, pos1, pos2 - pos1)
            Exit Function
        End If
        
        pos2 = InStr(pos1, src, """")
        If pos2 > 0 Then
            GetKeyValue = Mid(src, pos1, pos2 - pos1)
            Exit Function
        End If
        GetKeyValue = Mid(src, pos1)
    Else
        GetKeyValue = ""
    End If
End Function


Public Function GetKeyValueM(src, key) As String
Dim pos1 As Long
Dim i As Long
Dim aLine
Dim oneLine As String

    aLine = Split(src, vbCrLf)
    key = UCase(key)
    For i = 0 To UBound(aLine)
        If InStr(UCase(aLine(i)), key) > 0 Then
            oneLine = aLine(i)
            Exit For
        End If
    Next i
    
    If i > UBound(aLine) Then
        GetKeyValueM = ""
        Exit Function
    End If
    
    pos1 = InStr(UCase(oneLine), Trim(key) & "=")
    If pos1 > 0 Then
        pos1 = pos1 + Len(Trim(key) & "=")
        GetKeyValueM = Mid(oneLine, pos1)
    Else
        GetKeyValueM = ""
    End If
End Function


Public Function inArray(Arry, str As String, Optional Compare As Long = 0) As Long
Dim i As Long

inArray = -1

For i = LBound(Arry) To UBound(Arry)
    If Compare = 0 Then
        If Trim(UCase(Arry(i))) = Trim(UCase(str)) Then
            inArray = i
            Exit Function
        End If
    Else
        If Trim(Arry(i)) = Trim(str) Then
            inArray = i
            Exit Function
        End If
    End If
Next i
End Function

Public Function inArray2(Arry, str As String, Optional Compare As Integer = 0) As Integer
Dim i As Integer

inArray2 = -1

For i = LBound(Arry, 2) To UBound(Arry, 2)
    If Compare = 0 Then
        If Trim(UCase(Arry(0, i))) = Trim(UCase(str)) Or Trim(UCase(Arry(1, i))) = Trim(UCase(str)) Then
            inArray2 = i
            Exit Function
        End If
    Else
        If Trim(Arry(0, i)) = Trim(str) Or Trim(Arry(1, i)) = Trim(str) Then
            inArray2 = i
            Exit Function
        End If
    End If
Next i
End Function

Public Function RightPad(data As String, length As Long, padder As String) As String
Dim i As Long
Dim result As String
  
  RightPad = data
  If Len(padder) > 1 Then
    MsgBox ("Padder only can be one digit!")
    Exit Function
  End If
  If Len(data) <= length Then
    result = data & String(length - Len(data), padder)
  Else
    result = Left(data, length)
  End If
  
  RightPad = result
End Function

Public Function padding(data As String, length As Long, padder As String) As String
Dim i As Long
Dim result As String
  padding = data
  If Len(padder) > 1 Then
    MsgBox ("Padder only can be one digit!")
    Exit Function
  End If
  If Len(data) <= length Then
    result = data & String(length - Len(data), padder)
  Else
    result = Left(data, length)
  End If
  
  padding = result
End Function

Public Function LeftPad(data As String, length As Long, padder As String) As String
Dim i As Long
Dim result As String
  LeftPad = data
  If Len(padder) > 1 Then
    MsgBox ("Padder only can be one digit!")
    Exit Function
  End If
  If Len(data) <= length Then
    result = String(length - Len(data), padder) & data
  Else
    result = Left(data, length)
  End If
  
  LeftPad = result
End Function


Public Function StrReverse(str As String) As String
Dim i As Long
Dim strOut As String
    strOut = ""
    For i = Len(str) To 1 Step -1
        strOut = strOut & Mid(str, i, 1)
    Next i
    StrReverse = strOut
End Function


Public Function SetFrmCaption(Caption As String, connStr As String) As String
Dim DataSource As String, DB As String
    DataSource = GetKeyValue(connStr, "Server")
    If DataSource = "" Then
         DataSource = GetKeyValue(connStr, "Data Source")
    End If
    
    DB = GetKeyValue(connStr, "Database")
    If DB = "" Then
         DB = GetKeyValue(connStr, "Initial Catalog")
    End If
    SetFrmCaption = Caption & "; Server:" & DataSource & " ; DB:" & DB

End Function
   
' get name from full filename/path
'
Public Function ExtractName(sSpecIn As String, BaseOnly As Boolean) As String
   Dim sSpecOut As String
   Dim nCnt As Long
   Dim nDot As Long
   
   On Local Error Resume Next
   '
   ' strip path from front
   '
   If InStr(sSpecIn, "\") Then
      For nCnt = Len(sSpecIn) To 1 Step -1
         If Mid$(sSpecIn, nCnt, 1) = "\" Then
            sSpecOut = Mid$(sSpecIn, nCnt + 1)
            Exit For
         End If
      Next nCnt
   ElseIf InStr(sSpecIn, ":") = 2 Then
      sSpecOut = Mid$(sSpecIn, 3)
   Else
      sSpecOut = sSpecIn
   End If
   '
   ' if we're looking for only the base filename,
   ' strip out any extension
   '
   If BaseOnly Then
      nDot = InStr(sSpecOut, ".")
      If nDot Then
         sSpecOut = Left$(sSpecOut, nDot - 1)
      End If
   End If
   '
   ' return to caller
   '
   ExtractName = sSpecOut
End Function


'
' get path from full filename/path
'
Public Function ExtractPath(sSpecIn As String) As String
   Dim nCnt As Long
   Dim sSpecOut As String
   
   On Local Error Resume Next
   '
   ' strip filename from back
   '
   If InStr(sSpecIn, "\") Then
      For nCnt = Len(sSpecIn) To 1 Step -1
         If Mid$(sSpecIn, nCnt, 1) = "\" Then
            sSpecOut = Left$(sSpecIn, nCnt)
            Exit For
         End If
      Next nCnt
   ElseIf InStr(sSpecIn, ":") = 2 Then
      sSpecOut = CurDir$(sSpecIn)
      If Len(sSpecOut) = 0 Then
         sSpecOut = CurDir$
      End If
   Else
      sSpecOut = CurDir$
   End If
   '
   ' make sure we terminate with a \
   '
   If Right$(sSpecOut, 1) <> "\" Then
      sSpecOut = sSpecOut + "\"
   End If
   '
   ' return to caller
   '
   ExtractPath = UCase$(sSpecOut)
End Function


Public Function SNDate(strDate As String, ByVal SNDateFormat) As String
Dim Mon1 As String, WW As String
    
    Mon1 = Hex(Month(strDate))
    WW = DatePart("WW", strDate, vbThursday)
  
    Select Case UCase(SNDateFormat)
        Case "YM"
            SNDate = Mid(Year(strDate), 4, 1) & Mon1
        Case "YWW"
            SNDate = Mid(Year(strDate), 4, 1) & WW
        Case "YYWW"
            SNDate = Mid(Year(strDate), 3, 2) & WW
        Case "D"
            SNDate = Mid(BASE32, 1 + Day(strDate), 1)
        Case "NO DATE"
            SNDate = ""
        Case Else
            SNDate = ""
    End Select
End Function
