Attribute VB_Name = "modReadIniFile"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function ReadIniFile(ByVal strSection As String, ByVal strKey As String, strFname As String) As String
Dim strValue As String * 255
Dim intRet As Integer

On Error Resume Next
intRet = GetPrivateProfileString(strSection, strKey, "", strValue, Len(strValue), strFname)
ReadIniFile = Left(strValue, intRet)
End Function
