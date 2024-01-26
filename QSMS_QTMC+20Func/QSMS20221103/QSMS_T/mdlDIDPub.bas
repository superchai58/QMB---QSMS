Attribute VB_Name = "mdlDIDPub"
Option Explicit
Private sSql As String
Private rst As New ADODB.Recordset


Public Function ChkGroupClosed(ByVal GroupID As String) As Boolean

    ChkGroupClosed = False
    sSql = "select * from QSMS_WoGroup where GroupID='" & Trim(GroupID) & "' and ClosedFlag<>'Y'"
    Set rst = Conn.Execute(sSql)
    If rst.EOF Then
       ChkGroupClosed = True
    End If
    
End Function

Public Function DIDGetRefIDByResult(strResult As String) As String
    Dim Ipos As Integer
    Dim Jpos As Integer
    Ipos = InStr(1, strResult, ":")
    Jpos = InStr(1, strResult, ",")
    
    ' sCurrRefID length must be>10
    If Ipos < 1 Or Jpos < 10 Then
        DIDGetRefIDByResult = ""
        Exit Function
    End If
    
    DIDGetRefIDByResult = Mid(strResult, Ipos + 1, Jpos - Ipos - 1)
End Function
