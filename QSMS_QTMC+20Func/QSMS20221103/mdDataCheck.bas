Attribute VB_Name = "mdDataCheck"
Option Explicit

Public Function FunPartNumberCheck(ByVal PartNumber As String) As String
Dim strSQL As String
Dim rs As ADODB.Recordset
strSQL = "Exec CheckFormat 'PARTNUMBER','" & PartNumber & "'"
Set rs = Conn.Execute(strSQL)
If Not rs.EOF Then
    If rs("ErrorCode") = 0 Then
        FunPartNumberCheck = "PASS"
    Else
        FunPartNumberCheck = rs("Result")
    End If
Else
    FunPartNumberCheck = "Fail"
    Exit Function
End If
End Function
