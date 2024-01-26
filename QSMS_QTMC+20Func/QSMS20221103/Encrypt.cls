VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Encrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mKey(30) As Long
Private mSeed As Long
Private tmp As String

Public Property Let key(Str As String)
Dim I As Long
    mSeed = 0
    For I = 1 To Len(Str)
        mSeed = mSeed + (Asc(Mid(Str, I, 1)) * I) Mod 177
    Next I
    
    For I = 1 To 30
        mKey(I) = (mSeed * I) Mod (128 - I)
    Next I
End Property

Public Function Encrypt(Str As String)
Dim I As Long
Dim AdjPos As String
Dim code As Long, cnt As Long

    AdjPos = ""
    tmp = ""
    For I = 1 To Len(Str)
        code = Asc(Mid(Str, I, 1)) Xor mKey(I Mod 31)
        If code < 35 Then
            'skip ASC 1~35
            code = code + 35
            cnt = cnt + 1
            AdjPos = AdjPos & Chr(35 + I)
        End If
        tmp = tmp & Chr(code)
    Next I
    'skip ASC 1~35
    Encrypt = Chr(35 + cnt) & AdjPos & tmp
End Function

Public Function Decrypt(Str As String)
Dim I As Long, pos As Long
Dim AdjPos As String, data As String
Dim DataArry(100) As String
Dim code As Long, cnt As Long

    tmp = ""
    cnt = Asc(Left(Str, 1)) - 35
    AdjPos = Mid(Str, 2, cnt)
    data = Mid(Str, cnt + 2)
'    For i = 1 To Len(Data)
'        DataArry(i) = Mid(Data, i, 1)
'    Next i
    For I = 1 To cnt
        pos = Asc(Mid(AdjPos, I, 1)) - 35
        Mid(data, pos, 1) = Chr(Asc(Mid(data, pos, 1)) - 35)
    Next I
    For I = 1 To Len(data)
        tmp = tmp & Chr((Asc(Mid(data, I, 1)) Xor mKey(I Mod 31)))
    Next I
    Decrypt = tmp
End Function

