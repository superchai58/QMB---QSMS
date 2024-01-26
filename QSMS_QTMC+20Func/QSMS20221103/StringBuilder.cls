VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Buffer() As Byte   '������
Private m_lpBuffer As Long   '������ڴ��ָ��

Private m_AllocSize As Long  'ʵ�ʷ�����ڴ��С
Private m_Length As Long     '��ǰ�Ѿ�ʹ�õ��ڴ��С

Private Declare Sub CopyMemory Lib "ntdll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function StrCSpn Lib "shlwapi.dll" Alias "StrCSpnW" (ByVal lpStr&, ByVal lpCharacters&) As Long
Private Declare Function StrCSpnI Lib "shlwapi.dll" Alias "StrCSpnIW" (ByVal lpStr&, ByVal lpCharacters&) As Long

Private Declare Function StrRStr Lib "shell32.dll" Alias "StrRStrW" (ByVal lpStart&, ByVal lpEnd&, ByVal lpSrch&) As Long
Private Declare Function StrRStrI Lib "shell32.dll" Alias "StrRStrIW" (ByVal lpStart&, ByVal lpEnd&, ByVal lpSrch&) As Long

Dim nLen As Long '�¼�����ַ�������

'*************************************************************************
'   ���û�����������С
'*************************************************************************
Property Let Capacity(ByVal Size As Long)
    If Size < 1 Then Exit Property

    Size = Size + Size '����vb���ַ�����BSTR�ַ�ռ2���ֽ����Է���2�����ڴ�
    If Size <= m_AllocSize Then Exit Property '�ڴ�黹����

    '   ����һ��������
    If m_AllocSize = 0 Or m_Length = 0 Then 'û�����ݻ�û�з�����ڴ�
        ReDim m_Buffer(0 To Size)
    Else
        ReDim Preserve m_Buffer(0 To Size)
    End If
    m_lpBuffer = VarPtr(m_Buffer(0))        '������ָ��

    m_AllocSize = Size
End Property

'*************************************************************************
'   ��ȡ��ǰ�ַ�������
'*************************************************************************
Property Get Length() As Long
    Length = m_Length / 2
End Property

'*************************************************************************
'   �������
'*************************************************************************
Sub Clear()
    m_Length = 0
End Sub

'*************************************************************************
'� ���ַ����������ĺ��������ַ���
'   ��  ����Ҫ���ӵ��ַ���
'   ����ֵ����
'*************************************************************************
Sub Append(ByVal str As String)
    Const IncreaseStep = 10240

    nLen = Len(str)
    nLen = nLen + nLen

    '   ���m_AllocSize <= m_Length + nLen������ 128+ nLen ���ַ������������ڴ�
    If m_AllocSize <= m_Length + nLen Then
        Me.Capacity = (m_Length + nLen) / 2 + IncreaseStep
    End If

    '   �������ݵ�������
    Call CopyMemory(ByVal (m_lpBuffer + m_Length), ByVal StrPtr(str), nLen)
    m_Length = m_Length + nLen
End Sub

'*************************************************************************
'   ת��Ϊvb���ַ���
'*************************************************************************
Property Get ToString() As String
    If m_Length > 1 Then
        ToString = Left(m_Buffer, m_Length / 2) 'ע�ⲻҪ��Left$
    End If
End Property

'*************************************************************************
'   Ѱ�� str �ڻ������е�һ�γ��ֵ�λ��(�ַ�λ�ô� 1 ����)
'   ��  ����Ҫ���ҵ��ַ���,[�Ƿ���Ӵ�Сд(Ĭ�Ϻ��Ӵ�Сд)]
'   ����ֵ���ɹ���Ϊ-1
'*************************************************************************
Function IndexOf(ByVal str As String, Optional ByVal IsIgnoreCase As Boolean = True) As Long
    If m_Length < 1 Then
        IndexOf = -1
        Exit Function
    End If

    If IsIgnoreCase Then
        IndexOf = StrCSpnI(m_lpBuffer, StrPtr(str))
    Else
        IndexOf = StrCSpn(m_lpBuffer, StrPtr(str))
    End If

    If IndexOf >= 0 Then IndexOf = IndexOf + 1
End Function

'*************************************************************************
'   Ѱ�� str �ڻ����������һ�γ��ֵ�λ��(�ַ�λ�ô� 1 ����)
'   ��  ����Ҫ���ҵ��ַ���,[�Ƿ���Ӵ�Сд(Ĭ�Ϻ��Ӵ�Сд)]
'   ����ֵ���ɹ���Ϊ-1
'*************************************************************************
Function LastIndexOf(ByVal str As String, Optional ByVal IsIgnoreCase As Boolean = True) As Long
    If m_Length < 1 Then
        LastIndexOf = -1
        Exit Function
    End If

    If IsIgnoreCase Then
        LastIndexOf = StrRStrI(m_lpBuffer, m_lpBuffer + m_Length, StrPtr(str))
    Else
        LastIndexOf = StrRStr(m_lpBuffer, m_lpBuffer + m_Length, StrPtr(str))
    End If

    If LastIndexOf >= m_lpBuffer Then
        LastIndexOf = (LastIndexOf - m_lpBuffer) / 2 + 1
    Else
        LastIndexOf = -1
    End If
End Function

Private Sub Class_Initialize()
    m_AllocSize = 0
    m_Length = 0
End Sub

Private Sub Class_Terminate()
    Erase m_Buffer
End Sub
