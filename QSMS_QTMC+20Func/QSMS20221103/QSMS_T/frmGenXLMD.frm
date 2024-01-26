VERSION 5.00
Begin VB.Form frmGenXLMD 
   Caption         =   "��������"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1680
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.ComboBox cboFac 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenXLMD 
      Caption         =   "GenXLMD"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Factory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   5535
   End
End
Attribute VB_Name = "frmGenXLMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSql As String
Dim rs As New ADODB.Recordset

Private Sub cmdGenXLMD_Click()
Dim strSql As String
Dim rs As New ADODB.Recordset

If Trim(cboFac.Text) = "" Then
    MsgBox "Please select the Factory.", vbCritical, "Information"
    Exit Sub
End If
Me.cmdGenXLMD.Enabled = False
''''''''''(1006)''''''''''
If StrBU = "PO" Then
    strSql = "XL_JOB_8Hours_GenMD @OPID='" & g_userName & "',@Factory='" & Trim(cboFac.Text) & "'"
ElseIf StrBU = "NB4" Then   ''��1130��
    strSql = "XL_JOB_12Hours_GenMD @OPID='" & g_userName & "',@Factory='" & Trim(cboFac.Text) & "',@PNInterval='" & Trim(cboType.Text) & "'"
Else
    strSql = "XL_JOB_12Hours_GenMD @OPID='" & g_userName & "',@Factory='" & Trim(cboFac.Text) & "'"
End If

If rs.State Then rs.Close
Set rs = Conn.Execute(strSql)
If Not rs.EOF Then
    If rs("RESULT") <> "ok" Then
        MsgBox rs("MSG"), vbCritical, "Error Tips"
    Else
        If StrBU = "PO" Then
            MsgBox "�ͦ��ݨD���\", vbInformation, "Tips"
        Else
            MsgBox "��������ɹ�", vbInformation, "Tips"
        End If
    End If
End If

Me.cmdGenXLMD.Enabled = True

End Sub

Private Sub Form_Load()
Label2.Visible = False
cboType.Visible = False
Label1.Caption = "ע��:" & vbCrLf & "    �����ٴμ���XL�����ʱ���ǵ�һ��XL�ܹ�1H~5H֮��" & vbCrLf & _
                 "����:" & vbCrLf & "    XLʱ��Ϊ7:40 ��ô�����ٴμ��������ʱ���Ϊ8:40~12:40," & vbCrLf & _
                 "����������ʱ��㽫���������ֶ���,������ϵͳ�Զ�����."
strSql = "select distinct Factory from Site"
If rs.State Then rs.Close
Set rs = Conn.Execute(strSql)
While rs.EOF = False
    cboFac.AddItem rs.Fields("Factory")
    rs.MoveNext
Wend

If StrBU = "PO" Then
    frmGenXLMD.Caption = "���s�ݨD"
    Label1.Caption = "�`�N:" & vbCrLf & "�i�H�A���p��XL�ݨD���ɶ���7:00~10:00(���Z�ݨD), 15:00~18:00(�]�Z�ݨD)�� 23:00~02:00(���Z�ݨD)" & vbCrLf & "�p�G�W�L�o�Ӯɶ��I�N�����\��ʭ��ҭp��!!"
End If
If StrBU = "NB4" Then     ''��1130��
    Label2.Visible = True
    cboType.Visible = True
    strSql = "select  XL_Type from XL_TypeDateTime order by cast(XL_Type as int ) "    ''1144
    If rs.State Then rs.Close
    Set rs = Conn.Execute(strSql)
    While rs.EOF = False
        cboType.AddItem rs.Fields("XL_Type")
        rs.MoveNext
    Wend
'    cboType.AddItem "6"                                                             ''1144
'    cboType.AddItem "12"
End If
End Sub
