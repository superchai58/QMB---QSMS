VERSION 5.00
Begin VB.Form frmGenXLMD 
   Caption         =   "XL Material Demand"
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
      Left            =   1440
      TabIndex        =   5
      Top             =   720
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenXLMD 
      Caption         =   "GenXLMD"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   120
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
      Left            =   120
      TabIndex        =   4
      Top             =   720
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
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
      Top             =   1320
      Width           =   5535
   End
End
Attribute VB_Name = "frmGenXLMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL As String
Dim Rs As New ADODB.Recordset

Private Sub cmdGenXLMD_Click()
Dim strSQL As String
Dim Rs As New ADODB.Recordset

If Trim(cboFac.text) = "" Then
    MsgBox "Please select the Factory.", vbCritical, "Information"
    Exit Sub
End If
Me.cmdGenXLMD.Enabled = False
''''''''''(1006)''''''''''
If StrBU = "PO" Then
    strSQL = "XL_JOB_8Hours_GenMD @OPID='" & g_userName & "',@Factory='" & Trim(cboFac.text) & "'"
ElseIf StrBU = "NB4" Then   ''ㄗ1130ㄘ
    strSQL = "XL_JOB_12Hours_GenMD @OPID='" & g_userName & "',@Factory='" & Trim(cboFac.text) & "',@PNInterval='" & Trim(cboType.text) & "'"
Else
    strSQL = "XL_JOB_12Hours_GenMD @OPID='" & g_userName & "',@Factory='" & Trim(cboFac.text) & "'"
End If

If Rs.State Then Rs.Close
Set Rs = Conn.Execute(strSQL)
If Not Rs.EOF Then
    If Rs("RESULT") <> "ok" Then
        MsgBox Rs("MSG"), vbCritical, "Error Tips"
    Else
'        If StrBU = "PO" Then
'            MsgBox "生成需求成功", vbInformation, "Tips"
'        Else
            MsgBox "Generate demand successfully", vbInformation, "Tips"
'        End If
    End If
End If

Me.cmdGenXLMD.Enabled = True

End Sub

Private Sub Form_Load()
Label2.Visible = False
cboType.Visible = False
Label1.Caption = "Notice:" & vbCrLf & "    The time that XL demand can be calculated again is between 1H and 5H after the first XL execution." & vbCrLf & _
                 "For Example:" & vbCrLf & "    The execution time of XL is 7:40, and the time period during which XL demand can be calculated again is 8:40 ~ 12:40." & vbCrLf & _
                 "If it exceeds this time period, it will not be allowed to perform XL demand calculation manually, and it will be calculated automatically by the system."
strSQL = "select distinct Factory from Site with(nolock)"
If Rs.State Then Rs.Close
Set Rs = Conn.Execute(strSQL)
While Rs.EOF = False
    cboFac.AddItem Rs.Fields("Factory")
    Rs.MoveNext
Wend

'If StrBU = "PO" Then
'    frmGenXLMD.Caption = "祥龍需求"
'    Label1.Caption = "注意:" & vbCrLf & "可以再次計算XL需求的時間為7:00~10:00(中班需求), 15:00~18:00(夜班需求)或 23:00~02:00(早班需求)" & vbCrLf & "如果超過這個時間點將不允許手動重啟計算!!"
'End If
If StrBU = "NB4" Then     ''ㄗ1130ㄘ
    Label2.Visible = True
    cboType.Visible = True
    strSQL = "select XL_Type from XL_TypeDateTime order by cast(XL_Type as int) "    ''1144
    If Rs.State Then Rs.Close
    Set Rs = Conn.Execute(strSQL)
    While Rs.EOF = False
        cboType.AddItem Rs.Fields("XL_Type")
        Rs.MoveNext
    Wend
'    cboType.AddItem "6"                                                             ''1144
'    cboType.AddItem "12"
End If
End Sub
