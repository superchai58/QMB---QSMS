VERSION 5.00
Begin VB.Form FrmGenXLPrior 
   Caption         =   "XL Prior"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGenXLPrior 
      Caption         =   "GenXLPrior"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   1335
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
      TabIndex        =   0
      Top             =   240
      Width           =   2055
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
      Height          =   1815
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   5535
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
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FrmGenXLPrior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGenXLPrior_Click()
On Error GoTo errMsg
    Dim strSQL As String
    Dim Rs As New ADODB.Recordset
    Dim NowHour As String
    Dim NowMin As String
    
    If Trim(cboFac.text) = "" Then
        MsgBox "Please select the Factory.", vbCritical, "Information"
        Exit Sub
    End If
    
    NowHour = Hour(Now)
    NowMin = Minute(Now)
    
    If NowHour < 5 Or (NowHour = 5 And NowMin < 30) Or (NowHour = 7 And NowMin > 30) Or (NowHour > 7 And NowHour < 17) Or (NowHour = 17 And NowMin < 30) Or (NowHour = 19 And NowMin > 30) Or NowHour > 19 Then
        MsgBox "This program cannot be executed at this time.", vbCritical, "Information"
        Exit Sub
    End If

    Me.cmdGenXLPrior.Enabled = False
    

    strSQL = "QSMS_RegisterCheckBOM @WO='XLJob',@Type='0',@RtnCode=''"

    If Rs.State Then Rs.Close
    Set Rs = Conn.Execute(strSQL)

    strSQL = "XL_UploadToCurWOSeq "

    If Rs.State Then Rs.Close
    Set Rs = Conn.Execute(strSQL)

    strSQL = "XL_GetWoInputPlan "

    If Rs.State Then Rs.Close
    Set Rs = Conn.Execute(strSQL)

    strSQL = "XL_ReleaseMaterialDosage @DurationType=12,@Factory='" & Trim(cboFac.text) & "'"

    If Rs.State Then Rs.Close
    Set Rs = Conn.Execute(strSQL)

    strSQL = "XL_InheritDID @Factory='" & Trim(cboFac.text) & "',@XLType='12'"

    If Rs.State Then Rs.Close
    Set Rs = Conn.Execute(strSQL)

    strSQL = "QSMS_RegisterCheckBOM @WO='XLJob',@Type='1',@RtnCode=''"

    If Rs.State Then Rs.Close
    Set Rs = Conn.Execute(strSQL)

    strSQL = "AlarmMail_XL_Prior "

    If Rs.State Then Rs.Close
    Set Rs = Conn.Execute(strSQL)
    
    MsgBox "Program execution is complete.", vbInformation, "Information"
    
    Me.cmdGenXLPrior.Enabled = True
    
    Exit Sub
errMsg:
    MsgBox Err.Description + ",Please contact QMS"
End Sub

Private Sub Form_Load()

    Label1.Caption = "Notice:" & vbCrLf & " Only 5:30 ~ 7:30 and 17:30 ~ 19:30 can run the XL program, and it is prohibited to use it at other times. If you have any questions, you can contact QMS."

    strSQL = "select distinct Factory from Site with(nolock)"
    If Rs.State Then Rs.Close
    Set Rs = Conn.Execute(strSQL)
    While Rs.EOF = False
        cboFac.AddItem Rs.Fields("Factory")
        Rs.MoveNext
    Wend

End Sub
