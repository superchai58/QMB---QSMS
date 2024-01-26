VERSION 5.00
Begin VB.Form FrmUniReport 
   Caption         =   "UniReport"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   11970
   Begin VB.Frame FramePara 
      Height          =   5895
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   11775
   End
   Begin VB.TextBox TXT 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   13920
      TabIndex        =   5
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Frame ReportName 
      Caption         =   "ReportName"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.CommandButton CmdCleanPara 
         Caption         =   "&CleanPara"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6840
         Picture         =   "FrmUniReport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "&Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5520
         Picture         =   "FrmUniReport.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox CboReport 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label LblReportName 
         Caption         =   "ReportName"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Label ReportLabel 
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   12000
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
End
Attribute VB_Name = "FrmUniReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report As String
Dim Str As String
Dim rs As ADODB.Recordset
Dim WithEvents TempText As TextBox
Attribute TempText.VB_VarHelpID = -1
Dim WithEvents TempLabel As Label
Attribute TempLabel.VB_VarHelpID = -1
Dim WithEvents TempFrame As Frame
Attribute TempFrame.VB_VarHelpID = -1

Sub CreateLabelAndTextbox(ParaName As String, ID As Integer)
   ' 创建新的Textbox控件
'   Set TempLabel = Controls.Add("VB.label", labelName, TempFrame)
    Load ReportLabel(CInt(ID))
   ' 将控件移动到你所需要的地方
   ReportLabel(CInt(ID)).Move 100, 500 * CInt(ID), 1800, 300
   ' 创建时，所有的控件都是不可见的
   ReportLabel(CInt(ID)).Visible = True
   Set ReportLabel(CInt(ID)).Container = Me.FramePara
   ReportLabel(CInt(ID)).Caption = ParaName
   ' 创建新的Textbox控件
    Load TXT(CInt(ID))
   ' 将控件移动到你所需要的地方
   Set TXT(CInt(ID)).Container = Me.FramePara
   TXT(CInt(ID)).Move 1800, 500 * CInt(ID), 1800, 300
   ' 创建时，所有的控件都是不可见的
   TXT(CInt(ID)).Visible = True
End Sub
Private Sub CboReport_Click()
    Dim lngNum As Long
    Dim obj As Object
    Dim i As Integer
    For Each obj In Me.Controls
        If TypeName(obj) = "Label" And obj.Name = "ReportLabel" Then
            lngNum = lngNum + 1
        End If
    Next obj
    If lngNum > 1 Then
        MsgBox "Please Click CleanPara Button!"
        Exit Sub
    End If
    Str = "exec UniReport_GetParameter '" & Trim(CboReport) & "','1'"
    Set rs = Conn.Execute(Str)
    While Not rs.EOF
        Call CreateLabelAndTextbox(Trim(rs!Item), Trim(rs!ID))
        rs.MoveNext
    Wend
    Report = Trim(CboReport)
End Sub


Private Sub CmdCleanPara_Click()
    Dim lngNum As Long
    Dim obj As Object
    Dim i As Integer
    For Each obj In Me.Controls
        If TypeName(obj) = "Label" And obj.Name = "ReportLabel" Then
            lngNum = lngNum + 1
        End If
    Next obj
    For i = 1 To lngNum - 1
        Unload ReportLabel(i)
        Unload TXT(i)
    Next i
End Sub

Private Sub cmdExcel_Click()
Dim StrSQL As String
Dim i As Integer
    If Report <> "" Then
        Str = "exec UniReport_GetParameter '" & Trim(CboReport) & "','2'"
        Set rs = Conn.Execute(Str)
        If rs.EOF = False Then
            StrSQL = "EXEC " + Trim(rs!ObjectSP)
            For i = 1 To rs!cnt
                StrSQL = StrSQL + "  '" + TXT(i) + "',"
            Next i
            StrSQL = Left(StrSQL, Len(StrSQL) - 1)
        End If
        Set rs = Conn.Execute(StrSQL)
        If Not rs.EOF Then
           Call CopyToExcel(rs)
        Else
           MsgBox "No data found,please check!"
        End If
    Else
        MsgBox "Please Select ReportName!"
    End If
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
Str = "select distinct ObjectName from UniReport where Status<>0"
Set rs = Conn.Execute(Str)
CboReport.Clear
If rs.EOF Then MsgBox "No data"
While Not rs.EOF
      CboReport.AddItem Trim(rs!ObjectName)
      rs.MoveNext
Wend
Report = ""
End Sub


