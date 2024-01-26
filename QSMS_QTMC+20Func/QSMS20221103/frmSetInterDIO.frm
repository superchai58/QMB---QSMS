VERSION 5.00
Begin VB.Form frmSetInterDIO 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Set DIO and InterLock[20100810]"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMachine 
      BackColor       =   &H8000000A&
      Caption         =   "Check2"
      Height          =   375
      Index           =   15
      Left            =   7800
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "SET A LINE"
      Height          =   255
      Left            =   9480
      TabIndex        =   20
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Frame fraCondition 
      BackColor       =   &H00C0FFC0&
      Caption         =   "by Line or Machine"
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   5535
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   3960
         TabIndex        =   21
         Text            =   "Combo1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkAllLine 
         Alignment       =   1  'Right Justify
         Caption         =   "SET ALL LINE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Set Machine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   12
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame fraMachine 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Machine in Line"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8775
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   47
         Left            =   7680
         TabIndex        =   55
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   46
         Left            =   6600
         TabIndex        =   54
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   45
         Left            =   5520
         TabIndex        =   53
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   44
         Left            =   4440
         TabIndex        =   52
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   43
         Left            =   3360
         TabIndex        =   51
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   42
         Left            =   2280
         TabIndex        =   50
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   41
         Left            =   1200
         TabIndex        =   49
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   40
         Left            =   120
         TabIndex        =   48
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   39
         Left            =   7680
         TabIndex        =   47
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   38
         Left            =   6600
         TabIndex        =   46
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   37
         Left            =   5520
         TabIndex        =   45
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   36
         Left            =   4440
         TabIndex        =   44
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   35
         Left            =   3360
         TabIndex        =   43
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   34
         Left            =   2280
         TabIndex        =   42
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   33
         Left            =   1200
         TabIndex        =   41
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   32
         Left            =   120
         TabIndex        =   40
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   31
         Left            =   7680
         TabIndex        =   39
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   30
         Left            =   6600
         TabIndex        =   38
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   29
         Left            =   5520
         TabIndex        =   37
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   28
         Left            =   4440
         TabIndex        =   36
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   27
         Left            =   3360
         TabIndex        =   35
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   26
         Left            =   2280
         TabIndex        =   34
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   25
         Left            =   1200
         TabIndex        =   33
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   24
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   23
         Left            =   7680
         TabIndex        =   31
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   22
         Left            =   6600
         TabIndex        =   30
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   21
         Left            =   5520
         TabIndex        =   29
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   20
         Left            =   4440
         TabIndex        =   28
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   19
         Left            =   3360
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   18
         Left            =   2280
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   17
         Left            =   1200
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   16
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   14
         Left            =   6600
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   13
         Left            =   5520
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   12
         Left            =   4440
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   11
         Left            =   3360
         TabIndex        =   14
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   10
         Left            =   2280
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   9
         Left            =   1200
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   7
         Left            =   7680
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   6
         Left            =   6600
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   5
         Left            =   5520
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   4
         Left            =   4440
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   3
         Left            =   3360
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   2
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkMachine 
         BackColor       =   &H8000000A&
         Caption         =   "Check2"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmSetInterDIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit
Dim str As String
Dim cQty As Integer
Dim ClickFlg As Integer
Dim RS As New ADODB.Recordset

Private Function GetLine()
Dim str As String
Dim RS As ADODB.Recordset
str = "select distinct Line from QSMS_woGroup order by Line"
Set RS = Conn.Execute(str)
CboLine.Clear
While Not RS.EOF
    CboLine.AddItem RS!Line
    RS.MoveNext
Wend
End Function
Private Function GetSet()
Dim rsMachine As New ADODB.Recordset
Dim i As Integer
Call Refence
str = "select machine,DisableInterlock from Machine where substring(machine,1,1)='" & Trim(CboLine) & "'"
Set rsMachine = Conn.Execute(str)

'CboLine.Clear
'While Not rsMachine.EOF
'20100810  Denver  一条线 Machine 数量不应该超出48，如果大于48 应该Machine设置有问题
While Not rsMachine.EOF And i < 48
    chkMachine(i).Visible = True
    chkMachine(i).Caption = rsMachine!machine
    If rsMachine!DisableInterlock = "1" Then
        chkMachine(i).Value = 1
    Else
        chkMachine(i).Value = 0
    End If
    rsMachine.MoveNext
    i = i + 1
        
Wend
cQty = rsMachine.RecordCount
'fraMachine.Visible = True
End Function

Private Sub CboLine_Click()
Call GetSet
frmSetInterDIO.Height = 4920
fraMachine.Visible = True
ClickFlg = 2
End Sub

Private Sub chkAllLine_Click()
ClickFlg = 1
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
Dim strtemp As String
If ClickFlg = 1 Then
    If chkAllLine.Value = 1 Then
        str = "update Machine set DisableInterlock='1' "
    Else
        str = "update Machine set DisableInterlock='0' "
    End If
    Conn.Execute str
    'add log 2009/03/17 by lynn
    strtemp = Replace(str, "'", "")
    Conn.Execute ("insert into qms_log(system_name,event_no,sn,user_name,desc1,trans_date) values('SetDIOInterLock','1','','" & UID & "','" & strtemp & "',dbo.formatdate(getdate(),'YYYYMMDDHHNNSS'))")
ElseIf ClickFlg = 2 Then
    For i = 0 To cQty - 1
        If chkMachine(i).Value = 1 Then
            str = "update Machine set DisableInterlock='1'where machine='" & Trim(chkMachine(i).Caption) & "' "
        Else
            str = "update Machine set DisableInterlock='0'where machine='" & Trim(chkMachine(i).Caption) & "' "
        End If
        Conn.Execute str
        'add log 2009/03/17 by lynn
        strtemp = Replace(str, "'", "")
        Conn.Execute ("insert into qms_log(system_name,event_no,sn,user_name,desc1,trans_date) values('SetDIOInterLock','1','','" & UID & "','" & strtemp & "',dbo.formatdate(getdate(),'YYYYMMDDHHNNSS'))")
    Next i
End If
    MsgBox "set InterLock or DIO is OK!"
End Sub

Private Sub Form_Load()
Dim str1 As String, str2 As String
Dim Rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
str1 = "select * from Machine"
Set Rs1 = Conn.Execute(str1)
str2 = "select * from Machine where DisableInterlock='1'"
Set Rs2 = Conn.Execute(str2)
If Rs1.RecordCount = Rs2.RecordCount Then
    chkAllLine.Value = 1
End If
Call GetLine
fraMachine.Visible = False
End Sub

Private Sub Refence()
Dim i As Integer
For i = 0 To 47
    chkMachine(i).Caption = ""
    chkMachine(i).Visible = False
Next
End Sub
