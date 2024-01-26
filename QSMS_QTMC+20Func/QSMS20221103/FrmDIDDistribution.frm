VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FrmDIDDistribution 
   Caption         =   "DIDDistribution"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "PortType"
      Height          =   1665
      Left            =   480
      TabIndex        =   14
      Top             =   0
      Width           =   1695
      Begin VB.OptionButton OptComp 
         Caption         =   "COM Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.OptionButton OptPrint 
         Caption         =   "LPT Port"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   15
         Top             =   960
         Width           =   1200
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1665
      Left            =   2280
      TabIndex        =   9
      Top             =   0
      Width           =   2880
      Begin VB.TextBox TxtCompPort 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   11
         Text            =   "1"
         Top             =   480
         Width           =   1140
      End
      Begin VB.TextBox TxtCommSetting 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   10
         Text            =   "9600,N,8,1"
         Top             =   990
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "CompPort："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label Label7 
         Caption         =   "Settings："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   12
         Top             =   990
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Info"
      Height          =   1575
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   7335
      Begin VB.TextBox txtDID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   7
         Top             =   840
         Width           =   3615
      End
      Begin VB.ComboBox CmbSide 
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox CmbLine 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "DID"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Side"
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Line"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
   End
   Begin MSCommLib.MSComm MSComm_Com 
      Left            =   7800
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   480
      TabIndex        =   0
      Top             =   3480
      Width           =   7335
      Begin VB.Label location 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   2040
         TabIndex        =   17
         Top             =   720
         Width           =   2895
      End
   End
   Begin VB.Label lblstatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   7200
      Width           =   7215
   End
End
Attribute VB_Name = "FrmDIDDistribution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
Dim RS As ADODB.Recordset
Dim strsql As String

 

Private Sub Form_Load()
Call GetLineData
CmbSide.AddItem "S"
CmbSide.AddItem "C"
CmbSide.AddItem "Q"
CmbSide.AddItem "W"
End Sub
Private Sub ClearData()
    txtDID.Text = ""
    location.Caption = ""
    Call SendData("OFF")
End Sub
 
 

Private Sub txtDID_KeyPress(KeyAscii As Integer)
    Dim Line As String, Side As String, Machine As String
    Dim I As Integer
    If KeyAscii = 13 And txtDID.Text <> "" Then
        If UCase(Trim(txtDID.Text)) = "OFF" Then '' 0FF  表示灭灯
            MessageLabel "ML_INIT", lblstatus, ""
            Call ClearData
            Exit Sub
        End If
        strsql = "select  Line,Side,Machine from QSMS_Dispatch with(nolock) where DID ='" & Trim(txtDID.Text) & "'"
        Set RS = Conn.Execute(strsql)
        If RS.EOF = False Then
            Line = Trim(UCase(RS("line")))
            Side = Trim(UCase(RS("side")))
            Machine = Trim(UCase(RS("Machine")))
            If Line <> Trim(CmbLine.Text) Or Side <> Trim(CmbSide.Text) Then
                Call ChangeColor("OFF", "OFF")
                MessageLabel "ML_ERROR", lblstatus, "Line or side  is  not correct "
                Call ClearData
                Exit Sub
            End If
            strsql = "select LightLocation  from Machine_LightLocation_Map with(nolock) where  Line='" & Line & "' and Side='" & Side & "'   and  Machine='" & RS("machine") & "'"
            Set RS = Conn.Execute(strsql)
            If RS.EOF = False Then
                MessageLabel "ML_PASS", lblstatus, "OK"
                Call ChangeColor("ON", UCase(RS("LightLocation")))
                txtDID.Text = ""
                Exit Sub
            Else
                MessageLabel "ML_ERROR", lblstatus, "Please Call Shopfloor PE, and  Define relation between machine  and  LightLocaion"
                Call ChangeColor("OFF", "OFF")
                Call ClearData
                Exit Sub
            End If
        Else
            MessageLabel "ML_ERROR", lblstatus, " this DID  is not  exists !"
            Call ChangeColor("OFF", "OFF")
            Call ClearData
            Exit Sub
        End If
    End If
End Sub
Private Sub ChangeColor(status As String, content As String)
    If status = "ON" Then
        location.BackColor = vbGreen
        location.Caption = content
        
    ElseIf status = "OFF" Then         ''OFF 表示灭灯
        location.BackColor = &H8000000C
        location.Caption = ""
    End If
    Call SendData(content)
End Sub


Private Sub GetLineData()
strsql = "select distinct Line from Sap_WO_List where Trans_Date>dbo.FormatDate(getdate()-30,'YYYYMMDDHHNNSS') order by Line "
Set RS = Conn.Execute(strsql)
CmbLine.Clear
While Not RS.EOF
    CmbLine.AddItem UCase(Trim(RS!Line))
    RS.MoveNext
Wend
End Sub

Private Sub SendData(Command As String)
On Error GoTo Err1:
If MSComm_Com.PortOpen = True Then
    MSComm_Com.PortOpen = False
End If
MSComm_Com.Settings = Trim(TxtCommSetting.Text)
MSComm_Com.CommPort = Trim(TxtCompPort.Text)
MSComm_Com.PortOpen = True
If Command <> "" Then
    MSComm_Com.Output = Trim(UCase(Command))
End If
Exit Sub
MSComm_Com.PortOpen = False
Err1:
    MsgBox Err.Description
End Sub

