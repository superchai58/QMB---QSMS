VERSION 5.00
Begin VB.Form frmPrinterSetting 
   Caption         =   "PrinterSetting"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   7650
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame Fradpm 
      Caption         =   "dpm"
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
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   7215
      Begin VB.OptionButton Opt200 
         Caption         =   "200"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Opt300 
         Caption         =   "300"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtComm4 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5980
      TabIndex        =   15
      Text            =   "1"
      Top             =   4050
      Width           =   800
   End
   Begin VB.TextBox TxtComm3 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4930
      TabIndex        =   14
      Text            =   "8"
      Top             =   4050
      Width           =   800
   End
   Begin VB.TextBox TxtComm2 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3880
      TabIndex        =   13
      Text            =   "N"
      Top             =   4050
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4000
      TabIndex        =   12
      Top             =   4875
      Width           =   1000
   End
   Begin VB.Frame FraPort 
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   7200
      Begin VB.OptionButton OptNetwork 
         Caption         =   "Network"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   4450
         TabIndex        =   11
         Top             =   390
         Width           =   1600
      End
      Begin VB.OptionButton OptLPT 
         Caption         =   "LPT Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2520
         TabIndex        =   10
         Top             =   360
         Width           =   1600
      End
      Begin VB.OptionButton OptComp 
         Caption         =   "Comp Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   630
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1600
      End
   End
   Begin VB.Frame FraPrinter 
      Caption         =   "Printer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7200
      Begin VB.OptionButton OptSATO 
         Caption         =   "SATO printer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2550
         TabIndex        =   7
         Top             =   420
         Width           =   1600
      End
      Begin VB.OptionButton OptZebra 
         Caption         =   "Zebra printer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   630
         TabIndex        =   6
         Top             =   390
         Value           =   -1  'True
         Width           =   1600
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2000
      TabIndex        =   4
      Top             =   4875
      Width           =   1000
   End
   Begin VB.TextBox TxtComm1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2130
      TabIndex        =   3
      Text            =   "9600"
      Top             =   4050
      Width           =   1500
   End
   Begin VB.TextBox TxtCompPort 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2150
      TabIndex        =   1
      Text            =   "1"
      Top             =   3345
      Width           =   2000
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   5700
      X2              =   5980
      Y1              =   4275
      Y2              =   4275
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   4650
      X2              =   4930
      Y1              =   4275
      Y2              =   4275
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   3600
      X2              =   3880
      Y1              =   4275
      Y2              =   4275
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000000&
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   345
      TabIndex        =   2
      Top             =   4050
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000000&
      Caption         =   "CompPort"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   345
      TabIndex        =   0
      Top             =   3345
      Width           =   1695
   End
End
Attribute VB_Name = "frmPrinterSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'20101115 Maggie Save Printer setting in local Registry (1019)
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim aryComm() As String
Dim strComm As String
Dim strCompPort As String
     
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    strComm = GetSetting("SMT", "QSMS", "Comm", "9600,N,8,1")
    aryComm = Split(strComm, ",")
    
    If strComm <> "" Then
        TxtComm1.Text = aryComm(0)
        TxtComm2.Text = aryComm(1)
        TxtComm3.Text = aryComm(2)
        TxtComm4.Text = aryComm(3)
    End If
     
    strCompPort = GetSetting("SMT", "QSMS", "CommPort", "1")
    
    If strCompPort <> "" Then
        TxtCompPort.Text = strCompPort
    End If
         
End Sub


Private Sub cmdSave_Click()
Dim Printer As String
Dim Port As String
Dim Comm As String
Dim dpm As String
    Comm = TxtComm1.Text & "," & TxtComm2.Text & "," & TxtComm3.Text & "," & TxtComm4.Text
              
    If SaveCheck = False Then Exit Sub
    
    If OptZebra.Value = True Then
        Printer = "Zebra"
    Else
        Printer = "SATO"
    End If
    
    If OptComp.Value = True Then
        Port = "COM"
    ElseIf OptLPT.Value = True Then
        Port = "LPT"
    Else
        Port = "Network"
    End If
    ''(1080)
    If Opt300.Value = True Then
        dpm = "300"
    Else
        dpm = "200"
    End If
            
    PrinterType = Printer ''1044
   PrintDpm = dpm ''1044
   
    SaveSetting "SMT", "QSMS", "Printer", Printer
    SaveSetting "SMT", "QSMS", "Port", Port
    SaveSetting "SMT", "QSMS", "CommPort", TxtCompPort
    SaveSetting "SMT", "QSMS", "Comm", Comm
    SaveSetting "SMT", "QSMS", "DPM", dpm   ''(1080)
    
    MsgBox "Save OK!!", vbInformation
       
End Sub


Private Function SaveCheck() As Boolean
    SaveCheck = True
    
    If TxtCompPort = "" Then
       MsgBox "Please Input CompPort!!", vbExclamation, "Prompt"
       SaveCheck = False
    End If
    
    If Trim(TxtComm1) = "" Or Trim(TxtComm2) = "" Or Trim(TxtComm3) = "" Or Trim(TxtComm4) = "" Then
        MsgBox "Please Input Settings!!", vbExclamation, "Prompt"
        SaveCheck = False
    End If
    
End Function
