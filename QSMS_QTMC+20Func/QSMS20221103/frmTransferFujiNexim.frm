VERSION 5.00
Begin VB.Form frmTransferFujiNexim 
   Caption         =   "TransferFujiNexim[20160920]"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5216.73
   ScaleMode       =   0  'User
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtWO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1560
      TabIndex        =   19
      Top             =   2280
      Width           =   1845
   End
   Begin VB.CheckBox chk_MutiModel 
      Caption         =   "MutiModel"
      Height          =   375
      Left            =   7080
      TabIndex        =   14
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Upload"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtSide 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   8300
      TabIndex        =   11
      Top             =   1440
      Width           =   1000
   End
   Begin VB.TextBox txtBuidT 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   8300
      TabIndex        =   10
      Top             =   480
      Width           =   1000
   End
   Begin VB.TextBox txtRev 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4200
      TabIndex        =   9
      Top             =   1440
      Width           =   2500
   End
   Begin VB.TextBox txtJobGr 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4200
      TabIndex        =   8
      Top             =   480
      Width           =   2500
   End
   Begin VB.TextBox txtLine 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1600
      TabIndex        =   7
      Top             =   1440
      Width           =   1000
   End
   Begin VB.TextBox txtFactory 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1600
      TabIndex        =   1
      Top             =   480
      Width           =   1000
   End
   Begin VB.Label lblWO 
      BackColor       =   &H0080FF80&
      Caption         =   "WorKOrder"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Build Type = 3; Single Side(C Side)"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   4440
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Build Type = 2; Single Side(S Side)"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Build Type = 1; Double Side(S/C Side)"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label labMsg 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   6975
   End
   Begin VB.Label LabSide 
      BackColor       =   &H0080FF80&
      Caption         =   "Side"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7020
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label labBT 
      BackColor       =   &H0080FF80&
      Caption         =   "BuildType"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7020
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label LabVer 
      BackColor       =   &H0080FF80&
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2940
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label labJobG 
      BackColor       =   &H0080FF80&
      Caption         =   "JobGroup"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2940
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label LabLine 
      BackColor       =   &H0080FF80&
      Caption         =   "Line"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   300
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label LalFactory 
      BackColor       =   &H0080FF80&
      Caption         =   "Factory"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   300
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmTransferFujiNexim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'CmbServer.AddItem "ServerOld"
'CmbServer.AddItem "ServerNew"
'CmbServer.AddItem "ServerR12Test"
End Sub

Private Sub cmdLoad_Click()
        
   If Trim(txtFactory) = "" Or Trim(txtLine) = "" Or Trim(txtJobGr) = "" Or Trim(txtRev) = "" Or Trim(txtBuidT) = "" Or Trim(txtSide) = "" Then
        MsgBox "Factory,Line,JobGroup,Version,BuildType,Side, All fields cannot be empty, please comfirm."
        Exit Sub
    End If
   cmdLoad.Enabled = False
   labMsg = "Uploading...,please wait a moment,thanks"
   sSql = "exec  QSMS_InsertMEBom_Nexim  '" & Trim(txtFactory) & "', '" & (txtLine) & "','" & (txtJobGr) & "','" & (txtRev) & "','" & (txtBuidT) & "','" & (txtSide) & "'," & sq(g_userName) & "," & sq(txtWO) & ""
   If chk_MutiModel.Value = 1 Then
    sSql = "exec  QSMS_InsertMEBom_Nexim_191207  '" & Trim(txtFactory) & "', '" & (txtLine) & "','" & (txtJobGr) & "','" & (txtRev) & "','" & (txtBuidT) & "','" & (txtSide) & "'," & sq(g_userName) & ""

   End If
   
   
   
   Set Rs = Conn.Execute(sSql)
   If Rs.EOF = False Then
           MsgBox Trim(Rs!Description)
       End If
    labMsg = ""
   cmdLoad.Enabled = True
   

End Sub
