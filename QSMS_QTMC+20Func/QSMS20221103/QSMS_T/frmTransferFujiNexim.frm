VERSION 5.00
Begin VB.Form frmTransferFujiNexim 
   Caption         =   "TransferFujiNexim[20160920]"
   ClientHeight    =   2775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
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
      Left            =   7440
      TabIndex        =   12
      Top             =   2040
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
      Top             =   2160
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
        MsgBox "Factory,Line,JobGroup,Version,BuildType,Side,都不可以为空，请确认！！"
        Exit Sub
    End If
   cmdLoad.Enabled = False
   labMsg = "Uploading...,please wait a moment,thanks"
   sSql = "exec  QSMS_InsertMEBom_Nexim  '" & Trim(txtFactory) & "', '" & (txtLine) & "','" & (txtJobGr) & "','" & (txtRev) & "','" & (txtBuidT) & "','" & (txtSide) & "'," & sq(g_userName) & ""
   If chk_MutiModel.Value = 1 Then
    sSql = "exec  QSMS_InsertMEBom_Nexim_191207  '" & Trim(txtFactory) & "', '" & (txtLine) & "','" & (txtJobGr) & "','" & (txtRev) & "','" & (txtBuidT) & "','" & (txtSide) & "'," & sq(g_userName) & ""

   End If
   
   
   
   Set RS = Conn.Execute(sSql)
   If RS.EOF = False Then
           MsgBox Trim(RS!Description)
       End If
    labMsg = ""
   cmdLoad.Enabled = True
   

End Sub
