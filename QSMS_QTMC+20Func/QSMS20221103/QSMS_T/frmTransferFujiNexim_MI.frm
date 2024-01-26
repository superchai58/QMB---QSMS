VERSION 5.00
Begin VB.Form frmTransferFujiNexim_MI 
   Caption         =   "TransferFujiNexim_MI[20181009]"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtType 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1485
      TabIndex        =   13
      Top             =   2160
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
      Left            =   1485
      TabIndex        =   6
      Top             =   480
      Width           =   1000
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
      Left            =   1485
      TabIndex        =   5
      Top             =   1440
      Width           =   1000
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
      Left            =   4080
      TabIndex        =   4
      Top             =   480
      Width           =   2500
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
      Left            =   4080
      TabIndex        =   3
      Top             =   1440
      Width           =   2500
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
      Left            =   8175
      TabIndex        =   2
      Top             =   480
      Width           =   1000
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
      Left            =   8175
      TabIndex        =   1
      Text            =   "Q"
      Top             =   1440
      Width           =   1000
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
      Left            =   7320
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label LabType 
      BackColor       =   &H0080FF80&
      Caption         =   "Type"
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
      Index           =   1
      Left            =   180
      TabIndex        =   14
      Top             =   2160
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
      Left            =   180
      TabIndex        =   12
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
      Left            =   180
      TabIndex        =   11
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
      Left            =   2820
      TabIndex        =   10
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
      Left            =   2820
      TabIndex        =   9
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
      Left            =   6900
      TabIndex        =   8
      Top             =   480
      Width           =   1335
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
      Left            =   6900
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "frmTransferFujiNexim_MI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''1271
Private Sub Form_Load()
'CmbServer.AddItem "ServerOld"
'CmbServer.AddItem "ServerNew"
'CmbServer.AddItem "ServerR12Test"
End Sub

Private Sub cmdLoad_Click()
    If Trim(txtSide) <> "Q" Then ''20210520 BangLi RQ21051001
        If MsgBox("This Program is TransferFujiNexim_MI ,Side should be [Q],Are you sure to execute ?", vbOKCancel, "Tip") <> vbOK Then
            Exit Sub
        End If
    End If
   If Trim(txtFactory) = "" Or Trim(txtLine) = "" Or Trim(txtJobGr) = "" Or Trim(txtRev) = "" Or Trim(txtBuidT) = "" Or Trim(txtSide) = "" Or Trim(txtType) = "" Then
        MsgBox "Factory,Line,JobGroup,Version,BuildType,Side,都不可以为空，请确认！！"
        Exit Sub
    End If
   cmdLoad.Enabled = False
   labMsg = "Uploading...,please wait a moment,thanks"
   sSql = "exec  QSMS_InsertMEBom_Nexim_MI  '" & Trim(txtFactory) & "', '" & (txtLine) & "','" & (txtJobGr) & "','" & (txtRev) & "','" & (txtBuidT) & "','" & (txtSide) & "','" & (txtType) & "'," & sq(g_userName) & ""
   Set RS = Conn.Execute(sSql)
   If RS.EOF = False Then
           MsgBox Trim(RS!Description)
       End If
    labMsg = ""
   cmdLoad.Enabled = True
   

End Sub

''1271
