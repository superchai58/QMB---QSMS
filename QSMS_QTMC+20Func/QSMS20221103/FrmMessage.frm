VERSION 5.00
Begin VB.Form FrmMessage 
   Caption         =   "Message"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   20025
   End
End
Attribute VB_Name = "FrmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  lblmsg = ""
  Unload Me
End Sub
