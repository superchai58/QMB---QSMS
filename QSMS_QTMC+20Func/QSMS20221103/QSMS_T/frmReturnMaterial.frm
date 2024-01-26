VERSION 5.00
Begin VB.Form frmReturnMaterial 
   Caption         =   "Return Material (20071217)"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opBad 
      Caption         =   "Bad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton opGood 
      Caption         =   "Good"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtWO 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   555
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "WO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   570
   End
End
Attribute VB_Name = "frmReturnMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
