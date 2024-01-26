VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCompSelect 
   Caption         =   "frmCompSelect"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   5619.256
   ScaleMode       =   0  'User
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dgNewCompPNList 
      Height          =   4935
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8705
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   480
      Left            =   3000
      TabIndex        =   2
      Top             =   5880
      Width           =   2500
   End
   Begin VB.TextBox txtNewCompPN 
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label lblCompPN 
      Caption         =   "NewCompPN:"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmCompSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public NewCompPNList As ADODB.Recordset
Public NewCompPN As String

Private Sub btnOK_Click()
    Unload Me
End Sub

Private Sub dgNewCompPNList_Click()
    txtNewCompPN.text = dgNewCompPNList
    NewCompPN = dgNewCompPNList
End Sub

Private Sub Form_Load()
    Set dgNewCompPNList.DataSource = NewCompPNList.DataSource
End Sub
