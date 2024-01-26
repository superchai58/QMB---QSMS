VERSION 5.00
Begin VB.Form frmSelectGroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Plant 20211023"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboGroup 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
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
      Left            =   960
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scope : MaintainDIDAutoDispatch,       Report,DIDchkStock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Select Plant:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmSelectGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
''Public strBU As String

Private Sub cmdCancel_Click()
  Unload Me
  
End Sub

Private Sub CmdOk_Click()
If Me.cboGroup.text <> "" Then
    plant = Me.cboGroup.text
    mdiMain.Show
    Unload Me
End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String
    Dim Rs As New ADODB.Recordset
    
    ConnSMT.CursorLocation = adUseClient
    If ConnSMT.State = 1 Then ConnSMT.Close
    ConnSMT.Open strConnSMT
    strSQL = "Select distinct GroupName from Glue_Stock with(nolock)"
    Set Rs = ConnSMT.Execute(strSQL)
  
    Me.cboGroup.AddItem ("All")
'  Me.cboGroup.AddItem ("TH20")
'  Me.cboGroup.AddItem ("TH2C")
    Me.cboGroup.AddItem ("TB30")
  
'test

'Dim Rs1 As New ADODB.Recordset
'Dim strSQL1 As String
'ConnSMT.CursorLocation = adUseClient
'If ConnSMT.State = 1 Then ConnSMT.Close
'ConnSMT.Open strConnSMT
  
'plant = "all"
'
'        strSQL = "Exec XL_GetAllWOInfoList 'Line','','','','','','CS41212FB11',''"
'        Set Rs = Conn.Execute(strSQL)
'        If Rs.EOF = False Then
'            While Not Rs.EOF
'                ''“¿èSÖ^÷ÆLine ÃÓ»Î
'                strSQL1 = "Exec GetPlant2Line '" & plant & "','" & Trim(Rs!GroupValue) & "'"
'                Set Rs1 = ConnSMT.Execute(strSQL1)
'                If (Trim(Rs1!result) = "1") Then
'                    Me.cboGroup.AddItem (Trim(Rs!GroupValue))
'                End If
'                Rs.MoveNext
'            Wend
'        End If


   

  On Error Resume Next
  
End Sub

