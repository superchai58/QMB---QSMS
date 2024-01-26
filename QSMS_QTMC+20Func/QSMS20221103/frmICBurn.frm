VERSION 5.00
Begin VB.Form frmICBurn 
   Caption         =   "IC Burn 2013/02/01"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3165
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   7515
      Begin VB.CommandButton cmdLinkShearPin 
         Caption         =   "DIDLinkShearPin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   12
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton CmdShearPinExcel 
         Caption         =   "QueryCC_ShearPin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   11
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Query IC_Comp"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   10
         Top             =   1680
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.ComboBox cboPN 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   840
         TabIndex        =   9
         Text            =   "cboPN"
         Top             =   960
         Width           =   2865
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   8
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtModelName 
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   1995
      End
      Begin VB.TextBox txtCompPN 
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   1995
      End
      Begin VB.TextBox txtDID 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   810
         TabIndex        =   2
         Top             =   270
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "ModelName:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   6
         Top             =   990
         Width           =   1545
      End
      Begin VB.Label Label3 
         Caption         =   "PN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Comp PN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3840
         TabIndex        =   3
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "DID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   330
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmICBurn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboPN_Click()
    Dim sql As String
    Dim rs As ADODB.Recordset
    sql = "select * from ModelName where PN='" & cboPN.Text & "'"  '''1104
    Set rs = Conn.Execute(sql)
    If rs.EOF = False Then
        txtModelName = rs("ModelName")
    Else
        txtModelName = ""
    End If
End Sub

Private Sub cmdExcel_Click()
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim TemplateFileName As String
    TemplateFileName = "IC_Burn"
    '20110920 Maggie Ôö¼ÓBy PN/CompPN²éÑ¯
    'sql = "select ModelName, PN, CompPN, Location, UID from IC_CompPN where ModelName=" & sq(Trim(cboModelName))
    sql = "select * from IC_CompPN where PN like '" & Trim(cboPN.Text) & "%' and CompPN like '" & Trim(txtCompPN) & "%'"
    Set rs = Conn.Execute(sql)
    If Not rs.EOF Then
       Call CopyToTemplateExcel(rs, TemplateFileName)
    Else
       MsgBox ("No Data"), vbCritical
    End If
End Sub

Private Sub cmdLinkShearPin_Click()
    Dim sql As String
    Dim rs As ADODB.Recordset
    If Trim(txtDID) = "" Then
        MsgBox "Please input DID!", vbInformation
        txtDID.SetFocus
        Exit Sub
    End If
    If Trim(cboPN.Text) = "" Then
        MsgBox "Please select PN !", vbInformation
        cboPN.SetFocus
        Exit Sub
    End If
    sql = "Exec IC_ShearPinLinkDID '" & txtDID.Text & "','" & txtModelName.Text & "','" & cboPN.Text & "','" & txtCompPN.Text & "','" & Trim(g_userName) & "'"
    Set rs = Conn.Execute(sql)
    MsgBox rs("description")
End Sub

Private Sub cmdOK_Click()
    Dim sql As String
    Dim rs As ADODB.Recordset
    If Trim(txtDID) = "" Then
        MsgBox "Please input DID!", vbInformation
        txtDID.SetFocus
        Exit Sub
    End If
    If Trim(cboPN.Text) = "" Then
        MsgBox "Please select PN !", vbInformation
        cboPN.SetFocus
        Exit Sub
    End If
    sql = "Exec IC_CompPNLinkDID '" & txtDID.Text & "','" & txtModelName.Text & "','" & cboPN.Text & "','" & txtCompPN.Text & "','" & Trim(g_userName) & "'"
    Set rs = Conn.Execute(sql)
    MsgBox rs("description")
End Sub

Private Sub CmdShearPinExcel_Click()

    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim TemplateFileName As String
    TemplateFileName = "IC_ShearPin"
  
    sql = "select * from IC_ShearPin where PN like '" & Trim(cboPN.Text) & "%' and ModelName like '" & Trim(txtCompPN) & "%'"
    Set rs = Conn.Execute(sql)
    If Not rs.EOF Then
       Call CopyToTemplateExcel(rs, TemplateFileName)
    Else
       MsgBox ("No Data"), vbCritical
    End If
End Sub


Private Sub Form_Load()
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    '''1157
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    sql = "select * from ModelName"
    Set rs = Conn.Execute(sql)
    For i = 1 To rs.RecordCount
        cboPN.AddItem rs("PN")
        rs.MoveNext
    Next i
    
End Sub

Private Sub txtDID_KeyPress(KeyAscii As Integer)
    Dim sql As String
    Dim rs As ADODB.Recordset
    If KeyAscii = 13 Then
        sql = "select * from QSMS_DID where DID='" & Trim(txtDID.Text) & "'"
        Set rs = Conn.Execute(sql)
        If rs.EOF Then
            MsgBox "This DID is not in QSMS !"
            txtCompPN = ""
            txtDID = ""
            txtDID.SetFocus
        Else
            txtCompPN = rs("CompPN")
            cboPN.SetFocus
        End If
    End If
End Sub

Private Sub txtDID_LostFocus()
    Dim sql As String
    Dim rs As ADODB.Recordset
    If Trim(txtDID) <> "" Then
        sql = "select * from QSMS_DID where DID='" & Trim(txtDID.Text) & "'"
        Set rs = Conn.Execute(sql)
        If rs.EOF Then
            MsgBox "This DID is not in QSMS !"
            txtDID = ""
            txtDID.SetFocus
        Else
            txtCompPN = rs("CompPN")
        End If
    End If
End Sub
