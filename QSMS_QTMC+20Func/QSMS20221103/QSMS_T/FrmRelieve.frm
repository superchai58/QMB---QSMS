VERSION 5.00
Begin VB.Form FrmInRelieve 
   Caption         =   "RelieveIPQC"
   ClientHeight    =   4515
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   8357.554
   ScaleMode       =   0  'User
   ScaleWidth      =   12495.39
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRelieve 
      Caption         =   "Relieve"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txtDID 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label LblDID 
      Caption         =   "DID:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "FrmInRelieve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRelieve_Click() ''1258
    If (txtDID.Text = "") Then
      MsgBox ("DID不能为空！")
      Exit Sub
    End If
    If MsgBox("请核对DID信息是否正确？", vbYesNo, "提示信息") = vbYes Then
       ''Conn.Execute ("update QSMS_DID set IPQCFlag='Y' where DID='" & Trim(txtDID) & "'")
       Conn.Execute ("update QSMS_DID set IPQCFlag='Y' where DID='" & Trim(txtDID) & "';update qsms_DID_inspect set IPQCFlag='Y',TestResult='PASS' where DID='" & Trim(txtDID) & "'  ")
       MsgBox ("该DID Check OK!")
    End If
    txtDID.Text = ""
End Sub

''1258
