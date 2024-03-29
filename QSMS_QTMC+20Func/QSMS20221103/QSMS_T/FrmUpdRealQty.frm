VERSION 5.00
Begin VB.Form FrmUpdRealQty 
   Caption         =   "Update DID RealQty[2012/06/08]"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtReason 
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
      Left            =   4320
      TabIndex        =   16
      Top             =   1200
      Width           =   5175
   End
   Begin VB.TextBox txtTotalQty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      Picture         =   "FrmUpdRealQty.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "说明"
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   5415
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "3.在Reason后面输入原因，点Save保存。"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "2.在Update To 后面输入你要改的DID数量。"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "1.刷入你想要更改数量的DID号码, 会带出它当前的Real Qty。"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      Picture         =   "FrmUpdRealQty.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtUpdQty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Txtqty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox TxtDID 
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
      Left            =   1440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "Reason"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "Real Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   7320
      TabIndex        =   13
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblmsg 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   11
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "Update To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "Total Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "DID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "FrmUpdRealQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************************************************************
'**文 件 名: FrmUpdRealQty.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: Lynn Sun
'**日    期: 2008.08.12
'**描    述: DID Header
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**Sandy       20080919        mark realqty must large then 200 --------(0001)
'**Sandy       20091003        将转txtTotalQty，txtUpdQty换为int类型---（0002）
'***********************************************************************************/
Dim Line1 As String
Dim Side As String

Private Sub cmdReset_Click()
TxtDID = ""
Txtqty = ""
txtTotalQty = ""
txtUpdQty = ""
TxtDID.Locked = False
lblmsg.Caption = ""
txtReason.Text = ""
End Sub

Private Sub CmdSave_Click()
Dim str As String
Dim str1 As String
Dim rs As adodb.Recordset
If TxtDID <> "" And Txtqty <> "" And txtUpdQty <> "" And IsNumeric(Trim(txtUpdQty)) = True And txtReason <> "" Then
'    If txtTotalQty < txtUpdQty Or Txtqty > 200 Then
'**Sandy       20080919     mark realqty must large then 200 --------(0001)
   '''' If CInt(txtTotalQty) < CInt(txtUpdQty) Then
        If CLng(txtTotalQty) < CLng(txtUpdQty) Then         ''Fix Bug By Newton Qty较大时超出Cint转换范围
        MsgBox "Update Qty can not larger than Total Qty!"
        Exit Sub
    End If
    
    str = "update qsms_did set realqty=" & Trim(txtUpdQty) & " where did='" & Trim(TxtDID) & "'"
    Conn.Execute (str)
    '1098
    str1 = "Line=" & Line1 & ";Side=" & Side & ";TotalQty=" & Trim(txtTotalQty) & ";RealQTY=" & Trim(Txtqty) & ";UpdateTo=" & Trim(txtUpdQty) & ";"
    ''Add log in qsms_log
    str = "insert into qsms_log(system_name,event_no,did,[user_name],returnqty,trans_date) values ('SMT_QSMS','" & str1 & "','" & Trim(TxtDID) & "',N'" & g_userName & ";Reason=" & Trim(txtReason) & "',0,dbo.formatdate(getdate(),'yyyymmddhhnnss'))"
    Conn.Execute (str)
    Call cmdReset_Click
    lblmsg.Caption = "OK"
    TxtDID.SetFocus
End If
End Sub

Private Sub Form_Load()
    ''20100903 Kyle added to solve the encoding problem of UI.
    If StrBU = "PO" Then
        Frame1.Caption = "弧"
        Label4(2).Caption = "1.稱璶э计秖DID腹絏穦盿ウ讽玡Real Qty"
        Label4(3).Caption = "2.Update to块璶э眔DID计秖翴Save玂"
    End If
End Sub

Private Sub txtDID_KeyPress(KeyAscii As Integer)
Dim str As String
Dim rs As adodb.Recordset
If TxtDID <> "" And KeyAscii = 13 Then
    sql = "select did,qty,realqty,Line,Side from qsms_did where did='" & Trim(TxtDID) & "'"
    Set rs = Conn.Execute(sql)
    If rs.EOF = True Then
        MsgBox "Can not find this DID, check it please !!"
        Exit Sub
    Else
        txtTotalQty = rs!Qty
        Txtqty = rs!realqty
        TxtDID.Locked = True
        Line1 = rs!Line             '1098
        Side = rs!Side
        txtUpdQty.SetFocus
    End If
End If
End Sub
Private Sub txtReason_KeyPress(KeyAscii As Integer)
If txtReason <> "" And KeyAscii = 13 Then
    Call CmdSave_Click
End If
End Sub


Private Sub txtUpdQty_KeyPress(KeyAscii As Integer)
If txtUpdQty <> "" And KeyAscii = 13 Then
    txtReason.SetFocus
End If
End Sub
