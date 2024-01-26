VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmSpecialReturnArchive 
   BackColor       =   &H8000000B&
   Caption         =   "SpecialReturnArchive"
   ClientHeight    =   9885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDateTime 
      Height          =   375
      Left            =   7440
      TabIndex        =   17
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtDeleteSN 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton CMDDel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtReason 
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   1560
      Width           =   4815
   End
   Begin VB.TextBox txtEntityQty 
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox txtVerifyUID 
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      Top             =   840
      Width           =   2775
   End
   Begin MCI.MMControl MM1 
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   9240
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton CMDSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtReturnUID 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtReturnDID 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid DataGridResult 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   9763
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
            LCID            =   2052
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
            LCID            =   2052
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
   Begin VB.Label DateTime 
      BackColor       =   &H000080FF&
      Caption         =   "DateTime"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label DeleteSN 
      BackColor       =   &H000080FF&
      Caption         =   "DeleteSN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Reason 
      BackColor       =   &H0000FF00&
      Caption         =   "Reason"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label EntityQty 
      BackColor       =   &H0000FF00&
      Caption         =   "EntityQty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   10
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label VerifyUID 
      BackColor       =   &H0000FF00&
      Caption         =   "VerifyUID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label ReturnUID 
      BackColor       =   &H0000FF00&
      Caption         =   "ReturnUID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label ReturnDID 
      BackColor       =   &H0000FF00&
      Caption         =   "ReturnDID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmSpecialReturnArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim str As String
Dim strGroupID As String
Dim RS As ADODB.Recordset
Private Sub Err_Sound()
    MM1.FileName = App.Path & "\OO.wav"
    MM1.Command = "open"
    MM1.Command = "play"
     Do While MM1.Mode = mciModePlay
     Loop
     MM1.Command = "close"
End Sub

Private Sub OK_Sound()
    MM1.FileName = App.Path & "\OK.wav"
    MM1.Command = "open"
    MM1.Command = "play"
    Do While MM1.Mode = mciModePlay
    Loop
    MM1.Command = "close"
End Sub

Private Sub cmdSave_Click()
    If Trim(txtReturnDID) <> "" And Trim(txtEntityQty) <> "" And Trim(txtReturnUID) <> "" And Trim(txtVerifyUID) <> "" And Trim(txtReason) <> "" Then
    
        str = "select * from QSMS_Dispatch with(nolock) where DID='" & Trim(txtReturnDID) & "' "
        Set RS = Conn.Execute(str)
        If RS.EOF Then
            Call Err_Sound
            MsgBox "DID不存在于发料表中！", vbOKOnly Or vbInformation, "系统提示"
            Exit Sub
        End If

        str = "select RemainQty from QSMS_DID with(nolock) where DID='" & Trim(txtReturnDID) & "' union all select RemainQty from QSMS_DID_Log with(nolock) where DID='" & Trim(txtReturnDID) & "' "
        Set RS = Conn.Execute(str)
        If Not RS.EOF Then
            While Not RS.EOF
                If Trim(RS.Fields("RemainQty")) > 0 Then
                    Call Err_Sound
                    MsgBox "DID剩余数量不得大于0！", vbOKOnly Or vbInformation, "系统提示"
                    Exit Sub
                End If

                RS.MoveNext
            Wend
        Else
            Call Err_Sound
            MsgBox "无法查询到DID剩余数量信息！", vbOKOnly Or vbInformation, "系统提示"
            Exit Sub
        End If
        
        If Not IsNumeric(txtEntityQty) Then
            Call Err_Sound
            MsgBox ("实物数量(EntityQty)必须为整数!")
            Exit Sub
        End If

        If Val(txtEntityQty) <= 0 Then
            Call Err_Sound
            MsgBox ("实物数量(EntityQty)不可为0!")
            Exit Sub
        End If
        
        If Len(txtReturnUID) < 7 Or Len(txtReturnUID) > 8 Then
            Call Err_Sound
            MsgBox ("输入工号(ReturnUID)有误(不少于7码，不得超过8码)!请重新输入!")
            Exit Sub
        End If
        
        If Len(txtVerifyUID) < 7 Or Len(txtVerifyUID) > 8 Then
            Call Err_Sound
            MsgBox ("输入工号(VerifyUID)有误(不少于7码，不得超过8码)!请重新输入!")
            Exit Sub
        End If
        
        If Len(txtReason) < 10 Then
            Call Err_Sound
            MsgBox ("记录确认具体内容(Reason)，不得少于10个字符!")
            Exit Sub
        End If
        
        str = "insert into SpecialReturnArchive_LOG values('SpecialReturnArchive','1','" & Trim(txtReturnDID) & "','" & Trim(g_userName) & "',N'" & Trim(txtEntityQty) & ";" & Trim(txtReturnUID) & ";" & Trim(txtVerifyUID) & ";" & Trim(txtReason) & "',convert(char(8),getdate(),112) + left(replace(convert(char(8),getdate(),108), ':', ''),6)" & ")"
        Conn.Execute (str)
        Call OK_Sound
        txtReturnDID.Text = ""
        txtEntityQty.Text = ""
        txtReturnUID.Text = ""
        txtVerifyUID.Text = ""
        txtReason.Text = ""
    
        Call reFreshData
        Set RS = Nothing
    Else
        Call Err_Sound
        MsgBox "各输入值均不得为空！", vbOKOnly Or vbInformation, "系统提示"
        'txtReturnDID = ""
        'txtReturnDID.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Call reFreshData
End Sub

Private Sub reFreshData()
    Dim strl As String
    Dim rstl As Recordset
    
    strl = "select top 100 * from SpecialReturnArchive_LOG with(nolock) order by Trans_Date desc"
    Set rstl = Conn.Execute(strl)
    'If Not rstl.EOF Then
    Set DataGridResult.DataSource = rstl
    'End If
    Set rstl = Nothing
End Sub

Private Sub txtReturnDID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtReturnDID) <> "" Then
        
        txtEntityQty.SetFocus
        
        'str = "select * from QSMS_DID with(nolock) where DID='" & Trim(txtReturnDID) & "' "
        'Set RS = Conn.Execute(str)
        'If Not RS.EOF Then
        '    If Trim(RS.Fields("RemainQty")) > 0 Then
        '        Call Err_Sound
        '        MsgBox "DID剩余数量不得大于0！", vbOKOnly Or vbInformation, "系统提示"
        '        txtReturnDID = ""
        '        txtReturnDID.SetFocus
        '        Exit Sub
        '    End If
        'End If
        
        'str = "select * from QSMS_DID_Log with(nolock) where DID='" & Trim(txtReturnDID) & "' "
        'Set RS = Conn.Execute(str)
        'If Not RS.EOF Then
        '    If Trim(RS.Fields("RemainQty")) > 0 Then
        '        Call Err_Sound
        '        MsgBox "DID剩余数量不得大于0！", vbOKOnly Or vbInformation, "系统提示"
        '        txtReturnDID = ""
        '        txtReturnDID.SetFocus
        '        Exit Sub
        '    End If
       ' End If
        
    End If
End Sub

'Private Sub txtReturnDID_LostFocus()
'    str = "select * from QSMS_Dispatch with(nolock) where DID='" & Trim(txtReturnDID) & "' "
'    Set RS = Conn.Execute(str)
'    If RS.EOF Then
'        Call Err_Sound
'        MsgBox "DID不存在于发料表中！", vbOKOnly Or vbInformation, "系统提示"
'        txtReturnDID = ""
'        'txtReturnDID.SetFocus
'        Exit Sub
'    End If
'
'    str = "select RemainQty from QSMS_DID with(nolock) where DID='" & Trim(txtReturnDID) & "' union all select RemainQty from QSMS_DID_Log with(nolock) where DID='" & Trim(txtReturnDID) & "' "
'    Set RS = Conn.Execute(str)
'    If Not RS.EOF Then
'        While Not RS.EOF
'            If Trim(RS.Fields("RemainQty")) > 0 Then
'                Call Err_Sound
'                MsgBox "DID剩余数量不得大于0！", vbOKOnly Or vbInformation, "系统提示"
'                txtReturnDID = ""
'                'txtReturnDID.SetFocus
'                Exit Sub
'            End If
'
'            RS.MoveNext
'        Wend
'    Else
'        Call Err_Sound
'        MsgBox "无法查询到DID剩余数量信息！", vbOKOnly Or vbInformation, "系统提示"
'        txtReturnDID = ""
'        'txtReturnDID.SetFocus
'        Exit Sub
'    End If
'
'    'txtEntityQty.SetFocus
'
'    Set RS = Nothing
'End Sub

Private Sub txtEntityQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtEntityQty) <> "" Then
        txtReturnUID.SetFocus
    End If
End Sub

'Private Sub txtEntityQty_LostFocus()
'    If Not IsNumeric(txtEntityQty) Then
'        Call Err_Sound
'        MsgBox ("实物数量必须为整数!")
'        txtEntityQty = ""
'        'txtEntityQty.SetFocus
'        Exit Sub
'    End If
'
'    If Val(txtEntityQty) <= 0 Then
'        Call Err_Sound
'        MsgBox ("实物数量不可为0!")
'        txtEntityQty = ""
'        'txtEntityQty.SetFocus
'        Exit Sub
'    End If
'
'    'txtReturnUID.SetFocus
'End Sub

Private Sub txtReturnUID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtReturnUID) <> "" Then
        txtVerifyUID.SetFocus
    End If
End Sub

'Private Sub txtReturnUID_LostFocus()
'    If Len(txtReturnUID) < 7 Or Len(txtReturnUID) > 8 Then
'        Call Err_Sound
'        MsgBox ("输入工号有误!请重新输入!")
'        txtReturnUID = ""
'        'txtReturnUID.SetFocus
'        Exit Sub
'    End If
'
'    'txtVerifyUID.SetFocus
'End Sub

Private Sub txtVerifyUID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtVerifyUID) <> "" Then
        txtReason.SetFocus
    End If
End Sub

'Private Sub txtVerifyUID_LostFocus()
'    If Len(txtVerifyUID) < 7 Or Len(txtVerifyUID) > 8 Then
'        Call Err_Sound
'        MsgBox ("输入工号有误!请重新输入!")
'        txtVerifyUID = ""
'        'txtVerifyUID.SetFocus
'        Exit Sub
'    End If
'
'    'txtReason.SetFocus
'End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtReason) <> "" Then
        CMDSave.SetFocus
    End If
End Sub

'Private Sub txtReason_LostFocus()
'    If Len(txtReason) < 10 Then
'        Call Err_Sound
'        MsgBox ("记录确认具体内容，不得少于10个字符!")
'        txtReason = ""
'        'txtReason.SetFocus
'        Exit Sub
'    End If
'
'    'CMDSave.SetFocus
'End Sub

Private Sub txtDeleteSN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtDeleteSN) <> "" Then
        txtDateTime.SetFocus
    End If
End Sub

Private Sub txtDateTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtDateTime) <> "" Then
        CMDDel.SetFocus
    End If
End Sub

Private Sub cmdDel_Click()
    Dim strsql As String
    Dim Rst As ADODB.Recordset
    
    strsql = "select * from SpecialReturnArchive_LOG where SN='" & Trim(txtDeleteSN) & "' and Trans_Date='" & Trim(txtDateTime) & "' "
    Set Rst = Conn.Execute(strsql)
    
    'If Trim(txtDeleteSN) <> "" And Trim(txtDateTime) <> "" Then
    If Not Rst.EOF Then
        strsql = "delete from SpecialReturnArchive_LOG where SN='" & Trim(txtDeleteSN) & "' and Trans_Date='" & Trim(txtDateTime) & "' "
        Conn.Execute (strsql)
        Call OK_Sound
    
        Call reFreshData
    Else
        Call Err_Sound
        MsgBox ("没有对应数据!")
    End If
End Sub
