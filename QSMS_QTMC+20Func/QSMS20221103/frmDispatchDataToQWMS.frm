VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDispatchDataToQWMS 
   Caption         =   "DispatchDataToQWMS"
   ClientHeight    =   9120
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   13530
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUID 
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
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton cmdFix 
      Caption         =   "Fix"
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
      Left            =   9600
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
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
      Left            =   11160
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGridSN 
      Height          =   7215
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   12726
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin MSDataGridLib.DataGrid DataGDetail 
      Height          =   7455
      Left            =   4320
      TabIndex        =   5
      Top             =   1200
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   13150
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.Label lblRow 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Not Fix Dispatch Data:"
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
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dispatch UserID"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmDispatchDataToQWMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub CmdExcel_Click()
On Error GoTo ErrH
    If Not rs.EOF Then
        Call CopyToExcel(rs)
    Else
        MsgBox ("No Data"), vbCritical
    End If
ErrH:
    If Err.Number <> 0 Then
       MsgBox Err.Description
    End If
End Sub

Private Sub cmdFix_Click()
'If UCase(Trim(g_userName)) <> UCase(Trim(lstVewUID.SelectedItem.Text)) Then
'    MsgBox "当前登陆系统的工号:" & g_userName & ",所选工号:" & Trim(lstVewUID.SelectedItem.Text) & ",两者不匹配.只能Fix自己的发料记录!"
'Else
On Error GoTo ErrH

If Trim(txtUID) = "" Then
    MsgBox "Please fill in the UID of the person who dispatch the materials"
    Exit Sub
End If

If MsgBox("Are you sure you want to send the data of the materials?", vbOKCancel, "Message") = vbOK Then
    strSQL = "EXEC QWMS_GetNeedData 'SendData','" & Trim(txtUID.Text) & "','" & Trim(Factory) & "'"
    Conn.Execute (strSQL)
    MsgBox "Dispatch data is fix OK!"
    Call GetUID
End If
ErrH:
    If Err.Number <> 0 Then
       MsgBox Err.Description
    End If
End Sub

Private Sub DataGridSN_Click()
On Error GoTo ErrH
    If Trim(DataGridSN.Columns(0)) <> "" Then
        txtUID.Text = Trim(DataGridSN.Columns("UID").Value)
        strSQL = "EXEC QWMS_GetNeedData 'Detail','" & Trim(txtUID.Text) & "','" & Trim(Factory) & "'"
        
        Set rs = Conn.Execute(strSQL)
        lblRow.Caption = "Not Fix Dispatch Data:" & Trim(rs.RecordCount)
        Set DataGDetail.DataSource = rs
    End If
ErrH:
    If Err.Number <> 0 Then
       MsgBox Err.Description
    End If
End Sub

Private Sub Form_Load()
    Call GetUID
End Sub

Private Sub GetUID()
    strSQL = "EXEC QWMS_GetNeedData 'UID'"
    Set rs = Conn.Execute(strSQL)
    Set DataGridSN.DataSource = rs
End Sub
