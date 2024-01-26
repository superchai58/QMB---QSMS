VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmUpdateUID 
   Caption         =   "UpdateUID"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNewUID 
      Height          =   405
      Left            =   3960
      TabIndex        =   4
      Text            =   "请输入替换后的工号"
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtOldUID 
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Text            =   "请输入要替换的工号"
      Top             =   240
      Width           =   2775
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Height          =   3255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5741
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
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblNew 
      Caption         =   "新工号"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblOld 
      Caption         =   "旧工号"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "FrmUpdateUID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuery_Click()
Dim sql As String
Dim rs As New Recordset ''''
sql = "select distinct UID  from QSMS_DID_ToWH with(nolock) where UID NOT IN (SELECT Username from userdetail with(nolock))and UID<>''"
Set rs = Conn.Execute(sql)
If rs.EOF = False Then
    Set DataGrid.DataSource = Nothing
    Set DataGrid.DataSource = rs
    DataGrid.Refresh
    
End If
End Sub
Private Sub cmdUpdate_Click()
Dim sql As String
 If Trim(txtNewUID) = "" Then
    MsgBox "请先输入替换后的工号!", vbCritical, "ErrMessage"
    NewUID.SetFocus
    Exit Sub
Else
    sql = "update QSMS_DID_ToWH  set UID='" & Trim(txtNewUID) & "' where UID='" & Trim(txtOldUID) & "'"
    Conn.Execute (sql)
    MsgBox ("Update UID successfully!"), vbInformation
End If
End Sub
