VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmDIDBake 
   Caption         =   "DIDBake"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBakeOK 
      BackColor       =   &H8000000D&
      Caption         =   "StartBake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdBakeQ 
      BackColor       =   &H8000000D&
      Caption         =   "Query"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGridDIDBake 
      Height          =   3135
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5530
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
   Begin VB.CommandButton cmdEndBake 
      BackColor       =   &H8000000D&
      Caption         =   "EndBake"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtBakeDID 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "DID"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FrmDIDBake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim sql As String
   Dim Rst As ADODB.Recordset
Private Sub cmdBakeOK_Click()
   If Trim(txtBakeDID) <> "" Then
        sql = "Exec QSMS_DIDBake @DID='" & Trim(txtBakeDID.Text) & "',@UID='" & Trim(g_userName) & "',@Type='Bake'"
        Set Rst = Conn.Execute(sql)
        If Not Rst.EOF Then
           If Trim(Rst!result) <> "OK" Then
              MsgBox Rst!Desc
              txtBakeDID.Text = ""
              Exit Sub
           End If
         Call reFreshBakeData
         txtBakeDID.Text = ""
        End If
   End If
End Sub
Private Function reFreshBakeData()
   sql = "Exec QSMS_DIDBake @DID='" & Trim(txtBakeDID.Text) & "',@UID='" & Trim(g_userName) & "',@Type='Query'"
   Set Rst = Conn.Execute(sql)
   If Not Rst.EOF Then
      Set DataGridDIDBake.DataSource = Rst
   End If
End Function

Private Sub cmdBakeQ_Click()
  Call reFreshBakeData
End Sub
Private Sub cmdEndBake_Click()
   If Trim(txtBakeDID) <> "" Then
        sql = "Exec QSMS_DIDBake @DID='" & Trim(txtBakeDID.Text) & "',@UID='" & Trim(g_userName) & "',@Type='EndBake'"
        Set Rst = Conn.Execute(sql)
        If Not Rst.EOF Then
           If Trim(Rst!result) <> "OK" Then
              MsgBox Rst!Desc
              txtBakeDID.Text = ""
              Exit Sub
           End If
         Call reFreshBakeData
         txtBakeDID.Text = ""
        End If
    End If
End Sub
