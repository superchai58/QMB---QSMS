VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmPEMainTain_WO 
   BackColor       =   &H8000000E&
   Caption         =   "MainTain_WO"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
   Begin VB.Frame Frame1 
      Caption         =   "MainTain_WO"
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8895
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
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
         Left            =   4680
         TabIndex        =   14
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton CmdQuery 
         Caption         =   "Query"
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
         Left            =   2640
         TabIndex        =   13
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton CmdADD 
         Caption         =   "ADD"
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
         Left            =   480
         TabIndex        =   12
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox TxtPN 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   400
         Left            =   5880
         TabIndex        =   10
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox TxtDCode 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   400
         Left            =   6360
         TabIndex        =   9
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox TxtVCode 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   400
         Left            =   1680
         TabIndex        =   8
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox TxtLCode 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   400
         Left            =   1680
         TabIndex        =   7
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txtWO 
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   400
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblSMT 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VendorCode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label lblSMT 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DateCode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   3
         Left            =   4800
         TabIndex        =   4
         Top             =   1920
         Width           =   1545
      End
      Begin VB.Label lblSMT 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LotCode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   1545
      End
      Begin VB.Label lblSMT 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CompPN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   0
         Left            =   4680
         TabIndex        =   2
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label lblSMT 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "WO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Index           =   2
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1185
      End
   End
End
Attribute VB_Name = "FrmPEMainTain_WO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdADD_Click()
Dim strSQL As String
If txtWO <> "" And TxtPN <> "" Then
       strSQL = "EXEC QSMS_MainTain_WO @WO ='" & Trim(txtWO) & "',@CompPN ='" & Trim(TxtPN) & "',@VendorCode ='" & Trim(TxtVCode) & "',@DateCode ='" & Trim(TxtDCode) & "',@LotCode ='" & Trim(TxtLCode) & "',@Type='ADD'"
       Conn.Execute (strSQL)
       Call reFreshData
       MsgBox "OK", vbOKOnly Or vbInformation, "系统提示"
       txtWO = ""
       TxtPN = ""
       TxtVCode = ""
       TxtDCode = ""
       TxtLCode = ""
Else
       MsgBox "添加信息不能有空，请确认！", vbOKOnly Or vbInformation, "系统提示"
       txtWO = ""
       txtWO.SetFocus
       Exit Sub
End If
End Sub
Private Sub reFreshData()
Dim tmpSQL As String
Dim rs As New ADODB.Recordset
    tmpSQL = "Select * from QSMS_PEMainTain_WO where WO ='" & Trim(txtWO) & "' and CompPN ='" & Trim(TxtPN) & "' and VendorCode ='" & Trim(TxtVCode) & "'" & _
                " and DateCode ='" & Trim(TxtDCode) & "' and LotCode ='" & Trim(TxtLCode) & "'"
    Set rs = Conn.Execute(tmpSQL)
    Set DataGrid1.DataSource = rs
End Sub

Private Sub cmdDelete_Click()
Dim strSQL As String
Dim rs As New ADODB.Recordset
If txtWO <> "" And TxtPN <> "" Then
    strSQL = "EXEC QSMS_MainTain_WO @WO ='" & Trim(txtWO) & "',@CompPN ='" & Trim(TxtPN) & "',@VendorCode ='" & Trim(TxtVCode) & "',@DateCode ='" & Trim(TxtDCode) & "',@LotCode ='" & Trim(TxtLCode) & "',@Type='Delete'"
       Conn.Execute (strSQL)
       ''Call reFreshData
       MsgBox "删除成功", vbOKOnly Or vbInformation, "系统提示"
       txtWO = ""
       TxtPN = ""
       TxtVCode = ""
       TxtDCode = ""
       TxtLCode = ""
       
       strSQL = "SELECT * FROM QSMS_PEMainTain_WO"
       Set rs = Conn.Execute(strSQL)
       Set DataGrid1.DataSource = rs
End If
End Sub

Private Sub CmdQuery_Click()
Dim strSQL As String
Dim rs As New ADODB.Recordset

If txtWO <> "" Or TxtPN <> "" Or TxtVCode <> "" Or TxtDCode <> "" Or TxtLCode <> "" Then
    strSQL = "EXEC QSMS_MainTain_WO @WO ='" & Trim(txtWO) & "',@CompPN ='" & Trim(TxtPN) & "',@VendorCode ='" & Trim(TxtVCode) & "',@DateCode ='" & Trim(TxtDCode) & "',@LotCode ='" & Trim(TxtLCode) & "',@Type='Query'"
Else
    MsgBox "输入不能为空", vbOKOnly Or vbInformation, "系统提示"
    txtWO = ""
    txtWO.SetFocus
    Exit Sub
End If
    Set rs = Conn.Execute(strSQL)
    Set DataGrid1.DataSource = rs
End Sub

Private Sub DataGrid1_Click()
   txtWO = DataGrid1.Columns("WO").Value
   TxtPN = DataGrid1.Columns("CompPN").Value
   TxtVCode = DataGrid1.Columns("VendorCode").Value
   TxtDCode = DataGrid1.Columns("DateCode").Value
   TxtLCode = DataGrid1.Columns("LotCode").Value
End Sub
