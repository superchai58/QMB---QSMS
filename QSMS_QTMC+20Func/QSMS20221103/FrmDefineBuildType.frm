VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmDefineBuildType 
   Caption         =   "Define BuildType2015/01/04"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbLine 
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
      Left            =   4800
      TabIndex        =   22
      Top             =   2760
      Width           =   1695
   End
   Begin VB.ComboBox cmbSide 
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
      ItemData        =   "FrmDefineBuildType.frx":0000
      Left            =   1800
      List            =   "FrmDefineBuildType.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Frame FraBuildType 
      BackColor       =   &H80000004&
      Caption         =   "Build Type"
      Height          =   1695
      Left            =   360
      TabIndex        =   15
      Top             =   2400
      Width           =   10455
      Begin VB.Frame Frame1 
         Caption         =   "ËµÃ÷"
         ForeColor       =   &H000000FF&
         Height          =   1335
         Left            =   7440
         TabIndex        =   32
         Top             =   240
         Width           =   2895
         Begin VB.Label Label11 
            Caption         =   "Station£ºPrograms used in production"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   960
            Width           =   2775
         End
         Begin VB.Label Label10 
            Caption         =   "Side: Actual production side"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label9 
            Caption         =   "Line: Actual production line"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label8 
            Caption         =   "Label8"
            Height          =   15
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.ComboBox CboStation 
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
         Left            =   4440
         TabIndex        =   31
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H0000FF00&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox CboBuildType 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FF80&
         Caption         =   "Station"
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
         Left            =   3120
         TabIndex        =   30
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label labelLine 
         BackColor       =   &H0080FF80&
         Caption         =   "Build Line"
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
         Left            =   3120
         TabIndex        =   21
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label labelside 
         BackColor       =   &H0080FF80&
         Caption         =   "Build Side"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "BuildType"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame FraFile 
      BackColor       =   &H80000004&
      Caption         =   "Select Work Order"
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   10455
      Begin VB.TextBox txtBuild 
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
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox TxtWO 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox TxtModel 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   960
         Width           =   2535
      End
      Begin VB.ComboBox CboLine 
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
         Left            =   960
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TxtGroup 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox TxtWOQty 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox TxtWOType 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cboWO 
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
         Left            =   3840
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "A workorder can only define one PCB"
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
         Left            =   6360
         TabIndex        =   26
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000003&
         Caption         =   "BuildType"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   6360
         TabIndex        =   24
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Group(M/S)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   22
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   21
         Left            =   6360
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "WO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   13
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   2
         Left            =   3840
         TabIndex        =   9
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Line"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "WOType"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Work Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid DataGDetail 
      Height          =   3735
      Left            =   360
      TabIndex        =   23
      Top             =   4560
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6588
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
   Begin MSDataGridLib.DataGrid DataWOMulti 
      Height          =   3735
      Left            =   6480
      TabIndex        =   29
      Top             =   4560
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6588
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
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "WO Multi Line"
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
      Left            =   6480
      TabIndex        =   28
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "WO BuildType Data"
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
      Left            =   360
      TabIndex        =   27
      Top             =   4080
      Width           =   1815
   End
End
Attribute VB_Name = "FrmDefineBuildType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CboBuildType_Click()
    If CboBuildType.Text = "4" Then
        labelLine.Visible = True
        CmbLine.Visible = True
        labelside.Visible = True
        CmbSide.Visible = True
    Else
        labelLine.Visible = False
        CmbLine.Visible = False
        labelside.Visible = False
        CmbSide.Visible = False
    End If
End Sub

Private Sub CboLine_Click()
    If Trim(CboLine) <> "" Then
        Call GetWO(Trim(CboLine))
    End If

End Sub

Private Sub CboWo_Click()
    TxtWO = Trim(cboWO)
    Call GetWoinfoBasic(TxtWO)
End Sub

Private Sub cmdOK_Click()
Dim Rs1 As ADODB.Recordset
    If ChkErr(Trim(TxtWO)) = False Then
       Exit Sub
    End If

   strSQL = "Exec QSMS_SetBuildType '" & Trim(TxtWO) & "','" & Trim(CboBuildType) & "','" & Trim(CmbLine) & "','" & Trim(CmbSide) & "','" & Trim(g_userName) & "','" & Trim(CboStation) & "'"
   ''Conn.Execute strSql
   Set Rs1 = Conn.Execute(strSQL)  '''1185
   If Not Rs1.EOF Then
      If Rs1("Result") = "1" Then
          MsgBox Rs1("Msg"), vbOKOnly
          Exit Sub
      End If
   End If
   strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Define BuildType','" & Replace(strSQL, "'", """") & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
   Conn.Execute (strSQL)

   Call GetData
   
   If GetCheckBomFail(Trim(TxtWO), Trim(CboBuildType)) = True Then
   Else
      Exit Sub
   End If

   MsgBox "Set BuildType values is OK!", vbOKOnly

End Sub

Private Sub Form_Load()
    Dim str As String
    Dim RS As ADODB.Recordset
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
     
    Call GetLine
    CboBuildType.AddItem "1"
    CboBuildType.AddItem "2"
    CboBuildType.AddItem "3"
    CboBuildType.AddItem "4"
    Call GetData
    ''''''1185
    CboStation.AddItem "SP"
    CboStation.AddItem "SP2"
    ''''''1185
End Sub
 

Private Function GetLine()
Dim str As String
Dim RS As ADODB.Recordset
str = "select distinct Line from Sap_WO_List where Trans_Date>dbo.FormatDate(getdate()-180,'YYYYMMDDHHNNSS')"
Set RS = Conn.Execute(str)
CboLine.Clear
While Not RS.EOF
    CboLine.AddItem RS!Line
    CmbLine.AddItem RS!Line
    RS.MoveNext
Wend
End Function

Private Function GetWO(ByVal WOLine As String)
Dim str As String
Dim TransDate As String
Dim RS As ADODB.Recordset
cboWO.Clear
TxtWO = ""
TxtWOType = ""
TxtModel = ""
TxtGroup = ""
TxtWOQty = ""
txtBuild = ""
str = "select dbo.FormatDate(getdate()-60,'YYYYMMDDHHNNSS')"
Set RS = Conn.Execute(str)
TransDate = RS.Fields(0)

str = "select WO from Sap_Wo_List where Line='" & WOLine & "' and InitAOIFlag='Y' and Trans_Date>'" & TransDate & "' and QCCnt <=0 order by Trans_Date desc"
Set RS = Conn.Execute(str)
While Not RS.EOF
    cboWO.AddItem Trim(RS!WO)
    RS.MoveNext
Wend
End Function

Private Function GetWoinfoBasic(ByVal WO As String)
Dim str As String
Dim RS As ADODB.Recordset
str = "select PN, Qty ,MB_Rev,WO_Type,Line,[Group],BuildType from Sap_Wo_List where WO='" & Trim(WO) & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   TxtWOType = RS!PN
   TxtWOQty = RS!Qty
   TxtModel = RS!PN + "-" + Trim(RS!Mb_Rev)
   TxtGroup = Trim(RS![Group])
   txtBuild = Trim(RS!BuildType)
End If
End Function

Private Function GetCheckBomFail(ByVal Work_Order As String, BuildType) As Boolean
Dim str As String
Dim RS As ADODB.Recordset, rs2 As ADODB.Recordset
GetCheckBomFail = True
If Trim(Work_Order) = "" Then
   MsgBox "Please check the WO"
   GetCheckBomFail = False
   Exit Function
End If
str = "Delete QSMS_WO where Work_Order in(Select Wo from Sap_WO_List where [Group]='" & Trim(TxtGroup) & "')"
Conn.Execute (str)

'Check BOM
str = "Select Wo,BuildType from Sap_WO_List where [Group]='" & Trim(TxtGroup) & "'"
Set RS = Conn.Execute(str)
While Not RS.EOF
    str = "Exec QSMS_CheckBomSP '" & Trim(RS!WO) & "','N','" & Trim(RS!BuildType) & "'"
    Set rs2 = Conn.Execute(str)
    If rs2.EOF = False Then
       GetCheckBomFail = False
       MsgBox "Check bom fail"
    End If

    RS.MoveNext
Wend

str = "select *  from Sap_BOM_Fail  where Work_Order in (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "') "
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   GetCheckBomFail = False
   Call CopyToExcel(RS)
Else

End If
End Function

Private Function ChkErr(ByVal WO As String) As Boolean
Dim str As String
Dim RS As ADODB.Recordset
Dim TempRs As ADODB.Recordset
ChkErr = True
    Select Case CboBuildType
           Case "1", "2", "3", "4"
            If CboBuildType = "4" Then
                If Trim(CmbLine) = "" Or Trim(CmbSide) = "" Or Trim(CboStation) = "" Then  ''1185
                    ChkErr = False
                    MsgBox "BuildType=4,Please select the line and side and Station!", vbCritical
                    Exit Function
                End If
'                If UCase(Trim(CboLine)) = UCase(Trim(cmbLine)) Then   1185
'                    ChkErr = False
'                    MsgBox "BuildType=4,the build Line must be different with WO Line,please check!", vbCritical
'                    Exit Function
'                End If
            End If
           Case Else
               ChkErr = False
               MsgBox "BuildType values can only is 1,2,3 or 4,Please check!", vbCritical
               Exit Function
    End Select
 
    str = "Select distinct Work_Order from QSMS_Dispatch with(nolock) where Work_Order in(Select Wo from Sap_WO_List where [Group]='" & Trim(TxtGroup) & "')"
    Set TempRs = Conn.Execute(str)
    If Not TempRs.EOF Then
       ChkErr = False
       MsgBox "The PCB Work Order is dispatching ,can not be modify ,please check:" & TempRs!Work_Order
    End If

End Function

Private Function GetData()
    Dim str As String
    Dim RS As ADODB.Recordset, rs2 As ADODB.Recordset
    str = "select top 500 WO,PN,MB_Rev,BuildType,Line,Qty,WO_Type,CostBU,Trans_Date from Sap_Wo_List where BuildType<>'1' order by Trans_Date desc"
    Set RS = Conn.Execute(str)
    Set DataGDetail.DataSource = RS
    DataGDetail.Refresh

    str = "select top 500 * from WO_MultiLine order by TransDateTime desc"
    Set rs2 = Conn.Execute(str)
    Set DataWOMulti.DataSource = rs2
    DataWOMulti.Refresh
    
End Function
