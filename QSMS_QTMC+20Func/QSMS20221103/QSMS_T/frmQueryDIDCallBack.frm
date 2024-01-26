VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmQueryDIDCallBack 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Query DID Call Back"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dgDIDCallBack 
      Height          =   3555
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6271
      _Version        =   393216
      BackColor       =   -2147483629
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Query Condition"
      Height          =   2655
      Left            =   360
      TabIndex        =   1
      Top             =   300
      Width           =   9375
      Begin MSComCtl2.DTPicker dtFromTime 
         Height          =   375
         Left            =   3690
         TabIndex        =   14
         Top             =   1140
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61014018
         CurrentDate     =   40137
      End
      Begin VB.CommandButton cmdToExcel 
         Caption         =   "&ToExcel"
         Height          =   500
         Left            =   6420
         TabIndex        =   8
         Top             =   1800
         Width           =   1125
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "&Query"
         Height          =   500
         Left            =   4770
         TabIndex        =   7
         Top             =   1800
         Width           =   1365
      End
      Begin VB.TextBox txtInput 
         Height          =   360
         Left            =   1770
         TabIndex        =   0
         Top             =   1860
         Width           =   2685
      End
      Begin VB.OptionButton optCondition 
         BackColor       =   &H80000013&
         Caption         =   "CompPN"
         Height          =   525
         Index           =   3
         Left            =   5310
         TabIndex        =   5
         Top             =   450
         Width           =   1185
      End
      Begin VB.OptionButton optCondition 
         BackColor       =   &H80000013&
         Caption         =   "GroupID"
         Height          =   525
         Index           =   2
         Left            =   3420
         TabIndex        =   4
         Top             =   450
         Width           =   1185
      End
      Begin VB.OptionButton optCondition 
         BackColor       =   &H80000013&
         Caption         =   "DID"
         Height          =   525
         Index           =   1
         Left            =   1830
         TabIndex        =   3
         Top             =   450
         Width           =   825
      End
      Begin VB.OptionButton optCondition 
         BackColor       =   &H80000013&
         Caption         =   "WO"
         Height          =   525
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   450
         Value           =   -1  'True
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         Top             =   1140
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61014017
         CurrentDate     =   40101
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   375
         Left            =   2340
         TabIndex        =   10
         Top             =   1140
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61014019
         CurrentDate     =   40087
      End
      Begin MSComCtl2.DTPicker dtToTime 
         Height          =   375
         Left            =   7830
         TabIndex        =   15
         Top             =   1140
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Format          =   61014018
         CurrentDate     =   40137
      End
      Begin VB.Label lblCondition 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "WO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   13
         Top             =   1860
         Width           =   1155
      End
      Begin VB.Label lblSMT 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ToDateTime"
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
         Height          =   360
         Index           =   2
         Left            =   4860
         TabIndex        =   12
         Top             =   1140
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblSMT 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FromDateTime"
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
         Height          =   360
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   1140
         Visible         =   0   'False
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmQueryDIDCallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdQuery_Click()
    If Trim(txtInput) = "" Then
        txtInput.SetFocus
        MsgBox "Please input condition value!", vbInformation
    Else
        CmdQuery.Enabled = False
        cmdToExcel.Enabled = False
        Set dgDIDCallBack.DataSource = Query
        CmdQuery.Enabled = True
        cmdToExcel.Enabled = True
    End If
End Sub

Private Sub cmdToExcel_Click()
    If Trim(txtInput) = "" Then
        txtInput.SetFocus
        MsgBox "Please input condition value!", vbInformation
    Else
        CmdQuery.Enabled = False
        cmdToExcel.Enabled = False
        CopyToExcel Query
        CmdQuery.Enabled = True
        cmdToExcel.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    dtTO.Value = Now
    dtFrom.Value = DateDiff("D", 7, Now)
End Sub

Private Sub optCondition_Click(index As Integer)
    If index = 3 Then
        lblSMT(1).Visible = True
        lblSMT(2).Visible = True
        dtFrom.Visible = True
        dtTO.Visible = True
        dtFromTime.Visible = True
        dtToTime.Visible = True
    Else
        lblSMT(1).Visible = False
        lblSMT(2).Visible = False
        dtFrom.Visible = False
        dtTO.Visible = False
        dtFromTime.Visible = False
        dtToTime.Visible = False
        txtInput.SetFocus
    End If
    SetLabel index
    txtInput = ""
End Sub

Private Function Query() As ADODB.Recordset
    Dim sql As String
    Dim condition As String
    Dim sDT As String
    Dim eDT As String
    
    If optCondition(0).Value = True Then
        condition = "WO"
    ElseIf optCondition(1).Value = True Then
        condition = "DID"
    ElseIf optCondition(2).Value = True Then
        condition = "GroupID"
    ElseIf optCondition(3).Value = True Then
        condition = "CompPN"
        sDT = Format(dtFrom.Value, "YYYYMMDD") & Format(dtFromTime.Value, "HHMMSS")
        eDT = Format(dtTO.Value, "YYYYMMDD") & Format(dtToTime.Value, "HHMMSS")
    End If
    sql = "Exec QueryDIDCallBackByCondition " & sq(txtInput) & "," & sq(condition) & "," & sq(sDT) & "," & sq(eDT)
    Set Query = Conn.Execute(sql)
End Function

Private Sub SetLabel(index As Integer)
    Select Case index
        Case 0
            lblCondition = "WO:"
        Case 1
            lblCondition = "DID:"
        Case 2
            lblCondition = "GroupID:"
        Case 3
            lblCondition = "CompPN:"
    End Select
End Sub

Private Sub txtInput_Click()
    SendKeys "{HOME}+{END}"
End Sub
