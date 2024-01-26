VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSingleSideBrdConfirm 
   Caption         =   "FrmSingleSideBrdConfirm"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin VB.Frame FrmConfirSingleSideBrd 
         Caption         =   "Confirm Single Side Brd"
         Height          =   1335
         Left            =   120
         TabIndex        =   32
         Top             =   3600
         Width           =   9375
         Begin VB.ComboBox CboBuildType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1440
            TabIndex        =   37
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton CmdConfirm 
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   4680
            Picture         =   "FrmSingleSideBrdConfirm.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Opt_Both 
            Caption         =   "Both Side"
            Height          =   375
            Left            =   3120
            TabIndex        =   35
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton Opt_Solder 
            Caption         =   "Solder Side"
            Height          =   375
            Left            =   1800
            TabIndex        =   34
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton Opt_Comp 
            Caption         =   "Component Side"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackColor       =   &H000000FF&
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
            Index           =   3
            Left            =   240
            TabIndex        =   38
            Top             =   840
            Width           =   1095
         End
      End
      Begin VB.ComboBox cboWO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6600
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox TxtMBPN 
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
         Left            =   1560
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2055
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
         Left            =   7680
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1335
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "Group"
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton OptRelease 
         Caption         =   "Release"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdQuery 
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3360
         Picture         =   "FrmSingleSideBrdConfirm.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox CboGroupID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6600
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Width           =   2655
      End
      Begin VB.ComboBox CboNotFinishedWO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6600
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2655
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
         Left            =   1560
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox TxtCustomer 
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
         Left            =   4920
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1815
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
         Left            =   7680
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Frame FraSB 
         Caption         =   "Small Board WO"
         Height          =   615
         Left            =   6600
         TabIndex        =   4
         Top             =   1800
         Visible         =   0   'False
         Width           =   2655
         Begin VB.ComboBox CboSBWO 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   240
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
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
         Left            =   4920
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1815
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ComboBox CboNotChkBOM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6600
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   600
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   62390275
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   62390275
         CurrentDate     =   36482
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "BeginDate"
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
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "OK Work Order"
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
         Left            =   4440
         TabIndex        =   30
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MB PN"
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
         Left            =   120
         TabIndex        =   29
         Top             =   3000
         Width           =   1455
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
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   1455
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
         Left            =   6960
         TabIndex        =   27
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "GroupID"
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
         Left            =   4440
         TabIndex        =   26
         Top             =   120
         Width           =   2175
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
         TabIndex        =   25
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Customer"
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
         Index           =   16
         Left            =   3840
         TabIndex        =   24
         Top             =   2520
         Width           =   1095
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
         Left            =   6960
         TabIndex        =   23
         Top             =   2520
         Width           =   735
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
         Left            =   3840
         TabIndex        =   22
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Un OK Work Order"
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
         Index           =   4
         Left            =   4440
         TabIndex        =   21
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "End Date"
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
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Chk BOM fail/not Chk"
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
         Index           =   5
         Left            =   4440
         TabIndex        =   19
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmSingleSideBrdConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSql As String

Private Sub CboGroupID_Click()
Call GetGroupWO(CboGroupID)
End Sub

Private Sub CboNotFinishedWO_Click()
TxtWO = Trim(CboNotFinishedWO)

Call GetWoinfo(TxtWO)
End Sub

Private Sub CboNotFinishedWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CboNotFinishedWO_Click
End If
End Sub

Private Sub CboWo_Click()

TxtWO = Trim(cboWO)

Call GetWoinfo(TxtWO)
End Sub

Private Sub cboWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CboWo_Click
End If
End Sub

Private Sub CmdConfirm_Click()
Dim Str As String
Dim Rs As ADODB.Recordset
Dim Machine As String

If ChkErr(Trim(TxtWO)) = False Then
   Exit Sub
End If

If GetCheckBomFail(Trim(TxtWO), Trim(CboBuildType)) = False Then
   Exit Sub
End If
strSql = "Update Sap_WO_list set  buildtype='" & Trim(CboBuildType) & "', SType='1' where WO in (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "') "
Conn.Execute strSql
MsgBox "Confirm OK"

'Str = "Select Wo from Sap_WO_List where [Group]='" & Trim(TxtGroup) & "'"
'Set Rs = Conn.Execute(Str)
'While Not Rs.EOF
'      If Opt_Comp.Value = True Then
'        Call SingleSideConfirm(Trim(CboLine), "C", Trim(Rs!Wo)) ''Debug By gary ,transfer txtwo to rs!WO ,2006/12/17
'        Call InsertIntoQSMSLog("SingleSideBrd", "SingleSideBrd", "Work_Order :" & Rs!Wo & " Update to component side")
'        'SType=1 stand for SingleSideBoard;SType=2 stand for BothSideBoard
'        'BuildType='3',
'        strSql = "Update Sap_WO_list set  SType='1'" & _
'            " where WO='" & Trim(TxtWO) & "' and PN='" & Trim(TxtMBPN) & "'"
'        Conn.Execute strSql
'        MsgBox "Confirm OK"
'      End If
'
'      If Opt_Solder.Value = True Then
'        Call SingleSideConfirm(Trim(CboLine), "S", Trim(Rs!Wo)) ''Debug By gary ,transfer txtwo to rs!WO,2006/12/17
'        Call InsertIntoQSMSLog("SingleSideBrd", "SingleSideBrd", "Work_Order :" & Rs!Wo & " Update to Solder side")
'        'BuildType='2',
'        strSql = "Update Sap_WO_list set SType='1'" & _
'           " where WO='" & Trim(TxtWO) & "' and PN='" & Trim(TxtMBPN) & "'"
'        Conn.Execute strSql
'        MsgBox "Confirm OK"
'      End If
'
'     If Opt_Both.Value = True Then
'        Call InsertIntoQSMSLog("SingleSideBrd", "SingleSideBrd", "Work_Order :" & Rs!Wo & " Update to Both side")
'        'BuildType='1',
'        strSql = "Update Sap_WO_list set SType='1'" & _
'            " where WO='" & Trim(TxtWO) & "' and PN='" & Trim(TxtMBPN) & "'"
'        Conn.Execute strSql
'        MsgBox "Confirm OK"
'     End If
'
'    Rs.MoveNext
'Wend


End Sub



Private Sub CmdQuery_Click()
If Trim(CboLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If
Call GetGroupID
End Sub

Private Sub Form_Load()
Dim Str As String
Dim Rs As ADODB.Recordset
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
Str = "select getdate()"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
    'Date = Rs(0)
    'Time = Rs(0)
End If
dtpSDate = Date
dtpEDate = Date
Call GetLine
End Sub
Private Function GetLine()
Dim Str As String
Dim Rs As ADODB.Recordset
Str = "select distinct Line from QSMS_woGroup"
Set Rs = Conn.Execute(Str)
CboLine.Clear
While Not Rs.EOF
    CboLine.AddItem Rs!Line
    Rs.MoveNext
Wend
End Function


Private Function GetGroupID()
Dim Str As String
Dim BeginDate, EndDate As String
Dim GroupIDHead As String
Dim i As Long
Dim Rs As ADODB.Recordset
BeginDate = Format(dtpSDate, "YYYY/MM/DD")
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
EndDate = Format(dtpEDate, "YYYY/MM/DD")
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")
'GroupIDHead = Trim(CboLine) & TransDate
If OptRelease.Value = True Then
   Str = "select distinct GroupID from QSMS_WOGroup a,sap_wo_list b where a.WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' " & _
         "and a.line='" & CboLine & "' and a.work_order=b.wo and b.PN in (select MBPN from QSMS_SingleSideBrd)"
Else
    Str = "select distinct GroupID from QSMS_WOGroup a,sap_wo_list b   where substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' " & _
          "and a.line='" & CboLine & "' and a.work_order=b.wo and b.PN in (select MBPN from QSMS_SingleSideBrd)"
End If

Set Rs = Conn.Execute(Str)
i = 0
CboGroupID.Clear
While Not Rs.EOF
      CboGroupID.AddItem Trim(Rs!GroupID)
      Rs.MoveNext
      i = i + 1
Wend
If i = 0 Then
   MsgBox "No data"
   
End If

End Function
Private Function GetGroupWO(ByVal GroupID As String)
Dim Str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset


Str = "select a.Work_Order from QSMS_WOGroup a,Sap_Wo_List b  where a.GroupID= '" & GroupID & "' and a.work_order=b.wo  " & _
      " and b.PN in (select MBPN from QSMS_SingleSideBrd)order by Seq_NO"

Set Rs = Conn.Execute(Str)
cboWO.Clear
CboNotFinishedWO.Clear
CboNotChkBOM.Clear
While Not Rs.EOF
      If ChkMBWo(Rs!Work_Order) = True Then
            If ChkQSMS_WO(Trim(Rs!Work_Order)) = False Then
                CboNotChkBOM.AddItem Trim(Rs!Work_Order)
            Else
            
                
                If ChkWoFinished(Rs!Work_Order) = True Then
    
                    cboWO.AddItem Trim(Rs!Work_Order)
                Else
                    
                     CboNotFinishedWO.AddItem Trim(Rs!Work_Order)
                End If
            End If
      End If
      Rs.MoveNext
Wend

End Function
Private Function ChkErr(ByVal WO As String) As Boolean
Dim Str As String
Dim Rs As ADODB.Recordset
Dim TempRs As ADODB.Recordset
ChkErr = True
If Trim(TxtGroup) = "" Then
    ChkErr = False
    MsgBox "Please select the Work_Order"
End If
If CboBuildType.Text = "1" Or CboBuildType.Text = "2" Or CboBuildType.Text = "3" Then
Else
    ChkErr = False
    MsgBox "please select side"
End If
Str = "Select Wo from sap_wo_list a,QSMS_SingleSideBrd b where a.wo='" & WO & "' and a.PN =b.MBPN"
Set Rs = Conn.Execute(Str)
If Rs.EOF Then
   ChkErr = False
   MsgBox "The WO is not single side brd,Please check or maintain"
End If
Str = "Select Wo from Sap_WO_List where [Group]='" & Trim(TxtGroup) & "'"
Set Rs = Conn.Execute(Str)
If Rs.EOF Then
   ChkErr = False
   MsgBox "Can not find wo by the Group:" & Trim(TxtGroup) & ""
End If
While Not Rs.EOF
'    Str = "Select distinct Work_Order from QSMS_WO where work_order='" & Rs!WO & "'"
'    Set TempRs = Conn.Execute(Str)
'    If TempRs.EOF Then
'       ChkErr = False
'       MsgBox "Please check BOM first"
'
'    End If
    Str = "Select distinct Work_Order from QSMS_Dispatch where work_order='" & Rs!WO & "'"
    Set TempRs = Conn.Execute(Str)
    If Not TempRs.EOF Then
       ChkErr = False
       MsgBox "The Work Order is dispatching ,can not be modify ,please check:" & Rs!WO
    End If
    Rs.MoveNext
Wend
End Function


Private Function GetCheckBomFail(ByVal Work_Order As String, ByVal BuildType As String) As Boolean
Dim Str As String
Dim Rs As ADODB.Recordset, rs2 As ADODB.Recordset
GetCheckBomFail = True
If Trim(Work_Order) = "" Then
   MsgBox "Please check the WO"
   GetCheckBomFail = False
   Exit Function
End If

'delete old failure log
Str = "delete from Sap_BOM_Fail  where Work_Order  in (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "') "
Conn.Execute (Str)

'Check BOM
Str = "Select Wo from Sap_WO_List where [Group]='" & Trim(TxtGroup) & "'"
Set Rs = Conn.Execute(Str)
While Not Rs.EOF
    Str = "delete from QSMS_Wo where work_Order='" & Trim(Rs!WO) & "'"
    Conn.Execute Str
    
    Str = "Exec QSMS_CheckBomSP '" & Trim(Rs!WO) & "','N','" & Trim(BuildType) & "'"               ''(0019)
    Set rs2 = Conn.Execute(Str)
    If rs2.EOF = False Then
       GetCheckBomFail = False
       MsgBox "Check bom fail"
    End If
    
''    If CheckBom(Rs!WO, "N", BuildType) = False Then
''       GetCheckBomFail = False
''       MsgBox "Check bom fail"
''    End If
    Rs.MoveNext
Wend

Str = "select *  from Sap_BOM_Fail  where Work_Order in (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "') "
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   GetCheckBomFail = False
   Call CopyToExcel(Rs)
Else
  ' MsgBox "Check BOM OK"
End If
End Function

Private Function GetWoinfo(ByVal WO As String)
Dim Str As String
Dim Rs As ADODB.Recordset
Str = "select PN, Qty,[Group],Line from Sap_Wo_List where WO='" & Trim(WO) & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   TxtMBPN = Rs!PN
   TxtModel = Mid(TxtMBPN, 3, 3)
   TxtWOQty = Rs!Qty
   TxtGroup = Trim(Rs![Group])
   CboLine = Trim(Rs!Line)
End If
Str = "select Customer from ModelName where PN='" & TxtMBPN & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   TxtCustomer = Trim(Rs!Customer)
End If
End Function

Private Sub Opt_Both_Click()
CboBuildType.Text = "1"
End Sub

Private Sub Opt_Comp_Click()
CboBuildType.Text = "3"
End Sub

Private Sub Opt_Solder_Click()
CboBuildType.Text = "2"
End Sub
