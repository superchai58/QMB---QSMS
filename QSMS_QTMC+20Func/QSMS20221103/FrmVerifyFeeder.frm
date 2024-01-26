VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FrmVerifyFeeder 
   Caption         =   "Verify Feeder & Slot"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraDID 
      BackColor       =   &H80000013&
      Caption         =   "DID & Feeder"
      Height          =   5055
      Left            =   0
      TabIndex        =   14
      Top             =   2640
      Width           =   8055
      Begin VB.TextBox TxtLR 
         Enabled         =   0   'False
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
         Left            =   6120
         TabIndex        =   31
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox TxtCompPN 
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
         Left            =   1320
         TabIndex        =   28
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox TxtFeeder 
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
         Left            =   1320
         TabIndex        =   20
         Top             =   2280
         Width           =   3375
      End
      Begin VB.ComboBox CboMachine 
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
         Left            =   1320
         TabIndex        =   19
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox CboJobPN 
         Enabled         =   0   'False
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
         Left            =   1320
         TabIndex        =   18
         Top             =   960
         Width           =   2415
      End
      Begin VB.ComboBox CboVersion 
         Enabled         =   0   'False
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
         Left            =   5040
         TabIndex        =   17
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox TxtSlot 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
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
         Left            =   1320
         TabIndex        =   16
         Top             =   2880
         Width           =   3375
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
         Left            =   5160
         Picture         =   "FrmVerifyFeeder.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         Caption         =   "LR"
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
         Index           =   2
         Left            =   4800
         TabIndex        =   30
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFC0FF&
         BorderWidth     =   10
         X1              =   0
         X2              =   7920
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label4 
         BackColor       =   &H00808080&
         Caption         =   "CompPN"
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
         Left            =   0
         TabIndex        =   27
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "Feeder"
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
         Left            =   0
         TabIndex        =   26
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "Machine"
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
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "Job"
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
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "Revision"
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
         Index           =   3
         Left            =   3840
         TabIndex        =   23
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "Slot"
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
         Index           =   2
         Left            =   0
         TabIndex        =   22
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label LblMessage 
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   3480
         Width           =   7095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "WO Information"
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
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
         Left            =   1440
         TabIndex        =   6
         Top             =   1320
         Width           =   2535
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
         Left            =   1440
         TabIndex        =   5
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox TxtLine 
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
         TabIndex        =   4
         Top             =   840
         Width           =   1575
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
         Left            =   3000
         Picture         =   "FrmVerifyFeeder.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.ListBox ListNotDispatch 
         Height          =   840
         ItemData        =   "FrmVerifyFeeder.frx":0884
         Left            =   5160
         List            =   "FrmVerifyFeeder.frx":0886
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.ListBox ListClosed 
         Height          =   840
         ItemData        =   "FrmVerifyFeeder.frx":0888
         Left            =   5160
         List            =   "FrmVerifyFeeder.frx":088A
         TabIndex        =   1
         Top             =   1320
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   69795843
         CurrentDate     =   36482
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Line"
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
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
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
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "WO Closed "
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
         Height          =   615
         Index           =   0
         Left            =   4080
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Wo Not Disptach"
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
         Height          =   615
         Index           =   3
         Left            =   4080
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
   End
   Begin MCI.MMControl wave_control 
      Height          =   330
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "FrmVerifyFeeder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private VendorCode, DateCode, LotCode As String
Private Sub CboGroupID_Click()
Call GetWoByGroupID(Trim(CboGroupID))
End Sub

Private Sub CboGroupID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboGroupID_Click
End If
End Sub


Private Sub CboMachine_Click()
Call GetJobByMachine(Trim(CboMachine))
TxtFeeder.SetFocus
End Sub

Private Sub CboMachine_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboMachine_Click
End If
End Sub





Private Sub cboWO_Click()
Call GetMachineByWo(cboWO)
End Sub

Private Sub cboWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call cboWO_Click
End If
End Sub

Private Sub CmdQuery_Click()
Dim TransDate As String
TransDate = Format(dtpSDate, "YYYY/MM/DD")
TransDate = Replace(TransDate, "-", "")
TransDate = Replace(TransDate, "/", "")
If Trim(TxtLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If
Call GetGroupIDByLine(Trim(TxtLine), TransDate)
End Sub

Private Sub cmdReset_Click()
TxtFeeder.Enabled = True
TxtSlot.Enabled = True
TxtFeeder.Text = ""
TxtSlot.Text = ""
TxtFeeder.SetFocus

     
     

End Sub



Private Sub CmdSave_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Or KeyAscii = 9 Then
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
Dim str As String
Dim RS As ADODB.Recordset
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
str = "select getdate()"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
    'Date = Rs(0)
    'Time = Rs(0)
End If
dtpSDate = Date

End Sub








Private Sub TxtDID_Change()

End Sub

Private Sub TxtFeeder_KeyPress(KeyAscii As Integer)
Dim RS As ADODB.Recordset
Dim str As String

If KeyAscii = 13 Or KeyAscii = 9 Then
   VendorCode = ""
   DateCode = ""
   LotCode = ""
   str = "select CompPN,VendorCode,DateCode,LotCode from QSMS_Feeder where Feeder='" & Trim(TxtFeeder) & "'"
   Set RS = Conn.Execute(str)
   If Not RS.EOF Then
      TxtCompPN = Trim(RS!CompPN)
      VendorCode = Trim(RS!VendorCode)
      DateCode = Trim(RS!DateCode)
      LotCode = Trim(RS!LotCode)
      TxtFeeder.Enabled = False
      TxtSlot.SetFocus
   Else
      Call Warning_Sound
      TxtFeeder.Text = ""
      TxtFeeder.SetFocus
   End If
   
End If


End Sub




Private Function GetJobByMachine(ByVal Machine As String)
Dim str As String
Dim RS As ADODB.Recordset
If Machine = "" Then
   MsgBox "Please select Machine"
   Exit Function
End If

CboJobPN.Clear
str = "select Distinct JobPn from QSMS_MEBom where Machine='" & Machine & "' and JobPN in (select Jobpn from QSMS_JobBOM where Work_order='" & Trim(cboWO) & "')"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
      CboJobPN.Text = Trim(RS!Jobpn)
      
End If
End Function

Private Sub Warning_Sound()
      wave_control.FileName = App.Path & "\OO.wav"
      wave_control.Command = "open"
      wave_control.Command = "play"
      Do While wave_control.Mode = mciModePlay
      Loop
      wave_control.Command = "close"
End Sub
Private Sub OK_Sound()
    wave_control.FileName = App.Path & "\OK.wav"
    wave_control.Command = "open"
    wave_control.Command = "play"
    Do While wave_control.Mode = mciModePlay
    Loop
    wave_control.Command = "close"
End Sub




Private Function ChkFeeder() As Boolean
Dim RS As ADODB.Recordset
Dim str As String

ChkFeeder = True
   str = "select DID from QSMS_Feeder where Feeder='" & Trim(TxtFeeder) & "' "
   Set RS = Conn.Execute(str)
   If Not RS.EOF Then
      MsgBox "The Feeder has been used,Please check and delete. or use another Feeder"
      ChkFeeder = False
      Exit Function
   End If
End Function



Private Function GetGroupIDByLine(ByVal Line As String, ByVal TransDate As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim GroupIDHead As String
GroupIDHead = Line & TransDate
CboGroupID.Clear
str = "Select distinct GroupID from QSMS_WoGroup where GroupID>'" & GroupIDHead & "'" 'and Sap1Flag='Y' and ClosedFlag='N'"
Set RS = Conn.Execute(str)
While Not RS.EOF
     CboGroupID.AddItem Trim(RS!GroupID)
     RS.MoveNext
Wend

End Function

Private Function GetWoByGroupID(ByVal GroupID As String)
Dim str As String
Dim RS As ADODB.Recordset


cboWO.Clear
ListNotDispatch.Clear
ListClosed.Clear
str = "Select Work_Order,Sap1Flag,ClosedFlag from QSMS_WoGroup where GroupID='" & GroupID & "' "
Set RS = Conn.Execute(str)
While Not RS.EOF
     If UCase(Trim(RS!sap1flag)) = "Y" And UCase(Trim(RS!ClosedFlag)) = "N" Then
        cboWO.AddItem Trim(RS!Work_Order)
     End If
      If UCase(Trim(RS!sap1flag)) = "N" Then
         ListNotDispatch.AddItem Trim(RS!Work_Order)
      End If
      If UCase(Trim(RS!ClosedFlag)) = "Y" Then
         ListClosed.AddItem Trim(RS!Work_Order)
      End If
     RS.MoveNext
Wend

End Function

Private Function GetMachineByWo(ByVal WO As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim Rev As String

str = "Select Mb_Rev from Sap_Wo_List where Wo='" & WO & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   Rev = Trim(RS!Mb_Rev)
   CboVersion = Rev
End If
CboMachine.Clear
str = "select distinct Machine From QSMS_MEbom   where JobPN in (select JobPN from QSMS_JobBOM where Work_Order='" & WO & "') and  Version='" & Rev & "'"
Set RS = Conn.Execute(str)
While Not RS.EOF
      CboMachine.AddItem Trim(RS!Machine)
      RS.MoveNext
Wend
End Function


Private Sub TxtSlot_KeyPress(KeyAscii As Integer)
Dim str As String
Dim RS As ADODB.Recordset
Dim TempCompPN As String
If KeyAscii = 13 Or KeyAscii = 9 Then
   str = "select CompPN from QSMS_MEBom where Machine='" & CboMachine & "' and JObpN='" & CboJobPN & "' and version='" & CboVersion & "' and Slot='" & Trim(TxtSlot) & "' and LR='" & Trim(TxtLR) & "'"
   Set RS = Conn.Execute(str)
   If RS.EOF Then
      Call Warning_Sound
      TxtSlot.Text = ""
      TxtSlot.SetFocus
      LblMessage.Caption = "can not find the ComppN By the Slot,Please check"
      Exit Sub
   Else
       TempCompPN = Trim(RS!CompPN)
   End If
   If UCase(TempCompPN) = UCase(TxtCompPN) Then
      Call UpdateQSMS_Feeder
      Call ChkMachineVerifyFinished
      Call OK_Sound
      TxtFeeder.Enabled = True
      TxtFeeder.Text = ""
      TxtSlot.Text = ""
      TxtFeeder.SetFocus
   Else
       LblMessage.Caption = " ComppN is different with the Feeder ,Please check: " & TempCompPN
       Call Warning_Sound
      TxtSlot.Text = ""
      TxtSlot.SetFocus
   End If
End If
End Sub
Private Function UpdateQSMS_Feeder()
Dim str As String
Dim RS As ADODB.Recordset
str = "Update QSMS_Feeder set Slot='" & Trim(TxtSlot) & "' where Feeder='" & TxtFeeder & "'"
Conn.Execute str
End Function
Private Function InsertQSMS_Verify()
Dim str As String
Dim RS As ADODB.Recordset
Dim TransDateTime As String
str = "select getdate()"
Set RS = Conn.Execute(str)
 TransDateTime = Format(RS.Fields(0), "YYYYMMDDHHMMSS")
str = "Select top 1 EndDateTime from QSMS_Verify where Machine='" & Trim(CboMachine) & "' and CompPN='" & Trim(TxtCompPN) & "' and VendorCode='" & VendorCode & "' and LotCode='" & LotCode & "' and DateCode='" & DateCode & "' order by BeginDateTime Desc"
Set RS = Conn.Execute(str)
If RS.EOF Or Len(Trim(RS!EndDateTime)) = 14 Then
   str = "Insert into QSMS_Verify(Machine,JobPN,Version,CompPN,VendorCode,DateCode,LotCode,BeginDateTime,EndDateTime) valuse " & _
         " ('" & Trim(CboMachine) & "','" & Trim(CboJobPN) & "','" & Trim(CboVersion) & "','" & Trim(TxtCompPN) & "','" & VendorCode & "','" & DateCode & "','" & LotCode & "','" & TransDateTime & "','')"
   Conn.Execute (str)
End If

End Function

Private Function UpdateQSMS_Verify()
Dim str As String
Dim RS As ADODB.Recordset
Dim TempRs As ADODB.Recordset
Dim TransDateTime As String
str = "select getdate()"
Set RS = Conn.Execute(str)
TransDateTime = Format(RS.Fields(0), "YYYYMMDDHHMMSS")

str = "select CompPN,VendorCode,DateCoce,LotCode from qsms_verify where machine='" & CboMachine & "'' and enddatetime=''"
Set RS = Conn.Execute(str)
While Not RS.EOF
      str = "Select CompPN from QSMS_Feeder where Machine='" & Trim(CboMachine) & "' and JobPN='" & Trim(CboJobPN) & "' and Version='" & CboVersion & "' and CompPN='" & Trim(RS!CompPN) & "' and VendorCode='" & Trim(RS!VendorCode) & "'"
      Set TempRs = Conn.Execute(str)
      If TempRs.EOF Then
         str = "Update QSMS_Feeder set EndDateTime='" & TransDateTime & "' where Machine='" & CboMachine & "' and EndDateTime='' and CompPN='" & Trim(RS!CompPN) & "' and VendorCode='" & Trim(RS!VendorCode) & "'"
         Conn.Execute str
      End If
      RS.MoveNext
Wend
End Function
Private Function ChkMachineVerifyFinished()
Dim str As String
Dim RS As ADODB.Recordset
str = "select Machine from QSMS_Feeder where Machine='" & Trim(CboMachine) & "' and Jobpn='" & Trim(CboJobPN) & "' and Version='" & Trim(CboVersion) & "' and Slot=''"
Set RS = Conn.Execute(str)
If RS.EOF Then
   Call UpdateQSMS_Verify
End If
End Function
