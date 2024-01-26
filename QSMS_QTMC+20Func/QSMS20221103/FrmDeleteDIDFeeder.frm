VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FrmDeleteDIDFeeder 
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraFeederDelete 
      BackColor       =   &H80000013&
      Caption         =   "Delete Feeder"
      Height          =   2055
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   7575
      Begin VB.ComboBox CboFeeder 
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
         Left            =   1200
         TabIndex        =   9
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton cmdUnLink 
         Caption         =   "UnLink Feeder"
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
         Left            =   5640
         Picture         =   "FrmDeleteDIDFeeder.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Lblmess 
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   7095
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
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame FraDIDDelete 
      BackColor       =   &H80000013&
      Caption         =   "Delete DID"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7575
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete DID"
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
         Left            =   6000
         Picture         =   "FrmDeleteDIDFeeder.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox CboDID 
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
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label LblMessage 
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   7095
      End
      Begin VB.Label Label1 
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
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
   Begin MCI.MMControl wave_control 
      Height          =   330
      Left            =   0
      TabIndex        =   7
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
Attribute VB_Name = "FrmDeleteDIDFeeder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub CmdDelete_Click()

Call DeleteDID(Trim(CboDID))
LblMessage = "Delete OK"
End Sub

Private Sub cmdUnLink_Click()
Dim Str As String
Dim Rs As ADODB.Recordset
Dim TransDateTime As String
Str = "select DID from QSMS_Feeder where Feeder='" & CboFeeder & "'"
Set Rs = Conn.Execute(Str)
If Rs.EOF Then
   MsgBox "can not find the record,please check"
   Exit Sub
End If
  
Str = "select getdate()"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
    TransDateTime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
End If
'(1) Backup the Feeder
Str = "Insert into QSMS_Feeder_Delete(Machine,JobPN,Version,DID,VendorCode,DateCode,LotCode,Feeder,Slot,LR,UID,TransDateTime,DeleteDateTime) " & _
     " Select Machine,JobPN,Version,DID,VendorCode,DateCode,LotCode,Feeder,Slot,LR,UID,TransDateTime,'" & TransDateTime & "' from QSMS_Feeder where Feeder='" & Trim(CboFeeder) & "'"

Conn.Execute Str


'(2) Delete DID from QSMS_Feeder
Str = "delete from QSMS_Feeder where Feeder='" & Trim(CboFeeder) & "'"
Conn.Execute Str
Call OK_Sound
Lblmess = "Delete OK"

End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
Select Case UCase(DeleteType)
      Case "DID"
           
            FraFeederDelete.Visible = False
            FraDIDDelete.Visible = True
            Me.Caption = "Delete DID"
      Case "FEEDER"
            
            FraFeederDelete.Top = FraDIDDelete.Top
            FraFeederDelete.Left = FraDIDDelete.Left
            FraFeederDelete.Visible = True
            FraDIDDelete.Visible = False
            Me.Caption = "Delete Feeder"
End Select

End Sub

'Private Function GetDID()
'Dim Str As String
'Dim Rs As ADODB.Recordset
'CboDID.Clear
'Str = "select DID from QSMS_Feeder order by DID"
'Set Rs = Conn.Execute(Str)
'While Not Rs.EOF
'      CboDID.AddItem Trim(Rs!DID)
'      Rs.MoveNext
'Wend
'End Function
'Private Function GetMachine()
'Dim Str As String
'Dim Rs As ADODB.Recordset
'CboMachine.Clear
'Str = "select distinct Machine from QSMS_Feeder order by Machine"
'Set Rs = Conn.Execute(Str)
'While Not Rs.EOF
'      CboMachine.AddItem Trim(Rs!Machine)
'      Rs.MoveNext
'Wend
'
'End Function

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

'Private Function GetJobByMachine(ByVal Machine As String)
'Dim Str As String
'Dim Rs As ADODB.Recordset
'CboJobPN.Clear
'Str = "select distinct JobPN from QSMS_Feeder where Machine='" & Machine & "'"
'Set Rs = Conn.Execute(Str)
'While Not Rs.EOF
'      CboJobPN.AddItem Trim(Rs!JobPN)
'      Rs.MoveNext
'Wend
'
'End Function

'Private Function GetVersionByMachineJobPN(ByVal Machine As String, ByVal JobPN As String)
'Dim Str As String
'Dim Rs As ADODB.Recordset
'CboVersion.Clear
'Str = "select distinct Version from QSMS_Feeder where Machine='" & Machine & "' and JobPN='" & JobPN & "'"
'Set Rs = Conn.Execute(Str)
'While Not Rs.EOF
'      CboVersion.AddItem Trim(Rs!Version)
'      Rs.MoveNext
'Wend
'
'End Function
'
'Private Function GetFeeder(ByVal Machine As String, ByVal JobPN As String, ByVal Version As String)
'Dim Str As String
'Dim Rs As ADODB.Recordset
'CboFeeder.Clear
'Str = "select distinct Feeder from QSMS_Feeder where Machine='" & Machine & "' and JobPN='%" & JobPN & "' and Version like '%" & Version & "'"
'Set Rs = Conn.Execute(Str)
'While Not Rs.EOF
'      CboFeeder.AddItem Trim(Rs!Feeder)
'      Rs.MoveNext
'Wend
'
'End Function

Private Function DeleteDID(ByVal DID As String)
Dim Str As String
Dim Rs As ADODB.Recordset
Dim TransDateTime As String
Str = "select getdate()"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
    TransDateTime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
End If
'(1) Backup the DID
Str = "Insert into QSMS_Feeder_Delete(Machine,JobPN,Version,DID,VendorCode,DateCode,LotCode,Feeder,Slot,LR,UID,TransDateTime,DeleteDateTime) Values " & _
      "Select Machine,,JobPN,Version,DID,VendorCode,DateCode,LotCode,Feeder,Slot,LR,UID,TransDateTime,'" & TransDateTime & "' from QSMS_Feeder where DID='" & Trim(DID) & "'"
Conn.Execute Str
'(2) Delete DID from QSMS_Feeder
Str = "delete from QSMS_Feeder where DID='" & Trim(DID) & "'"
Conn.Execute Str
'(3) Delete DID from QSMS_DID

Str = "delete from QSMS_DID where DID='" & Trim(DID) & "'"
Conn.Execute Str
'(4) Update QSMS_dispatch
Str = "Update QSMS_Dispatch set DeletedFlag='Y' where DID='" & Trim(DID) & "' and DeletedFlag='N' "
Conn.Execute Str
Call OK_Sound
End Function


