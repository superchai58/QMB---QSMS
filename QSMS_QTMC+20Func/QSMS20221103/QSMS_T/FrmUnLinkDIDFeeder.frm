VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FrmUnlinkDIDFeeder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Frm Unlink DID Feeder[20090914]"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "UnLink DID Feeder"
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      Begin VB.TextBox TxtDID 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   9
         Top             =   840
         Width           =   3495
      End
      Begin VB.ComboBox Cbostatus 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CommandButton cmdSave 
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
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TxtDIDFeeder 
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
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
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label LblMessage 
         BackColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "DID status"
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
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "Feeder/DID"
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
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin MCI.MMControl wave_control 
      Height          =   330
      Left            =   10080
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "FrmUnlinkDIDFeeder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************************************************************
'**文 件 名: FrmUnlinkDIDFeeder.frm
'**Copyright (C) 2009-0422 QMS
'**文件编号:
'**创 建 人: Jeanson
'**日    期: 2009.08.21
'**描    述:为了解除Feeder/DID的关联，虽然我们已经在Maintain Feeder/DID时会自动解除之间的关系，但是如果之前Feeder左右都有上材料，现在只有一边上材料时，仍需使用此功能！
'
'**EQMS_ID                 修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'QMS                        Lynn       2009.08.21      Unlink feeder,cancel step 2, we do not to update DID RealQty! (0001)
'***********************************************************************************/
Option Explicit

Private Sub cmdSave_Click()
Dim Str As String
Dim Rs As ADODB.Recordset
Dim TransDateTime As String
Dim DID As String
If InStr(1, TxtDIDFeeder, "-") > 0 Then
   DID = Trim(TxtDIDFeeder)
Else
   Str = "select DID from QSMS_Feeder where Feeder='" & TxtDIDFeeder & "'"
   Set Rs = Conn.Execute(Str)
    If Rs.EOF Then
       MsgBox "can not find the record,please check"
       LblMessage.BackColor = &HFF00&
       LblMessage = "Unlink DID  OK"
       Exit Sub
    Else
       DID = Trim(Rs!DID)
    End If
  
End If

Str = "select getdate()"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
    TransDateTime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
End If
'(1) Backup the Feeder
Str = "Insert into QSMS_Feeder_Delete(Machine,JobPN,Version,DID,VendorCode,DateCode,LotCode,Feeder,Slot,LR,UID,TransDateTime,DeleteDateTime) " & _
     " Select Machine,JobPN,Version,DID,VendorCode,DateCode,LotCode,Feeder,Slot,LR,'" & g_userName & "',TransDateTime,'" & TransDateTime & "' from QSMS_Feeder where DID='" & Trim(DID) & "'"

Conn.Execute Str

'(2) Update DID status---used or remain Qty

'Select Case UCase(Cbostatus)
'       Case "FINISHED"
'             Str = "Update QSMS_DID set RealQty=0,UID='" & g_userName & "' where DID='" & DID & "'"
'             Conn.Execute Str
'       Case "NOTFINISHED"
'             Str = "Update QSMS_DID set UID='" & g_userName & "' where did='" & DID & "'"
'             Conn.Execute Str
'       Case Else
'End Select

'(3) Delete DID from QSMS_Feeder
Str = "delete from QSMS_Feeder where DID='" & Trim(DID) & "'"
Conn.Execute Str

Call OK_Sound
LblMessage.BackColor = &HFF00&
LblMessage = "Unlink DID  OK"

End Sub

Private Sub Form_Load()
Cbostatus.AddItem "Finished"
Cbostatus.AddItem "NotFinished"

End Sub


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

Private Sub TxtDIDFeeder_KeyPress(KeyAscii As Integer)
Dim Str As String
Dim Rs As ADODB.Recordset
Dim TransDateTime As String
Dim DID As String


If KeyAscii = 13 Then
   If InStr(1, TxtDIDFeeder, "-") > 0 Then
   DID = Trim(TxtDIDFeeder)
   TxtDID = DID
   
   ''20090914   Denver  如不用此项，需要Mark
   'Cbostatus.SetFocus
Else
   Str = "select DID from QSMS_Feeder where Feeder='" & TxtDIDFeeder & "'"
   Set Rs = Conn.Execute(Str)
    If Rs.EOF Then
       MsgBox "can not find the record,please check"
       Call Warning_Sound
       LblMessage.BackColor = &HFF&
       LblMessage = "Unlink DID  OK"
       Exit Sub
    Else
       DID = Trim(Rs!DID)
       TxtDID = DID
       ''20090914   Denver  如不用此项，需要Mark
       'Cbostatus.SetFocus
    End If
  
End If

End If
End Sub
