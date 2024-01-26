VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmChangeDIDslot 
   Caption         =   "FrmTransferDispatchedDID"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15645
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   15645
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Change Extra DID Slot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   360
      TabIndex        =   32
      Top             =   3480
      Width           =   11415
      Begin VB.TextBox txtQty 
         Height          =   375
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox cbooldslot 
         Height          =   315
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H0000FF00&
         Caption         =   "&Change Slot"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2160
         Width           =   1455
      End
      Begin VB.ComboBox CboDID 
         Height          =   315
         Left            =   1560
         TabIndex        =   39
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox cboNewslot 
         Height          =   315
         Left            =   7800
         TabIndex        =   38
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox CboMachine 
         Height          =   315
         Left            =   1560
         TabIndex        =   37
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Extra DID Qty"
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
         TabIndex        =   46
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "to"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   7440
         TabIndex        =   45
         Top             =   1320
         Width           =   165
      End
      Begin VB.Label Label2 
         Caption         =   "Change slot from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   3960
         TabIndex        =   44
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Detination Slot"
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
         Index           =   8
         Left            =   7800
         TabIndex        =   36
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Source Slot"
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
         Index           =   7
         Left            =   5520
         TabIndex        =   35
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Extra DID"
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
         TabIndex        =   34
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Machine"
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
         TabIndex        =   33
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin MSComCtl2.DTPicker DTPedate 
         Height          =   375
         Left            =   1680
         TabIndex        =   43
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   69730305
         CurrentDate     =   39062
      End
      Begin MSComCtl2.DTPicker DTPsdate 
         Height          =   375
         Left            =   1680
         TabIndex        =   42
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   69730305
         CurrentDate     =   39062
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
         TabIndex        =   19
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
         Left            =   8640
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2520
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
         Left            =   11640
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2520
         Width           =   735
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "Group"
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton OptRelease 
         Caption         =   "Release"
         Height          =   255
         Left            =   120
         TabIndex        =   15
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
         Picture         =   "FrmChangeDIDslot.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
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
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   120
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2295
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
         Left            =   4680
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Frame FraSB 
         Caption         =   "Small Board WO"
         Height          =   615
         Left            =   6600
         TabIndex        =   9
         Top             =   1320
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
            TabIndex        =   10
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
         Left            =   13800
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton cmdDELALL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdDEL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdADDALL 
         BackColor       =   &H00C0C0C0&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdADD 
         BackColor       =   &H00C0C0C0&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   480
         Width           =   495
      End
      Begin VB.ListBox ListWoNotFinish 
         Height          =   1425
         ItemData        =   "FrmChangeDIDslot.frx":0442
         Left            =   10320
         List            =   "FrmChangeDIDslot.frx":0444
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   2175
      End
      Begin VB.ListBox ListWoDispatching 
         Height          =   1230
         ItemData        =   "FrmChangeDIDslot.frx":0446
         Left            =   13200
         List            =   "FrmChangeDIDslot.frx":0448
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1695
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
         Left            =   7560
         TabIndex        =   29
         Top             =   2520
         Width           =   1095
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
         Left            =   10920
         TabIndex        =   27
         Top             =   2520
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
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Not Finished Work Order"
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
         Height          =   495
         Index           =   2
         Left            =   10320
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   2520
         Width           =   1455
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
         Left            =   3960
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
         Left            =   12480
         TabIndex        =   22
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Dispatching WO"
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
         Height          =   495
         Index           =   3
         Left            =   13200
         TabIndex        =   21
         Top             =   120
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
   End
End
Attribute VB_Name = "FrmChangeDIDslot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CboDID_Click()
Dim Str As String
Dim Rs As ADODB.Recordset

Str = "select a.slot,a.didqty from qsms_dispatch a,qsms_did b where a.work_order='" & cboWO & "' and a.machine='" & CboMachine & "' and a.DID='" & CboDID & "' and a.did=b.did and a.diddatetime=b.transdatetime"

Set Rs = Conn.Execute(Str)
cbooldslot.Clear
If Rs.EOF Then
    MsgBox "Can't find the DID information!", vbCritical
    Exit Sub
Else
    cbooldslot.Text = Rs!Slot
    txtQty = Rs!DIDQty
End If
'for NB1, modify substring(rtrim('" & CboDID & "'),1,11)+'%' to '%'+substring(rtrim('" & CboDID & "'),2,11)+'%'
Str = "select Distinct slot from qsms_wo where work_order='" & cboWO & "' and machine='" & CboMachine & "' and balanceqty<0 and comppn like '%'+substring(rtrim('" & CboDID & "'),1,11)+'%'"
Set Rs = Conn.Execute(Str)
cboNewslot.Clear
If Rs.EOF Then
    cboNewslot.Locked = True
    MsgBox "Can't find the slot to be transfer!"
Else
    cboNewslot.Locked = False
    Do While Not Rs.EOF
        cboNewslot.AddItem Rs!Slot
        Rs.MoveNext
    Loop
End If
End Sub

Private Sub CboGroupID_Click()
Call GetGroupWO(CboGroupID)
End Sub





Private Sub CboMachine_Click()
Dim Str As String, tempSlot As String
Dim tempqty As Long
Dim Rs As ADODB.Recordset
Dim TempRs As ADODB.Recordset

CboDID.Clear
Str = "select slot,balanceqty from qsms_wo where work_order='" & cboWO & "' and machine='" & CboMachine & "'and balanceqty>0 group by slot,balanceqty"
Set Rs = Conn.Execute(Str)
Do While Not Rs.EOF
    tempSlot = Rs!Slot
    tempqty = Rs!BalanceQty
    'txtQty = tempQTY
    Str = "select DID from qsms_dispatch where work_order='" & cboWO & "' and machine='" & CboMachine & "' and  slot='" & tempSlot & "' and didqty=" & tempqty & " and deletedflag<>'Y'"
    Set TempRs = Conn.Execute(Str)
    
    If TempRs.EOF = False Then
        
            CboDID.AddItem TempRs!DID

    Else
        Str = "select DID from qsms_dispatch where work_order='" & cboWO & "' and machine='" & CboMachine & "' and  slot='" & tempSlot & "' and didqty<" & tempqty & " and deletedflag<>'Y'"
        Set TempRs = Conn.Execute(Str)
        Do While Not TempRs.EOF
            CboDID.AddItem TempRs!DID
            TempRs.MoveNext
        Loop
    
    End If
    Rs.MoveNext
Loop
txtQty = ""
cboNewslot.Clear
cbooldslot.Clear
            
End Sub



Private Sub CboWo_Click()
TxtWO = Trim(cboWO)
Call GetWoinfo(TxtWO)
Call GetMachine(TxtWO)

End Sub

Private Sub cmdOK_Click()
Dim Str As String
Dim Rs As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim tempitem As String
Dim tempitemNew As String
On Error GoTo EcmdSave_Click
If CboMachine = "" Or cbooldslot = "" Or cboNewslot = "" Or CboDID = "" Then
MsgBox "Please input the machine & slot infomation", vbCritical
Exit Sub
End If

  Str = "Update qsms_dispatch set slot='" & cboNewslot & "' from qsms_dispatch a,qsms_did b where a.work_order='" & cboWO & "' and a.machine='" & CboMachine & "' and a.DID='" & CboDID & "' and a.DIDqty=" & txtQty & " and a.did=b.did and a.diddatetime=b.transdatetime "
  Conn.Execute (Str)

  Str = "Update qsms_wo set dispatchqty=dispatchqty+" & txtQty & ",balanceqty=dispatchqty-needqty+" & txtQty & " where work_order='" & cboWO & "' and machine='" & CboMachine & "' and slot='" & cboNewslot & "'"
  Conn.Execute (Str)
  
  Str = "Update qsms_wo set dispatchqty=dispatchqty-" & txtQty & ",balanceqty=balanceqty-" & txtQty & " where work_order='" & cboWO & "' and machine='" & CboMachine & "' and slot='" & cbooldslot & "'"
  Conn.Execute (Str)
    
    Call UpdateMachineFlagByWO(TxtWO)
    cboNewslot.Clear
    cbooldslot.Clear
    CboDID.Clear
    CboMachine.Clear
    txtQty = ""
    Call CboWo_Click

    MsgBox "OK ! "
    Exit Sub
EcmdSave_Click:
    MsgBox Err.Description + ",Please contact QSMS SMT Staff"
End Sub

Private Sub CmdQuery_Click()
If Trim(CboLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If
Call GetGroupID("")
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
DTPsdate = Date
DTPedate = Date
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

Private Function GetGroupID(ByVal Jobpn As String)
Dim Str As String
Dim BeginDate, EndDate As String
Dim GroupIDHead As String
Dim i As Long
Dim Rs As ADODB.Recordset
Dim TempJobPn As String
BeginDate = Format(DTPsdate, "YYYY/MM/DD")
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
EndDate = Format(DTPedate, "YYYY/MM/DD")
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")

Str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "'"
Set Rs = Conn.Execute(Str)

CboGroupID.Clear
If Rs.EOF Then MsgBox "No data"
While Not Rs.EOF
      CboGroupID.AddItem Trim(Rs!GroupID)
      Rs.MoveNext
Wend
End Function

Private Function GetWoinfo(ByVal WO As String)
Dim Str As String
Dim Rs As ADODB.Recordset
Str = "select PN, Qty ,MB_Rev,Line from Sap_Wo_List where WO='" & Trim(WO) & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   TxtMBPN = Rs!PN
   TxtWOQty = Rs!Qty
   'TxtRev = Trim(Rs!Mb_Rev)
   TxtModel = Rs!PN + "-" + Trim(Rs!Mb_Rev)
End If
'Str = "select Customer from ModelName where PN='" & TxtMBPN & "'"
'Set Rs = Conn.Execute(Str)
'If Not Rs.EOF Then
'   TxtCustomer = Trim(Rs!Customer)
'End If

End Function

Private Function GetMachine(ByVal WO As String)
Dim Str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset

Str = "select distinct machine from qsms_wo where work_order='" & TxtWO & "' and balanceqty>0"
Set Rs = Conn.Execute(Str)
If Rs.EOF Then
    MsgBox "No extra DID to change slot!", vbInformation + vbOKOnly
    Exit Function
End If

CboMachine.Clear

While Not Rs.EOF
    CboMachine.AddItem Trim(Rs!Machine)
    Rs.MoveNext
Wend
CboDID.Clear
cbooldslot.Clear
cboNewslot.Clear
txtQty = ""

End Function

Private Function GetGroupWO(ByVal GroupID As String)
Dim Str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset
cboWO.Clear
Str = "select distinct work_order from qsms_wogroup where groupid='" & CboGroupID & "'"
Set Rs = Conn.Execute(Str)
While Not Rs.EOF
          
        cboWO.AddItem Trim(Rs!Work_Order)
        Rs.MoveNext
Wend
End Function



