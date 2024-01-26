VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmCloseWO 
   Caption         =   "Close WO[20161226]"
   ClientHeight    =   8910
   ClientLeft      =   135
   ClientTop       =   375
   ClientWidth     =   10530
   FillColor       =   &H00004000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8910
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
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
      Left            =   9360
      Picture         =   "FrmCloseWO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5880
      Width           =   975
   End
   Begin VB.ListBox lstWO_SELECT 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   6960
      TabIndex        =   27
      Top             =   5040
      Width           =   2175
   End
   Begin VB.ListBox lstWOClosed 
      Enabled         =   0   'False
      Height          =   3570
      Left            =   120
      TabIndex        =   24
      Top             =   4920
      Width           =   3015
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5640
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6120
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6600
      Width           =   495
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7080
      Width           =   495
   End
   Begin VB.ListBox lstWOUnClose 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3660
      Left            =   3960
      TabIndex        =   19
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.Frame frameCHK 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Check Item"
         Height          =   1812
         Left            =   4680
         TabIndex        =   36
         Top             =   2160
         Visible         =   0   'False
         Width           =   5535
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   252
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Value           =   1  'Checked
            Width           =   252
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Check2"
            Height          =   252
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Value           =   1  'Checked
            Width           =   252
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Check3"
            Height          =   192
            Left            =   120
            TabIndex        =   38
            Top             =   1080
            Value           =   1  'Checked
            Width           =   252
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Check4"
            Height          =   192
            Left            =   120
            TabIndex        =   37
            Top             =   1440
            Value           =   1  'Checked
            Width           =   204
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Check if dispatch finished"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   195
            Left            =   600
            TabIndex        =   44
            Top             =   360
            Width           =   2205
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Check if AOI finished"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   43
            Top             =   720
            Width           =   1800
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Check if all has sent SAP1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   42
            Top             =   1080
            Width           =   2265
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Check if wo has sent SAP2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   195
            Index           =   2
            Left            =   600
            TabIndex        =   41
            Top             =   1440
            Width           =   2310
         End
      End
      Begin VB.Frame FraSB 
         Caption         =   "Small Board WO"
         Height          =   615
         Left            =   6600
         TabIndex        =   33
         Top             =   1440
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1695
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
         Left            =   1680
         TabIndex        =   9
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox TxtWO 
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
         Left            =   1680
         TabIndex        =   8
         Top             =   2160
         Width           =   2655
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
         Height          =   336
         Left            =   6840
         TabIndex        =   7
         Top             =   600
         Width           =   2655
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
         Left            =   3600
         Picture         =   "FrmCloseWO.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton OptRelease 
         Caption         =   "Release"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "Group"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1695
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
         Left            =   1680
         TabIndex        =   3
         Top             =   2640
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
         Left            =   1680
         TabIndex        =   2
         Top             =   3600
         Width           =   2655
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
         Height          =   336
         Left            =   6840
         TabIndex        =   1
         Top             =   1080
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
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
         Format          =   127926275
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   30
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
         Format          =   127926275
         CurrentDate     =   36482
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
         TabIndex        =   31
         Top             =   1080
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
         Left            =   120
         TabIndex        =   18
         Top             =   3120
         Width           =   1455
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
         TabIndex        =   17
         Top             =   2160
         Width           =   1455
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
         Height          =   372
         Index           =   1
         Left            =   4680
         TabIndex        =   16
         Top             =   600
         Width           =   2172
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
         Left            =   120
         TabIndex        =   15
         Top             =   2640
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
         TabIndex        =   14
         Top             =   1560
         Width           =   1455
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
         TabIndex        =   13
         Top             =   3600
         Width           =   1455
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
         Height          =   372
         Index           =   0
         Left            =   4680
         TabIndex        =   12
         Top             =   1080
         Width           =   2172
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Begin Date"
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
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Label LblMessage 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   120
      TabIndex        =   35
      Top             =   9000
      Width           =   9015
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "WO Will Close by manual"
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
      Index           =   5
      Left            =   6960
      TabIndex        =   28
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Wo Closed "
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
      Left            =   120
      TabIndex        =   26
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "WO Unclose"
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
      Index           =   4
      Left            =   3960
      TabIndex        =   25
      Top             =   4440
      Width           =   2175
   End
End
Attribute VB_Name = "FrmCloseWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CboGroupID_Click()
Call GetGroupWO(CboGroupID)
End Sub

Private Sub CboGroupID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboGroupID_Click
   Call DeleteWOQSMS(CboGroupID)       ''(0009)
End If
End Sub

Private Sub CboWo_Click()
TxtWO = Trim(CboWo)
Call GetWoinfo(TxtWO)
Call GetSBWO(TxtWO)
End Sub

Private Sub cboWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboWo_Click
End If
End Sub


Private Sub Check1_Click()

If Check1.Value = 0 Then
    If MsgBox("Are you sure to Un-check dispatch ?!", vbYesNo) = vbNo Then
        Check1.Value = 1
        Exit Sub
    End If
End If

If Check1.Value = 1 Then
Label2.ForeColor = &H4000&
Else
Label2.ForeColor = &HFF&

End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 0 Then
    If MsgBox("Are you sure to Un-check AOI Qty ?!", vbYesNo) = vbNo Then
        Check2.Value = 1
        Exit Sub
    End If
End If

If Check2.Value = 1 Then
Label5(0).ForeColor = &H4000&
Else
Label5(0).ForeColor = &HFF&
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 0 Then
    If MsgBox("Are you sure to Un-check if has been sent SAP1?!", vbYesNo) = vbNo Then
        Check3.Value = 1
        Exit Sub
    End If
End If

If Check3.Value = 1 Then
Label5(1).ForeColor = &H4000&
Else
Label5(1).ForeColor = &HFF&
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 0 Then
    If MsgBox("Are you sure to Un-check if has been sent SAP2 ?!", vbYesNo) = vbNo Then
        Check4.Value = 1
        Exit Sub
    End If
End If

If Check4.Value = 1 Then
Label5(2).ForeColor = &H4000&
Else
Label5(2).ForeColor = &HFF&
End If
End Sub

Private Sub CmdADD_Click()
    Dim Pointer As Long
    If lstWOUnClose.ListCount <= 0 Then Exit Sub
    If lstWOUnClose.ListIndex < 0 Then Exit Sub
    Pointer = lstWOUnClose.ListIndex
    lstWO_SELECT.AddItem Trim(lstWOUnClose.Text)
    lstWOUnClose.RemoveItem Pointer
    If lstWOUnClose.ListCount <> Pointer Then
       lstWOUnClose.ListIndex = Pointer
    End If
    
End Sub

Private Sub cmdADDALL_Click()
   
    If lstWOUnClose.ListCount <= 0 Then Exit Sub
    
    Do While lstWOUnClose.ListCount > 0
      lstWOUnClose.ListIndex = 0
      lstWO_SELECT.AddItem Trim(lstWOUnClose.Text)
      lstWOUnClose.RemoveItem 0
    Loop
    
End Sub

Private Sub CmdConfirm_Click()

 CmdConfirm.Enabled = False ''1246
 

 If lstWO_SELECT.ListCount <= 0 Then Exit Sub
    Do While lstWO_SELECT.ListCount > 0
        lstWO_SELECT.ListIndex = 0
           If MsgBox("do you make sure to close the work order by manual " & lstWO_SELECT.Text, vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
              If CloseWoByManual(Trim(lstWO_SELECT.Text), "Manual") = True Then
                 lstWOClosed.AddItem Trim(lstWO_SELECT.Text)
                 lstWO_SELECT.RemoveItem 0
              End If
           Else
              ''Exit Sub
              If BU = "NB5" Then
                  CmdConfirm.Enabled = True
                  Exit Sub
              Else
                  Exit Sub
              End If
           End If

    Loop
    
 CmdConfirm.Enabled = True ''1246


End Sub

Private Sub cmdDel_Click()
    Dim Pointer As Long
    If lstWO_SELECT.ListCount <= 0 Then Exit Sub
    If lstWO_SELECT.ListIndex < 0 Then Exit Sub
    Pointer = lstWO_SELECT.ListIndex
    
        lstWOUnClose.AddItem Trim(lstWO_SELECT.Text)
        lstWO_SELECT.RemoveItem Pointer
        If lstWO_SELECT.ListCount <> Pointer Then
           lstWO_SELECT.ListIndex = Pointer
        End If
   
End Sub

Private Sub cmdDELALL_Click()
    If lstWO_SELECT.ListCount <= 0 Then Exit Sub
    Do While lstWO_SELECT.ListCount > 0
        lstWO_SELECT.ListIndex = 0
       
           lstWOUnClose.AddItem Trim(lstWO_SELECT.Text)
           lstWO_SELECT.RemoveItem 0
       
    Loop
    
End Sub

Private Sub CmdQuery_Click()
If Trim(CboLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If
Call GetGroupID
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
Dim str As String
Dim rs As ADODB.Recordset
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
str = "select getdate()"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
    'Date = Rs(0)
    'Time = Rs(0)
End If
dtpSDate = Date
dtpEDate = Date
Call GetLine
Call GetAuthority

''20100903  Kyle Added to solve the encoding problem of UI
If StrBU = "PO" Then
    Label2.Caption = "Check if dispatch finished"
    Label5(0).Caption = "Check if AOI finished"
    Label5(1).Caption = "Check if all has sent SAP1"
    Label5(2).Caption = "Check if wo has sent SAP2"
End If
End Sub
Private Function GetLine()
Dim str As String
Dim rs As ADODB.Recordset
str = "select distinct Line from QSMS_woGroup"
Set rs = Conn.Execute(str)
CboLine.Clear
While Not rs.EOF
    CboLine.AddItem rs!Line
    rs.MoveNext
Wend
End Function

'Private Function GetAuthority()
'Dim str As String
'Dim rs As ADODB.Recordset
'str = "select * from userright where username='" & g_userName & "' AND USERRIGHT='PowerCloseWO'"
'Set rs = Conn.Execute(str)
'If rs.EOF = False Then
'    frameCHK.Visible = True
'End If
'End Function


''20090910   Denver  Add UnChkDispCloseWO function. it need not query DB
Private Sub GetAuthority()
Dim i As Integer
    For i = LBound(g_userRight) To UBound(g_userRight)
        If UCase(g_userRight(i)) = UCase("PowerCloseWO") Then
            frameCHK.Visible = True
            Exit Sub
        End If
        
        If UCase(g_userRight(i)) = UCase("UnChkDispCloseWO") Then
            frameCHK.Visible = True
            Check2.Enabled = False
            Check3.Enabled = False
            Check4.Enabled = False
        End If
           
    Next i
 
End Sub


Private Function GetGroupID()
Dim str As String
Dim BeginDate, EndDate As String
Dim GroupIDHead As String
Dim i As Long
Dim rs As ADODB.Recordset
BeginDate = Format(dtpSDate, "YYYY/MM/DD")
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
EndDate = Format(dtpEDate, "YYYY/MM/DD")
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")
'GroupIDHead = Trim(CboLine) & TransDate
If BU = "NB5" Then
    If OptRelease.Value = True Then
       str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )"
    Else
        str = "select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )"
    End If
Else
    If OptRelease.Value = True Then
       str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N'"
    Else
        str = "select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N'"
    End If
End If
Set rs = Conn.Execute(str)
i = 0
CboGroupID.Clear
While Not rs.EOF
      CboGroupID.AddItem Trim(rs!GroupID)
      rs.MoveNext
      i = i + 1
Wend
If i = 0 Then
   MsgBox "No data"
   
End If
End Function


Private Function GetGroupWO(ByVal GroupID As String)
Dim str As String
Dim TransDate As String
Dim rs As ADODB.Recordset
Dim wostr(1000) As String
Dim i, j As Integer
i = 1
For i = 1 To 1000
     wostr(i) = ""
Next i
str = "select Work_Order,ClosedFlag from QSMS_WOGroup  where GroupID= '" & GroupID & "'"

Set rs = Conn.Execute(str)
lstWOClosed.Clear
lstWOUnClose.Clear
lstWO_SELECT.Clear
CboWo.Clear
While Not rs.EOF
          If ChkMBWo(rs!Work_Order) = True Then
                
                If ChkIfSmallBoardExist(rs!Work_Order) = False Then
                    CboWo.AddItem Trim(rs!Work_Order)
                    If UCase(Trim(rs!ClosedFlag)) = "Y" Then
                       lstWOClosed.AddItem Trim(rs!Work_Order)
                    Else
                       lstWOUnClose.AddItem Trim(rs!Work_Order)
                    End If
                End If
          End If
      rs.MoveNext
Wend
End Function
Private Function ChkIfSmallBoardExist(ByVal WO As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
Dim i As Integer
str = ""
ChkIfSmallBoardExist = False

str = "select WO from Sap_Wo_list where wo='" & WO & "' and (PN like '21%' or PN like '31%')"
Set rs = Conn.Execute(str)
If rs.EOF Then
  
    str = "Select Wo from sap_wo_list where [group] in (select [group] from sap_wo_list where wo='" & WO & "')"
    Set rs = Conn.Execute(str)
    While Not rs.EOF
             For i = 1 To lstWOUnClose.ListCount
                lstWOUnClose.ListIndex = i - 1
                If Trim(lstWOUnClose.Text) = Trim(rs!WO) Then
                    ChkIfSmallBoardExist = True
                End If
            Next i
          rs.MoveNext
    Wend
End If


End Function


Private Function GetWoinfo(ByVal WO As String)
Dim str As String
Dim rs As ADODB.Recordset
str = "select PN, Qty from Sap_Wo_List where WO='" & Trim(WO) & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   TxtMBPN = rs!PN
   TxtWOQty = rs!Qty
End If
str = "select Customer from ModelName where PN='" & TxtMBPN & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   TxtCustomer = Trim(rs!Customer)
End If
End Function
Private Function GetSBWO(ByVal WO As String)
Dim str As String
Dim rs As ADODB.Recordset
Dim i As Long
Dim Group As String
i = 0
CboSBWO.Clear
FraSB.Visible = False
str = "Select [Group] from Sap_Wo_List where wo='" & WO & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   Group = Trim(rs!Group)
'   TxtGroup = Group
End If
str = "select Wo from Sap_Wo_list where [Group] ='" & Group & "' and wo<>'" & WO & "' order by wo"
Set rs = Conn.Execute(str)
While Not rs.EOF
     CboSBWO.AddItem Trim(rs!WO)
     rs.MoveNext
     i = i + 1
Wend
If i > 0 Then
    FraSB.Visible = True

End If
End Function

Private Function DeleteWOQSMS(ByVal GroupID As String)
Dim strSQL As String, strWO As String
Dim rs As ADODB.Recordset
    strSQL = "select distinct Work_Order from QSMS_WO where Work_Order in(select distinct WorkOrder from SMT_WO_Del where sSN not like 'Reduce%') and Work_Order not in(select WO from sap_Wo_List) and Work_Order in(select Work_Order from QSMS_WOGroup where GroupID='" & Trim(GroupID) & "')"
    Set rs = Conn.Execute(strSQL)
    If rs.EOF = False Then
        While Not rs.EOF
            strWO = Trim(rs!Work_Order) & ";" & strWO
            rs.MoveNext
        Wend
        If MsgBox("The SMT part have been deleted of the WO:" & strWO & ",do you want to delete the QSMS part of the WO?", vbYesNo, "Message") = vbYes Then
             While Not rs.EOF
                strSQL = "EXEC QSMS_DelWO '" & Trim(rs!Work_Order) & "'"
                Conn.Execute (strSQL)
                rs.MoveNext
            Wend
        End If
    End If
End Function
