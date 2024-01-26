VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FrmDispatchDIDAdditional 
   BackColor       =   &H0000C000&
   Caption         =   "DispatchDIDAdditional[2011/03/01]"
   ClientHeight    =   10620
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10620
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      Caption         =   "ChkDID Status"
      Height          =   1815
      Left            =   15480
      TabIndex        =   94
      Top             =   3840
      Visible         =   0   'False
      Width           =   3015
      Begin VB.ListBox ListDIDStatus 
         Height          =   1035
         ItemData        =   "FrmDispatchDIDAdditional.frx":0000
         Left            =   120
         List            =   "FrmDispatchDIDAdditional.frx":0007
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox TxtChkDID 
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
         Left            =   120
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      Height          =   2535
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   15375
      Begin VB.TextBox TxtBuildType 
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
         Left            =   10680
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2295
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
         Style           =   2  'Dropdown List
         TabIndex        =   76
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
         Left            =   10680
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   720
         Width           =   2295
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
         Left            =   14040
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "Group"
         Height          =   255
         Left            =   1680
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton OptRelease 
         Caption         =   "Release"
         Height          =   255
         Left            =   120
         TabIndex        =   72
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
         Picture         =   "FrmDispatchDIDAdditional.frx":001A
         Style           =   1  'Graphical
         TabIndex        =   71
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
         Style           =   2  'Dropdown List
         TabIndex        =   70
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
         TabIndex        =   69
         TabStop         =   0   'False
         Text            =   "CboNotFinishedWO"
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
         Left            =   10680
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   14040
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
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
         Left            =   14040
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   720
         Width           =   975
      End
      Begin VB.Frame FraSB 
         Caption         =   "Small Board WO"
         Height          =   615
         Left            =   6600
         TabIndex        =   64
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
            Style           =   2  'Dropdown List
            TabIndex        =   65
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
         Left            =   10680
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   62
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
         Style           =   2  'Dropdown List
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   480
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   77
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
         Format          =   130023427
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   78
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
         Format          =   130023427
         CurrentDate     =   36482
      End
      Begin MCI.MMControl wave_control 
         Height          =   450
         Left            =   -3840
         TabIndex        =   98
         Top             =   120
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   794
         _Version        =   393216
         PlayEnabled     =   -1  'True
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Build Type"
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
         Index           =   27
         Left            =   9240
         TabIndex        =   100
         Top             =   1680
         Width           =   1455
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
         TabIndex        =   91
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
         TabIndex        =   90
         Top             =   960
         Width           =   2295
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
         Left            =   9240
         TabIndex        =   89
         Top             =   720
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
         TabIndex        =   88
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
         Left            =   12960
         TabIndex        =   87
         Top             =   1200
         Width           =   1095
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
         TabIndex        =   86
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
         Left            =   9240
         TabIndex        =   85
         Top             =   240
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
         Left            =   12960
         TabIndex        =   84
         Top             =   240
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
         Left            =   12960
         TabIndex        =   83
         Top             =   720
         Width           =   1095
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
         Left            =   9240
         TabIndex        =   82
         Top             =   1200
         Width           =   1455
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
         TabIndex        =   81
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
         TabIndex        =   80
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
         TabIndex        =   79
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "PD prepair material"
      Height          =   6015
      Left            =   0
      TabIndex        =   12
      Top             =   2640
      Width           =   15375
      Begin VB.ComboBox CboJob 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4800
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   120
         Width           =   2775
      End
      Begin VB.ComboBox CboMathineNOK 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1560
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox CboMathineOK 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   10200
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox TxtMachine 
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
         Left            =   13440
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame FraWithoutDID 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Without DID"
         Height          =   735
         Left            =   120
         TabIndex        =   43
         Top             =   720
         Width           =   15255
         Begin VB.ComboBox CboWithout 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1800
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   240
            Width           =   4215
         End
         Begin VB.TextBox TxtWBalance 
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
            Left            =   9000
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtWTotal 
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
            Left            =   6840
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackColor       =   &H000000FF&
            Caption         =   "Without DID "
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
            Index           =   18
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Balance"
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
            Index           =   12
            Left            =   8040
            TabIndex        =   48
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Total"
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
            Left            =   6120
            TabIndex        =   47
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame FraDispatchDID 
         BackColor       =   &H00FF80FF&
         Caption         =   "DispatchDID"
         Height          =   1695
         Left            =   120
         TabIndex        =   24
         Top             =   3480
         Width           =   15015
         Begin VB.ComboBox CboLR 
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
            Left            =   8760
            Style           =   2  'Dropdown List
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TxtDID 
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
            Left            =   960
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   240
            Width           =   3255
         End
         Begin VB.TextBox TxtDispatchQty 
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
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox TxtDIDTotalQty 
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
            Left            =   7440
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox TxtCompBaseQty 
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
            Left            =   13080
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   840
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TxtDIDRemainQty 
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
            Left            =   10320
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox TxtConsumedQty 
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
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox TxtNeedQty 
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
            Left            =   1800
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1335
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
            Left            =   13320
            Picture         =   "FrmDispatchDIDAdditional.frx":045C
            Style           =   1  'Graphical
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox Cboslot 
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
            Left            =   5880
            TabIndex        =   25
            TabStop         =   0   'False
            Text            =   "Cboslot"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label LblSlot 
            BackColor       =   &H0000FF00&
            Caption         =   "LR"
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
            Left            =   7200
            TabIndex        =   93
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackColor       =   &H0000FF00&
            Caption         =   "DID"
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
            Index           =   14
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H0000FF00&
            Caption         =   "Dispatch Qty"
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
            Left            =   10080
            TabIndex        =   41
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "DID Total Qty"
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
            Index           =   17
            Left            =   6000
            TabIndex        =   40
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Comp Base Qty"
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
            Left            =   11400
            TabIndex        =   39
            Top             =   720
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "DID Remain Qty"
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
            Left            =   8520
            TabIndex        =   38
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Consumed Qty"
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
            Index           =   6
            Left            =   3120
            TabIndex        =   37
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Need Qty"
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
            TabIndex        =   36
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label LblMessage 
            BackColor       =   &H00FFFFC0&
            Height          =   495
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   10215
         End
         Begin VB.Label LblSlot 
            BackColor       =   &H0000FF00&
            Caption         =   "Slot"
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
            Left            =   4320
            TabIndex        =   34
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000000&
         Caption         =   "DID information "
         Height          =   735
         Left            =   0
         TabIndex        =   13
         Top             =   5160
         Width           =   15015
         Begin VB.TextBox TxtLotCode 
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
            Left            =   10200
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox TxtDateCode 
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
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox TxtVendorCode 
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
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox TxtCompPN 
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
            Left            =   1200
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   240
            Width           =   2295
         End
         Begin VB.TextBox TxtRackID 
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
            Left            =   12840
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Lot  Code"
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
            Index           =   11
            Left            =   9000
            TabIndex        =   23
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Date Code"
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
            Index           =   10
            Left            =   6360
            TabIndex        =   22
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Vendor Code"
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
            Index           =   9
            Left            =   3480
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Comp PN"
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
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "RackID"
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
            Index           =   19
            Left            =   11640
            TabIndex        =   19
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSDataGridLib.DataGrid DGAVL 
         Height          =   1815
         Left            =   4800
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3201
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
         Caption         =   "Vendor not approved"
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
      Begin MSDataGridLib.DataGrid DGDIDNotOK 
         Height          =   1815
         Left            =   9840
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3201
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
         Caption         =   "DID Not OK"
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
      Begin MSDataGridLib.DataGrid DGDIDOK 
         Height          =   1815
         Left            =   7320
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3201
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
         Caption         =   "DID OK"
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
      Begin MSDataGridLib.DataGrid DGCompNotOK 
         Height          =   1815
         Left            =   240
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3201
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
         Caption         =   "Component  Not OK"
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
      Begin MSDataGridLib.DataGrid DGDIDAvailable 
         Height          =   1815
         Left            =   12600
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3201
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
         Caption         =   "DID Available"
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
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "JobPN"
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
         Left            =   3720
         TabIndex        =   102
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Machine NOK"
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
         TabIndex        =   59
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Machine OK"
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
         Left            =   8880
         TabIndex        =   58
         Top             =   240
         Width           =   1335
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
         Index           =   15
         Left            =   12360
         TabIndex        =   57
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame FraQyeryDID 
      BackColor       =   &H00FFC0C0&
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   8760
      Width           =   15375
      Begin VB.Frame Frame4 
         Caption         =   "QueryWoBy DID"
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   7455
         Begin VB.TextBox TxtQryDID 
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
            Left            =   2040
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   360
            Width           =   3255
         End
         Begin VB.CommandButton cmdexcel 
            Caption         =   "&Excel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5400
            Picture         =   "FrmDispatchDIDAdditional.frx":0766
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
         Begin VB.CheckBox ChkAll 
            Caption         =   "By DID"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label4 
            BackColor       =   &H0000FF00&
            Caption         =   "Comp"
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
            Index           =   20
            Left            =   1080
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "QueryDIDByWO Machine"
         Height          =   855
         Left            =   7680
         TabIndex        =   1
         Top             =   120
         Width           =   7335
         Begin VB.CommandButton CmdExcelDID 
            Caption         =   "&Excel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   5760
            Picture         =   "FrmDispatchDIDAdditional.frx":0A70
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox TxtQryWO 
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
            Left            =   1080
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox TxtQryMachine 
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
            Left            =   3840
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackColor       =   &H0000FF00&
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
            Index           =   23
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label4 
            BackColor       =   &H0000FF00&
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
            Index           =   24
            Left            =   2880
            TabIndex        =   5
            Top             =   360
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "FrmDispatchDIDAdditional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/**********************************************************************************
'**文 件 名: FrmDispatchDIDAdditional.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: Jeanson
'**日    期: 2007.10.01
'**描    述: Dispatch DID Additionally
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**Jeanson      2007.10.31     Set cboLine order by line asc (0001)
'**Jeanson      2007.10.31     Modify auto-set LR function (0002)
'**Jing         2007.10.31     Add the message in detail in the dispatch interface for deleted DID (0003)
'**Udall        2007.11.05     Add check DID,the DID can't be dispatched to different line and side --------(0004)
'**Scofield     2010.04.18     Add checking IPQC before dispatch in MBU (0005)
'***********************************************************************************/
Dim DIDDateTime As String
Private Sub CboGroupID_Click()
Call GetGroupWO(CboGroupID)
End Sub

Private Sub CboGroupID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboGroupID_Click
End If
End Sub

Private Sub CboJob_Click()
   Call GetDID(Trim(TxtMachine), Trim(CboJob), TxtWO, CboLine)
End Sub

Private Sub CboJob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CboJob_Click
End If
End Sub

Private Sub CboLR_Click()
'    Call GetCompDispInfo(Trim(TxtWO), Trim(CboJob), Trim(TxtMachine), Trim(TxtCompPN), Trim(Cboslot), Trim(CboLR)) '(0002)
End Sub

Private Sub CboLR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CboLR_Click
End If
End Sub

Private Sub CboMathineNOK_Click()
TxtMachine = Trim(CboMathineNOK)
If ChkWOSeq(Trim(TxtWO), Trim(TxtMachine)) = False Then
   txtDID.Enabled = False
   Exit Sub
Else
   txtDID.Enabled = True
End If
Call GetJobForBuildType(Trim(TxtWO), Trim(TxtMachine), Trim(TxtBuildType))
Call GetDID(Trim(TxtMachine), Trim(CboJob), TxtWO, CboLine)
Call GetCompPnWithoutDID(Trim(TxtMachine), Trim(TxtWO))
txtDID.Text = ""
txtDID.SetFocus
CboMathineOK.Text = ""
End Sub

Private Sub CboMathineNOK_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboMathineNOK_Click
End If
End Sub

Private Sub CboMathineOK_Click()
TxtMachine = Trim(CboMathineOK)
If ChkWOSeq(Trim(TxtWO), Trim(TxtMachine)) = False Then
   txtDID.Enabled = False
   Exit Sub
Else
   txtDID.Enabled = True
End If
Call GetJobForBuildType(Trim(TxtWO), Trim(TxtMachine), Trim(TxtBuildType))
Call GetDID(Trim(TxtMachine), Trim(TxtMBPN), TxtWO, CboLine)
CboWithout.Clear
txtDID.Text = ""
txtDID.SetFocus
CboMathineNOK.Text = ""
End Sub

Private Sub CboMathineOK_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboMathineOK_Click
End If
End Sub

Private Sub CboNotChkBOM_Click()
TxtWO = Trim(CboNotChkBOM)

Call GetSBWO(TxtWO)
Call GetWoinfo(TxtWO)
End Sub

Private Sub CboNotFinishedWO_Click()
TxtWO = Trim(CboNotFinishedWO)
Call GetSBWO(TxtWO)
Call GetWoinfo(TxtWO)
Call GetMachine(TxtWO)

End Sub

Private Sub CboNotFinishedWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboNotFinishedWO_Click
End If
End Sub
'
'Private Sub CboSlot_Click()
' Call GetCompDispInfo(Trim(TxtWO), Trim(TxtMachine), Trim(TxtCompPN))
'End Sub
'
'Private Sub CboSlot_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 or KeyAscii = 9 Then
'   CboSlot_Click
'End If
'End Sub

Private Sub CboSBWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TxtWO = Trim(CboNotFinishedWO)
  
   Call GetWoinfo(TxtWO)
   Call GetMachine(TxtWO)
End If

End Sub

Private Sub Cboslot_Click()
   Call GetCompDispInfo(Trim(TxtWO), Trim(CboJob), Trim(TxtMachine), Trim(TxtCompPN), Trim(Cboslot), "")
End Sub

Private Sub Cboslot_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call Cboslot_Click
End If
End Sub

Private Sub CboWithout_Click()
Dim str As String
Dim rs As ADODB.Recordset
str = "Select sum(NeedQty) as NeedQty,sum(BalanceQty) as BalanceQty from QSMS_Wo where Work_order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') and Machine='" & Trim(TxtMachine) & "' and CompPN='" & Trim(CboWithout) & "' and BalanceQty<0"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   TxtWTotal.Text = Trim(rs!NeedQty)
   TxtWBalance.Text = Trim(rs!BalanceQty)
End If
End Sub

Private Sub cboWO_Click()
TxtWO = Trim(cboWO)
Call GetSBWO(TxtWO)
Call GetWoinfo(TxtWO)
Call GetMachine(TxtWO)
End Sub

Private Sub cboWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call cboWO_Click
End If
End Sub



Private Sub CmdConfirm_Click()
On Error GoTo EcmdSave_Click

'If ChkDIDInherit(Trim(TxtWO), Trim(TxtDID)) = False Then
'    MsgBox "The DID can not inherit,Please check"
'    Exit Sub
'End If
'If ChkDIDDispatchedToWo(TxtWO, TxtDID) = False Then
'   MsgBox "The DID has dispathcd to The WO, Please check"
'   Exit Sub
'End If

If ChkErr = True Then
   If Insert_QSMS_Out = True Then
      Call OK_Sound
      LblMessage.Caption = "insert OK"
   Else
      Call Warning_Sound
      LblMessage.Caption = "didn't dispatch the DID,Please check"
   End If
Else
    Call Warning_Sound
End If
CboMathineNOK = TxtMachine
txtDID.SetFocus
Exit Sub
EcmdSave_Click:
    MsgBox Err.Description + ",Please contact QSMS SMT Staff"
End Sub



Private Sub CmdExcel_Click()
Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object
 Dim fldCount As Integer, iCol As Integer
 Dim str As String
 Dim rs As ADODB.Recordset

 Dim strFileName, Trans_Date As String

    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add

    xlApp.DisplayAlerts = False
    Set xlWs = xlApp.Worksheets(1)
  
    xlApp.UserControl = True
''(1) first Get DID infomation
'    Str = "Select * from QSMS_DID where DID='" & Trim(TxtQryDID) & "'"
'    Set Rs = Conn.Execute(Str)
'
'    fldCount = Rs.Fields.Count
'
'    For iCol = 1 To fldCount
'        xlWs.Cells(1, iCol).Value = Rs.Fields(iCol - 1).Name
'    Next
'
'    ' Check veRsion of Excel
'
'   ' If Val(Left$(xlApp.Version, 1)) > 8 Then
'
'        xlWs.Cells(2, 1).CopyFromRecordset Rs
'
'  ' End If
'(2)second Get DID dispatch information
   If ChkAll.Value = 0 Then
   
     str = "Select * from QSMS_Dispatch where CompPN='" & Trim(TxtQryDID) & "'  and work_order='" & TxtWO & "'"
   Else
      str = "Select * from QSMS_Dispatch where DID = '" & Trim(TxtQryDID) & "' "

End If
    Set rs = Conn.Execute(str)
    
    fldCount = rs.Fields.Count
  
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = rs.Fields(iCol - 1).Name
    Next
        
    ' Check veRsion of Excel
    
    'If Val(Left$(xlApp.Version, 1)) > 8 Then

        xlWs.Cells(2, 1).CopyFromRecordset rs

  ' End If
 
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlApp.Visible = True
  
    rs.Close
    Set rs = Nothing
    Trans_Date = Format(Now, "YYYYMMDD")


'    Set xlWs = Nothing
    Set xlApp = Nothing
    Set xlsBook = Nothing
End Sub

Private Sub cmdexcel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   
   Call CmdExcel_Click
End If
End Sub

Private Sub CmdExcelDID_Click()
Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object
  Dim fldCount As Integer, iCol As Integer
 Dim str As String
 Dim rs As ADODB.Recordset

 Dim strFileName, Trans_Date As String

    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add

    xlApp.DisplayAlerts = False
    Set xlWs = xlApp.Worksheets(1)
  
    xlApp.UserControl = True
'(1) first Get DID infomation
    str = "Select * from QSMS_Dispatch where Work_Order='" & Trim(TxtQryWO) & "' and machine like '" & Trim(TxtQryMachine) & "%'"
    Set rs = Conn.Execute(str)
    
    fldCount = rs.Fields.Count
  
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = rs.Fields(iCol - 1).Name
    Next
        
    ' Check veRsion of Excel
    
    

    xlWs.Cells(2, 1).CopyFromRecordset rs

 
 
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlApp.Visible = True
  
    rs.Close
    Set rs = Nothing
    Trans_Date = Format(Now, "YYYYMMDD")


'    Set xlWs = Nothing
    Set xlApp = Nothing
    Set xlsBook = Nothing
End Sub

Private Sub CmdQuery_Click()
If Trim(CboLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If
Call GetGroupID
End Sub



Private Sub DGDIDAvailable_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 On Error Resume Next
        With DGDIDAvailable
             txtDID.Text = Trim(.Columns(0).Text)
              Call GetDIDInfo(txtDID, TxtWO)
              Call GetCompDispInfo(Trim(TxtWO), Trim(CboJob), Trim(TxtMachine), Trim(TxtCompPN), "", "")
              Call GetSlot(Trim(txtDID), Trim(TxtWO))
             
             If Err.Number <> 0 Then
                txtDID.Text = vbNullString
               
             End If
        End With
End Sub

Private Sub DGDIDNotOK_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 On Error Resume Next
        With DGDIDNotOK
             txtDID.Text = Trim(.Columns(0).Text)
              Call GetDIDInfo(txtDID, TxtWO)
              Call GetCompDispInfo(Trim(TxtWO), Trim(CboJob), Trim(TxtMachine), Trim(TxtCompPN), "", "")
              Call GetSlot(Trim(txtDID), Trim(TxtWO))
             
             If Err.Number <> 0 Then
                txtDID.Text = vbNullString
               
             End If
        End With
        
        
End Sub

Private Sub DGDIDOK_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 On Error Resume Next
        With DGDIDOK
             txtDID.Text = Trim(.Columns(0).Text)
              Call GetDIDInfo(txtDID, TxtWO)
              Call GetCompDispInfo(Trim(TxtWO), Trim(CboJob), Trim(TxtMachine), Trim(TxtCompPN), "", "")
              Call GetSlot(Trim(txtDID), Trim(TxtWO))
             
             If Err.Number <> 0 Then
                txtDID.Text = vbNullString
               
             End If
        End With
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
End Sub
'This not be used  --20070529
Private Function ListWO()
Dim str As String
Dim TransDate As String
Dim rs As ADODB.Recordset
TransDate = Format(dtpSDate, "YYYY/MM/DD")
TransDate = Replace(TransDate, "-", "/")
str = "select distinct WO from Sap_Wo_list where Trans_Date like '" & TransDate & "%'"
Set rs = Conn.Execute(str)
cboWO.Clear
While Not rs.EOF
      cboWO.AddItem Trim(rs!WO)
      rs.MoveNext
Wend
End Function
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
If OptRelease.Value = True Then
   str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "'"
Else
    str = "select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "'"
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

str = "select Work_Order from QSMS_WOGroup  where GroupID= '" & GroupID & "' order by Seq_NO"

Set rs = Conn.Execute(str)
cboWO.Clear
CboNotFinishedWO.Clear
CboNotChkBOM.Clear
While Not rs.EOF
      If ChkMBWo(rs!Work_Order) = True Then
            If ChkQSMS_WO(Trim(rs!Work_Order)) = False Then
                CboNotChkBOM.AddItem Trim(rs!Work_Order)
            Else
            
                
                If ChkWoFinished(rs!Work_Order) = True Then
    
                    cboWO.AddItem Trim(rs!Work_Order)
                Else
                    
                     CboNotFinishedWO.AddItem Trim(rs!Work_Order)
                End If
            End If
      End If
      rs.MoveNext
Wend
End Function


Private Function GetWoinfo(ByVal WO As String)
Dim str As String
Dim rs As ADODB.Recordset
str = "select PN, Qty,[Group],BuildType from Sap_Wo_List where WO='" & Trim(WO) & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   TxtMBPN = rs!PN
   TxtModel = Mid(TxtMBPN, 3, 3)
   TxtWOQty = rs!Qty
   TxtGroup = Trim(rs![Group])
   TxtBuildType = Trim(rs!BuildType)
   
End If
str = "select Customer from ModelName where PN='" & TxtMBPN & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   TxtCustomer = Trim(rs!Customer)
End If
End Function
Private Function GetLine()
Dim str As String
Dim rs As ADODB.Recordset
str = "select distinct Line from QSMS_woGroup order by Line"    '(0001)
Set rs = Conn.Execute(str)
CboLine.Clear
While Not rs.EOF
    CboLine.AddItem rs!Line
    rs.MoveNext
Wend
End Function



Private Function GetConsumeQty(WO, COMPPN) As Long
Dim str As String
Dim rs As ADODB.Recordset
str = "select ConsumedQty from SMT_QSMS_Out where Work_Order='" & Trim(WO) & "' and CompPN='" & COMPPN & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   GetConsumeQty = Trim(rs!ConsumedQty)
End If


End Function

Private Function ChkErr() As Boolean
ChkErr = True
If Trim(TxtWO) = "" Then
   MsgBox "The WO is error,Please check"
   ChkErr = False
End If
If Trim(TxtMachine) = "" Then
    MsgBox "Machine selected is error, Please check"
    ChkErr = False
End If
If Trim(CboGroupID) = "" Then
   MsgBox "GroupID is error,Please check"
   ChkErr = False
End If
If Trim(txtDID) = "" Then
    MsgBox "DID selected  is error,please check"
    ChkErr = False
End If
If Trim(TxtDispatchQty) = "" Then
   MsgBox "Dispatch Qty is error,Please check"
   ChkErr = False
End If
If ChkWOSeq(Trim(TxtWO), Trim(TxtMachine)) = False Then
   ChkErr = False
  
End If

If Trim(TxtDispatchQty) = "0" Then
   MsgBox "Dispatch Qty is error,please check"
   ChkErr = False
   Exit Function
End If
If CLng(TxtDispatchQty) > CLng(TxtDIDTotalQty) Then
   MsgBox "dispatch Qty can not larger than DID TotalQty"
   ChkErr = False
End If
If CLng(TxtDispatchQty) > CLng(TxtDIDRemainQty) Then
   MsgBox "dispatch Qty can not larger than DID remainQty"
   ChkErr = False
End If
If Trim(TxtCompPN) = "" Then
    MsgBox "CompPN can not be empty,Please check"
    ChkErr = False
End If


If ChkDIDCompPN(Trim(txtDID), Trim(TxtCompPN)) = False Then
    ChkErr = False
End If
If ChkNonAVL(Trim(txtDID), Trim(TxtCustomer), Trim(TxtModel), Trim(TxtMBPN), Trim(TxtWO)) = False Then
   ChkErr = False
  ' MsgBox "The DID doesn't been approved,Please check"
End If
If ChkAVL(Trim(TxtCompPN), Trim(TxtVendorCode), Trim(TxtCustomer), Trim(TxtModel)) = False Then
   MsgBox "Check AVL failed,please check "
   ChkErr = False
End If
If ChkDIDBelongToGroup(Trim(txtDID), Trim(CboGroupID)) = False Then
   ChkErr = False
End If
End Function

Private Function Insert_QSMS_Out() As Boolean
Dim str As String
Dim rs As ADODB.Recordset
Dim RsInsert As ADODB.Recordset
Dim RsWoarray As ADODB.Recordset
Dim transdatetime As String
Dim tempqty As Long, tempblaqty As Long
Dim position As Long
Dim COMPPN, SelectedCompPN As String
Dim TempDisQty, TempBalQty, TempDIDQty As Long
Dim DispQtybyWO, BalanceQtyByWO As Long
Dim Item, Slot, LR As String
Dim tempwo As String

str = "Select getdate()"
Set rs = Conn.Execute(str)
transdatetime = Format(rs(0), "YYYYMMDDHHMMSS")
Insert_QSMS_Out = True

'(1) update QSMS_WO

''''''maybe need update by item, check with SAP_BOM ,if the Item is unigue for a work order

Slot = Trim(Cboslot)
LR = Trim(CboLR)

str = "select Item,Slot,LR,NeedQty,DispatchQty,BalanceQty,Work_Order,WoQty,JobPN,JobGroup,Side from QSMS_Wo where (Work_Order ='" & TxtWO & "' or work_order in " & GetWoArray & ")  " & _
      "  and compPN='" & TxtCompPN & "' and Machine='" & Trim(TxtMachine) & "' and JobGroup like '" & CboJob & "%' and Slot='" & Trim(Slot) & "' and LR='" & LR & "' order by work_Order"
Set rs = Conn.Execute(str)
tempwo = ""
If Not rs.EOF Then
    If ChkDIDDispatchedToWo(rs!Work_Order, txtDID, Trim(TxtMachine), Trim(rs!Slot), Trim(rs!LR)) = False Then
       MsgBox "The DID has dispathcd to The WO, Please check" & rs!Work_Order
    Else

              TxtDIDRemainQty = CLng(TxtDIDRemainQty) - CLng(TxtDispatchQty)
              DispQtybyWO = CLng(TxtDispatchQty) + rs!DispatchQty
           
           ' (2) ##########insert into DID dispatch qty to QSMS_Dispatch------dispatch log
'           Str = "Insert Into QSMS_Dispatch (Work_Order,GroupID,Line,WoQty,JobPN,Machine,CompPN,Slot,LR,BaseQty,NeedQty,DID,TotalQty,DIDQty,VendorCode,DateCode,LotCode,UID,TransDateTime,DIDDateTime,DeletedFlag) " & _
'              " values('" & Trim(Rs!Work_Order) & "','" & Trim(CboGroupID) & "','" & Trim(CboLine) & "','" & Trim(Rs!WOqty) & "','" & Trim(Rs!Jobpn) & "','" & TxtMachine & "', '" & TxtCompPN & "','" & Slot & "','" & LR & "','" & TxtCompBaseQty & "' " & _
'               "," & Trim(Rs!NeedQty) & ",'" & TxtDID & "'," & CLng(TxtDIDTotalQty) & "," & CLng(TxtDispatchQty) & ",'" & TxtVendorCode & "','" & TxtDateCode & "','" & TxtLotCode & "','" & g_userName & "','" & TransDateTime & "','" & DIDDateTime & "','N')"
'
'           Conn.Execute Str
'
'           If Trim(Rs!Item) = "0" Then
'              Str = "Update QSMS_Wo set dispatchqty= " & DispQtybyWO & ",BalanceQty= " & BalanceQtyByWO & " where Work_Order='" & Trim(Rs!Work_Order) & "' and  CompPN='" & TxtCompPN & "' and Slot='" & Slot & "' and LR='" & LR & "' and Machine='" & Trim(TxtMachine) & "'"
'           Else
'             Str = "Update QSMS_Wo set dispatchqty= " & DispQtybyWO & ",BalanceQty= " & BalanceQtyByWO & " where Work_Order='" & Trim(Rs!Work_Order) & "' and  " & _
'                  " Item='" & Trim(Rs!Item) & "' and slot='" & Slot & "' and LR='" & Trim(LR) & "' and Machine='" & Trim(TxtMachine) & "'"
'
'           End If
'           Conn.Execute Str
              str = "exec  QSMSInsertDispatch  '" & rs!Work_Order & "','" & Trim(CboGroupID) & "' ,'" & Trim(CboLine) & "' ,'" & Trim(rs!WOqty) & "','" & Trim(rs!Jobpn) & "' ,'" & TxtMachine & "' " & _
                  " ,'" & TxtCompPN & "' ,'" & Slot & "','" & LR & "','" & TxtCompBaseQty & "'," & Trim(rs!NeedQty) & " ,'" & Trim(txtDID) & "'," & CLng(TxtDIDTotalQty) & " ," & TxtDispatchQty & " " & _
                  ",'" & TxtVendorCode & "' ,'" & txtDateCode & "','" & txtLotCode & "','" & g_userName & "' ,'" & transdatetime & "','" & Trim(DIDDateTime) & "','More','" & Trim(rs!Item) & "','" & Trim(rs!JobGroup) & "','" & Trim(rs!Side) & "'"

           Set RsInsert = Conn.Execute(str)
           If RsInsert.EOF Then
              MsgBox "Insert into QSMS_Dispatch Error,please retry again"
              Insert_QSMS_Out = False
              Exit Function
           Else
               If UCase(Trim(RsInsert.Fields(0))) = "PASS" Then
               Else
                      'if error retry again
                      Set RsInsert = Conn.Execute(str)
                        If RsInsert.EOF Then
                           MsgBox "Insert into QSMS_Dispatch Error,please retry again"
                           Insert_QSMS_Out = False
                           Exit Function
                        Else
                             If UCase(Trim(RsInsert.Fields(0))) = "PASS" Then
                             Else
                                    Insert_QSMS_Out = False
                                    MsgBox "Insert into QSMS_Dispatch Error,please retry again"
                                    Exit Function
                             End If
                        End If
                    
               End If

           End If
           tempwo = Trim(rs!Work_Order)
           BalanceQtyByWO = DispQtybyWO - rs!NeedQty
           TxtConsumedQty = CLng(TxtConsumedQty) + CLng(TxtDispatchQty)
    End If
    
End If

'(3)check if the DID has been dispatch finished,if Yes--update the DID Flag,if No----Update the remain Qty
'If TxtDIDRemainQty = 0 And TempDisQty <= 0 Then
'   'Str = "Delete from QSMS_DID where DID='" & TxtDID & "'"
'   Str = "Update QSMS_DID set UsedFlag='Y',RemainQty=0 where DID='" & TxtDID & "'"
'Else
'    Str = "Select UsedFlag from QSMS_DID where DID='" & Trim(TxtDID) & "' and UsedFlag='Y'"
'   Set Rs = Conn.Execute(Str)
'   If Not Rs.EOF Then
'      InsSAP_BOM_FAIL "Update DID", TxtDID, "Dispatch additional DID by WO : " & TxtWO
'   End If
'   Str = "Update QSMS_DID set RemainQty='" & TxtDIDRemainQty & "' where DID='" & Trim(TxtDID) & "'"
'End If
'Conn.Execute Str



'(4)  check if need refresh DID combo

   Call RefreshDID_Machine_WO("DID", TxtCompPN, TxtMachine, TxtWO, TxtMBPN, CboLine)
   Call UpdateMachineFlagByWO(Trim(TxtWO))
   Call ChkGroupFinished(Trim(TxtWO))
''(5)) check if the machine has been dispatch finished
'Str = "select distinct Work_Order from QSMS_WO where Work_Order ='" & TxtWO & "' "
'Set RsWoarray = Conn.Execute(Str)
'While Not RsWoarray.EOF
'    Str = "select Work_Order from QSMS_WO where Work_Order ='" & Trim(RsWoarray!Work_Order) & "' and machine='" & TxtMachine & "' and balanceQty<0"
'    Set Rs = Conn.Execute(Str)
'    If Rs.EOF Then
'       Str = "Update QSMS_WO set MachineFinishedFlag='Y' where   Work_Order ='" & Trim(RsWoarray!Work_Order) & "' and machine='" & TxtMachine & "'"
'       Conn.Execute Str
'       Call RefreshDID_Machine_WO("Machine", TxtCompPN, TxtMachine, TxtWO, TxtMBPN, CboLine)
'       '(6) Check if the WO has been dispatch finished
'       Str = "select WoFinishedFlag from QSMS_WO where  Work_Order ='" & Trim(RsWoarray!Work_Order) & "' and MachineFinishedFlag<>'Y'"
'       Set Rs = Conn.Execute(Str)
'       If Rs.EOF Then
'           Str = "Update QSMS_WO set WoFinishedFlag='Y' where   Work_Order ='" & Trim(RsWoarray!Work_Order) & "' "
'           Conn.Execute Str
'           Str = "Update QSMS_WoGroup set DispatchFlag='Y' where Work_Order ='" & Trim(RsWoarray!Work_Order) & "' "
'           Conn.Execute Str
'
'           Call RefreshDID_Machine_WO("WO", TxtCompPN, TxtMachine, Trim(RsWoarray!Work_Order), TxtMBPN, CboLine)
'
'           Call ChkGroupFinished(Trim(TxtWO))
'          MsgBox "The work Order finished the dispatching "
'       End If
'    End If
'    RsWoarray.MoveNext
'Wend
TxtCompPN.Text = ""
txtDID.Text = ""
txtDID.SetFocus
End Function
Public Function RefreshDID_Machine_WO(ByVal RefreshType As String, ByVal COMPPN As String, ByVal Machine As String, ByVal WO As String, ByVal MBPN As String, ByVal Line As String)
Dim str As String
Dim rs As ADODB.Recordset


Select Case UCase(RefreshType)
       Case "DID"
             Call GetDID(Machine, MBPN, WO, Line)
       Case "MACHINE"
             Call GetMachine(WO)
       Case "WO"
             Call GetGroupWO(CboGroupID)
End Select

End Function


Private Function GetDID(ByVal Machine As String, ByVal Jobpn As String, ByVal WO As String, ByVal Line As String)
'Dim Str As String
'Dim rs As ADODB.Recordset
'
'
'
'Str = "select a.CompPN from QSMS_NonAVL a,QSMS_Wo b where B.Work_Order ='" & TxtWO & "'" & _
'     "and a.CompPN=b.CompPN and a.Customer='" & Trim(TxtCustomer) & "' and Model='" & Trim(TxtModel) & "'"
'Set rs = Conn.Execute(Str)
'Set DGAVL.DataSource = rs
'DGAVL.Refresh
'Str = "Select  a.CompPn,-sum(a.BalanceQty) as NeedQty from QSMS_WO a where a.Work_Order ='" & TxtWO & "' " & _
'      "and a.JobGroup like '" & Jobpn & "%' and a.Machine='" & Trim(Machine) & "' and a.BalanceQty<0 group by a.comppN"
'Set rs = Conn.Execute(Str)
'DGCompNotOK.Caption = "Comp didn't dispatch:" & rs.RecordCount
'Set DGCompNotOK.DataSource = rs
'DGCompNotOK.Refresh

Dim str As String
Dim rs As ADODB.Recordset
Dim TempRs As ADODB.Recordset
Dim wostr As String
wostr = GetWoArray



str = "select a.CompPN from QSMS_NonAVL a,QSMS_Wo b where (B.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') or b.work_order in " & wostr & ") " & _
     "and a.CompPN=b.CompPN and a.Customer='" & Trim(TxtCustomer) & "' and Model='" & Trim(TxtModel) & "'"
Set rs = Conn.Execute(str)
Set DGAVL.DataSource = rs
DGAVL.Refresh
str = "Select  a.CompPn,-sum(a.DispatchQty-a.PlanNeedQty) as NeedQty from QSMS_WO a where (a.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') or a.work_order in " & wostr & ") " & _
      "and a.Machine='" & Trim(Machine) & "' and a.DispatchQty-a.PlanNeedQty<0 and JobGroup like '" & Trim(Jobpn) & "%' group by a.comppN"
Set rs = Conn.Execute(str)
DGCompNotOK.Caption = "Comp didn't dispatch:" & rs.RecordCount
Set DGCompNotOK.DataSource = rs

DGCompNotOK.Refresh


End Function
Private Function GetCompPnWithoutDID(ByVal Machine As String, ByVal WO As String)
Dim str As String
Dim rs As ADODB.Recordset
Dim TempRs As ADODB.Recordset
CboWithout.Clear
str = "select CompPN,Item from  QSMS_WO where Work_Order='" & Trim(WO) & "' and machine='" & Machine & "' and CompPN not in (select CompPN from QSMS_DID where UsedFlag='N')"
Set rs = Conn.Execute(str)
While Not rs.EOF
      If Trim(rs!Item) = "0" Then
         If ChkBalanceQty(Machine, WO, rs!COMPPN) = False Then
           CboWithout.AddItem Trim(rs!COMPPN)
         End If
      Else
         str = "select CompPN from  QSMS_WO where Work_Order='" & Trim(WO) & "' and machine='" & Machine & "' and Item='" & Trim(rs!Item) & "' and CompPN  in (select CompPN from QSMS_DID where UsedFlag='N')"
         Set TempRs = Conn.Execute(str)
         If TempRs.EOF Then
             If ChkBalanceQty(Machine, WO, rs!COMPPN) = False Then
                CboWithout.AddItem Trim(rs!COMPPN)
             End If
         End If
      End If
      rs.MoveNext
Wend
End Function
Private Function ChkBalanceQty(ByVal Machine As String, ByVal WO As String, ByVal COMPPN As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
ChkBalanceQty = True
str = "select Work_Order From QSMS_WO where Work_Order='" & WO & "' and Machine='" & Machine & "' and CompPN='" & COMPPN & "' and BalanceQty<0"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   ChkBalanceQty = False
End If
End Function
Private Function GetDIDInfo(ByVal DID As String, ByVal WO As String)
Dim str As String
Dim rs As ADODB.Recordset
Dim Used_Flag As String
TxtConsumedQty = ""
TxtCompBaseQty = ""
TxtNeedQty = ""
TxtDispatchQty = ""
TxtCompPN = ""
TxtVendorCode = ""

txtDateCode = ""
txtLotCode = ""
TxtDIDRemainQty = ""
DIDDateTime = ""

'Str = "select a.DID,a.CompPN,a.VendorCode,a.DateCode,a.LotCode,a.Qty,a.RemainQty,a.UsedFlag,b.BaseQty,b.DispatchQty from QSMS_DID a,QSMS_WO b " & _
      " where a.DID='" & Trim(DID) & "' and a.CompPN=b.CompPN and b.Work_Order='" & WO & "' and b.machine='" & Trim(TxtMachine) & "'"
str = "select DID,CompPN,VendorCode,DateCode,LotCode,Qty,RemainQty,UsedFlag,DIDLoc,TransDateTime from QSMS_DID where DID='" & Trim(txtDID) & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   TxtCompPN = Trim(rs!COMPPN)
   TxtVendorCode = Trim(rs!VendorCode)
   txtDateCode = Trim(rs!DateCode)
   txtLotCode = Trim(rs!LotCode)
   TxtRackID = Trim(rs!DIDLoc)
   TxtDIDTotalQty = Trim(rs!Qty)
   TxtDIDRemainQty = Trim(rs!RemainQty)
   DIDDateTime = Trim(rs!transdatetime)
    If UCase(Trim(rs!usedflag)) = "Y" Then

       TxtDispatchQty.Enabled = False
       TxtDispatchQty = 0
   Else

       TxtDispatchQty.Enabled = True
       
   End If
   
   
   
  
End If
End Function
Private Function GetCompDispInfo(ByVal WO As String, ByVal Jobpn As String, ByVal Machine As String, ByVal COMPPN As String, ByVal Slot As String, ByVal LR As String)
On Error GoTo Handler
Dim str As String
Dim rs As ADODB.Recordset
Dim WoArray As String
Dim LRNo As Integer
Dim LRStr As String
TxtConsumedQty = ""
TxtCompBaseQty = ""
TxtNeedQty = ""
LRNo = 0
CboLR.Clear
WoArray = GetWoArray
'Str = "Select BaseQty as BaseQty,NeedQty as NeedQty,DispatchQty as DispatchQty,LR from QSMS_WO where (Work_Order='" & TxtWO & "' or work_order in " & GetWoArray & ") " & _
'       "  and jobgroup like '" & Jobpn & "%' and machine='" & Machine & "' and comppn='" & CompPN & "' and slot like '" & Slot & "%' and LR like '" & LR & "%' "
'Set rs = Conn.Execute(Str)
'While Not rs.EOF
'    LRNo = LRNo + 1
'    BaseQty = BaseQty + rs!BaseQty
'    ConsumeQty = ConsumeQty + rs!DispatchQty
'    NeedQty = NeedQty + rs!NeedQty
'    LRStr = Trim(rs!LR)
'    CboLR.AddItem LRStr
'    rs.MoveNext
'Wend

'*******************************************************(0002)***************************************
str = "Select sum(BaseQty) as BaseQty,sum(NeedQty) as NeedQty,sum(DispatchQty) as DispatchQty from QSMS_WO where (Work_Order='" & TxtWO & "' or work_order in " & WoArray & ") " & _
       "  and jobgroup like '" & Jobpn & "%' and machine='" & Machine & "' and comppn='" & COMPPN & "' and slot like '" & Slot & "%' and LR like '" & LR & "%' "
Set rs = Conn.Execute(str)
If Not rs.EOF Then
    TxtCompBaseQty = Trim(rs!BaseQty)
    TxtConsumedQty = Trim(rs!DispatchQty)
    TxtNeedQty = Trim(rs!NeedQty)
End If
If Trim(TxtNeedQty) = "" Then
  Exit Function
End If

str = "Select distinct LR from QSMS_WO where (Work_Order='" & TxtWO & "' or work_order in " & WoArray & ") " & _
       "  and jobgroup like '" & Jobpn & "%' and machine='" & Machine & "' and comppn='" & COMPPN & "' and slot like '" & Slot & "%' and LR like '" & LR & "%' "
Set rs = Conn.Execute(str)
While Not rs.EOF
    LRNo = LRNo + 1
    CboLR.AddItem Trim(rs!LR)
    rs.MoveNext
Wend
'*******************************************************(0002)***************************************

If LRNo = 1 Then
   CboLR.ListIndex = 0
   CboLR.Enabled = False
   TxtDispatchQty.Enabled = True
   TxtDispatchQty.SetFocus
Else
   CboLR.Enabled = True
   CboLR.SetFocus
End If

Exit Function
Handler:
    MsgBox "Please call QMS for help"
End Function
Private Function GetSlot(ByVal DID As String, ByVal WO As String)
Dim str As String
Dim rs As ADODB.Recordset
Dim Used_Flag As String
Dim Slot As String
Dim i As Long
LblSlot(4).Visible = True
Cboslot.Visible = True
LblSlot(4).BackColor = &HFF&


Cboslot.Clear
Slot = ""
i = 0
str = "select distinct b.Slot from QSMS_DID a,QSMS_WO b where a.DID='" & Trim(DID) & "' and a.CompPN=b.CompPN and (b.Work_Order ='" & TxtWO & "' OR b.work_order in " & GetWoArray & ") and b.JobGroup like '" & CboJob & "%'" & _
      "and b.Machine='" & Trim(TxtMachine) & "'"
Set rs = Conn.Execute(str)
While Not rs.EOF
      Slot = Trim(rs!Slot)
      Cboslot.AddItem Trim(rs!Slot)
      rs.MoveNext
      If Slot <> "" Then
        i = i + 1
      End If
Wend
If i = 1 Then
   Cboslot.Text = Slot
   Cboslot.Enabled = False
   Call GetCompDispInfo(Trim(TxtWO), Trim(CboJob), Trim(TxtMachine), Trim(TxtCompPN), Trim(Slot), "")
Else
  Cboslot.Enabled = True

End If
End Function



Private Function GetMachine(ByVal WO As String)
Dim str As String
Dim TransDate As String
Dim rs As ADODB.Recordset
Dim rsMachine As ADODB.Recordset
str = "select distinct Machine from QSMS_WO where Work_Order ='" & TxtWO & "' and Line like '" & Trim(CboLine) & "%'"    '''1054

Set rs = Conn.Execute(str)
CboMathineOK.Clear
CboMathineNOK.Clear
While Not rs.EOF
     str = "select  Machine from QSMS_WO where Work_Order ='" & TxtWO & "' and MachinefinishedFlag='N' and machine='" & Trim(rs!Machine) & "'"
     Set rsMachine = Conn.Execute(str)
     If rsMachine.EOF Then
        CboMathineOK.AddItem Trim(rs!Machine)
     Else
        CboMathineNOK.AddItem Trim(rs!Machine)
       
     End If
     rs.MoveNext
Wend

End Function




Private Function ChkWOSeq(ByVal WO As String, ByVal Machine As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim TempGroupID As String
Dim Seq_No As Long
ChkWOSeq = True
str = "select GroupID,Seq_No from  QSMS_wogroup where Work_Order='" & Trim(WO) & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   TempGroupID = Trim(rs!GroupID)
   Seq_No = rs!Seq_No
Else
   ChkWOSeq = False
   MsgBox "The Work order has No group ID, please call ME  "
End If
'str = "select Work_Order from QSMS_WoGroup where GroupID='" & TempGroupID & "' and seq_no<" & Seq_NO & " order by seq_no"
'Set Rs = Conn.Execute(str)
'While Not Rs.EOF
'      str = "select Work_order from QSMS_Wo Where Work_Order='" & Trim(Rs!Work_Order) & "' and machine='" & Machine & "' and MachineFinishedFlag='N'"
'      Set rsTemp = Conn.Execute(str)
'      If Not rsTemp.EOF Then
'         ChkWOSeq = False
'         MsgBox "can not dispatch the wo:" & WO & "  you must dispatch the Wo :" & rsTemp!Work_Order & "  First"
'         Exit Function
'      End If
'      Rs.MoveNext
'Wend


End Function

Public Function ChkDIDBelongMachine(ByVal WO As String, ByVal Machine As String, ByVal DID As String, SAPWOGroup As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
Dim MessageString As String
Dim strSQL As String

str = "select UsedFlag from QSMS_DID where DID='" & DID & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   If UCase(Trim(rs!usedflag)) = "Y" Then
      MessageString = ""
      str = "select a.work_order,a.Machine,a.Slot From qsms_dispatch a join qsms_did b on a.did=b.did and a.diddatetime=b.transdatetime where a.did='" & DID & "'"
      Set rs = Conn.Execute(str)
      Do While Not rs.EOF
            MessageString = MessageString + "WO:" + rs!Work_Order + "  Machine:" + rs!Machine + "  Slot:" + rs!Slot + vbCrLf
            rs.MoveNext
        Loop
      MsgBox "The DID has been used at: " + vbCrLf + vbCrLf + MessageString + vbCrLf + " ===PLease check=== !"
      ChkDIDBelongMachine = False
      Exit Function
   Else        '''''(0004)Add by Udall 2007.11.05
      MessageString = ""
      str = "select a.work_order,a.Machine,a.Slot From qsms_dispatch a join qsms_did b on a.did=b.did and a.diddatetime=b.transdatetime where a.did='" & DID & "'"
      Set rs = Conn.Execute(str)
      Do While Not rs.EOF
            If Left(Trim(rs!Machine), 2) <> Left(Machine, 2) Then
               MessageString = MessageString + "WO:" + rs!Work_Order + "  Machine:" + rs!Machine + "  Slot:" + rs!Slot + vbCrLf
            End If
            rs.MoveNext
        Loop
      If MessageString <> "" Then
          MsgBox "DID can not be dispatched to different line and side,the DID has been dispatched to: " + vbCrLf + vbCrLf + MessageString + vbCrLf + " ===PLease check=== !"
          ChkDIDBelongMachine = False
          Exit Function
      End If      ''''(0004)
   End If
Else
'''''''''''''''''''''''''''''''''''''add by Jing 2007.10.31------------(0003)'''''''''''''''''''''''''''''''''''''''''''''''
    strSQL = "select * from qsms_did_log where did='" & DID & "'"
    Set rs = Conn.Execute(strSQL)
    If rs.RecordCount > 0 Then
        MsgBox "This DID had been deleted !"
    Else
       MsgBox "DID does not exist, please confirm."
    End If
   ChkDIDBelongMachine = False
   Exit Function
End If

str = "select b.DID,B.Qty from QSMS_WO a, QSMS_DID b where (a.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(SAPWOGroup) & "') or a.work_order in " & GetWoArray & ") " & _
       "and a.Machine='" & Trim(Machine) & "' and a.CompPN=b.CompPN  and b.DID='" & DID & "' and b.UsedFlag='N'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   ChkDIDBelongMachine = True
Else
   ChkDIDBelongMachine = False
   MsgBox "the DID does not belong to the machine :" & Machine & ".  Please Check."
End If


'Str = "select b.DID,B.Qty from QSMS_WO a, QSMS_DID b where a.Work_Order ='" & TxtWO & "' " & _
'       "and a.Machine='" & Trim(Machine) & "' and a.CompPN=b.CompPN  and b.DID='" & DID & "' and b.UsedFlag='N'"
'Set rs = Conn.Execute(Str)
'If Not rs.EOF Then
'   ChkDIDBelongMachine = True
'Else
'   ChkDIDBelongMachine = False
'   MsgBox "the DID does not belong to the machine :" & Machine & ".  Please Check."
'End If

End Function

'Private Sub TxtChkDID_KeyPress(KeyAscii As Integer)
'Dim Str As String
'Dim Rs As ADODB.Recordset
'If KeyAscii = 13 Or KeyAscii = 9 Then
'   ListDIDStatus.Clear
'   'str = "select b.DID,B.Qty ,B.UsedFlag from QSMS_WO a, QSMS_DID b where a.Work_Order='" & Trim(TxtWO) & "'  and a.Machine='" & Trim(TxtMachine) & "' and a.CompPN=b.CompPN  and b.DID='" & TxtChkDID & "' "
'   Str = "select b.DID,B.Qty ,B.UsedFlag from QSMS_WO a, QSMS_DID b where (a.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') or a.work_order in " & GetWoArray & ") " & _
'         "and a.Machine='" & Trim(TxtMachine) & "' and a.CompPN=b.CompPN  and b.DID='" & TxtChkDID & "' "
'
'   Set Rs = Conn.Execute(Str)
'   If Not Rs.EOF Then
'      If UCase(Rs!UsedFlag) = "Y" Then
'         ListDIDStatus.AddItem "Has been Dispatched"
'      Else
'         ListDIDStatus.AddItem "Not Dispatched"
'      End If
'   Else
'
'         ListDIDStatus.AddItem "Not belong to the Machine"
'   End If
'End If
'End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

End Sub

Private Sub txtDID_KeyPress(KeyAscii As Integer)
On Error GoTo Handle:
If KeyAscii = 13 Or KeyAscii = 9 Then
    
   If ChkDIDBelongMachine(Trim(TxtWO), Trim(TxtMachine), Trim(txtDID) & "", Trim(TxtGroup)) = False Then
     txtDID.Text = ""
     txtDID.SetFocus
     Exit Sub

   End If
   
'------------add check IPQCFlag='Y' before dispatch----------(0005)
   If StrBU = "MBU" And ChkIPQC(Trim(txtDID)) = False Then
        Call Warning_Sound
        MsgBox "DID:" & Trim(txtDID) & " Check IPQC failed,please check!!! "
        Exit Sub
   End If
'------------end (0005)-------------------------------------------
   Call GetDIDInfo(Trim(txtDID), Trim(TxtWO))
   Call GetSlot(Trim(txtDID), Trim(TxtWO))
   
   
End If

Exit Sub
Handle:
    MsgBox "Please call QMS for help"
End Sub

Private Function ChkIPQC(ByVal DID As String) As Boolean '(0005)
Dim str As String
Dim rs As ADODB.Recordset

str = "select IPQCFlag from QSMS_DID where DID='" & DID & "' and LEFT(CompPN,2) IN ('TH','CH','TS','CS','TU')"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
    If UCase(Trim(rs!IPQCFlag)) = "Y" Then
        ChkIPQC = True
        TxtDispatchQty.Enabled = True
        Exit Function
    End If
Else
    ChkIPQC = True
    TxtDispatchQty.Enabled = True
    Exit Function
End If

    ChkIPQC = False
    TxtDispatchQty.Enabled = False
    TxtDispatchQty = 0

End Function



Private Function ChkDIDInherit(ByVal WO As String, ByVal DID As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
Dim GroupID As String
ChkDIDInherit = False
If TxtDIDTotalQty = TxtDIDRemainQty Then
   ChkDIDInherit = True
   Exit Function
End If
str = "select GroupID from QSMS_WOGroup where Work_Order='" & WO & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   GroupID = Trim(rs!GroupID)
   str = "select DID from QSMS_Dispatch a, QSMS_WOGroup b where a.work_order=b.Work_Order and B.GroupID='" & GroupID & "' and a.DID='" & DID & "' "
   Set rs = Conn.Execute(str)
   If rs.EOF Then
      ChkDIDInherit = False
   Else
      ChkDIDInherit = True
   End If
End If


End Function

'Private Function ChkDIDDispatchedToWo(ByVal Wo As String, ByVal DID As String) As Boolean
'Dim Str As String
'Dim Rs As ADODB.Recordset
'
'   Str = "select DID from QSMS_Dispatch  where work_order='" & Wo & "'  and DID='" & DID & "' "
'   Set Rs = Conn.Execute(Str)
'   If Rs.EOF Then
'      ChkDIDDispatchedToWo = True
'   Else
'      ChkDIDDispatchedToWo = False
'   End If
'
'
'End Function
Private Function ChkDIDDispatchedToWo(ByVal WO As String, ByVal DID As String, ByVal Machine As String, ByVal Slot As String, LR As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
Dim DeleteFlag As Boolean
DeleteFlag = False
ChkDIDDispatchedToWo = False
   str = "select DID,DeletedFlag,Machine,Slot,LR from QSMS_Dispatch  where work_order='" & WO & "'  and DID='" & DID & "' "
   Set rs = Conn.Execute(str)
   If rs.EOF Then
      ChkDIDDispatchedToWo = True
      Exit Function
   End If
   While Not rs.EOF
      If UCase(Trim(rs!DeletedFlag = "Y")) Or (UCase(Machine) = UCase(Trim(rs!Machine)) And UCase(Slot) = UCase(Trim(rs!Slot)) And UCase(LR) = UCase(Trim(rs!LR))) Then
         ChkDIDDispatchedToWo = True
      Else
         ChkDIDDispatchedToWo = False
         Exit Function
      End If
      rs.MoveNext
  Wend
   
End Function
Private Function ChkDIDCompPN(ByVal DID As String, ByVal COMPPN As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
ChkDIDCompPN = True
str = "select DID from QSMS_DID where DID='" & DID & "' and CompPN='" & COMPPN & "'"
Set rs = Conn.Execute(str)
If rs.EOF Then
   MsgBox "The DID and CompPN doesn't match"
   ChkDIDCompPN = False
End If
End Function
'Private Function ChkMBWo(ByVal WO As String) As Boolean
'Dim Str As String
'Dim Rs As ADODB.Recordset
'ChkMBWo = False
'Str = "select WO from Sap_Wo_list where wo='" & WO & "'"
'Set Rs = Conn.Execute(Str)
'If Rs.EOF Then
'    Str = "Select count(*) from sap_wo_list where [group] in (select [group] from sap_wo_list where wo='" & WO & "')"
'    Set Rs = Conn.Execute(Str)
'    If Rs.Fields(0) > 1 Then
'       ChkMBWo = False
'    Else
'        ChkMBWo = True
'    End If
'
'Else
'   ChkMBWo = True
'
'End If
'End Function


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
   TxtGroup = Group
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

Private Sub TxtDispatchQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CmdConfirm_Click
End If
End Sub

Private Sub TxtQryMachine_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   
   Call CmdExcelDID_Click
End If
End Sub

Private Sub TxtQryWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   TxtQryMachine.SetFocus
   Call CmdExcel_Click
End If
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

Private Function GetJobForBuildType(ByVal WO As String, ByVal Machine As String, ByVal BuildType As String)
Dim str As String
Dim rs As ADODB.Recordset
CboJob.Clear
CboJob.Text = ""
Select Case BuildType
       Case "1"
               Label1(4).Visible = False
               CboJob.Visible = False
       Case "2", "3"
                Label1(4).Visible = True
                CboJob.Visible = True
                
                str = "select distinct JobGroup from qsms_wo where Work_Order='" & Trim(WO) & "' and machine='" & Machine & "'"
                Set rs = Conn.Execute(str)
                If rs.RecordCount > 1 Then
                   While Not rs.EOF
                         CboJob.AddItem Trim(rs!JobGroup)
                         rs.MoveNext
                   Wend
               End If
End Select
End Function

Private Function GetWoArray() As String
Dim WoArray As String
Dim str As String
Dim rs As ADODB.Recordset
Dim i As Long

    str = "select wo from Sap_WO_List where [Group] in (select [group] from sap_wo_list where wo='" & Trim(TxtWO) & "')"
        Set rs = Conn.Execute(str)
        While Not rs.EOF
               WoArray = WoArray + "'" + Trim(rs!WO) + "'" + ","
               rs.MoveNext
        Wend
    
'    For i = 1 To ListWoDispatching.ListCount
'        ListWoDispatching.ListIndex = i - 1
'        Str = "select wo from Sap_WO_List where [Group] in (select [group] from sap_wo_list where wo='" & ListWoDispatching.Text & "')"
'        Set rs = Conn.Execute(Str)
'        While Not rs.EOF
'               WoArray = WoArray + "'" + Trim(rs!WO) + "'" + ","
'               rs.MoveNext
'        Wend
        
 '   Next i
    WoArray = Mid(WoArray, 1, Len(WoArray) - 1)
    WoArray = "(" + WoArray + ")"
    GetWoArray = WoArray
End Function

