VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmInheritDIDByWO 
   Caption         =   "FrmInheritDIDByWO[20160106]"
   ClientHeight    =   9960
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   15420
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      Height          =   3135
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   15375
      Begin VB.CheckBox chkIncludeXL 
         Caption         =   "Include XL CompPN"
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
         Left            =   12840
         TabIndex        =   104
         Top             =   120
         Width           =   2415
      End
      Begin VB.OptionButton OptMachine 
         BackColor       =   &H000080FF&
         Caption         =   "By Machine"
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
         Left            =   12600
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   103
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton OptSide 
         BackColor       =   &H000080FF&
         Caption         =   "By Side"
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
         Left            =   11280
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   102
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox CboInheritingWO 
         Height          =   315
         Left            =   11280
         TabIndex        =   101
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdUnlink 
         BackColor       =   &H0000FF00&
         Caption         =   "Unlink Inherit"
         Height          =   375
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton CmdDispatchedDID 
         Caption         =   "&Dispatched DID"
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
         Left            =   9360
         Picture         =   "FrmInheritDIDByWO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton CmdExcelInheritDID 
         Caption         =   "&Excel-No Inherit DID"
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
         Left            =   9360
         Picture         =   "FrmInheritDIDByWO.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton CmdInherit 
         BackColor       =   &H000080FF&
         Caption         =   "Inherit"
         Height          =   375
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox CboInheritWO 
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
         Left            =   9360
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
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
         TabIndex        =   66
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
         Left            =   9000
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   2520
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
         Left            =   11760
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   2520
         Width           =   735
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "Group"
         Height          =   255
         Left            =   1680
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton OptRelease 
         Caption         =   "Release"
         Height          =   255
         Left            =   120
         TabIndex        =   62
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
         Style           =   1  'Graphical
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2520
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
         Left            =   6600
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Frame FraSB 
         Caption         =   "Small Board WO"
         Height          =   615
         Left            =   6600
         TabIndex        =   54
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
            TabIndex        =   55
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
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1335
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   1680
         TabIndex        =   52
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
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   480
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   67
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
         Format          =   2818051
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   68
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
         Format          =   2818051
         CurrentDate     =   36482
      End
      Begin VB.Label LblInherit 
         Height          =   495
         Left            =   11280
         TabIndex        =   97
         Top             =   1920
         Width           =   3975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Inherit from WO"
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
         Left            =   9360
         TabIndex        =   83
         Top             =   120
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         Left            =   7920
         TabIndex        =   80
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
         TabIndex        =   79
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
         Left            =   11040
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         Left            =   5880
         TabIndex        =   74
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
         TabIndex        =   73
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Inheriting WO"
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
         Left            =   11280
         TabIndex        =   72
         Top             =   120
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "PD prepair material"
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   15375
      Begin VB.Frame FraQyeryDID 
         BackColor       =   &H00FFC0C0&
         Height          =   1455
         Left            =   0
         TabIndex        =   86
         Top             =   5400
         Width           =   15375
         Begin VB.Frame Frame4 
            Caption         =   "QueryWoBy DID"
            Height          =   855
            Left            =   120
            TabIndex        =   93
            Top             =   240
            Width           =   5415
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
               Left            =   1080
               TabIndex        =   95
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
               Left            =   4440
               Style           =   1  'Graphical
               TabIndex        =   94
               TabStop         =   0   'False
               Top             =   240
               Width           =   735
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
               Left            =   120
               TabIndex        =   96
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "QueryDIDByWO Machine"
            Height          =   855
            Left            =   5640
            TabIndex        =   87
            Top             =   240
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
               Style           =   1  'Graphical
               TabIndex        =   90
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
               TabIndex        =   89
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
               TabIndex        =   88
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
               TabIndex        =   92
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
               TabIndex        =   91
               Top             =   360
               Width           =   855
            End
         End
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
         TabIndex        =   45
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
         Left            =   4800
         TabIndex        =   44
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
         Left            =   8040
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFF80&
         Caption         =   "ChkDID Status"
         Height          =   1815
         Left            =   12000
         TabIndex        =   38
         Top             =   1560
         Width           =   3015
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
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   240
            Width           =   2535
         End
         Begin VB.ListBox ListDIDStatus 
            Height          =   1035
            Left            =   120
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   720
            Width           =   2535
         End
      End
      Begin VB.Frame FraWithoutDID 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Without DID"
         Height          =   735
         Left            =   120
         TabIndex        =   30
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   34
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame FraDispatchDID 
         BackColor       =   &H00FF80FF&
         Caption         =   "DispatchDID"
         Height          =   1215
         Left            =   0
         TabIndex        =   12
         Top             =   3360
         Width           =   15015
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
            TabIndex        =   20
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
            Left            =   6000
            TabIndex        =   19
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
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   720
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
            Left            =   13800
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   240
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
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   720
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
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   720
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
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   720
            Width           =   1335
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
            Left            =   13080
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
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
            TabIndex        =   29
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
            Left            =   4440
            TabIndex        =   28
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
            TabIndex        =   27
            Top             =   720
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
            Left            =   12840
            TabIndex        =   26
            Top             =   240
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
            TabIndex        =   25
            Top             =   720
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
            TabIndex        =   24
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Total Need Qty"
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
            TabIndex        =   23
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label LblMessage 
            BackColor       =   &H00FFFFC0&
            Height          =   495
            Left            =   8880
            TabIndex        =   22
            Top             =   120
            Width           =   6015
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
            Left            =   11520
            TabIndex        =   21
            Top             =   720
            Visible         =   0   'False
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000000&
         Caption         =   "DID information "
         Height          =   855
         Left            =   0
         TabIndex        =   1
         Top             =   4560
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
            TabIndex        =   6
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
            TabIndex        =   5
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
            TabIndex        =   4
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
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   120
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
            TabIndex        =   2
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
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   9
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
            TabIndex        =   8
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
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSDataGridLib.DataGrid DGAVL 
         Height          =   1800
         Left            =   9600
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   3175
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
         Height          =   1800
         Left            =   6960
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3175
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
         Height          =   1800
         Left            =   4320
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3175
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
         Height          =   1800
         Left            =   120
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4150
         _ExtentX        =   7329
         _ExtentY        =   3175
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
         TabIndex        =   49
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
         Left            =   3480
         TabIndex        =   48
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
         Left            =   6960
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmInheritDIDByWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/**********************************************************************************
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**Jing       2007.10.31     Add the message in detail in the dispatch interface for deleted DID -------(0001)
'**Lynn       2007.11.04     Add Log who do Uninherit DID -------(0002)
'**Udall      2008.01.03    If the DID is dispatched by auto, the did can't be inherited.(b.Line='' and b.WoGroup='') -------(0003)
'**Giant      2008.03.01    Inherited separate for all CompPN from not include XL CompPN -------(0004)
'**Season     2015.11.12    承接料的时候，如果DID之前没有发过这个位置，则这个DID不能承接到这个位置 ----(0005)(1218)
'***********************************************************************************

Private Sub CboGroupID_Click()
Call GetGroupWO(CboGroupID)
End Sub

Private Sub CboGroupID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboGroupID_Click
End If
End Sub





Private Sub CboInheritingWO_Click()
TxtWO = Trim(CboInheritingWO)
Call GetSBWO(TxtWO)
Call GetWoinfo(TxtWO)
Call GetMachine(TxtWO)
End Sub

Private Sub CboInheritingWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CboInheritingWO_Click
End If
End Sub

Private Sub CboInheritWO_Click()
TxtWO = Trim(CboInheritWO)
Call GetSBWO(TxtWO)
Call GetWoinfo(TxtWO)
Call GetMachine(TxtWO)
End Sub

Private Sub CboInheritWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CboInheritWO_Click
End If
End Sub

Private Sub CboMathineNOK_Click()
TxtMachine = Trim(CboMathineNOK)
If ChkWOSeq(Trim(TxtWO), Trim(TxtMachine)) = False Then
   TxtDID.Enabled = False
   Exit Sub
Else
   TxtDID.Enabled = True
End If
Call GetDID(Trim(TxtMachine), Trim(TxtMBPN), TxtWO, CboLine)
Call GetCompPnWithoutDID(Trim(TxtMachine), Trim(TxtWO))
TxtDID.Text = ""
TxtDID.SetFocus
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
   TxtDID.Enabled = False
   Exit Sub
Else
   TxtDID.Enabled = True
End If
Call GetDID(Trim(TxtMachine), Trim(TxtMBPN), TxtWO, CboLine)
CboWithout.Clear
TxtDID.Text = ""
TxtDID.SetFocus
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







Private Sub CmdDispatchedDID_Click()
Dim str As String
Dim rs As ADODB.Recordset
str = "select A.MACHINE, b.* from qsms_dispatch a,qsms_DID b " & _
       "  where  a.work_order in (select wo from sap_wo_list  where [group] in (select [group] from sap_wo_list where wo='" & Trim(CboInheritWO) & "')) " & _
       "and a.did=b.did and b.remainqty>0  and b.UsedFlag<>'Y' and a.DIDDateTime=b.TransDateTime and a.TotalQty=B.Qty"
Set rs = Conn.Execute(str)
If rs.EOF Then
   MsgBox "No DID can be inherit"
   Exit Sub
Else
   Call CopyToExcel(rs)
End If
End Sub

Private Sub cmdExcel_Click()
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
    str = "Select * from QSMS_Dispatch where CompPN='" & Trim(TxtQryDID) & "'  and work_order='" & TxtWO & "'"
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
   
   Call cmdExcel_Click
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

Private Sub CmdExcelInheritDID_Click()
Dim rs As ADODB.Recordset
If Trim(CboInheritWO) = "" Then
   Exit Sub
End If
Set rs = GetNotInheritDIDByWO(Trim(CboInheritWO))
 If Not rs.EOF Then
       Call CopyToExcel(rs)
 Else
       MsgBox ("No Data"), vbCritical
End If

End Sub

Private Sub CmdInherit_Click()
Dim str As String
Dim rs As ADODB.Recordset
Dim RsDisp As ADODB.Recordset
Dim RsInsert As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim RsWoarray As ADODB.Recordset
Dim TempDispQty As Long
Dim TempDIDQty As Long
Dim DispQtybyWO As Long
Dim BalanceQty As Long
Dim tempwo As String
Dim Item, Slot, LR As String, Machine As String
Dim BaseQty, TotalQty As Long
Dim TransDateTime, DIDDateTime As String
Dim wostr As String
Dim Insert_QSMS_Out As Boolean
Dim FistTimeofDispatch As Integer ''To record if is the first dispatch --20070715
On Error GoTo EcmdSave_Click
If OptSide.Value = False And OptMachine.Value = False Then
   MsgBox "Please select the Inherit mode!By side or by Machine.", vbCritical, "Message"
   Exit Sub
End If
If MsgBox("Are you sure to do Inherit?!", vbYesNo) = vbNo Then
Exit Sub
End If

If ChkBelongSameGroup(Trim(CboInheritWO)) = False Then
   Exit Sub
End If
LblInherit.BackColor = &HFF&
LblInherit.Caption = "System is Inheriting,Please wait!!!!!"
If chkIncludeXL.Value = 0 Then   '' ------(0004)
''-------(0003)b.Line='' and b.WoGroup=''
    str = "select distinct a.machine,a.comppn,a.did,a.vendorcode,a.datecode,a.lotcode ,b.Qty,b.remainQty,B.TransDateTime from qsms_dispatch a,qsms_DID b " & _
           "  where  a.work_order in (select wo from sap_wo_list  where [group] in (select [group] from sap_wo_list where wo='" & Trim(CboInheritWO) & "')) " & _
           "and a.did=b.did and b.Line='' and b.WoGroup='' and b.remainqty>0 and UsedFlag<>'Y' and a.DIDDateTime=b.TransDateTime and a.TotalQty=b.Qty"
Else
    str = "select distinct a.machine,a.comppn,a.did,a.vendorcode,a.datecode,a.lotcode ,b.Qty,b.remainQty,B.TransDateTime from qsms_dispatch a,qsms_DID b " & _
           "  where  a.work_order in (select wo from sap_wo_list  where [group] in (select [group] from sap_wo_list where wo='" & Trim(CboInheritWO) & "')) " & _
           "and a.did=b.did and b.remainqty>0 and UsedFlag<>'Y' and a.DIDDateTime=b.TransDateTime and a.TotalQty=b.Qty"

End If
Set rs = Conn.Execute(str)
If rs.EOF Then
   MsgBox "No DID can be inherit"
   LblInherit.BackColor = &HFF00&
    LblInherit.Caption = ""
   Exit Sub
End If
wostr = GetWoArray
FistTimeofDispatch = 1 '---20070715

While Not rs.EOF
     TxtDID = rs!DID
     TxtCompPN = Trim(rs!CompPN)
     TxtMachine = Trim(rs!Machine)
     TxtVendorCode = Trim(rs!VendorCode)
     TxtDateCode = Trim(rs!DateCode)
     TxtLotCode = Trim(rs!LotCode)
     TempDispQty = rs!RemainQty
     TotalQty = rs!Qty
     DIDDateTime = Trim(rs!TransDateTime)
     If ChkAVL(Trim(TxtCompPN), Trim(TxtVendorCode), Trim(TxtCustomer), Trim(TxtModel)) = False Then
        MsgBox "Check AVL failed,please check "
        Exit Sub
     End If
     
     If ChkNonAVL(Trim(TxtDID), Trim(TxtCustomer), Trim(TxtModel), Trim(TxtMBPN), Trim(TxtWO)) = False Then
        'MsgBox "The DID doesn't been approved,Please check"
        Exit Sub
     End If
     If OptSide.Value = True Then  ''(1090)
'        Str = "select Item,Machine,Slot,LR,NeedQty as TotalNeedQty,PlanNeedQty as NeedQty,BaseQty ,DispatchQty,DispatchQty-PlanNeedQty as BalanceQty,Work_Order,WoQty,JobPN,JobGroup,Side from QSMS_Wo where " & _
'            "  work_order in " & wostr & " " & _
'              "and compPN='" & TxtCompPN & "' and left(ltrim(Machine),2)='" & Left(Trim(TxtMachine), 2) & "' and DispatchQty-PlanNeedQty<0 order by work_Order,Machine,Slot,LR"

        str = "select Item,Machine,Slot,LR,NeedQty as TotalNeedQty,PlanNeedQty as NeedQty,BaseQty ,DispatchQty,DispatchQty-PlanNeedQty as BalanceQty,Work_Order,WoQty,JobPN,JobGroup,Side from QSMS_Wo A where " & _
            "  work_order in " & wostr & " and compPN='" & TxtCompPN & "' and DispatchQty-PlanNeedQty<0 " & _
            " and line='" & Trim(CboLine.Text) & "' and side = (SELECT Side FROM Machine B WHERE A.Machine =B.Machine AND B.Line ='" & CboLine.Text & "' AND B.Machine='" & Trim(TxtMachine) & "') " & _
            " order by work_Order,Machine,Slot,LR"
     End If
     If OptMachine.Value = True Then
        str = "select Item,Machine,Slot,LR,NeedQty as TotalNeedQty,PlanNeedQty as NeedQty,BaseQty ,DispatchQty,DispatchQty-PlanNeedQty as BalanceQty,Work_Order,WoQty,JobPN,JobGroup,Side from QSMS_Wo where " & _
            "  work_order in " & wostr & " " & _
              "and compPN='" & TxtCompPN & "' and rtrim(ltrim(Machine))='" & Trim(TxtMachine) & "' and DispatchQty-PlanNeedQty<0 order by work_Order,Machine,Slot,LR"
     End If
        Set RsDisp = Conn.Execute(str)
''        tempwo = ""
        While Not RsDisp.EOF And TempDispQty > 0
            If ChkDIDDispatchedToWo(RsDisp!Work_Order, TxtDID, Trim(TxtMachine), Trim(RsDisp!Slot), Trim(RsDisp!LR)) = False Then
'               MsgBox "The DID can not dispatch to The WO " & RsDisp!Work_Order & ",may be Machine or slot or LR is not match, Please check! "
            Else
''mark by udall
''                If tempwo = "" Or tempwo <> Trim(RsDisp!Work_Order) Then
                   Cboslot.Visible = False
                   LblSlot(4).Visible = False
                   Item = Trim(RsDisp!Item)
                   Machine = Trim(RsDisp!Machine)
                   Slot = Trim(RsDisp!Slot)
                   LR = Trim(RsDisp!LR)
                   BaseQty = RsDisp!BaseQty

'Add by  leimo 20061226
                   If TempDispQty + RsDisp!BalanceQty > 0 Then
                      TempDIDQty = -RsDisp!BalanceQty
                   Else
                      TempDIDQty = TempDispQty
                   End If
                   DispQtybyWO = RsDisp!DispatchQty + TempDIDQty
                   'BalanceQtyByWO = DispQtybyWO - RsDisp!NeedQty
'
                   
                   ' (2) ##########insert into DID dispatch qty to QSMS_Dispatch------dispatch log
                   str = "Select getdate()"
                   Set rsTemp = Conn.Execute(str)
                   TransDateTime = Format(rsTemp(0), "YYYYMMDDHHMMSS")
                   TempDispQty = TempDispQty + RsDisp!DispatchQty - RsDisp!NeedQty

                   str = "exec  QSMSInsertDispatch  '" & RsDisp!Work_Order & "','" & Trim(CboGroupID) & "' ,'" & Trim(CboLine) & "' ,'" & CInt(RsDisp!WOqty) & "','" & Trim(RsDisp!Jobpn) & "' ,'" & Machine & "' " & _
                         " ,'" & TxtCompPN & "' ,'" & Slot & "','" & LR & "','" & CInt(BaseQty) & "'," & CLng((RsDisp!TotalNeedQty)) & " ,'" & Trim(TxtDID) & "'," & CLng(TotalQty) & " ," & TempDIDQty & " " & _
                        ",'" & TxtVendorCode & "' ,'" & TxtDateCode & "','" & TxtLotCode & "','" & g_userName & "' ,'" & TransDateTime & "','" & Trim(DIDDateTime) & "','" & Trim(CboInheritWO) & "','" & Item & "','" & Trim(RsDisp!jobgroup) & "','" & Trim(RsDisp!Side) & "'"

                   Set RsInsert = Conn.Execute(str)
                   If RsInsert.EOF Then
                      MsgBox "Insert into QSMS_Dispatch Error,please retry again"
                      Insert_QSMS_Out = False
                      Exit Sub
                   Else
                      If UCase(Trim(RsInsert.Fields(0))) = "PASS" Then
                            'record the first dispatch time ---20070715
                            If FistTimeofDispatch = 1 Then
                                str = "exec RecordDispatchFDT '" & RsDisp!Work_Order & "'"
                                Conn.Execute (str)
                                FistTimeofDispatch = FistTimeofDispatch + 1
                            End If
                      Else
                        'if error retry again
                         Set RsInsert = Conn.Execute(str)
                         If RsInsert.EOF Then
                           MsgBox "Insert into QSMS_Dispatch Error,please retry again"
                           Insert_QSMS_Out = False
                           Exit Sub
                         Else
                             If UCase(Trim(RsInsert.Fields(0))) = "PASS" Then
                             Else
                                    Insert_QSMS_Out = False
                                    MsgBox "Insert into QSMS_Dispatch Error,please retry again"
                                    Exit Sub
                             End If
                         End If
                    
                     End If

                  End If
              
''                tempwo = Trim(RsDisp!Work_Order)
''                Else
''                   Call GetSlot(TxtDID, tempwo)
''
''                   tempwo = Trim(RsDisp!Work_Order)
''                End If
            End If
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            
           RsDisp.MoveNext
        Wend

rs.MoveNext
     
Wend


'(5)) check if the machine has been dispatch finished
str = "select distinct Work_Order from QSMS_WO where  work_order in " & wostr & " "
Set RsWoarray = Conn.Execute(str)
While Not RsWoarray.EOF
      Call UpdateMachineFlagByWO(RsWoarray!Work_Order)
      RsWoarray.MoveNext
Wend
'(6) check if the item has been dispatch finished
Call ChkWOItemFinished(wostr)

LblInherit.BackColor = &HFF00&
LblInherit.Caption = "Inherit OK!!!"
Exit Sub
EcmdSave_Click:
    MsgBox Err.Description + ",Please contact QSMS SMT Staff"

End Sub

Private Sub CmdQuery_Click()
If Trim(CboLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If
Call GetGroupID
End Sub



Private Sub cmdUnLink_Click()
Dim str As String
Dim rs As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim tempwo As String, TempJobPn As String, TempJObGroup As String, tempSide As String, tempmachine As String, tempPN As String, tempSlot As String, tempLR As String
Dim tempqty As Integer
Dim wostr As String
On Error GoTo EcmdSave_Click
If MsgBox("Are you sure to do Unlink Inherit?!", vbYesNo) = vbNo Then
Exit Sub
End If


If ChkBelongSameGroup(Trim(CboInheritWO)) = False Then
   Exit Sub
End If

LblInherit.BackColor = &HFF&
LblInherit.Caption = "System is Unlink Inheriting,Please wait!!!!!"
wostr = GetWoArray

str = "select * from qsms_dispatch where work_order in " & wostr & " and inherit_wo='" & Trim(CboInheritWO) & "' and didqty>0"
Set rs = Conn.Execute(str)
If rs.EOF Then
   MsgBox "No DID can be Unlink inherit"
   LblInherit.BackColor = &HFF00&
   LblInherit.Caption = ""
   Exit Sub
End If

While Not rs.EOF
    tempwo = Trim(rs!Work_Order)
    tempmachine = Trim(rs!Machine)
    tempPN = Trim(rs!CompPN)
    tempSlot = Trim(rs!Slot)
    tempLR = Trim(rs!LR)
    tempqty = rs!DIDQty
    TempJobPn = Trim(rs!Jobpn)
    TempJObGroup = Trim(rs!jobgroup)
    tempSide = Trim(rs!Side)
    
    str = "update qsms_did set remainqty=remainqty+" & rs!DIDQty & ",usedflag='N',inheritflag='N' where did='" & rs!DID & "' and transdatetime='" & rs!DIDDateTime & "'"
    Conn.Execute (str)
    
    str = "select item from qsms_wo where work_order='" & rs!Work_Order & "' and JobPN='" & TempJobPn & "' and JobGroup='" & TempJObGroup & "' and Side='" & tempSide & "' and machine='" & rs!Machine & "' and comppn='" & rs!CompPN & "' and slot='" & rs!Slot & "' and lr='" & rs!LR & "' "
    Set rsTemp = Conn.Execute(str)
    
    If rsTemp!Item = "0" Then
            str = "Update QSMS_Wo set dispatchqty=dispatchqty-" & tempqty & " ,BalanceQty=BalanceQty-" & tempqty & ",machinefinishedflag='N', wofinishedflag='N' " & _
            "where work_order='" & tempwo & "' and JobPN='" & TempJobPn & "' and JobGroup='" & TempJObGroup & "' and Side='" & tempSide & "' and machine='" & tempmachine & "' and comppn='" & tempPN & "' and slot='" & tempSlot & "' and lr='" & tempLR & "'"
    Else
            str = "Update QSMS_Wo set dispatchqty=dispatchqty-" & tempqty & " ,BalanceQty=BalanceQty-" & tempqty & ",machinefinishedflag='N', wofinishedflag='N' " & _
            "where work_order='" & tempwo & "' and JobPN='" & TempJobPn & "' and JobGroup='" & TempJObGroup & "' and Side='" & tempSide & "' and machine='" & tempmachine & "' and item='" & rsTemp!Item & "' and slot='" & tempSlot & "' and lr='" & tempLR & "'"
    End If
    
    Conn.Execute (str)
rs.MoveNext
Wend
   
str = "delete qsms_dispatch where work_order in " & wostr & " and inherit_wo='" & Trim(CboInheritWO) & "' and didqty>0"
Conn.Execute (str)
    

'(6) check if the item has been dispatch finished
Call ChkWOItemFinished(wostr)

'Add log who do Inherit DID (0002)
str = "Insert into QSMS_LOG(System_name,event_no,DID,user_name,returnQty,trans_date) select '" & App.EXEName & "','UnInherit DID','" & Trim(CboInheritWO) & "','" & g_userName & "',0,DBO.FORMATDATE(GETDATE(),'YYYYMMDDHHNNSS')"
Conn.Execute (str)

LblInherit.BackColor = &HFF00&
LblInherit.Caption = "Unlink Inherit OK!!!"
Exit Sub
EcmdSave_Click:
    MsgBox Err.Description + ",Please contact QSMS SMT Staff"
End Sub

Private Sub DGDIDNotOK_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 On Error Resume Next
        With DGDIDNotOK
             TxtDID.Text = Trim(.Columns(0).Text)
              Call GetDIDInfo(TxtDID, TxtWO)
              Call GetCompDispInfo(Trim(TxtWO), Trim(TxtMachine), Trim(TxtCompPN))
'              Call GetSlot(Trim(TxtDID), Trim(TxtWO))
             
             If Err.Number <> 0 Then
                TxtDID.Text = vbNullString
               
             End If
        End With
        
        
End Sub

Private Sub DGDIDOK_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 On Error Resume Next
        With DGDIDOK
             TxtDID.Text = Trim(.Columns(0).Text)
              Call GetDIDInfo(TxtDID, TxtWO)
              Call GetCompDispInfo(Trim(TxtWO), Trim(TxtMachine), Trim(TxtCompPN))
              'Call GetSlot(Trim(TxtDID), Trim(TxtWO))
             
             If Err.Number <> 0 Then
                TxtDID.Text = vbNullString
               
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
Dim I As Long
Dim rs As ADODB.Recordset
BeginDate = Format(dtpSDate, "YYYY/MM/DD")
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
EndDate = Format(dtpEDate, "YYYY/MM/DD")
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")

If BU = "NB5" Then
    If OptRelease.Value = True Then
       str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )"
    Else
        str = "select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )"
    End If
Else
    If OptRelease.Value = True Then
       str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N'" '- ---(1220)
    Else
        str = "select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N'"  '- ---(1220)
    End If
End If

Set rs = Conn.Execute(str)
I = 0
CboGroupID.Clear
While Not rs.EOF
      CboGroupID.AddItem Trim(rs!GroupID)
      rs.MoveNext
      I = I + 1
Wend
If I = 0 Then
   MsgBox "No data"
   
End If
End Function
Private Function GetGroupWO(ByVal GroupID As String)
Dim str As String
Dim TransDate As String
Dim rs As ADODB.Recordset
CboInheritingWO.Clear

str = "select Work_Order from QSMS_WOGroup  where GroupID= '" & GroupID & "' order by Seq_NO"

Set rs = Conn.Execute(str)
cboWO.Clear
CboNotFinishedWO.Clear
CboNotChkBOM.Clear
CboInheritWO.Clear

While Not rs.EOF
      If ChkMBWo(rs!Work_Order) = True Then
            If ChkQSMS_WO(Trim(rs!Work_Order)) = False Then
                CboNotChkBOM.AddItem Trim(rs!Work_Order)
            Else
            
                 CboInheritWO.AddItem Trim(rs!Work_Order)
                If ChkWoFinished(rs!Work_Order) = True Then
    
                    cboWO.AddItem Trim(rs!Work_Order)
                   
                Else
                     CboInheritingWO.AddItem Trim(rs!Work_Order)
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
str = "select PN, Qty,[Group] from Sap_Wo_List where WO='" & Trim(WO) & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   TxtMBPN = rs!PN
   TxtModel = Mid(TxtMBPN, 3, 3)
   TxtWOQty = rs!Qty
   TxtGroup = Trim(rs![Group])
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
str = "select distinct Line from QSMS_woGroup"
Set rs = Conn.Execute(str)
CboLine.Clear
While Not rs.EOF
    CboLine.AddItem rs!Line
    rs.MoveNext
Wend
End Function



Private Function GetConsumeQty(WO, CompPN) As Long
Dim str As String
Dim rs As ADODB.Recordset
str = "select ConsumedQty from SMT_QSMS_Out where Work_Order='" & Trim(WO) & "' and CompPN='" & CompPN & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   GetConsumeQty = Trim(rs!ConsumedQty)
End If


End Function

Private Function ChkErr() As Boolean
ChkErr = True
If Trim(CboInheritWO) = "" Then
   MsgBox "Please select the work order which you want to inherit from "
   ChkErr = False
End If
If Trim(CboInheritingWO) = "" Then
    MsgBox "Please select the work order which you want to inherit from "
    ChkErr = False
End If

End Function

Public Function RefreshDID_Machine_WO(ByVal RefreshType As String, ByVal CompPN As String, ByVal Machine As String, ByVal WO As String, ByVal MBPN As String, ByVal Line As String)
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
Dim str As String
Dim rs As ADODB.Recordset


'Str = "select b.DID from QSMS_WO a, QSMS_DID b where (a.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "')  or a.work_order in " & GetWoArray & ") " & _
       "and a.Machine='" & Trim(Machine) & "' and a.CompPN=b.CompPN and a.balanceQty>=0  And a.dispatchqty > 0 and b.UsedFlag='N'" & _
      " Union all select distinct b.DID from QSMS_dispatch a, QSMS_DID b where (a.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') or a.work_order in " & GetWoArray & ") " & _
      "and a.Machine='" & Trim(Machine) & "' and a.DID=b.DID and  b.UsedFlag='Y' Order by B.DID"
str = "select distinct b.DID from QSMS_dispatch a, QSMS_DID b where (a.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') or a.work_order in " & GetWoArray & ") " & _
      "and a.Machine='" & Trim(Machine) & "' and a.DID=b.DID Order by B.DID"
Set rs = Conn.Execute(str)
DGDIDOK.Caption = "DID OK :" & rs.RecordCount
Set DGDIDOK.DataSource = rs


str = "select b.DID from QSMS_WO a, QSMS_DID b where (a.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') or a.work_order in " & GetWoArray & ")" & _
      "and a.Machine='" & Trim(Machine) & "' and a.CompPN=b.CompPN and a.balanceQty<0 and b.UsedFlag<>'Y' order by b.DID"
Set rs = Conn.Execute(str)
DGDIDNotOK.Caption = "DID Not OK :" & rs.RecordCount
Set DGDIDNotOK.DataSource = rs
DGDIDNotOK.Refresh

str = "select a.CompPN from QSMS_NonAVL a,QSMS_Wo b where (B.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') or b.work_order in " & GetWoArray & ") " & _
     "and a.CompPN=b.CompPN and a.Customer='" & Trim(TxtCustomer) & "' and Model='" & Trim(TxtModel) & "'"
Set rs = Conn.Execute(str)
Set DGAVL.DataSource = rs
DGAVL.Refresh
str = "Select  a.CompPn,-sum(balanceQty) as NeedQty from QSMS_WO a where (a.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') or a.work_order in " & GetWoArray & ") " & _
      "and a.Machine='" & Trim(Machine) & "' and a.balanceQty<0 group by a.comppN"
Set rs = Conn.Execute(str)
DGCompNotOK.Caption = "Comp didn't dispatch:" & rs.RecordCount
Set DGCompNotOK.DataSource = rs
DGCompNotOK.Refresh
'While Not Rs.EOF
'    If ChkAVL(Rs!DID) = False Then
'
'        CboAVL.AddItem Trim(Rs!DID)
'    End If
'    Rs.MoveNext
'Wend

'Str = "select b.DIDfrom QSMS_WO a, QSMS_DID b where a.Work_Order='" & Trim(WO) & "'  and a.Machine='" & Trim(Machine) & "' and a.CompPN=b.CompPN and a.balanceQty>=0  And a.dispatchqty > 0 and b.UsedFlag='N'"
'Set Rs = Conn.Execute(Str)
'CboDIDOK.Clear
'While Not Rs.EOF
'      CboDIDOK.AddItem Trim(Rs!DID)
'      Rs.MoveNext
'Wend
'
'Str = "select b.DID from QSMS_WO a, QSMS_DID b where a.Work_Order='" & Trim(WO) & "'  and a.Machine='" & Trim(Machine) & "' and a.CompPN=b.CompPN and  b.UsedFlag='Y'"
'Set Rs = Conn.Execute(Str)
'
'While Not Rs.EOF
'       CboDIDOK.AddItem Trim(Rs!DID)
'       Rs.MoveNext
'Wend


'Str = "select b.DID,B.Qty from QSMS_WO a, QSMS_DID b where a.Work_Order='" & Trim(WO) & "'  and a.Machine='" & Trim(Machine) & "' and a.CompPN=b.CompPN and a.balanceQty<0 and b.UsedFlag='N'"
'Set Rs = Conn.Execute(Str)
'
'CboDIDNOK.Clear
'While Not Rs.EOF
'    If ChkAVL(Rs!DID) = True Then
'        CboDIDNOK.AddItem Trim(Rs!DID)
'    Else
'        CboAVL.AddItem Trim(Rs!DID)
'    End If
'    Rs.MoveNext
'Wend



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
         If ChkBalanceQty(Machine, WO, rs!CompPN) = False Then
           CboWithout.AddItem Trim(rs!CompPN)
         End If
      Else
         str = "select CompPN from  QSMS_WO where Work_Order='" & Trim(WO) & "' and machine='" & Machine & "' and Item='" & Trim(rs!Item) & "' and CompPN  in (select CompPN from QSMS_DID where UsedFlag='N')"
         Set TempRs = Conn.Execute(str)
         If TempRs.EOF Then
             If ChkBalanceQty(Machine, WO, rs!CompPN) = False Then
                CboWithout.AddItem Trim(rs!CompPN)
             End If
         End If
      End If
      rs.MoveNext
Wend
End Function
Private Function ChkBalanceQty(ByVal Machine As String, ByVal WO As String, ByVal CompPN As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
ChkBalanceQty = True
str = "select Work_Order From QSMS_WO where Work_Order='" & WO & "' and Machine='" & Machine & "' and CompPN='" & CompPN & "' and BalanceQty<0"
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

TxtDateCode = ""
TxtLotCode = ""
TxtDIDRemainQty = ""

'Str = "select a.DID,a.CompPN,a.VendorCode,a.DateCode,a.LotCode,a.Qty,a.RemainQty,a.UsedFlag,b.BaseQty,b.DispatchQty from QSMS_DID a,QSMS_WO b " & _
      " where a.DID='" & Trim(DID) & "' and a.CompPN=b.CompPN and b.Work_Order='" & WO & "' and b.machine='" & Trim(TxtMachine) & "'"
str = "select DID,CompPN,VendorCode,DateCode,LotCode,Qty,RemainQty,UsedFlag,DIDLoc from QSMS_DID where DID='" & Trim(TxtDID) & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   TxtCompPN = Trim(rs!CompPN)
   TxtVendorCode = Trim(rs!VendorCode)
   TxtDateCode = Trim(rs!DateCode)
   TxtLotCode = Trim(rs!LotCode)
   TxtRackID = Trim(rs!DIDLoc)
   TxtDIDTotalQty = Trim(rs!Qty)
   TxtDIDRemainQty = Trim(rs!RemainQty)

    If UCase(Trim(rs!usedflag)) = "Y" Then

       TxtDispatchQty.Enabled = False
       TxtDispatchQty = 0
   Else

       TxtDispatchQty.Enabled = True
       
   End If
   
   
   
  
End If
End Function
Private Function GetCompDispInfo(ByVal WO As String, ByVal Machine As String, ByVal CompPN As String)
Dim str As String
Dim rs As ADODB.Recordset
Dim WoArray As String
TxtConsumedQty = ""
TxtCompBaseQty = ""
TxtNeedQty = ""
WoArray = GetWoArray
str = "Select sum(BaseQty) as BaseQty, sum(NeedQty) as NeedQty ,sum(DispatchQty) as DispatchQty from QSMS_WO where (Work_Order " & _
       " in (select WO from sap_wo_list where [group]='" & Trim(TxtGroup) & "' )or work_order in " & WoArray & ") and machine='" & Machine & "' and comppn='" & CompPN & "' "
Set rs = Conn.Execute(str)
If Not rs.EOF Then
    TxtCompBaseQty = Trim(rs!BaseQty)
    TxtConsumedQty = Trim(rs!DispatchQty)
    TxtNeedQty = Trim(rs!NeedQty)
End If
If Trim(TxtNeedQty) = "" Then
  Exit Function
End If
If CLng(Trim(TxtDIDRemainQty)) > CLng(Trim(TxtNeedQty)) Then
       TxtDispatchQty = TxtNeedQty
Else
       TxtDispatchQty = TxtDIDRemainQty
End If
End Function
'Private Function GetSlot(ByVal DID As String, ByVal WO As String)
'Dim Str As String
'Dim Rs As ADODB.Recordset
'Dim Used_Flag As String
'Dim Slot As String
'Dim i As Long
'LblSlot(4).Visible = True
'Cboslot.Visible = True
'LblSlot(4).BackColor = &HFF&
'
'
'Cboslot.Clear
'Slot = ""
'i = 0
'Str = "select b.Slot from QSMS_DID a,QSMS_WO b where a.DID='" & Trim(DID) & "' and a.CompPN=b.CompPN and (b.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') or b.work_order in " & GetWoArray & ") " & _
'      "and b.Machine='" & Trim(Machine) & "'"
'Set Rs = Conn.Execute(Str)
'While Not Rs.EOF
'      Slot = Trim(Rs!Slot)
'      Cboslot.AddItem Trim(Rs!Slot)
'      Rs.MoveNext
''      If Slot <> "" Then
''        I = I + 1
''      End If
'Wend
'''If I = 1 Then
'''   CboSlot.Text = Slot
'''   CboSlot.Enabled = False
''   Call GetCompDispInfo(Trim(TxtWO), Trim(TxtMachine), Trim(TxtCompPN))
'''Else
'''  CboSlot.Enabled = True
''
'''End If
'End Function



Private Function GetMachine(ByVal WO As String)
Dim str As String
Dim TransDate As String
Dim rs As ADODB.Recordset
Dim rsMachine As ADODB.Recordset
str = "select distinct Machine from QSMS_WO where Work_Order in (select Wo from Sap_Wo_List where [Group]='" & TxtGroup & "') and machine like '" & Trim(CboLine) & "%'"

Set rs = Conn.Execute(str)
CboMathineOK.Clear
CboMathineNOK.Clear
While Not rs.EOF
     str = "select  Machine from QSMS_WO where Work_Order in (select Wo from Sap_Wo_List where [Group]='" & TxtGroup & "') and MachinefinishedFlag='N' and machine='" & Trim(rs!Machine) & "'"
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
   MsgBox "The Work order has No group ID, please call PMC "
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
Dim strSql As String

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
   End If
Else
'''''''''''''''''''''''''''''''''''''add by Jing 2007.10.31------------(0001)'''''''''''''''''''''''''''''''''''''''''''''''
    strSql = "select * from qsms_did_log where did='" & DID & "'"
    Set rs = Conn.Execute(strSql)
    If rs.RecordCount > 0 Then
        MsgBox "This DID had been deleted !"
    Else
        MsgBox "找不到DID，请检查"
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

End Function

Private Sub TxtChkDID_KeyPress(KeyAscii As Integer)
Dim str As String
Dim rs As ADODB.Recordset
If KeyAscii = 13 Or KeyAscii = 9 Then
   ListDIDStatus.Clear
   'str = "select b.DID,B.Qty ,B.UsedFlag from QSMS_WO a, QSMS_DID b where a.Work_Order='" & Trim(TxtWO) & "'  and a.Machine='" & Trim(TxtMachine) & "' and a.CompPN=b.CompPN  and b.DID='" & TxtChkDID & "' "
   str = "select b.DID,B.Qty ,B.UsedFlag from QSMS_WO a, QSMS_DID b where (a.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') or a.work_order in " & GetWoArray & ") " & _
         "and a.Machine='" & Trim(TxtMachine) & "' and a.CompPN=b.CompPN  and b.DID='" & TxtChkDID & "' "
   
   Set rs = Conn.Execute(str)
   If Not rs.EOF Then
      If UCase(rs!usedflag) = "Y" Then
         ListDIDStatus.AddItem "Has been Dispatched"
      Else
         ListDIDStatus.AddItem "Not Dispatched"
      End If
   Else
      
         ListDIDStatus.AddItem "Not belong to the Machine"
   End If
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

End Sub

Private Sub txtDID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
    
   If ChkDIDBelongMachine(Trim(TxtWO), Trim(TxtMachine), Trim(TxtDID) & "", Trim(TxtGroup)) = False Then
     TxtDID.Text = ""
     TxtDID.SetFocus
     Exit Sub

   End If
   Call GetDIDInfo(Trim(TxtDID), Trim(TxtWO))
   Call GetWoArray
   Call GetCompDispInfo(Trim(TxtWO), Trim(TxtMachine), Trim(TxtCompPN))
   TxtDispatchQty.SetFocus
  ' Call GetSlot(Trim(TxtDID), Trim(TxtWO))
End If
End Sub
Private Function ResetTxt()

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

Private Function ChkDIDDispatchedToWo(ByVal WO As String, ByVal DID As String, ByVal Machine As String, ByVal Slot As String, LR As String) As Boolean
Dim str As String
Dim SqlConfig As String
Dim rs As ADODB.Recordset
Dim Rssql As ADODB.Recordset
Dim DeleteFlag As Boolean
DeleteFlag = False
ChkDIDDispatchedToWo = False

 '0005 '承接料的时候，Flag:InheritDIDBySlot开启 如果DID之前没有发过这个位置，则这个DID不能承接到这个位置，否则按照DID来承接 1218
SqlConfig = "select 0 from QSMS_ProConfig where Station = 'QSMS' and [Key] = 'InheritDIDBySlot' and Value = 'Y'"
Set Rssql = Conn.Execute(SqlConfig)
If Not Rssql.EOF Then
    str = "select DID,DeletedFlag,Machine,Slot,LR from QSMS_Dispatch  where DID='" & DID & "' and Slot= '" & Slot & "' and LR= '" & LR & "' "
Else
    str = "select DID,DeletedFlag,Machine,Slot,LR from QSMS_Dispatch  where DID='" & DID & "' "
End If

'Str = "select DID,DeletedFlag,Machine,Slot,LR from QSMS_Dispatch  where DID='" & DID & "' and Slot= '" & Slot & "' and LR= '" & LR & "' "
Set rs = Conn.Execute(str)
If rs.EOF Then
    ChkDIDDispatchedToWo = True
'   ChkDIDDispatchedToWo = False    最初的版本应该是True！
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

Private Function ChkDIDCompPN(ByVal DID As String, ByVal CompPN As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
ChkDIDCompPN = True
str = "select DID from QSMS_DID where DID='" & DID & "' and CompPN='" & CompPN & "'"
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
'Str = "select WO from Sap_Wo_list where wo='" & WO & "' and (PN like '21%' or PN like '31%')"
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
Private Function GetWoArray() As String
Dim WoArray As String
Dim str As String
Dim rs As ADODB.Recordset
Dim I As Long
    If CboInheritingWO = "" Then
       WoArray = "('')"
       GetWoArray = WoArray
       Exit Function   ''Add by Udall  07/03/13
    End If

        str = "select wo from Sap_WO_List where [Group] in (select [group] from sap_wo_list where wo='" & CboInheritingWO.Text & "')"
        Set rs = Conn.Execute(str)
        While Not rs.EOF
               WoArray = WoArray + "'" + Trim(rs!WO) + "'" + ","
               rs.MoveNext
        Wend
        WoArray = WoArray + "'" + CboInheritingWO.Text + "'" + ","

    WoArray = Mid(WoArray, 1, Len(WoArray) - 1)
    WoArray = "(" + WoArray + ")"
    GetWoArray = WoArray
End Function

Private Function GetSBWO(ByVal WO As String)
Dim str As String
Dim rs As ADODB.Recordset
Dim I As Long
Dim Group As String
I = 0
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
     I = I + 1
Wend
If I > 0 Then
    FraSB.Visible = True

End If
End Function



Private Sub TxtQryMachine_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   
   Call CmdExcelDID_Click
End If
End Sub

Private Sub TxtQryWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   TxtQryMachine.SetFocus
   Call cmdExcel_Click
End If
End Sub

Private Function ChkBelongSameGroup(ByVal InheritWO As String) As Boolean
Dim rs As ADODB.Recordset
Dim str As String
ChkBelongSameGroup = True
str = "Select GroupID from QSMS_WoGroup where Work_Order='" & InheritWO & "'"
Set rs = Conn.Execute(str)
If rs.EOF Then
   ChkBelongSameGroup = False
   MsgBox "Can not find the group,Please check"
   Exit Function
Else
   If UCase(Trim(rs!GroupID)) <> UCase(Trim(CboGroupID)) Then
      ChkBelongSameGroup = False
      MsgBox "The Inherit Wo doesn't belong to the GroupID:" & Trim(CboGroupID)
      Exit Function
   End If
End If

str = "select work_order from QSMS_WoGroup where work_order='" & CboInheritingWO.Text & "' and GroupID='" & Trim(CboGroupID) & "'"
Set rs = Conn.Execute(str)
If rs.EOF Then
   ChkBelongSameGroup = False
   MsgBox "The inheriting Wo does not belong to the GroupID :" & Trim(CboGroupID)
   
End If
End Function


