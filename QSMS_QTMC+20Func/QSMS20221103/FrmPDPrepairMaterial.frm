VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FrmPDPrepairMaterial 
   Caption         =   "frm PD prepair material 14-09-03"
   ClientHeight    =   10470
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraQyeryDID 
      BackColor       =   &H00FFC0C0&
      Height          =   1455
      Left            =   120
      TabIndex        =   72
      Top             =   8880
      Width           =   15375
      Begin VB.Frame Frame5 
         Caption         =   "QueryDIDByWO Machine"
         Height          =   855
         Left            =   7680
         TabIndex        =   77
         Top             =   120
         Width           =   7335
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
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   360
            Width           =   1815
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
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   360
            Width           =   1815
         End
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
            Picture         =   "FrmPDPrepairMaterial.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   360
            Width           =   735
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
            TabIndex        =   82
            Top             =   360
            Width           =   855
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
            TabIndex        =   80
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "QueryWoBy DID"
         Height          =   855
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   7455
         Begin VB.CheckBox ChkAll 
            Caption         =   "By DID"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   360
            Width           =   855
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
            Picture         =   "FrmPDPrepairMaterial.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
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
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   360
            Width           =   3255
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
            TabIndex        =   74
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "PD prepair material"
      Height          =   5655
      Left            =   120
      TabIndex        =   10
      Top             =   3240
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
         Left            =   4680
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
      End
      Begin MSDataGridLib.DataGrid DGSlot 
         Height          =   1575
         Left            =   10200
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   2778
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
         Caption         =   "Dispatch Qty By slot"
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
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh Machine"
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
         Left            =   13920
         Picture         =   "FrmPDPrepairMaterial.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   120
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000000&
         Caption         =   "DID information "
         Height          =   1095
         Left            =   120
         TabIndex        =   63
         Top             =   4560
         Width           =   15015
         Begin VB.TextBox TxtDIDDateTime 
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
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   600
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
            TabIndex        =   97
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
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   240
            Width           =   2295
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
            TabIndex        =   66
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
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
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
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "DateaTime"
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
            Index           =   25
            Left            =   120
            TabIndex        =   101
            Top             =   600
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
            TabIndex        =   96
            Top             =   240
            Width           =   1215
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
            TabIndex        =   71
            Top             =   240
            Width           =   1095
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
            TabIndex        =   70
            Top             =   240
            Width           =   1455
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
            TabIndex        =   69
            Top             =   240
            Width           =   1335
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
            TabIndex        =   68
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame FraDispatchDID 
         BackColor       =   &H00FF80FF&
         Caption         =   "DispatchDID"
         Height          =   1215
         Left            =   120
         TabIndex        =   46
         Top             =   3360
         Width           =   15015
         Begin VB.TextBox TxtTotalQty 
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
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox TxtBalanceQty 
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
            Left            =   8760
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   720
            Width           =   1095
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
            Left            =   14640
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
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
            Left            =   7800
            Picture         =   "FrmPDPrepairMaterial.frx":1016
            Style           =   1  'Graphical
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   120
            Width           =   975
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
            Left            =   3840
            TabIndex        =   55
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
            Left            =   6480
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   720
            Width           =   975
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
            Left            =   14040
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   720
            Width           =   855
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
            Left            =   14760
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
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
            Left            =   11280
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   720
            Width           =   975
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
            Left            =   6240
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
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
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   240
            Width           =   3855
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "T_Qty"
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
            Index           =   30
            Left            =   0
            TabIndex        =   120
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Balance Qty"
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
            Index           =   26
            Left            =   7440
            TabIndex        =   109
            Top             =   720
            Width           =   1335
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
            Left            =   14760
            TabIndex        =   91
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label LblMessage 
            BackColor       =   &H00FFFFC0&
            Height          =   495
            Left            =   8880
            TabIndex        =   62
            Top             =   120
            Width           =   6015
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Plan Need Qty"
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
            Left            =   2280
            TabIndex        =   60
            Top             =   720
            Width           =   1575
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
            Left            =   4920
            TabIndex        =   59
            Top             =   720
            Width           =   1575
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
            Left            =   12360
            TabIndex        =   58
            Top             =   720
            Width           =   1695
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
            TabIndex        =   57
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
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
            Left            =   9840
            TabIndex        =   56
            Top             =   720
            Width           =   1455
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
            Left            =   4800
            TabIndex        =   50
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
            TabIndex        =   49
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame FraWithoutDID 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Without DID"
         Height          =   735
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   15255
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
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
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
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   240
            Width           =   1095
         End
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
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   240
            Width           =   4215
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
            TabIndex        =   45
            Top             =   240
            Width           =   735
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
            TabIndex        =   44
            Top             =   240
            Width           =   855
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
            TabIndex        =   40
            Top             =   240
            Width           =   1695
         End
      End
      Begin MSDataGridLib.DataGrid DGAVL 
         Height          =   1815
         Left            =   4680
         TabIndex        =   33
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFF80&
         Caption         =   "ChkDID Status"
         Height          =   1815
         Left            =   7080
         TabIndex        =   30
         Top             =   1440
         Width           =   3015
         Begin VB.ListBox ListDIDStatus 
            Height          =   1035
            ItemData        =   "FrmPDPrepairMaterial.frx":1320
            Left            =   120
            List            =   "FrmPDPrepairMaterial.frx":1327
            TabIndex        =   32
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
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   240
            Width           =   2535
         End
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
         Left            =   12240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
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
         Left            =   9000
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   2055
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
         ItemData        =   "FrmPDPrepairMaterial.frx":133A
         Left            =   1560
         List            =   "FrmPDPrepairMaterial.frx":133C
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "CboMathineNOK"
         Top             =   240
         Width           =   1935
      End
      Begin MSDataGridLib.DataGrid DGCompNotOK 
         Height          =   1815
         Left            =   240
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3201
         _Version        =   393216
         DefColWidth     =   96
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
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1154.835
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "JobGroup"
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
         Left            =   3600
         TabIndex        =   115
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
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
         Left            =   11160
         TabIndex        =   24
         Top             =   240
         Width           =   975
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
         Left            =   7680
         TabIndex        =   20
         Top             =   240
         Width           =   1335
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
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin VB.TextBox TxtPlanQty 
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
         Height          =   480
         Left            =   10320
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox TxtLine 
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
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   2040
         Width           =   735
      End
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
         Left            =   13800
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox CboClosed 
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
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   480
         Width           =   2655
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
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ListBox ListWoDispatching 
         Height          =   1230
         ItemData        =   "FrmPDPrepairMaterial.frx":133E
         Left            =   13200
         List            =   "FrmPDPrepairMaterial.frx":1340
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   600
         Width           =   2055
      End
      Begin VB.ListBox ListWoNotFinish 
         Height          =   1425
         ItemData        =   "FrmPDPrepairMaterial.frx":1342
         Left            =   10320
         List            =   "FrmPDPrepairMaterial.frx":1344
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   600
         Width           =   2175
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
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   480
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
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   840
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
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   1200
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
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   1560
         Width           =   495
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
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Frame FraSB 
         Caption         =   "Small Board WO"
         Height          =   615
         Left            =   6600
         TabIndex        =   34
         Top             =   1920
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
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
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
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1335
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
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2520
         Width           =   975
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
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2295
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1560
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
         Height          =   360
         Left            =   6600
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "CboGroupID"
         Top             =   120
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
         Left            =   3360
         Picture         =   "FrmPDPrepairMaterial.frx":1346
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptRelease 
         Caption         =   "Release"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
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
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2520
         Width           =   735
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2055
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
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
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
         Format          =   134545411
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   94
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
         Format          =   134545411
         CurrentDate     =   36482
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000FF&
         Caption         =   "Plan Qty"
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
         Index           =   29
         Left            =   9240
         TabIndex        =   119
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
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
         Index           =   28
         Left            =   11040
         TabIndex        =   116
         Top             =   2040
         Width           =   735
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
         Left            =   12480
         TabIndex        =   112
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF80FF&
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
         Index           =   6
         Left            =   4440
         TabIndex        =   105
         Top             =   480
         Width           =   2175
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
         TabIndex        =   99
         Top             =   840
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
         TabIndex        =   95
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
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
         TabIndex        =   90
         Top             =   1560
         Width           =   2175
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
         TabIndex        =   89
         Top             =   120
         Width           =   2175
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
         TabIndex        =   36
         Top             =   2520
         Width           =   1215
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
         TabIndex        =   28
         Top             =   2520
         Width           =   735
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
         TabIndex        =   26
         Top             =   2520
         Width           =   1095
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
         TabIndex        =   22
         Top             =   2520
         Width           =   1455
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
         TabIndex        =   18
         Top             =   120
         Width           =   2175
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
         TabIndex        =   16
         Top             =   120
         Width           =   2175
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
         TabIndex        =   8
         Top             =   2520
         Width           =   735
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
         TabIndex        =   7
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
         Left            =   7920
         TabIndex        =   5
         Top             =   2520
         Width           =   1095
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
         TabIndex        =   3
         Top             =   1200
         Width           =   2175
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
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
   End
   Begin MSDataGridLib.DataGrid DGDIDNotOK 
      Height          =   1815
      Left            =   20280
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   7200
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
      Left            =   17880
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   7200
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MCI.MMControl wave_control 
      Height          =   450
      Left            =   0
      TabIndex        =   111
      Top             =   0
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   794
      _Version        =   393216
      PlayEnabled     =   -1  'True
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "FrmPDPrepairMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************************************************************
'**文 件 名: FrmPDPrepairMaterial.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: Sandy
'**日    期: 2007.10.26
'**描    述: QSMS Prepair Material
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**Sandy      2007.10.27     add check DID and CompPN if matched --------(0001)
'**Jing       2007.10.31     Add the message in detail in the dispatch interface for deleted DID -------(0002)
'**Udall      2007.11.05     Add check DID,the DID can't be dispatched to different line and side --------(0003)
'**Kane       2007.11.15     Modify bug add rule a.diddatetime=b.diddatetime--------(0004)
'**Sandy      2007.11.20     Add slot in WO group but not only machine name---------(0005)
'**Sandy      2007.11.23     Add Choose WO input priority by MCC.---------(0006)
'**Steven     2008.02.21     Add check IPQC Test flag ---------(0007)
'**Giant      2008.03.05     add query dispatch data by comppn ---------(0008)
'**Lynn       2008.05.22     do not allow dispatch if the work order has been closed ---------(0009)
'**Sandy      2008.05.29     Query WO dispatch infromation from live DB and histroy DB---------(0010)
'**Kane       2009.08.20     NB3创建的DID要在NB2发料，所以NB2输入DID后如果没有就到NB3把DID转到NB2系统中(0011)
'**Archer     2009/12/04     Query PrepairMaterial Information use SP: QSMS_PrepairMaterial (0012)
'**Scofield   2010/04/18     Add checking IPQC before dispatch in MBU (0013)
'**Austin     2010/09/06     Confirm with Scofield,marked his 0013 Code (0014)
'***********************************************************************************/
Option Explicit
Dim strGetDIDFromSourceBU As String

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
'Call GetCompPnWithoutDID(Trim(TxtMachine), Trim(TxtWO))
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
Call GetDID(Trim(TxtMachine), Trim(CboJob), TxtWO, CboLine)
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
Dim wostr As String
TxtWO = Trim(CboNotFinishedWO)
Call GetSBWO(TxtWO)
Call GetWoinfo(TxtWO)
wostr = GetWoArray
Call GetMachine(TxtWO, wostr)

End Sub

Private Sub CboNotFinishedWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboNotFinishedWO_Click
End If
End Sub

Private Sub CboWithout_Click()
Dim str As String
Dim Rs As ADODB.Recordset
str = "Select sum(NeedQty) as NeedQty,sum(BalanceQty) as BalanceQty from QSMS_Wo where Work_order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') and Machine='" & Trim(TxtMachine) & "' and CompPN='" & Trim(CboWithout) & "' and BalanceQty<0"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   TxtWTotal.Text = Trim(Rs!NeedQty)
   TxtWBalance.Text = Trim(Rs!BalanceQty)
End If
End Sub

Private Sub CboWo_Click()
Dim wostr As String

TxtWO = Trim(cboWO)
Call GetSBWO(TxtWO)
Call GetWoinfo(TxtWO)
wostr = GetWoArray
Call GetMachine(TxtWO, wostr)
End Sub

Private Sub cboWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboWo_Click
End If
End Sub

Private Sub CmdADD_Click()
    Dim Pointer As Long
    If ListWoNotFinish.ListCount <= 0 Then Exit Sub
    If ListWoNotFinish.ListIndex < 0 Then Exit Sub
    Pointer = ListWoNotFinish.ListIndex
    ListWoDispatching.AddItem Trim(ListWoNotFinish.Text)
    ListWoNotFinish.RemoveItem Pointer
    If ListWoNotFinish.ListCount <> Pointer Then
       ListWoNotFinish.ListIndex = Pointer
    End If
   
End Sub

Private Sub cmdADDALL_Click()

    If ListWoNotFinish.ListCount <= 0 Then Exit Sub

    Do While ListWoNotFinish.ListCount > 0
     
      ListWoNotFinish.ListIndex = 0
      ListWoDispatching.AddItem Trim(ListWoNotFinish.Text)
      ListWoNotFinish.RemoveItem 0
     
    Loop
   
End Sub

Private Sub cmdDel_Click()
    Dim Pointer As Long
    If ListWoDispatching.ListCount <= 0 Then Exit Sub
    If ListWoDispatching.ListIndex < 0 Then Exit Sub
    Pointer = ListWoDispatching.ListIndex

    ListWoNotFinish.AddItem Trim(ListWoDispatching.Text)
    ListWoDispatching.RemoveItem Pointer
    If ListWoDispatching.ListCount <> Pointer Then
       ListWoDispatching.ListIndex = Pointer
    End If

End Sub

Private Sub cmdDELALL_Click()
    If ListWoDispatching.ListCount <= 0 Then Exit Sub
    Do While ListWoDispatching.ListCount > 0
        ListWoDispatching.ListIndex = 0
       
        ListWoNotFinish.AddItem Trim(ListWoDispatching.Text)
        ListWoDispatching.RemoveItem 0
    Loop
    
End Sub

Private Sub CmdConfirm_Click()
On Error GoTo EcmdSave_Click

If ChkErr = True Then
   If Insert_QSMS_Out = True Then
      Call OK_Sound
      LblMessage.Caption = "insert OK"
   Else
      Call Warning_Sound
      LblMessage.Caption = "didn't dispatch the DID"
   End If
   
Else
   Call Warning_Sound
End If
CboMathineNOK = TxtMachine
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
 Dim Rs As ADODB.Recordset

 Dim strFileName, Trans_Date As String

    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add

    xlApp.DisplayAlerts = False
    Set xlWs = xlApp.Worksheets(1)
  
    xlApp.UserControl = True

    If ChkAll.Value = 0 Then
        If Trim(TxtWO.Text) = "" Then   '--(0008)
            str = "Select * from QSMS_Dispatch where CompPN='" & Trim(TxtQryDID) & "'"
        Else
            str = "Select * from QSMS_Dispatch where CompPN='" & Trim(TxtQryDID) & "'  and work_order='" & TxtWO & "'"
        End If
    Else
        str = "Select a.Work_Order,a.GroupID,a.Line,a.WoQty,a.JobPN,a.Machine,a.CompPN,a.Slot,a.LR,a.BaseQTY,a.NeedQty," & _
            "a.DID,a.TotalQty,a.DIDQty,a.VendorCode,a.DateCode,a.LotCode,a.UID,a.TransDateTime,a.DIDDateTime,a.DeletedFlag," & _
            "a.Inherit_wo,a.JobGroup,a.Side,b.ReturnFlag,b.UID,b.TransDateTime from QSMS_Dispatch a left join QSMS_GroupDID b " & _
            "on a.did=b.did and a.diddatetime=b.diddatetime where a.did='" & Trim(TxtQryDID) & "'"  '------(0004)
    
    End If
    Set Rs = Conn.Execute(str)
    
    fldCount = Rs.Fields.Count
  
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = Rs.Fields(iCol - 1).Name
    Next
        
    xlWs.Cells(2, 1).CopyFromRecordset Rs

    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlApp.Visible = True
  
    Rs.Close
    Set Rs = Nothing
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
 Dim Rs As ADODB.Recordset

 Dim strFileName, Trans_Date As String

    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add

    xlApp.DisplayAlerts = False
    Set xlWs = xlApp.Worksheets(1)
  
    xlApp.UserControl = True
'(1) first Get DID infomation---(0010)
    str = "Select * from QSMS_Dispatch where Work_Order='" & Trim(TxtQryWO) & "' and machine like '" & Trim(TxtQryMachine) & "%'" & _
            "union Select * from [QSMS_History].dbo.QSMS_Dispatch where Work_Order='" & Trim(TxtQryWO) & "' and machine like '" & Trim(TxtQryMachine) & "%'"
    Set Rs = Conn.Execute(str)
    
    fldCount = Rs.Fields.Count
  
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = Rs.Fields(iCol - 1).Name
    Next
        
    xlWs.Cells(2, 1).CopyFromRecordset Rs

    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlApp.Visible = True
  
    Rs.Close
    Set Rs = Nothing
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

Private Sub CmdRefresh_Click()
Dim wostr As String
Dim WoSPStr
Dim str As String
'(1) update machine and wo dispatch flag
wostr = GetWoArray
WoSPStr = Replace(wostr, "'", "")
WoSPStr = Mid(WoSPStr, 2, Len(WoSPStr) - 2)
str = "Exec UpdateDispatchFlag '" & Trim(TxtMachine) & "','" & WoSPStr & "'"
Conn.Execute str
'refresh machine --show cbobox
 Call GetMachine(Trim(TxtWO), wostr)
'refresh wo --show in cbobox
 Call GetGroupWO(Trim(CboGroupID))

End Sub



Private Sub DGCompNotOK_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
        With DGCompNotOK
             txtCompPN.Text = Trim(.Columns(0).Text)
'              Call GetDIDInfo(TxtDID, TxtWO)
'              Call GetCompDispInfo(Trim(TxtWO), Trim(TxtMachine), Trim(TxtCompPN))
'              Call GetSlot(Trim(TxtDID), Trim(TxtWO))
               Call GetDispatchQtyBySlot(Trim(TxtWO), Trim(TxtMachine), Trim(txtCompPN))
             
             If Err.Number <> 0 Then
                txtDID.Text = vbNullString
               
             End If
        End With
        
End Sub

Private Sub DGDIDNotOK_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
 On Error Resume Next
        With DGCompNotOK
             txtDID.Text = Trim(.Columns(0).Text)
              Call GetDIDInfo(txtDID, TxtWO)
              Call GetCompDispInfo(Trim(TxtWO), Trim(CboJob), Trim(TxtMachine), Trim(txtCompPN))
'              Call GetSlot(Trim(TxtDID), Trim(TxtWO))
             
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
              Call GetCompDispInfo(Trim(TxtWO), Trim(CboJob), Trim(TxtMachine), Trim(txtCompPN))
              'Call GetSlot(Trim(TxtDID), Trim(TxtWO))
             
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
Dim Rs As ADODB.Recordset
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
str = "select getdate()"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
    'Date = Rs(0)
    'Time = Rs(0)
End If
dtpSDate = Date
dtpEDate = Date
Call GetLine
strGetDIDFromSourceBU = ReadIniFile("QSMS", "GetDIDFromSourceBU", App.Path & "\set.ini")
End Sub

'This not be used --20070529
Private Function ListWO()
Dim str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset
TransDate = Format(dtpSDate, "YYYY/MM/DD")
TransDate = Replace(TransDate, "-", "/")
str = "select distinct WO from Sap_Wo_list where Trans_Date like '" & TransDate & "%'"
Set Rs = Conn.Execute(str)
cboWO.Clear
While Not Rs.EOF
      cboWO.AddItem Trim(Rs!WO)
      Rs.MoveNext
Wend
End Function

Private Function GetGroupID()
Dim str As String
Dim BeginDate, EndDate As String
Dim GroupIDHead As String
Dim I As Long
Dim Rs As ADODB.Recordset
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
Set Rs = Conn.Execute(str)
I = 0
CboGroupID.Clear
While Not Rs.EOF
      CboGroupID.AddItem Trim(Rs!GroupID)
      Rs.MoveNext
      I = I + 1
Wend
If I = 0 Then
   MsgBox "No data"
   
End If
End Function

Private Function GetGroupWO(ByVal GroupID As String)
Dim str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset
ListWoNotFinish.Clear
CboClosed.Clear
str = "select Work_Order,ClosedFlag from QSMS_WOGroup  where GroupID= '" & GroupID & "' order by Seq_NO"

Set Rs = Conn.Execute(str)
cboWO.Clear
CboNotFinishedWO.Clear
CboNotChkBOM.Clear
While Not Rs.EOF
      If UCase(Trim(Rs!ClosedFlag)) = "Y" Then
          CboClosed.AddItem Trim(Rs!Work_Order)
      Else
          If ChkMBWo(Rs!Work_Order) = True Then
                If ChkQSMS_WO(Trim(Rs!Work_Order)) = False Then
                    CboNotChkBOM.AddItem Trim(Rs!Work_Order)
                Else
                
                    
                    If ChkWoFinished(Rs!Work_Order) = True Then
        
                        cboWO.AddItem Trim(Rs!Work_Order)
                    Else
                         ListWoNotFinish.AddItem Trim(Rs!Work_Order)
                         CboNotFinishedWO.AddItem Trim(Rs!Work_Order)
                    End If
                End If
          End If
      End If
      Rs.MoveNext
Wend
End Function

Private Function GetWoinfo(ByVal WO As String)
Dim str As String
Dim Rs As ADODB.Recordset
str = "select PN, Qty,[Group],BuildType,Line from Sap_Wo_List where WO='" & Trim(WO) & "'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   TxtMBPN = Rs!PN
   TxtModel = Mid(TxtMBPN, 3, 3)
   TxtWOQty = Rs!Qty
   TxtGroup = Trim(Rs![Group])
   TxtBuildType = Trim(Rs!BuildType)
   txtLine = Rs!Line
End If
str = "select TotalQty from QSMS_WoInputPlan where Work_Order='" & WO & "'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   TxtPlanQty = Rs!TotalQty
Else
   TxtPlanQty = ""
End If
str = "select Customer from ModelName where PN='" & TxtMBPN & "'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   TxtCustomer = Trim(Rs!Customer)
End If
End Function

Private Function GetLine()
Dim str As String
Dim Rs As ADODB.Recordset
str = "select distinct Line from QSMS_woGroup order by line"
Set Rs = Conn.Execute(str)
CboLine.Clear
While Not Rs.EOF
    CboLine.AddItem Rs!Line
    Rs.MoveNext
Wend
End Function

Private Function GetConsumeQty(WO, COMPPN) As Long
Dim str As String
Dim Rs As ADODB.Recordset
str = "select ConsumedQty from SMT_QSMS_Out where Work_Order='" & Trim(WO) & "' and CompPN='" & COMPPN & "'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   GetConsumeQty = Trim(Rs!ConsumedQty)
End If

End Function

Private Function ChkErr() As Boolean
Dim str As String
Dim Rs As ADODB.Recordset
ChkErr = True
If UCase(Trim(CboLine)) <> UCase(Trim(txtLine)) Then
   MsgBox "Line does not match with GroupID,please check the line "
   ChkErr = False
   
End If
If Trim(TxtWO) = "" Then
   MsgBox "The WO is error,Please check"
   ChkErr = False
End If
If Trim(CboGroupID) = "" Then
   MsgBox "The GroupID is error,Please check"
   ChkErr = False
End If
If Trim(TxtMachine) = "" Then
    MsgBox "Machine selected is error, Please check"
    ChkErr = False
End If
If Trim(txtDID) = "" Then
    MsgBox "DID selected  is error,please check"
    ChkErr = False
End If
If Trim(TxtDispatchQty) = "" Or Trim(TxtDispatchQty) = "0" Then
   MsgBox "Dispatch Qty is error,Please check"
   ChkErr = False
   Exit Function
End If
If ChkWOSeq(Trim(TxtWO), Trim(TxtMachine)) = False Then
   ChkErr = False
  
End If
If CLng(TxtDispatchQty) > CLng(TxtDIDRemainQty) Then
   MsgBox "dispatch Qty can not larger than DID remainQty"
   ChkErr = False
End If
If Trim(txtCompPN) = "" Then
    MsgBox "CompPN can not be empty,Please check"
    ChkErr = False
End If


If ChkDIDCompPN(Trim(txtDID), Trim(txtCompPN)) = False Then
    ChkErr = False
End If
If ChkNonAVL(Trim(txtDID), Trim(TxtCustomer), Trim(TxtModel), Trim(TxtMBPN), Trim(TxtWO)) = False Then
   ChkErr = False
   'MsgBox "The Comp doesn't been approved,Please check"
End If
If ChkAVL(Trim(txtCompPN), Trim(TxtVendorCode), Trim(TxtCustomer), Trim(TxtModel)) = False Then
   ChkErr = False
   MsgBox "Check AVL failed,please check "
End If
str = "Select UsedFlag from QSMS_DID where DID='" & Trim(txtDID) & "' and UsedFlag='Y'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
    ChkErr = False
    MsgBox "The DID has been used,please check"
End If
str = "select * from QSMS_WOGroup where GroupID='" & CboGroupID & "' and work_order='" & TxtWO & "'"
Set Rs = Conn.Execute(str)
If Rs.EOF Then
    ChkErr = False
    MsgBox "The Work Order: " & TxtWO & " does not belong to the GroupID :" & CboGroupID & " Please check"
ElseIf UCase(Trim(Rs!ClosedFlag)) = "Y" Then   '''(0009)
    ChkErr = False
    MsgBox "The Work Order: " & TxtWO & " has closed, Please check !!"
End If
If ChkDIDBelongToGroup(Trim(txtDID), Trim(CboGroupID)) = False Then
   ChkErr = False
End If

str = "selecT IPQCFlag from qsms_did where comppn in (selecT comppn from qsms_inspect_rule ) and did='" & txtDID & "'  "   '----(0007)
If Rs.State Then Rs.Close
Rs.Open str, Conn
    If Rs.EOF = False Then
       If Rs("IPQCFlag") = "N" Or Trim(Rs("IPQCFlag")) = "" Then
          ChkErr = False
          MsgBox "IPQC test fail or not test"
       End If
    End If

End Function

Private Function Insert_QSMS_Out() As Boolean
    Dim str As String
    Dim Rs As ADODB.Recordset
    Dim RsWoarray As ADODB.Recordset
    Dim RsInsert As ADODB.Recordset
    Dim FistTimeofDispatch As Integer ''To record if is the first dispatch --20070715
    Dim transdatetime As String
    Dim tempqty As Long, tempblaqty As Long
    Dim position As Long
    Dim COMPPN, SelectedCompPN As String
    Dim TempDisQty, TempBalQty, TempDIDQty As Long
    Dim DispQtybyWO, BalanceQtyByWO As Long
    Dim LoopNum As Long
    Dim Item, Slot, LR As String
    Dim tempwo As String
    Dim aryWO As Variant
    Dim WoArrayStr As String
    Dim BaseQty As Integer
    Dim j As Integer
    Dim I As Integer
    Insert_QSMS_Out = True
    WoArrayStr = ""
    LoopNum = 0
    str = "Select getdate()"
    Set Rs = Conn.Execute(str)
    transdatetime = Format(Rs(0), "YYYYMMDDHHMMSS")
    WoArrayStr = GetWoArrayForsp
    
    If StrBU <> "NB5" Then
        '''''''''''''''1176
        str = "select * from XL_ImplementPN with(nolock) where charindex(PrefixPN,'" & Trim(txtCompPN.Text) & "')=1"
        Set Rs = Conn.Execute(str)
        If Not Rs.EOF Then
            MsgBox "This material is dispatched using XL, and cannot be dispatched using the QSMS interface."
            Insert_QSMS_Out = False
            Exit Function
        End If
        
        '''''''''''''''1176
    End If
        
'(1) update QSMS_WO

''''''maybe need update by item, check with SAP_BOM ,if the Item is unigue for a work order
TempDisQty = CLng(TxtDispatchQty)
TxtConsumedQty = CLng(TxtConsumedQty) + CLng(TxtDispatchQty)
tempblaqty = TxtConsumedQty - CLng(TxtNeedQty)
TxtBalanceQty = tempblaqty
TxtDIDRemainQty = CLng(TxtDIDRemainQty) - CLng(TxtDispatchQty)

        '''(0012)
        str = "Exec QSMS_PrepairMaterial @WOGroup='" & Trim(TxtGroup.Text) & "',@WOString='" & Trim(WoArrayStr) & "',@CompPN='" & _
            Trim(txtCompPN.Text) & "',@JobGroup='" & Trim(CboJob.Text) & "',@Machine='" & Trim(TxtMachine.Text) & "',@VendorCode='" & _
            Trim(TxtVendorCode.Text) & "'"
        
        Set Rs = Conn.Execute(str)
        tempwo = ""
        FistTimeofDispatch = 1 '---20070715
        
        While Not Rs.EOF And TempDisQty > 0
            If ChkDIDDispatchedToWo(Rs!Work_Order, txtDID, Trim(TxtMachine), Trim(Rs!Slot), Trim(Rs!LR)) = False Then
               MsgBox "The DID has dispathcd to The WO, Please check" & Rs!Work_Order
            Else
        '        If tempwo = "" Or tempwo <> Trim(Rs!Work_Order) Then
                   Cboslot.Visible = False
                   LblSlot(4).Visible = False
                   Item = Trim(Rs!Item)
                   Slot = Trim(Rs!Slot)
                   LR = Trim(Rs!LR)
                   BaseQty = CLng(Rs!BaseQty)
        'Add by  leimo 20061226
                   If TempDisQty + Rs!BalanceQty > 0 Then
                       TempDIDQty = -Rs!BalanceQty
                   Else
                       TempDIDQty = TempDisQty
                   End If
                   DispQtybyWO = Rs!DispatchQty + TempDIDQty
                   BalanceQtyByWO = DispQtybyWO - Rs!NeedQty
                   
        
        ''add by leimo 20061226
                    TempDisQty = TempDisQty - TempDIDQty
        
        '           Conn.Execute Str
        'Update by Leimo 20070118--use SP to insert into QSMS_Dispatch and update QSMS_WO
                    str = "exec  QSMSInsertDispatch  '" & Rs!Work_Order & "','" & Trim(CboGroupID) & "' ,'" & Trim(CboLine) & "' ,'" & Trim(Rs!WOqty) & "','" & Trim(Rs!Jobpn) & "' ,'" & TxtMachine & "' " & _
                          " ,'" & txtCompPN & "' ,'" & Slot & "','" & LR & "','" & BaseQty & "'," & Trim(Rs!TotalNeedQty) & " ,'" & Trim(txtDID) & "'," & CLng(TxtDIDTotalQty) & " ," & TempDIDQty & " " & _
                          ",'" & TxtVendorCode & "' ,'" & TxtDateCode & "','" & TxtLotCode & "','" & g_userName & "' ,'" & transdatetime & "','" & Trim(TxtDIDDateTime) & "','','" & Item & "','" & Trim(Rs!jobgroup) & "','" & Trim(Rs!Side) & "'"
        
                   Set RsInsert = Conn.Execute(str)
                   If RsInsert.EOF Then
                      MsgBox "Insert into QSMS_Dispatch Error,please retry again"
                      Insert_QSMS_Out = False
                      Exit Function
                   Else
                       If UCase(Trim(RsInsert.Fields(0))) = "PASS" Then
                            'record the first dispatch time ---20070715
                            If FistTimeofDispatch = 1 Then
                                str = "exec RecordDispatchFDT '" & Rs!Work_Order & "'"
                                Conn.Execute (str)
                                FistTimeofDispatch = FistTimeofDispatch + 1
                            End If
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
                                            MsgBox "Insert into QSMS_Dispatch Error,please retry again "
                                            Exit Function
                                     End If
                                End If
                            
                       End If
        
                   End If
                
                   tempwo = Trim(Rs!Work_Order)
                   LoopNum = LoopNum + 1
            End If
            
           Rs.MoveNext
        Wend
        If LoopNum = 0 Then
           MsgBox "didn't dispatch the  material ,please check"
           Insert_QSMS_Out = False
           Exit Function
        End If
        
        '(4)  check if need refresh DID combo
        If tempblaqty >= 0 Or CLng(TxtDIDRemainQty) = 0 Then
           Call RefreshDID_Machine_WO("DID", txtCompPN, TxtMachine, TxtWO, Trim(CboJob), CboLine, WoArrayStr)
        End If
        
        Call ChkWOItemFinished(WoArrayStr)
'    Next j
    txtCompPN.Text = ""
    txtDID.Text = ""
    txtDID.SetFocus
End Function
Public Function RefreshDID_Machine_WO(ByVal RefreshType As String, ByVal COMPPN As String, ByVal Machine As String, ByVal WO As String, ByVal MBPN As String, ByVal Line As String, ByVal wostr As String)
Dim str As String
Dim Rs As ADODB.Recordset

Select Case UCase(RefreshType)
       Case "DID"
             Call GetDID(Machine, MBPN, WO, Line)
       Case "MACHINE"
             Call GetMachine(WO, wostr)
       Case "WO"
             Call GetGroupWO(CboGroupID)
End Select

End Function


Private Function GetDID(ByVal Machine As String, ByVal Jobpn As String, ByVal WO As String, ByVal Line As String)
Dim str As String
Dim Rs As ADODB.Recordset
Dim TempRs As ADODB.Recordset
Dim wostr As String
wostr = GetWoArray



str = "select a.CompPN from QSMS_NonAVL a,QSMS_Wo b where (B.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') or b.work_order in " & wostr & ") " & _
     "and a.CompPN=b.CompPN and a.Customer='" & Trim(TxtCustomer) & "' and Model='" & Trim(TxtModel) & "'"
Set Rs = Conn.Execute(str)
Set DGAVL.DataSource = Rs
DGAVL.Refresh
'---------(0005)
str = "Select  a.CompPn,a.Slot,a.LR,-sum(a.DispatchQty-a.PlanNeedQty) as NeedQty from QSMS_WO a where (a.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') or a.work_order in " & wostr & ") " & _
      "and a.Machine='" & Trim(Machine) & "' and a.DispatchQty-a.PlanNeedQty<0 and JobGroup like '" & Trim(Jobpn) & "%' group by a.comppN,a.Slot,a.LR"
Set Rs = Conn.Execute(str)
DGCompNotOK.Caption = "Comp didn't dispatch:" & Rs.RecordCount
Set DGCompNotOK.DataSource = Rs
DGCompNotOK.Columns(1).Width = 550
DGCompNotOK.Columns(2).Width = 550
If Rs.EOF Then
   Call CmdRefresh_Click
End If
DGCompNotOK.Refresh

End Function

Private Function ChkBalanceQty(ByVal Machine As String, ByVal WO As String, ByVal COMPPN As String) As Boolean
Dim str As String
Dim Rs As ADODB.Recordset
ChkBalanceQty = True
str = "select Work_Order From QSMS_WO where Work_Order='" & WO & "' and Machine='" & Machine & "' and CompPN='" & COMPPN & "' and BalanceQty<0"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   ChkBalanceQty = False
End If
End Function
Private Function GetDIDInfo(ByVal DID As String, ByVal WO As String)
Dim str As String
Dim Rs As ADODB.Recordset
Dim Used_Flag As String
TxtConsumedQty = ""
TxtCompBaseQty = ""
TxtNeedQty = ""
TxtDispatchQty = ""
txtCompPN = ""
TxtVendorCode = ""
TxtDIDDateTime = ""

TxtDateCode = ""
TxtLotCode = ""
TxtDIDRemainQty = ""

str = "select DID,CompPN,VendorCode,DateCode,LotCode,Qty,RemainQty,UsedFlag,DIDLoc,TransDateTime from QSMS_DID where DID='" & Trim(txtDID) & "'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   txtCompPN = Trim(Rs!COMPPN)
   TxtVendorCode = Trim(Rs!VendorCode)
   TxtDateCode = Trim(Rs!DateCode)
   TxtLotCode = Trim(Rs!LotCode)
   TxtRackID = Trim(Rs!DIDLoc)
   TxtDIDTotalQty = Trim(Rs!Qty)
   TxtDIDRemainQty = Trim(Rs!RemainQty)
   TxtDIDDateTime = Trim(Rs!transdatetime)

    If UCase(Trim(Rs!usedflag)) = "Y" Then

       TxtDispatchQty.Enabled = False
       TxtDispatchQty = 0
   Else

       TxtDispatchQty.Enabled = True
       
   End If
   
End If
End Function
Private Function GetCompDispInfo(ByVal WO As String, ByVal Jobpn As String, ByVal Machine As String, ByVal COMPPN As String)
Dim str As String
Dim Rs As ADODB.Recordset
Dim WoArray As String
Dim TempBalanceQty As String
Dim TempConsumeQty As String

TxtConsumedQty = ""
TxtCompBaseQty = ""
TxtNeedQty = ""
WoArray = GetWoArray
str = "Select sum(BaseQty) as BaseQty, sum(NeedQty) as TotalNeedQty ,sum(PlanNeedQty) as NeedQty,sum(DispatchQty) as DispatchQty, sum(case when DispatchQty-PlanNeedQty<0 then DispatchQty-PlanNeedQty else 0 end) as BalanceQty " & _
    "from QSMS_WO where (Work_Order " & _
       " in (select WO from sap_wo_list where [group]='" & Trim(TxtGroup) & "' )or work_order in " & WoArray & ") and Jobgroup like '" & Jobpn & "%' and machine='" & Machine & "' and comppn='" & COMPPN & "' "
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
    TxtCompBaseQty = Trim(Rs!BaseQty)
    TxtConsumedQty = Trim(Rs!DispatchQty)
    TxtBalanceQty = Trim(Rs!BalanceQty)
    
    TxtNeedQty = Trim(Rs!NeedQty)
    TxtTotalQty = Trim(Rs!TotalNeedQty)
End If
If Trim(TxtNeedQty) = "" Then
  Exit Function
End If

If CLng(Trim(TxtDIDRemainQty)) > -CLng(Trim(TxtBalanceQty)) Then
   TxtDispatchQty = -CLng(TxtBalanceQty)
Else
   TxtDispatchQty = TxtDIDRemainQty
End If

 Call GetDispatchQtyBySlot(Trim(WO), Trim(Machine), Trim(COMPPN))

End Function
Private Function GetSlot(ByVal DID As String, ByVal WO As String)
Dim str As String
Dim Rs As ADODB.Recordset
Dim Used_Flag As String
Dim Slot As String
Dim I As Long

LblSlot(4).BackColor = &HFF&


Cboslot.Clear
Slot = ""
I = 0
str = "select b.Slot from QSMS_DID a,QSMS_WO b where a.DID='" & Trim(DID) & "' and a.CompPN=b.CompPN and (b.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') or b.work_order in " & GetWoArray & ") " & _
      "and b.Machine='" & Trim(TxtMachine) & "'"
Set Rs = Conn.Execute(str)
While Not Rs.EOF
      Slot = Trim(Rs!Slot)
      Cboslot.AddItem Trim(Rs!Slot)
      Rs.MoveNext
Wend

End Function

Private Function GetMachine(ByVal WO As String, ByVal wostr As String)
Dim str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset
Dim rsMachine As ADODB.Recordset
Dim tempmachine As String
Dim Rs1 As ADODB.Recordset

 'and machine like '" & Trim(CboLine) & "%'
str = "select Wo from Sap_Wo_List where [Group]='" & TxtGroup & "' "
Set Rs1 = Conn.Execute(str)
wostr = Left(wostr, Len(wostr) - 1)
While Not Rs1.EOF
    wostr = wostr & ",'" & Rs1("wo") & "'"
    Rs1.MoveNext
Wend
wostr = wostr & ")"

str = "select distinct Machine,MachinefinishedFlag from QSMS_WO where work_order in " & wostr & _
      "AND DispatchQty-PlanNeedQty<0 order by machine,MachinefinishedFlag"
            
Set Rs = Conn.Execute(str)
CboMathineOK.Clear
CboMathineNOK.Clear

While Not Rs.EOF
     If tempmachine = "" Or tempmachine <> UCase(Trim(Rs!Machine)) Then
        If UCase(Trim(Rs!MachinefinishedFlag)) = "N" Then
            CboMathineNOK.AddItem Trim(Rs!Machine)
        Else
            CboMathineOK.AddItem Trim(Rs!Machine)
        End If
     End If
     tempmachine = UCase(Trim(Rs!Machine))
        
     Rs.MoveNext
Wend

End Function

Private Function ChkWOSeq(ByVal WO As String, ByVal Machine As String) As Boolean
Dim str As String
Dim Rs As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim TempGroupID As String
Dim Seq_No As Long
ChkWOSeq = True
str = "select GroupID,Seq_No from  QSMS_wogroup where Work_Order='" & Trim(WO) & "'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   TempGroupID = Trim(Rs!GroupID)
   Seq_No = Rs!Seq_No
Else
   ChkWOSeq = False
   MsgBox "The Work order has No group ID, please call PMC "
End If

End Function

Public Function ChkDIDBelongMachine(ByVal WO As String, ByVal Machine As String, ByVal DID As String, SAPWOGroup As String) As Boolean
Dim str As String
Dim Rs As ADODB.Recordset
Dim MessageString As String
Dim strSQL As String

str = "select UsedFlag from QSMS_DID where DID='" & DID & "'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then

   If UCase(Trim(Rs!usedflag)) = "Y" Then
      MessageString = ""
      str = "select a.work_order,a.Machine,a.Slot From qsms_dispatch a join qsms_did b on a.did=b.did and a.diddatetime=b.transdatetime where a.did='" & DID & "'"
      Set Rs = Conn.Execute(str)
      Do While Not Rs.EOF
            MessageString = MessageString + "WO:" + Rs!Work_Order + "  Machine:" + Rs!Machine + "  Slot:" + Rs!Slot + vbCrLf
            Rs.MoveNext
        Loop
      MsgBox "The DID has been used at: " + vbCrLf + vbCrLf + MessageString + vbCrLf + " ===PLease check=== !"
      ChkDIDBelongMachine = False
      Exit Function
   Else        '''''(0003)Add by Udall 2007.11.05
      MessageString = ""
      str = "select a.work_order,a.Machine,a.Slot From qsms_dispatch a join qsms_did b on a.did=b.did and a.diddatetime=b.transdatetime where a.did='" & DID & "'"
      Set Rs = Conn.Execute(str)
      Do While Not Rs.EOF
            If Left(Trim(Rs!Machine), 2) <> Left(Machine, 2) Then
               MessageString = MessageString + "WO:" + Rs!Work_Order + "  Machine:" + Rs!Machine + "  Slot:" + Rs!Slot + vbCrLf
            End If
            Rs.MoveNext
        Loop
      If MessageString <> "" Then
          MsgBox "DID can not be dispatched to different line and side,the DID has been dispatched to: " + vbCrLf + vbCrLf + MessageString + vbCrLf + " ===PLease check=== !"
          ChkDIDBelongMachine = False
          Exit Function
      End If      ''''(0003)
   End If
Else
'''''''''''''''''''''''''''''''''''''add by Jing 2007.10.31------------(0002)'''''''''''''''''''''''''''''''''''''''''''''''
    strSQL = "select * from qsms_did_log where did='" & DID & "'"
    Set Rs = Conn.Execute(strSQL)
    If Rs.RecordCount > 0 Then
        MsgBox "This DID had been deleted ! "
    Else
       MsgBox "DID does not exist"
    End If
    ChkDIDBelongMachine = False
    Exit Function
End If

str = "select b.DID,B.Qty from QSMS_WO a, QSMS_DID b where (a.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(SAPWOGroup) & "') or a.work_order in " & GetWoArray & ") " & _
       "and a.Machine='" & Trim(Machine) & "' and a.CompPN=b.CompPN  and b.DID='" & DID & "' and b.UsedFlag='N'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   ChkDIDBelongMachine = True
Else
   ChkDIDBelongMachine = False
   MsgBox "the DID does not belong to the machine :" & Machine & ".  Please Check."
End If

End Function

Private Sub TxtChkDID_KeyPress(KeyAscii As Integer)
Dim str As String
Dim Rs As ADODB.Recordset
If KeyAscii = 13 Or KeyAscii = 9 Then
   ListDIDStatus.Clear
   'str = "select b.DID,B.Qty ,B.UsedFlag from QSMS_WO a, QSMS_DID b where a.Work_Order='" & Trim(TxtWO) & "'  and a.Machine='" & Trim(TxtMachine) & "' and a.CompPN=b.CompPN  and b.DID='" & TxtChkDID & "' "
   str = "select b.DID,B.Qty ,B.UsedFlag from QSMS_WO a, QSMS_DID b where (a.Work_Order in (select Wo from Sap_Wo_List where [Group]='" & Trim(TxtGroup) & "') or a.work_order in " & GetWoArray & ") " & _
         "and a.Machine='" & Trim(TxtMachine) & "' and a.CompPN=b.CompPN  and b.DID='" & TxtChkDID & "' "
   
   Set Rs = Conn.Execute(str)
   If Not Rs.EOF Then
      If UCase(Rs!usedflag) = "Y" Then
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
Dim StrPN As String
If KeyAscii = 13 Or KeyAscii = 9 Then
    If Trim(txtDID) = "" Then Exit Sub
    txtDID = Trim(txtDID)
   If strGetDIDFromSourceBU = "Y" Then '(0011)
      Call GetDIDFromSourceBU(txtDID)
   End If
   If ChkDIDBelongMachine(Trim(TxtWO), Trim(TxtMachine), Trim(txtDID) & "", Trim(TxtGroup)) = False Then
     txtDID.Text = ""
     txtDID.SetFocus
     Call Warning_Sound
     Exit Sub
   End If
   
   Call GetDIDInfo(Trim(txtDID), Trim(TxtWO))
   
   ''0014
 '------------add check IPQCFlag='Y' before dispatch----------(0013)
   'If StrBU = "MBU" And ChkIPQC(Trim(TxtDID)) = False Then
        'Call Warning_Sound
        'MsgBox "DID:" & Trim(TxtDID) & " Check IPQC failed,please check!!! "
        'Exit Sub
   'End If
'------------end (0013)-------------------------------------------
   
'-------------add check DID and CompPN if matched --------(0001)
     If Check_DID = "Y" Then
       StrPN = InputBox("Please Input CompPN", "Input CompPN")
       If UCase(Trim(StrPN)) <> UCase(Trim(txtCompPN)) Then
            Call Warning_Sound
            MsgBox ("DID and CompPN aren't matched,Please check DID and CompPN")
            txtDID.Text = ""
            txtDID.SetFocus
            Exit Sub
        End If
    End If

'--------------the end of add check DID and CompPN if matched ------------
   If ChkAVL(Trim(txtCompPN), Trim(TxtVendorCode), Trim(TxtCustomer), Trim(TxtModel)) = False Then
        Call Warning_Sound
        MsgBox "Check AVL failed,please check "
        Exit Sub
   End If
   If ChkNonAVL(Trim(txtDID), Trim(TxtCustomer), Trim(TxtModel), Trim(TxtMBPN), Trim(TxtWO)) = False Then
     
      'MsgBox "The Comp doesn't been approved,Please check"
   End If
   Call GetWoArray
   Call GetCompDispInfo(Trim(TxtWO), Trim(CboJob), Trim(TxtMachine), Trim(txtCompPN))
   TxtDispatchQty.SetFocus
  ' Call GetSlot(Trim(TxtDID), Trim(TxtWO))
End If
End Sub
Private Function ChkIPQC(ByVal DID As String) As Boolean '(0013)
Dim str As String
Dim Rs As ADODB.Recordset

str = "select IPQCFlag from QSMS_DID where DID='" & DID & "' and LEFT(CompPN,2) IN ('TH','CH','TS','CS','TU')"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
    If UCase(Trim(Rs!IPQCFlag)) = "Y" Then
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

Private Function GetDIDFromSourceBU(DID As String) As Boolean
Dim Rs As New ADODB.Recordset
Dim strSQL As String
On Error GoTo ErrHdl:
strSQL = "EXEC QSMS_GetDIDFromSourceBU @DID='" & DID & "'"
Conn.Execute (strSQL)
Exit Function
ErrHdl:
    MsgBox Err.Description
End Function

Private Function ChkDIDInherit(ByVal WO As String, ByVal DID As String) As Boolean
Dim str As String
Dim Rs As ADODB.Recordset
Dim GroupID As String
ChkDIDInherit = False
If TxtDIDTotalQty = TxtDIDRemainQty Then
   ChkDIDInherit = True
   Exit Function
End If
str = "select GroupID from QSMS_WOGroup where Work_Order='" & WO & "'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   GroupID = Trim(Rs!GroupID)
   str = "select DID from QSMS_Dispatch a, QSMS_WOGroup b where a.work_order=b.Work_Order and B.GroupID='" & GroupID & "' and a.DID='" & DID & "' "
   Set Rs = Conn.Execute(str)
   If Rs.EOF Then
      ChkDIDInherit = False
   Else
      ChkDIDInherit = True
   End If
End If


End Function

Private Function ChkDIDDispatchedToWo(ByVal WO As String, ByVal DID As String, ByVal Machine As String, ByVal Slot As String, LR As String) As Boolean
Dim str As String
Dim Rs As ADODB.Recordset
Dim DeleteFlag As Boolean

DeleteFlag = False
ChkDIDDispatchedToWo = False
   str = "select DID,DeletedFlag,Machine,Slot,LR from QSMS_Dispatch  where work_order='" & WO & "'  and DID='" & DID & "' "
   Set Rs = Conn.Execute(str)
   If Rs.EOF Then
      ChkDIDDispatchedToWo = True
      Exit Function
   End If
   While Not Rs.EOF
      If UCase(Trim(Rs!DeletedFlag = "Y")) Or (UCase(Machine) = UCase(Trim(Rs!Machine)) And UCase(Slot) = UCase(Trim(Rs!Slot)) And UCase(LR) = UCase(Trim(Rs!LR))) Then
         ChkDIDDispatchedToWo = True
      Else
         ChkDIDDispatchedToWo = False
         Exit Function
      End If
      Rs.MoveNext
  Wend
   
End Function
Private Function ChkDIDCompPN(ByVal DID As String, ByVal COMPPN As String) As Boolean
Dim str As String
Dim Rs As ADODB.Recordset
ChkDIDCompPN = True
str = "select DID from QSMS_DID where DID='" & DID & "' and CompPN='" & COMPPN & "'"
Set Rs = Conn.Execute(str)
If Rs.EOF Then
   MsgBox "The DID and CompPN doesn't match"
   ChkDIDCompPN = False
End If
End Function

Private Function GetWoArrayForsp() As String
Dim WoArray As String
Dim str As String
Dim Rs As ADODB.Recordset
Dim I As Long

    str = "select wo from Sap_WO_List where [Group] in (select [group] from sap_wo_list where wo='" & Trim(TxtWO) & "')"
        Set Rs = Conn.Execute(str)
        While Not Rs.EOF
               WoArray = WoArray + Trim(Rs!WO) + ","
               Rs.MoveNext
        Wend
    
    For I = 1 To ListWoDispatching.ListCount
        ListWoDispatching.ListIndex = I - 1
        str = "select wo from Sap_WO_List where [Group] in (select [group] from sap_wo_list where wo='" & ListWoDispatching.Text & "')"
        Set Rs = Conn.Execute(str)
        While Not Rs.EOF
               WoArray = WoArray + Trim(Rs!WO) + ","
               Rs.MoveNext
        Wend
        'WoArray = WoArray + "'" + ListWoDispatching.Text + "'" + ","
        
    Next I
    WoArray = Mid(WoArray, 1, Len(WoArray) - 1)
    'WoArray = "(" + WoArray + ")"
    GetWoArrayForsp = WoArray
End Function

Private Function GetWoArray() As String
Dim WoArray As String
Dim str As String
Dim Rs As ADODB.Recordset
Dim I As Long

    str = "select wo from Sap_WO_List where [Group] in (select [group] from sap_wo_list where wo='" & Trim(TxtWO) & "')"
        Set Rs = Conn.Execute(str)
        While Not Rs.EOF
               WoArray = WoArray + "'" + Trim(Rs!WO) + "'" + ","
               Rs.MoveNext
        Wend
    
    For I = 1 To ListWoDispatching.ListCount
        ListWoDispatching.ListIndex = I - 1
        str = "select wo from Sap_WO_List where [Group] in (select [group] from sap_wo_list where wo='" & ListWoDispatching.Text & "')"
        Set Rs = Conn.Execute(str)
        While Not Rs.EOF
               WoArray = WoArray + "'" + Trim(Rs!WO) + "'" + ","
               Rs.MoveNext
        Wend
        
    Next I
    WoArray = Mid(WoArray, 1, Len(WoArray) - 1)
    WoArray = "(" + WoArray + ")"
    GetWoArray = WoArray
End Function

Private Function GetSBWO(ByVal WO As String)
Dim str As String
Dim Rs As ADODB.Recordset
Dim I As Long
Dim Group As String
I = 0
CboSBWO.Clear
FraSB.Visible = False
str = "Select [Group] from Sap_Wo_List where wo='" & WO & "'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   Group = Trim(Rs!Group)
   TxtGroup = Group
End If
str = "select Wo from Sap_Wo_list where [Group] ='" & Group & "' and wo<>'" & WO & "' order by wo"
Set Rs = Conn.Execute(str)
While Not Rs.EOF
     CboSBWO.AddItem Trim(Rs!WO)
     Rs.MoveNext
     I = I + 1
Wend
If I > 0 Then
    FraSB.Visible = True

End If
End Function

Private Function GetDispatchQtyBySlot(ByVal WO As String, ByVal Machine As String, ByVal COMPPN As String)
Dim str As String
Dim Rs As ADODB.Recordset
Dim WoArray As String
WoArray = GetWoArray
str = "Select work_order,machine,slot,NeedQty,DispatchQty,BalanceQty,PlanQty,PlanNeedQty,DispatchQty-PlanNeedQty as PlanBalanceQty " & _
    "from QSMS_WO where (Work_Order " & _
       " in (select WO from sap_wo_list where [group]='" & Trim(TxtGroup) & "' )or work_order in " & WoArray & ") and machine='" & Machine & "' and comppn='" & COMPPN & "' "
Set Rs = Conn.Execute(str)
Set DGSlot.DataSource = Rs
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
    Call EffectSound("OO.wav") '1092
'      wave_control.FileName = App.Path & "\OO.wav"
'      wave_control.Command = "open"
'      wave_control.Command = "play"
'      Do While wave_control.Mode = mciModePlay
'      Loop
'      wave_control.Command = "close"
End Sub
Private Sub OK_Sound()
    Call EffectSound("OK.wav") '1092
'    wave_control.FileName = App.Path & "\OK.wav"
'    wave_control.Command = "open"
'    wave_control.Command = "play"
'    Do While wave_control.Mode = mciModePlay
'    Loop
'    wave_control.Command = "close"
End Sub

Private Function GetJobForBuildType(ByVal WO As String, ByVal Machine As String, ByVal BuildType As String)
Dim str As String
Dim Rs As ADODB.Recordset
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
                Set Rs = Conn.Execute(str)
                If Rs.RecordCount > 1 Then
                   While Not Rs.EOF
                         CboJob.AddItem Trim(Rs!jobgroup)
                         Rs.MoveNext
                   Wend
               End If
End Select
End Function
