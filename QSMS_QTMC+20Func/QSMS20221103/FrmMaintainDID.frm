VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FrmMaintainDID 
   BackColor       =   &H80000009&
   Caption         =   "Maintain DID  2016/01/19"
   ClientHeight    =   9420
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmaVendorDesc 
      BackColor       =   &H80000013&
      Caption         =   "Vendor description"
      Height          =   1815
      Left            =   7920
      TabIndex        =   35
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Comp Port data maintain"
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   7695
      Begin VB.OptionButton optNetwork 
         Caption         =   "Network"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   1200
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Printer"
         Height          =   615
         Left            =   1440
         TabIndex        =   37
         Top             =   960
         Width           =   4695
         Begin VB.OptionButton OptZebra 
            Caption         =   "Zebra printer"
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
            Left            =   720
            TabIndex        =   39
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton OptSATO 
            Caption         =   "SATO printer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   2760
            TabIndex        =   38
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.OptionButton OptPrint 
         Caption         =   "Print Port"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton OptComp 
         Caption         =   "Comp Port"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton CmdCommSave 
         Caption         =   "CommSave"
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
         Left            =   6240
         Picture         =   "FrmMaintainDID.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtComm 
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
         Left            =   4560
         TabIndex        =   18
         Text            =   "9600,N,8,1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox TxtCompPort 
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
         Left            =   2880
         TabIndex        =   16
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "Settings"
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
         Left            =   3480
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
         Caption         =   "CompPort"
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
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   14160
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSDataGridLib.DataGrid DG1 
      Height          =   3855
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   6800
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
   Begin VB.Frame FraConnection 
      BackColor       =   &H80000013&
      Caption         =   "DID maintain "
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   14415
      Begin VB.TextBox txtInspection 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   43
         Top             =   1680
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.TextBox txtMSD 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   42
         Top             =   1680
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CommandButton cmdForceDel 
         Caption         =   "Command1"
         Height          =   255
         Left            =   13560
         TabIndex        =   40
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton CmdVendorCode 
         Caption         =   "Vendor code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6840
         Picture         =   "FrmMaintainDID.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton CmdReprint 
         Caption         =   "Reprint"
         DragIcon        =   "FrmMaintainDID.frx":0614
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8160
         Picture         =   "FrmMaintainDID.frx":6226
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2400
         Picture         =   "FrmMaintainDID.frx":BE38
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3120
         Picture         =   "FrmMaintainDID.frx":C27A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3840
         Picture         =   "FrmMaintainDID.frx":C584
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   960
         Picture         =   "FrmMaintainDID.frx":C9C6
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
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
         Height          =   855
         Left            =   240
         Picture         =   "FrmMaintainDID.frx":CE08
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1680
         Picture         =   "FrmMaintainDID.frx":D24A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "&Excel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6000
         Picture         =   "FrmMaintainDID.frx":D68C
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "EXIT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5280
         Picture         =   "FrmMaintainDID.frx":D996
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2280
         Width           =   735
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4560
         Picture         =   "FrmMaintainDID.frx":DCA0
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox TxtGroupQty 
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
         Left            =   1800
         TabIndex        =   21
         Text            =   "1"
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox CboDID 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8880
         TabIndex        =   13
         Top             =   1200
         Width           =   4575
      End
      Begin VB.ComboBox CboCompPN 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1800
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox CboVendorCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   6240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox CboDateCode 
         BackColor       =   &H00FFFFFF&
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
         Left            =   10440
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox CboLotCode 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1800
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox TxtQty 
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
         Left            =   5280
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label labInpectionNo 
         BackColor       =   &H0000FF00&
         Caption         =   "InspectionNo."
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
         Left            =   4560
         TabIndex        =   45
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label labelMSD 
         BackColor       =   &H0000FF00&
         Caption         =   "MSD"
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
         Left            =   240
         TabIndex        =   44
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
         Caption         =   "Group Qty"
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
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
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
         Index           =   1
         Left            =   7320
         TabIndex        =   11
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H0000FF00&
         Caption         =   "Date Code"
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
         Left            =   8880
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080FF80&
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
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H0000FF00&
         Caption         =   "Vendor Code"
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
         Left            =   4560
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Lot Code"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "Qty"
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
         Left            =   4560
         TabIndex        =   6
         Top             =   1200
         Width           =   735
      End
   End
   Begin VB.Menu mnubasic 
      Caption         =   "BasicData"
      Begin VB.Menu mnuCompPort 
         Caption         =   "CompPort"
      End
   End
End
Attribute VB_Name = "FrmMaintainDID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/**********************************************************************************
'**文 件 名: FrmMaintainDID.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: Jeanson
'**日    期: 2007.10.01
'**描    述: QSMS Maintain DID
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**             Jeanson      2007.10.15     update the condition of checking DID is using or not (0001)
'**             Jing         2007.10.24     check the printer is Zebra or Sato for print barcode (0002)
'**             Sandy        2007.11.05     if SATO printer will not delay time                  (0003)
'**             Jeanson      2007.11.07     update realqty/lastupdatetime instead of transdatetime when updating (0004)
'**             Sandy        2007.11.16     update the DID deleted function. (delete DID both in QSMS and in FUJI DB)(0005)
'**             Jing         2007.11.26     control the delete and reprint (0006)
'**             Sandy        2007.11.26     update to drive out the blank space in DID.---------(0007)
'**             Sandy        2007.11.30     added that don't allow to delete dispatched DID in MaintainDID----(0008)
'**             Sandy        2007.12.13     set default Printer according to each BU whether have FUJITrax ----(0009)
'**             Sandy        2007.12.19     checking DID dispatched or DIDSlotlink TO combine the delete DiD ----(0010)
'**             Jing         2008.01.10     Add error alarm when can not find label file     (0011)
'**             Jing         2008.01.10     Changed from 'NB25' to 'NB5' (0012)
'**             Salon        2008.01.25     not set the datecode/lotcode/VendorCode value automatically (0013)
'**             Sandy        2008.03.26      check ' in the datecode/lotcode/VendorCode value---(0014)
'**             Jeanson      2008.03.26     to check whether the vendor length is less than 8  (0015)
'**             Jeanson      2008.09.09     not to display the DID which Qty=0  (0016)
'**             Udall        2009.11.30     针对AP代工Foxconn的机种，其刷入的barcode中包含料号，vendorCode,DateCode,LotCode，以"|"分割，特殊处理 (0017)
'**RQ09122849   Archer       2010/03/06     Modify program to use new label format as default option(0018)
'**             Austin       2010/05/15     Add Para @MSD for GenRegisterDID (0019)
'***********************************************************************************1/
Dim Rs2 As ADODB.Recordset
Dim CommandType As Long
Dim isZebra As Boolean
Dim strSql  As String

Private Sub CboCompPN_Click()
Dim str As String, I As Integer
Dim rs As ADODB.Recordset
If Check_AVL = "Y" Then
   str = "Select distinct VendorCode from QSMS_AVL where CompPN='" & Trim(CboCompPN) & "'"
   Set rs = Conn.Execute(str)
    I = 0
    CboVendorCode.Clear
    While Not rs.EOF
      CboVendorCode.AddItem Trim(rs!VendorCode)
      rs.MoveNext
      I = I + 1
    Wend
End If

CboVendorCode.SetFocus

End Sub

Private Sub CboCompPN_KeyPress(KeyAscii As Integer)
Dim tempComp() As String, I As Integer
Dim NewComp() As String, index As Integer
If (KeyAscii = 13 Or KeyAscii = 9) And CboCompPN <> "" Then
'******************************
'****add by jeanson 2007/09/03
    CboCompPN.Text = Replace(Replace(Replace(CboCompPN.Text, " ", ""), vbCr, ""), vbLf, "")
    strErrMessage = ""

''''(0017) Start''''''''''
    tempComp = Split(Trim(CboCompPN), "|")
    
    For I = 0 To UBound(tempComp)
        If I = 0 Then
            CboCompPN = Trim(tempComp(I))
        End If
        If I = 1 Then
            CboVendorCode = Trim(tempComp(I))
        End If
        If I = 2 Then
            CboDateCode = Trim(tempComp(I))
        End If
        If I = 3 Then
            CboLotCode = Trim(tempComp(I))
        End If
        If I = 4 Then
            TxtQty = Trim(tempComp(I))
        End If
    Next I
''''(0017) End''''''''''

''(1018) Start ---------------------''
    NewComp = Split(Trim(CboCompPN.Text), ";")
    For index = 0 To UBound(NewComp)
        If index = 0 Then
            CboCompPN.Text = Trim(NewComp(index))
        ElseIf index = 1 Then
            CboDateCode.Text = Trim(NewComp(index))
        ElseIf index = 2 Then
            CboVendorCode.Text = Trim(NewComp(index))
        ElseIf index = 3 Then
            CboLotCode.Text = Trim(NewComp(index))
        ElseIf index = 4 Then
            TxtQty.Text = Trim(NewComp(index))
        End If
            
    Next index
''(1018) End -----------------------''

    strErrMessage = FunPartNumberCheck(CboCompPN.Text)
    If strErrMessage <> "PASS" Then
        MsgBox strErrMessage
        CboCompPN.SetFocus
        Exit Sub
    End If
    '******************************
    If UBound(tempComp) = 0 And UBound(NewComp) = 0 Then ''(1018)
       CboCompPN_Click
    Else
'        Call TxtQty_KeyPress(13)
        TxtQty.SetFocus
        Call TxtQty_Click
    End If
End If
End Sub

Private Sub CboDateCode_Click()
CboLotCode.SetFocus
End Sub

Private Sub CboDateCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboDateCode_Click
End If
End Sub

Private Sub CboDID_Click()
Call cmdFind_Click
End Sub

Private Sub CboLotCode_Click()
TxtQty.SetFocus
End Sub

Private Sub CboLotCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboLotCode_Click
End If
End Sub

Private Sub CboVendorCode_Click()
Dim str As String
Dim rs As ADODB.Recordset
If Check_AVL = "Y" Then
    str = "Select Desc1 from QSMS_AVL where CompPN='" & Trim(CboCompPN) & "' and VendorCode='" & Trim(CboVendorCode) & "'"
    Set rs = Conn.Execute(str)
    txtDesc.Text = Trim(rs!Desc1)
End If
CboDateCode.SetFocus
'If ChkAVL(Trim(CboCompPN), Trim(CboVendorCode)) = True Then
'    CboDateCode.SetFocus
'Else
'    CboVendorCode.Text = ""
'    Exit Sub
'End If
End Sub

Private Sub CboVendorCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboVendorCode_Click
End If
End Sub

Private Sub cmdADD_Click()
    If IsInteger(TxtQty) = False Then
        MsgBox ("Check the qty of print please !"), vbCritical
        TxtQty.SetFocus
        Exit Sub
    End If

    cmdAdd.Enabled = False
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdExit.Enabled = True
    cmdFind.Enabled = True
    CboCompPN.Enabled = True
    CboVendorCode.Enabled = True
    CboDateCode.Enabled = True
    CboLotCode.Enabled = True
    TxtQty.Enabled = True
    CommandType = 1
    cmdSave.SetFocus
End Sub

Private Sub cmdCancel_Click()
CboCompPN.Text = ""
CboVendorCode.Text = ""
CboDateCode.Text = ""
CboLotCode.Text = ""
TxtQty.Text = ""
CboDID.Text = ""
'CboCompPN.SetFocus
TxtGroupQty.SetFocus
End Sub

Private Sub CmdCommSave_Click()
SaveSetting "SMT", "QSMS", "CommPort", TxtCompPort
SaveSetting "SMT", "QSMS", "Comm", TxtComm

End Sub

Private Sub cmdexcel_Click()
Dim str As String
'Dim Rs As ADODB.Recordset
If Not Rs2.EOF Then
       Call CopyToExcel(Rs2)
    Else
       MsgBox ("No Data"), vbCritical
End If
End Sub

Private Sub cmdFind_Click()

    strSql = "Select top 30 DID,CompPN,VendorCode,DateCode,LotCOde,Qty,UID,remainQty,TransDateTime,UsedFlag From QSMS_DID Where CompPN like '" & Trim(CboCompPN) & "%' and DID like '" & Trim(CboDID) & "%' " & _
              " Order by CompPN,DID desc "
    Set Rs2 = Conn.Execute(strSql)
    Set DG1.DataSource = Rs2
    
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdSave.Enabled = True
    cmdExcel.Enabled = True
End Sub
'Private Sub CmdDelete_Click()
'Dim strsql As String
'Dim rs As ADODB.Recordset
'    If MsgBox("Are you Sure ??", vbYesNo) = vbNo Then
'       Exit Sub
'    End If
'    cmdAdd.Enabled = True
'    cmdUpdate.Enabled = True
'    cmdDelete.Enabled = True
'    cmdSave.Enabled = True
'    cmdCancel.Enabled = True
'    cmdExit.Enabled = True
'    strsql = "Select UsedFlag,Qty,RemainQty,UID from QSMS_DID where DID='" & CboDID & "' and CompPN='" & Trim(CboCompPN) & "' "
'    Set rs = Conn.Execute(strsql)
'    If rs.EOF Then
'       MsgBox "can not find the DID and CompPN"
'       Exit Sub
'    End If
'    If UCase(Trim(rs!UID)) = "FUJIDB" Then
'       MsgBox "The DID is from FuJi,can not delete in QSMS "
'       Exit Sub
'    End If
'    If UCase(Trim(rs!usedflag)) = "Y" Or rs!RemainQty <> rs!Qty Then
'       MsgBox "The DID has been used,can not delete"
'       Exit Sub
'    Else
'       strsql = "Delete from QSMS_DID Where DID='" & CboDID & "' and CompPN='" & Trim(CboCompPN) & "' and UsedFlag='N' "
'       Conn.Execute (strsql)
'    End If
'
'
'    Call RefreshDg(CboCompPN)
'    CboVendorCode.Text = ""
'    CboDateCode.Text = ""
'    CboLotCode.Text = ""
'    TxtQty.Text = ""
'    CboDID.Text = ""
'End Sub

Private Sub CmdDelete_Click() '--0005
Dim strSql As String
On Error GoTo errHandler
Dim rs As ADODB.Recordset

'''''''''''''''Add by Jing  20071126   (0006) '''''''''''
    frmLogin.Authorized = False
    BU = ReadIniFile("Common", "BU", App.Path & "\set.ini")
    
    ''''''Updated by Jing 2008.01.10    (0012)''''''
    'If BU = "NB25" Or BU = "NB3" Then
    If BU = "NB5" Or BU = "NB3" Then
        frmLogin.delright = "DeleteDID"
        frmLogin.Show vbModal
        If frmLogin.Authorized = False Then
            Exit Sub
        End If
    End If
    
    If MsgBox("Are you Sure to Delete DID from QSMS system ?!", vbYesNo) = vbNo Then
       Exit Sub
    End If
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdExit.Enabled = True

'    strsql = "Select UsedFlag,Qty,RemainQty,UID from QSMS_DID where DID='" & CboDID & "' and CompPN='" & Trim(CboCompPN) & "' "
'    Set Rs = Conn.Execute(strsql)
'    If Rs.EOF Then
'       MsgBox "Can not find the DID and CompPN"
'       Exit Sub
'    End If
''**Sandy        2007.11.30     added that don't allow to delete dispatched DID in MaintainDID----(0008)
'    If UCase(Trim(Rs!usedflag)) = "Y" Or Rs!RemainQty <> Rs!Qty Then
'       MsgBox "The DID has been used,can not delete"
'       Exit Sub
'    End If
''******************************************
'    strsql = "select machine,MSlot From qsms_feederdid_current where did='" & CboDID & "'"
'    Set Rs = Conn.Execute(strsql)
'    If Not Rs.EOF Then
'       MsgBox "The DID is using at " & Rs!Machine & ", " & Rs!mslot & ",please check it! "
'       Exit Sub
'    End If

    
    If BU = "" Then
        MsgBox "Can not find the values of BU , Call QMS for help please !", vbInformation
        Exit Sub
    Else
    '**Sandy        2007.12.19     checking DID dispatched or DIDSlotlink TO combine the delete DiD ----(0010)
        strSql = "Exec DeleteDIDByBU '" & BU & "','" & g_userName & "','" & CboDID & "'"
        Set rs = Conn.Execute(strSql)
        If Not rs.EOF Then
            If rs("ErrorCode") <> 0 Then
                MsgBox rs("Result")
                Exit Sub
            End If
        End If
    End If

    Call RefreshDg(CboCompPN)
    CboVendorCode.Text = ""
    CboDateCode.Text = ""
    CboLotCode.Text = ""
    TxtQty.Text = ""
    CboDID.Text = ""
Exit Sub
errHandler: MsgBox Err.Description & "delete fail!"
End Sub

Private Sub CmdRefresh_Click()
Call RefreshDg("")
End Sub

Private Sub cmdReprint_Click()
Dim str As String
Dim rs As ADODB.Recordset

''''''''''''''''''''''''''Add by Jing  20071126   (0006) ''''''''''''''''''''''''''''
frmLogin.Authorized = False
frmLogin.delright = "RePrintDID"
frmLogin.Show vbModal
If frmLogin.Authorized = False Then
    Exit Sub
End If

If Trim(CboDID) = "" Then
   MsgBox "Please input the DID"
   CboDID.SetFocus
   Exit Sub
End If

If IsInteger(TxtQty) = False Then
    MsgBox ("Check the qty of print please !"), vbCritical
    TxtQty.SetFocus
    Exit Sub
End If

str = "Select DID,firstmachine from QSMS_DID Where DID='" & Trim(CboDID) & "'"
Set rs = Conn.Execute(str)
If rs.EOF Then
   MsgBox "can not find the DID,Please check"
   CboDID.SetFocus
   Exit Sub
ElseIf UCase(Trim(rs!FirstMachine)) = "RETURN" Or UCase(Trim(rs!FirstMachine)) = "CALLBACK" Then  ''(0041)
    MsgBox "Can not do reprint, this DID has been " + Trim(rs!FirstMachine) + "ed !"
    CboDID.SetFocus
    Exit Sub
End If

Call PrintLabel
End Sub

Private Sub cmdSave_Click()
   Dim strSql As String
    Dim rs As ADODB.Recordset
    Dim TempDID As String
    Dim TransDate As String
    Dim I As Long, RetryCnt As Integer
    Dim InsertDIDOk As Boolean
    
    Dim adoCmd As New ADODB.Command
    Dim p1 As Integer, p2 As String
    
    
    cmdAdd.Enabled = True
    cmdFind.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdExit.Enabled = True
    CboCompPN.Text = Replace(Replace(Replace(CboCompPN.Text, " ", ""), vbCr, ""), vbLf, "")     '--------(7)
    If IsInteger(TxtGroupQty) = False Or IsInteger(TxtQty) = False Then
        MsgBox "Check the DID qty and Print qty please ！", vbCritical
        Exit Sub
    End If
    

    CboVendorCode = Trim(CboVendorCode)
    If Len(CboVendorCode) > 7 Then      'to check whether the vendor length is less than 8  (0015)
            MsgBox ("Vendor Code must be less than 8"), vbCritical
            Exit Sub
    End If
    
    '******************************
    '****add by jeanson 2007/09/03
    strErrMessage = ""
    strErrMessage = FunPartNumberCheck(CboCompPN.Text)
    If strErrMessage <> "PASS" Then
        MsgBox strErrMessage
        CboCompPN.SetFocus
        Exit Sub
    End If
    '******************************
'    If Len(Trim(CboCompPN)) <> 11 Or Trim(CboVendorCode) = "" Or Trim(CboDateCode) = "" Or Trim(CboLotCode) = "" Or Trim(txtQty) = "" Or Trim(TxtGroupQty) = "" Or txtQty > 20000 Then
    If Trim(CboVendorCode) = "" Or Trim(CboDateCode) = "" Or Trim(CboLotCode) = "" Or Trim(TxtQty) = "" Or Trim(TxtGroupQty) = "" Or TxtQty > 60000 Then '(1050)
        MsgBox ("Verdorcode or datecode or lotcode or qty can't be empty,DID Qty can't over 60000 ！"), vbCritical '(1050)
        CboCompPN.Enabled = True
        CboCompPN.SetFocus
        Exit Sub
    End If
    
    If ChkAVL(Trim(CboCompPN), Trim(CboVendorCode)) = False Then
        CboVendorCode.Text = ""
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''Add by Jing  20071126   (0006) ''''''''''''''''''''''''''''
    strSql = "select * from QSMS_DID where did='" & CboDID.Text & "'"
    Set rs = Conn.Execute(strSql)
    If Not rs.EOF Then
        MsgBox ("This DID have exist !")
        Exit Sub
    End If
    '----------------------add by Sandy 20080326-(0014)----------------
    If InStr(1, Trim(CboVendorCode), "'") <> 0 Then
        MsgBox ("Please check VendorCode,Don't allow to include '")
        Exit Sub
    End If
    If InStr(1, Trim(CboDateCode), "'") <> 0 Then
        MsgBox ("Please check DateCode,Don't allow to include '")
        Exit Sub
    End If
    If InStr(1, Trim(CboLotCode), "'") <> 0 Then
        MsgBox ("Please check LotCode,Don't allow to include '")
        Exit Sub
    End If
    '''''''''(1011)''''''''''
    If UCase(Trim(StrBU)) = "AS" Then
        If CheckDataCode(txtInspection, CboVendorCode, CboDateCode) = False Then
            Exit Sub
        End If
'        If ChkDateCodeSpecial(Trim(CboVendorCode), Trim(CboCompPN), Trim(CboDateCode)) = False Then
'            Exit Sub
'        End If
    Else
        txtInspection = ""
    End If
    
    If UCase(Trim(ChkDateCode)) = "Y" Then ''1222
        If ChkDateCodeSpecial(Trim(CboVendorCode), Trim(CboCompPN), Trim(CboDateCode)) = False Then
           Exit Sub
        End If
    End If
    
    If UCase(Trim(ScanMSD)) = "Y" Then
        If CheckMSD(Trim(txtMSD), CboCompPN) = False Then
            Exit Sub
        End If
    Else
        txtMSD = ""
    End If
    ''''''''''''''(1011)''''''''''''''''
    
    '----------------------add by Sandy 20080326-(0014)----------------
    strSql = "select getdate()"
    Set rs = Conn.Execute(strSql)
    TransDate = Format(rs(0), "YYYYMMDDHHMMSS")
    
    Select Case CommandType
        Case 1
            LockTheForm (False)
            For I = 1 To CLng(Trim(TxtGroupQty))
                RetryCnt = 0
                InsertDIDOk = False
                
                
                'If there are more than one person add same component DID, the DID# may conflict
                On Error GoTo Retry
                
                ''----------------------------------------------------------------------------------------------
                ''20100508  Denver   add Check CompPN(defined: BAM16400002;BAM16410002) length must be 5 Codes.
                Set adoCmd = New ADODB.Command
                adoCmd.ActiveConnection = Conn
                adoCmd.CommandText = "GenRegisterDID"
                adoCmd.CommandType = adCmdStoredProc
                
                adoCmd.Parameters.Append adoCmd.CreateParameter("@DID", adVarChar, adParamInput, 30, TempDID)
                adoCmd.Parameters.Append adoCmd.CreateParameter("@CompPN", adVarChar, adParamInput, 50, Trim(CboCompPN))
                adoCmd.Parameters.Append adoCmd.CreateParameter("@Qty", adInteger, adParamInput, 4, CInt(TxtQty))
                adoCmd.Parameters.Append adoCmd.CreateParameter("@VendorCode", adVarChar, adParamInput, 50, Left(Trim(CboVendorCode), 7))
                adoCmd.Parameters.Append adoCmd.CreateParameter("@DateCode", adVarChar, adParamInput, 50, Trim(CboDateCode))
                adoCmd.Parameters.Append adoCmd.CreateParameter("@LotCode", adVarChar, adParamInput, 50, Trim(CboLotCode))
                adoCmd.Parameters.Append adoCmd.CreateParameter("@DIDLoc", adVarChar, adParamInput, 50, Trim(txtInspection))         '''''''''(1011)''''''''''
                adoCmd.Parameters.Append adoCmd.CreateParameter("@DIDMEM", adVarChar, adParamInput, 50, "")
                adoCmd.Parameters.Append adoCmd.CreateParameter("@UID", adVarChar, adParamInput, 20, g_userName)
                adoCmd.Parameters.Append adoCmd.CreateParameter("@TransDateTime", adChar, adParamInput, 14, TransDate)
                adoCmd.Parameters.Append adoCmd.CreateParameter("@Type", adChar, adParamInput, 1, 0)
                adoCmd.Parameters.Append adoCmd.CreateParameter("@MSD", adVarChar, adParamInput, 50, Trim(txtMSD))                   '''''''''(1011)''''''''''
                adoCmd.Parameters.Append adoCmd.CreateParameter("@RtnCode", adInteger, adParamOutput)   ''0019
                adoCmd.Parameters.Append adoCmd.CreateParameter("@RtnMessage", adVarChar, adParamOutput, 1000)
                
                Do While Not InsertDIDOk And RetryCnt < 10
                    '**Sandy        2007.12.2  update to drive out the blank space in DID.---------(0007)
                    TempDID = Trim(GetDID(Trim(CboCompPN), TransDate))
'                    strSql = "exec GenRegisterDID '" & TempDID & "','" & Trim(CboCompPN) & "'," & TxtQty & ",'" & Left(Trim(CboVendorCode), 7) & "','" & Trim(CboDateCode) & "','" & Trim(CboLotCode) & "','','','" & g_userName & "','" & TransDate & "'"
'                    Conn.Execute (strSql)
'                    InsertDIDOk = True
                    
                    adoCmd.Parameters.Item("@DID").Value = TempDID
                    adoCmd.Execute
                    p1 = adoCmd.Parameters("@RtnCode").Value
                    p2 = adoCmd.Parameters("@RtnMessage").Value & ""
                    If p1 < 0 Then
                        MsgBox p2
                        
                         LockTheForm (True)
                        cmdExcel.Enabled = False
                        Exit Sub
                    ElseIf p1 > 0 Then
                        InsertDIDOk = True
                    End If
                    
Retry:
                    RetryCnt = RetryCnt + 1
                    DoEvents
                Loop
                
                CboDID = TempDID
                Call PrintLabel
                
'                if SATO printer will not delay time                  (0003)
                If OptZebra.Value = True Then
                    Call Delay_Time(1)
                End If
            Next I
            
            ''----------------------------------------------------------------------------------------------
            Call cmdFind_Click
            LockTheForm (True)
            cmdExcel.Enabled = False
            
        Case 2
            If Trim(CboDID) = "" Then
                MsgBox ("DID can't be empty!!"), vbCritical
                CboDID.Enabled = True
                CboDID.SetFocus
                Exit Sub
            End If
            
            If ChkAVL(Trim(CboCompPN), Trim(CboVendorCode)) = False Then
                CboVendorCode.Text = ""
                Exit Sub
            End If
    
            TempDID = Trim(CboDID)
            strSql = "Select Qty,RemainQty from QSMS_DID where DID='" & Trim(TempDID) & "'"
            Set rs = Conn.Execute(strSql)
            'If rs!RemainQty < rs!Qty Or CLng(TxtQty) > rs!Qty Then
            If rs!RemainQty <> rs!Qty Then  '***Modify by jeanson 2007/10/15    (0001)
               MsgBox "The DID is using,can not update"
               Exit Sub
            End If
            
            strSql = "Update QSMS_DID Set CompPN='" & Trim(CboCompPN) & "',VendorCode='" & Trim(CboVendorCode) & "',DateCode='" & CboDateCode & "',LotCode='" & Trim(CboLotCode) & "',Qty='" & Trim(TxtQty) & "' " & _
                    ",RemainQty=" & TxtQty & ",RealQty=" & TxtQty & ",LastUpdateDT='" & TransDate & "' Where DID='" & Trim(CboDID) & "'"          '(0004)
'            strsql = "exec UpdateDIDData @DID='" & Trim(CboDID) & "',@CompPN='" & Trim(CboCompPN) & "', @VendorCode='" & Trim(CboVendorCode) & "',@DateCode='" & CboDateCode & "',@LotCode='" & Trim(CboLotCode) & "',@Qty=" & Trim(txtQty) & ",@RemainQty=" & Trim(txtQty) & ",@RealQty=" & Trim(txtQty) & ",@LastUpdateDT='" & TransDate & "',@Type='1'"
            Conn.Execute strSql
            
             If Trim(CboDID) = "" Then
                      CboDID = TempDID
             End If
            Call cmdFind_Click
    End Select
    Call RefreshDg("")
 
    CommandType = 0
    TxtGroupQty = 1
    Call cmdCancel_Click
End Sub

Private Sub cmdUpdate_Click()
    cmdAdd.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdExit.Enabled = True
    cmdFind.Enabled = True
    
    CboCompPN.Enabled = True
    CboVendorCode.Enabled = True
    CboDateCode.Enabled = True
    CboLotCode.Enabled = True
    TxtQty.Enabled = True
    

    CommandType = 2
End Sub

Private Sub CmdVendorCode_Click()
Dim str As String
Dim rs As ADODB.Recordset

str = "select * from QSMS_DID where vendorcode = '" & CboVendorCode & "'order by CompPN"

Set rs = Conn.Execute(str)
 If Not rs.EOF Then
       Call CopyToExcel(rs)
    Else
       MsgBox ("No Data"), vbCritical
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DG1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    On Error Resume Next
    With DG1
    ''''''''''''''''''''''''''''''(0013)
        CboDID = ""
        CboCompPN = .Columns(1).Value
        CboVendorCode = ""
        CboDateCode = ""
        CboLotCode = ""
        TxtQty = ""
    End With
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    cmdCancel.Enabled = True
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

Dim str As String
Dim rs As ADODB.Recordset

'TxtCompPort.Text = GetSetting("SMT", "QSMS", "CommPort", "1")
'TxtComm.Text = GetSetting("SMT", "QSMS", "Comm", "9600,N,8,1")

'20101115 Maggie Save Printer setting in local Registry (1019)
Call GetPrinterSetting(FrmMaintainDID)

''(1080)
If OptZebra.Value = True Then
    isZebra = True
Else
    isZebra = False
End If

If UCase(Trim(StrBU)) = "AS" Then
    labInpectionNo.Visible = True
    txtInspection.Visible = True
End If

If UCase(ScanMSD) = "Y" Then
    labelMSD.Visible = True
    txtMSD.Visible = True
End If

End Sub

Private Function RefreshCompPN()
Dim str As String
Dim rs As ADODB.Recordset
str = "select CompPn from QSMS_RackID order by CompPN"
Set rs = Conn.Execute(str)
CboCompPN.Clear
While Not rs.EOF
      CboCompPN.AddItem Trim(rs!compPN) & vbNullString
      
      rs.MoveNext
Wend

End Function

Private Function RefreshDg(ByVal compPN As String)
Dim str As String
'not to display the DID which Qty=0  (0016)
str = "select top 20 DID,CompPN,VendorCode,DateCode,LotCode,Qty,RemainQty,UID,TransDateTime from QSMS_DID where CompPN like '" & compPN & "%' and Qty<>0 order by DID desc"
Set Rs2 = Conn.Execute(str)
Set DG1.DataSource = Rs2
DG1.Refresh
End Function

'Private Function PrintLabel()
'If OptComp.Value = True Then
   'Call PrintLabelCompPort
'End If
'If OptPrint.Value = True Then
   'Call PrintLabelPrintPort
'End If
'End Function
Private Function PrintLabel() As String
Dim I As Integer
Dim lptPort As Integer
Dim hFile As Long
Dim hString As String
Dim strDID As String, tmpDID As String, strQty As String
Dim strDay As String, strSql As String
Dim LabelFile As String
Dim rsTmp As ADODB.Recordset
Dim tmpPrintStr As String
          '''(1224)''''
Dim rsTime As ADODB.Recordset
On Error GoTo errHandler
       
 BU = ReadIniFile("Common", "BU", App.Path & "\set.ini")
        strSql = "select getdate()"
        Set rsTime = Conn.Execute(strSql)
        If BU = "NB5" Then
        strDay = Format(rsTime(0), "YYYYMMDDHHNNSS")
        Else
        strDay = Format(Now, "YYYY/MM/DD")
        End If
        '''(1224)'''
        ''(1080) replace by 1080
'        If OptZebra.Value = True Then
'            isZebra = True
'            LabelFile = Settings.LabelAFile
'        Else
'            isZebra = False
'            LabelFile = Settings.LabelSATOFIle
'        End If

        LabelFile = GetDIDLabelFile(FrmMaintainDID) ''(1080) get labelfile
        strDID = UCase(Trim(CboDID))
        strQty = Trim(TxtQty.Text)
        
        If Dir(LabelFile) = vbNullString Then
            ''''''Added by Jing 2008.01.10  (0011)''''''
            MsgBox ("Can not find label file !"), vbCritical
            PrintLabel = "PRN_FileNoExist"
            Exit Function
        End If
        
        If OptComp.Value = True Then
            MSComm.CommPort = TxtCompPort 'Settings.PRNa_Port
            MSComm.Settings = TxtComm 'Settings.PRNa_Settings
            MSComm.OutBufferCount = 0 '清空输出缓存
            If MSComm.PortOpen = False Then MSComm.PortOpen = True
        ElseIf OptPrint.Value = True Then
            lptPort = OpenOutputFile("LPT1")
            If lptPort = 0 Then
                MsgBox "Open print port LPT1 error!"
                Exit Function
            End If
        End If

        tmpPrintStr = ""
        hFile = FreeFile
    ''''''Updated by Kyle 2010.08.10  (0077)''''''
    If FileReadAll(tmpPrintStr, LabelFile) <= 0 Then
        MsgBox "Open file:" & LabelFile & " fail!!", vbCritical
    Exit Function
    End If
'    Open LabelFile For Input As #hFile
'    Do
'       Select Case EOF(hFile)
'          Case True
'            Close #hFile
'            PrintLabel = "PRN_Succeed"
'            Exit Do
'          Case False
'            Line Input #hFile, hString
'            hString = Trim(hString)
'            tmpPrintStr = tmpPrintStr & Trim(hString)
'      End Select
'    Loop
'    Close #hFile
    ''''''Updated by Kyle 2010.08.10  (0077)''''''
        
        tmpDID = Trim(strDID) '***************add by jeanson 20070814******
         'for Code 128 barcode, the ^ must be tranfer to ><
        If InStr(tmpPrintStr, "<DID_CODE>") > 0 Then
             ''********************************updated by jing 20071024 (0002) ***********
             If isZebra Then
                 tmpDID = Replace(strDID, "^", "><")
             End If
            tmpPrintStr = Replace(tmpPrintStr, "<DID_CODE>", tmpDID)
         End If
         
         'for text ^, must be use ^FH_ and the use _5E (the ascii of ^)
         If InStr(tmpPrintStr, "<DID_TEXT>") > 0 Then
             ''********************************updated by jing 20071024 (0002) ***********
             If isZebra Then
                tmpDID = Replace(strDID, "^", "_5E")
             End If
            tmpPrintStr = Replace(tmpPrintStr, "<DID_TEXT>", tmpDID)
         End If

         tmpPrintStr = Replace(tmpPrintStr, "<UID>", UID)
'                hString = Replace(hString, "<RACKID>", TxtRackID)
         tmpPrintStr = Replace(tmpPrintStr, "<DATE>", strDay)
         tmpPrintStr = Replace(tmpPrintStr, "<QTY>", strQty)
         
        Select Case Trim(tmpPrintStr)
           Case vbNullString
           Case Else
                If OptComp.Value = True Then
                    For I = 1 To Len(tmpPrintStr) Step 100
                        MSComm.Output = Mid(tmpPrintStr, I, 100)
                        DoEvents
                    Next I
                    MSComm.PortOpen = False
                ElseIf OptPrint.Value = True Then
                    For I = 1 To Len(tmpPrintStr) Step 50
                        Print #lptPort, Mid(tmpPrintStr, I, 50)
                        DoEvents
                    Next I
                    Close #lptPort
                Else
                    Printer.Print tmpPrintStr
                    Printer.EndDoc
                    Printer.KillDoc
                End If
        End Select
        Exit Function
        
errHandler:
        MsgBox Err.Description
        If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
        End If
End Function

Private Function PrintLabelCompPort() As String
Dim hFile As Long
Dim hString As String
Dim strDID As String, tmpDID As String, strQty As String
Dim strDay As String, strSql As String
Dim LabelFile As String
Dim rsTmp As ADODB.Recordset
        
On Error GoTo errHandler
        strDay = Format(Now, "YYYY/MM/DD")
        If OptZebra.Value = True Then
            isZebra = True
            LabelFile = Settings.LabelAFile
        Else
            isZebra = False
            LabelFile = Settings.LabelSATOFIle
        End If
        strDID = UCase(Trim(CboDID))
        strQty = Trim(TxtQty.Text)
        
        If Dir(LabelFile) = vbNullString Then
            ''''''Added by Jing 2008.01.10  (0011)''''''
            MsgBox ("Can not find label file !"), vbCritical
            PrintLabelCompPort = "PRN_FileNoExist"
            Exit Function
        End If
        
        MSComm.CommPort = TxtCompPort 'Settings.PRNa_Port
        MSComm.Settings = TxtComm 'Settings.PRNa_Settings
        MSComm.OutBufferCount = 0 '清空输出缓存
        
        If MSComm.PortOpen = False Then MSComm.PortOpen = True
        
        hFile = FreeFile
        Open LabelFile For Input As #hFile
        Do
           Select Case EOF(hFile)
              Case True
                Close #hFile
                PrintLabelCompPort = "PRN_Succeed"
                Exit Do
              Case False
                Line Input #hFile, hString
                hString = Trim(hString)
                tmpDID = Trim(strDID) '***************add by jeanson 20070814******
                'for Code 128 barcode, the ^ must be tranfer to ><
                If InStr(hString, "<DID_CODE>") > 0 Then
                    ''********************************updated by jing 20071024 (0002) ***********
                    If isZebra Then
                        tmpDID = Replace(strDID, "^", "><")
                    End If
                   hString = Replace(hString, "<DID_CODE>", tmpDID)
                End If
                'for text ^, must be use ^FH_ and the use _5E (the ascii of ^)
                If InStr(hString, "<DID_TEXT>") > 0 Then
                    ''********************************updated by jing 20071024 (0002) ***********
                    If isZebra Then
                       tmpDID = Replace(strDID, "^", "_5E")
                    End If
                   hString = Replace(hString, "<DID_TEXT>", tmpDID)
                End If

                hString = Replace(hString, "<UID>", UID)
'                hString = Replace(hString, "<RACKID>", TxtRackID)
                hString = Replace(hString, "<DATE>", strDay)
                hString = Replace(hString, "<QTY>", strQty)
                
               Select Case Trim(hString)
                  Case vbNullString
                  Case Else
                    MSComm.Output = hString
                    Debug.Print hString
               End Select
          End Select
        Loop
       
        Close #hFile
        MSComm.PortOpen = False
        Exit Function
        
errHandler:
        MsgBox Err.Description
        If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
        End If
End Function

Private Function PrintLabelPrintPort() As String
Dim hFile As Long
Dim hString As String
Dim strDID As String, tmpDID As String, strQty As String
Dim FileNum As Integer, lptPort As Integer
Dim strDay As String
Dim LabelFile, strLabelFileContent As String
Dim strPort As String, strSql As String
Dim rsTmp As ADODB.Recordset
        
On Error GoTo errhandle
    strDay = Format(Now, "YYYY/MM/DD")
    strDID = UCase(Trim(CboDID))
    strQty = Trim(TxtQty)
        
    If OptZebra.Value = True Then
        isZebra = True
        LabelFile = Settings.LabelAFile
    Else
        isZebra = False
        LabelFile = Settings.LabelSATOFIle
    End If
    
'        strLabelFileContent = funGetTxtFileContent(LabelFile)

    If Dir(LabelFile) = vbNullString Then
        ''''''Added by Jing 2008.01.10  (0011)''''''
        MsgBox ("Can not find label file !"), vbCritical
        PrintLabelPrintPort = "PRN_FileNoExist"
        Exit Function
    End If
    
    lptPort = OpenOutputFile("LPT1")
    If lptPort = 0 Then
        MsgBox "Open print port LPT1 error!"
        Exit Function
    End If
    
'        strPort = "LPT1"

'        strLabelFileContent = Replace(strLabelFileContent, "<DID>", CboDID)
'
'        strLabelFileContent = Replace(strLabelFileContent, "<UID>", UID)
'        strLabelFileContent = Replace(strLabelFileContent, "<RACKID>", TxtRackID)
'        strLabelFileContent = Replace(strLabelFileContent, "<DATE>", strDay)

    FileNum = FreeFile()
    Open LabelFile For Input As #FileNum
    While Not EOF(FileNum)
       Line Input #FileNum, hString
            hString = Trim(hString)
            tmpDID = Trim(strDID)  '***************add by jeanson 20070814******
            'for Code 128 barcode, the ^ must be tranfer to ><
            If InStr(hString, "<DID_CODE>") > 0 Then
                ''********************************updated by jing 20071024 (0002) ***********
                If isZebra Then
                    tmpDID = Replace(strDID, "^", "><")
                End If
               hString = Replace(hString, "<DID_CODE>", tmpDID)
            End If
            'for text ^, must be use ^FH_ and the use _5E (the ascii of ^)
            If InStr(hString, "<DID_TEXT>") > 0 Then
                ''********************************updated by jing 20071024 (0002) ***********
                If isZebra Then
                   tmpDID = Replace(strDID, "^", "_5E")
                End If
               hString = Replace(hString, "<DID_TEXT>", tmpDID)
            End If
            
            hString = Replace(hString, "<UID>", UID)
'                hString = Replace(hString, "<RACKID>", TxtRackID)
            hString = Replace(hString, "<DATE>", strDay)
            hString = Replace(hString, "<QTY>", strQty)
            
            Print #lptPort, hString & Chr(13)
    Wend
'        Open strPort For Output As #FileNum
'        Print #FileNum, strLabelFileContent
    Close #FileNum
    Close #lptPort
    Exit Function
    
errhandle:
     MsgBox Err.Description
End Function

Function OpenOutputFile(ByVal fname As String)
  Dim Fnumber As Integer
  
  On Error GoTo ErrorProcedure
  OpenOutputFile = 0
  Fnumber = FreeFile
  OpenOutputFile = Fnumber
  Open fname For Output As #Fnumber
  Exit Function
ErrorProcedure:
  OpenOutputFile = 0
End Function


Private Function funGetTxtFileContent(ByVal strFile As String) As String
    Dim NFile As Long
    Dim strFileContent As String
    NFile = FreeFile()
    Open strFile For Input As #NFile
        strFileContent = Input(LOF(NFile), NFile)
    Close NFile
    funGetTxtFileContent = strFileContent
End Function

Private Sub TxtGroupQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TxtGroupQty <> "" Then
    CboCompPN.SetFocus
End If
End Sub

Private Sub TxtQty_Click()
    SendKeys "{HOME}+{END}"
End Sub

'****************************
'*****add Qty check by jeanson 2006/09/07
Private Sub TxtQty_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13 Or KeyAscii = 9) And TxtQty <> "" Then
    If IsInteger(TxtQty) = False Then
        MsgBox ("Check the qty of print please !"), vbCritical
        TxtQty.SetFocus
        Exit Sub
    End If
    cmdAdd.SetFocus
End If
End Sub

Private Sub LockTheForm(lockCtl As Boolean)
 'On Error Resume Next
 Dim ctl As Control
 
' For Each ctl In Me.Controls
'
'   Debug.Print ctl
'   If ctl <> False Or ctl <> True Then
'   ctl.Enabled = lockCtl
'   End If
'
' Next ctl
 OptComp.Enabled = lockCtl
 OptPrint.Enabled = lockCtl
 CmdCommSave.Enabled = lockCtl
 TxtGroupQty.Enabled = lockCtl
 CboCompPN.Enabled = lockCtl
 CboVendorCode.Enabled = lockCtl
 CboDateCode.Enabled = lockCtl
 CboLotCode.Enabled = lockCtl
 TxtQty.Enabled = lockCtl
 CboDID.Enabled = lockCtl
 cmdFind.Enabled = lockCtl
 cmdAdd.Enabled = lockCtl
 cmdUpdate.Enabled = lockCtl
 cmdDelete.Enabled = lockCtl
 cmdSave.Enabled = lockCtl
 cmdCancel.Enabled = lockCtl
 CmdRefresh.Enabled = lockCtl
 cmdExit.Enabled = lockCtl
 CmdReprint.Enabled = lockCtl
 DG1.Enabled = lockCtl
 End Sub

Private Function ChkAVL(ByVal compPN As String, ByVal VendorCode As String) As Boolean
Dim strSql As String
Dim rs As ADODB.Recordset
ChkAVL = True
If Check_AVL <> "Y" Then
    ChkAVL = True 'add by Giant  --20070618
Else
    strSql = "Select TOP 1 * from QSMS_AVL where CompPN='" & Trim(CboCompPN) & "' and VendorCode='" & Trim(CboVendorCode) & "' "
    Set rs = Conn.Execute(strSql)
    If rs.EOF Then
        ChkAVL = False
        MsgBox "CompPN and VendorCode not match!! please check "
    End If
End If

End Function
