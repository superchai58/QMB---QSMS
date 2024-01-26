VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmReturnDID 
   Caption         =   "Return DID   2023-05-16"
   ClientHeight    =   10530
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10530
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraReturnDIDALL 
      BackColor       =   &H80000013&
      Caption         =   "Return DID ALL"
      Height          =   1365
      Left            =   120
      TabIndex        =   45
      Top             =   9600
      Width           =   13215
      Begin MSCommLib.MSComm MSComm1 
         Left            =   12480
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin VB.CommandButton CmdChkWO 
         BackColor       =   &H008080FF&
         Caption         =   "&Check if can close"
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
         Left            =   2160
         Picture         =   "FrmReturnDID.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CmdCloseWO 
         BackColor       =   &H0000FF00&
         Caption         =   "&Close WO"
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
         Left            =   4080
         Picture         =   "FrmReturnDID.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CmdReturnDIDAll 
         Caption         =   "&Return ALL"
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
         Left            =   240
         Picture         =   "FrmReturnDID.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label LblALL 
         BackColor       =   &H80000000&
         Height          =   405
         Left            =   120
         TabIndex        =   47
         Top             =   870
         Width           =   6255
      End
   End
   Begin VB.Frame FraReturnDID 
      BackColor       =   &H80000013&
      Caption         =   "Return DID"
      Height          =   2835
      Left            =   120
      TabIndex        =   20
      Top             =   6750
      Width           =   13215
      Begin VB.Frame Frame2 
         Height          =   450
         Left            =   8880
         TabIndex        =   68
         Top             =   240
         Width           =   4275
         Begin VB.CheckBox chkAutoGetRefID 
            Caption         =   "AutoGetRefID(No RefID Label)"
            Enabled         =   0   'False
            Height          =   345
            Left            =   2400
            TabIndex        =   78
            Top             =   120
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.OptionButton optGoodMaterial 
            Caption         =   "Good"
            Height          =   345
            Left            =   120
            TabIndex        =   77
            Top             =   120
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton optBadMaterial 
            Caption         =   "Bad"
            Height          =   345
            Left            =   960
            TabIndex        =   76
            Top             =   120
            Width           =   615
         End
         Begin VB.CheckBox ChkHUA 
            Caption         =   "HUA"
            Height          =   345
            Left            =   1680
            TabIndex        =   75
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdReprint 
         Caption         =   "&Reprint"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   720
         Width           =   1200
      End
      Begin VB.CommandButton cmdGetRefID 
         Caption         =   "&GetRefID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   720
         Width           =   1200
      End
      Begin VB.Frame FraPrinter 
         BackColor       =   &H80000013&
         Caption         =   "Printer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   54
         Top             =   1950
         Width           =   13035
         Begin VB.Frame Frame3 
            Height          =   615
            Left            =   5500
            TabIndex        =   71
            Top             =   180
            Width           =   1665
            Begin VB.OptionButton opNewLabel 
               Caption         =   "New"
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
               Left            =   840
               TabIndex        =   73
               Top             =   240
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton opOldLabel 
               Caption         =   "Old"
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
               Left            =   120
               TabIndex        =   72
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame4 
            Height          =   615
            Left            =   120
            TabIndex        =   61
            Top             =   180
            Width           =   1800
            Begin VB.OptionButton OptZebra 
               Caption         =   "Zebra"
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
               Left            =   60
               TabIndex        =   63
               Top             =   240
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton OptSATO 
               Caption         =   "SATO"
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
               Left            =   960
               TabIndex        =   62
               Top             =   240
               Width           =   855
            End
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
            Height          =   555
            Left            =   11400
            Picture         =   "FrmReturnDID.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   240
            Width           =   1485
         End
         Begin VB.TextBox TxtComm 
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
            Left            =   10000
            TabIndex        =   59
            Text            =   "9600,N,8,1"
            Top             =   300
            Width           =   1320
         End
         Begin VB.TextBox TxtCompPort 
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
            Left            =   8520
            TabIndex        =   58
            Text            =   "1"
            Top             =   300
            Width           =   500
         End
         Begin VB.Frame Frame5 
            Height          =   615
            Left            =   1920
            TabIndex        =   55
            Top             =   180
            Width           =   3600
            Begin VB.OptionButton optNetwork 
               Caption         =   "NetWork"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2400
               TabIndex        =   74
               Top             =   240
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton OptPrint 
               Caption         =   "Print Port"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1320
               TabIndex        =   57
               Top             =   240
               Width           =   1155
            End
            Begin VB.OptionButton OptComp 
               Caption         =   "Comp Port"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   56
               Top             =   240
               Width           =   1365
            End
         End
         Begin VB.Label Label4 
            BackColor       =   &H0080FF80&
            Caption         =   "Settings"
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
            Index           =   6
            Left            =   9000
            TabIndex        =   65
            Top             =   300
            Width           =   1000
         End
         Begin VB.Label Label3 
            BackColor       =   &H0080FF80&
            Caption         =   "CompPort"
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
            Index           =   2
            Left            =   7200
            TabIndex        =   64
            Top             =   300
            Width           =   1320
         End
      End
      Begin VB.CommandButton cmdReturn 
         BackColor       =   &H0000FF00&
         Caption         =   "Update ReturnQty"
         Height          =   375
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2280
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox TxtCompPN 
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
         Height          =   450
         Left            =   6840
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   2000
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
         Left            =   8880
         Picture         =   "FrmReturnDID.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   720
         Width           =   1200
      End
      Begin VB.TextBox TxtDIDReturnedQty 
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
         Height          =   450
         Left            =   4560
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox TxtDIDTotalQty 
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
         Height          =   450
         Left            =   1680
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox TxtReturnQty 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6840
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   720
         Width           =   2000
      End
      Begin VB.TextBox TxtDID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1680
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   720
         Width           =   3885
      End
      Begin VB.Label lblFeedBack 
         Caption         =   "Qty FeedBack: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   345
         Left            =   120
         TabIndex        =   70
         Top             =   1710
         Width           =   8700
      End
      Begin VB.Label LblMessage 
         BackColor       =   &H80000000&
         Height          =   465
         Left            =   120
         TabIndex        =   32
         Top             =   1230
         Width           =   8700
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
         Height          =   450
         Index           =   3
         Left            =   5640
         TabIndex        =   31
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DID Returned Qty"
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
         Height          =   450
         Index           =   5
         Left            =   2760
         TabIndex        =   30
         Top             =   240
         Width           =   1800
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
         Height          =   450
         Index           =   17
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1600
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Return Qty"
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
         Height          =   450
         Index           =   2
         Left            =   5640
         TabIndex        =   28
         Top             =   720
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
         Height          =   450
         Index           =   14
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "DID information"
      Height          =   4125
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   13215
      Begin TabDlg.SSTab sstabDID 
         Height          =   2145
         Left            =   30
         TabIndex        =   50
         Top             =   210
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   3784
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "DID Return"
         TabPicture(0)   =   "FrmReturnDID.frx":0F32
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "DGDIDInfo"
         Tab(0).Control(1)=   "DGDIDNeedReturned"
         Tab(0).Control(2)=   "DGDIDReturned"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "DID To WareHouse Top 1000"
         TabPicture(1)   =   "FrmReturnDID.frx":0F4E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "gridDIDtoWH"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid DGDIDReturned 
            Height          =   1695
            Left            =   -69720
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   360
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   2990
            _Version        =   393216
            AllowUpdate     =   0   'False
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
            Caption         =   "Returned DID"
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
         Begin MSDataGridLib.DataGrid DGDIDNeedReturned 
            Height          =   1695
            Left            =   -74760
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   360
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   2990
            _Version        =   393216
            AllowUpdate     =   0   'False
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
            Caption         =   "Need Return DID"
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
         Begin MSDataGridLib.DataGrid DGDIDInfo 
            Height          =   1695
            Left            =   -66180
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   390
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   2990
            _Version        =   393216
            AllowUpdate     =   0   'False
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
            Caption         =   "Comp Information"
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
         Begin MSDataGridLib.DataGrid gridDIDtoWH 
            Height          =   1635
            Left            =   60
            TabIndex        =   69
            Top             =   450
            Width           =   12945
            _ExtentX        =   22834
            _ExtentY        =   2884
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   19
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
               Size            =   9.75
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
         Left            =   1560
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2400
         Width           =   3495
      End
      Begin MSDataGridLib.DataGrid DGCompInfo 
         Height          =   1215
         Left            =   60
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2880
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         Caption         =   "Comp Return & Dispatch information"
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
      Begin VB.Label LblChk 
         BackColor       =   &H80000000&
         Height          =   495
         Left            =   5160
         TabIndex        =   35
         Top             =   2340
         Width           =   7935
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "DID:"
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
         Left            =   0
         TabIndex        =   33
         Top             =   2400
         Width           =   1455
      End
   End
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   13095
      Begin VB.ComboBox CboReportType 
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
         ItemData        =   "FrmReturnDID.frx":0F6A
         Left            =   8760
         List            =   "FrmReturnDID.frx":0F6C
         Sorted          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1980
         Width           =   2775
      End
      Begin VB.CommandButton cmdExcel 
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
         Height          =   615
         Left            =   11520
         Picture         =   "FrmReturnDID.frx":0F6E
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1860
         Width           =   975
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   1680
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1320
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
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   900
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
         Left            =   6000
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2100
         Width           =   2655
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2100
         Width           =   2655
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "Group"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton OptRelease 
         Caption         =   "Release"
         Height          =   255
         Left            =   120
         TabIndex        =   5
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
         Left            =   3480
         Picture         =   "FrmReturnDID.frx":1278
         Style           =   1  'Graphical
         TabIndex        =   4
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
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   480
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
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2655
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
         Left            =   6000
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
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
         Format          =   134348803
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   900
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
         Format          =   134348803
         CurrentDate     =   36482
      End
      Begin VB.Label Lbl1 
         BackColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   9240
         TabIndex        =   44
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Report Type"
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
         Left            =   8760
         TabIndex        =   42
         Top             =   1680
         Width           =   1455
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
         TabIndex        =   38
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
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
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   480
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
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   17
         Top             =   900
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
         Left            =   4440
         TabIndex        =   16
         Top             =   2100
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
         TabIndex        =   15
         Top             =   1320
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
         Left            =   120
         TabIndex        =   14
         Top             =   2100
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
         Height          =   375
         Index           =   1
         Left            =   4440
         TabIndex        =   13
         Top             =   480
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
         TabIndex        =   12
         Top             =   1680
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
         Left            =   4440
         TabIndex        =   11
         Top             =   1680
         Width           =   1455
      End
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   90
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "FrmReturnDID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/**********************************************************************************
'**文 件 名: FrmReturnDID.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: LynnSun
'**日    期: 2007.12.03
'**描    述: QSMS Return DID
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**LynnSun      2007.12.03     If one DID dispatch again, we will allow this DID return again --------(00001)
'**Sandy        2007.12.29     don't check some data where sstabDID<>0--------(00002)
'**Jing         2008.01.10     Add Error Alarm when Can not find Label file     (00003)
'**Jing         2008.01.10     Changed from '0' to '4' for type(4:ReturnDID)      (00004)
'**Denver       2008.01.10     check WareHouseID when DID return firstly (00004)
'**Jeanson      2008.01.22     delete the data of QSMS_DID_ToWH (00005)
'**Sandy        2008.02.11     add ReturnDIDByGroupID and ReturnDIDByWO (00006)
'**Jeanson      2008.03.27     to check the form:frmDIDChkStock exists or not (00007)
'**Denver       2008.04.02     show return Qty and LineMC stock Qty   --(0008)
'**Denver       2008.04.07     it need not select GroupID and WO when Return DID   --(0009)
'**Udall        2008.04.13     Add Factory into the XL_DIDAutoDispatch SP    --(0010)
'**Kane         2008.04.17     Save log when autodispatch did not success--(0011)
'**Denver       2008.04.25     因为接料(A==>B) ,B 的Remain数量会〉总数，故放量30%(Move to procedure)--(0012)
'**Sandy        2008.05.06     update the Auto dispatch label  (0013)
'**Udall        2008.06.10     Add a function for query the DID which need to return when select the WO  (0014)
'**Kane         2008.06.11     Add wotype on did label,when did dispatched to wo which pilot is new then new,
'                              when pilot is eol and not exists pilot is new then eol else is empty (0015)
'**Sandy        2008.08.01      Cancel the function of UpdateReturnQty  (0016)
'**Jing         2008.05.30     Cancel:save the DID print log (0038)
'**Denver       2009.02.20      因为最初没有建立历史DB，最初的功能是为不影响效能，备份在QSMS不同的table. 现在可去掉。 （0039）
'**Sandy        2009.03.03     it will remind when ReturnQty>DIDRealQty  (0040)
'**Denver       2009.03.03     如果有选择GroupID,就需要检查属于GroupID，  不选不检查   --(0041)
'**Kane         2009.03.16     Save error log when generate unknow error (0042)
'*Sandy         2009.04.21     add tmpRS("LR") in new auto diapach (0043)
'*Denver        2009.07.23     NB2&NB3分家，但是用同一个QSMS可作NB2NB3 的退料
'*Austin        2009.11.03     记录Log(当line<>Left(machine,1)     (00044)
'*Kane          2009.12.28     打印自动发料条码时LPT口区分新旧Label'(00045)
'*RQ09122849    Archer          2010/03/06      Modify program to use new label format as default option(0046)
'*RQ10042001    Kane           2010/05/18       如果是禁用料则不能退料'(0047)
'QMS            Feix            20180403        ESBU添加提示信息（1267）
'***********************************************************************************/

Dim TempDID As String
Public returnDIDflag As Boolean
Private sSql As String
Private rstDIDtoWH As New ADODB.Recordset
Private Rst As New ADODB.Recordset
Dim PrintData As PtData
Dim PreGroupID As String
Dim isZebra As Boolean
Dim WOType As String
Dim IsAnotherBUDID As String
Dim strDelaytime As Long
Dim strCheckScaner As String





Private Sub CboGroupID_Click()
Dim str As String
Dim Rs As ADODB.Recordset
On Error GoTo errhandle
Lbl1.BackColor = &HFF&
Lbl1.Caption = "Please wait while get data"
DoEvents
DoEvents
DoEvents
If ChkGroupClosed(Trim(CboGroupID)) = True Then
   MsgBox "The Group has been closed,can not return DID"
   Exit Sub
End If
str = "Exec QSMSGroupCompQty '" & Trim(CboGroupID) & "'"
Conn.Execute str
Call GetGroupWO(CboGroupID)
Call GetReturned_NotReturnDID(Trim(CboGroupID))
Lbl1.Caption = ""
Lbl1.BackColor = &HC0C0C0
Exit Sub
errhandle:
  MsgBox "System Error,Please contact QMS"
  Resume Next
End Sub

Private Sub CboGroupID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboGroupID_Click
End If
End Sub

Private Sub CboWo_Click()
Dim strSQL As String
Dim Rs As ADODB.Recordset

TxtWO = Trim(cboWO)
Call GetWoinfo(TxtWO)
''---------------(0014)
 strSQL = "exec QSMS_WONeedReturnDID '" & TxtWO & "'"
 Set Rs = Conn.Execute(strSQL)
 If Rs.EOF = False Then
    Set DGDIDNeedReturned.DataSource = Rs
    DGDIDNeedReturned.Caption = "(Need Return DID ) Total: " & Rs.RecordCount
 End If
 ''---------------(0014)
End Sub

Private Sub cboWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboWo_Click
End If
End Sub
Private Sub CmdCommSave_Click()
    SaveSetting "SMT", "QSMS", "CommPort", TxtCompPort
    SaveSetting "SMT", "QSMS", "Comm", TxtComm
End Sub

Private Sub CmdChkWO_Click()
Dim str As String
Dim Rs As ADODB.Recordset
'str = "select Distinct DID from QSMS_GroupDID where GroupID='" & CboGroupID & "' and ReturnFlag<>'Y'  Order by DID"
'Set RS = Conn.Execute(str)
'If Not RS.EOF Then
'    MsgBox "The DID didn't return finish,please return first"
'     Call CopyToExcel(RS)
'End If
LblALL.Caption = "System is checking if WO can be closed ,please wait"
DoEvents
DoEvents
DoEvents
DoEvents
If ChkCLoseWoByAuto(CboGroupID) = True Then
    MsgBox "Check OK,you can close the GroupID Now."
    
 Else
       MsgBox "Can not close the wo,please check the excel file."
       Exit Sub
 End If
End Sub

Private Sub CmdCloseWO_Click()
Dim str As String
Dim Rs As ADODB.Recordset
str = "select Distinct DID from QSMS_GroupDID where GroupID='" & CboGroupID & "' and ReturnFlag<>'Y'  Order by DID"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
    MsgBox "The DID didn't return finish,please return first"
     Call CopyToExcel(Rs)
     Exit Sub
End If

LblALL.Caption = "System is checking if WO can be closed ,please wait"
DoEvents
DoEvents
DoEvents
DoEvents
 If ChkCLoseWoByAuto(CboGroupID) = True Then
       LblALL.Caption = "System is negerating SAP2 data,please wait"
       If UpdateQSMSGroupCompQty(CboGroupID) = True Then
            LblALL.Caption = "System is backup data,please wait"
            DoEvents
            DoEvents
            DoEvents
            DoEvents
            Call DeleteDID(Trim(CboGroupID))
       Else
            Exit Sub
       End If
 Else
       MsgBox "Can not close the wo,please check the excel file."
       Exit Sub
 End If
End Sub

Private Sub CmdExcel_Click()
Call Sap_Return(Trim(CboReportType))
End Sub

Private Function Sap_Return(ByVal Report_Type)
Dim str As String
Dim Rs As ADODB.Recordset
'CboReportType.AddItem "SAP1"
'CboReportType.AddItem "SAP2"
'CboReportType.AddItem "ReturnDID"
'CboReportType.AddItem "DispatchDID"
'CboReportType.AddItem "Return_Dispatch"
Select Case Report_Type
       Case "SAP1"
            str = "select * from QSMS_Sap where work_order='" & Trim(cboWO) & "' and status='open' order by UpCompPN,Item"
       Case "SAP2"
            str = "select * from QSMS_Sap where work_order='" & Trim(cboWO) & "' and status='close' order by UpCompPN,Item"
       Case "ReturnDID"
            str = "Select * from QSMS_GroupDID where GroupID='" & Trim(CboGroupID) & "'order by returnFlag,CompPN"
       Case "DispatchDID"
            str = "Select * from QSMS_Dispatch where Work_Order in (select Work_Order from QSMS_Wogroup where GroupID='" & Trim(CboGroupID) & "') order by Work_order,machine"
        '**Sandy        2008.02.11     add ReturnDIDByGroupID and ReturnDIDByWO (00006)
       Case "ReturnDIDByGroupID"
            str = "exec XL_ReturnDIDByGroupID '" & Trim(CboGroupID) & "'"
       Case "ReturnDIDByWO"
            str = "exec XL_ReturnDIDByWO '" & Trim(CboGroupID) & "'"
       Case "Return_Dispatch"
             str = "Select * from QSMS_GroupCompQty where GroupID='" & Trim(CboGroupID) & "' Order by CompPN"
       Case "CastQty"
             str = "exec QSMSGetCastQty '" & Trim(CboGroupID) & "'"
             Conn.Execute str
             str = "Select * from QSMS_GroupCompQty where GroupID='" & Trim(CboGroupID) & "' order by CompPN"
End Select
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   Call CopyToExcel(Rs)
Else
   MsgBox "No data"
End If
End Function

Private Sub cmdGetRefID_Click()
    Dim sCurrRefID As String
    Dim sMsg As String
    
    
'    TxtReturnQty.Locked = False ''1163
    sSql = "exec XL_DIDGetRefID @Type='Return', @IsGood=" & sq(IIf(optGoodMaterial.Value = True, "Y", "N")) & ",@UserName=" & sq(g_userName) & ",@Factory=" & sq(Trim(Factory)) & ",@IsAnotherBUDID=" & sq(Trim(IsAnotherBUDID))
    Set Rst = Conn.Execute(sSql)
    If Rst.EOF = False Then
        
        If Rst("Result") <> 0 Then
            MsgBox Rst("Description"), vbExclamation, "Prompt"
            Exit Sub
        End If
        
        sMsg = Trim(Rst("Description") & "")
        sCurrRefID = DIDGetRefIDByResult(sMsg)
        
        ''打印Label for GetRefID
        With DIDInfo
            .DID = sCurrRefID
            .COMPPN = sCurrRefID
            .Qty = -10000
            .IsGood = IIf(optGoodMaterial.Value = True, "Y", "N")
            .DIDType = ""
        End With
        
        If chkAutoGetRefID <> "1" Then
            Call DIDPrintLabel(OptZebra.Value, CInt(Trim(TxtCompPort)), Trim(TxtComm))
        End If
        
        ''Check Stock qty By RefID
        sSql = "exec XL_DIDChkStockByRefID @Type='Auto',@RefID=" & sq(sCurrRefID) & ",@UserName=" & sq(g_userName)
        Set Rst = Conn.Execute(sSql)
        If Rst.EOF = False Then
            If Rst("Result") <> 0 Then
                MsgBox Rst("Description"), vbExclamation, "Prompt"
                Exit Sub
            End If
            
            If chkAutoGetRefID <> "1" Then
                'to check the from:frmDIDChkStock exists or not (00007)
                Dim frm As Form
                For Each frm In Forms
                    If frm.Name = "frmDIDChkStock" Then
                        Unload frm
                        Exit For
                    End If
                Next frm
                'to check the from:frmDIDChkStock exists or not (00007)
                
                frmDIDChkStock.FuncType = "AutoChk"
                Set Rst = Rst.NextRecordset
                Set frmDIDChkStock.rstCompPN = Rst
                frmDIDChkStock.lblmsg = sMsg
                frmDIDChkStock.Show 1
            End If
            
        End If
    
    End If
End Sub

Private Sub CmdQuery_Click()
If Trim(CboLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If
Call GetGroupID
End Sub

Private Function GetLine()
Dim str As String
Dim Rs As ADODB.Recordset
str = "select distinct Line from QSMS_woGroup"
Set Rs = Conn.Execute(str)
CboLine.Clear
While Not Rs.EOF
    CboLine.AddItem Rs!Line
    Rs.MoveNext
Wend
End Function



Private Sub cmdReprint_Click()
     ''check printer
    Dim IsByDIDInput As Boolean
    
    IsByDIDInput = False
    
    If Trim(TxtCompPort) = "" Or Trim(TxtComm) = "" Then
        MsgBox "Printer have not set!!", vbInformation
        Exit Sub
    End If
    
'    If sstabDID.Tab <> 1 Then
'        MsgBox "Please select related CallBack DID to print!!", vbInformation
'        sstabDID.Tab = 1
'        Exit Sub
'    End If
    
    If Trim(TxtDID) = "" Then
        LblMessage = "Please select or Input DID to reprint!!"
        Exit Sub
    End If
    
    
    '''2008/03/24   Denver  Modify for Reprint DID --(0001)
    If gridDIDtoWH.row >= 0 Then
        If Trim(TxtDID) = Trim(gridDIDtoWH.Columns(1).text) Then
            With DIDInfo
                .DID = Trim(gridDIDtoWH.Columns(1).text)
                .COMPPN = Trim(gridDIDtoWH.Columns(2).text)
                .Qty = Trim(gridDIDtoWH.Columns(3).text)
                .IsGood = Trim(gridDIDtoWH.Columns(10).text)
                If ChkPrintDIDType = "Y" Then
                    .DIDType = Trim(gridDIDtoWH.Columns(14).text)
                Else
                    .DIDType = ""
                End If
            End With
        Else
            IsByDIDInput = True
        End If
            
    Else
        IsByDIDInput = True
    End If
    
    If IsByDIDInput = True Then
        sSql = "exec XL_DIDGetToWHInfo 'Return'," & sq(Trim(TxtDID)) & "," & sq(Trim(Factory)) & ",@IsAnotherBUDID=" & sq(Trim(IsAnotherBUDID))
        Set Rst = Conn.Execute(sSql)
        If Rst.EOF = True Then
            LblMessage = "There is no DID:" & sq(Trim(TxtDID)) + " !!"
            TxtDID.text = ""
            TxtReturnQty.text = ""
            TxtDID.SetFocus
            Exit Sub
        Else
            With DIDInfo
                .DID = Trim(Rst("DID") & "")
                .COMPPN = Trim(Rst("CompPN") & "")
                .Qty = Rst("Qty")
                .IsGood = Trim(Rst("IsGood") & "")
                .DateCode = Trim(Rst("DateCode") & "")
                .VendorCode = Trim(Rst("VendorCode") & "")
                .LotCode = Trim(Rst("VendorCode") & "")
                ''.Location = Trim(rst("Location") & "")  '1242
                
                If ChkPrintDIDType = "Y" Then
                    .DIDType = Trim(Rst("DIDType"))
                Else
                    .DIDType = ""
                End If
            End With
        End If
    End If
    
    Call DIDPrintLabel(OptZebra.Value, CInt(Trim(TxtCompPort)), Trim(TxtComm))
    
    TxtDID = ""
End Sub

Private Sub cmdReturn_Click()
Dim str As String
Dim Rs As ADODB.Recordset
Dim transdatetime As String
Dim tempNewDID As String, tempqty As String

If MsgBox("You have only once time to update the retrunQty, are you sure to update now? ", vbYesNo + vbQuestion) = vbNo Then
    Exit Sub
End If

If Trim(TxtReturnQty) = "" Or IsNumeric(TxtReturnQty) = False Then
   MsgBox "The Return Qty can not be empty or must be numeric"
   Exit Sub
End If
If CLng(TxtDIDTotalQty) < CLng(TxtReturnQty) Then
   str = "Select Qty from QSMS_DID where DID='" & Trim(TxtDID) & "'"
   Set Rs = Conn.Execute(str)
   If Rs!Qty < CLng(TxtReturnQty) Then
       MsgBox "Return Qty can not larger than total qty"
       Exit Sub
   End If
End If
cmdSave.Enabled = False

str = "exec QSMS_UpdateReturnQty '" & Trim(TxtDID) & "','" & CboGroupID & "'," & TxtReturnQty & ",'" & g_userName & "'"
Set Rs = Conn.Execute(str)
If Rs!result <> "PASS" Then
    MsgBox Rs!result
Else
    MsgBox "Update OK ! "
End If
cmdSave.Enabled = True
End Sub

Private Sub CmdReturnDIDAll_Click()
Dim str As String
Dim Rs As ADODB.Recordset
Dim transdatetime As String
str = "select getdate()"
Set Rs = Conn.Execute(str)
transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHNNSS")

DoEvents
DoEvents
DoEvents
DoEvents
If ChkErrAll = False Then
   Exit Sub
End If
LblALL.Caption = "System is Updating data,please wait"
str = "Update QSMS_GroupDID set ReturnFlag='Y',TransDateTime='" & transdatetime & "',UID='" & g_userName & "' where GroupID='" & CboGroupID & "' and returnFlag<>'Y'"
Conn.Execute (str)
Call CmdCloseWO_Click
LblALL.Caption = "Return all successfully"
End Sub
Private Function ChkErrAll() As Boolean
ChkErrAll = True
If Trim(CboGroupID) = "" Then
    ChkErrAll = False
End If
End Function

Private Sub cmdSave_Click()
On Error GoTo EcmdSave_Click
    Dim TransDate As String, strSQL As String
    Dim Rs As New ADODB.Recordset
    
    Dim sDID As String
    Dim intReturnQty As String    ''(1052)
    Dim sProcessStatus  As String
    Dim sNewDispDID As String
    Dim PreDID As String
    Dim RestQty As Integer
    
'    TxtReturnQty.Locked = False ''1163
    sProcessStatus = ""
    lblFeedBack = "Qty FeedBack:"
    If ChkErr = False Then
        GoTo Normal_Eixt
'       Exit Sub
    End If
    
    ''0040
    sDID = Trim(TxtDID)
    
    intReturnQty = Trim(TxtReturnQty)
    sSql = "exec XL_CheckReturnQty @DID=" & sq(Trim(sDID)) & ",@CompPN=" & sq(Trim(TxtCompPN)) & ", @ReturnQty=" & Trim(intReturnQty) & ",@GroupID=" & sq(PreGroupID) & ",@IsAnotherBUDID=" & sq(Trim(IsAnotherBUDID)) & ",@CheckForbiddenPN=" & sq(Trim(CheckReturnForbiddenPN))
    Set Rst = Conn.Execute(sSql)
    If UCase(Rst("Result")) = "F" Then '(0047)
        LblMessage.Caption = Rst("Description")
        MsgBox (Rst("Description"))
        GoTo Normal_Eixt
'        Exit Sub
    End If
    If Rst("Result") = 0 Then
        If MsgBox(Rst("Description"), vbYesNo) = vbNo Then
            GoTo Normal_Eixt
'            Exit Sub
        End If
        LblMessage.Caption = Rst("Description")
    End If
    
    If StrBU = "ESBU" And Rst("Result") = 3 Then           '''(1272)
        If MsgBox(Rst("Description"), vbYesNo) = vbNo Then
            GoTo Normal_Eixt
'            Exit Sub
        End If
    End If
    ''20081230   Denver  Check OK后，运用中间变量保存，以防运行中修改  sProcessStatus
    ''======20080422  Denver  need get QSMS_groupcompqty data====
    ''If UpdateReturnQty(Trim(CboGroupID), Trim(TxtCompPN), Trim(TxtDID), Trim(TxtReturnQty)) = False Then
    If UpdateReturnQty(Trim(Trim(PreGroupID)), Trim(TxtCompPN), Trim(sDID), Trim(intReturnQty)) = False Then
        GoTo Normal_Eixt
'        Exit Sub
    End If
        
    '**Denver       2010.03.19     Add IC comp check function  （0068）
    ''1267
    If BU = "ESBU" And IC_CompChk = "Y" Then
        If IC_CompNeedBurn(Trim(TxtCompPN)) = True Then
            Exit Sub
        End If
    End If
    ''1267

    ''2007.12.27 Denver  modify DID return Print label and get ReferenceID (0003)
    If PrtCallBKandReturn = "Y" Then
        sProcessStatus = "Return Start"
        sSql = "exec XL_DIDGetNewID @Type='Return',@DID=" & sq(Trim(sDID)) & ",@IsGood=" & sq(IIf(optGoodMaterial.Value = True, "Y", "N")) & ", @ReturnQty=" & Trim(intReturnQty) & ",@UserName=" & sq(g_userName) & ",@Factory=" & sq(Trim(Factory)) & ",@IsAnotherBUDID=" & sq(Trim(IsAnotherBUDID))
        Set Rst = Conn.Execute(sSql)
        If Rst("Result") <> 0 Then
            LblMessage.Caption = Rst("Description")
        Else
            LblMessage.BackColor = &H80FF80
            Set Rst = Rst.NextRecordset
            'PN/Qty/PU/NG/UID/Date   (bad)   'DID/Qty/PU/UID/Date (Good)
            If Rst.EOF = True Then
                LblMessage.Caption = "Get DID information fail,print DID fail!!"
                GoTo Normal_Eixt
            End If
            
            ''2008/04/02 denver    show return Qty and LineMC stock Qty   --(0008)
            lblFeedBack = Trim(Rst("QtyFeedback") & "")
            lblFeedBack = Mid(lblFeedBack, InStr(lblFeedBack, "##") + 2, 600)
            
            With DIDInfo
                .DID = Trim(Rst("DID") & "")
                .COMPPN = Trim(Rst("CompPN") & "")
                .Qty = Rst("Qty")
                .IsGood = Trim(Rst("IsGood") & "")
                .VendorCode = Trim(Rst("VendorCode"))
                .DateCode = Trim(Rst("DateCode"))
                .LotCode = Trim(Rst("LotCode"))
                If BU = "NB5" Then
                    .WareHouseID = Trim(Rst("WareHouseID"))        '(1252)
                End If
                If ChkPrintDIDType = "Y" Then  ''1142
                    .DIDType = Trim(Rst("DIDType"))
                Else
                    .DIDType = ""
                End If
            End With
            
            If DIDInfo.IsGood = "Y" Then
                Set Rst = Conn.Execute("select getdate()")
                TransDate = Format(Rst(0), "yyyymmddhhnnss")
                TempDID = Trim(GetDID(DIDInfo.COMPPN, TransDate))
                
                sProcessStatus = "Dispatch_Start"
                ''''''Changed from '0' to '4' for Type  (00004)'''''''' Add Factory''''(0010)
                ''(1160) add a new parameter @OldDID for return 先满足原有线别需求
                
                'Add superchai 20240125
                strSQL = "exec XL_DIDAutoDispatch_superchai " & _
                        "'" & TempDID & "'," & _
                        "'" & DIDInfo.COMPPN & "'," & _
                        "" & DIDInfo.Qty & "," & _
                        "" & DIDInfo.Qty & "," & _
                        "'" & DIDInfo.VendorCode & "'," & _
                        "'" & DIDInfo.DateCode & "'," & _
                        "'" & DIDInfo.LotCode & "'," & _
                        "''," & _
                        "''," & _
                        "'" & g_userName & "'," & _
                        "'4'," & _
                        "'', " & _
                        "'', " & _
                        "'', " & _
                        "'', " & _
                        "'', " & _
                        "'', " & _
                        "'', " & _
                        "'', " & _
                        "'0', " & _
                        "'" & Trim(Factory) & "'," & _
                        "'', " & _
                        "'" & Trim(sDID) & "'"
                Set Rst = Conn.Execute(strSQL)
                
                sProcessStatus = "Dispatch_End"
                If Rst.EOF = False Then
                    
                    LblMessage = Rst("ErrDesc")
                    If Rst("result") <> 1 Then
                        '--(0011)
                        ''Conn.Execute ("insert into qms_log(system_name,event_no,sn,user_name,desc1,trans_date) values ('SMT_QSMS','1','" & Trim(TxtDID) & "','" & g_userName & "',N'" & Trim(LblMessage) & "',DBO.FORMATDATE(GETDATE(),'YYYYMMDDHHNNSS'))")
                        sSql = "insert  into QSMS_Error_Log(AppName,SubFunction,SubID,DetailDesc,Col2,Col3,TransDateTime)  " _
                            & " values('QSMS_Return','ByDID_Dispatch','Log',N'DID=" + Trim(sDID) + "; Message:" + Trim(LblMessage) + "' ,app_name(),substring(host_name(),1,20), DBO.FORMATDATE(GETDATE(),'YYYYMMDDHHNNSS')  )"
                        
                        Conn.Execute sSql
                        GoTo PrintLabel
                    Else

DoPrintLabel:
                        '*Denver 2009.08.16 NB2&NB3分家，但是用同一个QSMS可作NB2NB3的退料
                        '20090205  Denver  if ChkDispatchIsOK, it need Print Label
                        sProcessStatus = "Del_ToWH_DID"
                        'delete the data of QSMS_DID_ToWH (00005)
                        'Conn.Execute ("delete QSMS_DID_ToWH where oldDID='" & Trim(sDID) & "' and ToWHType='Return' and IsGood='Y' and UID='" & Trim(g_userName) & "'")
                        'Set rs = Conn.Execute("exec [XL_GetDidPrintInfo] @DID='" & Trim(TempDID) & "'")

                        sNewDispDID = Trim(Rst("DID") & "")
                        TempDID = sNewDispDID
                        sSql = " exec XL_GetDidPrintInfo_Return @DID=" & sq(sNewDispDID) & ",@OldDID=" & sq(sDID) & ",@IsAnotherBU='N',@Factory='',@PrinterType='" & Trim(PrinterType) & "',@PrintDpm='" & Trim(PrintDpm) & "'"
                        Set Rst = Conn.Execute(sSql)
                        
                        If Rst.EOF = True Then
                            LblMessage = "Can not get auto dispatch did"
                            GoTo PrintLabel
                        Else
                            If Rst("Result") <> 0 Then
                                LblMessage = Trim(Rst("Description") & "")
                                GoTo PrintLabel
                            Else
                                Set Rst = Rst.NextRecordset
                                PrintData.Line = Trim(Rst!Line)
                                PrintData.Machine = Trim(Rst!FirstMachine)
                                PrintData.Side = Trim(Rst!Side)
                                PrintData.DIDWOGROUP = Trim(Rst!woGroup)
                                WOType = Trim(Rst!WOType)
                                PrintData.BU = Trim(Rst!DIDHead)   ''print DID head
                                PrintData.location = Trim(Rst!location) '1242
                                PrintData.Mark = Trim(Rst!Mark) '1255
                                DIDInfo.location = Trim(Rst!location) '1242
                                DIDInfo.Mark = Trim(Rst!Mark) '1255
                                
                                If opNewLabel.Value = True Then
                                    Dim x As Integer
                                    For x = 0 To 4
                                        WO(x) = ""
                                        Model(x) = ""
                                        Machine(x) = ""
                                        Work_Order(x) = ""
                                        DIDType(x) = ""
                                        ISCYL(x) = ""
                                        Slot(x) = ""
                                        VenderCode(x) = ""
                                        LR(x) = ""
                                    Next x
                            
                                    Set Rst = Rst.NextRecordset
                                    If Rst.EOF = False Then
                                        Dim i As Integer, j As Integer, ff As Integer
                                        j = Rst.RecordCount
                                        If j > 5 Then j = 5
                                        For i = 0 To j - 1
                                            WO(i) = Rst("Machine") + " " + Rst("Slot") + "-" + Rst("LR") '0043
                                            Model(i) = Rst("model")
                                            Machine(i) = Mid(Rst("Machine"), 2, 1) + "-" + Mid(Rst("Slot"), 1, 1) + "-" + Mid(Rst("Machine"), 6, 1)
                                            Work_Order(i) = Rst("Work_Order")       '''(1093)
                                            DIDType(i) = Rst("DIDType")
                                            
                                            MachineCH(i) = Rst("MachineCH") ''1247
                                            SideCH(i) = Rst("SideCH")       ''1247
                                            LRCH(i) = Rst("LRCH")           ''1247
                                            SlotCH(i) = Rst("Slot")         ''1247
                                            PN(i) = Rst("PN")               ''1247
                                            If BU = "NB5" Then
                                              ReelWidth(i) = Rst("ReelWidth")
                                            End If
                                            
                                            If PrintedSeqID = "Y" Then
                                                SeqID(i) = Rst("SeqID")           '(1148)
                                            End If
                                            If PrintedVenderCode = "Y" Then       ''1227
                                                VenderCode(i) = Rst("VenderCode")
                                                LR(i) = Rst("SLR")
                                            End If
                                            For ff = 0 To Rst.Fields.Count - 1    ''(1109)
                                                If UCase(Rst.Fields(ff).Name) = "ISCYL" Then
                                                    ISCYL(i) = Rst("ISCYL")
                                                End If
                                            Next ff
                                            Rst.MoveNext
                                        Next i
                                    End If
                                End If
                                
                                Call PrintAutoDispatchLabel
    '                            lelmessage = "Return DID auto dispatch successful!"
                                'there is no this item
                                LblMessage = "Return DID auto dispatch successful!"
                                
                                GoTo Normal_Eixt
                            
                            End If
                        End If
                    End If
                Else
                    LblMessage = "Auto dispatch return DID fail!"
                    GoTo Normal_Eixt
                End If
            End If
PrintLabel:
            sProcessStatus = "Print_Label"
            Call DIDPrintLabel(OptZebra.Value, CInt(Trim(TxtCompPort)), Trim(TxtComm))
            If chkAutoGetRefID.Value = "1" Then
                Call cmdGetRefID_Click
            End If
        End If
    End If
    
    If IsAnotherBUDID <> "Y" Then
        Call GetReturned_NotReturnDID(Trim(CboGroupID))
        Call GetDIDInfo(Trim(sDID), Trim(CboGroupID))
    End If
    
Normal_Eixt:
    TxtDID.text = ""
    TxtReturnQty.text = ""
    TxtDID.SetFocus
    
    Exit Sub
EcmdSave_Click:

    ''20081230   Denver   Check if DispatchIsOK
    Dim sErrMsg As String
    sErrMsg = "ErrNum: " & Err.Number & ",ErrDesc: " & Err.Description
    
    Call InsertIntoQSMSLog("QSMS", "ReturnDID", "DID:" & TempDID & ";SDID:" & sDID & ";ProcessStatus:" & sProcessStatus & ";ErrDetail:" & sErrMsg)  '(0042)
    If UCase(sProcessStatus) = UCase("Dispatch_Start") Or UCase(sProcessStatus) = UCase("Del_ToWH_DID") Then
       If ChkDispatchIsOK(Trim(TempDID), sDID, sErrMsg) = True Then
           GoTo DoPrintLabel
       End If
    Else
        MsgBox sErrMsg + ",Please contact QSMS SMT Staff " + sProcessStatus
    End If
     
End Sub
''1267
Private Function IC_CompNeedBurn(COMPPN As String) As Boolean
On Error GoTo errHandler:
    
    Dim Rs As New ADODB.Recordset
    Dim NeedBurn As Boolean
    
    NeedBurn = False
    
    strSQL = "exec IC_CompNeedBurn " & sq(COMPPN)
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = False Then
        If Rs("Result") = 0 Then
            If MsgBox(Rs("Description") & " DO you burn IC for it firstly!", vbYesNo) = vbYes Then
                NeedBurn = True
            End If
       
        End If
    End If
    
    IC_CompNeedBurn = NeedBurn
    Exit Function
    
errHandler:

    IC_CompNeedBurn = False
    MsgBox Err.Number & "," & Err.Description
End Function
''1267


Private Sub DGDIDInfo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim str As String
On Error Resume Next
    With DGDIDInfo
        If Not IsNumeric(LastRow) Then Exit Sub
        If .row <= 0 Then Exit Sub
    
         TxtDID = Trim(.Columns(0).text)
         TxtDIDTotalQty.text = Trim(.Columns(1).text)
         TxtDIDReturnedQty.text = Trim(.Columns(2).text)
         TxtCompPN = Mid(Trim(.Columns(0).text), 1, 11)
         
         If Err.Number <> 0 Then
            TxtDIDTotalQty.text = vbNullString
            TxtDIDReturnedQty.text = vbNullString
         End If
         
      
    End With

End Sub

Private Sub DGDIDNeedReturned_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    With DGDIDNeedReturned
        If Not IsNumeric(LastRow) Then Exit Sub
        If .row <= 0 Then Exit Sub
        
        TxtChkDID = Trim(.Columns(0).text)
        
        Call TxtChkDID_KeyPress(13)
    End With
End Sub


Private Sub DGDIDReturned_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    With DGDIDReturned
        If Not IsNumeric(LastRow) Then Exit Sub
        If .row <= 0 Then Exit Sub
        
        TxtChkDID = Trim(.Columns(0).text)
        
        Call TxtChkDID_KeyPress(13)
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
    CboReportType.AddItem "SAP1"
    CboReportType.AddItem "SAP2"
    CboReportType.AddItem "ReturnDID"
    CboReportType.AddItem "DispatchDID"
    CboReportType.AddItem "Return_Dispatch"
    CboReportType.AddItem "ReturnDIDByGroupID"
    CboReportType.AddItem "ReturnDIDByWO"
    CboReportType.AddItem "CastQty"
'    If returnDIDflag = True Then
'        cmdReturn.Visible = True
'    End If
    
'    TxtCompPort = GetSetting("SMT", "QSMS", "CommPort", "1")
'    TxtComm = GetSetting("SMT", "QSMS", "Comm", "9600,N,8,1")

    '20101115 Maggie Save Printer setting in local Registry (1019)
    Call GetPrinterSetting(FrmReturnDID)
    
    ''(1080)
    If OptZebra.Value = True Then
        isZebra = True
    Else
        isZebra = False
    End If

    sstabDID.Tab = 0
    
     ''20100507    Denver     好坏料不让user 选择
    optGoodMaterial.Enabled = False
    optBadMaterial.Enabled = False
      
    strCheckScaner = ReadIniFile("QSMS", "DIDScan", App.path & "\set.ini")
    If PrtCallBKandReturn <> "Y" Then
        cmdGetRefID.Visible = False
        cmdReprint.Visible = False
        FraPrinter.Visible = False
    End If
'    MSComm1.CommPort = GetSetting("SMT", "QSMS", "CommPort") ''(1163)
'    MSComm1.Settings = "9600,n,8,1"
'    MSComm1.InputMode = comInputModeText
'    MSComm1.DTREnable = True
'    MSComm1.Handshaking = comRTS
'    MSComm1.InputLen = 0
'    MSComm1.PortOpen = True
End Sub

Private Function DeleteDID(ByVal GroupID As String)
Dim str As String
Dim Rs As ADODB.Recordset
'(1)backup qsms_dispatch
   str = "Insert into QSMS_Dispatch_bak(Work_Order,Line,WoQty,JobPN ,Machine,CompPN ,Slot,BaseQty,NeedQty, DID,DIDQty ,VendorCode,DateCode,LotCode ,UID,TransDateTime,DeletedFlag ) " & _
      "select Work_Order,Line,WoQty,JobPN ,Machine,CompPN ,Slot,BaseQty,NeedQty, DID,DIDQty ,VendorCode,DateCode,LotCode ,UID,TransDateTime,DeletedFlag  from " & _
      "QSMS_Dispatch where work_Order in (select work_Order from QSMS_WOGroup where GroupID='" & GroupID & "') "
  Conn.Execute str
'(2)Delete qsms_dispatch
str = "delete QSMS_Dispatch  from QSMS_Dispatch Where work_Order in (select work_Order from QSMS_WOGroup where GroupID='" & GroupID & "') " 'and (b.Returnflag<>'Y' or b.ReturnQty=0)"
Conn.Execute str
'(3)Backup QSMS_WO
str = "Insert into QSMS_Wo_Bak (Work_Order,Line,WoQty, JobPN,JobGroup, Machine,CompPN , Slot, LR,Item ,BaseQty,NeedQty, DispatchQty, BalanceQty,  MachineFinishedFlag ,WoFinishedFlag ,RefreshFlag,BuildType,Side )" & _
    "select Work_Order,Line,WoQty, JobPN,JobGroup, Machine,CompPN , Slot, LR,Item ,BaseQty,NeedQty, DispatchQty, BalanceQty,  MachineFinishedFlag ,WoFinishedFlag ,RefreshFlag,BuildType,Side from QSMS_WO " & _
    " where work_order in (select work_order from QSMS_WoGroup where GroupID='" & GroupID & "')"
Conn.Execute str
'(4)Delete from QSMS_WO
str = "Delete from QSMS_Wo where work_Order in (select work_Order from QSMS_WOGroup where GroupID='" & GroupID & "')"
Conn.Execute str
 '(5) backup DID
' str = "Insert into QSMS_DID_Bak(DID,CompPN,Qty,RemainQty,RealQty,VendorCode,DateCode,LotCode,DIDLoc,UID,TransDateTime,UsedFlag,InheritFlag) " & _
'       "select a.DID,a.CompPN,a.Qty,a.RemainQty,a.RealQty,a.VendorCode,a.DateCode,a.LotCode,a.DIDLoc,a.UID,a.TransDateTime,a.UsedFlag,a.InheritFlag from  " & _
'       "QSMS_DID a,QSMS_GroupDID b where  B.GroupID='" & Trim(GroupDID) & "' and a.DID= b.NewDID"
' Conn.Execute str
''(6)Delete from QSMS_DID
'
'str = "Delete   QSMS_DID  from QSMS_DID a,QSMS_GroupDID b where  B.GroupID='" & Trim(GroupDID) & "' and a.DID= b.NewDID " 'and (b.Returnflag<>'Y' or b.ReturnQty=0)"
'Conn.Execute str
'


End Function
Private Function GetGroupID()
Dim str As String
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

If BU = "NB5" Then
    If OptRelease.Value = True Then
       str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag<>'Y' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )"
    Else
        str = "select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag<>'Y' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )"
    End If
Else
    If OptRelease.Value = True Then
       str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag<>'Y'"
    Else
        str = "select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag<>'Y'"
    End If
End If
Set Rs = Conn.Execute(str)
CboGroupID.Clear
If Rs.EOF Then MsgBox "No data"
While Not Rs.EOF
      CboGroupID.AddItem Trim(Rs!GroupID)
      Rs.MoveNext
Wend
End Function

Private Function GetGroupWO(ByVal GroupID As String)
Dim str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset

str = "select Work_Order,ClosedFlag from QSMS_WOGroup  where GroupID= '" & GroupID & "'"

Set Rs = Conn.Execute(str)

cboWO.Clear
While Not Rs.EOF
     
          cboWO.AddItem Trim(Rs!Work_Order)
      
      Rs.MoveNext
Wend
End Function


Private Function GetWoinfo(ByVal WO As String)
    Dim str As String
    Dim Rs As ADODB.Recordset
    str = "select PN, Qty from Sap_Wo_List where WO='" & Trim(WO) & "'"
    Set Rs = Conn.Execute(str)
    If Not Rs.EOF Then
       TxtMBPN = Rs!PN
       TxtWOQty = Rs!Qty
    End If
    str = "select Customer from ModelName where PN='" & TxtMBPN & "'"
    Set Rs = Conn.Execute(str)
    If Not Rs.EOF Then
       TxtCustomer = Trim(Rs!Customer)
    End If
    
End Function

Private Function GetReturned_NotReturnDID(ByVal GroupID As String)
    Dim str As String
    Dim Rs As ADODB.Recordset
    '(1) get Needn't return DID---has been consumed
    str = "select Distinct DID from QSMS_GroupDID where GroupID='" & GroupID & "' and realQty=0 Order by DID"
    Set Rs = Conn.Execute(str)
    
    '(2) get return ComppN
    str = "select distinct DID from QSMS_GroupDID where GroupID='" & GroupID & "' and ReturnFlag='Y' and realQty<>0 order by DID"
    Set Rs = Conn.Execute(str)
    Set DGDIDReturned.DataSource = Rs
    DGDIDReturned.Caption = "(Returned DID)  Total: " & Rs.RecordCount

    ''2007.12.27 Denver  modify Return Print label and get ReferenceID (0003)
    ''DID to WareHouse top 1000
    If PrtCallBKandReturn = "Y" Then
        sSql = "exec XL_DIDGetToWHInfo 'Return',''," & sq(Trim(Factory))
        If rstDIDtoWH.State Then rstDIDtoWH.Close
        rstDIDtoWH.CursorLocation = adUseClient
        Set rstDIDtoWH = Conn.Execute(sSql)
        Set gridDIDtoWH.DataSource = rstDIDtoWH
        gridDIDtoWH.Refresh
    End If
End Function

Private Function ChkReturnFinished(ByVal GroupID As String)
Dim str As String
Dim Rs As ADODB.Recordset
Dim TotalCompQty, ReturnCompQty As Long
'
'Str = "select count(distinct CompPN) from QSMS_GroupDID where GroupID='" & GroupID & "'"
'Set Rs = Conn.Execute(Str)
'TotalCompQty = Rs.Fields(0)
'
'Str = "select count(distinct CompPN) from QSMS_GroupDID where GroupID='" & GroupID & "' and Returnflag='Y'"
'Set Rs = Conn.Execute(Str)
'ReturnCompQty = Rs.Fields(0)

str = "select Distinct DID from QSMS_GroupDID where GroupID='" & GroupID & "' and ReturnFlag<>'Y' and realQty<>0 Order by DID"
Set Rs = Conn.Execute(str)

'If TotalCompQty = ReturnCompQty Then
If Rs.EOF Then
    If ChkCLoseWoByAuto(GroupID) = True Then
'       If UpdateQSMSGroupCompQty(GroupID) = True Then
          Call DeleteDID(Trim(GroupID))
'       Else
'          Exit Function
'       End If
    Else
       MsgBox "Can not close the wo,please check the excel file."
       Exit Function
    End If
End If


End Function

Private Function ChkCLoseWoByAuto(ByVal GroupID As String) As Boolean
Dim str As String
Dim Rs As ADODB.Recordset
Dim RsWo As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
ChkCLoseWoByAuto = True
str = "select distinct b.[Group] from QSMS_WoGroup a,Sap_Wo_List b where a.GroupID='" & GroupID & "' and a.work_Order=b.wo and a.closedflag<>'Y'"
Set Rs = Conn.Execute(str)
While Not Rs.EOF
      str = "select min(wo) as WO from sap_wo_list where [group]='" & Rs![Group] & "'"
      Set RsWo = Conn.Execute(str)
      If Not RsWo.EOF Then
            If CloseWoByManual(Trim(RsWo!WO), "Auto") = True Then
            Else
               ChkCLoseWoByAuto = False
              
            End If
      End If
      Rs.MoveNext
      DoEvents
      DoEvents
      DoEvents
      DoEvents
Wend

End Function



Private Sub gridDIDtoWH_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     With gridDIDtoWH
        If Not IsNumeric(LastRow) Then Exit Sub
        If .row <= 0 Then Exit Sub
        TxtDID = Trim(.Columns(1).text)
        TxtDID.SetFocus
        
      End With
End Sub

'Private Sub MSComm1_OnComm()
'Dim Instring As String, NewComp() As String
'    Call Delay_Time(0.5)
'    Do While MSComm1.InBufferCount <> 0
'        Instring = Instring & MSComm1.Input
'        DoEvents
'    Loop
'    NewComp = Split(Trim(Instring), """")
'    If UBound(NewComp) > 0 Then
'        If TxtDID.Text = "" Then
'            MsgBox "请先刷入料号！！"
'        Else
'            TxtReturnQty.Text = Trim(NewComp(1))
'            TxtReturnQty.Locked = True
'        End If
'    End If
'End Sub

Private Sub TxtChkDID_KeyPress(KeyAscii As Integer)
Dim str As String
Dim Rs As ADODB.Recordset
str = "select A.DID,A.Qty as TotalQty,B.ReturnQty ,A.RealQty from QSMS_DID A left join QSMS_GroupDID B on  A.DID=B.DID where A.DID='" & Trim(TxtChkDID) & "'"
Set Rs = Conn.Execute(str)
If Rs.EOF Then
  LblChk.Caption = "The CompPN does not belong to the GroupID"
Else
   Set DGDIDInfo.DataSource = Rs
   
End If

End Sub

Private Sub txtDID_Click()
Sendkeys "{HOME}+{END}"
End Sub

Private Sub TxtDID_KeyDown(KeyCode As Integer, Shift As Integer)

If strKeyInPNByManual = True Then  '1071
    strCheckScaner = "N"
End If
If strCheckScaner = "Y" Then    '1071
    If App.Title <> App.EXEName Then
        If Shift = 2 Then
            MsgBox "Can't use Ctrl+V and Ctrl+C,input is void!", vbCritical
            TxtDID.text = ""
        End If
    End If
End If
End Sub

Private Sub txtDID_KeyPress(KeyAscii As Integer)
 
Dim strSQL, strComPN As String    '20230508 Kaelyn 穝糤DID跑计      'superchai add dataType 20230516
Dim Rs As ADODB.Recordset
Dim PreDID As String, RestQty As Long, TotalQty As Long
    
If strKeyInPNByManual = True Then            '1071
    strCheckScaner = "N"
End If
If strCheckScaner = "Y" Then          '1071
    If Len(Trim(TxtDID.text)) < 1 Then strDelaytime = 0
        If strDelaytime <> 0 Then
            If GetTickCount - strDelaytime > 100 Then
                MsgBox "Please use scaner!"
                TxtDID.text = ""
                strDelaytime = 0
                Exit Sub
            End If
        End If
    strDelaytime = GetTickCount
End If

If strCheckScaner = "Y" And KeyAscii = 13 Then   '1071
    strDelaytime = 0
End If

If (KeyAscii = 13 Or KeyAscii = 9) And Trim(TxtDID) <> "" Then
    
     If ChkFujiSPL = "Y" Then  ''（1107）
'        strsql = "select DID,RealQty from QSMS_DID WHERE DID = (select DID from QSMS_DID where NextDID='" & Trim(TxtDID) & "') AND RealQty>0"
        strSQL = "select DID,RealQty from QSMS_DID WHERE DID = (select top 1 DID from QSMS_DID where NextDID='" & Trim(TxtDID) & "' order by  SplicingDT desc) AND RealQty>0"  ''(1113)
        Set Rs = Conn.Execute(strSQL)
        If Rs.EOF = False Then
            PreDID = Rs!DID
            RestQty = Rs!realqty
            strSQL = "select DID,RealQty from QSMS_DID WHERE DID = '" & Trim(TxtDID) & "'"
            Set Rs = Conn.Execute(strSQL)
            MsgBox "This DID has received material and has not been used up." + vbCrLf + "Previous DID:" + PreDID + "; Qty:" + CStr(RestQty) + "." + vbCrLf + "Next DID：" + Rs!DID + "; Qty:" + CStr(Rs!realqty) + "."
        End If
    End If
    '1162
    strSQL = "select * from CompPN_Data WHERE type='Sole' and CompPN='" & Trim(TxtCompPN) & "'"
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = False Then
        MsgBox "The material of this DID is for exclusive use. Please confirm whether the return quantity is consistent with the actual quantity."
    End If
    
    '--superchai comment (Begin) QMB0001 20230516--
    'If BU = "NB5" Then
    '    strSQL = "SELECT 1 FROM  msd_data where CompPN=Left('" & Trim(TxtDID) & "',11)" '1248
    '    Set Rs = Conn.Execute(strSQL)
    '    If Not Rs.EOF Then
    '            MsgBox "This material needs to be vacuum packed and then returned to the warehouse."
    '    End If      '1248
    'End If
    '--superchai comment (End) QMB0001 20230516--
    
    '--superchai add function from QSMC (Begin) QMB0001 20230516--
    If BU = "NB5" Then
        '20230508 Kaelyn forPN э20絏
        strSQL = "SELECT CompPN FROM QSMS_DID with(nolock) where DID='" & Trim(TxtDID) & "'"
        Set Rs = Conn.Execute(strSQL)
        strComPN = CStr(Rs!COMPPN)
        strSQL = "SELECT 1 FROM  msd_data where CompPN='" & strComPN & "'"
        
        'strSQL = "SELECT 1 FROM  msd_data where CompPN=Left('" & Trim(TxtDID) & "',11)" '1248
        Set Rs = Conn.Execute(strSQL)
        If Not Rs.EOF Then
                MsgBox "This material needs to be vacuum packed and then returned to the warehouse."
        End If      '1248
    End If
    '--superchai add function from QSMC (End) QMB0001 20230516--
    
    ''(1116)(1240)
    strSQL = "Exec GetDIDRealQty '" & Trim(TxtDID.text) & "'"
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = False Then
        TxtReturnQty.text = Rs!realqty
        TotalQty = Rs!TotalQty
        
        If TotalQty = TxtReturnQty Then    ''''1197
          MsgBox "This DID has not been used yet, please double check whether it is true and requires a Return."
        End If
    End If
   ''1221
    If ChkHUA.Value = 1 Then ''1280
       strSQL = "EXEC XL_DIDReturnCheck @DID='" & Trim(TxtDID) & "',@ChkPN='Y'"
        Set Rs = Conn.Execute(strSQL)
        If Rs.EOF = False Then
           If Rs!result = 1 Then
               MsgBox Rs!Message
               TxtDID.text = ""
               Exit Sub
           End If
         End If
    Else
        strSQL = "EXEC XL_DIDReturnCheck @DID='" & Trim(TxtDID) & "'"
        Set Rs = Conn.Execute(strSQL)
        If Rs.EOF = False Then
           If Rs!result = 1 Then
               MsgBox Rs!Message
               TxtDID.text = ""
               Exit Sub
           End If
         End If
    End If
    
    TxtReturnQty.SetFocus
    Call TxtReturnQty_Click
'    TxtReturnQty.Text = ""
    
    LblMessage.BackColor = &H80000000
    LblMessage.Caption = ""
 
End If


End Sub

Private Function GetDIDInfo(ByVal DID As String, ByVal GroupID As String)
Dim str As String
Dim Rs As ADODB.Recordset
'Str = "Select Qty,RemainQty from QSMS_DID where DID='" & DID & "' "
'Set Rs = Conn.Execute(Str)
'If Not Rs.EOF Then
'   TxtDIDTotalQty = Trim(Rs!Qty)
'   TxtDIDRemainQty = Trim(Rs!RemainQty)
'End If
If Trim(TxtDID) = "" Then Exit Function '(00002)
str = "Select ReturnQty,TotalQty,CompPN from QSMS_GroupDID where DID='" & DID & "' and GroupID like'" & GroupID & "'+'%'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   TxtDIDTotalQty = Trim(Rs!TotalQty)
   TxtDIDReturnedQty = Trim(Rs!ReturnQty)
   TxtCompPN = Trim(Rs!COMPPN)
Else
    '**Denver       2008.04.07     it need not select GroupID and WO when Return DID   --(0009)
    str = "Select 0 as ReturnQty,Qty as TotalQty,CompPN from QSMS_DID where DID=" & sq(DID)
    Set Rs = Conn.Execute(str)
    If Not Rs.EOF Then
        TxtDIDTotalQty = Trim(Rs!TotalQty)
        TxtDIDReturnedQty = Trim(Rs!ReturnQty)
        TxtCompPN = Trim(Rs!COMPPN)
    Else
        MsgBox "DID:" & DID & " is not existed in QSMS_DID!!", vbExclamation, "Prompt"
        Exit Function
    End If
End If

str = "Select * from QSMS_GroupCompQty where GroupID='" & GroupID & "' and CompPN='" & Trim(TxtCompPN) & "'"
Set Rs = Conn.Execute(str)
Set DGCompInfo.DataSource = Rs
DGCompInfo.Refresh
End Function

'Private Function SaveReturnDID(ByVal DID As String, ByVal GroupID As String)
'Dim Str As String
'Dim Rs As ADODB.Recordset
'Dim Qty As Long
'Qty = CLng(TxtReturnQty)
'
'
'Str = "Update QSMS_GroupCompQty set ReturnQty=ReturnQty+'" & Qty & "',ReturnFlag='Y' where GroupID='" & GroupID & "' and DID='" & DID & "'"
'Conn.Execute (Str)
'
'
'Str = "Update QSMS_GroupCompQty set ConsumedQty=DispatchedQty-ReturnQty,AutoQty=dispatchedQty-ReturnQty-ManualQty where GroupID='" & GroupID & "' and DID='" & DID & "'"
'Conn.Execute (Str)
'
'End Function

Private Function ChkDIDBelongToGroupID(ByVal GroupID As String, ByVal DID As String) As Boolean
    Dim str As String
    Dim ReturnTime As String
    Dim Rs As ADODB.Recordset
    
    ChkDIDBelongToGroupID = True
    If Trim(TxtDID) = "" Then Exit Function '(00002)

    '**Denver       2009.03.03     如果有选择GroupID,就只需要检查属于GroupID，  不选不检查   --(0041)
    If Trim(GroupID) = "" Then
        ChkDIDBelongToGroupID = True
        Exit Function
    End If
    
    str = "Select top 1 DID,ReturnFlag,RealQty,transdatetime,returnTimes from QSMS_GroupDID where GroupID='" & Trim(GroupID) & "' and DID='" & Trim(DID) & "' order by transdatetime desc"
    
    Set Rs = Conn.Execute(str)
    If Rs.EOF Then
       ChkDIDBelongToGroupID = False
       MsgBox "The DID does not belong to the GroupID,Please check"
       
       
'    Else

'         If Trim(rs!ReturnFlag) = "Y" Then
'            ReturnTime = rs!TransDateTime
'         End If
'
'         If Trim(rs!ReturnFlag) = "Y" And NoExistsSecondDispatch(ReturnTime, Trim(DID)) = True Then  ''''(00001)
'            ChkDIDBelongToGroupID = False
'            MsgBox "The DID has been returned,Please check"
'         End If
'       'End If
'       If rs!realqty = 0 Then
'           ChkDIDBelongToGroupID = False
'           MsgBox "The DID RemainQty is zero,Needn't return"
'       End If
    End If

End Function

Private Function ChkErr() As Boolean
    Dim str As String
    Dim Rs As ADODB.Recordset
    Dim BalanceQty As Long
    
    
    ChkErr = True
    If Trim(TxtDID) = "" Then
        ChkErr = False
        Exit Function
    End If
    
    ''--*2009.07.24    Denver    NB2&NB3分家，但是用同一个QSMS可作NB2&NB3 的退料,返回True
    If IsAnotherBUDID = "Y" Then
        Exit Function
    End If
 
 
    '**Denver       2008.04.07     it need not select GroupID and WO when Return DID   --(0009)
    '**Denver       2009.03.03     如果有选择GroupID,就需要检查属于GroupID，  不选不检查   --(0041)
    If ChkDIDBelongToGroupID(Trim(CboGroupID), Trim(TxtDID)) = False Then
       ChkErr = False
       Exit Function
    End If

    If Trim(TxtReturnQty) = "" Or IsNumeric(TxtReturnQty) = False Then
       MsgBox "The Return Qty can not be empty or must be numeric"
       ChkErr = False
       Exit Function
    End If
    
    TxtReturnQty = Abs(Int(Trim(TxtReturnQty)))
    
    ''20081230  Denver   其实此处检查数量不为0
    If Trim(TxtReturnQty) <= 0 Then
        MsgBox "The Return Qty must be >0 !!"
        ChkErr = False
        Exit Function
    End If
   
   
    ''20080425  Denver  因为接料(A==>B) ,B 的Remain数量会〉总数，故放量30%(Move to procedure)--(0012)
'    If CLng(TxtDIDTotalQty) < CLng(TxtReturnQty) Then
'       Str = "Select Qty from QSMS_DID where DID='" & Trim(TxtDID) & "'"
'       Set Rs = Conn.Execute(Str)
'       If Rs!Qty < CLng(TxtReturnQty) Then
'           MsgBox "Return Qty can not larger than total qty"
'           ChkErr = False
'       End If
'
'
'    End If

    
    If ChkDIDInMachine(Trim(TxtDID)) = False Then
         ChkErr = False
         Exit Function
     End If
     
     '**Denver       2008.04.07     it need not select GroupID and WO when Return DID   --(0009)
'    If ChkGroupClosed(Trim(CboGroupID)) = True Then
'       ChkErr = False
'        MsgBox "The Group has been closed,can not return DID"
'    End If
    
    ''check printer
    If Trim(TxtCompPort) = "" Or Trim(TxtComm) = "" Then
        MsgBox "Printer have not set!!", vbInformation
        Exit Function
    End If
    
     ''1200
   If CheckMSDCallBack = "Y" Then
        strSQL = "select * from MSD_Data WHERE CompPN='" & Trim(TxtCompPN) & "'"
        Set Rs = Conn.Execute(strSQL)
        If Rs.EOF = False Then
            LblMessage.Caption = "This is MSD Material! "
            strSQL = "exec [PD_MSD_LinkDIDAuto] @DID='',@ReturnDID='" & Trim(TxtDID.text) & "',@CompPN='" & Trim(TxtCompPN) & "',@Inherit_WO='',@ReturnFlag='Y',@UID='" & g_userName & "' "
                Set Rs = Conn.Execute(strSQL)
                If UCase(Rs!result) = "CHECKFAIL" Then ''1200
                    MsgBox ("Message: " & Rs!ErrDesc), vbCritical
                    ChkErr = False
                    Exit Function
                End If
        End If
   End If
    
    

End Function

Private Function UpdateReturnQty(ByVal GroupID As String, COMPPN As String, DID As String, ReturnQty As Long) As Boolean
    Dim str As String
    Dim Rs As ADODB.Recordset
    Dim transdatetime As String
    Dim ReturnDIDSeq As String
    Dim intPos As String
    
    UpdateReturnQty = False
    
    str = "select GetDate()"
    Set Rs = Conn.Execute(str)
'    transdatetime = Format(rs.Fields(0), "YYMMDDHHNNSS")
    transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHNNSS")  ''(1111)
    
    ReturnDIDSeq = COMPPN + "-A" + transdatetime
    
    ''20080110 Denver check WareHouseID when DID return firstly
    str = "exec QSMS_ReturnDID " & sq(ReturnDIDSeq) & "," & sq(Trim(DID)) & "," & sq(COMPPN) & "," & ReturnQty & "," & _
            sq(g_userName) & "," & sq(GroupID) & "," & sq(transdatetime) & "," & sq(IIf(optGoodMaterial.Value = True, "Y", "N")) & "," & sq(PrtCallBKandReturn) & "," & sq(Trim(Factory)) & "," & sq(Trim(IsAnotherBUDID))
            
    Set Rs = Conn.Execute(str)
    If Rs.EOF = False Then
        LblMessage.Caption = Trim(Rs("Description"))
        If Rs("Result") = 0 Then
            intPos = InStr(1, Trim(Rs("Description")), "PreGroupID:")
            PreGroupID = Mid(Trim(Rs("Description")), intPos + Len("PreGroupID:"))
            
            UpdateReturnQty = True
            LblMessage.BackColor = &H80FF80
        End If
        
    End If
    
End Function


Private Function UpdateQSMSGroupCompQty(ByVal GroupID As String) As Boolean
Dim str As String
Dim Rs As ADODB.Recordset
Dim transdatetime As String
str = "select GetDate()"
Set Rs = Conn.Execute(str)
transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHNNSS")
UpdateQSMSGroupCompQty = True


'str = "exec QSMSSap2 '" & Trim(GroupID) & "'"
'Set RS = Conn.Execute(str)
'If Not RS.EOF Then
'    If UCase(RS.Fields(0)) = "PASS" Then
'    Else
'        UpdateQSMSGroupCompQty = False
'        MsgBox RS.Fields(0)
'        Exit Function
'    End If
'End If
End Function

Private Sub txtDID_LostFocus()
    If Trim(TxtDID) = "" Then Exit Sub
    
    ''20080324 Denver    for DID reprint
    If Mid(Right(Trim(TxtDID), 3), 1, 1) = "R" Then Exit Sub
    
  
    If sstabDID.Tab = 0 Then '(00002)
    
        ''--*2009.07.24    Denver    NB2&NB3分家，但是用同一个QSMS可作NB2&NB3 的退料
        IsAnotherBUDID = "N"
        If AutoDispatchForAnotherBU <> "" Then
            ''CompPN 长度至少大于10,如果为True,根据DID生成规则,应该属于另一BU的.
'            If InStr(UCase(Trim(TxtDID)), "-" & UCase(DIDHead)) < 10 Then
'                IsAnotherBUDID = "Y"
'                If XL_ChkAnotherBUDID(UCase(Trim(TxtDID))) = False Then
'                    TxtDID.Text = ""
'                    TxtDID.SetFocus
'                    Exit Sub
'
'                End If
'
'                Exit Sub
'            End If
            ''20090912   Denver    因为对于未导入XL 的CompPN DID 都是在NB3 注册，然后如有需求，再传至NB2,故改为不用DIDHead判别
            If XL_ChkAnotherBUDID(UCase(Trim(TxtDID))) = False Then
                TxtDID.text = ""
                TxtDID.SetFocus
                Exit Sub

            End If
            If IsAnotherBUDID = "Y" Then
                Exit Sub
            End If

        End If
        
    
        '**Denver       2008.04.07     it need not select GroupID and WO when Return DID   --(0009)
        '**Denver       2009.03.03     如果有选择GroupID,就需要检查属于GroupID，  不选不检查(原均不检查)   --(0041)
        If ChkDIDBelongToGroupID(Trim(CboGroupID), Trim(TxtDID)) = False Then

            TxtDID.text = ""
            TxtDID.SetFocus
            Exit Sub
        End If
        
        
        If ChkDIDInMachine(Trim(TxtDID)) = False Then
            TxtDID.text = ""
            TxtDID.SetFocus
            Exit Sub
        End If
        Call GetDIDInfo(Trim(TxtDID), Trim(CboGroupID))
    End If
End Sub

Private Sub TxtDID_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim strCheckScaner As String

 
If strKeyInPNByManual = True Then    '1071
    strCheckScaner = "N"
End If
If strCheckScaner = "Y" Then       '1071
    If App.Title <> App.EXEName Then
        If Button = 2 Then
            MsgBox "Please use scaner to do it!!", vbDefaultButton1
            TxtDID.text = ""
            TxtDID.SetFocus
        End If
    End If
End If
End Sub

Private Sub TxtReturnQty_Click()
    Sendkeys "{home}+{end}"
End Sub

Private Sub TxtReturnQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call cmdSave_Click
End If
End Sub

'
'Private Function GetReturnDIDSeq(ByVal CompPN As String) As String
'Dim Str As String
'Dim Rs As ADODB.Recordset
'Dim I, TempDID As Long
'TempDID = 0
'Str = "select  MAX(right(DID,4)) from QSMS_DID where DID like '" & CompPN & "-A%' "
'Set Rs = Conn.Execute(Str)
'
'If IsNumeric(Trim(Rs.Fields(0))) = False Then
'   GetReturnDIDSeq = CompPN + "-A" + "0001"
'   Exit Function
'Else
'   TempDID = CLng(Rs.Fields(0)) + 1
'   GetReturnDIDSeq = CompPN + "-A" + Format(Trim(TempDID), "0000")
'End If
''TempDID = 1
''While Not Rs.EOF
''      If TempDID <> CLng(Rs.Fields(0)) Then
''         GetDID = CompPN + "-" + Format(Trim(TempDID), "00000")
''         Exit Function
''      End If
''      Rs.MoveNext
''      TempDID = TempDID + 1
''Wend
'
'If GetReturnDIDSeq = "" Then GetReturnDIDSeq = CompPN + "-A" + Format(Trim(TempDID), "0000")
'End Function

Private Function GetReturnDIDSeq(ByVal COMPPN As String) As String
Dim str As String
Dim Rs As ADODB.Recordset
Dim i, TempDID As Integer
str = "select  right(DID,4) from QSMS_DID where DID like '" & COMPPN & "-A%' order by right(DID,4) "
Set Rs = Conn.Execute(str)

If Rs.EOF Then
   GetReturnDIDSeq = COMPPN + "-A" + "0001"
   Exit Function
End If
TempDID = 1
While Not Rs.EOF
      If TempDID <> CLng(Rs.Fields(0)) Then
         GetReturnDIDSeq = COMPPN + "-A" + Format(Trim(TempDID), "0000")
         Exit Function
      End If
      Rs.MoveNext
      TempDID = TempDID + 1
Wend

If GetReturnDIDSeq = "" Then GetReturnDIDSeq = COMPPN + "-A" + Format(Trim(TempDID), "0000")
End Function

Private Function ChkDIDInMachine(ByVal DID As String) As Boolean
Dim str As String
Dim Rs As ADODB.Recordset
ChkDIDInMachine = True
If Trim(TxtDID) = "" Then Exit Function '(00002)
str = "Select Machine,Feeder from QSMS_FeederDID_Current where DID='" & DID & "'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   ChkDIDInMachine = False
   MsgBox "The DID is in Machine :" & Rs!Machine & " Feeder :" & Trim(Rs!Feeder) & "  Please delete first"
   
End If
End Function

Private Function ChkGroupClosed(ByVal GroupID As String) As Boolean
Dim str As String
Dim Rs As ADODB.Recordset
ChkGroupClosed = False
str = "select * from QSMS_WoGroup where GroupID='" & Trim(GroupID) & "' and ClosedFlag<>'Y'"
Set Rs = Conn.Execute(str)
If Rs.EOF Then
   ChkGroupClosed = True
End If
End Function

'Private Function CancelReturnQty(ByVal GroupID As String, CompPN As String, DID As String, ChkDIDBelongToGroupID As Boolean)
'Dim Str As String
'Dim Rs As ADODB.Recordset
'Dim TransDateTime As String
'Dim tempNewDID As String, tempqty As String
'
'Str = "select GetDate()"
'Set Rs = Conn.Execute(Str)
'TransDateTime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
'
'Str = "select NewDID,returnqty,groupid from qsms_groupdid where DID='" & Trim(DID) & "' and groupid='" & GroupID & "'"
'Set Rs = Conn.Execute(Str)
'If Rs.EOF Then
'
'    ChkDIDBelongToGroupID = False
'    Exit Function
'Else
'    tempNewDID = Rs!NewDID
'    tempqty = Rs!ReturnQty
'
'    Str = "delete qsms_did where DID='" & Trim(DID) & "'"
'    Conn.Execute Str
'
'    Str = "Update QSMS_DID set DID='" & Trim(DID) & "' where DID='" & tempNewDID & "'"
'    Conn.Execute Str
'
'
'    Str = "Update QSMS_GroupDID set ReturnQty=ReturnQty-" & tempqty & ",TransDateTime='" & TransDateTime & "' ,NewDID='" & Trim(DID) & "', UID='" & g_userName & "',ReturnTimes=ReturnTimes+1 where GroupID='" & GroupID & "' and DID='" & DID & "'"
'    Conn.Execute Str
'
'    Str = "Update QSMS_GroupCompQty set ReturnQty=ReturnQty-" & tempqty & ",TransDateTime='" & TransDateTime & "' where GroupID='" & GroupID & "' and CompPN='" & CompPN & "'"
'    Conn.Execute Str
'
'    Str = "Update QSMS_Dispatch Set DID='" & Trim(DID) & "' where did='" & tempNewDID & "' and work_order in (select work_order from QSMS_Wogroup where GroupID='" & GroupID & "')"
'    Conn.Execute (Str)
'End If
'
'End Function

Private Function NoExistsSecondDispatch(ReturnTime As String, DID As String) As Boolean
Dim str As String
Dim Rs As ADODB.Recordset
NoExistsSecondDispatch = True
If ReturnTime = "" Then
    NoExistsSecondDispatch = True
    Exit Function
End If

str = "QSMS_NoExistsSecondDispatch '" & ReturnTime & "','" & DID & "'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
    NoExistsSecondDispatch = Rs!result
Else
    NoExistsSecondDispatch = True
End If

ReturnTime = ""
End Function

Public Function ChkDIDBelongToPCB(ByVal WO As String, ByVal DID As String) As Boolean
    'Restore DID from QSMS for splicit 2007-04-11
    sSql = "Exec DIDRestoreForCallBK " & sq(WO) & "," & sq(DID)
    Set Rst = Conn.Execute(sSql)

    ChkDIDBelongToPCB = True
    sSql = "Select DID,ReturnFlag from QSMS_DIDCallBack where work_order in(Select WO from dbo.GetWOGroup('" & Trim(WO) & "')) and DID='" & Trim(DID) & "'"
    Set Rst = Conn.Execute(sSql)
    If Rst.EOF Then
        ChkDIDBelongToPCB = False
        MsgBox "The DID does not belong to the work order,Please check!!"
    Else
        'validation when scan DID Directly
        'can multi CallBack
'       If Trim(rst!ReturnFlag) = "Y" Then
'           ChkDIDBelongToPCB = False
'           MsgBox "The DID has been Call Back,Please check!!"
'       End If
    End If

End Function

Public Function OpenOutputFile(ByVal fname As String)
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
Private Function PrintAutoDispatchLabel()
If OptComp.Value = True Then
    Call PrintAutoDispatchLabelCompPort
ElseIf OptPrint.Value = True Then
    Call PrintAutoDispatchLabelPrintPort
Else
    Call PrintAutoDispatchLabelNetWorkPort
End If
End Function

Private Function PrintAutoDispatchLabelNetWorkPort() As String '(1067)
Dim M As Integer
Dim tmpPrintStr As String
Dim hFile As Long
Dim hString As String
Dim tmpDID As String, strDID As String, strQty As String
Dim strDay As String
Dim LabelFile As String
Dim strPrinterType As String
Dim tmpStr As String
Dim tmpRS As ADODB.Recordset
Dim strSQL As String
Dim rsTime As ADODB.Recordset

        On Error GoTo errHandler
        strDID = UCase(TempDID)
        strQty = Trim(DIDInfo.Qty)
        '1112
        strSQL = "select getdate()"
        Set rsTime = Conn.Execute(strSQL)
        strDay = Format(rsTime(0), "YYMMDDHHNNSS")  '1101
        If StrBU = "NB4" Then                         '1148
            strDay = Format(rsTime(0), "YYYYMMDD")    '
        End If
        ''(1080) replace by 1080
'        If OptZebra.Value = True Then
'            isZebra = True
'
'            ''''''updated by Jing   (0032)''''''
'            If opOldLabel.Value = True Then
'                LabelFile = Settings.AutoDispatchLabel
'            Else
'                LabelFile = Settings.AutoDispatchNewLabel
'            End If
'            strPrinterType = "Zebra"
'        Else
'            isZebra = False
'
'            ''''''updated by Jing   (0032)''''''
'            If opOldLabel.Value = True Then
'                LabelFile = Settings.AutoDispatchSatoLabel
'            Else
'                LabelFile = Settings.AutoDispatchSatoNewLabel
'            End If
'            strPrinterType = "SATO"
'        End If
        
        LabelFile = GetDIDLabelFile(FrmReturnDID, IIf(opOldLabel.Value = True, "OLD", "NEW")) ''(1080) get labelfile
        If Dir(LabelFile) = vbNullString Then
            ''''''Added by Jing 2008.01.10  (0019)''''''
            MsgBox ("Can not find label file !"), vbCritical
            PrintAutoDispatchLabelNetWorkPort = "PRN_FileNoExist"
            Exit Function
        End If
        
        tmpPrintStr = ""
        hFile = FreeFile
        If FileReadAll(tmpPrintStr, LabelFile) <= 0 Then
            MsgBox "Open file:" & LabelFile & " fail!!", vbCritical
            Exit Function
        End If

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
        
         tmpPrintStr = Replace(tmpPrintStr, "<UID>", uId)
         tmpPrintStr = Replace(tmpPrintStr, "<DATE>", strDay)
         tmpPrintStr = Replace(tmpPrintStr, "<QTY>", strQty)
         
         ''''''updated by Jing (0032)''''''
         If opNewLabel.Value = True Then
            tmpPrintStr = Replace(tmpPrintStr, "<BU>", PrintData.Line) '(0037)
         Else
            tmpPrintStr = Replace(tmpPrintStr, "<LINE>", PrintData.Line)
         End If
         tmpPrintStr = Replace(tmpPrintStr, "<SIDE>", PrintData.Side)
         tmpPrintStr = Replace(tmpPrintStr, "<MACHINE>", PrintData.Machine)
         tmpPrintStr = Replace(tmpPrintStr, "<DIDWOGROUP>", PrintData.DIDWOGROUP)
         tmpPrintStr = Replace(tmpPrintStr, "<WOTYPE>", WOType) '(0015)
         tmpPrintStr = Replace(tmpPrintStr, "<Location>", PrintData.location) '1242
         tmpPrintStr = Replace(tmpPrintStr, "<MARK>", PrintData.Mark) '1255
         tmpPrintStr = Replace(tmpPrintStr, "<WHID>", DIDInfo.WareHouseID)    '1252
         ''tmpPrintStr = Replace(tmpPrintStr, "<REELWIDTH>", DIDInfo.ReelWidth)
'         If PrintData.Line <> Left(PrintData.machine, 1) Then '(00044)
'            Conn.Execute ("Insert into QSMS_Error_log(Appname,SubFunction,SubID,Col1,Col2,Col3,DetailDesc,TransDateTime) values(" & _
'                         "'QSMS','PrintDID','Log','" & tmpDID & "','" & PrintData.Line & "','" & PrintData.machine & "'," & _
'                         "'Line and Machine did not match',dbo.formatdate(getdate(),'yyyymmddhhnnss'))")
'         End If
         
        ''''''added by Jing 2008.04.05  (0032)''''''
        If opNewLabel.Value = True Then
            tmpPrintStr = Replace(tmpPrintStr, "<WO1>", WO(0))
            tmpPrintStr = Replace(tmpPrintStr, "<WO2>", WO(1))
            tmpPrintStr = Replace(tmpPrintStr, "<WO3>", WO(2))
            tmpPrintStr = Replace(tmpPrintStr, "<WO4>", WO(3))
            tmpPrintStr = Replace(tmpPrintStr, "<WO5>", WO(4))
            
            tmpPrintStr = Replace(tmpPrintStr, "<MACHINE1>", Machine(0))         '(1084)
            tmpPrintStr = Replace(tmpPrintStr, "<MACHINE2>", Machine(1))
            tmpPrintStr = Replace(tmpPrintStr, "<MACHINE3>", Machine(2))
            tmpPrintStr = Replace(tmpPrintStr, "<MACHINE4>", Machine(3))
            tmpPrintStr = Replace(tmpPrintStr, "<MACHINE5>", Machine(4))
            
            tmpPrintStr = Replace(tmpPrintStr, "<SLOT1>", Slot(0))        '(1086)
            tmpPrintStr = Replace(tmpPrintStr, "<SLOT2>", Slot(1))
            tmpPrintStr = Replace(tmpPrintStr, "<SLOT3>", Slot(2))
            tmpPrintStr = Replace(tmpPrintStr, "<SLOT4>", Slot(3))
            tmpPrintStr = Replace(tmpPrintStr, "<SLOT5>", Slot(4))
                       
            tmpPrintStr = Replace(tmpPrintStr, "<Model1>", Model(0))
            tmpPrintStr = Replace(tmpPrintStr, "<Model2>", Model(1))
            tmpPrintStr = Replace(tmpPrintStr, "<Model3>", Model(2))
            tmpPrintStr = Replace(tmpPrintStr, "<Model4>", Model(3))
            tmpPrintStr = Replace(tmpPrintStr, "<Model5>", Model(4))
            
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order1>", Work_Order(0))  ''''1093
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order2>", Work_Order(1))
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order3>", Work_Order(2))
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order4>", Work_Order(3))
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order5>", Work_Order(4))
        
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType1>", DIDType(0)) ''''1093
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType2>", DIDType(1))
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType3>", DIDType(2))
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType4>", DIDType(3))
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType5>", DIDType(4))
            
            tmpPrintStr = Replace(tmpPrintStr, "<CYL1>", ISCYL(0)) ''''1109
            tmpPrintStr = Replace(tmpPrintStr, "<CYL2>", ISCYL(1))
            tmpPrintStr = Replace(tmpPrintStr, "<CYL3>", ISCYL(2))
            tmpPrintStr = Replace(tmpPrintStr, "<CYL4>", ISCYL(3))
            tmpPrintStr = Replace(tmpPrintStr, "<CYL5>", ISCYL(4))
            
            tmpPrintStr = Replace(tmpPrintStr, "<COUNT1>", SeqID(0))  '(1148)
            tmpPrintStr = Replace(tmpPrintStr, "<COUNT2>", SeqID(1))
            tmpPrintStr = Replace(tmpPrintStr, "<COUNT3>", SeqID(2))
            tmpPrintStr = Replace(tmpPrintStr, "<COUNT4>", SeqID(3))
            tmpPrintStr = Replace(tmpPrintStr, "<COUNT5>", SeqID(4))
            
            tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE1>", VenderCode(0))    ''1227
            tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE2>", VenderCode(1))
            tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE3>", VenderCode(2))
            tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE4>", VenderCode(3))
            tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE5>", VenderCode(4))
         
            tmpPrintStr = Replace(tmpPrintStr, "<LR1>", LR(0))                      ''1227
            tmpPrintStr = Replace(tmpPrintStr, "<LR2>", LR(1))
            tmpPrintStr = Replace(tmpPrintStr, "<LR3>", LR(2))
            tmpPrintStr = Replace(tmpPrintStr, "<LR4>", LR(3))
            tmpPrintStr = Replace(tmpPrintStr, "<LR5>", LR(4))
            
        tmpPrintStr = Replace(tmpPrintStr, "<MachineCH1>", MachineCH(0))                    '1247
        tmpPrintStr = Replace(tmpPrintStr, "<MachineCH2>", MachineCH(1))
        tmpPrintStr = Replace(tmpPrintStr, "<MachineCH3>", MachineCH(2))
        tmpPrintStr = Replace(tmpPrintStr, "<MachineCH4>", MachineCH(3))
        tmpPrintStr = Replace(tmpPrintStr, "<MachineCH5>", MachineCH(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH1>", SideCH(0))                    '1247
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH2>", SideCH(1))
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH3>", SideCH(2))
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH4>", SideCH(3))
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH5>", SideCH(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH1>", LRCH(0))                    '1247
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH2>", LRCH(1))
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH3>", LRCH(2))
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH4>", LRCH(3))
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH5>", LRCH(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH1>", SlotCH(0))                    '1247
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH2>", SlotCH(1))
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH3>", SlotCH(2))
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH4>", SlotCH(3))
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH5>", SlotCH(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<REELWIDTH>", ReelWidth(0))                    '1247
        tmpPrintStr = Replace(tmpPrintStr, "<REELWIDTH1>", ReelWidth(1))
        tmpPrintStr = Replace(tmpPrintStr, "<REELWIDTH2>", ReelWidth(2))
        tmpPrintStr = Replace(tmpPrintStr, "<REELWIDTH3>", ReelWidth(3))
        tmpPrintStr = Replace(tmpPrintStr, "<REELWIDTH4>", ReelWidth(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<PN1>", PN(0))                    '1247
        tmpPrintStr = Replace(tmpPrintStr, "<PN2>", PN(1))
        tmpPrintStr = Replace(tmpPrintStr, "<PN3>", PN(2))
        tmpPrintStr = Replace(tmpPrintStr, "<PN4>", PN(3))
        tmpPrintStr = Replace(tmpPrintStr, "<PN5>", PN(4))
            
            
            
            
            
            
            '(1063)
            tmpPrintStr = Replace(tmpPrintStr, "<MACHINETYPE>", Mid(PrintData.Machine, Len(PrintData.Machine) - 3, 3))
            tmpPrintStr = Replace(tmpPrintStr, "<MACHINECODE>", Right(PrintData.Machine, 1))
            If InStr(WO(0), " ") > 1 Then
                tmpPrintStr = Replace(tmpPrintStr, "<SLOT>", Mid(WO(0), InStr(WO(0), " ") + 1, Len(WO(0)) - InStr(WO(0), " ")))
            End If
        End If

        Select Case Trim(tmpPrintStr)
           Case vbNullString
           Case Else
                Printer.Print tmpPrintStr
                Printer.EndDoc
                Printer.KillDoc
        End Select

        Exit Function
errHandler:
        MsgBox Err.Description
End Function

Private Function PrintAutoDispatchLabelCompPort() As String '(0013)
Dim M As Integer
Dim tmpPrintStr As String
Dim hFile As Long
Dim hString As String
Dim tmpDID As String, strDID As String, strQty As String
Dim strDay As String
Dim LabelFile As String
Dim strPrinterType As String
Dim tmpStr As String
Dim tmpRS As ADODB.Recordset
Dim strSQL As String
Dim rsTime As ADODB.Recordset

        On Error GoTo errHandler
        strDID = UCase(TempDID)
        strQty = Trim(DIDInfo.Qty)
        '1112
        strSQL = "select getdate()"
        Set rsTime = Conn.Execute(strSQL)
        strDay = Format(rsTime(0), "YYMMDDHHNNSS") '1101
        
        ''20090818   Denver   ''get DID head after return did
'        tmpStr = "Select DIDHead from site" '(0037)
'        Set tmpRS = Conn.Execute(tmpStr)
'        If tmpRS.EOF Then
'           MsgBox "can not find the DIDHead in the Table,Please check"
'           Exit Function
'        Else
'            PrintData.BU = Trim(tmpRS!DIDHead)
'        End If
        
        
        '''*Denver        2009.07.23     NB2&NB3分家，但是用同一个QSMS可作NB2NB3 的退料
        ''不需要在内部调用
        ''''''added by Jing 2008.04.05  (0032)''''''
'        If opNewLabel.Value = True Then
'            Dim X As Integer
'            For X = 0 To 4
'                WO(X) = ""
'                Model(X) = ""
'            Next X
        
'            tmpStr = "select distinct a.Machine,a.Slot,A.LR,substring(b.PN,3,3) as Model from QSMS_Dispatch a,sap_wo_list b where a.did='" & Trim(strDID) & "' and a.work_order=b.wo"
'            Set tmpRS = Conn.Execute(tmpStr)
'            If tmpRS.EOF = False Then
'                Dim i As Integer, j As Integer
'                j = tmpRS.RecordCount
'                If j > 5 Then j = 5
'                For i = 0 To j - 1
'                    WO(i) = tmpRS("Machine") + " " + tmpRS("Slot") + "-" + tmpRS("LR") '0043
'                    Model(i) = tmpRS("model")
'                    tmpRS.MoveNext
'                Next i
'            End If
'        End If
        ''(1080) replace by 1080
'        If OptZebra.Value = True Then
'            isZebra = True
'
'            ''''''updated by Jing   (0032)''''''
'            If opOldLabel.Value = True Then
'                LabelFile = Settings.AutoDispatchLabel
'            Else
'                LabelFile = Settings.AutoDispatchNewLabel
'            End If
'            strPrinterType = "Zebra"
'        Else
'            isZebra = False
'
'            ''''''updated by Jing   (0032)''''''
'            If opOldLabel.Value = True Then
'                LabelFile = Settings.AutoDispatchSatoLabel
'            Else
'                LabelFile = Settings.AutoDispatchSatoNewLabel
'            End If
'            strPrinterType = "SATO"
'        End If
        
        LabelFile = GetDIDLabelFile(FrmReturnDID, IIf(opOldLabel.Value = True, "OLD", "NEW")) ''(1080) get labelfile
        
        If Dir(LabelFile) = vbNullString Then
            ''''''Added by Jing 2008.01.10  (0019)''''''
            MsgBox ("Can not find label file !"), vbCritical
            PrintAutoDispatchLabelCompPort = "PRN_FileNoExist"
            Exit Function
        End If
        
        MSComm.CommPort = TxtCompPort 'Settings.PRNa_Port
        MSComm.Settings = TxtComm 'Settings.PRNa_Settings
        MSComm.OutBufferCount = 0 '清空输出缓存
        
        If MSComm.PortOpen = False Then MSComm.PortOpen = True
        tmpPrintStr = ""
        hFile = FreeFile
        If FileReadAll(tmpPrintStr, LabelFile) <= 0 Then
            MsgBox "Open file:" & LabelFile & " fail!!", vbCritical
            Exit Function
        End If
'        Open LabelFile For Input As #hFile
'        Do
'           Select Case EOF(hFile)
'              Case True
'                Close #hFile
'                PrintAutoDispatchLabelCompPort = "PRN_Succeed"
'                Exit Do
'              Case False
'                Line Input #hFile, hString
'                hString = Trim(hString)
'                tmpPrintStr = tmpPrintStr & Trim(hString)
'          End Select
'        Loop
'        Close #hFile
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
        
         tmpPrintStr = Replace(tmpPrintStr, "<UID>", uId)
         tmpPrintStr = Replace(tmpPrintStr, "<DATE>", strDay)
         tmpPrintStr = Replace(tmpPrintStr, "<QTY>", strQty)
         
         ''''''updated by Jing (0032)''''''
         If opNewLabel.Value = True Then
            tmpPrintStr = Replace(tmpPrintStr, "<BU>", PrintData.Line) '(0037)
         Else
            tmpPrintStr = Replace(tmpPrintStr, "<LINE>", PrintData.Line)
         End If
         tmpPrintStr = Replace(tmpPrintStr, "<SIDE>", PrintData.Side)
         tmpPrintStr = Replace(tmpPrintStr, "<MACHINE>", PrintData.Machine)
         tmpPrintStr = Replace(tmpPrintStr, "<DIDWOGROUP>", PrintData.DIDWOGROUP)
         tmpPrintStr = Replace(tmpPrintStr, "<WOTYPE>", WOType) '(0015)
         tmpPrintStr = Replace(tmpPrintStr, "<Location>", PrintData.location) '1242
         tmpPrintStr = Replace(tmpPrintStr, "<MARK>", PrintData.Mark) '1255
         tmpPrintStr = Replace(tmpPrintStr, "<WHID>", DIDInfo.WareHouseID)     '1252
         
'         If PrintData.Line <> Left(PrintData.machine, 1) Then '(00044)
'            Conn.Execute ("Insert into QSMS_Error_log(Appname,SubFunction,SubID,Col1,Col2,Col3,DetailDesc,TransDateTime) values(" & _
'                         "'QSMS','PrintDID','Log','" & tmpDID & "','" & PrintData.Line & "','" & PrintData.machine & "'," & _
'                         "'Line and Machine did not match',dbo.formatdate(getdate(),'yyyymmddhhnnss'))")
'         End If
         
        ''''''added by Jing 2008.04.05  (0032)''''''
        If opNewLabel.Value = True Then
            tmpPrintStr = Replace(tmpPrintStr, "<WO1>", WO(0))
            tmpPrintStr = Replace(tmpPrintStr, "<WO2>", WO(1))
            tmpPrintStr = Replace(tmpPrintStr, "<WO3>", WO(2))
            tmpPrintStr = Replace(tmpPrintStr, "<WO4>", WO(3))
            tmpPrintStr = Replace(tmpPrintStr, "<WO5>", WO(4))
            
            tmpPrintStr = Replace(tmpPrintStr, "<Model1>", Model(0))
            tmpPrintStr = Replace(tmpPrintStr, "<Model2>", Model(1))
            tmpPrintStr = Replace(tmpPrintStr, "<Model3>", Model(2))
            tmpPrintStr = Replace(tmpPrintStr, "<Model4>", Model(3))
            tmpPrintStr = Replace(tmpPrintStr, "<Model5>", Model(4))
            
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order1>", Work_Order(0))  ''''1093
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order2>", Work_Order(1))
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order3>", Work_Order(2))
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order4>", Work_Order(3))
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order5>", Work_Order(4))
        
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType1>", DIDType(0)) ''''1093
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType2>", DIDType(1))
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType3>", DIDType(2))
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType4>", DIDType(3))
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType5>", DIDType(4))
            
            tmpPrintStr = Replace(tmpPrintStr, "<CYL1>", ISCYL(0))  ''(1109）
            tmpPrintStr = Replace(tmpPrintStr, "<CYL2>", ISCYL(1))
            tmpPrintStr = Replace(tmpPrintStr, "<CYL3>", ISCYL(2))
            tmpPrintStr = Replace(tmpPrintStr, "<CYL4>", ISCYL(3))
            tmpPrintStr = Replace(tmpPrintStr, "<CYL5>", ISCYL(4))
            
            tmpPrintStr = Replace(tmpPrintStr, "<COUNT1>", SeqID(0))  '(1148)
            tmpPrintStr = Replace(tmpPrintStr, "<COUNT2>", SeqID(1))
            tmpPrintStr = Replace(tmpPrintStr, "<COUNT3>", SeqID(2))
            tmpPrintStr = Replace(tmpPrintStr, "<COUNT4>", SeqID(3))
            tmpPrintStr = Replace(tmpPrintStr, "<COUNT5>", SeqID(4))
            
            tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE1>", VenderCode(0))
            tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE2>", VenderCode(1))
            tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE3>", VenderCode(2))
            tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE4>", VenderCode(3))
            tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE5>", VenderCode(4))
         
            tmpPrintStr = Replace(tmpPrintStr, "<LR1>", LR(0))
            tmpPrintStr = Replace(tmpPrintStr, "<LR2>", LR(1))
            tmpPrintStr = Replace(tmpPrintStr, "<LR3>", LR(2))
            tmpPrintStr = Replace(tmpPrintStr, "<LR4>", LR(3))
            tmpPrintStr = Replace(tmpPrintStr, "<LR5>", LR(4))
            
            
            
        tmpPrintStr = Replace(tmpPrintStr, "<MachineCH1>", MachineCH(0))                    '1247
        tmpPrintStr = Replace(tmpPrintStr, "<MachineCH2>", MachineCH(1))
        tmpPrintStr = Replace(tmpPrintStr, "<MachineCH3>", MachineCH(2))
        tmpPrintStr = Replace(tmpPrintStr, "<MachineCH4>", MachineCH(3))
        tmpPrintStr = Replace(tmpPrintStr, "<MachineCH5>", MachineCH(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH1>", SideCH(0))                    '1247
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH2>", SideCH(1))
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH3>", SideCH(2))
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH4>", SideCH(3))
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH5>", SideCH(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH1>", LRCH(0))                    '1247
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH2>", LRCH(1))
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH3>", LRCH(2))
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH4>", LRCH(3))
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH5>", LRCH(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH1>", SlotCH(0))                    '1247
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH2>", SlotCH(1))
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH3>", SlotCH(2))
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH4>", SlotCH(3))
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH5>", SlotCH(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<PN1>", PN(0))                    '1247
        tmpPrintStr = Replace(tmpPrintStr, "<PN2>", PN(1))
        tmpPrintStr = Replace(tmpPrintStr, "<PN3>", PN(2))
        tmpPrintStr = Replace(tmpPrintStr, "<PN4>", PN(3))
        tmpPrintStr = Replace(tmpPrintStr, "<PN5>", PN(4))
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            '(1063)
            tmpPrintStr = Replace(tmpPrintStr, "<MACHINETYPE>", Mid(PrintData.Machine, Len(PrintData.Machine) - 3, 3))
            tmpPrintStr = Replace(tmpPrintStr, "<MACHINECODE>", Right(PrintData.Machine, 1))
            If InStr(WO(0), " ") > 1 Then
                tmpPrintStr = Replace(tmpPrintStr, "<SLOT>", Mid(WO(0), InStr(WO(0), " ") + 1, Len(WO(0)) - InStr(WO(0), " ")))
            End If
        End If

        Select Case Trim(tmpPrintStr)
           Case vbNullString
           Case Else
                 For M = 1 To Len(tmpPrintStr) Step 50
                     MSComm.Output = Mid(tmpPrintStr, M, 50)
                     'Debug.Print Mid(hString, m, 50)
                     DoEvents
                 Next M
        End Select

        MSComm.PortOpen = False
        Exit Function
errHandler:
        MsgBox Err.Description
        If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
        End If
End Function
'Private Function PrintAutoDispatchLabelCompPort() As String
'Dim hFile As Long
'Dim hString As String, isZebra As Boolean
'Dim strDID As String, tmpDID As String, strQty As String
'Dim strDay As String
'Dim LabelFile As String
'Dim m As Integer
'Dim strPrinterType As String
'        On Error GoTo ErrHandler
'        strDay = Format(Now, "YYYY/MM/DD")
'        If OptZebra.Value = True Then
'            isZebra = True
'            LabelFile = Settings.AutoDispatchLabel
'            strPrinterType = "Zebra-Return"
'        Else
'            isZebra = False
'            LabelFile = Settings.AutoDispatchSatoLabel
'            strPrinterType = "SATO-Return"
'        End If
'        strDID = UCase(TempDID)
'        strQty = Trim(DIDInfo.Qty)
'
'        '**********************save the DID print log (0025)**********************
'        strSQL = "insert  into QSMS_Error_Log(AppName,SubFunction,SubID,DetailDesc,Col1,Col2,Col3,TransDateTime) values('QSMS_XL','PrintDID','Print DID Log','" & Trim(strDID) & "','" & Trim(strPrinterType) & "','" & g_userName & "','Com',dbo.FormatDate(getdate(),'yyyymmddhhnnss'))"
'        Conn.Execute strSQL
'        '**********************save the DID print log (0025)**********************
'
'        If Dir(LabelFile) = vbNullString Then
'            ''''''Added by Jing 2008.01.10  (00003)''''''
'            MsgBox ("Can not find label file !"), vbCritical
'            PrintAutoDispatchLabelCompPort = "PRN_FileNoExist"
'            Exit Function
'        End If
'
'        MSComm.CommPort = TxtCompPort 'Settings.PRNa_Port
'        MSComm.Settings = TxtComm 'Settings.PRNa_Settings
'        MSComm.OutBufferCount = 0 '清空输出缓存
'
'        If MSComm.PortOpen = False Then MSComm.PortOpen = True
'
'        hFile = FreeFile
'        Open LabelFile For Input As #hFile
'        Do
'           Select Case EOF(hFile)
'              Case True
'                Close #hFile
'                PrintAutoDispatchLabelCompPort = "PRN_Succeed"
'                Exit Do
'              Case False
'                Line Input #hFile, hString
'                hString = Trim(hString)
'                tmpDID = Trim(strDID) '***************add by jeanson 20070814******
'                'for Code 128 barcode, the ^ must be tranfer to ><
'                If InStr(hString, "<DID_CODE>") > 0 Then
'                    ''********************************updated by jing 20071024 (0002) ***********
'                    If isZebra Then
'                        tmpDID = Replace(strDID, "^", "><")
'                    End If
'                   hString = Replace(hString, "<DID_CODE>", tmpDID)
'                End If
'                'for text ^, must be use ^FH_ and the use _5E (the ascii of ^)
'                If InStr(hString, "<DID_TEXT>") > 0 Then
'                    ''********************************updated by jing 20071024 (0002) ***********
'                    If isZebra Then
'                       tmpDID = Replace(strDID, "^", "_5E")
'                    End If
'                   hString = Replace(hString, "<DID_TEXT>", tmpDID)
'                End If
'
'                hString = Replace(hString, "<UID>", UID)
''                hString = Replace(hString, "<RACKID>", TxtRackID)
'                hString = Replace(hString, "<DATE>", strDay)
'                hString = Replace(hString, "<QTY>", strQty)
'                hString = Replace(hString, "<LINE>", PrintData.Line)
'                hString = Replace(hString, "<SIDE>", PrintData.Side)
'                hString = Replace(hString, "<MACHINE>", PrintData.Machine)
'                hString = Replace(hString, "<DIDWOGROUP>", PrintData.DIDWOGROUP) '(0015)
'                Select Case Trim(hString)
'                    Case vbNullString
'                    Case Else
'                    If isZebra Then
'                        MSComm.Output = hString
'                    Else '   (0016)
'                        For m = 1 To Len(hString) Step 50
'                            MSComm.Output = Mid(hString, m, 50)
'                        Next m
'                    End If
'                End Select
'
''               Select Case Trim(hString)
''                  Case vbNullString
''                  Case Else
''                    MSComm.Output = hString
''                    Debug.Print hString
''               End Select
'          End Select
'        Loop
'
'        Close #hFile
'        MSComm.PortOpen = False
'        Exit Function
'ErrHandler:
'        MsgBox Err.Description
'        If MSComm.PortOpen = True Then
'            MSComm.PortOpen = False
'        End If
'End Function

Private Function PrintAutoDispatchLabelPrintPort() As String
        Dim hFile As Long       '',isZebra As Boolean
        Dim hString As String
        Dim strDID As String, tmpDID As String, strQty As String
        Dim FileNum As Integer, lptPort As Integer
        Dim strDay As String
        Dim LabelFile, strLabelFileContent As String
        Dim strPort As String
        Dim strPrinterType As String
        Dim strSQL As String
        Dim rsTime As ADODB.Recordset
        
        On Error GoTo errHandler
        '1112
        strSQL = "select getdate()"
        Set rsTime = Conn.Execute(strSQL)
        strDay = Format(rsTime(0), "YYMMDDHHNNSS") '1101
        strDID = UCase(TempDID)
        strQty = Trim(DIDInfo.Qty)
        ''(1080) replace by 1080
'        If OptZebra.Value = True Then
'            isZebra = True
'            If opNewLabel.Value = True Then '(00045)
'                LabelFile = Settings.AutoDispatchNewLabel
'            Else
'                LabelFile = Settings.AutoDispatchLabel
'            End If
'            strPrinterType = "Zebra-Return"
'
'        Else
'            isZebra = False
'            If opNewLabel.Value = True Then '(00045)
'                LabelFile = Settings.AutoDispatchSatoNewLabel
'            Else
'                LabelFile = Settings.AutoDispatchSatoLabel
'            End If
'            strPrinterType = "SATO-Return"
'        End If
'        strLabelFileContent = funGetTxtFileContent(LabelFile)
        
        LabelFile = GetDIDLabelFile(FrmReturnDID, IIf(opNewLabel.Value = True, "NEW", "OLD")) ''(1080) get labelfile
        
        If Dir(LabelFile) = vbNullString Then
            ''''''Added by Jing 2008.01.10  (00003)''''''
            MsgBox ("Can not find label file !"), vbCritical
            PrintAutoDispatchLabelPrintPort = "PRN_FileNoExist"
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
'        If FileReadAll(hString, LabelFile) <= 0 Then
'            MsgBox "Open file:" & LabelFile & " fail!!", vbCritical
'            Exit Function
'        End If
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
                hString = Replace(hString, "<UID>", uId)
'                hString = Replace(hString, "<RACKID>", TxtRackID)
                hString = Replace(hString, "<DATE>", strDay)
                hString = Replace(hString, "<QTY>", strQty)
                If opNewLabel.Value = True Then  '(00045)
                    hString = Replace(hString, "<BU>", PrintData.Line)
                Else
                    hString = Replace(hString, "<LINE>", PrintData.Line)
                End If
                hString = Replace(hString, "<SIDE>", PrintData.Side)
                hString = Replace(hString, "<MACHINE>", PrintData.Machine)
                hString = Replace(hString, "<DIDWOGROUP>", PrintData.DIDWOGROUP) '(0015)
                hString = Replace(hString, "<WOTYPE>", WOType) '(0015)
                hString = Replace(hString, "<Location>", PrintData.location) '1242
                hString = Replace(hString, "<MARK>", PrintData.Mark) '1255
                hString = Replace(hString, "<WHID>", DIDInfo.WareHouseID)     '1252
                
                If opNewLabel.Value = True Then '(00045)
                    hString = Replace(hString, "<WO1>", WO(0))
                    hString = Replace(hString, "<WO2>", WO(1))
                    hString = Replace(hString, "<WO3>", WO(2))
                    hString = Replace(hString, "<WO4>", WO(3))
                    hString = Replace(hString, "<WO5>", WO(4))
                    
                    hString = Replace(hString, "<Model1>", Model(0))
                    hString = Replace(hString, "<Model2>", Model(1))
                    hString = Replace(hString, "<Model3>", Model(2))
                    hString = Replace(hString, "<Model4>", Model(3))
                    hString = Replace(hString, "<Model5>", Model(4))
                    
                    hString = Replace(hString, "<Work_Order1>", Work_Order(0))  ''''1093
                    hString = Replace(hString, "<Work_Order2>", Work_Order(1))
                    hString = Replace(hString, "<Work_Order3>", Work_Order(2))
                    hString = Replace(hString, "<Work_Order4>", Work_Order(3))
                    hString = Replace(hString, "<Work_Order5>", Work_Order(4))
        
                    hString = Replace(hString, "<DIDType1>", DIDType(0)) ''''1093
                    hString = Replace(hString, "<DIDType2>", DIDType(1))
                    hString = Replace(hString, "<DIDType3>", DIDType(2))
                    hString = Replace(hString, "<DIDType4>", DIDType(3))
                    hString = Replace(hString, "<DIDType5>", DIDType(4))
                    
                    hString = Replace(hString, "<CYL1>", ISCYL(0)) ''(1109）
                    hString = Replace(hString, "<CYL2>", ISCYL(1))
                    hString = Replace(hString, "<CYL3>", ISCYL(2))
                    hString = Replace(hString, "<CYL4>", ISCYL(3))
                    hString = Replace(hString, "<CYL5>", ISCYL(4))
                    
                    hString = Replace(hString, "<COUNT1>", SeqID(0))  '(1148)
                    hString = Replace(hString, "<COUNT2>", SeqID(1))
                    hString = Replace(hString, "<COUNT3>", SeqID(2))
                    hString = Replace(hString, "<COUNT4>", SeqID(3))
                    hString = Replace(hString, "<COUNT5>", SeqID(4))
                   
                    hString = Replace(hString, "<VENDORCODE1>", VenderCode(0))    ''1227
                    hString = Replace(hString, "<VENDORCODE2>", VenderCode(1))
                    hString = Replace(hString, "<VENDORCODE3>", VenderCode(2))
                    hString = Replace(hString, "<VENDORCODE4>", VenderCode(3))
                    hString = Replace(hString, "<VENDORCODE5>", VenderCode(4))
                
                    hString = Replace(hString, "<LR1>", LR(0))                   ''1227
                    hString = Replace(hString, "<LR2>", LR(1))
                    hString = Replace(hString, "<LR3>", LR(2))
                    hString = Replace(hString, "<LR4>", LR(3))
                    hString = Replace(hString, "<LR5>", LR(4))
                    
                    
        hString = Replace(hString, "<MachineCH1>", MachineCH(0))                    '1247
        hString = Replace(hString, "<MachineCH2>", MachineCH(1))
        hString = Replace(hString, "<MachineCH3>", MachineCH(2))
        hString = Replace(hString, "<MachineCH4>", MachineCH(3))
        hString = Replace(hString, "<MachineCH5>", MachineCH(4))
        
        hString = Replace(hString, "<SideCH1>", SideCH(0))                    '1247
        hString = Replace(hString, "<SideCH2>", SideCH(1))
        hString = Replace(hString, "<SideCH3>", SideCH(2))
        hString = Replace(hString, "<SideCH4>", SideCH(3))
        hString = Replace(hString, "<SideCH5>", SideCH(4))
        
        hString = Replace(hString, "<LRCH1>", LRCH(0))                    '1247
        hString = Replace(hString, "<LRCH2>", LRCH(1))
        hString = Replace(hString, "<LRCH3>", LRCH(2))
        hString = Replace(hString, "<LRCH4>", LRCH(3))
        hString = Replace(hString, "<LRCH5>", LRCH(4))
        
        hString = Replace(hString, "<SlotCH1>", SlotCH(0))                    '1247
        hString = Replace(hString, "<SlotCH2>", SlotCH(1))
        hString = Replace(hString, "<SlotCH3>", SlotCH(2))
        hString = Replace(hString, "<SlotCH4>", SlotCH(3))
        hString = Replace(hString, "<SlotCH5>", SlotCH(4))
        
        hString = Replace(hString, "<PN1>", PN(0))                    '1247
        hString = Replace(hString, "<PN2>", PN(1))
        hString = Replace(hString, "<PN3>", PN(2))
        hString = Replace(hString, "<PN4>", PN(3))
        hString = Replace(hString, "<PN5>", PN(4))
                        
                    
                    
                    
                    
                    
                    
                    
                    
                    
                End If
                Print #lptPort, hString & Chr(13)
        Wend
        
'        If PrintData.Line <> Left(PrintData.machine, 1) Then '(00044)
'            Conn.Execute ("Insert into QSMS_Error_log(Appname,SubFunction,SubID,Col1,Col2,Col3,DetailDesc,TransDateTime) values(" & _
'                         "'QSMS','PrintDID','Log','" & tmpDID & "','" & PrintData.Line & "','" & PrintData.machine & "'," & _
'                         "'Line and Machine did not match',dbo.formatdate(getdate(),'yyyymmddhhnnss'))")
'         End If
        
        
'        Open strPort For Output As #FileNum
'        Print #FileNum, strLabelFileContent
        Close #FileNum
        Close #lptPort
        Exit Function
errHandler:
     MsgBox Err.Description
End Function


''20071226 Denver Print DID for CallBack
'20100407 Denver    BU Name change (ESBU to CC,ASBU to LC)  （0070）
''===================================================
'Private Function DIDPrintLabel(blnCompPort As Boolean, blnZebra As Boolean, intCompPort As Integer, sCommString As String)

   ' If blnCompPort = True Then
       'Call PrintLabelCompPort(blnZebra, intCompPort, sCommString)
    'Else
       'Call PrintLabelPrintPort(blnZebra)
    'End If
'End Function
Private Function DIDPrintLabel(blnZebra As Boolean, intCompPort As Integer, sCommString As String) As String
    Dim hFile As Long
    Dim hString As String
    Dim strDID As String, tmpDID As String, strQty As String, strDIDType As String
    Dim strDay As String
    Dim LabelFile As String
    Dim M As Integer
    Dim lptPort As Integer
    Dim tmpPrintStr As String
    Dim strSQL As String
    Dim VendorCode As String
    Dim rsTime As ADODB.Recordset
    Dim DIDLocation As String
    Dim DIDMark As String
    
        On Error GoTo errHandler
        '1112
        strSQL = "select getdate()"
        Set rsTime = Conn.Execute(strSQL)
        strDay = Format(rsTime(0), "YYMMDDHHNNSS")  '1101
        
        DIDLocation = ""
        
        ''20100423   Denver    DID Return后，如无需求，现不直接退库，打印时DID位置显示CompPN(0071)
        If blnZebra = True Then
'            LabelFile = Settings.AutoDispatchLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
'                LabelFile = Settings.DIDLabelGood
                strDID = DIDInfo.DID
                
                 ''20110127   denver    CCBU DIDnotToQWMS Label need print DID Information
'                If DIDnotToQWMS = "Y" Then
'                    strDID = DIDInfo.compPN
'                Else
'                    strDID = DIDInfo.DID
'                End If
            Else
'                LabelFile = Settings.DIDLabelBad
                strDID = DIDInfo.COMPPN
            End If
        Else
'            LabelFile = Settings.AutoDispatchSatoLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
'                LabelFile = Settings.DIDLabelSATOGood
                strDID = DIDInfo.DID
                ''20110127   denver    CCBU DIDnotToQWMS Label need print DID Information
'                If DIDnotToQWMS = "Y" Then
'                    strDID = DIDInfo.compPN
'                Else
'                    strDID = DIDInfo.DID
'                End If
            Else
'                LabelFile = Settings.DIDLabelSATOBad
                strDID = DIDInfo.COMPPN
            End If
        End If
        strDIDType = DIDInfo.DIDType
        
        sSql = "select * from MSD_DATA where CompPN=left(" & sq(strDID) & ",11)"    '''(1272)
        Set Rst = Conn.Execute(sSql)
        
        If BU = "ESBU" And Rst.EOF = False Then
            LabelFile = GetDIDLabelFile(frmDIDCallBack_New, "GOOD_MSD")
        Else
            LabelFile = GetDIDLabelFile(FrmReturnDID, IIf(DIDInfo.IsGood = "Y", "GOOD", "BAD")) ''(1080) Get labelfile
        End If
        
        
        DIDLocation = DIDInfo.location  ''1242
        DIDMark = DIDInfo.Mark ''1255
        
        ''特殊处理(打印RefID Label)
        If DIDInfo.Qty <= -10000 Then
            strQty = "RefID"
        Else
            strQty = CStr(DIDInfo.Qty)
        End If
     
        If (Dir(LabelFile) = vbNullString) Then
            ''''''Added by Jing 2008.01.10  (00003)'''''
            MsgBox ("Can not find Lable file !"), vbCritical
            DIDPrintLabel = "PRN_FileNoExist"
            Exit Function
        End If
        
        'TxtCompPort TxtComm
        If OptComp.Value = True Then
            MSComm.CommPort = intCompPort
            MSComm.Settings = sCommString
            MSComm.OutBufferCount = 0 '清空输出缓存
            
            If MSComm.PortOpen = False Then MSComm.PortOpen = True
        ElseIf OptPrint.Value = True Then
            lptPort = OpenOutputFile("LPT1")
            If lptPort = 0 Then
                MsgBox "Open print port LPT1 error!"
                Exit Function
            End If
        End If
        
        ''(0077)
        hFile = FreeFile
        If FileReadAll(tmpPrintStr, LabelFile) <= 0 Then
            MsgBox "Open file:" & LabelFile & " fail!!", vbCritical
            Exit Function
        End If
'        Open LabelFile For Input As #hFile
'        Do
'           Select Case EOF(hFile)
'              Case True
'                Close #hFile
'                DIDPrintLabel = "PRN_Succeed"
'                Exit Do
'              Case False
'                Line Input #hFile, hString
'                hString = Trim(hString)
'                tmpPrintStr = tmpPrintStr & Trim(hString)
'            End Select
'        Loop
'        Close #hFile
         tmpDID = Trim(strDID) '***************add by jeanson 20070814******
         VendorCode = DIDInfo.COMPPN + ";" + DIDInfo.DateCode + ";" + DIDInfo.VendorCode + ";" + DIDInfo.LotCode    ''1227
         'for Code 128 barcode, the ^ must be tranfer to ><
         If InStr(tmpPrintStr, "<DID_CODE>") > 0 Then
             ''********************************updated by jing 20071024 (0002) ***********
             If blnZebra Then
                 tmpDID = Replace(strDID, "^", "><")
             End If
             tmpPrintStr = Replace(tmpPrintStr, "<DID_CODE>", tmpDID)
         End If
        
        If InStr(tmpPrintStr, "<DID_2D>") > 0 Then
      
            If blnZebra Then
                tmpDID = Trim(Replace(strDID, "^", "_5E"))
            End If
             tmpPrintStr = Replace(tmpPrintStr, "<DID_2D>", tmpDID)
        End If
         
         'for text ^, must be use ^FH_ and the use _5E (the ascii of ^)
         If InStr(tmpPrintStr, "<DID_TEXT>") > 0 Then
             ''********************************updated by jing 20071024 (0002) ***********
             If blnZebra Then
                tmpDID = Replace(strDID, "^", "_5E")
             End If
             tmpPrintStr = Replace(tmpPrintStr, "<DID_TEXT>", tmpDID)
         End If

         tmpPrintStr = Replace(tmpPrintStr, "<UID>", uId)
         tmpPrintStr = Replace(tmpPrintStr, "<DATE>", strDay)
         tmpPrintStr = Replace(tmpPrintStr, "<QTY>", strQty)
         tmpPrintStr = Replace(tmpPrintStr, "<DIDType>", strDIDType)
         tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE1>", VendorCode)   ''1227
         tmpPrintStr = Replace(tmpPrintStr, "<Location>", DIDLocation) '1242
         tmpPrintStr = Replace(tmpPrintStr, "<MARK>", DIDMark) '1255
         tmpPrintStr = Replace(tmpPrintStr, "<WHID>", DIDInfo.WareHouseID) '1252
         
         If strQty = "RefID" Then
             tmpPrintStr = Replace(tmpPrintStr, "<LINE>", BUDIDShow)
         Else
             tmpPrintStr = Replace(tmpPrintStr, "<LINE>", IIf(IsAnotherBUDID = "Y", AutoDispatchForAnotherBU, BUDIDShow))
         End If
         tmpPrintStr = Replace(tmpPrintStr, "<SIDE>", "")
'                hString = Replace(hString, "<MACHINE>", IIf(DIDInfo.IsGood = "Y", "", "NG"))
         tmpPrintStr = Replace(tmpPrintStr, "<MACHINE>", "NG")
         ''Debug.Print hString
         
        Select Case Trim(tmpPrintStr)
           Case vbNullString
           Case Else
                If OptComp.Value = True Then
                    If blnZebra = True Then                      '(1017)
                        'MSComm.Output = tmpPrintStr
                        For M = 1 To Len(tmpPrintStr) Step 50
                            MSComm.Output = Mid(tmpPrintStr, M, 50)
                            'Debug.Print Mid(hString, m, 50)
                        Next M
                    Else '   (0016)
                        For M = 1 To Len(tmpPrintStr) Step 50
                            MSComm.Output = Mid(tmpPrintStr, M, 50)
                            'Debug.Print Mid(hString, m, 50)
                        Next M
                    End If
                    MSComm.PortOpen = False
             ElseIf OptPrint.Value = True Then
                If blnZebra = True Then
                    Print #lptPort, tmpPrintStr & Chr(13)
                Else '   (0016)
                    For M = 1 To Len(tmpPrintStr) Step 50
                        Print #lptPort, Mid(tmpPrintStr, M, 50)
                        'Debug.Print Mid(hString, m, 50)
                    Next M
                End If
             Else
                Printer.Print tmpPrintStr
                Printer.EndDoc
                Printer.KillDoc
             End If
             
        End Select
        Close #lptPort   '(1003)
        Close #hFile
        Exit Function
errHandler:
        MsgBox Err.Description
        If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
        End If
End Function


Private Function PrintLabelCompPort(blnZebra As Boolean, intCompPort As Integer, sCommString As String) As String
    Dim hFile As Long
    Dim hString As String
    Dim strDID As String, tmpDID As String, strQty As String
    Dim strDay As String
    Dim LabelFile As String
    Dim M As Integer
    Dim strSQL As String
    Dim rsTime As ADODB.Recordset
    
        On Error GoTo errHandler
        '1112
        strSQL = "select getdate()"
        Set rsTime = Conn.Execute(strSQL)
        strDay = Format(rsTime(0), "YYMMDDHHNNSS")    '1101
        
        ''20100423   Denver    DID Return后，如无需求，现不直接退库，打印时DID位置显示CompPN(0071)
        If blnZebra = True Then
'            LabelFile = Settings.AutoDispatchLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
                LabelFile = Settings.DIDLabelGood
                ''strDID = DIDInfo.DID
                If DIDnotToQWMS = "Y" Then
                    strDID = DIDInfo.COMPPN
                Else
                    strDID = DIDInfo.DID
                End If
            Else
                LabelFile = Settings.DIDLabelBad
                strDID = DIDInfo.COMPPN
            End If
        Else
'            LabelFile = Settings.AutoDispatchSatoLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
                LabelFile = Settings.DIDLabelSATOGood
                ''strDID = DIDInfo.DID
                If DIDnotToQWMS = "Y" Then
                    strDID = DIDInfo.COMPPN
                Else
                    strDID = DIDInfo.DID
                End If
            Else
                LabelFile = Settings.DIDLabelSATOBad
                strDID = DIDInfo.COMPPN
            End If
        End If
        
        ''特殊处理(打印RefID Label)
        If DIDInfo.Qty <= -10000 Then
            strQty = "RefID"
        Else
            strQty = CStr(DIDInfo.Qty)
        End If
     
        If (Dir(LabelFile) = vbNullString) Then
            ''''''Added by Jing 2008.01.10  (00003)'''''
            MsgBox ("Can not find Lable file !"), vbCritical
            PrintLabelCompPort = "PRN_FileNoExist"
            Exit Function
        End If
        
        'TxtCompPort   TxtComm
        MSComm.CommPort = intCompPort
        MSComm.Settings = sCommString
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
                    If blnZebra Then
                        tmpDID = Replace(strDID, "^", "><")
                    End If
                    hString = Replace(hString, "<DID_CODE>", tmpDID)
                End If
                'for text ^, must be use ^FH_ and the use _5E (the ascii of ^)
                If InStr(hString, "<DID_TEXT>") > 0 Then
                    ''********************************updated by jing 20071024 (0002) ***********
                    If blnZebra Then
                       tmpDID = Replace(strDID, "^", "_5E")
                    End If
                    hString = Replace(hString, "<DID_TEXT>", tmpDID)
                End If

                hString = Replace(hString, "<UID>", uId)
                
'                hString = Replace(hString, "<DATE>", strDay)
'                hString = Replace(hString, "<QTY>", strQTY)
'                hString = Replace(hString, "<LINE>", PrintData.Line)
'                hString = Replace(hString, "<SIDE>", PrintData.Side)
'                hString = Replace(hString, "<MACHINE>", PrintData.Machine)
                
                hString = Replace(hString, "<DATE>", strDay)
                hString = Replace(hString, "<QTY>", strQty)
                If strQty = "RefID" Then
                    hString = Replace(hString, "<LINE>", BUDIDShow)
                Else
                    hString = Replace(hString, "<LINE>", IIf(IsAnotherBUDID = "Y", AutoDispatchForAnotherBU, BUDIDShow))
                End If
                hString = Replace(hString, "<SIDE>", "")
'                hString = Replace(hString, "<MACHINE>", IIf(DIDInfo.IsGood = "Y", "", "NG"))
                hString = Replace(hString, "<MACHINE>", "NG")
                ''Debug.Print hString
                
               Select Case Trim(hString)
                  Case vbNullString
                  Case Else
'                    MSComm.Output = hString
'                    Debug.Print hString

                    If blnZebra = True Then
                        MSComm.Output = hString
                    Else '   (0016)
                        For M = 1 To Len(hString) Step 50
                            MSComm.Output = Mid(hString, M, 50)
                            'Debug.Print Mid(hString, m, 50)
                        Next M
                    End If
                    
                    
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

Private Function PrintLabelPrintPort(blnZebra As Boolean) As String
    Dim hFile As Long
    Dim hString As String
    Dim strDID As String, tmpDID As String, strQty As String
    Dim FileNum As Integer, lptPort As Integer
    Dim strDay As String
    Dim LabelFile, strLabelFileContent As String
    Dim strPort As String, PrintLabel As String
    Dim M As Integer
    Dim strSQL As String
    Dim rsTime As ADODB.Recordset
    
    On Error GoTo errHandler
        '1112
        strSQL = "select getdate()"
        Set rsTime = Conn.Execute(strSQL)
        strDay = Format(rsTime(0), "YYMMDDHHNNSS") '1101
        
        ''特殊处理(打印RefID Label)
        If DIDInfo.Qty <= -10000 Then
            strQty = "RefID"
        Else
            strQty = CStr(DIDInfo.Qty)
        End If
        
         ''20100423   Denver    DID Return后，如无需求，现不直接退库，打印时DID位置显示CompPN(0071)
        If blnZebra = True Then
'            LabelFile = Settings.AutoDispatchLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
                LabelFile = Settings.DIDLabelGood
                 
                ''strDID = DIDInfo.DID
                If DIDnotToQWMS = "Y" Then
                    strDID = DIDInfo.COMPPN
                Else
                    strDID = DIDInfo.DID
                End If
            Else
                LabelFile = Settings.DIDLabelBad
                strDID = DIDInfo.COMPPN
            End If
        Else
'            LabelFile = Settings.AutoDispatchSatoLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
                LabelFile = Settings.DIDLabelSATOGood
                 
                ''strDID = DIDInfo.DID
                If DIDnotToQWMS = "Y" Then
                    strDID = DIDInfo.COMPPN
                Else
                    strDID = DIDInfo.DID
                End If
            Else
                LabelFile = Settings.DIDLabelSATOBad
                strDID = DIDInfo.COMPPN
            End If
            
        End If
'        strLabelFileContent = funGetTxtFileContent(LabelFile)
        If Dir(LabelFile) = vbNullString Then
            ''''''Added by Jing 2008.01.10  (00003)''''''
            MsgBox ("Can not find Label file !"), vbCritical
            PrintLabel = "PRN_FileNoExist"
            Exit Function
        End If
        lptPort = OpenOutputFile("LPT1")
        If lptPort = 0 Then
            MsgBox "Open print port LPT1 error!"
            Exit Function
        End If
        
        FileNum = FreeFile()
        Open LabelFile For Input As #FileNum
        While Not EOF(FileNum)
            Line Input #FileNum, hString
            hString = Trim(hString)
            tmpDID = Trim(strDID)  '***************add by jeanson 20070814******
            'for Code 128 barcode, the ^ must be tranfer to ><
            If InStr(hString, "<DID_CODE>") > 0 Then
                ''********************************updated by jing 20071024 (0002) ***********
                If blnZebra Then
                    tmpDID = Replace(strDID, "^", "><")
                End If
               hString = Replace(hString, "<DID_CODE>", tmpDID)
            End If
            'for text ^, must be use ^FH_ and the use _5E (the ascii of ^)
            If InStr(hString, "<DID_TEXT>") > 0 Then
                ''********************************updated by jing 20071024 (0002) ***********
                If blnZebra Then
                   tmpDID = Replace(strDID, "^", "_5E")
                End If
               hString = Replace(hString, "<DID_TEXT>", tmpDID)
            End If
            
            hString = Replace(hString, "<UID>", uId)
'                hString = Replace(hString, "<RACKID>", TxtRackID)
            hString = Replace(hString, "<DATE>", strDay)
            hString = Replace(hString, "<QTY>", strQty)
            hString = Replace(hString, "<LINE>", IIf(IsAnotherBUDID = "Y", AutoDispatchForAnotherBU, BUDIDShow))
            hString = Replace(hString, "<SIDE>", "")
            hString = Replace(hString, "<MACHINE>", IIf(DIDInfo.IsGood = "Y", "", "NG"))
            
'            Debug.Print hString
'            Print #lptPort, hString & Chr(13)
            If blnZebra = True Then
                Print #lptPort, hString & Chr(13)
            Else '   (0016)
                For M = 1 To Len(hString) Step 50
                    Print #lptPort, Mid(hString, M, 50)
                    'Debug.Print Mid(hString, m, 50)
                Next M
            End If
         
        Wend
'        Open strPort For Output As #FileNum
'        Print #FileNum, strLabelFileContent
        Close #FileNum
        Close #lptPort
        Exit Function
errHandler:
     MsgBox Err.Description
End Function

Private Function ChkDispatchIsOK(sNewDID As String, sOldDID As String, Msg As String) As Boolean

    ChkDispatchIsOK = False
    sSql = "exec XL_DIDChk_ReturnDisp " & sq(sNewDID) & "," & sq(sOldDID) & "," & sq(Msg) & "," & sq(Trim(IsAnotherBUDID))
    Set Rst = Conn.Execute(sSql)
    If Rst.EOF = False Then
        If Rst("ReSult") = 0 Then
            ChkDispatchIsOK = True
        End If
    End If

End Function



Private Function XL_ChkAnotherBUDID(sDID As String) As Boolean
On Error GoTo errHandler

    XL_ChkAnotherBUDID = False
    
    'QMS             Denver         2011/01/12     Unify SP:XL_ChkAnotherBUDID in CallBack/ReturnDID   (1048)
    sSql = "exec XL_ChkAnotherBUDID " & sq(sDID) & "," & sq(IsAnotherBUDID) & "," & sq(Trim(Factory))
    Set Rst = Conn.Execute(sSql)
    
    If Rst("Result") <> 0 Then
        LblMessage.Caption = Rst("Description")
        Exit Function
    Else
        LblMessage.BackColor = &H80FF80
        Set Rst = Rst.NextRecordset
       
        If Rst.EOF = False Then
            TxtDIDTotalQty = Trim(Rst!TotalQty)
            TxtDIDReturnedQty = Trim(Rst!ReturnQty)
            TxtCompPN = Trim(Rst!COMPPN)
            IsAnotherBUDID = Trim(Rst!IsAnotherBUDID)
        End If
    End If
    
    XL_ChkAnotherBUDID = True

    Exit Function
errHandler:
     MsgBox Err.Description
     
End Function



