VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDIDCallBack_New 
   Caption         =   "DID CallBack[20150203]"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   14250
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sstabDID 
      Height          =   3105
      Left            =   90
      TabIndex        =   12
      Top             =   1170
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   5477
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "DID Not CallBack"
      TabPicture(0)   =   "frmDIDCallBack_New.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "gridDIDNotCall"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "DID CallBacked"
      TabPicture(1)   =   "frmDIDCallBack_New.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "gridDIDCalled"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "DID Information"
      TabPicture(2)   =   "frmDIDCallBack_New.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "gridDIDInfo"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "WO Information"
      TabPicture(3)   =   "frmDIDCallBack_New.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "gridWOGroup"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "DID to WareHouse Top 1000"
      TabPicture(4)   =   "frmDIDCallBack_New.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "gridDIDtoWH"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin MSDataGridLib.DataGrid gridDIDInfo 
         Height          =   2325
         Left            =   -74910
         TabIndex        =   15
         Top             =   720
         Width           =   13905
         _ExtentX        =   24527
         _ExtentY        =   4101
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
      Begin MSDataGridLib.DataGrid gridDIDNotCall 
         Height          =   2325
         Left            =   -74880
         TabIndex        =   13
         Top             =   720
         Width           =   13875
         _ExtentX        =   24474
         _ExtentY        =   4101
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
      Begin MSDataGridLib.DataGrid gridDIDCalled 
         Height          =   2325
         Left            =   -74910
         TabIndex        =   14
         Top             =   720
         Width           =   13875
         _ExtentX        =   24474
         _ExtentY        =   4101
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
      Begin MSDataGridLib.DataGrid gridWOGroup 
         Height          =   2355
         Left            =   -74910
         TabIndex        =   29
         Top             =   690
         Width           =   13905
         _ExtentX        =   24527
         _ExtentY        =   4154
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
      Begin MSDataGridLib.DataGrid gridDIDtoWH 
         Height          =   2355
         Left            =   120
         TabIndex        =   67
         Top             =   690
         Width           =   13905
         _ExtentX        =   24527
         _ExtentY        =   4154
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
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      Height          =   4275
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   14145
      Begin VB.Frame Frame1 
         Height          =   915
         Left            =   9660
         TabIndex        =   43
         Top             =   180
         Width           =   1245
         Begin VB.OptionButton OptRelease 
            Caption         =   "Release"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   210
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optGroup 
            Caption         =   "Group"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   510
            Width           =   1065
         End
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "Filter DID by WO"
         Height          =   285
         Left            =   7200
         TabIndex        =   16
         Top             =   780
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.ComboBox CboLine 
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
         Left            =   8250
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   270
         Width           =   1275
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
         Left            =   4950
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   690
         Width           =   2175
      End
      Begin VB.CommandButton CmdQuery 
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   11010
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   1455
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
         Left            =   1350
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   660
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1350
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
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
         Format          =   165281795
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   4950
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   270
         Width           =   2175
         _ExtentX        =   3836
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
         Format          =   165281795
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
         Left            =   3630
         TabIndex        =   11
         Top             =   270
         Width           =   1275
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
         Left            =   3600
         TabIndex        =   10
         Top             =   690
         Width           =   1305
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
         Left            =   7200
         TabIndex        =   9
         Top             =   270
         Width           =   945
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
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   1185
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
         TabIndex        =   7
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Frame FraReturnDID 
      BackColor       =   &H80000013&
      Caption         =   "CallBack DID"
      Height          =   5055
      Left            =   30
      TabIndex        =   17
      Top             =   4350
      Width           =   14145
      Begin VB.Frame fraPrinter 
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
         Height          =   825
         Left            =   30
         TabIndex        =   55
         Top             =   1770
         Width           =   14025
         Begin VB.Frame Frame4 
            Height          =   615
            Left            =   60
            TabIndex        =   62
            Top             =   210
            Width           =   3105
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
               Left            =   60
               TabIndex        =   64
               Top             =   210
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
               Left            =   1500
               TabIndex        =   63
               Top             =   240
               Width           =   1455
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
            Left            =   12480
            Picture         =   "frmDIDCallBack_New.frx":008C
            Style           =   1  'Graphical
            TabIndex        =   61
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
            Left            =   10920
            TabIndex        =   60
            Text            =   "9600,N,8,1"
            Top             =   300
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
            Left            =   9240
            TabIndex        =   59
            Text            =   "1"
            Top             =   300
            Width           =   495
         End
         Begin VB.Frame Frame5 
            Height          =   615
            Left            =   3210
            TabIndex        =   56
            Top             =   210
            Width           =   4395
            Begin VB.OptionButton optNetwork 
               Caption         =   "Network"
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
               Left            =   2760
               TabIndex        =   68
               Top             =   240
               Value           =   -1  'True
               Width           =   1455
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
               Left            =   1440
               TabIndex        =   58
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
               Left            =   60
               TabIndex        =   57
               Top             =   240
               Width           =   1365
            End
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
            Index           =   0
            Left            =   9840
            TabIndex        =   66
            Top             =   300
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
            Index           =   2
            Left            =   7800
            TabIndex        =   65
            Top             =   300
            Width           =   1335
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
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1185
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
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1185
      End
      Begin VB.Frame Frame2 
         Height          =   585
         Left            =   4170
         TabIndex        =   50
         Top             =   1200
         Width           =   1815
         Begin VB.OptionButton optGoodMaterial 
            Caption         =   "Good"
            Height          =   345
            Left            =   120
            TabIndex        =   52
            Top             =   150
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton optBadMaterial 
            Caption         =   "Bad"
            Height          =   285
            Left            =   1080
            TabIndex        =   51
            Top             =   180
            Width           =   615
         End
      End
      Begin TabDlg.SSTab sstabDispatched 
         Height          =   2295
         Left            =   30
         TabIndex        =   47
         Top             =   2700
         Width           =   14115
         _ExtentX        =   24897
         _ExtentY        =   4048
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "DID Dispatched by PCB"
         TabPicture(0)   =   "frmDIDCallBack_New.frx":0396
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "gridDIDDispatched"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Similar DID dispatched by PCB"
         TabPicture(1)   =   "frmDIDCallBack_New.frx":03B2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "gridSimilarDIDByPCB"
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid gridDIDDispatched 
            Height          =   1755
            Left            =   120
            TabIndex        =   48
            Top             =   420
            Width           =   13905
            _ExtentX        =   24527
            _ExtentY        =   3096
            _Version        =   393216
            DefColWidth     =   100
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
         Begin MSDataGridLib.DataGrid gridSimilarDIDByPCB 
            Height          =   1755
            Left            =   -74910
            TabIndex        =   49
            Top             =   420
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   3096
            _Version        =   393216
            DefColWidth     =   100
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
      Begin VB.TextBox txtRemainQty 
         Height          =   375
         Left            =   9120
         TabIndex        =   46
         Top             =   750
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmdADD 
         BackColor       =   &H00C0C0C0&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdADDALL 
         BackColor       =   &H00C0C0C0&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   770
         Width           =   495
      End
      Begin VB.CommandButton cmdDEL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1060
         Width           =   495
      End
      Begin VB.CommandButton cmdDELALL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   12120
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1350
         Width           =   495
      End
      Begin VB.Frame fraOPt 
         Height          =   585
         Left            =   0
         TabIndex        =   33
         Top             =   1200
         Width           =   4155
         Begin VB.OptionButton optCallAll 
            Caption         =   "Call back All"
            Height          =   285
            Left            =   2910
            TabIndex        =   42
            Top             =   180
            Width           =   1185
         End
         Begin VB.OptionButton optRatebyPCB 
            Caption         =   "Call by PCB"
            Height          =   285
            Left            =   1770
            TabIndex        =   35
            Top             =   180
            Width           =   1575
         End
         Begin VB.OptionButton optRatebySelWO 
            Caption         =   "Call by selected WO"
            Height          =   345
            Left            =   30
            TabIndex        =   34
            Top             =   150
            Value           =   -1  'True
            Width           =   2085
         End
      End
      Begin VB.ListBox lstCallBackWO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   12630
         TabIndex        =   32
         Top             =   420
         Width           =   1425
      End
      Begin VB.ListBox lstAvailableWO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   10530
         TabIndex        =   31
         Top             =   420
         Width           =   1545
      End
      Begin VB.TextBox TxtCompPN 
         Enabled         =   0   'False
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
         Left            =   7140
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   1965
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1155
      End
      Begin VB.TextBox TxtDIDReturnedQty 
         Enabled         =   0   'False
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
         Left            =   4920
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtDIDTotalQty 
         Enabled         =   0   'False
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
         Left            =   1590
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtReturnQty 
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
         Left            =   7530
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   690
         Width           =   1575
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
         Height          =   480
         Left            =   1590
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   690
         Width           =   4305
      End
      Begin VB.Label lblCallBackWO 
         Caption         =   "CallBack WO"
         Height          =   255
         Left            =   12630
         TabIndex        =   37
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label lblAvailableWO 
         Caption         =   "Available WO"
         Height          =   255
         Left            =   10530
         TabIndex        =   36
         Top             =   150
         Width           =   1095
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
         Index           =   3
         Left            =   6060
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DID CallBacked Qty"
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
         Left            =   2730
         TabIndex        =   27
         Top             =   240
         Width           =   2175
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
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Call Back Qty"
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
         Left            =   5940
         TabIndex        =   25
         Top             =   750
         Width           =   1575
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
         TabIndex        =   24
         Top             =   743
         Width           =   1695
      End
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   30
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label LblMessage 
      BackColor       =   &H80000000&
      Height          =   405
      Left            =   60
      TabIndex        =   30
      Top             =   9420
      Width           =   14115
   End
End
Attribute VB_Name = "frmDIDCallBack_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/**********************************************************************************
'**文 件 名: DID CallBack.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: Jeanson
'**日    期: 2007.12.18
'**描    述: QSMS DID CallBack
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**Jeanson      2007.12.18     add Good/Bad Material info and Print a new label for XL (0001)
'**Denver       2007.12.20     modify callback error info (0002)
'**Denver       2007.12.27     modify callback Print label and get ReferenceID (0003)
'**Denver       2008.01.10     Add error alarm when can not find label file     (0004)
'**Giant        2008.03.12     mared get date from local compture     (0005)

'*Denver        2009.08.14     NB2&NB3分家，但是用同一个QSMS可作NB2NB3 的CallBack

'***********************************************************************************
Private rstDID As New ADODB.Recordset
Private rstDIDCalled As New ADODB.Recordset
Private rstDIDNotCall As New ADODB.Recordset
Private rstWOGroup As New ADODB.Recordset
Private rstDIDSimilar As New ADODB.Recordset
Private rstDIDtoWH As New ADODB.Recordset
Private Rst As New ADODB.Recordset
Private sSql As String
Private mWOStatus As String   'Y is Input, N is Not Input

Private IsAnotherBUDID As String


Private Sub CboGroupID_Click()
    Call GetGroupWO(CboGroupID)
    
End Sub

Private Sub CboWo_Click()
    Screen.MousePointer = vbHourglass
    sSql = "Select GroupID from QSMS_WOGroup where Work_Order='" & Trim(cboWO) & "'"
    Set Rst = Conn.Execute(sSql)
        If Rst.EOF Then
            MsgBox "Can not find the GroupID for the Work Order:" & Trim(cboWO)
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            If ChkGroupClosed(Trim(Rst!GroupID)) = True Then
                MsgBox "The Group has been closed,can not return DID"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    
    
'    Call GetSBWO(TxtWO)
    Call GetWOGroupinfo(cboWO)
    sSql = "Exec QSMSDIDCallBack '" & Trim(cboWO) & "'"
    Conn.Execute sSql
'    wostr = GetWoArray
    Call GetDIDInfoCallBack(Trim(cboWO))
    TxtDID = ""
    TxtReturnQty = ""
    Screen.MousePointer = vbDefault
    sstabDID.Tab = 0

End Sub

Private Sub CmdADD_Click()
    Dim I As Integer
    If lstAvailableWO.ListCount <= 0 Then Exit Sub
    If lstAvailableWO.ListIndex < 0 Then Exit Sub
    I = lstAvailableWO.ListIndex
    lstCallBackWO.AddItem Trim(lstAvailableWO.Text)
    lstAvailableWO.RemoveItem I
    If lstAvailableWO.ListCount > 0 Then
        If lstAvailableWO.ListCount - 1 >= I Then
            lstAvailableWO.ListIndex = I
        Else
            lstAvailableWO.ListIndex = lstAvailableWO.ListCount - 1
        End If
    End If
End Sub

Private Sub cmdADDALL_Click()
    If lstAvailableWO.ListCount <= 0 Then Exit Sub
    
    Do While lstAvailableWO.ListCount > 0
      lstAvailableWO.ListIndex = 0
      lstCallBackWO.AddItem Trim(lstAvailableWO.Text)
      lstAvailableWO.RemoveItem 0
    Loop
End Sub

Private Sub CmdCommSave_Click()
    SaveSetting "SMT", "QSMS", "CommPort", TxtCompPort
    SaveSetting "SMT", "QSMS", "Comm", TxtComm
End Sub

Private Sub cmdDel_Click()
    Dim I As Integer
    If lstCallBackWO.ListCount <= 0 Then Exit Sub
    If lstCallBackWO.ListIndex < 0 Then Exit Sub
    I = lstCallBackWO.ListIndex
    lstAvailableWO.AddItem Trim(lstCallBackWO.Text)
    lstCallBackWO.RemoveItem I
    If lstCallBackWO.ListCount > 0 Then
        If lstCallBackWO.ListCount - 1 >= I Then
            lstCallBackWO.ListIndex = I
        Else
            lstCallBackWO.ListIndex = lstCallBackWO.ListCount - 1
        End If
    End If
End Sub

Private Sub cmdDELALL_Click()
    If lstCallBackWO.ListCount <= 0 Then Exit Sub
    
    Do While lstCallBackWO.ListCount > 0
      lstCallBackWO.ListIndex = 0
      lstAvailableWO.AddItem Trim(lstCallBackWO.Text)
      lstCallBackWO.RemoveItem 0
    Loop
End Sub

Private Sub cmdGetRefID_Click()
    Dim sCurrRefID As String
    Dim sMsg As String
   
   
    sSql = "exec XL_DIDGetRefID @Type='CallBack', @IsGood=" & sq(IIf(optGoodMaterial.Value = True, "Y", "N")) & ",@UserName=" & sq(g_userName) & ",@Factory=" & sq(Trim(Factory)) & ",@IsAnotherBUDID=" & sq(Trim(IsAnotherBUDID))
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
            .compPN = sCurrRefID
            .Qty = -10000
            .IsGood = IIf(optGoodMaterial.Value = True, "Y", "N")
            .DIDType = ""
        End With
        Call DIDPrintLabel(OptZebra.Value, CInt(Trim(TxtCompPort)), Trim(TxtComm))
        
        ''Check Stock qty By RefID
        sSql = "exec XL_DIDChkStockByRefID @Type='Auto',@RefID=" & sq(sCurrRefID) & ",@UserName=" & sq(g_userName)
        Set Rst = Conn.Execute(sSql)
        If Rst.EOF = False Then
            If Rst("Result") <> 0 Then
                MsgBox Rst("Description"), vbExclamation, "Prompt"
                Exit Sub
            End If
            
            frmDIDChkStock.FuncType = "AutoChk"
            Set Rst = Rst.NextRecordset
            Set frmDIDChkStock.rstCompPN = Rst
            frmDIDChkStock.lblMsg = sMsg
            frmDIDChkStock.Show 1
            
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

Private Sub cmdReprint_Click()
    ''check printer
    Dim IsByDIDInput As Boolean
    
    IsByDIDInput = False
    
    If Trim(TxtCompPort) = "" Or Trim(TxtComm) = "" Then
        MsgBox "Printer have not set!!", vbInformation
        Exit Sub
    End If

    If Trim(TxtDID) = "" Then
        LblMessage = "Please select or Input DID to reprint!!"
        Exit Sub
    End If
    
    
    '''2008/03/24   Denver  Modify for Reprint DID --(0001)
    If gridDIDtoWH.row >= 0 Then
        If Trim(TxtDID) = Trim(gridDIDtoWH.Columns(1).Text) Then
            With DIDInfo
                .DID = Trim(gridDIDtoWH.Columns(1).Text)
                .compPN = Trim(gridDIDtoWH.Columns(2).Text)
                .Qty = Trim(gridDIDtoWH.Columns(3).Text)
                .IsGood = Trim(gridDIDtoWH.Columns(10).Text)
                If ChkPrintDIDType = "Y" Then   ''1142
                    .DIDType = Trim(gridDIDtoWH.Columns(14).Text)
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
        sSql = "exec XL_DIDGetToWHInfo 'CallBack'," & sq(Trim(TxtDID)) & ",@Factory=" & sq(Trim(Factory))
        Set Rst = Conn.Execute(sSql)
        If Rst.EOF = True Then
            LblMessage = "There is no DID:" & sq(Trim(TxtDID)) + " !!"
            Exit Sub
        Else
            With DIDInfo
                .DID = Trim(Rst("DID") & "")
                .compPN = Trim(Rst("CompPN") & "")
                .Qty = Rst("Qty")
                .IsGood = Trim(Rst("IsGood") & "")
                If ChkPrintDIDType = "Y" Then    ''1142
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

Private Sub cmdSave_Click()
On Error GoTo Err_Handler
    Dim sSelWo As String
    Dim sDID As String
    Dim intReturnQty As Long  ''(1094)Integer to long
    
    
    
    cmdSave.Enabled = False
    If ChkErr = False Then
       GoTo Normal_Eixt
    End If
    
    ''20081230   Denver  Chedk OK后，运用中间变量保存，以防运行中修改
    sDID = Trim(TxtDID)
    intReturnQty = Trim(TxtReturnQty)
    
    If optRatebySelWO.Value = True Then
        sSelWo = GetSelWO(lstCallBackWO)
        If sSelWo = "" Then Exit Sub
        If chkCallBackQty(Trim(sDID), sSelWo, Val(Trim(intReturnQty)), Val(Trim(txtRemainQty))) = False Then Exit Sub

        sSql = "Exec DIDCallBackByType @CallType='CallBySelWO', @WO=" & sq(sSelWo) & ",@CompPN=" & sq(Trim(TxtCompPN)) & ",@DID=" & sq(Trim(sDID)) & ",@ReturnQty=" & Trim(intReturnQty) & ",@UserName=" & sq(g_userName) & ",@IsGood=" & sq(IIf(optGoodMaterial.Value = True, "Y", "N")) & ",@IsAnotherBUDID =" & sq(IsAnotherBUDID)
    ElseIf optRatebyPCB.Value = True Then
        
        sSql = "Exec DIDCallBackByType @CallType='CallbyPCB', @WO=" & sq(Trim(cboWO)) & ",@CompPN=" & sq(Trim(TxtCompPN)) & ",@DID=" & sq(Trim(sDID)) & ",@ReturnQty=" & Trim(intReturnQty) & ",@UserName=" & sq(g_userName) & ",@IsGood=" & sq(IIf(optGoodMaterial.Value = True, "Y", "N")) & ",@IsAnotherBUDID =" & sq(IsAnotherBUDID)

    Else  'CallBack All
        If MsgBox("DID:" & sDID & ",Total = CallBackQty = " & TxtDIDTotalQty & ",Are you sure CallBack All", vbExclamation + vbYesNo, "Prompt") = vbNo Then
            GoTo Normal_Eixt
        End If

        sSql = "Exec DIDCallBackByType @CallType='CallAll', @WO=" & sq(Trim(cboWO)) & ",@CompPN=" & sq(Trim(TxtCompPN)) & ",@DID=" & sq(Trim(sDID)) & ",@ReturnQty=" & Trim(intReturnQty) & ",@UserName=" & sq(g_userName) & ",@IsGood=" & sq(IIf(optGoodMaterial.Value = True, "Y", "N")) & ",@IsAnotherBUDID =" & sq(IsAnotherBUDID)
    End If
    
    
'    If rst.State Then rst.Close
    Set Rst = Conn.Execute(sSql)
    If Rst.EOF = False Then
        LblMessage.Caption = Rst("Description")
        If Rst("Result") = 0 Then
            LblMessage.BackColor = &H80FF80
            ''2007.12.27 Denver  modify callback Print label and get ReferenceID (0003)
            If PrtCallBKandReturn = "Y" Then
                sSql = "exec XL_DIDGetNewID @Type='CallBack',@DID=" & sq(Trim(sDID)) & ",@IsGood=" & sq(IIf(optGoodMaterial.Value = True, "Y", "N")) & ", @ReturnQty=" & Trim(intReturnQty) & ",@UserName=" & sq(g_userName) & ",@Factory=" & sq(Trim(Factory)) & ",@IsAnotherBUDID =" & sq(IsAnotherBUDID)
                Set Rst = Conn.Execute(sSql)
                If Rst("Result") <> 0 Then
                    LblMessage.Caption = Rst("Description")
                    
                Else
    '                LblMessage.BackColor = &H80FF80
                    Set Rst = Rst.NextRecordset
                    'PN/Qty/PU/NG/UID/Date   (bad)   'DID/Qty/PU/UID/Date (Good)
                    If Rst.EOF = True Then
                        LblMessage.Caption = "Get DID information fail,print DID fail!!"
                        GoTo Normal_Eixt
                    End If
                    With DIDInfo
                        .DID = Trim(Rst("DID") & "")
                        .compPN = Trim(Rst("CompPN") & "")
                        .Qty = Rst("Qty")
                        .IsGood = Trim(Rst("IsGood") & "")
                        If ChkPrintDIDType = "Y" Then   ''1142
                            .DIDType = Trim(Rst("DIDType"))
                        Else
                            .DIDType = ""
                        End If
                    End With
                    
                    Call DIDPrintLabel(OptZebra.Value, CInt(Trim(TxtCompPort)), Trim(TxtComm))
                     
                End If
            End If
        End If
    End If
    
    Call GetDIDInfoCallBack(Trim(cboWO))

    
Normal_Eixt:
    cmdSave.Enabled = True
    TxtDID.Text = ""
    TxtReturnQty.Text = ""
    TxtDIDTotalQty = ""
    TxtCompPN = ""
    TxtDIDReturnedQty = ""
    TxtDID.SetFocus
    Exit Sub
    
    
Err_Handler:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "Prompt"
    cmdSave.Enabled = True
End Sub

Private Sub Form_Load()
'*************Marked by Giant 080312***********()
'    sSql = "select getdate()"
'    Set rst = Conn.Execute(sSql)
'    If Not rst.EOF Then
'        Date = rst(0)
'        Time = rst(0)
'    End If
'    dtpSDate.Value = Format(Date, "YYYY/MM/DD")
'    dtpEDate.Value = Format(Date, "YYYY/MM/DD")
'*************Marked by Giant 080312***********
    dtpSDate = Date 'add by Giant 080312
    dtpEDate = Date
    Call GetLine
    sstabDID.Tab = 0
    sstabDispatched.Tab = 0
    
'    TxtCompPort = GetSetting("SMT", "QSMS", "CommPort", "1")
'    TxtComm = GetSetting("SMT", "QSMS", "Comm", "9600,N,8,1")
    
    '20101115 Maggie Save Printer setting in local Registry
    Call GetPrinterSetting(frmDIDCallBack_New)
    
    ''20100507    Denver     好坏料不让user 选择
    optGoodMaterial.Enabled = False
    optBadMaterial.Enabled = False
    
    If PrtCallBKandReturn <> "Y" Then
        cmdGetRefID.Visible = False
        CmdReprint.Visible = False
        FraPrinter.Visible = False
    End If
    
End Sub
 
 
Private Function GetGroupWO(ByVal GroupID As String)
    sSql = "select Work_Order,ClosedFlag from QSMS_WOGroup  where GroupID= '" & GroupID & "'"
    Set Rst = Conn.Execute(sSql)

    cboWO.Clear
    Do While Rst.EOF = False
        cboWO.AddItem Trim(Rst!Work_Order)
        Rst.MoveNext
    Loop
End Function

Private Function GetLine()
    sSql = "select distinct Line from QSMS_woGroup Order by line"
    Set Rst = Conn.Execute(sSql)
    CboLine.Clear
    Do While Rst.EOF = False
        CboLine.AddItem Trim(Rst!Line)
        Rst.MoveNext
    Loop
End Function

Private Function GetGroupID()
    Dim str As String
    Dim sSDate As String
    Dim sEDate As String
'    Dim GroupIDHead As String
'    Dim i As Long
'    Dim Rs As ADODB.Recordset
    sSDate = Format(dtpSDate, "YYYYMMDD")
    sEDate = Format(dtpEDate, "YYYYMMDD")
    
'    sSDate = Replace(Replace(dtpSDate, "/", ""), "-", "")
'    sEDate = Replace(Replace(dtpEDate, "/", ""), "-", "")
    
    If OptRelease.Value = True Then
       sSql = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & sSDate & "' and '" & sEDate & "' and line='" & CboLine & "' and closedflag<>'Y'"
    Else
        sSql = "select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '" & sSDate & "' and '" & sEDate & "' and line='" & CboLine & "' and closedflag<>'Y'"
    End If
    
    Set Rst = Conn.Execute(sSql)
    CboGroupID.Clear
    If Rst.EOF Then MsgBox "No data"
    Do While Rst.EOF = False
          CboGroupID.AddItem Trim(Rst!GroupID)
          Rst.MoveNext
    Loop
End Function

Private Function GetWOGroupinfo(ByVal WO As String)
    '分BU 和 CostBU 获得 是否投入
'    sSql = "select WO,PrdStatus=case when B.workorder is null then 'N' else 'Y' end,PN,MB_Rev,Line,Qty,CombineQty,Trans_Date,WO_Type,[Group],Pilot,BuildType" _
'        & Chr(13) & " from Sap_Wo_List A left join (Select distinct WorkOrder from smt_sp where WorkOrder in(select WO from dbo.GetWOGroup('" & Trim(WO) & "'))) B on" _
'        & Chr(13) & " a.WO=B.WorkOrder Where WO in(select WO from dbo.GetWOGroup('" & Trim(WO) & "')) order by WO"

    sSql = "Exec GetWOPCBStatus @WO=" & sq(WO)

    Set rstWOGroup = Conn.Execute(sSql)
    Set gridWOGroup.DataSource = rstWOGroup
    gridWOGroup.Refresh
    
    
    
End Function

Private Function GetDIDInfoCallBack(ByVal WO As String)
    Dim str As String
    Dim rs As ADODB.Recordset
    
    sSql = "select DID,CompPN,TotalQty,ReturnQty,Transdatetime,UID from QSMS_DIDCallBack where Work_Order in(Select WO from dbo.GetWOGroup(" & WO & ")) and ReturnFlag='Y' order by DID"
    If rstDIDCalled.State Then rstDIDCalled.Close
    rstDIDCalled.CursorLocation = adUseClient
    Set rstDIDCalled = Conn.Execute(sSql)
    Set gridDIDCalled.DataSource = rstDIDCalled
    gridDIDCalled.Caption = "(CallBacked DID)  Total: " & rstDIDCalled.RecordCount
    gridDIDCalled.Refresh
'    sSql = "select DID,CompPN,TotalQty,ReturnQty from QSMS_GroupDID where GroupID='" & GroupID & "' and ReturnFlag<>'Y' Order by DID"
    sSql = "select DID,CompPN,TotalQty,ReturnQty from QSMS_DIDCallBack where Work_Order in(Select WO from dbo.GetWOGroup(" & WO & ")) and ReturnFlag<>'Y' Order by DID"
    If rstDIDNotCall.State Then rstDIDNotCall.Close
    rstDIDNotCall.CursorLocation = adUseClient
    Set rstDIDNotCall = Conn.Execute(sSql)
    Set gridDIDNotCall.DataSource = rstDIDNotCall
    gridDIDNotCall.Caption = "(Not CallBack DID ) Total: " & rstDIDNotCall.RecordCount
    gridDIDNotCall.Refresh
    
    sSql = "select A.DID,A.CompPN,A.Qty as TotalQty,A.RemainQty,A.RealQty,A.VendorCode,A.LotCode,A.UID,A.Transdatetime,A.UsedFlag,A.Inheritflag from QSMS_DID A,QSMS_DIDCallBack B where A.DID=B.DID and B.Work_Order in(Select WO from dbo.GetWOGroup(" & WO & ")) and B.ReturnFlag<>'Y' Order by A.DID"
    If rstDID.State Then rstDID.Close
    rstDID.CursorLocation = adUseClient
    Set rstDID = Conn.Execute(sSql)
    Set gridDIDInfo.DataSource = rstDID
    gridDIDInfo.Refresh
    gridDIDInfo.Caption = "(DID Information) Total: " & rstDID.RecordCount
    
    
    ''2007.12.27 Denver  modify callback Print label and get ReferenceID (0003)
    ''DID to WareHouse top 1000
    If PrtCallBKandReturn = "Y" Then
        sSql = "exec XL_DIDGetToWHInfo 'CallBack',''," & sq(Trim(Factory))
        If rstDIDtoWH.State Then rstDIDtoWH.Close
        rstDIDtoWH.CursorLocation = adUseClient
        Set rstDIDtoWH = Conn.Execute(sSql)
        Set gridDIDtoWH.DataSource = rstDIDtoWH
        gridDIDtoWH.Refresh
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set rstDID = Nothing
    Set rstDIDCalled = Nothing
    Set rstDIDNotCall = Nothing
    Set rstWOGroup = Nothing
    Set Rst = Nothing
    Set rstDIDtoWH = Nothing
    
End Sub

Private Sub gridDIDNotCall_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With gridDIDNotCall
        If Not IsNumeric(LastRow) Then Exit Sub
        If .row <= 0 Then Exit Sub
        
        lstAvailableWO.Clear
        lstCallBackWO.Clear
        Set gridDIDDispatched.DataSource = Nothing
        gridDIDDispatched.Refresh
        
        TxtDID = Trim(.Columns(0).Text)
        TxtDID.SetFocus
        Call txtDID_Click
'        Call TxtDID_KeyPress(13)
    End With
'    If lstAvailableWO.ListCount = 1 Then
'        optRatebySelWO.Value = True
'        Call cmdADDALL_Click
'        TxtReturnQty.SetFocus
'
'    End If
        
End Sub

Private Sub gridDIDtoWH_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With gridDIDtoWH
        If Not IsNumeric(LastRow) Then Exit Sub
        If .row <= 0 Then Exit Sub
        
        lstAvailableWO.Clear
        lstCallBackWO.Clear
        Set gridDIDDispatched.DataSource = Nothing
        gridDIDDispatched.Refresh
        
        TxtDID = Trim(.Columns(1).Text)
        TxtDID.SetFocus
      End With
         
End Sub

Private Sub optCallAll_Click()
    Call setCtrlSelWOStatus(False)
End Sub

Private Sub optRatebyPCB_Click()
    If lstAvailableWO.ListCount + lstCallBackWO.ListCount = 1 Then
        MsgBox "This PCB has only one WO dispathed!!", vbExclamation, "Prompt"
        optRatebySelWO.Value = True
    Else
        Call setCtrlSelWOStatus(False)
    End If
End Sub

Private Sub optRatebySelWO_Click()
    Call setCtrlSelWOStatus(True)
End Sub


Private Sub txtDID_Click()
    SendKeys "{home}+{end}"
End Sub

Private Sub TxtDID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(TxtDID) <> "" Then
        Dim blnDIDRemain As Boolean
        
        ''20080324 Denver    for DID reprint
        If Mid(Right(Trim(TxtDID), 3), 1, 1) = "C" Then Exit Sub
        
        
        ''--*2009.08.14    Denver    NB2&NB3分家，但是用同一个QSMS可作NB2&NB3 的CallBack 退料
        IsAnotherBUDID = "N"
        If AutoDispatchForAnotherBU <> "" Then
'            ''CompPN 长度至少大于10,如果为True,根据DID生成规则,应该属于另一BU的.
'            If InStr(UCase(Trim(TxtDID)), "-" & UCase(DIDHead)) < 10 Then
'                IsAnotherBUDID = "Y"
'
'                If XL_ChkAnotherBUDID(UCase(Trim(TxtDID))) = False Then
'                    TxtDID.Text = ""
'                    TxtDID.SetFocus
'                    Exit Sub
'
'                End If
'
'                Exit Sub
'            End If

             ''20090920   Denver    因为对于未导入XL 的CompPN DID 都是在NB3 注册，然后如有需求，再传至NB2,故改为不用DIDHead判别
            If XL_ChkAnotherBUDID(UCase(Trim(TxtDID))) = False Then
                TxtDID.Text = ""
                TxtDID.SetFocus
                Exit Sub

            End If
            
            If IsAnotherBUDID = "Y" Then
                TxtReturnQty.SetFocus
                Exit Sub
            End If

        End If
        
     
        If ChkDIDBelongToPCB(Trim(cboWO), Trim(TxtDID)) = False Then
           Exit Sub
        End If
        
        Call GetDIDInfo(Trim(TxtDID), Trim(cboWO))
        If TxtCompPN = "" Then
            Exit Sub
        End If
        
        sSql = "Exec DIDSimilarDispByPCB @WO=" & sq(Trim(cboWO)) & ", @DID=" & sq(Trim(TxtDID))
        If rstDIDSimilar.State Then rstDIDSimilar.Close
        rstDIDSimilar.CursorLocation = adUseClient
        Set rstDIDSimilar = Conn.Execute(sSql)
        
        If rstDIDSimilar.EOF = False Then
            If UCase(Rst.Fields(0).Name) <> "RESULT" Then
                If rstDIDSimilar("RemainQty") > 0 And rstDIDSimilar("DID") <> Trim(TxtDID) Then
                    blnDIDRemain = True
                    sstabDispatched.Tab = 1
                    MsgBox "There are similar DID for your refrence!! ", vbExclamation, "Prompt"
                End If
            End If
        End If
        Set gridSimilarDIDByPCB.DataSource = rstDIDSimilar
        If blnDIDRemain = False Then
            sstabDispatched.Tab = 0
        End If
        
        sSql = "select Work_Order,WoQty,NeedQty,DID,TotalQty,DIDQty,JobPN,JobGroup,Machine,CompPN,Slot,LR,Side,BaseQty,GroupID,Line,VendorCode,DateCode,UID,TransDateTime,DeletedFlag,Inherit_WO from QSMS_Dispatch where work_order in (Select WO from dbo.GetWOGroup('" & Trim(cboWO) & "')) and DID='" & Trim(TxtDID) & "'"
        If Rst.State Then Rst.Close
        Rst.CursorLocation = adUseClient
        Set Rst = Conn.Execute(sSql)
        lstAvailableWO.Clear
        lstCallBackWO.Clear
        
        Dim PreWO As String
        PreWO = ""
        Do While Rst.EOF = False
            If PreWO = "" Or PreWO <> Trim(Rst!Work_Order) Then
                lstAvailableWO.AddItem Trim(Rst!Work_Order)
                PreWO = Rst!Work_Order
            End If
            Rst.MoveNext
        Loop
'        rst.MoveFirst
        Set gridDIDDispatched.DataSource = Rst
        TxtReturnQty = ""
        If lstAvailableWO.ListCount = 1 Then
            optRatebySelWO.Value = True
            Call cmdADDALL_Click
            TxtReturnQty.SetFocus
            Call TxtReturnQty_Click
        End If
        
    End If
End Sub

Private Sub setCtrlSelWOStatus(blnEnable As Boolean)
    lstAvailableWO.Enabled = blnEnable
    lstCallBackWO.Enabled = blnEnable
    CmdADD.Enabled = blnEnable
    cmdADDALL.Enabled = blnEnable
    cmdDel.Enabled = blnEnable
    cmdDELALL.Enabled = blnEnable
    
End Sub

Private Function GetDIDInfo(sDID As String, sWO As String)
    
'    sSql = "select A.DID,A.ComppN,isnull(A.TotalQty,0) TotalQty,isnull(A.ReturnQty,0) ReturnQty,isnull(B.RemainQty,0) RemainQty from QSMS_DIDCallBack A,QSMS_DID B where A.DID=B.DID and work_order in (Select WO from dbo.GetWOGroup('" & sWO & "')) and A.DID='" & sDID & "'"

    '**Denver       2007.12.20     modify callback error info--------------- (0002)
    sSql = "exec DIDInfoForCallBK @WO=" & sq(sWO) & ",@DID=" & sq(sDID)
    
    Set Rst = Conn.Execute(sSql)
    If Rst.EOF = True Then
        MsgBox "System can not get data,Please Retry or contact QMS!!", vbExclamation, "Prompt"
    Else
        If Rst("Result") <> 0 Then
            MsgBox Trim(Rst("Description") & ""), vbExclamation, "Prompt"
        Else
            Set Rst = Rst.NextRecordset
            
            TxtDIDTotalQty = Rst!TotalQty
            TxtDIDReturnedQty = Rst!ReturnQty
            TxtCompPN = Rst!compPN
            txtRemainQty = Rst!RemainQty
        End If
    End If

End Function

Private Function ChkErr() As Boolean
    Dim str As String
    Dim rs As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    ChkErr = True
    
    If IsAnotherBUDID = "Y" Then
        Exit Function
    End If
    
    If Trim(cboWO) = "" Then
        ChkErr = False
        Exit Function
    End If
    
    If Trim(TxtDID) = "" Then
        ChkErr = False
        Exit Function
    End If

    If ChkDIDBelongToPCB(Trim(cboWO), Trim(TxtDID)) = False Then
       ChkErr = False
       Exit Function
    End If
    If Trim(TxtDIDTotalQty) = "" Or Trim(TxtCompPN) = "" Then
       MsgBox "DID total Qty or comppN Can not be empty,Please press enter key in DID txtbox"
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
    
    If CLng(TxtDIDTotalQty) < CLng(TxtReturnQty) Then
       MsgBox "Return Qty can not larger than total qty"
       ChkErr = False
       Exit Function
    End If
    str = "Select SUM(DIDQty) as PCBQty from QSMS_Dispatch where work_order in (Select WO from dbo.GetWOGroup('" & Trim(cboWO) & "')) and DID='" & Trim(TxtDID) & "'"
    Set rs = Conn.Execute(str)
    
    str = "select * from qsms_dispatch where did='" & Trim(TxtDID) & "' and deletedFlag='N' and work_order not in (Select WO from dbo.GetWOGroup(" & Trim(cboWO) & "))"
    Set rsTemp = Conn.Execute(str)
    
    If IsNull(Trim(rs.Fields(0))) = True Then
        MsgBox "This DID did not dispatched!"
        ChkErr = False
        Exit Function
    End If
    
    If rsTemp.EOF = False And CLng(TxtReturnQty) > CLng(Trim(txtRemainQty)) + rs!PCBQty Then
        MsgBox "This DID has dispatched to more than one PCB,CallBack Qty can not larger than the dispatched Qty : " & CLng(Trim(txtRemainQty)) + rs!PCBQty & " of one PCB!", vbInformation
        ChkErr = False
    End If
    
    ''check printer
    If Trim(TxtCompPort) = "" Or Trim(TxtComm) = "" Then
        MsgBox "Printer have not set!!", vbInformation
        ChkErr = False
    End If
    
End Function

Private Sub TxtReturnQty_Click()
    SendKeys "{home}+{end}"
End Sub

Private Sub TxtReturnQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(TxtReturnQty) <> "" Then
        If IsNumeric(Trim(TxtReturnQty)) = False Then
            MsgBox "Please input numeric!!", vbExclamation, "Prompt"
        Else
            TxtReturnQty = Abs(CLng(Trim(TxtReturnQty))) '''(1094) int to CLng
            cmdSave.SetFocus
        End If
    End If
End Sub

Private Function GetSelWO(lstB As ListBox) As String
    Dim I As Integer
    Dim stempWO As String
    
    stempWO = ""
    GetSelWO = ""
    mWOStatus = "N"
    With lstB
        If .ListCount <= 0 Then
            MsgBox "Please select WO to CallBack!!"
            Exit Function
        End If
        For I = 0 To .ListCount - 1
            .ListIndex = I
            
'            If mWOStatus = "N" Then
'                mWOStatus = GetSelWOStatus(Trim(.Text))
'            End If
'            stempWO = stempWO & "''" + Trim(.Text) & "'',"
            stempWO = stempWO & Trim(.Text) & ","
        Next I
        
        stempWO = Mid(stempWO, 1, Len(stempWO) - 1)
        
'        stempWO = "'('+" + stempWO + "+')'"
        GetSelWO = stempWO
        
    End With
End Function

Private Sub GetCallBackType(sWO As String, sDID As String, intDIDQty As Long, intCallBackQty As Long)
    If intDIDQty = intCallBackQty Then
        If chkCallBackQty(sDID, sWO, intCallBackQty, Trim(txtRemainQty), "ByPCB") = True Then
            optCallAll.Value = True
        Else
            TxtReturnQty.SetFocus
            Call TxtReturnQty_Click
        End If
    Else
        optRatebySelWO.Value = True
    End If
    
End Sub

Private Function GetSelWOStatus(sWO) As String
    Dim I As Integer
    With gridWOGroup
        For I = 0 To .ApproxCount
            If sWO = Trim(.Columns(0).Text) And Trim(.Columns(2)) = "Y" Then
                GetSelWOStatus = "Y"
                Exit Function
            End If
        Next I
    End With
    GetSelWOStatus = "N"
End Function

Private Function chkCallBackQty(sDID As String, sSelWo As String, intCallQty As Long, intRemainQty As Long, Optional sType As String = "BySelWO") As Boolean
'    sSql = "Select SUM(DIDQty) as CallQty from QSMS_Dispatch where work_order in ('" & Replace(sSelWo, ",", "','") & "') and DID='" & sDID & "'"
    If IsAnotherBUDID = "Y" Then
        chkCallBackQty = True
        Exit Function
    End If
    
    If sType = "ByPCB" Then
        sSql = "Select SUM(DIDQty) as CallQty from QSMS_Dispatch where work_order in (Select WO from dbo.GetWOGroup('" & sSelWo & "')) and DID='" & sDID & "' and deletedFlag<>'Y'"
    Else
        sSql = "Select SUM(DIDQty) as CallQty from QSMS_Dispatch where work_order in ('" & Replace(sSelWo, ",", "','") & "') and DID='" & sDID & "' and deletedFlag<>'Y'"
    End If
    Set Rst = Conn.Execute(sSql)
    chkCallBackQty = True
    If Rst.EOF = False Then
        If intCallQty - intRemainQty > Rst!CallQty Then
            MsgBox "CallBack Qty:" & intCallQty & " > Qty:" & intRemainQty + Rst!CallQty & " of these Selected WOs dispatched DID!!", vbExclamation, "Prompt"
            chkCallBackQty = False
        End If
    Else
        chkCallBackQty = False
    End If
    
End Function

Private Sub TxtReturnQty_LostFocus()
    If Trim(TxtDID) = "" Or Trim(TxtDIDTotalQty) = "" Then Exit Sub
    If Trim(TxtReturnQty) = "" Or Not IsNumeric(Trim(TxtReturnQty)) Then Exit Sub
    
    Call GetCallBackType(Trim(cboWO), Trim(TxtDID), Val(Trim(TxtDIDTotalQty)), Val(Trim(TxtReturnQty)))
    
End Sub

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

''20071226 Denver Print DID for CallBack
'20100407 Denver    BU Name change (ESBU to CC,ASBU to LC)  （0070）
''===================================================
'Private Function DIDPrintLabel(blnCompPort As Boolean, blnZebra As Boolean, intCompPort As Integer, sCommString As String)

    'If blnCompPort = True Then
       'Call PrintLabelCompPort(blnZebra, intCompPort, sCommString)
    'Else
       'Call PrintLabelPrintPort(blnZebra)
    'End If
'End Function
Private Function DIDPrintLabel(blnZebra As Boolean, intCompPort As Integer, sCommString As String)
    Dim hFile As Long
    Dim hString As String
    Dim strDID As String, tmpDID As String, strQty As String, strDIDType As String
    Dim strDay As String
    Dim LabelFile As String
    Dim m As Integer
    Dim tmpPrintStr As String
    Dim lptPort As Integer
        
    On Error GoTo errHandler
    
        strDay = Format(Now, "YYYY/MM/DD")
        If blnZebra = True Then
            If UCase(DIDInfo.IsGood) = "Y" Then
'                LabelFile = Settings.DIDLabelGood
                strDID = DIDInfo.DID
            Else
'                LabelFile = Settings.DIDLabelBad
                strDID = DIDInfo.compPN
            End If
        Else
'            LabelFile = Settings.AutoDispatchSatoLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
'                LabelFile = Settings.DIDLabelSATOGood
                strDID = DIDInfo.DID
            Else
'                LabelFile = Settings.DIDLabelSATOBad
                strDID = DIDInfo.compPN
            End If
        End If
        ''(1080) get labelfile
        sSql = "select * from MSD_DATA where CompPN=" & sq(TxtCompPN.Text)    '''(1272)
        Set Rst = Conn.Execute(sSql)
        
        If BU = "ESBU" And Rst.EOF = False Then
            LabelFile = GetDIDLabelFile(frmDIDCallBack_New, "GOOD_MSD")
        Else
            LabelFile = GetDIDLabelFile(frmDIDCallBack_New, IIf(DIDInfo.IsGood = "Y", "GOOD", "BAD"))
        End If
        
        strDIDType = DIDInfo.DIDType
        ''特殊处理(打印RefID Label)
        If DIDInfo.Qty <= -10000 Then
            strQty = "RefID"
        Else
            strQty = CStr(DIDInfo.Qty)
        End If
        
        If Dir(LabelFile) = vbNullString Then
            ''''''Added by Jing 2008.01.10  (0004)''''''
            MsgBox ("Can not find label file !"), vbCritical
            DIDPrintLabel = "PRN_FileNoExist"
            Exit Function
        End If
        
        'TxtCompPort   TxtComm
        
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
        
        hFile = FreeFile
        
        If FileReadAll(tmpPrintStr, LabelFile) <= 0 Then
            MsgBox "Open file:" & LabelFile & " fail!!", vbCritical
            Exit Function
        End If
        
        tmpDID = Trim(strDID) '***************add by jeanson 20070814******
            'for Code 128 barcode, the ^ must be tranfer to ><
        If InStr(tmpPrintStr, "<DID_CODE>") > 0 Then
            ''********************************updated by jing 20071024 (0002) ***********
            If blnZebra Then
                tmpDID = Replace(strDID, "^", "><")
            End If
            tmpPrintStr = Replace(tmpPrintStr, "<DID_CODE>", tmpDID)
        End If
            'for text ^, must be use ^FH_ and the use _5E (the ascii of ^)
        If InStr(tmpPrintStr, "<DID_TEXT>") > 0 Then
            ''********************************updated by jing 20071024 (0002) ***********
            If blnZebra Then
               tmpDID = Replace(strDID, "^", "_5E")
            End If
            tmpPrintStr = Replace(tmpPrintStr, "<DID_TEXT>", tmpDID)
        End If

        tmpPrintStr = Replace(tmpPrintStr, "<UID>", UID)
        tmpPrintStr = Replace(tmpPrintStr, "<DIDType>", strDIDType)
            
        tmpPrintStr = Replace(tmpPrintStr, "<DATE>", strDay)
        tmpPrintStr = Replace(tmpPrintStr, "<QTY>", strQty)
            
        ''20090920   Denver NB2&NB3分家，但是用同一个QSMS可作NB2&NB3 的CallBack
        If strQty = "RefID" Then
            tmpPrintStr = Replace(tmpPrintStr, "<LINE>", BUDIDShow)
        Else
            tmpPrintStr = Replace(tmpPrintStr, "<LINE>", IIf(IsAnotherBUDID = "Y", AutoDispatchForAnotherBU, BUDIDShow))
        End If
            
        
        tmpPrintStr = Replace(tmpPrintStr, "<SIDE>", "")
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINE>", IIf(DIDInfo.IsGood = "Y", "", "NG"))
                
        Select Case Trim(tmpPrintStr)
           Case vbNullString
           Case Else
                If OptComp.Value = True Then
                    If blnZebra Then    '(1017)
                        For m = 1 To Len(tmpPrintStr) Step 50
                            MSComm.Output = Mid(tmpPrintStr, m, 50)
                            DoEvents
                        Next m
                    Else '   (0016)
                        For m = 1 To Len(tmpPrintStr) Step 50
                            MSComm.Output = Mid(tmpPrintStr, m, 50)
                        Next m
                    End If
                    MSComm.PortOpen = False
                ElseIf OptPrint.Value = True Then
                    If blnZebra = True Then
                        Print #lptPort, tmpPrintStr & Chr(13)
                    Else '   (0016)
                        For m = 1 To Len(tmpPrintStr) Step 50
                            Print #lptPort, Mid(tmpPrintStr, m, 50)
                        Next m
                    End If
                Else
                    Printer.Print tmpPrintStr
                    Printer.EndDoc
                    Printer.KillDoc
                End If
        End Select
       
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
    Dim m As Integer
        
    On Error GoTo errHandler
    
        strDay = Format(Now, "YYYY/MM/DD")
        If blnZebra = True Then
'            LabelFile = Settings.AutoDispatchLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
                LabelFile = Settings.DIDLabelGood
                strDID = DIDInfo.DID
            Else
                LabelFile = Settings.DIDLabelBad
                strDID = DIDInfo.compPN
            End If
        Else
'            LabelFile = Settings.AutoDispatchSatoLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
                LabelFile = Settings.DIDLabelSATOGood
                strDID = DIDInfo.DID
            Else
                LabelFile = Settings.DIDLabelSATOBad
                strDID = DIDInfo.compPN
            End If
        End If
        
        ''特殊处理(打印RefID Label)
        If DIDInfo.Qty <= -10000 Then
            strQty = "RefID"
        Else
            strQty = CStr(DIDInfo.Qty)
        End If
        
        If Dir(LabelFile) = vbNullString Then
            ''''''Added by Jing 2008.01.10  (0004)''''''
            MsgBox ("Can not find label file !"), vbCritical
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

                hString = Replace(hString, "<UID>", UID)
                
'                hString = Replace(hString, "<DATE>", strDay)
'                hString = Replace(hString, "<QTY>", strQTY)
'                hString = Replace(hString, "<LINE>", PrintData.Line)
'                hString = Replace(hString, "<SIDE>", PrintData.Side)
'                hString = Replace(hString, "<MACHINE>", PrintData.Machine)
                
                hString = Replace(hString, "<DATE>", strDay)
                hString = Replace(hString, "<QTY>", strQty)
'                hString = Replace(hString, "<LINE>", BU)
                
                ''20090920   Denver NB2&NB3分家，但是用同一个QSMS可作NB2&NB3 的CallBack
                If strQty = "RefID" Then
                    hString = Replace(hString, "<LINE>", BUDIDShow)
                Else
                    hString = Replace(hString, "<LINE>", IIf(IsAnotherBUDID = "Y", AutoDispatchForAnotherBU, BUDIDShow))
                End If
                
                
                hString = Replace(hString, "<SIDE>", "")
'                hString = Replace(hString, "<MACHINE>", IIf(DIDInfo.IsGood = "Y", "", "NG"))
                hString = Replace(hString, "<MACHINE>", IIf(DIDInfo.IsGood = "Y", "", "NG"))
                Debug.Print hString
                
               Select Case Trim(hString)
                  Case vbNullString
                  Case Else
'                    MSComm.Output = hString
'                    Debug.Print hString
                    
                    If blnZebra Then
                        MSComm.Output = hString
                    Else '   (0016)
                        For m = 1 To Len(hString) Step 50
                            MSComm.Output = Mid(hString, m, 50)
                            'Debug.Print Mid(hString, m, 50)
                        Next m
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
        Dim hString As String, PrintLabel As String
        Dim strDID As String, tmpDID As String, strQty As String
        Dim FileNum As Integer, lptPort As Integer
        Dim strDay As String
        Dim LabelFile, strLabelFileContent As String
        Dim strPort As String
        Dim m As Integer
        
        On Error GoTo errHandler
        strDay = Format(Now, "YYYY/MM/DD")
        
        ''特殊处理(打印RefID Label)
        If DIDInfo.Qty <= -10000 Then
            strQty = "RefID"
        Else
            strQty = CStr(DIDInfo.Qty)
        End If
        
        If blnZebra = True Then
'            LabelFile = Settings.AutoDispatchLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
                LabelFile = Settings.DIDLabelGood
                strDID = DIDInfo.DID
            Else
                LabelFile = Settings.DIDLabelBad
                strDID = DIDInfo.compPN
            End If
        Else
'            LabelFile = Settings.AutoDispatchSatoLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
                LabelFile = Settings.DIDLabelSATOGood
                strDID = DIDInfo.DID
            Else
                LabelFile = Settings.DIDLabelSATOBad
                strDID = DIDInfo.compPN
                
            End If
            
        End If
'        strLabelFileContent = funGetTxtFileContent(LabelFile)
        If Dir(LabelFile) = vbNullString Then
            ''''''Added by Jing 2008.01.10  (0004)''''''
            MsgBox ("Can not find label file !"), vbCritical
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
            
            hString = Replace(hString, "<UID>", UID)
'                hString = Replace(hString, "<RACKID>", TxtRackID)
            hString = Replace(hString, "<DATE>", strDay)
            hString = Replace(hString, "<QTY>", strQty)
            
            ''hString = Replace(hString, "<LINE>", BU)
            ''20090920   Denver NB2&NB3分家，但是用同一个QSMS可作NB2&NB3 的CallBack
            If strQty = "RefID" Then
                hString = Replace(hString, "<LINE>", BUDIDShow)
            Else
                hString = Replace(hString, "<LINE>", IIf(IsAnotherBUDID = "Y", AutoDispatchForAnotherBU, BUDIDShow))
            End If
                
            hString = Replace(hString, "<SIDE>", "")
            hString = Replace(hString, "<MACHINE>", IIf(DIDInfo.IsGood = "Y", "", "NG"))
            
            'Debug.Print hString
            'Print #lptPort, hString & Chr(13)
            If blnZebra = True Then
                Print #lptPort, hString & Chr(13)
            Else '   (0016)
                For m = 1 To Len(hString) Step 50
                    Print #lptPort, Mid(hString, m, 50)
                Next m
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


''DID CallBack 检查另一个BU 的DID 信息,并显示相关的信息
Private Function XL_ChkAnotherBUDID(sDID As String) As Boolean
On Error GoTo errHandler

    XL_ChkAnotherBUDID = False
    
    'QMS             Denver         2011/01/12     Unify SP:XL_ChkAnotherBUDID in CallBack/ReturnDID   (1048)
    'sSql = "exec DIDChkAnotherBU  " & sq(sDID) & "," & sq(IsAnotherBUDID)
    sSql = "exec DIDChkAnotherBU  " & sq(sDID) & "," & sq(IsAnotherBUDID) & "," & sq(Trim(Factory))
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
            TxtCompPN = Trim(Rst!compPN)
            IsAnotherBUDID = Trim(Rst!IsAnotherBUDID)
        End If
        
        ''Fill WO info and dispatch info
        Set Rst = Rst.NextRecordset
        lstAvailableWO.Clear
        lstCallBackWO.Clear
        Dim PreWO As String
        PreWO = ""
        Do While Rst.EOF = False
            If PreWO = "" Or PreWO <> Trim(Rst!Work_Order) Then
                lstAvailableWO.AddItem Trim(Rst!Work_Order)
                PreWO = Rst!Work_Order
            End If
            
            Rst.MoveNext
        Loop
'        rst.MoveFirst
        Set gridDIDDispatched.DataSource = Rst
        TxtReturnQty = ""
        If lstAvailableWO.ListCount = 1 Then
            optRatebySelWO.Value = True
            Call cmdADDALL_Click
            TxtReturnQty.SetFocus
            Call TxtReturnQty_Click
        End If
        
    End If

    XL_ChkAnotherBUDID = True


    Exit Function
errHandler:
     MsgBox Err.Description

End Function




