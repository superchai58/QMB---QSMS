VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmMaintainFeeder 
   Caption         =   "Maintain Feeder[20140922]"
   ClientHeight    =   9825
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9825
   ScaleWidth      =   14985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraDID 
      BackColor       =   &H80000013&
      Caption         =   "DID & Feeder"
      Height          =   6855
      Left            =   120
      TabIndex        =   30
      Top             =   3000
      Width           =   13815
      Begin VB.ComboBox cboJobGroup 
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
         Left            =   1440
         TabIndex        =   53
         Top             =   840
         Width           =   3135
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
         Left            =   4680
         Picture         =   "FrmMaintainFeeder.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "outout Feeder& DID "
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptFuJi 
         Caption         =   "Without LR"
         Height          =   255
         Left            =   5760
         TabIndex        =   39
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton OptPanal 
         BackColor       =   &H80000004&
         Caption         =   "Has LR"
         Height          =   255
         Left            =   5760
         TabIndex        =   38
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset"
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
         Left            =   12240
         Picture         =   "FrmMaintainFeeder.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TxtLR 
         BackColor       =   &H00FFFFFF&
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
         Left            =   8640
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox TxtCompPN 
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
         Left            =   8640
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   600
         Width           =   3375
      End
      Begin VB.ComboBox CboMachine 
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
         ItemData        =   "FrmMaintainFeeder.frx":074C
         Left            =   1440
         List            =   "FrmMaintainFeeder.frx":074E
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   240
         Width           =   3135
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
         Height          =   375
         Left            =   8640
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   120
         Width           =   3375
      End
      Begin VB.TextBox TxtFeeder 
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
         Left            =   8640
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3375
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
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
         Left            =   12240
         Picture         =   "FrmMaintainFeeder.frx":0750
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   360
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid DGMachine 
         Height          =   2055
         Left            =   120
         TabIndex        =   41
         Top             =   2640
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Caption         =   "Information of Feeder and DID bound to the current machine"
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
      Begin MSDataGridLib.DataGrid DGDID 
         Height          =   1935
         Left            =   120
         TabIndex        =   42
         Top             =   4680
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3413
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Caption         =   "Material information of the machine on the workorder"
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
      Begin MSDataGridLib.DataGrid DGNeed 
         Height          =   2055
         Left            =   6840
         TabIndex        =   43
         Top             =   2640
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3625
         _Version        =   393216
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Machine information of unbound feeder"
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
      Begin MSDataGridLib.DataGrid DgDIDSlot 
         Height          =   1935
         Left            =   6840
         TabIndex        =   51
         Top             =   4680
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3413
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Caption         =   "Feeder and DID information that is being used on the machine"
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
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Job Group"
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
         TabIndex        =   50
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label LblMessage 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   49
         Top             =   2160
         Width           =   13575
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "LR"
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
         Left            =   7320
         TabIndex        =   48
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "Comp PN"
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
         TabIndex        =   47
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "Machine"
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
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Left            =   7320
         TabIndex        =   45
         Top             =   120
         Width           =   1215
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
         Index           =   0
         Left            =   7320
         TabIndex        =   44
         Top             =   1080
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   12960
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CompQty"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "WO Information"
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12615
      Begin VB.CommandButton cmdCheck 
         Caption         =   "&Check"
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
         Left            =   4440
         TabIndex        =   52
         Top             =   1800
         Width           =   975
      End
      Begin VB.ListBox ListNoChkBOM 
         Height          =   255
         ItemData        =   "FrmMaintainFeeder.frx":0A5A
         Left            =   9960
         List            =   "FrmMaintainFeeder.frx":0A5C
         TabIndex        =   28
         Top             =   360
         Width           =   2535
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
         Left            =   6720
         TabIndex        =   27
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Frame FraSB 
         Caption         =   "Small Board WO"
         Height          =   615
         Left            =   1320
         TabIndex        =   25
         Top             =   2160
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
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   3120
         TabIndex        =   22
         Top             =   840
         Width           =   1215
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
         Left            =   6720
         TabIndex        =   19
         Top             =   360
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
         Left            =   6720
         TabIndex        =   18
         Top             =   1800
         Width           =   2055
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
         Left            =   6720
         TabIndex        =   15
         Top             =   840
         Width           =   2055
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
         Left            =   6720
         TabIndex        =   14
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ListBox ListClosed 
         Height          =   255
         ItemData        =   "FrmMaintainFeeder.frx":0A5E
         Left            =   9960
         List            =   "FrmMaintainFeeder.frx":0A60
         TabIndex        =   11
         Top             =   1920
         Width           =   2535
      End
      Begin VB.ListBox ListNotDispatch 
         Height          =   255
         ItemData        =   "FrmMaintainFeeder.frx":0A62
         Left            =   9960
         List            =   "FrmMaintainFeeder.frx":0A64
         TabIndex        =   10
         Top             =   1200
         Width           =   2535
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
         Height          =   855
         Left            =   4440
         Picture         =   "FrmMaintainFeeder.frx":0A66
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   975
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
         ItemData        =   "FrmMaintainFeeder.frx":0EA8
         Left            =   1440
         List            =   "FrmMaintainFeeder.frx":0EAA
         TabIndex        =   8
         Top             =   1800
         Width           =   2535
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
         Left            =   1440
         TabIndex        =   7
         Top             =   1320
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   129892355
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   1440
         TabIndex        =   23
         Top             =   840
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   129892355
         CurrentDate     =   36482
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Not chk bom"
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
         Height          =   615
         Index           =   5
         Left            =   8880
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
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
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1335
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
         Index           =   3
         Left            =   5520
         TabIndex        =   21
         Top             =   360
         Width           =   1095
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
         Left            =   5520
         TabIndex        =   20
         Top             =   1800
         Width           =   1095
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
         Left            =   5520
         TabIndex        =   17
         Top             =   840
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
         Left            =   5520
         TabIndex        =   16
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Wo Not Disptach"
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
         Height          =   615
         Index           =   3
         Left            =   8880
         TabIndex        =   13
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "WO Closed "
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
         Height          =   615
         Index           =   0
         Left            =   8880
         TabIndex        =   12
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label3 
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
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
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
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
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
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "group"
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
         Index           =   4
         Left            =   5520
         TabIndex        =   2
         Top             =   2280
         Width           =   1095
      End
   End
   Begin MCI.MMControl wave_control 
      Height          =   330
      Left            =   4560
      TabIndex        =   0
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
Attribute VB_Name = "FrmMaintainFeeder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************************************************************
'**文 件 名: FrmMaintainFeeder.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: Jeanson
'**日    期: 2007.10.01
'**描    述: QSMS Maintain Feeder DID
'
'EQMS_ID        **修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'               Jing        2007.10.19      Modify Maintain DID to check if DID already dispatch (0001)
'               Jeanson     2007.10.21      Add auto get the focus after the DID is not already dispatched  (0002)
'               Kane        2007.10.26      Add function to check if did has been spliced for NB25 (0003)
'               Kane        2007.11.06      Don't check did dispatch information by PCB group(0004)
'               Jeanson     2007.11.07      Check did dispatch information by WO group(0005)
'               Jing        2008.01.10      Changed from 'NB25' to 'NB5' (0006)
'               Jing        2008.03.03      maintain Feeder by Line for NB5  (0007)
'               Lynn        2008.09.02      add a interface let PD update DID real qty  (0008)
'               Kane        2009.03.18      Check feeder's line whether match with machine's line (0009)
'               Kane        2009.03.19      Check feeder whether need repair '(0010)
'               Kane        2009.05.12      Check feeder used time before save data '(0011)
'QMS            Archer      2009.05.26      If user can not bind Feeder and DID, maybe the substitute data causes such issue (0012)
'QMS            Archer      2009/12/02      Check DID use SP:CheckDIDValidity (0013)
'QMS            Kevin       2010/05/16      Modify get machine list methods  (0014)
'QMS            Lynn        2012/07/12      Add Line For QSMS_Feeder,NB6 will transfer data to fuji trax according line  (0015)
'***********************************************************************************/


Option Explicit
Dim LastMachine As String
Dim rsMachine As ADODB.Recordset
Dim ChkCompPN As String
Dim ChkFeederLine As String
Dim CheckFeeder As String
Dim strDelaytime As Long
Dim t As Long

Private Sub CboGroupID_Click()
Call GetWoByGroupID(Trim(CboGroupID))
End Sub

Private Sub CboGroupID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboGroupID_Click
End If
End Sub
 
Private Sub cboJobGroup_click()
    If cboJobGroup.Text <> "" Then
        RefreshMachineFeeder
    End If
End Sub

Private Sub CboMachine_Click()
    Dim str As String
    Dim temp As String
    Call GetJobByMachine(Trim(cboMachine))
    If LastMachine <> "" Then
        str = "exec GetUnlinkFeederSlot '" & Trim(LastMachine) & "','','" & Trim(cboWO) & "'"
        Set RS = Conn.Execute(str)
        If RS.RecordCount <> 0 Then
            If MsgBox("Some Feeder  not Link DID at  Machine: " + LastMachine + ",is Cut in Machine?", vbYesNo) = vbNo Then
                temp = LastMachine
                LastMachine = ""
                cboMachine.Text = temp
                Exit Sub
            End If
        End If
    End If
 
    RefreshMachineFeeder
    LastMachine = cboMachine
    If TxtDID.Enabled = True Then
        TxtDID.SetFocus
    End If
End Sub

Private Sub CboMachine_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboMachine_Click
End If
End Sub
 
Sub RefreshMachineFeeder()
Dim str As String
Dim RS As New ADODB.Recordset

    '''''(0014)''''''''(1151)
    str = "select distinct  Machine,Feeder,Slot,LR ,CompPN,TransDateTime,DID, Used, VendorCode, DateCode, LotCode, JobGroup,JobPN, Version  from QSMS_Feeder where jobGroup='" & Trim(cboJobGroup.Text) & "'  and machine='" & Trim(cboMachine) & "'order by Slot,LR  "
    If RS.State Then RS.Close
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
        Set DGMachine.DataSource = RS
    End If
 
    str = "exec GetUnlinkFeederSlot '" & Trim(cboMachine) & "','" & Trim(cboJobGroup.Text) & "','" & Trim(cboWO) & "'"
    Set RS = Conn.Execute(str)
    Set DGNeed.DataSource = RS
      
    str = "select distinct Machine ,CompPN ,Slot ,LR from QSMS_Dispatch where Machine='" & Trim(cboMachine) & "' and Work_Order= '" & Trim(cboWO) & "' order by Slot,LR"
    Set RS = Conn.Execute(str)
    Set DGDID.DataSource = RS
   
    str = "select distinct  Machine,Feeder,Slot ,LR ,DID ,DIDCompPN  from QSMS_FeederDID_Current where Machine='" & Trim(cboMachine) & "' and WorkOrder= '" & Trim(cboWO) & "' order by Slot,LR"
    Set RS = Conn.Execute(str)
    Set DgDIDSlot.DataSource = RS
   
    
End Sub
 

Private Sub CboWo_Click()

Call GetSBWO(cboWO)

Call GetMachineByWo(cboWO)
Call GetWoinfo(Trim(cboWO))
End Sub

Private Sub cboWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboWo_Click
End If
End Sub

Private Sub cmdCheck_Click()             ''''(1151)
Dim strSQL As String
    If cboWO = "" Then
       MsgBox "WO  can not be empty,Please check"
       Exit Sub
    End If
    strSQL = "exec CheckWOLinkFeeder '" & Trim(cboWO) & "' "
    If RS.State Then RS.Close
    Set RS = Conn.Execute(strSQL)
    If Trim(RS!result) <> "PASS" Then
        MsgBox (RS!Msg), vbCritical
        Call CopyToExcel(RS.NextRecordset)  '1159
    Else
        MsgBox ("WO Link Feeder is  OK !!!"), vbCritical
    End If
End Sub

Private Sub CmdExcel_Click()
Dim rsTemp As ADODB.Recordset
Dim strSQL As String

  strSQL = "select distinct  Machine,Feeder,CompPN,TransDateTime,DID,LR, Slot, Used, VendorCode, DateCode, LotCode, JobGroup,JobPN, Version  from QSMS_Feeder where jobGroup='" & Trim(cboJobGroup.Text) & "'  and machine='" & Trim(cboMachine) & "'order by Slot,LR  "
  If RS.State Then RS.Close
  Set RS = Conn.Execute(strSQL)
  If Not RS.EOF Then
        ''query unlink feeder and DID by WO'1061==>move it to unireport
'          If CboMachine.Text = "" Then
'               Set rsTemp = Conn.Execute(str)
'               if
'               Call CopyToExcel(rsMachine)
'          Else
              Call CopyToExcel(RS)
'          End If
    Else
          MsgBox ("Please choose the Machine and JobGroup First"), vbCritical
    End If
  
End Sub

Private Sub CmdQuery_Click()
Dim str As String
Dim BeginDate As String
Dim EndDate As String
Dim RS As ADODB.Recordset
BeginDate = Format(dtpSDate, "YYYY/MM/DD")
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
EndDate = Format(dtpEDate, "YYYY/MM/DD")
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")
If Trim(CboLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If
Call GetGroupIDByLine(Trim(CboLine), BeginDate, EndDate)
End Sub

Private Sub cmdReset_Click()
TxtDID.Enabled = True
TxtCompPN.Enabled = True
TxtFeeder.Enabled = True

TxtLR.Enabled = True

TxtDID.Text = ""
TxtCompPN.Text = ""
TxtFeeder.Text = ""

TxtLR.Text = ""
With DIDInfo
    .DID = ""
    .COMPPN = ""
    .VendorCode = ""
    .DateCode = ""
    .LotCode = ""
End With
TxtDID.SetFocus

End Sub
Private Function GetWoinfo(ByVal WO As String)
Dim str As String
Dim RS As ADODB.Recordset

str = "select PN, Qty from Sap_Wo_List where WO='" & Trim(WO) & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   TxtMBPN = RS!PN
   TxtModel = Mid(TxtMBPN, 3, 3)
   TxtWOQty = RS!Qty
End If
str = "select Customer from ModelName where PN='" & TxtMBPN & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   TxtCustomer = Trim(RS!Customer)
End If
End Function

Private Function ChkValid() As Boolean
Dim str As String
Dim RS As ADODB.Recordset
Dim transdatetime As String
Dim LR As String
ChkValid = True
If ChkDID(Trim(TxtDID)) = False Then
   ChkValid = False
End If
If MaintainFeederDID <> "Y" Or OptFuJi.Value <> True Then ''(1103)
    If OptPanal.Value = True Then
      If UCase(Trim(TxtLR)) = "L" Or UCase(Trim(TxtLR)) = "R" Or UCase(Trim(TxtLR)) = "0" Then
           
      Else
        MsgBox "LR Is invalid,Pleaes check"
        TxtLR.Enabled = True
        TxtLR.SetFocus
        ChkValid = False
      End If
    End If
    If Trim(cboJobGroup.Text) = "" Or Trim(cboMachine) = "" Then
       MsgBox "JObPN or Machine or Version can not be empty,Please check"
       ChkValid = False
    End If
End If
'check LR,,L--1,R--2
'ChkIfInCurretnFeeder
If ChkIfInCurretnFeeder(Trim(TxtFeeder), Trim(TxtDID)) = True Then
  MsgBox "The DID or Feeder In machine use,Please clear the link relationship on DID & Slot Link"
  ChkValid = False
End If
End Function
Private Function GetLRMapping() As String
If OptFuJi.Value = True Then
   GetLRMapping = 0
   Exit Function
End If
If OptPanal.Value = True Then
   Select Case UCase(Trim(TxtLR))
          Case "0"
               GetLRMapping = "0"
          Case "L"
               GetLRMapping = "1"
          Case "R"
               GetLRMapping = "2"
   End Select
End If

End Function
Private Sub cmdSave_Click()
Dim str As String, strLine As String, tmpLine As String
Dim RS As ADODB.Recordset
Dim transdatetime As String
Dim Slot, LR As String
Dim ErrorMsg As String
Dim TempRealQty As Integer
Dim TempOldRealQty As Integer
Dim TempOldDID As String
Dim tempCompPN As String
On Error GoTo EcmdSave_Click        'add by jeanson 20061122
str = "select getdate()"
Set RS = Conn.Execute(str)

If Not RS.EOF Then
    transdatetime = Format(RS.Fields(0), "YYYYMMDDHHMMSS")
End If

''1087
strSQL = "select realQty,DID,CompPN from qsms_DID where NextDID='" & Trim(TxtDID) & "'"
Set RS = Conn.Execute(strSQL)
If Not RS.EOF Then
    TempRealQty = RS!realqty
    TempOldDID = RS!DID
    tempCompPN = RS!COMPPN
    If TempRealQty > 0 Then
        TxtDID = TempOldDID
        TxtCompPN = tempCompPN
        strSQL = "select realQty,DID,CompPN from qsms_DID where NextDID='" & Trim(TempOldDID) & "'"
        Set RS = Conn.Execute(strSQL)
        If Not RS.EOF Then
            TempOldRealQty = RS!realqty
            If TempOldRealQty > 0 Then
                TxtDID = RS!DID
                TxtCompPN = RS!COMPPN
            End If
        End If
        LblMessage.Caption = "DID:" & Trim(TxtDID) & " is Inherit Material."
    End If
End If
''''1087

If ChkValid = False Then
    Exit Sub
End If

''''''added by Jing 2008.03.03  (0007)''''''
If CheckFeeder = "Y" Then
    str = "select line from qsms_dispatch where did='" & Trim(TxtDID) & "'"
    Set RS = Conn.Execute(str)
    If RS.EOF = False Then
        strLine = Trim(RS("Line"))
    Else
        MsgBox "This DID not have dispatch record", vbCritical
        Exit Sub
    End If
    
    str = "select line from FMS_FeederList where FeederID='" & Trim(TxtFeeder) & "'"
    Set RS = Conn.Execute(str)
    If RS.EOF = False Then
        'tmpLine = Trim(rs("Line"))
        If ChkFeederLine = "Y" Then '(0009)
            If strLine <> Trim(RS("Line")) Then
                MsgBox ("This Feeder must be used in line: " & Trim(RS("Line")) & " !"), vbCritical
                Exit Sub
            End If
            '(0010)
            str = "select * from fms_maintain where FeederId='" & Trim(TxtFeeder) & "' and actionmeasure=''"
            Set RS = Conn.Execute(str)
            If RS.EOF = False Then
                MsgBox "This feeder must be repair first", vbCritical
                Exit Sub
            End If
        End If
    End If
     '(0011)============
    str = "exec FeederCheck '" & Trim(TxtFeeder) & "'"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
        If RS("ErrorCode") <> 0 Then
            If RS("ErrorCode") = 100 Then
                MsgBox RS("Result")
                TxtFeeder.Text = ""
                TxtFeeder.SetFocus
                Exit Sub
            Else
                MsgBox RS("Result")
            End If
        End If
    End If
     '(0011)==========
End If

''''(0012)
 LR = GetLRMapping   ''(1046) LR取值放在前面，不能放在pana下面
 
If MaintainFeederDID <> "Y" Or OptFuJi.Value <> True Then ''(1103)
    str = "Select distinct Slot,LR from QSMS_wo where JobGroup='" & Trim(cboJobGroup.Text) & "' and CompPN='" & Trim(TxtCompPN) & "' and machine='" & cboMachine & "'  and   work_order in (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "')"
    Set RS = Conn.Execute(str)
    Slot = ""
    ErrorMsg = "Component is error,Please check the Machine Name!!"
    If RS.EOF Then
        MsgBox ErrorMsg
    '    Str = "Exec QSMS_CheckLostReplacePN @WO='" & Trim(cboWO.Text) & "'"            '''mark by Archer 200090611
    '    Conn.Execute Str
    '    MsgBox "若Machine/JobGroup/CompPN/WO都没问题，可能是替代料原因，请重新CheckBom后再试!"
        Exit Sub
    Else
        If InStr(1, UCase(Trim(cboMachine)), "MSF") > 0 Then
            While Not RS.EOF
                 Slot = Slot & Trim(RS!Slot) & ","
                 RS.MoveNext
            Wend
        Else
            While Not RS.EOF
                If RS!LR = Trim(LR) Then Slot = Slot & Trim(RS!Slot) & ","
                RS.MoveNext
            Wend
            ErrorMsg = "LR is error,Please check the LR!!"
        End If
        If Slot <> "" Then
            Slot = Mid(Slot, 1, Len(Slot) - 1)
        Else
            MsgBox ErrorMsg
            Exit Sub
        End If
    End If
ElseIf MaintainFeederDID = "Y" And OptFuJi.Value = True Then
    str = "Select top 1 0 from QSMS_wo with(nolock) where  CompPN='" & Trim(TxtCompPN) & "'"
    Set RS = Conn.Execute(str)
    
    ErrorMsg = "Component is error,Please check the Machine Name!!"
    If RS.EOF Then
        MsgBox ErrorMsg
        Exit Sub
    End If
''1210

str = "select * from qsms_ProConfig where line='" & strLine & "' and [key]='FujiCheckFeeder'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
    If RS("value") = "Y" Then

    str = "exec Fuji_FeederCheck " & _
        "'" & Trim(TxtFeeder) & "','" & strLine & "'"
    Set RS = Conn.Execute(str)
    
        
        If RS("result") <> 0 Then
            MsgBox RS("msg")
            Exit Sub
        End If
    End If
 End If
        
 ''1210
   
    
End If
''(0015)
str = "exec QSMS_Feeder_MainTain " & _
    "'" & Trim(cboMachine) & "','" & Trim(cboJobGroup.Text) & "','','" & Trim(TxtDID) & "','" & Trim(TxtCompPN) & "','" & DIDInfo.VendorCode & "','" & DIDInfo.DateCode & "','" & DIDInfo.LotCode & "' " & _
    ",'" & Trim(TxtFeeder) & "','" & Left(Slot, 20) & "','" & Trim(LR) & "','" & g_userName & "','" & OptPanal.Value & "','" & strLine & "' "
Set RS = Conn.Execute(str)
If Not RS.EOF Then  '''1177
    If RS("Result") = 1 Then
        MsgBox RS("Msg")
        TxtDID.Text = ""
        TxtDID.SetFocus
        Exit Sub
    End If
    If StrBU = "NB3" Then ''1281
      If RS("Shim") = "Y" Then
          FrmMessage.WindowState = 2
          FrmMessage.lblmsg = RS("Descr")
          FrmMessage.Show vbModal
      End If
    End If
End If  '''1177
strSQL = "insert into qsms_log(system_name,event_no,did,user_name,returnqty,trans_date) values('QSMS','MaintainFeeder'," & _
                  "'Line:" & CboLine.Text & ";DID:" & Trim(TxtDID.Text) & ";Machine:" & Trim(cboMachine.Text) & ";Feeder:" & Trim(TxtFeeder.Text) & "','" & g_userName & "',0,dbo.formatdate(getdate(),'yyyymmddhhnnss'))"        ''(1081)
Conn.Execute (strSQL)

RefreshMachineFeeder

LblMessage.Caption = "Insert OK ,Feeder:" & Trim(TxtFeeder) & ",    Slot:" & Left(Slot, 20) & ",    LR:" & Trim(LR)
Call cmdReset_Click
Call OK_Sound


Exit Sub
EcmdSave_Click:
    MsgBox Err.Description + ",Please contact QSMS SMT Staff"
End Sub

Private Sub CmdSave_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Or KeyAscii = 9 Then
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
Dim str As String
Dim RS As ADODB.Recordset
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
BU = ReadIniFile("Common", "BU", App.Path & "\set.ini")
str = "select getdate()"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
    'Date = Rs(0)
    'Time = Rs(0)
End If
dtpSDate = Date
dtpEDate = Date
cboMachine.Clear
str = "select distinct Machine From QSMS_MEbom order by Machine"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
    RS.MoveFirst
    While Not RS.EOF
          cboMachine.AddItem Trim(RS!Machine)
          RS.MoveNext
    Wend
Else
   MsgBox "can not find Machine from ME BOM"
   Exit Sub
End If
LastMachine = ""
Call GetLine
ChkFeederLine = ReadIniFile("VerifyPart", "CheckFeederLine", App.Path & "\set.ini")
CheckFeeder = ReadIniFile("VerifyPart", "CheckFeeder", App.Path & "\set.ini")

If App.Title <> App.EXEName Then  ' 调试时可以使用拷备和粘贴功能
    Call Hook(TxtCompPN.hWnd)   '启动钩子
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call UnHook   '撤消钩子
End Sub

Private Sub OptFuJi_Click()
TxtLR.Enabled = False
TxtLR.BackColor = &H808080
End Sub

Private Sub OptPanal_Click()
TxtLR.Enabled = True
TxtLR.BackColor = &H80000005
End Sub

Private Sub txtCompPN_KeyPress(KeyAscii As Integer)
Dim I As Long
Dim tempCompPN As String
Dim NewComp() As String

If Len(Trim(TxtCompPN)) < 1 Then t = 0
If t <> 0 Then
    If GetTickCount - t > 100 Then
        MsgBox "Please use scaner!"
        TxtCompPN.Text = ""
        t = 0
        Exit Sub
    End If
End If
t = GetTickCount
If KeyAscii = 13 Then
    t = 0
End If

If KeyAscii = 13 Or KeyAscii = 9 Then
'    I = InStrRev(TxtDID, "-")
'    If I = 0 Then
'       MsgBox "DID format is error,Please check"
'       TxtDID.Enabled = True
'       TxtDID.SetFocus
'       Exit Sub
'    End If
'    TempCompPN = Mid(TxtDID, 1, I - 1)
    If InStr(1, Trim(TxtCompPN.Text), ";") > 0 Then  '(1045)
         NewComp = Split(Trim(TxtCompPN.Text), ";")  '(1045)
         TxtCompPN.Text = NewComp(0)  '(1045)
    End If  '(1045)
    If UCase(DIDInfo.COMPPN) = UCase(Trim(TxtCompPN)) Then
       TxtCompPN.Enabled = False
       TxtFeeder.SetFocus
    Else
       Call Warning_Sound
       TxtCompPN.Text = ""
       TxtCompPN.SetFocus
       LblMessage.Caption = "Comp format error"
    End If
End If

End Sub
Private Sub txtDID_Click()
SendKeys "{HOME}+{END}" '(0002)
End Sub

Private Sub TxtDID_KeyDown(KeyCode As Integer, Shift As Integer)
If App.Title <> App.EXEName Then
    If Shift = 2 Then   '1076
        MsgBox "Can't use Ctrl+V and Ctrl+C,input is void!", vbCritical
        TxtDID.Text = ""
    End If
End If
End Sub

Private Sub txtDID_KeyPress(KeyAscii As Integer)
Dim RS As ADODB.Recordset
Dim str As String
Dim I As Long

If Len(Trim(TxtDID.Text)) < 1 Then strDelaytime = 0   '''''1076
    If strDelaytime <> 0 Then
        If GetTickCount - strDelaytime > 100 Then
            MsgBox "Please use scaner!"
            TxtDID.Text = ""
            strDelaytime = 0
            Exit Sub
        End If
    End If
    strDelaytime = GetTickCount
If KeyAscii = 13 Or KeyAscii = 9 Then
    strDelaytime = 0
    If ChkDID(Trim(TxtDID)) = False Then
        TxtDID = ""
        TxtDID.SetFocus
        Exit Sub
    End If
    
    If ChkAVL(Trim(DIDInfo.COMPPN), Trim(DIDInfo.VendorCode), Trim(TxtCustomer), Trim(TxtModel)) = False Then
       Exit Sub
    End If

    TxtDID.Enabled = False
    TxtCompPN.Enabled = True
    TxtCompPN.SetFocus
    
    If UnChkCompPN = "Y" Then   ''1187
        I = InStrRev(TxtDID, "-")
        TxtCompPN = Mid(TxtDID, 1, I - 1)
        TxtCompPN.Enabled = False
        TxtFeeder.SetFocus
    End If

End If
End Sub

Private Sub TxtDID_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If App.Title <> App.EXEName Then
    If Button = 2 Then  '1076
        MsgBox "Please use scaner to do it!!", vbDefaultButton1
        TxtDID.Text = ""
        TxtDID.SetFocus
    End If
End If
End Sub

Private Sub TxtFeeder_KeyPress(KeyAscii As Integer)
Dim RS As ADODB.Recordset
Dim str As String
Dim tempmachine As String
Dim TempFeeder As String
Dim tempSlot  As String
On Error GoTo Handler
If KeyAscii = 13 Or KeyAscii = 9 And TxtFeeder <> "" Then
'************************add by jeanson 2007/09/14*****************************
    str = "exec FeederCheck '" & Trim(TxtFeeder) & "'"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
        If RS("ErrorCode") <> 0 Then
            If RS("ErrorCode") = 100 Then
                MsgBox RS("Result")
                TxtFeeder.Text = ""
                TxtFeeder.SetFocus
                Exit Sub
            Else
                MsgBox RS("Result")
            End If
        End If
    End If
    str = "select * from QSMS_Feeder where Feeder='" & Trim(TxtFeeder) & "'"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
        tempSlot = RS!Slot
        tempmachine = RS!Machine
        TempFeeder = RS!Feeder
        If MsgBox("Feeder:" + Trim(TempFeeder) + "is Link Machine " + tempmachine + ",is Delete Current Link?", vbYesNo) = vbNo Then

            TxtFeeder.Text = ""
            TxtFeeder.SetFocus
            LblMessage = "Feeder:" + Trim(TxtFeeder) + "is Link Machine " + tempmachine + ";Slot:" + tempSlot + ",Please Check"
            Exit Sub
        End If
    End If
'************************add by jeanson 2007/09/14*****************************
     If OptPanal.Value = True Then
        TxtLR.Enabled = True
        TxtLR.SetFocus
        Exit Sub
     End If
     If OptFuJi.Value = True Then
         TxtLR.Enabled = False
         Call cmdSave_Click
     End If
End If
Exit Sub
Handler:
'************************add by jeanson 2007/09/14*****************************
    MsgBox Err.Description + ", Please call QMS"
'************************add by jeanson 2007/09/14*****************************
End Sub

Private Function GetJobByMachine(ByVal Machine As String)
Dim str As String
Dim RS As ADODB.Recordset

If Machine = "" Then
   MsgBox "Please select Machine"
   Exit Function
End If

cboJobGroup.Clear



str = "select  distinct JobGroup from QSMS_Wo where work_order in (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "') and machine='" & Machine & "'"    ''1186
Set RS = Conn.Execute(str)
While Not RS.EOF
          cboJobGroup.AddItem Trim(RS!jobgroup)
          RS.MoveNext
Wend
     cboJobGroup.Enabled = True
End Function
'TxtJobGroup = ""
'str = "select  distinct JobGroup from QSMS_Wo where work_order in (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "' and InitAOIFlag='Y') and machine='" & Machine & "'"    ''1186
'Set RS = Conn.Execute(str)
'If Not RS.EOF Then
'    TxtJobGroup = Trim(RS!jobgroup)
'Else
'    str = "select  distinct JobGroup from QSMS_Wo where work_order in (select wo from sap_wo_list where [group]='" & Trim(TxtGroup) & "') and machine='" & Machine & "'"    ''(1175)
'    Set RS = Conn.Execute(str)
'    TxtJobGroup = Trim(RS!jobgroup)
'End If
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

Private Function GetDIDInfo(ByVal DID As String)
Dim str As String
Dim RS As ADODB.Recordset
str = "select DID,CompPN,VendorCode,DateCode,LotCode from QSMS_DID where DID='" & DID & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
    With DIDInfo
         .DID = DID
         .COMPPN = Trim(RS!COMPPN)
         .VendorCode = Trim(RS!VendorCode)
         .DateCode = Trim(RS!DateCode)
         .LotCode = Trim(RS!LotCode)
    End With
End If
End Function

Private Function ChkDID(ByVal DID As String) As Boolean
Dim RS As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim str As String
Dim wostr As String
ChkDID = True
If ChkNonAVL(DID, Trim(TxtCustomer), Trim(TxtModel), Trim(TxtMBPN), Trim(cboWO)) = False Then
   ChkDID = False
   Exit Function
End If

''''(0013) '(1002)
str = "Exec CheckDIDValidity @DID='" & DID & "',@LineSide='" & Left(Trim(cboMachine.Text), 2) & "',@WOGroup='" & Trim(CboGroupID.Text) & "',@CurrentWO='" & Trim(cboWO.Text) & "',@no_output='N'" & " ,@Machine='" & Trim(cboMachine.Text) & "',@Line='" & Trim(CboLine.Text) & "'"   '(1034)
Set RS = Conn.Execute(str)
If Not RS.EOF Then
    If Left(RS!result, 4) = "PASS" Then
        With DIDInfo
            .DID = DID
            .COMPPN = Trim(RS!DIDCompPN)
            .VendorCode = Trim(RS!VendorCode)
            .DateCode = Trim(RS!DateCode)
            .LotCode = Trim(RS!LotCode)
        End With
    Else
        MsgBox RS!result
        ChkDID = False
        Exit Function
    End If
End If
End Function

Private Sub TxtLR_Click()
SendKeys "{HOME}+{END}"
End Sub

Private Sub TxtLR_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13 Or KeyAscii = 9) Then
   If UCase(Trim(TxtLR)) = "L" Or UCase(Trim(TxtLR)) = "R" Or UCase(Trim(TxtLR)) = "0" Then
  
        Call cmdSave_Click
    Else
        MsgBox "Please input L or R or 0 "
        TxtLR.SetFocus
        Call TxtLR_Click
   End If
End If
End Sub

Private Function GetGroupIDByLine(ByVal Line As String, ByVal BeginDate As String, ByVal EndDate As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim GroupIDHead As String

CboGroupID.Clear
str = "Select distinct GroupID from QSMS_WoGroup where substring(Group_TransDateTime,1,8) between  '" & BeginDate & "'  and '" & EndDate & "' and line='" & Trim(CboLine) & "'"
Set RS = Conn.Execute(str)
While Not RS.EOF
     CboGroupID.AddItem Trim(RS!GroupID)
     RS.MoveNext
Wend

End Function

Private Function GetWoByGroupID(ByVal GroupID As String)
Dim str As String
Dim RS As ADODB.Recordset

cboWO.Clear
ListNotDispatch.Clear
ListClosed.Clear
ListNoChkBOM.Clear
str = "Select Work_Order,Sap1Flag,ClosedFlag from QSMS_WoGroup where GroupID='" & GroupID & "' "
Set RS = Conn.Execute(str)
While Not RS.EOF
     If ChkMBWo(RS!Work_Order) = True Then
        If ChkQSMS_WO(Trim(RS!Work_Order)) = True Then
            cboWO.AddItem Trim(RS!Work_Order)
            If UCase(Trim(RS!sap1flag)) = "N" Then
               ListNotDispatch.AddItem Trim(RS!Work_Order)
            End If
            If UCase(Trim(RS!ClosedFlag)) = "Y" Then
               ListClosed.AddItem Trim(RS!Work_Order)
            End If
        Else
           ListNoChkBOM.AddItem Trim(RS!Work_Order)
        End If
     End If
      
     RS.MoveNext
Wend

End Function

Private Function GetMachineByWo(ByVal WO As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim Rev As String

Rev = ""
 
cboMachine.Clear
str = "select distinct Machine From QSMS_Wo    where work_order in (select wo from Sap_Wo_List where [Group]='" & TxtGroup & "') "
Set RS = Conn.Execute(str)
If RS.EOF Then
   MsgBox "can not find Machine ,Please check if the work order bom Check OK"
End If
While Not RS.EOF
      cboMachine.AddItem Trim(RS!Machine)
      RS.MoveNext
Wend
End Function

Private Function GetLine()
Dim str As String
Dim RS As ADODB.Recordset
str = "select distinct Line from QSMS_woGroup Order by Line"
Set RS = Conn.Execute(str)
CboLine.Clear
If Not RS.EOF Then
    RS.MoveFirst
    While Not RS.EOF
        CboLine.AddItem Trim(RS!Line)
        RS.MoveNext
    Wend
Else
    MsgBox "Didn't get the Line from QSMS_woGroup"
    Exit Function
End If
End Function

Private Function GetSBWO(ByVal WO As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim I As Long
Dim Group As String
I = 0
CboSBWO.Clear
FraSB.Visible = False
str = "Select [Group] from Sap_Wo_List where wo='" & WO & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   Group = Trim(RS!Group)
   TxtGroup = Group
End If
str = "select Wo from Sap_Wo_list where [Group] ='" & Group & "' and wo<>'" & WO & "' order by wo"
Set RS = Conn.Execute(str)
While Not RS.EOF
     CboSBWO.AddItem Trim(RS!WO)
     RS.MoveNext
     I = I + 1
Wend
If I > 0 Then
    FraSB.Visible = True

End If
End Function

Private Function ChkIfInCurretnFeeder(ByVal Feeder As String, ByVal DID As String) As Boolean
Dim str As String
Dim RS As ADODB.Recordset
ChkIfInCurretnFeeder = True
str = "Select * from QSMS_FeederDID_Current where Feeder='" & Feeder & "' or DID='" & DID & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
 
   ChkIfInCurretnFeeder = True
Else
   ChkIfInCurretnFeeder = False
End If
Set DgDIDSlot.DataSource = RS
DgDIDSlot.Refresh
End Function


