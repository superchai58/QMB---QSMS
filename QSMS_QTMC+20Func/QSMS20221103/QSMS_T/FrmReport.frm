VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmReport 
   Caption         =   "Report 2019-06-04"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   14790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Report Type"
      Height          =   1575
      Left            =   0
      TabIndex        =   23
      Top             =   5400
      Width           =   14775
      Begin VB.CheckBox chkDual 
         Caption         =   "Dual"
         Height          =   255
         Left            =   5400
         TabIndex        =   63
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "&Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6720
         Picture         =   "FrmReport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   180
         Width           =   975
      End
      Begin VB.ComboBox CboReportType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "FrmReport.frx":030A
         Left            =   1680
         List            =   "FrmReport.frx":030C
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lBLmESSAGE 
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   6255
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
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   9600
         Top             =   4920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Upload"
         Height          =   375
         Left            =   8400
         TabIndex        =   61
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtFilePath 
         Height          =   375
         Left            =   1440
         TabIndex        =   60
         Top             =   4800
         Width           =   5895
      End
      Begin VB.CommandButton CmdSFile 
         Caption         =   "SelectFiles"
         Height          =   375
         Left            =   7320
         TabIndex        =   59
         Top             =   4800
         Width           =   1095
      End
      Begin VB.ComboBox CboShift 
         Height          =   360
         ItemData        =   "FrmReport.frx":030E
         Left            =   1560
         List            =   "FrmReport.frx":0310
         TabIndex        =   58
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox txtDID 
         Height          =   375
         Left            =   1560
         TabIndex        =   56
         Top             =   4320
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox TxtJobGroup 
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
         Left            =   5880
         TabIndex        =   55
         Top             =   3840
         Width           =   2655
      End
      Begin VB.ListBox ListselectingJobGroup 
         Height          =   255
         ItemData        =   "FrmReport.frx":0312
         Left            =   12480
         List            =   "FrmReport.frx":0314
         TabIndex        =   52
         Top             =   3000
         Width           =   2175
      End
      Begin VB.CommandButton CmdAddGroup 
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
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3000
         Width           =   495
      End
      Begin VB.CommandButton cmdADDALLGroup 
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
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3480
         Width           =   495
      End
      Begin VB.CommandButton cmdDELGroup 
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
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3960
         Width           =   495
      End
      Begin VB.CommandButton cmdDELALLGroup 
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
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   4440
         Width           =   495
      End
      Begin VB.ListBox ListAllJobGroup 
         Height          =   255
         ItemData        =   "FrmReport.frx":0316
         Left            =   9600
         List            =   "FrmReport.frx":0318
         TabIndex        =   46
         Top             =   3000
         Width           =   2175
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
         Left            =   6720
         TabIndex        =   44
         Top             =   960
         Width           =   2655
      End
      Begin VB.ComboBox CboJobPN 
         Height          =   360
         Left            =   1560
         TabIndex        =   41
         Top             =   1920
         Width           =   2655
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   1680
         TabIndex        =   40
         Top             =   1440
         Width           =   1695
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
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2280
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
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1800
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
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1320
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
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   840
         Width           =   495
      End
      Begin VB.ListBox ListWoall 
         Height          =   255
         ItemData        =   "FrmReport.frx":031A
         Left            =   9600
         List            =   "FrmReport.frx":031C
         TabIndex        =   31
         Top             =   840
         Width           =   2175
      End
      Begin VB.ListBox ListWoSelecting 
         Height          =   255
         ItemData        =   "FrmReport.frx":031E
         Left            =   12480
         List            =   "FrmReport.frx":0320
         TabIndex        =   30
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox TxtRev 
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
         Left            =   5880
         TabIndex        =   29
         Top             =   3360
         Width           =   2655
      End
      Begin VB.ComboBox CboMachine 
         Height          =   360
         Left            =   5880
         TabIndex        =   20
         Top             =   1920
         Width           =   2655
      End
      Begin VB.ComboBox CboComp 
         Height          =   360
         Left            =   5880
         TabIndex        =   19
         Top             =   2400
         Width           =   2655
      End
      Begin VB.ComboBox CboWo 
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
         Left            =   6720
         TabIndex        =   9
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox TxtMBPN 
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
         Left            =   5880
         TabIndex        =   8
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox TxtWOQty 
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
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   2880
         Width           =   2655
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "Group"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton OptRelease 
         Caption         =   "Release"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
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
         Height          =   975
         Left            =   3480
         Picture         =   "FrmReport.frx":0322
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   6720
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox TxtWO 
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
         Left            =   1560
         TabIndex        =   2
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox TxtCustomer 
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
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   3360
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   10
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
         Format          =   81002499
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   38
         Top             =   960
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
         Format          =   81002499
         CurrentDate     =   36482
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FilePath"
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
         Left            =   120
         TabIndex        =   62
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "Shift"
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
         Left            =   120
         TabIndex        =   57
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         Index           =   7
         Left            =   4560
         TabIndex        =   54
         Top             =   3840
         Width           =   1335
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
         Index           =   6
         Left            =   120
         TabIndex        =   53
         Top             =   4320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Selecting JobGroup"
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
         Height          =   255
         Index           =   5
         Left            =   12480
         TabIndex        =   51
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "All JobGroup"
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
         Height          =   255
         Index           =   4
         Left            =   9600
         TabIndex        =   45
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
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
         Index           =   4
         Left            =   4560
         TabIndex        =   43
         Top             =   960
         Width           =   2175
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
         Index           =   5
         Left            =   120
         TabIndex        =   42
         Top             =   1920
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
         Index           =   4
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "All Work Order"
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
         Height          =   255
         Index           =   2
         Left            =   9600
         TabIndex        =   37
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Selecting  WO"
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
         Height          =   255
         Index           =   3
         Left            =   12480
         TabIndex        =   36
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Revision"
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
         Left            =   4560
         TabIndex        =   28
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   4560
         TabIndex        =   22
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "CompPN"
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
         Left            =   4560
         TabIndex        =   21
         Top             =   2400
         Width           =   1335
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
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "Check BOM OK"
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
         Left            =   4560
         TabIndex        =   17
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "MB/Job PN"
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
         Left            =   4560
         TabIndex        =   16
         Top             =   2880
         Width           =   1335
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
         Top             =   1440
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
         Top             =   2880
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
         Left            =   4560
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
         Top             =   2400
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
         TabIndex        =   11
         Top             =   3360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/**********************************************************************************
'**文 件 名: FrmReport
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: Sandy Sun
'**日    期: 2007.11.20
'**描    述: QSMS Report
'
'** EQMS             修 改 人     修改日期        描    述
 '--------------------------------------------------------------------------------------------------------------------------------
                    '**Sandy      2007.11.20     在DID Dispatch report里增加超发材料的项目--------(0001)
                    '**Kane       2007/12/11     增加祥龙计划计算需求数量报表---------------------(0002）
                    '**Jing       2007.12.12     Add a report for WoInputPlan-------------------(0003)
                    '**Kane       2007/12/25     增加祥龙计划QSMS_WO需求数量报表------------------(0004）
                    '**Jing       2007.12.26     Update Report for DispatchDID(从未使用的)-------(0005)
                    '**Jing       2008.01.08     Add a report for NOUseDID----------------------(0006)
                    '**Jing       2008.01.09     Mark (0006)------------------------------------(0007)
                    '**Sandy      2008.01.09     Can not check Bom for one WO at the same time for multi-user, only one by one;(0008)
                    '**Sandy      2008.01.16     add PrepariMaterialMonitor report-------------(0009)
                    '**Kane       2008.01.21     增加by线别查询祥龙计划发料状况报表(0010)
                    '**Sandy      2008.01.24     增加 ReturnDID to WH 状况报表----------------(0011)
                    '**Sandy      2008.02.02     add XL_ReelBaseQty report----------------(0012)
                    '**Giant      2008.02.19     add add check forbidden PN function----------------(0013)
                    '**Sandy      2008.03.05     add ReturnDIDByGroupID and ReturnDIDByWO (00014)
                    '**Archer     2008.03.13     modify the Report of XL_MaterialDemand (0015)
                    '**Jing       2008.03.18     check wo if it was closed when check Bom   (0016)
                    '**Kane       2008.03.25     Query dispatchdid information by date & shift & groupid --(0017)
                    '**Kane       2008.04.02     当有新的ReplacePN上传时，对可能产生影响的工单重新CheckBom PASS后纪录下来--(0018)
                    '**Udall      2008.04.07     From call function to call stored procedures when check bom and move the GetCheckBomData function to ChkBom Modules--(0019)
                    '**Kane       2008.05.27     Get all dispatch information both in current & history database by groupid ---(0020)
                    '**Jing       2008.06.01     Query MEBom by JobPN+Revision  (0021)
                    '**Sandy      2008.06.02     add a new report type: WoInputPlanBySide to get the  WO Plan  by Side  (0022)
                    '**Udall      2008.06.10     Update the SAP1 and SAP2 Report for NB5  (0023)
                    '**Udall      2008.06.13     Update the Report for get the ReplacePN(Update JobPN function)  (0024)
                    '**Udall      2008.06.16     Add a new report GetSapCostSum data  (0025)
                    '**Sandy      2008.06.26     应NB5的要求将EndDay的默认日期更为当天的下一天；（0026）
                    '**Kane       2008.08.18     Add new report for DID compare (0027)
                    '**Udall      2008.08.28     Add new report for GroupID Cost CompPN Qty (0028)
                    '**Udall      2008.09.28     Add new report for get GroupID data by CompPN (0029)
                    '**Kane       2008.10.14     getdata by SP for DispatchDID report '(0030)
                    '**Kane       2008.10.28     Add new report for check splice replace pn data'(0031)
                    '**Kane       2008.11.07     Add new report for forbidden pn '(0032)
                    '**Kevin      2008.12.08     Add new Function for query Sapbom Report (0033)
                    '**Kevin      2008.12.23     Add new Function for query qsms_wo by wo list(0034)
                    '**Sandy      2009.01.06     add new report for MaterialReturn (0035)
                    '**Sandy      2009.01.20     add new report for RPTGlue_Consumptio (0036)
                    '**Sandy      2009.01.20     add new report for Glue_DATEBYDAY(0037)
                    '**Giant      2009.03.19     add new report for Glue_CallOffDATa (0038)
                    '**Salon      2009.03.26     From:"\PrepareMaterialReport.xls" to "\Template\PrepareMaterialReport.xls" (0039)
                    '**Kevin      2009.04.27     add new report for PanalnterLock (0040)
                    '**Kane       2009.08.13     add report get fuji avl data by wo(0041)
''QMS                **Sandy      2009.08.12     check CompPN in JobPN (0059)
''RQ09101401         **Lynn       2009.10.23     add CheckBom_log Query tool (0060)
''RQ09101401         **Lynn       2009.10.23     add CheckBom log on confirm (0061)
''RQ09092203         **Sandy      2009.11.02     比对料头材料之间的差异(0062)
''QMS                **Austin     2009.11.10     DispatchQTYByWO没有资料的时候，返回信息改为"NO Data" 之前返回信息"refresh bom ok" 不准确 (0063)
''QMS                **Austin     2009/12/21     记录CheckBOM的结果信息,保存到QSMS_Log,并通过CheckBOM_Result查询        (0073)
''QMS                **Kane       2010/01/22     下拉列表在下拉时增加宽度显示所有信息'(0074)
''QMS                **Sandy      2010/02/02     Add Tips for 点击CboReportType可以看到定义该FuncType有何作用 （0075）
''QMS                **Feix       2017/07/25     添加导出条件Slot不等于空（1260）
'**********************************************************************************/

Dim mclsToolTip As New clsToolTip
Dim arryTipData() As String
Dim strAddress As String
Private Sub CboGroupID_Click()
Call GetGroupWO(CboGroupID)

End Sub

Private Sub CboGroupID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboGroupID_Click
End If
End Sub



Private Sub CboJobPN_Click()
Call GetGroupID(cboJobPN)
Call GetJobGroupByJobRev("", Trim(TxtMBPN), Trim(txtRev))
End Sub

Private Sub CboJobPN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call GetGroupID(cboJobPN)
End If
End Sub

Private Sub CboLine_Click()

Dim str As String
Dim BeginDate, EndDate As String
Dim GroupIDHead As String
Dim I As Long
Dim RS As ADODB.Recordset
BeginDate = Format(dtpSDate, "YYYY/MM/DD")
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
EndDate = Format(dtpEDate, "YYYY/MM/DD")
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")

    If OptRelease.Value = True Then
       str = "select distinct c.jobpn,b.Mb_Rev from QSMS_WOGroup a ,sap_wo_list b,qsms_JobBOM c  where " & _
       "a.WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and a.line='" & CboLine & "'" & _
       " and a.work_Order=b.wo and a.work_order=c.work_order order by c.jobpn,b.mb_rev"
    Else
        str = "select distinct c.jobpn,b.Mb_Rev  from QSMS_WOGroup a ,sap_wo_list b,qsms_JobBOM c where" & _
        " substring(a.GroupID,2,8) between '" & BeginDate & "' and '" & EndDate & "' and a.line='" & CboLine & "'" & _
        " and a.work_Order=b.wo and a.work_order=c.work_order  order by c.jobpn,b.mb_rev"
    End If


Set RS = Conn.Execute(str)
cboJobPN.Clear
While Not RS.EOF
     cboJobPN.AddItem Trim(RS!Jobpn) & "-" & Trim(RS!Mb_Rev)
     RS.MoveNext
Wend
End Sub

Private Sub CboLine_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CboLine_Click
End If
End Sub

Private Sub CboMachine_Click()
Call GetComp(Trim(CboMachine), Trim(TxtMBPN), Trim(txtWO), Trim(CboLine))
Call GetJobGroupByJobRev(Trim(CboMachine), Trim(TxtMBPN), Trim(txtRev))
End Sub

Private Sub CboMachine_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboMachine_Click
End If
End Sub

Private Sub CboNotChkBOM_Click()
txtWO = Trim(CboNotChkBOM)
Call GetWoinfo(txtWO)
Call GetMachine(txtWO)
End Sub

Private Sub CboNotChkBOM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboNotChkBOM_Click
End If
End Sub



Private Sub CboReportType_Click()
    

Dim RS As New ADODB.Recordset
Dim k As Integer
    If CboReportType = "DIDCallBack" Then
        frmQueryDIDCallBack.Show
    End If

    If Trim(CboReportType) = "" Then Exit Sub
    For k = 1 To UBound(arryTipData)           ''''(00075)
        If CboReportType.Text = arryTipData(k, 1) Then
            mclsToolTip.ToolText(CboReportType) = arryTipData(k, 2) + vbCrLf + vbCrLf + vbCrLf
            Exit For
        Else
            mclsToolTip.ToolText(CboReportType) = ""
        End If
    Next k
End Sub

Private Sub CboWo_Click()
txtWO = Trim(cboWO)
Call GetWoinfo(txtWO)
Call GetMachine(txtWO)
End Sub

Private Sub cboWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboWo_Click
End If
End Sub

Private Sub CmdADD_Click()
    Dim Pointer As Integer
    If ListWoall.ListCount <= 0 Then Exit Sub
    If ListWoall.ListIndex < 0 Then Exit Sub
    Pointer = ListWoall.ListIndex
    ListWoSelecting.AddItem Trim(ListWoall.Text)
    ListWoall.RemoveItem Pointer
    If ListWoall.ListCount <> Pointer Then
       ListWoall.ListIndex = Pointer
    End If
   
End Sub

Private Sub cmdADDALL_Click()

    If ListWoall.ListCount <= 0 Then Exit Sub

    Do While ListWoall.ListCount > 0
     
      ListWoall.ListIndex = 0
      ListWoSelecting.AddItem Trim(ListWoall.Text)
      ListWoall.RemoveItem 0
     
    Loop
   
End Sub

Private Sub cmdDel_Click()
    Dim Pointer As Long
    If ListWoSelecting.ListCount <= 0 Then Exit Sub
    If ListWoSelecting.ListIndex < 0 Then Exit Sub
    Pointer = ListWoSelecting.ListIndex

    ListWoall.AddItem Trim(ListWoSelecting.Text)
    ListWoSelecting.RemoveItem Pointer
    If ListWoSelecting.ListCount <> Pointer Then
       ListWoSelecting.ListIndex = Pointer
    End If

   
    
End Sub

Private Sub cmdDELALL_Click()
    If ListWoSelecting.ListCount <= 0 Then Exit Sub
    Do While ListWoSelecting.ListCount > 0
        ListWoSelecting.ListIndex = 0
       
        ListWoall.AddItem Trim(ListWoSelecting.Text)
        ListWoSelecting.RemoveItem 0
  
    Loop
    
End Sub




Private Sub cmdADDGroup_Click()
    Dim Pointer As Integer
    If ListAllJobGroup.ListCount <= 0 Then Exit Sub
    If ListAllJobGroup.ListIndex < 0 Then Exit Sub
    Pointer = ListAllJobGroup.ListIndex
    ListselectingJobGroup.AddItem Trim(ListAllJobGroup.Text)
    ListAllJobGroup.RemoveItem Pointer
    If ListAllJobGroup.ListCount <> Pointer Then
       ListAllJobGroup.ListIndex = Pointer
    End If
   
End Sub

Private Sub cmdADDALLGroup_Click()

    If ListAllJobGroup.ListCount <= 0 Then Exit Sub

    Do While ListAllJobGroup.ListCount > 0
     
      ListAllJobGroup.ListIndex = 0
      ListselectingJobGroup.AddItem Trim(ListAllJobGroup.Text)
      ListAllJobGroup.RemoveItem 0
     
    Loop
   
End Sub

Private Sub cmdDELGroup_Click()
    Dim Pointer As Long
    If ListselectingJobGroup.ListCount <= 0 Then Exit Sub
    If ListselectingJobGroup.ListIndex < 0 Then Exit Sub
    Pointer = ListselectingJobGroup.ListIndex

    ListAllJobGroup.AddItem Trim(ListselectingJobGroup.Text)
    ListselectingJobGroup.RemoveItem Pointer
    If ListselectingJobGroup.ListCount <> Pointer Then
       ListselectingJobGroup.ListIndex = Pointer
    End If

   
    
End Sub

Private Sub cmdDELALLGroup_Click()
    If ListselectingJobGroup.ListCount <= 0 Then Exit Sub
    Do While ListselectingJobGroup.ListCount > 0
        ListselectingJobGroup.ListIndex = 0
       
        ListAllJobGroup.AddItem Trim(ListselectingJobGroup.Text)
        ListselectingJobGroup.RemoveItem 0
  
    Loop
    
End Sub

Private Sub CmdExcel_Click()
Dim str As String
Dim RS As ADODB.Recordset
Dim I As Long
Dim WO As String
Dim strSQL As String
Dim DualModel As String
Dim tmpRS As New ADODB.Recordset

'On Error GoTo errHandler

Select Case CboReportType
        Case "PrepariMaterialList"
            If Trim(CboMachine) = "" Then
               MsgBox "Please select Machine"
               Exit Sub
            End If
            If ListWoSelecting.ListCount <= 0 Then
                MsgBox "Please select the wo to listbox---wo selecting"
                Exit Sub
            End If
            For I = 0 To ListWoSelecting.ListCount - 1
                ListWoSelecting.ListIndex = I
                WO = WO & Trim(ListWoSelecting.Text) & ","
            Next I
            WO = Mid(WO, 1, Len(WO) - 1)
            Call CopyToExcelPrepareMaterialList("PrepariMaterialList", Trim(WO), Trim(CboMachine), Trim(CboLine))
              
         Case "PrepariMaterialByJobPN"
            If ListWoSelecting.ListCount <= 0 Then
                MsgBox "Please select the wo to listbox---wo selecting"
                Exit Sub
            End If
            Call PrepareMaterialByWONew("By_JobPN", Trim(CboLine)) ''1169
        Case "QSMS_CheckCompPN"
                Call Load_QSMS_CheckCompPN(1)
        Case "CheckReplacePNBySAPBOM" ''0062
                Call Load_CheckReplacePNBySAPBOM(1)
        Case "PrepariMaterialByLineShift"                   ''# 2008-01-18 add by archer
            If Trim(CboShift) = "" Then
               MsgBox "Please select the shift"
               Exit Sub
            End If
            If Trim(CboLine) = "" Then
               MsgBox "Please select the Line"
               Exit Sub
            End If
            If (Trim(dtpSDate) > Trim(dtpEDate)) Then
               MsgBox "Please Select BeginDate and EndDate Again"
               Exit Sub
            End If
            Call PrepareMaterialByWONew("By_Shift", Trim(CboLine)) ''1169
              
        Case "PrepariMaterialByGroup"
            Call PrepareMaterialByWONew("By_Group", Trim(CboLine)) ''1169
            
        Case "PrepariMaterialByWos"
            Call PrepareMaterialByWONew("By_WorkOrders", Trim(CboLine)) ''1169
            
        Case "PrepariMaterialByWo"
            Call PrepareMaterialByWONew("By_WorkOrder", Trim(CboLine)) ''1169
            
        Case "LineChangeStatisticsByall"
            Call LineChangeStatisticsByall
            
        Case "DIDDeleteRecords"
            Call DIDDeleteRecords
        
        Case "LineChangeStatistics"
            Call LineChangeStatistics
            
        Case "PrepariMaterialMonitor" 'add PrepariMaterialMonitor report-------------(0009)
            Call XL_MonitorReport

        Case "DispatchQTYByWO" '**Sandy      2008.01.24     增加 ReturnDID to WH 状况报表----------------(0011)
            Call DispatchQTYByWO
        Case "QSMS_DID_ToWH" '**Sandy      2008.01.24     增加 ReturnDID to WH 状况报表----------------(0011)
            Call QSMS_DID_ToWH
        Case "QSMS_WO"
            Call QSMS_WO
        Case "WipByMaterial"
            If Trim(CboComp) = "" Then
               MsgBox "CompPN can not be empty,Please check"
               Exit Sub
            End If
            Call CopyToExcelWipByMaterial("WipByMaterial", Trim(CboComp))
        Case "WipByDate"
            Call CopyToExcelWipByDate("WipByDate")
            
        Case "WipByGroup"
            If Trim(CboGroupID) = "" Then
                MsgBox "Please check the GroupID"
                Exit Sub
            End If
            Call CopyToExcelWipByGroup("WipByGroup", Trim(CboGroupID))
            
        Case "WipLackbyWo"
              If Trim(cboWO) = "" Then
                 MsgBox "Please check the WO"
                 Exit Sub
              End If
              Call CopyToExcelWipLackByWo("WipLackbyWo", Trim(txtWO))
              
        Case "MaterialDifferentList"
            If Trim(cboWO) = "" Then
              MsgBox "Please check the WO"
              Exit Sub
            End If
            Call CopyToExcelWipDifferentMaterial("MaterialDifferentList", Trim(txtWO))
        Case "PDUsedByCompLine"
            Call PDUsedByCompLine(CboLine, CboComp)
        Case "SAP_BOM"
               Call GetSapBom(Trim(txtWO))
        Case "SAP_GroupByWo"
               Call GetSapGroupByWo(Trim(txtWO))
        Case "ME_BOM"
               Call GetMEBom(Trim(txtWO))
        Case "CheckWO_WastagePN"                    ''1212
               Call CheckWOWastagePN(Trim(txtWO))
        Case "GetGroupIDDataByCompPN"               ''0029
               If Trim(CboGroupID) = "" Then
                 MsgBox "Please input the GroupID"
                 CboGroupID.SetFocus
                 Exit Sub
               End If
               If Trim(CboComp) = "" Then
                 MsgBox "Please input the CompPN"
                 CboComp.SetFocus
                 Exit Sub
               End If
               Call GETGROUPIDDATABYCOMPPN(Trim(CboGroupID), Trim(CboComp))
        Case "MEBOM_Delete_Log"
               Call GetMEBom_DeleteLog(Trim(TxtMBPN))
        Case "ME_BOM_WO"
                Call GetMEBom_WO(Trim(txtWO))
        Case "ReplacePN" '1059
                If Trim(txtWO) <> "" Then
                  Call GetReplacePN(Trim(txtWO))
                ElseIf ListWoSelecting.ListCount <= 0 Then
                    MsgBox "Please select the wo to listbox---wo selecting,or upload the wolist"
                    Exit Sub
                Else
                    For I = 0 To ListWoSelecting.ListCount - 1
                        ListWoSelecting.ListIndex = I
                        WO = WO & Trim(ListWoSelecting.Text) & ","
                    Next I
                    WO = Mid(WO, 1, Len(WO) - 1)
                    Call GetReplacePN(Trim(WO))           ''''(0024)
                End If
        Case "CheckBOM"
            ''''''added by Jing 2008.03.18  (0016)
            If chkDual = 1 Then    ''(1179)
                DualModel = "Y"
            Else
                DualModel = "N"
            End If
            
            
        ''''1279
           strSQL = "Exec QSMS_CheckBOM_CheckCycleTime '" & Trim(txtWO) & "'"
           Set tmpRS = Conn.Execute(strSQL)
           If Not tmpRS.EOF Then
               If tmpRS("Result") <> 0 Then
                   MsgBox ("" & Trim(tmpRS("Descr")) & "")
               End If
           End If
            
            
            
            strSQL = "select * from qsms_wogroup where work_order='" & Trim(txtWO) & "' and ClosedFlag='Y'"
            Set tmpRS = Conn.Execute(strSQL)
            If tmpRS.EOF = False Then
                MsgBox ("This wo had been closed !"), vbCritical
                Exit Sub
            End If
            BomTest = Trim(txtRev)
            If CheckBomLogon = "Y" Then   ''(0061)
                If CheckBomRight = False Then
                    MsgBox "You have no right to check bom!", vbCritical
                    Exit Sub
                Else
                    Call GetCheckBomData(Trim(txtWO), Trim(g_userName), DualModel)    ''(1179)
                End If
            Else
                Call GetCheckBomData(Trim(txtWO), Trim(g_userName), DualModel)             ''(0019)  ''(1179)
            End If
        Case "CheckBOMDiff"
               Call GetChkBOMDiff(Trim(txtWO))
        Case "RefreshBOM"
              Call RefreshBoM(Trim(txtWO))
        Case "CheckBOM_Rate"
              Call CheckBOM_Rate
''        Case "DeleteME_BOM"
''              Call DeleteME_BOM(Trim(TxtMBPN), Trim(TxtRev), Trim(CboMachine.Text))
        
        Case "SAPCostSum", "SAP1", "SAP1His", "SAP2", "ReturnDID", "ReturnDID_ByDate", "DispatchDID", "Return_Dispatch", "DIDCallBack", "SAPFileChk", "CastQty", "WO_SingleCompPNData", "GroupIDCostQty"
              Call Sap_Return(Trim(CboReportType))
        Case "ReturnDIDByGroupID"
            Call ReturnDID
        Case "ReturnDIDByWO"
             Call ReturnDID
        Case "AOIQtySummary" '1059
                If Trim(txtWO) <> "" Then
                  Call AOIQtySummary(Trim(txtWO))
                ElseIf ListWoSelecting.ListCount <= 0 Then
                    MsgBox "Please select the wo to listbox---wo selecting,or upload the wolist"
                    Exit Sub
                Else
                    For I = 0 To ListWoSelecting.ListCount - 1
                        ListWoSelecting.ListIndex = I
                        WO = WO & Trim(ListWoSelecting.Text) & ","
                    Next I
                    WO = Mid(WO, 1, Len(WO) - 1)
                    Call AOIQtySummary(Trim(WO))
                End If
        Case "AOIDetail"
              Call AOIDetail(Trim(txtWO))
        Case "MachineType"
            Call MachineType
        Case "TraySlot"
            Call TraySlot
        Case "VerifyReport"
              Call ToExcelVerifyReport
        Case "VerifyReportWOChged"
              Call ToExcelWOChged
        Case "VerifyJobFailLog"
              Call ToExcelVerifyJobFailLog
        Case "UnDispatchList"
            Call GetUnDispatchList(Trim(txtWO))
        Case "NonAVL"
            Call NonAVL(Trim(CboComp))
        Case "CastRate"
            Call CastRate
        Case "OneByOne"
            Call OneByOne
        Case "SameGroupWO"
            Call SameGroupWO(Trim(txtWO))
        Case "CompPN_DIDData"
            Call GetCompPNDIDData(Trim(CboComp))
        Case "CompPNQty"
            Call GetCompPNQty(Trim(CboComp))
        Case "CheckDispatchQty"
            Call CheckDispatchQty(Trim(txtWO))
        Case "UnCloseGroupID"
            Call CheckUnCloseGroupID
        Case "XL_MaterialDemand" '-------------------------0002
            Call XL_MaterialDemand '''(0015)
        Case "WoInputPlan"
            Call GetWoInputPlan
        Case "WoInputPlanBySide"
            Call GetWoInputPlanBySide
        Case "XL_DemandDetail" '-------------------------0004
            Call XL_DemandDetail
        Case "XL_DispatchStatus" '-------------(0010)
            Call XL_DispatchStatus
        Case "XL_ReelBaseQty" '-------------(0011)
            If Trim(CboComp) = "" Then
               MsgBox "CompPN can not be empty,Please check"
               Exit Sub
            End If
            Call XL_ReelBaseQty
        Case "AllDispatchByGroupID"  '---- (0020)
            Call GetAllDispatchInforByGroupID(CboGroupID)
'        Case "Report_NoUseDID"
'            Call GetNoUseDID

        '''(0021)
        Case "MEBom_Model"
            Call MEBom_Model(Trim(TxtMBPN), Trim(txtRev))
        Case "DIDCompare"
            Call DIDCompare
        Case "CheckSpliceReplacePN"
            Call CheckSpliceReplacePN
        Case "ForbiddenPN" '(0032)
            Call ForbiddenPN
        Case "Glue_Consumption" '(0032)
            Call Glue_Consumption
        Case "Glue_DataByDay" '(0037)
            Call Glue_DataByDay
        Case "Glue_CallOff" '(0038)
            Call Glue_CallOff
        Case "MaterialReturn" '(0035)
            Call MaterialReturn
        Case "PanalnterLock"  '(0040)
            Call PanalnterLock
        Case "FUJI_AVLList" '0041
            Call FUJI_AVLList
        Case "CheckBom_Log" ''(0060)
            Call CheckBom_Log(Trim(txtWO))
        Case "CheckBom_Result"   ''0073
            Call CheckBom_Result(Trim(txtWO))
        Case "DIDIntegration"
            Call DIDIntegration(Trim(CboGroupID.Text))
        Case "SpliceReplacePN"                '（1077）
            Call SpliceReplacePN
        Case "SplicePN"                       '（1077）
            Call SplicePN
        Case "MaintainFeeder"                       '（1082）
            Call MaintainFeeder
        Case "PDA_DistributeDIDLog"                       '（1089）
            Call PDA_DistributeDIDLog
        Case "ME_BOM_GroupID"
            Call GetMEBom_ByGroupID(Trim(CboGroupID.Text))   '(1127)
        Case "MEBom_EQProgram"
            Call GetMEBom_EQProgram(Trim(TxtJobGroup))     '(1219)
        Case Else
             MsgBox "Please select the function type"

             
End Select
Exit Sub

'errHandler:
'    MsgBox ("cmdExcel_Click, " & Err.Description)
    
End Sub

Private Sub CmdQuery_Click()
If Trim(CboLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If
Call GetGroupID("")
Call GetJobPN
End Sub

'Add by Kevin 2008.12.08 (0033)
Private Sub CmdSFile_Click()
    CommonDialog1.ShowOpen
    txtFilePath = CommonDialog1.FileName
End Sub

'Add by Kevin 2008.12.08 (0033)
Private Sub cmdUpload_Click()

Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rCount As Long


    If Trim(txtFilePath) = "" Then
       Exit Sub
    End If
    If CboReportType.Text = "QSMS_CheckCompPN" Or CboReportType.Text = "CheckReplacePNBySAPBOM" Then
        Call CmdExcel_Click
    Else
        Set xlApp = CreateObject("Excel.Application")
        Let xlApp.Visible = False
        Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
        xlApp.DisplayAlerts = False
    
        rCount = 2
        ListWoSelecting.Clear
        'strWO = ""
        With xlsBook.Worksheets(1)
            While Trim(.Cells(rCount, 1)) <> ""
                ListWoSelecting.AddItem (.Cells(rCount, 1))
                'StrWO = StrWO & ",''" & ListWoSelecting.List(rCount - 2) & "''"
                rCount = rCount + 1
                DoEvents
            Wend
        End With
        'StrWO = Mid(StrWO, 2, Len(StrWO))
        xlsBook.Close
        xlApp.Quit
        Set xlApp = Nothing
        Set xlsBook = Nothing
    End If
End Sub
Private Sub Load_QSMS_CheckCompPN(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rsTmp As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim strSQL As String, strJobPN As String, strCompPN As String, strUID As String
Dim tmpRow As Integer
Dim blerr As Boolean
blerr = False

On Error GoTo errhandle:
    
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    
    strSQL = "Truncate table QSMS_CompPNcheck_Temp"
    Set rsTmp = Conn.Execute(strSQL)
    
    tmpRow = 2
    strUID = Trim(g_userName)
    
    With xlsBook.Worksheets(1)
        While Trim(.Cells(tmpRow, 1)) <> "" And Trim(.Cells(tmpRow, 2)) <> ""
            strJobPN = Trim(.Cells(tmpRow, 1))
            strCompPN = Trim(.Cells(tmpRow, 2))
            
            strSQL = "EXEC QSMS_CheckCompPN '" & strJobPN & "','" & strCompPN & "','" & strUID & "','W'"
            Set rsTmp = Conn.Execute(strSQL)
            
            If Trim(rsTmp("result")) <> "0" Then
                MsgBox ("Err: " + rsTmp("desc1"))
                GoTo NormalHandle
            End If
            tmpRow = tmpRow + 1
        Wend
    End With
    
strSQL = "EXEC QSMS_CheckCompPN "
Set RS = Conn.Execute(strSQL)
If Not RS.EOF Then
       Call CopyToExcel(RS)
    Else
       MsgBox ("No Data"), vbCritical      '0063
End If

NormalHandle:
    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox ("*** Load  Finish ! ***")
    Exit Sub
    
errhandle:
    MsgBox Err.Description
End Sub
Private Sub Load_CheckReplacePNBySAPBOM(Shift_Item As String) ''0062
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rsTmp As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim strSQL As String, strJobPN As String, strCompPN As String, strUID As String
Dim tmpRow As Integer
Dim blerr As Boolean
blerr = False

On Error GoTo errhandle:
    
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    
    strSQL = "Truncate table QSMS_CompPNcheck_Temp"  ''(1110)
    Set rsTmp = Conn.Execute(strSQL)
    
    tmpRow = 2
    strUID = Trim(g_userName)
    
    With xlsBook.Worksheets(1)
        While Trim(.Cells(tmpRow, 1)) <> ""
            strCompPN = Trim(.Cells(tmpRow, 1))
            strSQL = "select top 1 * from sapbom where MBPN='" & strCompPN & "'"
            Set rsTmp = Conn.Execute(strSQL)
            If rsTmp.EOF Then
                MsgBox ("this MBPN '" & strCompPN & "' have not exist in SAPBOM,please chceck it;")
                Exit Sub
            End If
            strSQL = "insert into QSMS_CompPNcheck_Temp(JobPN, CompPN, UserID, TransDateTime) select '','" & strCompPN & "','',''"  '''(1110)
            Set rsTmp = Conn.Execute(strSQL)
            tmpRow = tmpRow + 1
        Wend
    End With
    
strSQL = "EXEC QSMS_QueryReplacePN "
Set RS = Conn.Execute(strSQL)
If Not RS.EOF Then
       Call CopyToExcel(RS)
    Else
       MsgBox ("Those PN have no different ReplacePN!"), vbCritical
End If

NormalHandle:
    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    Exit Sub
    
errhandle:
    MsgBox ("*** Load  Finish ! ***")
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
Dim str As String
Dim RS As ADODB.Recordset
Dim lRetVal As Long
Dim k As Integer
Dim ctrl As Control
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
k = 1
strSQL = "Select [Key],Value from B_ToolTip_Config Where Station='QSMS' order by [Key]"        '''(0075)
Set RS = Conn.Execute(strSQL)
ReDim arryTipData(RS.RecordCount, 2) As String
If RS.EOF = False Then
    While Not RS.EOF
        arryTipData(k, 1) = Trim(RS!key)
        arryTipData(k, 2) = Trim(RS!Value)
        Debug.Print arryTipData(k, 1)
        Debug.Print arryTipData(k, 2)
        k = k + 1
        RS.MoveNext
    Wend
End If

With mclsToolTip
    Call .Create(Me)
    .MaxTipWidth = 240
    .DelayTime(ttDelayShow) = 20000
    For Each ctrl In Controls
        If TypeOf ctrl Is ComboBox Then
            Call .AddTool(ctrl)
        End If
    Next
End With
'''''''''''''''''''''''''''''''''''''''''''''''''''



lRetVal = SendMessage(CboReportType.hWnd, CB_SETDROPPEDWIDTH, 300, 0) '(0074)
str = "select getdate()"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
    'Date = Rs(0)
    'Time = Rs(0)
End If

dtpSDate = Date
dtpEDate = Date + 1 '(0026)
CboShift.AddItem "Day_Shift"
CboShift.AddItem "Night_Shift"

''1168
Call GetReportType
''CboReportType.AddItem "PrepariMaterialList"
''CboReportType.AddItem "QSMS_CheckCompPN"
''CboReportType.AddItem "Glue_Consumption"
''CboReportType.AddItem "Glue_CallOff"
''CboReportType.AddItem "Glue_DataByDay"
''CboReportType.AddItem "DispatchQTYByWO"
''CboReportType.AddItem "PrepariMaterialByGroup"
''CboReportType.AddItem "PrepariMaterialByWo"
''CboReportType.AddItem "PrepariMaterialByLineShift"          ''# 2008-01-18 add by archer
''CboReportType.AddItem "ReturnDIDByGroupID"
''CboReportType.AddItem "ReturnDIDByWO"
''CboReportType.AddItem "PrepariMaterialByWos"
''CboReportType.AddItem "PrepariMaterialByJobPN"
'CboReportType.AddItem "CheckReplacePNBySAPBOM" ''0062
''CboReportType.AddItem "PrepariMaterialMonitor"
''CboReportType.AddItem "WipByDate"
''CboReportType.AddItem "WipByGroup"
''CboReportType.AddItem "WipLackbyWo"
''CboReportType.AddItem "PDUsedByCompLine"
''CboReportType.AddItem "QSMS_DID_ToWH"
''CboReportType.AddItem "QSMS_WO"
''CboReportType.AddItem "SAP_BOM"
''CboReportType.AddItem "SAP_GroupByWo"
''CboReportType.AddItem "ME_BOM"
''CboReportType.AddItem "MEBOM_Delete_Log"
''CboReportType.AddItem "ME_BOM_WO"
''CboReportType.AddItem "MaterialReturn"
''CboReportType.AddItem "ReplacePN"
''CboReportType.AddItem "CheckBOM"
''CboReportType.AddItem "CheckBOMDiff"
''CboReportType.AddItem "RefreshBOM"
''CboReportType.AddItem "CheckBOM_Rate"
''CboReportType.AddItem "CheckDispatchQty"
'''''''CboReportType.AddItem "DeleteME_BOM"
''CboReportType.AddItem "SAPCostSum"
''CboReportType.AddItem "GroupIDCostQty"
''CboReportType.AddItem "GetGroupIDDataByCompPN"
''CboReportType.AddItem "SAP1His"
''CboReportType.AddItem "SAP1"
''CboReportType.AddItem "SAP2"
''CboReportType.AddItem "SAPFileChk"
''CboReportType.AddItem "CastQty"
''CboReportType.AddItem "AOIQtySummary"
''CboReportType.AddItem "AOIDetail"
''CboReportType.AddItem "DIDCallBack"
''CboReportType.AddItem "ReturnDID"
''CboReportType.AddItem "DispatchDID"
''CboReportType.AddItem "Return_Dispatch"
''CboReportType.AddItem "MachineType"
''CboReportType.AddItem "VerifyReport"
''CboReportType.AddItem "VerifyReportWOChged"
''CboReportType.AddItem "VerifyJobFailLog"
''CboReportType.AddItem "TraySlot"
''CboReportType.AddItem "CastRate"
''CboReportType.AddItem "OneByOne"
''CboReportType.AddItem "UnDispatchList"
''CboReportType.AddItem "NonAVL"
''CboReportType.AddItem "SameGroupWO"
''CboReportType.AddItem "CompPN_DIDData"
''CboReportType.AddItem "CompPNQty"
''CboReportType.AddItem "LineChangeStatistics"
''CboReportType.AddItem "LineChangeStatisticsByall"
''CboReportType.AddItem "UnCloseGroupID"
''CboReportType.AddItem "DIDDeleteRecords"
''CboReportType.AddItem "XL_MaterialDemand"
''CboReportType.AddItem "WoInputPlan"
''CboReportType.AddItem "WO_SingleCompPNData"
''CboReportType.AddItem "DIDCompare"
''CboReportType.AddItem "CheckSpliceReplacePN"
''CboReportType.AddItem "PanalnterLock"
''CboReportType.AddItem "FUJI_AVLList"
''CboReportType.AddItem "CheckBom_Log"
''CboReportType.AddItem "CheckBom_Result"  ''0073
''CboReportType.AddItem "DIDIntegration"
''str = "Select Site from Site"
''Set rs = Conn.Execute(str)
''If rs.EOF = False Then
    ''If rs.Fields(0) = "ES" Or rs.Fields(0) = "ESBU" Or rs.Fields(0) = "CC" Then
        ''CboReportType.AddItem "WoInputPlanBySide" '(0022)
    ''End If
''End If

''CboReportType.AddItem "XL_DemandDetail" '0004
''CboReportType.AddItem "XL_DispatchStatus" '0010
''CboReportType.AddItem "XL_ReelBaseQty"
''CboReportType.AddItem "AllDispatchByGroupID"
''CboReportType.AddItem "MEBom_Model"
''CboReportType.AddItem "ForbiddenPN"
''CboReportType.AddItem "SplicePN"              '（1077）
''CboReportType.AddItem "SpliceReplacePN"       '（1077）
''CboReportType.AddItem "MaintainFeeder"        '（1082）
''CboReportType.AddItem "PDA_DistributeDIDLog"        '（1089）
''CboReportType.AddItem "ME_BOM_GroupID"        '（1127）
Call GetLine
If StrBU = "NB4" Or StrBU = "NB5" Then    ''(1179)
    chkDual.Visible = True
End If

End Sub
Private Function GetLine()
Dim str As String
Dim RS As ADODB.Recordset
str = "select distinct Line from SAP_WO_List order by line"
Set RS = Conn.Execute(str)
CboLine.Clear
While Not RS.EOF
    CboLine.AddItem RS!Line
    RS.MoveNext
Wend
End Function

Private Function GetGroupID(ByVal Jobpn As String)
Dim str As String
Dim BeginDate, EndDate As String
Dim GroupIDHead As String
Dim I As Long
Dim RS As ADODB.Recordset
Dim TempJobPn As String
BeginDate = Format(dtpSDate, "YYYY/MM/DD")
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
EndDate = Format(dtpEDate, "YYYY/MM/DD")
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")

If Jobpn = "" Then
    If BU = "NB5" Then
        If OptRelease.Value = True Then
           str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )"
        Else
           str = "select distinct GroupID from QSMS_WOGroup  where substring(GroupID,4,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )"
        End If
    Else
        If OptRelease.Value = True Then
           str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N'"
        Else
            ''''1041
            If StrBU = "PO" Then
                str = "select distinct GroupID from QSMS_WOGroup  where substring(GroupID,2,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N' "
            Else
                str = "select distinct GroupID from QSMS_WOGroup  where substring(GroupID,4,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N' "
            End If
        End If
    End If
Else
    If InStr(1, Jobpn, "-") > 0 Then
       
       TempJobPn = Mid(Jobpn, 1, 11)
       TxtMBPN = TempJobPn
       txtRev.Text = Right(Jobpn, 3)
    Else
      TempJobPn = Jobpn
    End If
    If OptRelease.Value = True Then
       str = "select distinct GroupID from QSMS_WOGroup a,QSMS_JobBOM b   where a.WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' " & _
             "and a.line='" & CboLine & "' and a.work_order=b.work_order and b.jobpn='" & TempJobPn & "'and closedflag='N' "
    Else
        str = "select distinct GroupID from QSMS_WOGroup a,QSMS_JobBOM b  where a.WO_TransDateTime between '" & BeginDate & "' and '" & EndDate & "' " & _
              "and a.line='" & CboLine & "' and a.work_order=b.work_order and b.jobpn='" & TempJobPn & "' and closedflag='N'"
    End If
End If

Set RS = Conn.Execute(str)
CboGroupID.Clear
If RS.EOF Then MsgBox "No data"
While Not RS.EOF
      CboGroupID.AddItem Trim(RS!GroupID)
      RS.MoveNext
Wend
End Function
Private Function GetJobPN()
Dim str As String
Dim BeginDate, EndDate As String
Dim GroupIDHead As String
Dim I As Long
Dim RS As ADODB.Recordset
BeginDate = Format(dtpSDate, "YYYY/MM/DD")
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
EndDate = Format(dtpEDate, "YYYY/MM/DD")
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")

If OptRelease.Value = True Then
   str = "select distinct JobPN from QSMS_WOGroup a,QSMS_JobBOM b where a.WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' " & _
         "and a.line='" & CboLine & "' and A.work_order=b.work_order"
Else
    ''''1041
    If StrBU = "PO" Then
        str = "select distinct JobPN from QSMS_WOGroup a,QSMS_JobBOM b where substring(GroupID,2,8) between '" & BeginDate & "' and '" & EndDate & "' " & _
              "and a.line='" & CboLine & "' and a.work_Order=b.work_Order"
    Else
        str = "select distinct JobPN from QSMS_WOGroup a,QSMS_JobBOM b where substring(GroupID,4,8) between '" & BeginDate & "' and '" & EndDate & "' " & _
              "and a.line='" & CboLine & "' and a.work_Order=b.work_Order"
    End If
End If

Set RS = Conn.Execute(str)
cboJobPN.Clear
If RS.EOF Then MsgBox "No data"
While Not RS.EOF
      cboJobPN.AddItem Trim(RS!Jobpn)
      RS.MoveNext
Wend
End Function

Private Function GetGroupWO(ByVal GroupID As String)
Dim str As String
Dim TransDate As String
Dim RS As ADODB.Recordset
Dim TempJobPn As String
TempJobPn = ""
If Trim(cboJobPN) <> "" Then
   TempJobPn = Mid(cboJobPN, 1, 11)
End If
str = "select distinct a.Work_Order from QSMS_WOGroup a,QSMS_JobBOM b  where a.GroupID= '" & GroupID & "' and a.work_Order=b.work_order and b.jobpn like '" & TempJobPn & "%'"

Set RS = Conn.Execute(str)
ListWoall.Clear
cboWO.Clear
CboNotChkBOM.Clear
While Not RS.EOF
          If ChkQSMS_WO(Trim(RS!Work_Order)) = False Then
             CboNotChkBOM.AddItem Trim(RS!Work_Order)
          Else
             ListWoall.AddItem Trim(RS!Work_Order)
             cboWO.AddItem Trim(RS!Work_Order)
          End If
          RS.MoveNext
Wend
End Function


Private Function GetWoinfo(ByVal WO As String)
Dim str As String
Dim RS As ADODB.Recordset
str = "select PN, Qty ,MB_Rev,Line,BuildType from Sap_Wo_List where WO='" & Trim(WO) & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   TxtMBPN = RS!PN
   TxtWOQty = RS!Qty
   txtRev = Trim(RS!Mb_Rev)
   CboLine.Text = Trim(RS!Line)
End If
''''Lynn modify 2007/07/31
'Str = "select Customer from ModelName where PN='" & TxtMBPN & "'"
str = "select Customer from ModelName where modelname='" & TxtMBPN & "-" & txtRev & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   TxtCustomer = Trim(RS!Customer)
End If
cboJobPN.Clear
str = "select jobPn from QSMS_JobBOM where work_Order='" & WO & "'"
Set RS = Conn.Execute(str)
While Not RS.EOF
     cboJobPN.AddItem Trim(RS!Jobpn)
     RS.MoveNext
Wend
End Function


Private Function GetMachine(ByVal WO As String)
Dim str As String
Dim TransDate As String
Dim RS As ADODB.Recordset

str = "select distinct Machine,MachinefinishedFlag from QSMS_WO where Work_Order= '" & WO & "' "

Set RS = Conn.Execute(str)
CboMachine.Clear
CboMachine.AddItem "ALL"
While Not RS.EOF
    
        CboMachine.AddItem Trim(RS!Machine)
    
     RS.MoveNext
Wend

End Function

Private Function GetComp(ByVal Machine As String, ByVal MBPN As String, ByVal WO As String, ByVal Line As String)


Dim str As String
Dim RS As ADODB.Recordset

If Machine <> "ALL" Then
    str = "select CompPN from QSMS_WO  where Work_Order='" & Trim(WO) & "'  and Machine='" & Trim(Machine) & "'"
Else
    str = "select CompPN from QSMS_WO  where Work_Order='" & Trim(WO) & "'"
End If
Set RS = Conn.Execute(str)
CboComp.Clear
While Not RS.EOF
  
       CboComp.AddItem Trim(RS!compPN)
       

    RS.MoveNext
Wend

End Function
'Private Function JobGroup(ByVal Machine As String, ByVal MBPN As String, ByVal Rev As String)
'Dim Str As String
'Dim Rs As ADODB.Recordset
'Str = "select distinct JobGroup from QSMS_MEBom where Machine='" & Trim(Machine) & "' and JobPN='" & Trim(MBPN) & "' and Version='" & Trim(Rev) & "'"
'Set Rs = Conn.Execute(Str)
'CmbJobGroup.Clear
'While Not Rs.EOF
'     CmbJobGroup.AddItem Trim(Rs!JobGroup)
'     Rs.MoveNext
'Wend
'End Function

Private Function CopyToExcelPrepareMaterialList(ByVal SheetName As String, ByVal WO As String, ByVal Machine As String, ByVal Line As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim strDate As String
Dim strFileName  As String, strSheetName As String
Dim blnActiveSheet As Boolean
Dim MyXlsApp As Excel.Application
Dim xlWorkSheet As Excel.Worksheet
Dim strFlag As String
Dim LocalPath As String
LocalPath = "D:\QSMS_Report\"
If Dir(LocalPath, vbDirectory) = "" Then MkDir LocalPath

LblMessage = ""
str = "exec QSMSRptPrepareMaterial '" & WO & "','" & Machine & "'"
Set RS = Conn.Execute(str)
If RS.EOF Then
   MsgBox "No data"
   Exit Function
End If


Set MyXlsApp = CreateObject("Excel.Application")
MyXlsApp.Visible = False
MyXlsApp.UserControl = True
MyXlsApp.DisplayAlerts = False


strFileName = LocalPath & "SMTMaterialReport_" & WO & ".xls"
If Dir(strFileName) <> "" Then
Kill strFileName
End If
FileCopy App.Path & "\SMTMaterialReport.xls", strFileName
'Copy Sample sheet to new file
blnActiveSheet = True
MyXlsApp.Workbooks.Open FileName:=strFileName
MyXlsApp.Visible = False
MyXlsApp.UserControl = True


MyXlsApp.ActiveWorkbook.Sheets(SheetName).Activate

    
MyXlsApp.Sheets(SheetName).Cells(2, 3).Value = Line
MyXlsApp.Sheets(SheetName).Cells(2, 5).Value = WO
MyXlsApp.Sheets(SheetName).Cells(2, 7).Value = Machine

MyXlsApp.Sheets(SheetName).Cells(4, 1).CopyFromRecordset RS
MyXlsApp.Visible = True
'MyXlsApp.ActiveWorkbook.SaveAs strFileName


'MyXlsApp.Quit
'Set MyXlsApp = Nothing
LblMessage.Caption = "report OK" + strAddress



End Function

Private Function ToExcelVerifyReport()
 Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object
 Dim fldCount As Integer, iCol As Integer
 Dim str As String
 Dim RS As ADODB.Recordset
 Dim RsLine As ADODB.Recordset
 Dim strLine As String
 Dim strFileName, Trans_Date As String
 Dim I As Integer
 Dim ReportFileName As String
    str = "EXEC GenVerifyReportByBU "
    Conn.Execute (str)
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add
    xlApp.DisplayAlerts = False
    xlApp.UserControl = True
    
    
str = "select distinct LEFT(machine, 1) AS Expr1 from machine WHERE machine<>'' order by LEFT(machine, 1) "
Set RsLine = Conn.Execute(str)
If Not RsLine.EOF Then
    RsLine.MoveFirst
    While Not RsLine.EOF
    I = I + 1
    strLine = Trim(RsLine("Expr1"))
    
    xlApp.Sheets(I).Select
    xlApp.Sheets.Add
    xlApp.Sheets(I).Select
    xlApp.Sheets(I).Move After:=xlApp.Sheets(I)
    
    Set xlWs = xlApp.Worksheets(I)
    
    xlApp.Worksheets(I).Name = strLine
    str = "Exec GenVerifyReportToExcel '" & strLine & "'"
    Set RS = Conn.Execute(str)
    fldCount = RS.Fields.Count
  
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = RS.Fields(iCol - 1).Name
    Next
        xlWs.Cells(2, 1).CopyFromRecordset RS
    RsLine.MoveNext
    Wend
End If

    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlApp.Visible = True
    RS.Close
    Set RS = Nothing
    Trans_Date = Format(Now, "YYYYMMDD")
    xlApp.Sheets("Sheet1").Select
    xlApp.ActiveWindow.SelectedSheets.Delete
    xlApp.Sheets("Sheet2").Select
    xlApp.ActiveWindow.SelectedSheets.Delete
    xlApp.Sheets("Sheet3").Select
    xlApp.ActiveWindow.SelectedSheets.Delete
    xlApp.Worksheets(1).Select
'    xlApp.ActiveWorkbook.SaveAs ReportFileName
    Set xlApp = Nothing
    Set xlsBook = Nothing
End Function
Private Function ToExcelWOChged()
 Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object
 Dim fldCount As Integer, iCol As Integer
 Dim str As String
 Dim RS As ADODB.Recordset
 Dim RsLine As ADODB.Recordset
 Dim strLine As String
 Dim strFileName, Trans_Date As String
 Dim I As Integer
 Dim ReportFileName As String

    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add
    xlApp.DisplayAlerts = False
    xlApp.UserControl = True
    
    
str = "select distinct LEFT(machine, 1) AS Expr1 from machine WHERE machine<>'' order by LEFT(machine, 1) "
Set RsLine = Conn.Execute(str)
If Not RsLine.EOF Then
    RsLine.MoveFirst
    While Not RsLine.EOF
    I = I + 1
    strLine = Trim(RsLine("Expr1"))
    
    xlApp.Sheets(I).Select
    xlApp.Sheets.Add
    xlApp.Sheets(I).Select
    xlApp.Sheets(I).Move After:=xlApp.Sheets(I)
    
    Set xlWs = xlApp.Worksheets(I)
    
    xlApp.Worksheets(I).Name = strLine
    str = "EXEC SP8_VerificationReportWOChged '" & strLine & "' "
    Set RS = Conn.Execute(str)
    fldCount = RS.Fields.Count
  
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = RS.Fields(iCol - 1).Name
    Next
        xlWs.Cells(2, 1).CopyFromRecordset RS
    RsLine.MoveNext
    Wend
End If

    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlApp.Visible = True
    RS.Close
    Set RS = Nothing
    Trans_Date = Format(Now, "YYYYMMDD")
    xlApp.Sheets("Sheet1").Select
    xlApp.ActiveWindow.SelectedSheets.Delete
    xlApp.Sheets("Sheet2").Select
    xlApp.ActiveWindow.SelectedSheets.Delete
    xlApp.Sheets("Sheet3").Select
    xlApp.ActiveWindow.SelectedSheets.Delete
    xlApp.Worksheets(1).Select
    Set xlApp = Nothing
    Set xlsBook = Nothing
End Function

Private Function ToExcelVerifyJobFailLog()
 Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object
 Dim fldCount As Integer, iCol As Integer
 Dim str As String
 Dim RS As New ADODB.Recordset
 Dim RsLine As ADODB.Recordset
 Dim strLine As String
 Dim strFileName, Trans_Date As String
 Dim I As Integer
 Dim ReportFileName As String
 
 If DateDiff("d", dtpSDate, dtpEDate) > 7 Then
    MsgBox ("The date range must be <= 7 days!")
    Exit Function
 End If
   
        
str = "Exec GenVerifyFailReport" & sq(Format(dtpSDate, "YYYYMMDD")) & "," & sq(Format(dtpEDate, "YYYYMMDD"))
RS.Open str, Conn, adOpenForwardOnly, adLockReadOnly
If Not RS.EOF Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found"
End If


End Function

Private Function CopyToExcelWipByMaterial(ByVal SheetName As String, ByVal compPN As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim strDate As String
Dim strFileName  As String, strSheetName As String
Dim blnActiveSheet As Boolean
Dim MyXlsApp As Excel.Application
Dim xlWorkSheet As Excel.Worksheet
Dim strFlag As String
Dim LocalPath As String
LocalPath = "D:\QSMS_Report\"
If Dir(LocalPath, vbDirectory) = "" Then MkDir LocalPath
LblMessage = ""
str = "exec QSMSRptWipByMaterial '" & compPN & "'"
Set RS = Conn.Execute(str)
If RS.EOF Then
   MsgBox "No data"
   Exit Function
End If


Set MyXlsApp = CreateObject("Excel.Application")
MyXlsApp.Visible = False
MyXlsApp.UserControl = True
MyXlsApp.DisplayAlerts = False

'strFileName = App.Path & "\SMTMaterialReport_" & CompPN & ".xls"
strFileName = LocalPath & "SMTMaterialReport_" & compPN & ".xls"
If Dir(strFileName) <> "" Then
Kill strFileName
End If
FileCopy App.Path & "\SMTMaterialReport.xls", strFileName
'Copy Sample sheet to new file
blnActiveSheet = True
MyXlsApp.Workbooks.Open FileName:=strFileName
MyXlsApp.Visible = False
MyXlsApp.UserControl = True


MyXlsApp.ActiveWorkbook.Sheets(SheetName).Activate
MyXlsApp.Sheets(SheetName).Cells(2, 4).Value = compPN
MyXlsApp.Sheets(SheetName).Cells(2, 6).Value = RS.Fields(0)

Set RS = RS.NextRecordset
MyXlsApp.Sheets(SheetName).Cells(4, 1).CopyFromRecordset RS
MyXlsApp.Visible = True
'MyXlsApp.ActiveWorkbook.SaveAs strFileName
'
'
'MyXlsApp.Quit
'Set MyXlsApp = Nothing
LblMessage.Caption = "report OK" + strAddress



End Function


Private Function CopyToExcelWipByDate(ByVal SheetName As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim strDate As String
Dim strFileName  As String, strSheetName As String
Dim blnActiveSheet As Boolean
Dim MyXlsApp As Excel.Application
Dim xlWorkSheet As Excel.Worksheet
Dim strFlag As String
Dim TransDate As String
Dim LocalPath As String
LocalPath = "D:\QSMS_Report\"
If Dir(LocalPath, vbDirectory) = "" Then MkDir LocalPath
LblMessage = ""
str = "select Getdate()"
Set RS = Conn.Execute(str)
TransDate = Format(RS.Fields(0), "YYYYMMDD")

str = "exec QSMSRptWipBydate "
Set RS = Conn.Execute(str)
If RS.EOF Then
   MsgBox "No data"
   Exit Function
End If


Set MyXlsApp = CreateObject("Excel.Application")
MyXlsApp.Visible = False
MyXlsApp.UserControl = True
MyXlsApp.DisplayAlerts = False

strFileName = LocalPath & "SMTMaterialReport_" & TransDate & ".xls"
If Dir(strFileName) <> "" Then
Kill strFileName
End If
FileCopy App.Path & "\SMTMaterialReport.xls", strFileName
'Copy Sample sheet to new file
blnActiveSheet = True
MyXlsApp.Workbooks.Open FileName:=strFileName
MyXlsApp.Visible = False
MyXlsApp.UserControl = True


MyXlsApp.ActiveWorkbook.Sheets(SheetName).Activate
MyXlsApp.Sheets(SheetName).Cells(2, 2).Value = TransDate
MyXlsApp.Sheets(SheetName).Cells(2, 5).Value = RS.Fields(0)

Set RS = RS.NextRecordset
MyXlsApp.Sheets(SheetName).Cells(4, 1).CopyFromRecordset RS
'
'MyXlsApp.ActiveWorkbook.SaveAs strFileName
'
'
'MyXlsApp.Quit
'Set MyXlsApp = Nothing
LblMessage.Caption = "report OK" + strAddress



End Function


Private Function CopyToExcelWipByGroup(ByVal SheetName As String, ByVal GroupID As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim strDate As String
Dim strFileName  As String, strSheetName As String
Dim blnActiveSheet As Boolean
Dim MyXlsApp As Excel.Application
Dim xlWorkSheet As Excel.Worksheet
Dim strFlag As String
Dim TransDate As String
Dim LocalPath As String
LocalPath = "D:\QSMS_Report\"
If Dir(LocalPath, vbDirectory) = "" Then MkDir LocalPath
LblMessage = ""
str = "select Getdate()"
Set RS = Conn.Execute(str)
TransDate = Format(RS.Fields(0), "YYYYMMDD")

str = "exec QSMSRptWipByGroup '" & GroupID & "'"
Set RS = Conn.Execute(str)
If RS.EOF Then
   MsgBox "No data"
   Exit Function
End If


Set MyXlsApp = CreateObject("Excel.Application")
MyXlsApp.Visible = False
MyXlsApp.UserControl = True
MyXlsApp.DisplayAlerts = False

strFileName = LocalPath & "\SMTMaterialReport_" & GroupID & ".xls"
If Dir(strFileName) <> "" Then
Kill strFileName
End If
FileCopy App.Path & "\SMTMaterialReport.xls", strFileName
'Copy Sample sheet to new file
blnActiveSheet = True
MyXlsApp.Workbooks.Open FileName:=strFileName
MyXlsApp.Visible = False
MyXlsApp.UserControl = True


MyXlsApp.ActiveWorkbook.Sheets(SheetName).Activate
MyXlsApp.Sheets(SheetName).Cells(2, 2).Value = GroupID
MyXlsApp.Sheets(SheetName).Cells(2, 6).Value = RS.Fields(0)

Set RS = RS.NextRecordset
MyXlsApp.Sheets(SheetName).Cells(4, 1).CopyFromRecordset RS

'MyXlsApp.ActiveWorkbook.SaveAs strFileName
'
'
'MyXlsApp.Quit
'Set MyXlsApp = Nothing
LblMessage.Caption = "report OK" + strAddress



End Function

Private Function CopyToExcelWipLackByWo(ByVal SheetName As String, ByVal Work_Order As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim strDate As String
Dim strFileName  As String, strSheetName As String
Dim blnActiveSheet As Boolean
Dim MyXlsApp As Excel.Application
Dim xlWorkSheet As Excel.Worksheet
Dim strFlag As String
Dim TransDate As String
Dim LocalPath As String
LocalPath = "D:\QSMS_Report\"
If Dir(LocalPath, vbDirectory) = "" Then MkDir LocalPath
LblMessage = ""
str = "select Getdate()"
Set RS = Conn.Execute(str)
TransDate = Format(RS.Fields(0), "YYYYMMDD")

str = "exec QSMSRptLackCompByWo '" & Work_Order & "'"
Set RS = Conn.Execute(str)
If RS.EOF Then
   MsgBox "No data"
   Exit Function
End If


Set MyXlsApp = CreateObject("Excel.Application")
MyXlsApp.Visible = False
MyXlsApp.UserControl = True
MyXlsApp.DisplayAlerts = False

strFileName = LocalPath & "SMTMaterialReport_" & Work_Order & ".xls"
If Dir(strFileName) <> "" Then
Kill strFileName
End If
FileCopy App.Path & "\SMTMaterialReport.xls", strFileName
'Copy Sample sheet to new file
blnActiveSheet = True
MyXlsApp.Workbooks.Open FileName:=strFileName
MyXlsApp.Visible = False
MyXlsApp.UserControl = True


MyXlsApp.ActiveWorkbook.Sheets(SheetName).Activate

MyXlsApp.Sheets(SheetName).Cells(2, 5).Value = RS.Fields(0)

Set RS = RS.NextRecordset
MyXlsApp.Sheets(SheetName).Cells(4, 1).CopyFromRecordset RS

'MyXlsApp.ActiveWorkbook.SaveAs strFileName
'
'
'MyXlsApp.Quit
'Set MyXlsApp = Nothing
LblMessage.Caption = "report OK" + strAddress



End Function


Private Function CopyToExcelWipDifferentMaterial(ByVal SheetName As String, ByVal Work_Order As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim strDate As String
Dim strFileName  As String, strSheetName As String
Dim blnActiveSheet As Boolean
Dim MyXlsApp As Excel.Application
Dim xlWorkSheet As Excel.Worksheet
Dim strFlag As String
Dim TransDate As String, LocalPath As String

LblMessage = ""
str = "select Getdate()"
Set RS = Conn.Execute(str)
TransDate = Format(RS.Fields(0), "YYYYMMDD")

str = "exec QSMSRptWipDifferentMaterial '" & Work_Order & "'"
Set RS = Conn.Execute(str)
If RS.EOF Then
   MsgBox "No data"
   Exit Function
End If


Set MyXlsApp = CreateObject("Excel.Application")
MyXlsApp.Visible = False
MyXlsApp.UserControl = True
MyXlsApp.DisplayAlerts = False

strFileName = LocalPath & "SMTMaterialReport_" & Work_Order & ".xls"
If Dir(strFileName) <> "" Then
Kill strFileName
End If
FileCopy App.Path & "\SMTMaterialReport.xls", strFileName
'Copy Sample sheet to new file
blnActiveSheet = True
MyXlsApp.Workbooks.Open FileName:=strFileName
MyXlsApp.Visible = False
MyXlsApp.UserControl = True


MyXlsApp.ActiveWorkbook.Sheets(SheetName).Activate

MyXlsApp.Sheets(SheetName).Cells(2, 2).Value = TransDate


MyXlsApp.Sheets(SheetName).Cells(4, 1).CopyFromRecordset RS

'MyXlsApp.ActiveWorkbook.SaveAs strFileName
'
'
'MyXlsApp.Quit
'Set MyXlsApp = Nothing
LblMessage.Caption = "report OK" + strAddress



End Function


Private Function GetSapBom(ByVal Work_Order As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim strWO As String
'Add by Kevin 2008.12.08 (0033)
'Call GetWO from ListWoSelecting get WO list
If txtFilePath <> "" And Me.ListWoSelecting.ListCount > 0 Then
    strWO = GetWO("BY_WORKORDERS", "N")
    str = "exec QSMS_QuerySapBom '" & strWO & "'"
Else
    If Trim(txtWO) = "" Then
        MsgBox "Please check the WO"
        Exit Function
    Else
        str = "select * from sap_bom where work_Order='" & Trim(Work_Order) & "' order by CompPN,Item,CompLevel"
    End If
End If
Set RS = Conn.Execute(str)
If Not RS.EOF Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found <GetSapBom> !", vbInformation
End If
Me.txtFilePath = ""
Me.ListWoSelecting.Clear
End Function

Sub MachineType()
Dim str As String
Dim RS As New ADODB.Recordset

RS.CursorLocation = adUseClient
str = "  select Vendor, Line,SeqIDByLine,Side , Factory,Machine,Unit,Qty, MaxSlotNum, LR, MappingID,FujiData,DIOCircuit from Machine"
RS.Open str, Conn, adOpenForwardOnly, adLockReadOnly
If Not RS.EOF Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found"
End If

RS.Close
Set RS = Nothing
End Sub


Sub TraySlot()
Dim str As String
Dim RS As New ADODB.Recordset

RS.CursorLocation = adUseClient
str = "select * from TraySlot"
RS.Open str, Conn, adOpenForwardOnly, adLockReadOnly
If Not RS.EOF Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found"
End If

RS.Close
Set RS = Nothing
End Sub

Sub CastRate()
Dim str As String
Dim RS As New ADODB.Recordset

RS.CursorLocation = adUseClient
str = "select * from QSMS_CastRate"
RS.Open str, Conn, adOpenForwardOnly, adLockReadOnly
If Not RS.EOF Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found"
End If

RS.Close
Set RS = Nothing
End Sub
Sub AOIQtySummary(ByVal Work_Order As String) '1059
Dim str As String
Dim RS As New ADODB.Recordset

RS.CursorLocation = adUseClient
str = "exec QSMSGetAOISummary '" & Work_Order & "'"
RS.Open str, Conn, adOpenForwardOnly, adLockReadOnly
If Not RS.EOF Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found"
End If

RS.Close
Set RS = Nothing
End Sub
Sub AOIDetail(ByVal Work_Order As String)
Dim str As String
Dim RS As New ADODB.Recordset

RS.CursorLocation = adUseClient
str = "select * from QSMS_AOI where wo= '" & Work_Order & "' order by station,transdatetime"
RS.Open str, Conn, adOpenForwardOnly, adLockReadOnly
If Not RS.EOF Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found"
End If

RS.Close
Set RS = Nothing
End Sub
Sub SameGroupWO(ByVal Work_Order As String)
Dim str As String
Dim RS As New ADODB.Recordset

RS.CursorLocation = adUseClient
str = "select WO,PN,MB_Rev,Line,QTY,CombineQty,Trans_Date,WO_Type from sap_wo_list where [group] in (select [group] from sap_wo_list where wo= '" & Work_Order & "')"
RS.Open str, Conn, adOpenForwardOnly, adLockReadOnly
If Not RS.EOF Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found"
End If

RS.Close
Set RS = Nothing
End Sub
Sub OneByOne()
Dim str As String
Dim RS As New ADODB.Recordset

RS.CursorLocation = adUseClient
str = "select * from QSMS_OneByOne"
RS.Open str, Conn, adOpenForwardOnly, adLockReadOnly
If Not RS.EOF Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found"
End If

RS.Close
Set RS = Nothing
End Sub
Sub LineChangeStatistics()
Dim str As String
Dim RS As New ADODB.Recordset
Dim sheet1 As String
Dim IntRow As Integer, fldCount As Integer
Dim j, k As Integer, iCol As Integer, jCol As Integer, ColCount As Integer
Dim TempStr As String
Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object

RS.CursorLocation = adUseClient
str = "Exec GenChangeLineReport1" & sq(Format(dtpSDate, "YYYYMMDD")) & "," & sq(Format(dtpEDate, "YYYYMMDD"))
RS.Open str, Conn, adOpenForwardOnly, adLockReadOnly
If Not RS.EOF Then
    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add
    xlApp.DisplayAlerts = False
    xlApp.UserControl = True
    Set xlWs = xlApp.Worksheets(1)
    k = 2
    fldCount = RS.Fields.Count
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = RS.Fields(iCol - 1).Name
    Next
'        xlWs.Cells(2, 1).CopyFromRecordset Rs
    ColCount = RS.RecordCount
    For j = 1 To ColCount
        For jCol = 1 To fldCount
            xlWs.Cells(k, jCol).Value = RS.Fields(jCol - 1).Value
        Next
        If xlWs.Cells(k, 12).Value > 15 Then
            xlWs.Cells(k, 12).Select
            With Selection.Interior
                .ColorIndex = 3
                .Pattern = xlSolid
            End With
        End If
        'REMARK
'        xlWs.Cells(K, 9).AddComment
'        xlWs.Cells(K, 9).Comment.Visible = False
'        If IsNull(Rs.Fields(jCol - 1).Value) Then
'            tempStr = "NULL"
'        Else
'            tempStr = Rs.Fields(jCol - 1).Value
'        End If
'
'        xlWs.Cells(K, 9).Comment.Text Text:="Note:" & Chr(10) & tempStr
        k = k + 1
        RS.MoveNext
    Next
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlApp.Visible = True
    sheet1 = "Sheet1"
    IntRow = RS.RecordCount + 1
    xlApp.Sheets(sheet1).Range(xlApp.Sheets(sheet1).Cells(1, 1), xlApp.Sheets(sheet1).Cells(IntRow, 13)).Select
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Range("A1:M1").Select
    Selection.Font.ColorIndex = 3
    Selection.Font.Bold = True
    'Set backColor
    Range("A1:M1").Select
    With Selection.Interior
        .ColorIndex = 35
        .Pattern = xlSolid
    End With
'    xlApp.Sheets(sheet1).Range(xlApp.Sheets(sheet1).Cells(2, 1), xlApp.Sheets(sheet1).Cells(IntRow, 13)).Select
'    With Selection.Interior
'        .ColorIndex = 36
'        .Pattern = xlSolid
'    End With
    
    xlApp.Sheets("Sheet1").Select
    Set xlApp = Nothing
    Set xlsBook = Nothing
Else
    MsgBox "No data found"
End If

RS.Close
Set RS = Nothing
End Sub
Sub LineChangeStatisticsByall()
Dim str As String
Dim RS As New ADODB.Recordset

RS.CursorLocation = adUseClient
str = "Exec GenChangeLineReport2" & sq(Format(dtpSDate, "YYYYMMDD")) & "," & sq(Format(dtpEDate, "YYYYMMDD"))
RS.Open str, Conn, adOpenForwardOnly, adLockReadOnly
If Not RS.EOF Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found"
End If

RS.Close
Set RS = Nothing
End Sub
Sub DIDDeleteRecords()
Dim str As String
Dim RS As New ADODB.Recordset

RS.CursorLocation = adUseClient
str = "Exec QueryDIDUsed" & sq(Format(dtpSDate, "YYYYMMDD")) & "," & sq(Format(dtpEDate, "YYYYMMDD"))
RS.Open str, Conn, adOpenForwardOnly, adLockReadOnly
If Not RS.EOF Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found"
End If

RS.Close
Set RS = Nothing
End Sub
Private Function GetSapGroupByWo(ByVal Work_Order As String)
Dim str As String
Dim RS As ADODB.Recordset

If Trim(txtWO) = "" Then
   MsgBox "Please check the WO"
   Exit Function
End If
str = "select * from sap_wo_list  where [group] in (select [group] from sap_wo_list where wo='" & Trim(Work_Order) & "') order by wo"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found"
End If
End Function

Private Function GetMEBom(ByVal Work_Order As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim Line As String
Dim Rst As ADODB.Recordset ''''1073
Dim TempGroup As String
Dim TempJobPn As String, TempJObGroup As String
Dim Machine As String
Dim MultiLineStr As Variant
Dim I As Integer

If Trim(txtWO) <> "" Then

      str = "Select [Group],Line,BuildType from sap_wo_list where wo='" & Work_Order & "'"
      Set RS = Conn.Execute(str)
      If Not RS.EOF Then
         TempGroup = Trim(RS![Group])
         TempJObGroup = GetJobGroup(TempGroup)
         '2007-03-30 Denver for BuildType=4
         If Trim(RS("BuildType")) = 4 Then
            Machine = GetMultiLine(Work_Order, "WO")
         Else
            Machine = Trim(RS!Line) & "%"
         End If
         '2007-03-30 End
      End If

Else
    'If Trim(CboLine) = "" Or Trim(TxtMBPN) = "" Or Trim(TxtRev) = "" Then
 
    TempJObGroup = GetSelectingJobGroup()
        
    If CboMachine = "ALL" Then
        Machine = Trim(CboLine) & "%"
    Else
        If Len(Trim(CboMachine)) > 0 Then
           CboLine = Mid(Trim(CboMachine), 1, 1)
        End If
       
        Machine = CboMachine & "%"
    End If
    If Trim(CboLine) = "" Then
        MsgBox "Please check Line "
        Exit Function
    End If
End If

If TempJObGroup = "" Then
    If CboMachine = "" Then
        MsgBox "请选择 Machine"
        Exit Function
    End If
    
    MsgBox "请选择 JobGroup or work order"
    Exit Function
End If

'2007-03-30 Denver for BuildType=4
'second char of Machine must be "S" for BuildType=4
If InStr(Machine, ",") > 0 Then
    MultiLineStr = Split(Machine, ",")
    Machine = ""
    For I = 0 To UBound(MultiLineStr)
        If Right(MultiLineStr(I), 1) = "Q" Then
            Machine = Machine & "((a.machine like '" & Left(MultiLineStr(I), 1) & "Q%' OR a.machine like '" & Left(MultiLineStr(I), 1) & "W%') and (a.Side like 'Q%' or a.Side like 'W%' )) OR "
        Else
            Machine = Machine & "(a.machine like '" & Left(MultiLineStr(I), 1) & "S%' and a.Side like '" & Right(MultiLineStr(I), 1) & "%') OR "
        End If
    Next I
    Machine = Left(Machine, Len(Machine) - 3)
    Machine = "(" & Machine & ")"
Else
    Machine = "a.machine like '" & Machine & "'"
End If
''(1031)
''If TxtMBPN = "" Then
If txtWO <> "" Then
        '''''''''''''''1073
    'str = "select * from WO_MultiLine where WO='" & Trim(TxtWO) & "'"
    ''1164
    str = "select a.Line from WO_MultiLine a, Sap_Wo_List b where a.WO=b.WO and B.[Group] in(select [Group] from Sap_Wo_List where BuildType='4' and WO='" & Trim(txtWO) & "')"
    Set Rst = Conn.Execute(str)
    If Not Rst.EOF Then
        Line = Rst("Line")
        str = "select a.Machine,a.JobGroup,a.JObpN,a.Version,A.CompPN,A.lR,''''+ a.Slot as Slot,a.Qty,a.BuildType,a.Side,a.Factory,a.UID,a.TransDateTime,a.Location from QSMS_MeBom a where  a.version like '" & Trim(txtRev) & "%' " & _
            " and (a.line='" & Trim(RS!Line) & "'or a.line='" & Line & "') and a.jobpn in (select distinct jobpn from qsms_Jobbom where Work_Order='" & Trim(txtWO) & "') and jobgroup in " & TempJObGroup & " order by a.JobpN,a.machine,CompPN,Slot"      '''(1033)
    Else
        str = "select a.Machine,a.JobGroup,a.JObpN,a.Version,A.CompPN,A.lR,''''+ a.Slot as Slot,a.Qty,a.BuildType,a.Side,a.Factory,a.UID,a.TransDateTime,a.Location from QSMS_MeBom a where  a.version like '" & Trim(txtRev) & "%' " & _
            " and a.line='" & Trim(RS!Line) & "' and a.jobpn in (select distinct jobpn from qsms_Jobbom where Work_Order='" & Trim(txtWO) & "') and jobgroup in " & TempJObGroup & " order by a.JobpN,a.machine,CompPN,Slot"        '''(1033)
    End If
    ''''''''''''''1073
Else
    str = "select a.Machine,a.JobGroup,a.JObpN,a.Version,A.CompPN,A.lR,''''+ a.Slot as Slot,a.Qty,a.BuildType,a.Side,a.Factory,a.UID,a.TransDateTime,a.Location from QSMS_MeBom a where a.version='" & Trim(txtRev) & "' " & _
        " and a.line='" & Trim(CboLine) & "'  and jobgroup in " & TempJObGroup & " order by a.JobpN,a.machine,CompPN,Slot"
                
                
End If
'2007-03-30 end

Set RS = Conn.Execute(str)
If Not RS.EOF Then
   Call CopyToExcel(RS)
Else
   MsgBox "No data found,Please check"
End If
End Function
Private Function CheckWOWastagePN(ByVal Work_Order As String)  ''1212
Dim str As String
Dim RS As ADODB.Recordset
Dim TempGroup As String
Dim TempJobPn As String
Dim Machine As String

If Trim(txtWO) = "" Then
    MsgBox "Please Input WO !!!", vbOKOnly, "Fail"
    Exit Function
End If
str = "Select b.* from CompPN_Data a,QSMS_WO b with(nolock) where b.work_order='" & Trim(Work_Order) & "'and a.CompPN=b.CompPN  and a.type='WastagePN' order by jobpn,machine,slot,lr,comppn"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
    Call CopyToExcel(RS)
Else
    str = "Select b.* from CompPN_Data a,QSMS_History..QSMS_WO b with(nolock) where b.work_order='" & Trim(Work_Order) & "' and a.CompPN=b.CompPN and a.type='WastagePN' order by jobpn,machine,slot,lr,comppn"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
        Call CopyToExcel(RS)
    Else
        MsgBox "No Data in current and history !!", vbOKOnly, "Fail"
    End If
End If

End Function

Private Function GetMEBom_WO(ByVal Work_Order As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim TempGroup As String
Dim TempJobPn As String
Dim Machine As String

If Trim(txtWO) = "" Then
    MsgBox "Please Input WO !!!", vbOKOnly, "Fail"
    Exit Function
End If
If StrBU = "NB5" Then  ''1270
    str = "exec QSMSRpt_QSMS_WO '" & Work_Order & "','ME_BOM_WO'"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
        Call CopyToExcel(RS)
    Else
        MsgBox "No Data in current and history !!", vbOKOnly, "Fail"
    End If
Else
   str = "Select * from QSMS_WO with(nolock) where work_order='" & Trim(Work_Order) & "' order by jobpn,machine,slot,lr,comppn"
   Set RS = Conn.Execute(str)
   If Not RS.EOF Then
       Call CopyToExcel(RS)
   Else
       str = "Select * from QSMS_History..QSMS_WO with(nolock) where work_order='" & Trim(Work_Order) & "' order by jobpn,machine,slot,lr,comppn"
       Set RS = Conn.Execute(str)
       If Not RS.EOF Then
          Call CopyToExcel(RS)
       Else
          MsgBox "No Data in current and history !!", vbOKOnly, "Fail"
       End If
    End If
End If

End Function

Private Function GetMEBom_DeleteLog(ByVal DID As String)
    Dim str As String
    Dim RS As ADODB.Recordset
    Dim TempGroup As String
    Dim TempJobPn As String
    Dim Machine As String
    
    If Trim(TxtMBPN) = "" Then
        MsgBox "Please Input 'MB/Job PN' !!!", vbOKOnly, "Fail"
        Exit Function
    End If
    str = "Select * from QSMS_Log where DID like '%" & Trim(DID) & "%' AND Event_No='Delete Me_BOM' order by Trans_Date desc"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
       Call CopyToExcel(RS)
    Else
       MsgBox "No data found!!"
    End If
End Function
Private Function GetUnDispatchList(ByVal Work_Order As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim TempGroup As String
Dim TempJobPn As String
Dim Machine As String

If Trim(txtWO) = "" Then
    MsgBox "Please Input WO !!!", vbOKOnly, "Fail"
    Exit Function
End If

Set RS = Conn.Execute("QSMS_WOUnDispatch " & Work_Order)
Set RS = RS.NextRecordset

If Not RS.EOF Then
    Call CopyToExcel(RS)
Else
    Set RS = Conn.Execute("SELECT COUNT(*) FROM QSMS_WO WHERE Work_Order='" & Work_Order & "'")
    If RS(0) = 0 Then
        MsgBox "The WO hasn't begined to dispatch!", vbOKOnly
    Else
        MsgBox "The WO has complete to dispatch!", vbOKOnly
    End If
End If
End Function

Private Function GetCompPNDIDData(ByVal compPN As String)
Dim str As String
Dim RS As New ADODB.Recordset
    str = "GetCompPNDIDData '" & compPN & "'"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
       Call CopyToExcel(RS)
    Else
       MsgBox "No data found,please check!"
    End If
    
End Function
Private Function GetCompPNQty(ByVal compPN As String)
Dim str As String
Dim RS As New ADODB.Recordset
    str = "GetCompPNQty '" & compPN & "'"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
       Call CopyToExcel(RS)
    Else
       MsgBox "No data found,please check!"
    End If
End Function


Private Function CheckDispatchQty(ByVal WO As String)
Dim str As String
Dim RS As New ADODB.Recordset
   str = "select * from Sap_Wo_List where BuildType='1' and WO='" & Trim(WO) & "'"
   Set RS = Conn.Execute(str)
   If RS.EOF Then
      MsgBox "Can't find the WO!Please check!"
      Exit Function
   End If
      str = "QSMS_ReCountDispatchQty '" & WO & "','2'"
      Conn.Execute (str)
      MsgBox "That's OK!"
End Function



Private Function GetReplacePN(ByVal WO As String) '1059
Dim str As String
Dim iCol As Integer
Dim RS As ADODB.Recordset

If Trim(WO) = "" Then
    MsgBox "Please Input WO !!!", vbOKOnly, "Fail"
    Exit Function
End If
str = "exec QSMS_GetReplacePNByWOList '" & Trim(WO) & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   Call CopyToExcel(RS)
Else
    MsgBox "no data "
End If

End Function
Private Function Sap_Return(ByVal Report_Type)
Dim str As String
Dim RS As ADODB.Recordset
Dim sDateTime As String
Dim eDateTime As String
Select Case Report_Type
       Case "SAPCostSum"                                            ''(0025)
            str = "exec QSMS_GetSAPCostSum '" & Trim(cboWO) & "'"
       Case "SAP1His"
            str = "select A.Work_Order,A.Item,A.CompPN,A.UpCOmpPN,A.Qty,A.Status,A.TransDateTime,A.BeginDateTime,A.EndDateTime,A.AOIQty,A.SendFlag,A.SendID,B.Sap1FileName,B.FilePath,B.CheckFlag from qsms_saphis a,SAP1_Seq_Num b  where a.work_order='" & Trim(cboWO) & "' and a.SendID=b.SendID and a.SendID<>'' order by a.SendID"
       Case "SAP1"
            str = "select * from QSMS_Sap where work_order='" & Trim(cboWO) & "' order by UpCompPN,Item"
       Case "SAP2"
            str = "exec QSMS_GetSAP2 '" & Trim(cboWO) & "' "
       Case "WO_SingleCompPNData"
            str = "exec  QSMS_WOSingleCompPNData '" & Trim(cboWO) & "'"
       Case "GroupIDCostQty"                ''(0028)
            If Trim(CboGroupID) = "" Then
                MsgBox "Please choose GroupID first!", vbCritical
                Exit Function
            End If
            str = "Exec QSMS_GroupIDCostQty '" & Trim(CboGroupID) & "'"
       Case "ReturnDID"
            str = "Select * from QSMS_GroupDID where GroupID='" & Trim(CboGroupID) & "' order by returnFlag,CompPN"
            
       Case "ReturnDID_ByDate"          '''''（1237）
            sDateTime = Format(dtpSDate, "yyyymmdd") & "000000"
            eDateTime = Format(dtpEDate, "yyyymmdd") & "240000"

            str = "Select * from QSMS_GroupDID with(nolock) where TransDateTime between '" & Trim(sDateTime) & "' and '" & Trim(eDateTime) & "' and ReturnFlag='Y'" & _
                    "union All Select * from QSMS_History.dbo.QSMS_GroupDID with(nolock) where TransDateTime between '" & Trim(sDateTime) & "' and '" & Trim(eDateTime) & "' and ReturnFlag='Y'"
'*******************-(0001)
'*******************-(0005)
       Case "DispatchDID"       '--0017  '(0030)
'            str = "select a.*,CASE WHEN a.Inherit_WO='more' THEN N'超发料' else N'正常发料' END AS DispatchAdditional,case when a.transdatetime=b.lastupdatedt and b.deletedt='' and b.Qty=b.realQty then N'从未使用' " & _
'                  "when a.did in (select did from qsms_feederdid_current) and b.deletedt='' then N'使用中' " & _
'                  "when b.realqty<0 and b.deletedt='' then N'用完未删除' when b.deletedt<>'' and b.realqty<b.qty then N'用完删除' " & _
'                  "when b.realqty=a.didqty AND b.DELETEDT<>'' then N'未用删除' else N'当前未使用' end as status " & _
'                  "from qsms_dispatch a," & _
'                  "(select did,qty,remainqty,realqty,transdatetime,lastupdatedt,'' as deleteDT from qsms_did " & _
'                  "Union All " & _
'                  "select did,qty,remainqty,realqty,transdatetime,lastupdatedt,deletedt from qsms_did_log ) b " & _
'                  "where a.did=b.did and a.diddatetime=b.transdatetime "
'                  '"order by A.Work_order,A.DID"
            If dtpSDate = dtpEDate Then
                If CboShift = "" Then
                    MsgBox "Please choose shift first!", vbCritical
                    Exit Function
                End If
                sDateTime = Format(dtpSDate, "yyyymmdd") & IIf(CboShift = "Day_Shift", "0740", "1940")
                eDateTime = IIf(CboShift = "Day_Shift", Format(dtpEDate, "yyyymmdd") & "1940", Format(dtpEDate + 1, "yyyymmdd") & "0740")
                'str = str & " and a.transdatetime between '" & Format(DTPsdate, "yyyymmdd") & IIf(CboShift = "Day_Shift", "0740", "1940") & "' and " & _
                '      " '" & IIf(CboShift = "Day_Shift", Format(DTPedate, "yyyymmdd") & "1940", Format(DTPedate + 1, "yyyymmdd") & "0740") & "'"
            Else
                If DateDiff("d", dtpSDate, dtpEDate) > 1 And CboGroupID = "" Then
                    MsgBox "Date Range can not over 1 days when did not choose one wo group!", vbCritical
                    Exit Function
                End If
                sDateTime = Format(dtpSDate, "yyyymmdd") & "0740"
                eDateTime = Format(dtpEDate, "yyyymmdd") & "0740"
                'str = str & " and a.transdatetime between '" & Format(DTPsdate, "yyyymmdd") & "0740" & "' and '" & Format(DTPedate, "yyyymmdd") & "0740" & "'"
            End If
            str = "exec QSMS_DispatchDID @Sdatetime='" & sDateTime & "',@Edatetime='" & eDateTime & "',@GroupID='" & Trim(CboGroupID) & "',@CompPN='" & Trim(CboComp) & "'"    ''(1138)
'            If CboGroupID <> "" Then
'                str = str & " and a.work_order in(select Work_Order from QSMS_Wogroup where GroupID='" & Trim(CboGroupID) & "') "
'            End If
'            str = str & " order by A.Work_order,A.DID"
       Case "Return_Dispatch"
             str = "Select * from QSMS_GroupCompQty where GroupID='" & Trim(CboGroupID) & "' Order by CompPN"
       Case "CastQty"
             str = "exec QSMSGetCastQty '" & Trim(CboGroupID) & "'"
             Conn.Execute str
             str = "Select * from QSMS_GroupCompQty where GroupID='" & Trim(CboGroupID) & "' order by CompPN"
              str = "Select * from QSMS_GroupCompQty where GroupID='" & Trim(CboGroupID) & "'"
       Case "DIDCallBack"
'            If txtDID.Visible = True Then
'                str = "Select isnull(GroupID,'') AS GroupID,A.Work_Order,DID,CompPN,TotalQty,ReturnQty AS CallBackQty,ReturnFlag AS CallBackFlag,DeleteFlag,TransDateTime,A.UID " & _
'                "from QSMS_DIDCallBack A left join QSMS_WoGroup B on A.Work_Order=B.Work_Order where DID='" & Trim(txtDID) & "'"
'            Else
                str = "Select * from qsms_didcallback where work_order='" & Trim(cboWO) & "' order by CompPN"
'            End If
       Case "SAPFileChk"
             str = "select distinct A.Work_Order,A.Status,A.SendID,B.Sap1FileName,B.TransDateTime,B.FilePath from qsms_saphis a,SAP1_Seq_Num b  where a.work_order='" & Trim(cboWO) & "' and a.SendID=b.SendID and a.SendID<>'' order by b.Sap1FileName"
End Select

Set RS = Conn.Execute(str)
If Not RS.EOF Then
   Call CopyToExcel(RS)
Else
   MsgBox "No data"
End If

End Function
Private Function NonAVL(ByVal compPN As String)
Dim str As String
Dim RS As ADODB.Recordset
If Trim(compPN) = "" Then
   MsgBox "Please select the CompPN"
   Exit Function
End If
str = "Select * from QSMS_NonAVL where CompPN like '" & Trim(compPN) & "%'"

Set RS = Conn.Execute(str)
If Not RS.EOF Then
   Call CopyToExcel(RS)
Else
   MsgBox "No data"
End If

End Function

Private Function GetChkBOMDiff(ByVal WO As String)
Dim str As String
Dim RS As ADODB.Recordset
If Trim(WO) = "" Then
   MsgBox "Please select WO"
   Exit Function
End If
str = "Select * from QSMS_Wo_Diff where Work_Order='" & Trim(WO) & "' "

Set RS = Conn.Execute(str)
If Not RS.EOF Then
   Call CopyToExcel(RS)
Else
   MsgBox "No data"
End If

End Function

''Private Function DeleteME_BOM(ByVal MBPN As String, ByVal Rev As String, Machine As String)
''Dim Str As String
''Dim Rs As ADODB.Recordset
''Dim JobGroup As String
''If Trim(MBPN) = "" Or Trim(Rev) = "" Then
''   MsgBox "Please input MBPN or Rev"
''   Exit Function
''End If
''If Trim(Machine) = "" Then
''   MsgBox "Please input Machine (set Machine=All to delete all)"
''   Exit Function
''End If
''
'''If Trim(JobGroup) = "" Then
'''   MsgBox "Please input JobGroup (set JobGroup=All to delete all)"
'''   Exit Function
'''End If
''
''JobGroup = GetSelectingJobGroup()
''If JobGroup = "" Then
''   MsgBox "请选择 JobGroup"
''   Exit Function
''End If
''
''
''If UCase(Machine) = "ALL" Then
''    Machine = "%"
''Else
''    Machine = Machine & "%"
''End If
''
'''If JobGroup = "All" Then
'''   JobGroup = "%"
'''Else
'''   JobGroup = JobGroup & "%"
'''End If
''
''Str = "select *  FROM QSMS_MEBOM where (jobpn ='" & MBPN & "' or jobpn in (select jobpn from qsms_jobbom where mbpn='" & MBPN & "')) and JobGroup in  " & JobGroup & " and version='" & Rev & "' and Machine like '" & Machine & "'"
''Set Rs = Conn.Execute(Str)
''If Rs.EOF Then
''  MsgBox "can not find ME BOM ,Please check the MBPN or Rev"
''  Exit Function
''End If
''Str = "delete  FROM QSMS_MEBOM where (jobpn ='" & MBPN & "' or jobpn in (select jobpn from qsms_jobbom where mbpn='" & MBPN & "')) and JobGroup in " & JobGroup & "  and version='" & Rev & "' and Machine like '" & Machine & "'"
''Conn.Execute Str
'''Record who delete which MEBOM
''strsql = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Delete ME_BOM','" & Trim(MBPN) & "+" & Trim(Rev) & "+" & Machine & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
''Conn.Execute (strsql)
''
''MsgBox "Delete ME bom OK"
''
''End Function


''1169 使用PrepareMaterialByWONew代替此Funtion,取消模板
Private Function PrepareMaterialByWO(ByVal SheetName As String, ByVal Line As String)
Dim str As String, LocalPath As String
Dim RS As ADODB.Recordset
Dim strDate As String
Dim strFileName  As String, strSheetName As String
Dim blnActiveSheet As Boolean
Dim MyXlsApp As Excel.Application
Dim xlWorkSheet As Excel.Worksheet
Dim strFlag As String
Dim Shift As String
Dim WO As String
Dim Jobpn As String, step As Integer
Dim BiginDate As String
Dim EndDate As String

BiginDate = Format(dtpSDate, "YYYYMMDD")
EndDate = Format(dtpEDate, "YYYYMMDD")

On Error GoTo errHandler

step = 0
LocalPath = "D:\QSMS_Report\"
If Dir(LocalPath, vbDirectory) = "" Then MkDir LocalPath
step = 1
LblMessage = ""
Select Case SheetName
       Case "By_WorkOrder"
             WO = GetWO("BY_WorkOrder", "N")
             If Trim(WO) = "" Then
                MsgBox "Please select the work order"
                Exit Function
             End If
             str = "exec QSMSRptPrepareMaterialByWO '" & WO & "'"
       Case "By_Shift"                          ''# 2008-01-18 add by archer
             If Trim(CboShift) = "" Then
                MsgBox "Please select Shift"
                Exit Function
             End If
             Shift = Mid(CboShift, 1, 1)
             If Shift <> "D" And Shift <> "N" Then
                MsgBox "Wrong Shift for you select"
                Exit Function
             End If
             WO = GetWO("BY_Shift", "N")
             str = "exec QSMSRptPrepareMaterialByLineShift '" & WO & "'"
       Case "By_WorkOrders"
       ''''modify by Kevin 2008.12.18 (0034)
'             If txtFilePath <> "" And Me.ListWoSelecting.ListCount > 0 Then
'                wo = GetWO("BY_WorkOrders", "N")
'                Str = "exec QSMSRptPrepareMaterialByGroup '" & wo & "'"
'             Else
                WO = GetWO("BY_WorkOrders", "N")
                If Trim(WO) = "" Then
                    Exit Function
                End If
                str = "exec QSMSRptPrepareMaterialByGroup '" & WO & "'"
'             End If
             
       Case "By_Group"
             WO = GetWO("BY_Group", "N")
             If Trim(WO) = "" Then
                Exit Function
             End If
             str = "exec QSMSRptPrepareMaterialByGroup '" & WO & "'"
       Case "By_JobPN"
             WO = GetWO("BY_WorkOrders", "N")
             If cboJobPN = "" Then
                 MsgBox "Please select jobpn"
                 Exit Function
             Else
                Jobpn = Mid(cboJobPN, 1, 11)
             End If
             str = "exec QSMSRptPrepareMaterialByJobPN '" & WO & "','" & Jobpn & "'"
       
End Select
Set RS = Conn.Execute(str)
If RS.EOF Then
   MsgBox "No data"
   Exit Function
End If
step = 2

Set MyXlsApp = CreateObject("Excel.Application")
MyXlsApp.Visible = False
MyXlsApp.UserControl = True
MyXlsApp.DisplayAlerts = False

step = 3
strFileName = LocalPath & "PrepareMaterialReport_" & SheetName & ".xls"
If Dir(strFileName) <> "" Then
    Kill strFileName
End If
FileCopy App.Path & "\Template\PrepareMaterialReport.xls", strFileName
step = 4
'Copy Sample sheet to new file
blnActiveSheet = True
MyXlsApp.Workbooks.Open FileName:=strFileName
MyXlsApp.Visible = False
MyXlsApp.UserControl = True

step = 5
MyXlsApp.ActiveWorkbook.Sheets(SheetName).Activate
'MyXlsApp.ActiveWorkbook.Sheets(1).Activate

'MyXlsApp.Sheets(SheetName).Cells(1, 1).Value = "Line"
'MyXlsApp.Sheets(SheetName).Cells(1, 2).Value = Line
'MyXlsApp.Sheets(SheetName).Cells(1, 3).Value = "W/O"
'MyXlsApp.Sheets(SheetName).Cells(1, 4).Value = GetWo(SheetName, "Y")
'MyXlsApp.Sheets(SheetName).Cells(1, 5).Value = "Qty"
'MyXlsApp.Sheets(SheetName).Cells(1, 6).Value = rs.Fields(0)
'MyXlsApp.Sheets(SheetName).Cells(1, 7).Value = "JobPN"
'MyXlsApp.Sheets(SheetName).Cells(1, 8).Value = CboJobPN
If (SheetName <> "By_Shift") Then
    MyXlsApp.Sheets(SheetName).Cells(1, 1).Value = "Line"
    MyXlsApp.Sheets(SheetName).Cells(1, 2).Value = Line
    MyXlsApp.Sheets(SheetName).Cells(1, 3).Value = "W/O"
    MyXlsApp.Sheets(SheetName).Cells(1, 4).Value = GetWO(SheetName, "Y")
    MyXlsApp.Sheets(SheetName).Cells(1, 5).Value = "Qty"
    MyXlsApp.Sheets(SheetName).Cells(1, 6).Value = RS.Fields(0)
    MyXlsApp.Sheets(SheetName).Cells(1, 7).Value = "JobPN"
    MyXlsApp.Sheets(SheetName).Cells(1, 8).Value = cboJobPN
Else
    MyXlsApp.Sheets(SheetName).Cells(1, 1).Value = "Line"
    MyXlsApp.Sheets(SheetName).Cells(1, 2).Value = Line
    MyXlsApp.Sheets(SheetName).Cells(1, 3).Value = "Shift"
    MyXlsApp.Sheets(SheetName).Cells(1, 4).Value = Shift
    MyXlsApp.Sheets(SheetName).Cells(1, 5).Value = "Date Time"
    If Trim(dtpEDate) <> Trim(dtpSDate) Then
        MyXlsApp.Sheets(SheetName).Cells(1, 6).Value = BiginDate & " To " & EndDate
    Else
        MyXlsApp.Sheets(SheetName).Cells(1, 6).Value = BiginDate
    End If
    MyXlsApp.Sheets(SheetName).Cells(1, 7).Value = "Schedule Qty"
    MyXlsApp.Sheets(SheetName).Cells(1, 8).Value = RS.Fields(0)
End If
Set RS = RS.NextRecordset
MyXlsApp.Sheets(SheetName).Cells(3, 1).CopyFromRecordset RS
MyXlsApp.Visible = True
'MyXlsApp.ActiveWorkbook.SaveAs strFileName


'MyXlsApp.Quit
'Set MyXlsApp = Nothing
LblMessage.Caption = "report OK" + strAddress


Exit Function

errHandler:
    MsgBox ("PrepareMaterialByWO, Step:" & CStr(step) & ", " & Err.Description)

End Function

Private Sub ReturnDID()
'**Sandy        2008.03.05     add ReturnDIDByGroupID and ReturnDIDByWO (00014)
Dim RS As New ADODB.Recordset
Dim str As String
If CboGroupID = "" Then
    MsgBox ("the GroupID is empty.please select GroupID")
End If
Select Case CboReportType
       Case "ReturnDIDByGroupID"
            str = "exec XL_ReturnDIDByGroupID '" & Trim(CboGroupID) & "'"
       Case "ReturnDIDByWO"
            str = "exec XL_ReturnDIDByWO '" & Trim(CboGroupID) & "'"
End Select
Set RS = Conn.Execute(str)
If RS.EOF = False Then
    Call CopyToExcel(RS)
Else
    MsgBox ("NO data")
End If
End Sub
Private Sub XL_MaterialDemand()
Dim RS As New ADODB.Recordset
Dim strSQL As String
Dim sDate As String
Dim eDate As String
sDate = Format(dtpSDate, "yyyymmdd")
eDate = Format(dtpEDate, "yyyymmdd")
strSQL = "exec XL_RptMaterialDemand '" & Trim(sDate) & "','" & Trim(eDate) & "'"
    Call ToExcel(strSQL)
'Set Rs = Conn.Execute(strSQL)   ''''(0015)
'If Rs.EOF = False Then
'    Call CopyToExcel(Rs)
'End If
End Sub

Private Function CheckBOM_Rate()

Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object
Dim str As String
Dim BeginDate, EndDate As String
Dim GroupIDHead As String
Dim I As Long, fldCount As Integer, iCol As Integer
Dim RS As ADODB.Recordset
BeginDate = Format(dtpSDate, "YYYY/MM/DD")
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
BeginDate = BeginDate + "000000"

EndDate = Format(dtpEDate, "YYYY/MM/DD")
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")
EndDate = EndDate + "240000"

If chkDomain = "N" Then ''1165
    str = "Exec Wo_CheckBom_status '" & BeginDate & "','" & EndDate & "'"
    Set RS = Conn.Execute(str)
    Call CopyToExcel(RS)
    str = "Exec sap_wo_status '" & BeginDate & "','" & EndDate & "'"
    Set RS = Conn.Execute(str)
    Sleep (500)
    Call CopyToExcel(RS)
    Exit Function
End If

 Set xlApp = CreateObject("Excel.Application")
 Set xlsBook = xlApp.Workbooks.Add

 xlApp.DisplayAlerts = False
 Set xlWs = xlApp.Worksheets(1)

 xlApp.UserControl = True
 '''''1 get rate
 str = "Exec Wo_CheckBom_status '" & BeginDate & "','" & EndDate & "'"
 Set RS = Conn.Execute(str)
 fldCount = RS.Fields.Count

 For iCol = 1 To fldCount
     xlWs.Cells(1, iCol).Value = RS.Fields(iCol - 1).Name
 Next
 
  xlWs.Cells(2, 1).CopyFromRecordset RS
  
''''get summary
 str = "Exec sap_wo_status '" & BeginDate & "','" & EndDate & "'"
 Set RS = Conn.Execute(str)
 fldCount = RS.Fields.Count

 For iCol = 1 To fldCount
     xlWs.Cells(4, iCol).Value = RS.Fields(iCol - 1).Name
 Next
 
  xlWs.Cells(5, 1).CopyFromRecordset RS

 ' Auto-fit the column widths and row heights
 xlApp.Selection.CurrentRegion.Columns.AutoFit
 xlApp.Selection.CurrentRegion.Rows.AutoFit
 xlApp.Visible = True
 ' Close ADO objects
 RS.Close
 Set RS = Nothing
 'Trans_Date = Format(Now, "YYYYMMDD")

 Set xlApp = Nothing
 Set xlsBook = Nothing
End Function
Private Function RefreshBoM(ByVal WO As String)
Dim str As String
Dim RS As ADODB.Recordset
str = "select * from Qsms_WO where work_order='" & WO & "'"
Set RS = Conn.Execute(str)
If RS.EOF Then
  MsgBox "Please check bom first"
  Exit Function
End If

str = "Exec QSMS_CheckBomSP '" & Trim(WO) & "','Y'"               ''(0019)
Set RS = Conn.Execute(str)
If RS.EOF = False Then
    MsgBox "Check bom fail"
End If

''If CheckBom(WO, "Y") = False Then
''   MsgBox "Check bom fail"
''End If

str = "select *  from Sap_BOM_Fail  where Work_Order ='" & Trim(WO) & "' "
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   Call CopyToExcel(RS)
Else
   MsgBox "refresh BOM OK"
End If
End Function

Private Function GetWO(ByVal Ctype As String, cOutPut As String) As String
Dim str As String, I As Integer
Dim RS As ADODB.Recordset
Dim WO As String
Dim WoOutPut As String
Dim Shift As String
Dim Line As String
Dim sDateTime, eDateTime As String
sDateTime = Format(dtpSDate, "YYYY/MM/DD")
sDateTime = Replace(sDateTime, "-", "")
sDateTime = Replace(sDateTime, "/", "")
sDateTime = sDateTime + "0000"
eDateTime = Format(dtpEDate, "YYYY/MM/DD")
eDateTime = Replace(eDateTime, "-", "")
eDateTime = Replace(eDateTime, "/", "")
eDateTime = eDateTime + "2400"
WO = ""
WoOutPut = ""



Select Case UCase(Ctype)
       Case "BY_WORKORDERS"
            For I = 0 To ListWoSelecting.ListCount - 1
                              ListWoSelecting.ListIndex = I
                              WO = WO & Trim(ListWoSelecting.Text) & ","
                              WoOutPut = WoOutPut & ListWoSelecting.Text & Chr(10) & Chr(13)
            Next I
            If WO <> "" Then
                WO = Mid(WO, 1, Len(WO) - 1)
            Else
               MsgBox "Please select the work order list"
            End If
       Case "BY_GROUP"
            If Trim(CboGroupID) = "" Then
               MsgBox "Please select GroupID"
               Exit Function
            End If
            str = "select Work_Order from QSMS_WoGroup where GroupID='" & Trim(CboGroupID) & "'"
            Set RS = Conn.Execute(str)
            While Not RS.EOF
                 WO = WO & Trim(RS!Work_Order) & ","
                 WoOutPut = WoOutPut & Trim(RS!Work_Order) & Chr(10) & Chr(13)
                 RS.MoveNext
            Wend
             WO = Mid(WO, 1, Len(WO) - 1)
       Case "BY_WORKORDER"
              WO = Trim(txtWO)
              WoOutPut = WO
       Case "BY_SHIFT"                      ''# 2008-01-18 add by archer
            Shift = Mid(Trim(CboShift), 1, 1)
            Line = Mid(Trim(CboLine), 1, 1)
            str = "select distinct Wo from  XL_WOPlanSeq where shift = '" & Shift & "' and Line = '" & Line & "' and BeginDateTime between '" & sDateTime & "' and '" & eDateTime & "'"
            Set RS = Conn.Execute(str)
            If RS.EOF Then
                MsgBox "No data"
                Exit Function
            End If
            While Not RS.EOF
                 WO = WO & Trim(RS!WO) & ","
                 WoOutPut = WoOutPut & Trim(RS!WO) & Chr(10) & Chr(13)
                 RS.MoveNext
            Wend
             WO = Mid(WO, 1, Len(WO) - 1)
       Case Else
       
End Select
If cOutPut = "Y" Then
   GetWO = WoOutPut
Else
   GetWO = WO
End If
End Function
Private Function GetJobGroupByJobRev(ByVal Machine As String, ByVal Jobpn As String, Version As String)
Dim str As String
Dim RS As ADODB.Recordset
If Len(Machine) > 6 Then
    str = "Select distinct jobgroup from qsms_mebom where (jobpn='" & Jobpn & "' or jobpn in (select distinct jobpn from QSMS_JObBom where MBPN='" & Jobpn & "')) and version='" & Version & "'" & _
          " and Machine like '" & Machine & "%' "
Else
    str = "Select distinct jobgroup from qsms_mebom where (jobpn='" & Jobpn & "' or jobpn in (select distinct jobpn from QSMS_JObBom where MBPN='" & Jobpn & "')) and version='" & Version & "'"
End If
Set RS = Conn.Execute(str)
ListAllJobGroup.Clear
ListselectingJobGroup.Clear
While Not RS.EOF
     ListAllJobGroup.AddItem Trim(RS!jobgroup)
     RS.MoveNext
     
Wend

End Function
Private Function GetSelectingJobGroup() As String
Dim SelectingJobGroup As String, I As Integer

    If ListselectingJobGroup.ListCount <= 0 Then
           If Len(Trim(TxtJobGroup)) > 0 And InStr(1, Trim(TxtJobGroup), "-") > 0 Then
               GetSelectingJobGroup = "('" & Trim(TxtJobGroup) & "')"
           Else
               SelectingJobGroup = ""
           End If
           Exit Function
    End If
    
    For I = 1 To ListselectingJobGroup.ListCount
        ListselectingJobGroup.ListIndex = I - 1
        SelectingJobGroup = SelectingJobGroup + "'" + Trim(ListselectingJobGroup.Text) + "'" + ","
    
    Next I
    
    GetSelectingJobGroup = "(" + Mid(SelectingJobGroup, 1, Len(SelectingJobGroup) - 1) + ")"
End Function

Private Sub txtDID_Click()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtDID_LostFocus()
    If Trim(TxtDID) = "" Then
        TxtDID.SetFocus
        MsgBox "Please input DID !"
    End If
End Sub

Private Sub TxtMBPN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtRev.Enabled = True
   txtRev.SetFocus
End If
End Sub

Private Sub TxtRev_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call GetJobGroupByJobRev(Trim(CboMachine), Trim(TxtMBPN), Trim(txtRev))
End If
End Sub


Private Function PDUsedByCompLine(ByVal Line As String, ByVal Comp As String)
Dim str As String
Dim RS As ADODB.Recordset
str = "select * from qsms_verify where machine like '" & Trim(Line) & "%' and comppn='" & Comp & "' Order by Machine,Slot,LR,BegindateTime Desc"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
  Call CopyToExcel(RS)
Else
   MsgBox "No Data...."
End If

End Function
Private Function GETGROUPIDDATABYCOMPPN(ByVal GroupID As String, ByVal compPN As String)
Dim str As String
Dim RS As ADODB.Recordset
str = "exec GetGroupIDDataByCompPN '" & GroupID & "', '" & compPN & "'"         ''''0029
Set RS = Conn.Execute(str)
If chkDomain = "N" Then '1165
    Call CopyToExcel(RS)
Else
    Call CopyToExcel(RS)
    Set RS = RS.NextRecordset
    Call CopyToExcel(RS)
    Set RS = RS.NextRecordset
    Call CopyToExcel(RS)
    Set RS = RS.NextRecordset   ''''1118
    If Not RS Is Nothing Then
        Call CopyToExcel(RS)
    End If
End If
End Function


Private Function CheckUnCloseGroupID()
Dim str As String
Dim RS As New ADODB.Recordset
   'str = "select GroupID,substring(GroupID,2,4)+'/'+substring(GroupID,6,2)+'/'+substring(GroupID,8,2) as Create_date from qsms_wogroup where closedflag<>'Y'"
   str = "select distinct a.GroupID,substring(GroupID,4,4)+'/'+substring(GroupID,8,2)+'/'+substring(GroupID,10,2) as Create_date from qsms_wogroup a,sap_wo_list b where a.work_order=b.wo and a.closedflag<>'Y' order by create_date desc" ''(1056)
   Set RS = Conn.Execute(str)
   If Not RS.EOF Then
        Call CopyToExcel(RS)
   Else
      MsgBox "All wo has been closed!!"
      Exit Function
   End If
      
End Function

Private Function GetMultiLine(sInputStr As String, sType As String) As String
    Dim sSql As String
    Dim Rst As New ADODB.Recordset
    Dim sLineList As String
    If sType = "WO" Then
        'sSql = "select distinct left(Machine,1) as Line from QSMS_WO where Work_Order=" & sq(sInputStr)
        sSql = "select rtrim(Line)+ltrim(Side) as Line from WO_MultiLine where WO=" & sq(sInputStr)
    Else
        'sSql = "select distinct left(Machine,1) as Line from QSMS_WO where jobgroup in " & sInputStr   '已带括号
        
    End If
    Rst.Open sSql, Conn, adOpenKeyset, adLockReadOnly
    Do While Rst.EOF = False
        sLineList = sLineList & Rst("Line") & ","
        Rst.MoveNext
    Loop
    
    If sLineList <> "" Then
        sLineList = Left(sLineList, Len(sLineList) - 1)
        
    End If
    GetMultiLine = sLineList
    
    
    Set Rst = Nothing
End Function
'**Sandy      2008.01.24     增加 ReturnDID to WH 状况报表----------------(0011)
Private Function DispatchQTYByWO()
    Dim strSQL As String, sDate As String, eDate As String, strWO As String
    Dim RS As New ADODB.Recordset
    'sDate = Format(dtpSDate, "YYYYMMDD") '1060
    'eDate = Format(dtpEDate + 1, "YYYYMMDD") '1060
    strWO = Trim(txtWO)
    
   ' If dtpSDate > dtpEDate Then '1060
     '   MsgBox ("The StartDate must be smaller than Today !"), vbCritical
    '    Exit Function
   ' End If
    
      If strWO = "" Then
        MsgBox ("WO Can not be empty !"), vbCritical
        Exit Function
    End If
    
    strSQL = " EXEC QSMS_DispatchQTYByWO '','','" & Trim(txtWO) & "'"
    Set RS = Conn.Execute(strSQL)
    If Not RS.EOF Then
           Call CopyToExcel(RS)
    Else
           MsgBox ("No Data"), vbCritical
    End If

End Function

Private Function QSMS_DID_ToWH()
    Dim strSQL As String, sDate As String, eDate As String, strWO As String
    Dim RS As New ADODB.Recordset
    sDate = Format(dtpSDate, "YYYYMMDD")
    eDate = Format(dtpEDate + 1, "YYYYMMDD")
    strWO = Trim(txtWO)
    
    If dtpSDate > dtpEDate Then
        MsgBox ("The StartDate must be smaller than Today !"), vbCritical
        Exit Function
    End If
    
    strSQL = "select substring(TransDateTime,5,2)+'/'+substring(TransDateTime,7,2)as [Date],count(*) as TotalNumber from QSMS_DID_ToWH where ToWHType='Return' and isGood='Y' and WareHouseID<>'' " & _
                "and TransDateTime Between '" & Trim(sDate) & "' and '" & Trim(eDate) & "' group by substring(TransDateTime,5,2)+'/'+substring(TransDateTime,7,2)" & _
                "order by substring(TransDateTime,5,2)+'/'+substring(TransDateTime,7,2) "
    Set RS = Conn.Execute(strSQL)
    If Not RS.EOF Then
           Call Copy2Excel(RS)
    Else
           MsgBox ("No Data"), vbCritical
    End If

End Function

'**Kevin      2008.12.23     增加查询QSMS_WO的数据----------------(0034)
Private Function QSMS_WO()
    Dim strSQL As String, strWO As String
    Dim RS As New ADODB.Recordset
    If txtFilePath <> "" And Me.ListWoSelecting.ListCount > 0 Then
        strWO = GetWO("BY_WorkOrders", "N")
        strSQL = "exec QSMSRpt_QSMS_WO '" & strWO & "'"
    Else
        MsgBox "Please upload the wo list !", vbCritical
        Exit Function
    End If
    
    Set RS = Conn.Execute(strSQL)
    If Not RS.EOF Then
        Call CopyToExcel(RS)
    Else
        MsgBox ("No Data"), vbCritical
    End If
    
End Function

'add PrepariMaterialMonitor report-------------(0009)-----------20070116 by Sandy-----------
Private Function XL_MonitorReport()
    Dim strSQL As String, sDate As String, eDate As String, strWO As String
    Dim RS As New ADODB.Recordset
    sDate = Format(dtpSDate, "YYYYMMDD")
    eDate = Format(dtpEDate + 1, "YYYYMMDD")
    strWO = Trim(txtWO)
    
    If dtpSDate > dtpEDate Then
        MsgBox ("The StartDate must be smaller than Today !"), vbCritical
        Exit Function
    End If
    
    strSQL = "exec XL_MonitorReport '" & sDate & "'"
    Set RS = Conn.Execute(strSQL)
    If Not RS.EOF Then
        Call CopyToExcel(RS)
    Else
        MsgBox ("No Data"), vbCritical
    End If

End Function
'**Sandy      2008.02.02     add XL_ReelBaseQty report----------------(0012)
Private Function XL_ReelBaseQty()
    Dim strSQL As String, sDate As String, eDate As String, strCompPN As String
    Dim RS As New ADODB.Recordset
    sDate = Format(dtpSDate, "YYYYMMDD")
    eDate = Format(dtpEDate + 1, "YYYYMMDD")
    strCompPN = Trim(CboComp)
    
    If dtpSDate > dtpEDate Then
        MsgBox ("The StartDate must be smaller than Today !"), vbCritical
        Exit Function
    End If
    
    strSQL = "select A.Plant,A.CompPN,A.BaseReelQty,B.Location,B.StockQty,B.Status,B.WorkDate,B.Shift,B.TransDateTime from XL_ReelBaseQty A, XL_StockQtyByLocation B where A.comppn=B.comppn AND A.comppn='" & strCompPN & "' order by B.transdatetime desc "
    Set RS = Conn.Execute(strSQL)
    If Not RS.EOF Then
           Call CopyToExcel(RS)
    Else
           MsgBox ("No Data"), vbCritical
    End If

End Function
'''''''''''''''''''''''''''''''''''added by Jing 20071212 (0003) ''''''''''''''''''''''''''''''''
Private Function GetWoInputPlanBySide()
Dim strSQL As String, sDate As String, eDate As String, strWO As String, str As String
Dim RS As New ADODB.Recordset

On Error GoTo errhandle
    sDate = Format(dtpSDate, "YYYYMMDD")
    eDate = Format(dtpEDate + 1, "YYYYMMDD")
    strWO = Trim(txtWO)
    
    If dtpSDate > dtpEDate Then
        MsgBox ("The StartDate must be smaller than EndDate !"), vbCritical
        Exit Function
    End If
    
    If dtpEDate - dtpSDate > 31 Then
        MsgBox ("The day must less than 31 days !"), vbCritical
        Exit Function
    End If
    strSQL = "exec XL_GetWoInputPlanBySide '" & sDate & "','" & eDate & "','" & strWO & "'"
    Call ToExcel(strSQL)
    Exit Function
    
errhandle:
    MsgBox Err.Description
End Function
Private Function GetWoInputPlan()
Dim strSQL As String, sDate As String, eDate As String, strWO As String
Dim rsTmp As New ADODB.Recordset

On Error GoTo errhandle
    sDate = Format(dtpSDate, "YYYYMMDD")
    eDate = Format(dtpEDate + 1, "YYYYMMDD")
    strWO = Trim(txtWO)
    
    If dtpSDate > dtpEDate Then
        MsgBox ("The StartDate must be smaller than EndDate !"), vbCritical
        Exit Function
    End If
    
    If dtpEDate - dtpSDate > 31 Then
        MsgBox ("The day must less than 31 days !"), vbCritical
        Exit Function
    End If
    
    strSQL = "exec Rpt_XL_GetWoInputPlan '" & sDate & "','" & eDate & "','" & strWO & "'"
    Set RS = Conn.Execute(strSQL)
    Call CopyToExcel(RS)
    Set RS = RS.NextRecordset
    If RS.EOF = False Then
        Call CopyToExcel(RS)
    Else
        MsgBox "No Detail data"
    End If
    Exit Function
    
errhandle:
    MsgBox Err.Description
End Function

Private Sub XL_DemandDetail()
Dim RS As New ADODB.Recordset
Dim strSQL As String

strSQL = "select w.GroupID, x.* from xl_qsms_wo x join qsms_wogroup w on x.work_order=w.work_order where x.workdate between '" & Format(dtpSDate, "yyyymmdd") & "' and '" & Format(dtpEDate, "yyyymmdd") & "' order by x.workdate,x.shift,x.line,x.work_order,x.comppn"
Set RS = Conn.Execute(strSQL)
If RS.EOF = False Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found!"
End If
End Sub
Private Sub XL_DispatchStatus() '=====================(0010)
Dim RS As New ADODB.Recordset
Dim strSdate As String
Dim strEDate As String
Dim strShift As String
If DateDiff("d", dtpSDate, dtpEDate) > 3 Then
    MsgBox "Date range over 3 days!", vbCritical
    Exit Sub
End If
If CboShift = "" Then
    MsgBox "Please select the shift!", vbCritical
    Exit Sub
End If
If CboLine = "" Then
    MsgBox "Please select the line", vbCritical
    Exit Sub
End If
strSdate = Format(dtpSDate, "yyyymmdd")
strEDate = Format(dtpEDate, "yyyymmdd")
strShift = Left(CboShift, 1)
strSQL = "exec XL_DispatchStatus '" & Trim(CboLine) & "','" & strSdate & "','" & strEDate & "','" & strShift & "'"
Set RS = Conn.Execute(strSQL)
If RS.EOF = False Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found!"
End If
End Sub
'---- (0020)
Private Sub GetAllDispatchInforByGroupID(GroupID As String)
Dim RS As New ADODB.Recordset
If GroupID = "" Then
    MsgBox "Please select one groupid"
    Exit Sub
End If
strSQL = "exec QSMS_GetAllDispatchInforByGroupID '" & GroupID & "'"
Set RS = Conn.Execute(strSQL)
If RS.EOF = False Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data found"
End If
End Sub

'''''''''''''''''''''''''Mark by Jing 2008.01.09    (0007)''''''''''''''''''''''
'''''''''''''''''''''''''Added by Jing 2008.01.08   (0006)''''''''''''''''''''''
'Private Sub GetNoUseDID()
'Dim tmpRS As New ADODB.Recordset
'Dim strSQL As String
'
'On Err GoTo ErrHandler:
'    strSQL = "exec Report_NOUseDID"
'    Set tmpRS = Conn.Execute(strSQL)
'    If tmpRS.EOF Then
'        MsgBox ("NO DATA !")
'    Else
'        Call CopyToExcel(tmpRS)
'    End If
'    Exit Sub
'ErrHandler:
'    MsgBox Err.Description, vbCritical
'End Sub

'''''''''''''''''Added by Jing 2008.06.01   (0021)''''''''''''
Private Sub MEBom_Model(tmpJobPN As String, tmpRev As String)
Dim strSQL As String
Dim tmpRS As New ADODB.Recordset

On Error GoTo errHandler:

If Trim(TxtMBPN) = "" Or Trim(txtRev) = "" Then
    MsgBox ("Please input JobPN and Revision !")
    TxtMBPN.SetFocus
Else
    strSQL = "select * from qsms_mebom where JobPN='" & Trim(tmpJobPN) & "' and Version='" & Trim(tmpRev) & "'"
    Set tmpRS = Conn.Execute(strSQL)
    If tmpRS.EOF Then
        MsgBox ("NO DATA")
    Else
        Call CopyToExcel(tmpRS)
    End If
End If

Exit Sub
errHandler:
    MsgBox Err.Description
End Sub
Private Sub DIDCompare()  '(0027)
Dim sDate As String
Dim eDate As String
Dim strSQL As String
Dim RS As New ADODB.Recordset
On Error GoTo ErrHdl:
sDate = Format(dtpSDate, "yyyymmdd") & "080000"
eDate = Format(dtpEDate + 1, "yyyymmdd") & "080000"

''20100615   Denver    Add OPID for Compare DID
''20100705   Denver    DID='' for check CompPN action data(delete DID<>'')
strSQL = "SELECT Line  as Line,case when substring(transdatetime,9,6)>'080000' and substring(transdatetime,9,6)<'20000' " & _
         "then 'D' else 'N' end as Shift,Machine,slot+'-'+cast(lr as char(1)) as Slot,DID,NewDID,ScanDID,TransDateTime," & _
         "case CheckResult when 'Y' then N'相同' else N'不同' end as CheckResult,OPID " & _
         "FROM qsms_checkcomplog WHERE Slot<>'' and transdatetime between " & sq(sDate) & " and " & sq(eDate)
''添加Slot不等于空1260
Set RS = Conn.Execute(strSQL)
If RS.EOF = False Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data"
End If
Exit Sub
ErrHdl:
    MsgBox Err.Description
End Sub
Private Sub CheckSpliceReplacePN()  '(0031)
Dim sDate As String
Dim eDate As String
Dim strSQL As String
Dim RS As New ADODB.Recordset
On Error GoTo ErrHdl:
sDate = Format(dtpSDate, "yyyymmdd") & "074000"
eDate = Format(dtpEDate + 1, "yyyymmdd") & "074000"
strSQL = "exec QSMS_rptChkSpliceReplacePN '" & sDate & "','" & eDate & "'"
Set RS = Conn.Execute(strSQL)
If RS.EOF Then
    MsgBox "No data"
Else
    Call CopyToExcel(RS)
End If
Exit Sub
ErrHdl:
    MsgBox Err.Description
End Sub
Public Sub DIDIntegration(ByVal GroupID As String)
Dim sDate As String
Dim eDate As String
Dim strSQL As String
Dim RS As New ADODB.Recordset
 
 

sDate = Format(dtpSDate, "YYYYMMDD")
eDate = Format(dtpEDate, "YYYYMMDD")


strSQL = "exec QSMS_DIDIntegration @Item=" & sq("Report") & ", @GroupID=" & sq(GroupID) & " , @BeginTime=" & sDate & " , @EndTime=" & sq(eDate)
Set RS = Conn.Execute(strSQL)
If RS.EOF Then
    MsgBox "No data"
Else
    Call CopyToExcel(RS)
End If

End Sub

Private Sub ForbiddenPN() '(0032)
Dim strSQL As String
Dim RS As New ADODB.Recordset
Dim sDate As String
Dim eDate As String
On Error GoTo ErrHdl:
sDate = Format(dtpSDate, "yyyymmdd") & "074000"
eDate = Format(dtpEDate + 1, "yyyymmdd") & "074000"
'1056
strSQL = "select top 60000 RefID,ModelName,PN,VendorCode,DateCode,LotCode,Status,[User],TransDateTime,LastUpdateTime,Trans_Flag,''as DelDateTime " & _
         "from forbiddenpn where PN like '" & Trim(CboComp) & "%' and transdatetime between '" & sDate & "' and '" & eDate & "' " & _
         "Union All " & _
         "select top 60000 RefID,ModelName,PN,VendorCode,DateCode,LotCode,Status,[User],TransDateTime,LastUpdateTime,Trans_Flag,DelDateTime" & _
         " from forbiddenpn_trace WHERE PN like '" & Trim(CboComp) & "%' and transdatetime between '" & sDate & "' and '" & eDate & "' "
Set RS = Conn.Execute(strSQL)
If RS.EOF = False Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data"
End If
Exit Sub
ErrHdl:
    MsgBox Err.Description
End Sub
Private Sub Glue_DataByDay() '(0032)
Dim strSQL As String
Dim RS As New ADODB.Recordset
Dim sDate As String
Dim eDate As String
On Error GoTo ErrHdl:
sDate = Format(dtpSDate, "yyyy/mm/dd")
eDate = Format(dtpEDate, "yyyy/mm/dd")
strSQL = "EXEC Glue_DataByDay '" & sDate & "' ,'" & eDate & "'"
Set RS = Conn.Execute(strSQL)
If RS.EOF = False Then
    If chkDomain = "N" Then ''1165
        Call CopyToExcel(RS)
        Exit Sub
    Else
        Call CopyToExcel(RS)
        Set RS = RS.NextRecordset
        If RS.EOF = False Then
            Call CopyToExcel(RS)
        Else
            MsgBox "No Detail data"
        End If
    End If
Else
    MsgBox "No data"
End If
Exit Sub
ErrHdl:
    MsgBox Err.Description
End Sub

Private Sub Glue_Consumption() '(0032)
Dim strSQL As String
Dim RS As New ADODB.Recordset
Dim sDate As String
Dim eDate As String
On Error GoTo ErrHdl:
sDate = Format(dtpSDate, "yyyymmdd")
eDate = Format(dtpEDate, "yyyymmdd")
strSQL = "EXEC RPTGlue_Consumption '" & sDate & "' ,'" & eDate & "'"
Set RS = Conn.Execute(strSQL)
If RS.EOF = False Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data"
End If
Exit Sub
ErrHdl:
    MsgBox Err.Description
End Sub

Private Sub Glue_CallOff() '(0038)
Dim strSQL As String
Dim RS As New ADODB.Recordset
Dim sDate As String
Dim eDate As String
On Error GoTo ErrHdl:
sDate = Format(dtpSDate, "yyyymmdd")
eDate = Format(dtpEDate, "yyyymmdd")
strSQL = "EXEC RPTGlue_CallOff '" & sDate & "' ,'" & eDate & "'"
Set RS = Conn.Execute(strSQL)
If RS.EOF = False Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data"
End If
Exit Sub
ErrHdl:
    MsgBox Err.Description
End Sub

Private Sub MaterialReturn() '(0035)
Dim strSQL As String
Dim RS As New ADODB.Recordset
Dim sDate As String
Dim eDate As String
On Error GoTo ErrHdl:
sDate = Format(dtpSDate, "yyyymmdd") & "000000"
eDate = Format(dtpEDate + 1, "yyyymmdd") & "240000"
strSQL = "select BU,  ReferenceID, Status, DID, CompPN, Qty, VendorCode, DateCode, LotCode, OldDID, OldDIDDateTime, ToWHType,SAPClient, InPlant, OutPlant, OutLineMC, WHType, BatchNo, Material_Cost_Center, IsGood, WareHouseID, UID, TransDateTime, GenRefIDDateTime, WHTransDateTime" & _
         " from dbo.QSMS_DID_ToWH where CompPN like '" & Trim(CboComp) & "%' and TransDateTime between '" & sDate & "' and '" & eDate & "' order by ReferenceID,TransDateTime"
Set RS = Conn.Execute(strSQL)
If RS.EOF = False Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data"
End If
Exit Sub
ErrHdl:
    MsgBox Err.Description
End Sub

Private Sub PanalnterLock()
Dim strSQL As String
Dim RS As New ADODB.Recordset
Dim sDate As String
Dim eDate As String
On Error GoTo ErrHdl:
sDate = Format(dtpSDate, "yyyymmdd") & "000000"
eDate = Format(dtpEDate + 1, "yyyymmdd") & "240000"
strSQL = "select Machine,FeederID,DID,Slot,LR,JobPN,BeginDatetime as TransDateTime from QSMS_Verify " & _
         "where BeginDatetime between '" & sDate & "' and '" & eDate & "' and EndDatetime='' order by Machine,Slot,LR"
Set RS = Conn.Execute(strSQL)
If RS.EOF = False Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data"
End If
Exit Sub
ErrHdl:
    MsgBox Err.Description
End Sub

Private Sub FUJI_AVLList() '0041
Dim strSQL As String
Dim RS As New ADODB.Recordset
On Error GoTo ErrHdl:
strSQL = "Select * from QSMS_FujiAVL Where [group] in(select [Group] from sap_wo_list where wo='" & Trim(txtWO) & "') order by TransDateTime Desc"
Set RS = Conn.Execute(strSQL)
If RS.EOF = False Then
    Call CopyToExcel(RS)
Else
    MsgBox "No data"
End If
Exit Sub
ErrHdl:
    MsgBox Err.Description
End Sub

Private Function CheckBom_Log(ByVal WO As String)
    Dim str As String
    Dim RS As ADODB.Recordset

    str = "select DID as Wo, User_Name,Trans_date from qsms_log where did='" & WO & "' and system_name='SMT_QSMS' and Event_No='CheckBom' order by trans_date desc"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
       Call CopyToExcel(RS)
    Else
       MsgBox "No data found!!"
    End If
End Function
'0073
Private Function CheckBom_Result(ByVal WO As String)
    Dim str As String
    Dim RS As ADODB.Recordset
    
    str = "select DID as Wo,user_Name,case when ReturnQty='0' then 'Y' else 'N' end as Result, Trans_Date from qsms_log where did='" & WO & "' and system_name='SMT_QSMS' and Event_No='CheckBOMResult' order by trans_date desc"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
        Call CopyToExcel(RS)
    Else
        MsgBox "NO Data Found!!"
    End If
End Function
Private Function SplicePN()  '（1077）
    Dim str As String
    Dim RS As ADODB.Recordset
    
    str = "Select * from QSMS_Log with(nolock) where Event_No like '%SplicePN%' order by Trans_Date desc"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
       Call CopyToExcel(RS)
    Else
       MsgBox "No data found!!"
    End If
End Function

Private Function SpliceReplacePN()    '（1077）
    Dim str As String
    Dim RS As ADODB.Recordset
    
    str = "Select * from QSMS_Log where Event_No like '%SpliceReplacePN%' order by Trans_Date desc"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
       Call CopyToExcel(RS)
    Else
       MsgBox "No data found!!"
    End If
End Function
Private Function MaintainFeeder()  '（1082）
    Dim str As String
    Dim RS As ADODB.Recordset
    
    str = "Select * from QSMS_Log with(nolock) where Event_No='MaintainFeeder' order by Trans_Date desc"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
       Call CopyToExcel(RS)
    Else
       MsgBox "No data found!!"
    End If
End Function
Private Function PDA_DistributeDIDLog()  '（1089）
    Dim str As String
    Dim RS As ADODB.Recordset
    Dim sDate As String
    Dim eDate As String
    sDate = Format(dtpSDate, "yyyymmdd") & "000000"
    eDate = Format(dtpEDate, "yyyymmdd") & "240000"
    str = "Select * from PDA_DistributeDIDLog with(nolock) where line like '" & Trim(CboLine) & "%' and TransDateTime between '" & sDate & "' and '" & eDate & "' order by TransDateTime desc"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
       Call CopyToExcel(RS)
    Else
       MsgBox "No data found!!"
    End If
End Function
Private Function GetMEBom_ByGroupID(ByVal GroupID As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim TempGroup As String
Dim TempJobPn As String
Dim Machine As String

If Trim(CboGroupID.Text) = "" Then
    MsgBox "Please Input GroupID !!!", vbOKOnly, "Fail"
    Exit Function
End If
str = "EXEC QSMS_MEBOM_ByGroupID '" & GroupID & "'"
Set RS = Conn.Execute(str)

If Not RS.EOF Then
    Call CopyToExcel(RS)
End If

End Function

Private Function GetMEBom_EQProgram(ByVal FullJobGroup As String)    ''1219
Dim str As String
Dim RS As ADODB.Recordset

If Trim(TxtJobGroup.Text) = "" Then
    MsgBox "Please Input JobGroup !!!", vbOKOnly, "Fail"
    Exit Function
End If
str = "EXEC QSMS_MEBOM_ByEQProgram '" & FullJobGroup & "'"
Set RS = Conn.Execute(str)

If Not RS.EOF Then
    Call CopyToExcel(RS)
End If

End Function


Private Function GetReportType() ''1168
Dim str As String
Dim RS As ADODB.Recordset
str = "select distinct value from Program_DefineItem where AppName='QSMS' and FuncType='Report' and item='ReportType'"
Set RS = ConnSMT.Execute(str)
CboReportType.Clear
While Not RS.EOF
    CboReportType.AddItem RS!Value
    RS.MoveNext
Wend
End Function


Private Function PrepareMaterialByWONew(ByVal SheetName As String, ByVal Line As String) ''1169
Dim str As String, LocalPath As String
Dim RS As ADODB.Recordset
Dim strDate As String
Dim strFileName  As String, strSheetName As String
Dim blnActiveSheet As Boolean
Dim MyXlsApp As Excel.Application
Dim xlWorkSheet As Excel.Worksheet
Dim strFlag As String
Dim Shift As String
Dim WO As String
Dim Jobpn As String, step As Integer
Dim BiginDate As String
Dim EndDate As String


On Error GoTo errHandler


Select Case SheetName
       Case "By_WorkOrder"
             WO = GetWO("BY_WorkOrder", "N")
             If Trim(WO) = "" Then
                MsgBox "Please select the work order"
                Exit Function
             End If
             str = "exec QSMSRptPrepareMaterialByWO '" & WO & "'"
       Case "By_Shift"                          ''# 2008-01-18 add by archer
             If Trim(CboShift) = "" Then
                MsgBox "Please select Shift"
                Exit Function
             End If
             Shift = Mid(CboShift, 1, 1)
             If Shift <> "D" And Shift <> "N" Then
                MsgBox "Wrong Shift for you select"
                Exit Function
             End If
             WO = GetWO("BY_Shift", "N")
             If Trim(WO) = "" Then
                    Exit Function
             End If
                str = "exec QSMSRptPrepareMaterialByLineShift '" & WO & "'"
       Case "By_WorkOrders"
       ''''modify by Kevin 2008.12.18 (0034)
'             If txtFilePath <> "" And Me.ListWoSelecting.ListCount > 0 Then
'                wo = GetWO("BY_WorkOrders", "N")
'                Str = "exec QSMSRptPrepareMaterialByGroup '" & wo & "'"
'             Else
                WO = GetWO("BY_WorkOrders", "N")
                If Trim(WO) = "" Then
                    Exit Function
                End If
                str = "exec QSMSRptPrepareMaterialByGroup '" & WO & "'"
'             End If
             
       Case "By_Group"
             WO = GetWO("BY_Group", "N")
             If Trim(WO) = "" Then
                Exit Function
             End If
             str = "exec QSMSRptPrepareMaterialByGroup '" & WO & "'"
       Case "By_JobPN"
             WO = GetWO("BY_WorkOrders", "N")
             If cboJobPN = "" Then
                 MsgBox "Please select jobpn"
                 Exit Function
             Else
                Jobpn = Mid(cboJobPN, 1, 11)
             End If
             str = "exec QSMSRptPrepareMaterialByJobPN '" & WO & "','" & Jobpn & "'"
       
End Select

Call ToExcel(str)


Exit Function

errHandler:
    MsgBox ("PrepareMaterialByWONew, " & Err.Description)

End Function


