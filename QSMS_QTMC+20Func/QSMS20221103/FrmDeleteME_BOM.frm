VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmDeleteME_BOM 
   Caption         =   "DeleteME_BOM(2023/05/16)"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      ForeColor       =   &H00808080&
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      Begin VB.CommandButton CmdDeleteByLine 
         BackColor       =   &H8000000D&
         Caption         =   "Delete By Line"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         MaskColor       =   &H00808080&
         TabIndex        =   55
         Top             =   6600
         Width           =   1335
      End
      Begin VB.TextBox txtside 
         Height          =   375
         Left            =   1560
         TabIndex        =   54
         Top             =   6720
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.ComboBox cboslot 
         Height          =   360
         Left            =   1560
         TabIndex        =   52
         Top             =   6240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton CmdDelete 
         BackColor       =   &H8000000D&
         Caption         =   "Delete ME BOM"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         MaskColor       =   &H00808080&
         TabIndex        =   50
         Top             =   6600
         Width           =   1335
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
         TabIndex        =   28
         Top             =   3360
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
         TabIndex        =   27
         Top             =   2400
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
         TabIndex        =   26
         Top             =   120
         Width           =   2655
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
         Picture         =   "FrmDeleteME_BOM.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptRelease 
         Caption         =   "Release"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "Group"
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   240
         Width           =   1695
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
         TabIndex        =   22
         Top             =   2880
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
         Left            =   1560
         TabIndex        =   21
         Top             =   4800
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
         Left            =   6600
         TabIndex        =   20
         Top             =   1080
         Width           =   2655
      End
      Begin VB.ComboBox CboComp 
         Height          =   360
         Left            =   1560
         TabIndex        =   19
         Top             =   4320
         Width           =   2655
      End
      Begin VB.ComboBox CboMachine 
         Height          =   360
         Left            =   1560
         TabIndex        =   18
         Top             =   3840
         Width           =   2655
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
         Left            =   1560
         TabIndex        =   17
         Top             =   5280
         Width           =   2655
      End
      Begin VB.ListBox ListWoSelecting 
         Height          =   1425
         ItemData        =   "FrmDeleteME_BOM.frx":0442
         Left            =   7440
         List            =   "FrmDeleteME_BOM.frx":0444
         TabIndex        =   16
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ListBox ListWoall 
         Height          =   1425
         ItemData        =   "FrmDeleteME_BOM.frx":0446
         Left            =   4320
         List            =   "FrmDeleteME_BOM.frx":0448
         TabIndex        =   15
         Top             =   2160
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2160
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2640
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3120
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3600
         Width           =   495
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox CboJobPN 
         Height          =   360
         Left            =   1560
         TabIndex        =   9
         Top             =   1920
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
         TabIndex        =   8
         Top             =   600
         Width           =   2655
      End
      Begin VB.ListBox ListAllJobGroup 
         Height          =   1425
         ItemData        =   "FrmDeleteME_BOM.frx":044A
         Left            =   4320
         List            =   "FrmDeleteME_BOM.frx":044C
         TabIndex        =   7
         Top             =   4440
         Width           =   2175
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5880
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5400
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4920
         Width           =   495
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4440
         Width           =   495
      End
      Begin VB.ListBox ListselectingJobGroup 
         Height          =   1425
         ItemData        =   "FrmDeleteME_BOM.frx":044E
         Left            =   7440
         List            =   "FrmDeleteME_BOM.frx":0450
         TabIndex        =   2
         Top             =   4440
         Width           =   1815
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
         Left            =   1560
         TabIndex        =   1
         Top             =   5760
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   29
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
         Format          =   124715011
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   30
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
         Format          =   124715011
         CurrentDate     =   36482
      End
      Begin VB.Label lblside 
         BackColor       =   &H0000FF00&
         Caption         =   "side"
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
         TabIndex        =   53
         Top             =   6720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblslot 
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
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   6240
         Visible         =   0   'False
         Width           =   1335
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
         TabIndex        =   49
         Top             =   3360
         Width           =   1335
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
         TabIndex        =   48
         Top             =   2400
         Width           =   1335
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
         TabIndex        =   47
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
         Left            =   120
         TabIndex        =   46
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
         TabIndex        =   45
         Top             =   1440
         Width           =   1455
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
         Left            =   120
         TabIndex        =   44
         Top             =   4800
         Width           =   1335
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
         Left            =   4440
         TabIndex        =   43
         Top             =   1080
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
         TabIndex        =   42
         Top             =   480
         Width           =   1455
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
         Left            =   120
         TabIndex        =   41
         Top             =   4320
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
         Left            =   120
         TabIndex        =   40
         Top             =   3840
         Width           =   1335
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
         Left            =   120
         TabIndex        =   39
         Top             =   5280
         Width           =   1335
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
         Height          =   375
         Index           =   3
         Left            =   7200
         TabIndex        =   38
         Top             =   1680
         Width           =   1935
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
         Height          =   375
         Index           =   2
         Left            =   4320
         TabIndex        =   37
         Top             =   1680
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
         Index           =   4
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   1455
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
         TabIndex        =   35
         Top             =   1920
         Width           =   1335
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
         Left            =   4440
         TabIndex        =   34
         Top             =   600
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
         Left            =   4680
         TabIndex        =   33
         Top             =   4080
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
         Left            =   7200
         TabIndex        =   32
         Top             =   4080
         Width           =   2055
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
         Left            =   120
         TabIndex        =   31
         Top             =   5760
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmDeleteME_BOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/**********************************************************************************
'**文 件 名: FrmDeleteME_Bom
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人:
'**日    期:
'**描    述:
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'** Kevin       2008.09.17     DeleteMe_Bom which are deleted did not need select groupid (0001)
'***********************************************************************************/

Dim strSQL As String


Private Sub CboGroupID_Click()
Call GetGroupWO(CboGroupID)
End Sub

Private Sub CboGroupID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboGroupID_Click
End If
End Sub
Private Sub CboJobPN_Click()
Call GetGroupID(CboJobPN)
Call GetJobGroupByJobRev("", Trim(TxtMBPN), Trim(TxtRev))
End Sub

Private Sub CboJobPN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call GetGroupID(CboJobPN)
End If
End Sub

Private Sub CboLine_Click()

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

    If OptRelease.Value = True Then
       str = "select distinct c.jobpn,b.Mb_Rev from QSMS_WOGroup a ,sap_wo_list b,qsms_JobBOM c  where " & _
       "a.WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and a.line='" & CboLine & "'" & _
       " and a.work_Order=b.wo and a.work_order=c.work_order order by c.jobpn,b.mb_rev"
    Else
        str = "select distinct c.jobpn,b.Mb_Rev  from QSMS_WOGroup a ,sap_wo_list b,qsms_JobBOM c where" & _
        " a.WO_TransDateTime between '" & BeginDate & "' and '" & EndDate & "' and a.line='" & CboLine & "'" & _
        " and a.work_Order=b.wo and a.work_order=c.work_order  order by c.jobpn,b.mb_rev"
    End If


Set Rs = Conn.Execute(str)
CboJobPN.Clear
While Not Rs.EOF
     CboJobPN.AddItem Trim(Rs!Jobpn) & "-" & Trim(Rs!Mb_Rev)
     Rs.MoveNext
Wend
'for DeleteMe_Bom which are deleted did not need select groupid   by kevin 20080917 (0001)
If StrBU = "NB5" Then
'    str = "Select distinct machine from qsms_mebom where machine like '" & Me.CboLine & "%'"
    str = "Select distinct machine from qsms_mebom where line like '" & Me.CboLine & "%'"  ''(1035)
    Set Rs = Conn.Execute(str)
    Me.CboMachine.Clear
    CboMachine.AddItem "ALL"
    While Not Rs.EOF
        Me.CboMachine.AddItem Trim(Rs!Machine)
        Rs.MoveNext
    Wend
End If
End Sub

Private Sub CboLine_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call CboLine_Click
End If
End Sub

Private Sub CboMachine_Click()
Call GetComp(Trim(CboMachine), Trim(TxtMBPN), Trim(TxtWO), Trim(CboLine))
Call GetJobGroupByJobRev(Trim(CboMachine), Trim(TxtMBPN), Trim(TxtRev))
'DeleteMe_Bom which are deleted did not need select groupid for NB5 by kevin 20080917 (0001)
If StrBU = "NB5" And CboMachine <> "all" Then
    Call GetSlot(Trim(CboMachine))
End If
End Sub

Private Sub CboMachine_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboMachine_Click
End If
End Sub

Private Sub CboNotChkBOM_Click()
TxtWO = Trim(CboNotChkBOM)
Call GetWoinfo(TxtWO)
Call GetMachine(TxtWO)
End Sub

Private Sub CboNotChkBOM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboNotChkBOM_Click
End If
End Sub

Private Sub CboWo_Click()
TxtWO = Trim(CboWo)
Call GetWoinfo(TxtWO)
Call GetMachine(TxtWO)
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
    ListWoSelecting.AddItem Trim(ListWoall.text)
    ListWoall.RemoveItem Pointer
    If ListWoall.ListCount <> Pointer Then
       ListWoall.ListIndex = Pointer
    End If
   
End Sub

Private Sub cmdADDALL_Click()

    If ListWoall.ListCount <= 0 Then Exit Sub

    Do While ListWoall.ListCount > 0
     
      ListWoall.ListIndex = 0
      ListWoSelecting.AddItem Trim(ListWoall.text)
      ListWoall.RemoveItem 0
     
    Loop
   
End Sub

Private Sub cmdDel_Click()
    Dim Pointer As Long
    If ListWoSelecting.ListCount <= 0 Then Exit Sub
    If ListWoSelecting.ListIndex < 0 Then Exit Sub
    Pointer = ListWoSelecting.ListIndex

    ListWoall.AddItem Trim(ListWoSelecting.text)
    ListWoSelecting.RemoveItem Pointer
    If ListWoSelecting.ListCount <> Pointer Then
       ListWoSelecting.ListIndex = Pointer
    End If

   
    
End Sub

Private Sub cmdDELALL_Click()
    If ListWoSelecting.ListCount <= 0 Then Exit Sub
    Do While ListWoSelecting.ListCount > 0
        ListWoSelecting.ListIndex = 0
       
        ListWoall.AddItem Trim(ListWoSelecting.text)
        ListWoSelecting.RemoveItem 0
  
    Loop
    
End Sub
Private Sub cmdADDGroup_Click()
    Dim Pointer As Integer
    If ListAllJobGroup.ListCount <= 0 Then Exit Sub
    If ListAllJobGroup.ListIndex < 0 Then Exit Sub
    Pointer = ListAllJobGroup.ListIndex
    ListselectingJobGroup.AddItem Trim(ListAllJobGroup.text)
    ListAllJobGroup.RemoveItem Pointer
    If ListAllJobGroup.ListCount <> Pointer Then
       ListAllJobGroup.ListIndex = Pointer
    End If
   
End Sub
Private Sub cmdADDALLGroup_Click()
    If ListAllJobGroup.ListCount <= 0 Then Exit Sub
    Do While ListAllJobGroup.ListCount > 0
     
      ListAllJobGroup.ListIndex = 0
      ListselectingJobGroup.AddItem Trim(ListAllJobGroup.text)
      ListAllJobGroup.RemoveItem 0
     
    Loop
   
End Sub

Private Sub cmdDelete_Click()
Dim str As String
Dim strSQL As String
Dim Rs As New ADODB.Recordset
Dim Machine As String
Dim jobgroup As String

'DeleteMe_Bom which are deleted did not need select groupid for NB5 by kevin 20080917 (0001)
Rs.CursorLocation = adUseClient
jobgroup = Trim(Me.TxtJobGroup.text)
If Me.TxtWO = "" And (StrBU = "NB5" Or StrBU = "NB3") Then  '1253
'    If Me.CboMachine = "All" Then
'        machine = ""
'    Else
'        machine = Trim(Me.CboMachine.Text)
'    End If
'
'    If CheckData = False Then
'        MsgBox "Please input the all need data", vbCritical
'        Exit Sub
'    End If
'    str = "select count(*) as num FROM QSMS_MEBOM where (jobpn ='" & TxtMBPN & "' or jobpn in (select distinct jobpn from qsms_jobbom where mbpn='" & TxtMBPN & "') or jobPN like '" & Trim(Me.CboJobPN.Text) & "%') and  JobGroup like '" & jobgroup & "%'  and version='" & Trim(Me.TxtRev.Text) & "' and Machine like '" & machine & "' and side like '" & Trim(Me.txtside.Text) & "%' and comppn like '" & Trim(Me.CboComp.Text) & "%' and slot like '" & Trim(Me.cboslot.Text) & "%' and line like '" & Me.CboLine.Text & "%'" ''(1035)
'    Set rs = Conn.Execute(str)
'    If rs.Fields("num") <= 0 Then
'        MsgBox "The data which you want to delete is not exist !", vbCritical
'        Exit Sub
'    Else
'        '删除数据
'        str = "delete  FROM QSMS_MEBOM where (jobpn ='" & TxtMBPN & "' or jobpn in (select distinct jobpn from qsms_jobbom where mbpn='" & TxtMBPN & "') or jobPN like '" & Trim(Me.CboJobPN.Text) & "%') and  JobGroup like '" & jobgroup & "%'  and version='" & Trim(Me.TxtRev.Text) & "' and Machine like '" & machine & "' and side like '" & Trim(Me.txtside.Text) & "%' and comppn like '" & Trim(Me.CboComp.Text) & "%' and slot like '" & Trim(Me.cboslot.Text) & "%' and line like '" & Me.CboLine.Text & "%'" '(1035)    (1078)
'        Conn.Execute (str)
'        '插入日志
'        strsql = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Delete ME_BOM','" & Trim(Me.CboLine) & "+" & Trim(Me.CboJobPN.Text) & "+" & Trim(Me.TxtJobGroup.Text) & "+" & Trim(Me.TxtRev.Text) & "+" & machine & "+" & txtside & "+" & Me.CboComp.Text & "+" & Me.cboslot.Text & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
'        Conn.Execute (strsql)
'        MsgBox "Delete ME bom OK", vbInformation
'    End If
     If Me.TxtJobGroup <> "" Then     ''1198
        str = "delete  FROM QSMS_MEBOM where JobGroup ='" & Trim(jobgroup) & "'"
        str = Replace(str, Chr(13) + Chr(10), "")
        Conn.Execute (str)
        '插入日志
        strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Delete ME_BOM','" & Trim(Me.CboLine) & "+" & Trim(Me.CboJobPN.text) & "+" & Trim(Me.TxtJobGroup.text) & "+" & Trim(Me.TxtRev.text) & "+" & Machine & "+" & txtside & "+" & Me.CboComp.text & "+" & Me.cboslot.text & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
        Conn.Execute (strSQL)
        MsgBox "Delete ME bom OK", vbInformation
     End If
Else
    Call DeleteME_BOM(Trim(TxtMBPN), Trim(TxtRev), Trim(CboMachine.text))
End If
End Sub
Function CheckData() As Boolean
    If Me.CboLine = "" Then
        CheckData = False
        Exit Function
    ElseIf Me.CboJobPN = "" Then
        CheckData = False
        Exit Function
    ElseIf Me.TxtRev = "" Then
        CheckData = False
        Exit Function
    ElseIf Me.TxtJobGroup = "" Then
        CheckData = False
        Exit Function
    Else
        CheckData = True
    End If
End Function

Private Sub CmdDeleteByLine_Click() '1131
Dim str As String
Dim strSQL As String
Dim Rs As New ADODB.Recordset
    If MsgBox("Are you sure to delete this ME_BOM by line " & CboLine.text & " ?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    If CboJobPN.text = "" Then
        If StrBU = "NB3" Then  '1263
            str = "delete  FROM QSMS_MEBOM where line='" & CboLine.text & "' AND Machine not LIKE '%Others%'" '1264
        Else
            str = "delete  FROM QSMS_MEBOM where line='" & CboLine.text & "'"
        End If
    Else
        If StrBU = "NB3" Then  '1263
            str = "delete  FROM QSMS_MEBOM where line='" & CboLine.text & "' and JobPN='" & CboJobPN.text & "' AND Machine not LIKE '%Others%'"  ''1136  1264
        Else
            str = "delete  FROM QSMS_MEBOM where line='" & CboLine.text & "' and JobPN='" & CboJobPN.text & "'"
        End If
    End If
    Conn.Execute (str)
    '插入日志
    strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Delete ME_BOM By Line','" & Trim(Me.CboLine) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
    Conn.Execute (strSQL)
    MsgBox "Delete ME bom OK", vbInformation
End Sub

Private Sub cmdDELGroup_Click()
    Dim Pointer As Long
    If ListselectingJobGroup.ListCount <= 0 Then Exit Sub
    If ListselectingJobGroup.ListIndex < 0 Then Exit Sub
    Pointer = ListselectingJobGroup.ListIndex

    ListAllJobGroup.AddItem Trim(ListselectingJobGroup.text)
    ListselectingJobGroup.RemoveItem Pointer
    If ListselectingJobGroup.ListCount <> Pointer Then
       ListselectingJobGroup.ListIndex = Pointer
    End If
End Sub

Private Sub cmdDELALLGroup_Click()
    If ListselectingJobGroup.ListCount <= 0 Then Exit Sub
    Do While ListselectingJobGroup.ListCount > 0
        ListselectingJobGroup.ListIndex = 0
       
        ListAllJobGroup.AddItem Trim(ListselectingJobGroup.text)
        ListselectingJobGroup.RemoveItem 0
  
    Loop
End Sub

Private Sub CmdQuery_Click()
If Trim(CboLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If
Call GetGroupID("")
Call GetJobPN
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

Me.CmdDeleteByLine.Visible = False  '1131

If DeleteMeBomByLine = True Then    '1131
    Me.CmdDeleteByLine.Visible = True
End If

'DeleteMe_Bom which are deleted did not need select groupid for NB5 by kevin 20080917 (0001)
If StrBU = "NB5" Then
    Me.lblslot.Visible = True
    Me.cboslot.Visible = True
    Me.lblside.Visible = True
    Me.txtside.Visible = True
End If
Call GetLine
End Sub
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

Private Function GetGroupID(ByVal Jobpn As String)
Dim str As String
Dim BeginDate, EndDate As String
Dim GroupIDHead As String
Dim i As Long
Dim Rs As ADODB.Recordset
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
            str = "select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )"
        End If
   Else
        If OptRelease.Value = True Then
           str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N'"
        Else
            str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N' "  ''(1035)
        End If
   End If
Else
    If InStr(1, Jobpn, "-") > 0 Then
       
       '--superchai comments (Begin) QMB0002 20230516--
       'TempJobPn = Mid(Jobpn, 1, 11)
       'TxtMBPN = TempJobPn
       'TxtRev.text = Right(Jobpn, 3)
       '--superchai comments (End) QMB0002 20230516--
       
       '--superchai add function from QSMC (Begin) QMB0002 20230516--
       '20230508 Kaelyn forPN э20絏
       'TempJobPn = Mid(Jobpn, 1, 11)
       str = "SELECT CompPN FROM QSMS_DID with(nolock) where DID='" & Jobpn & "'"
       Set Rs = Conn.Execute(str)
       TempJobPn = CStr(Rs!COMPPN)
       
       TxtMBPN = TempJobPn
       TxtRev.text = Right(Jobpn, 3)
       '--superchai add function from QSMC (End) QMB0002 20230516--
    Else
      TempJobPn = Jobpn
    End If
    If OptRelease.Value = True Then
       str = "select distinct GroupID from QSMS_WOGroup a,QSMS_JobBOM b   where a.WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' " & _
             "and a.line='" & CboLine & "' and a.work_order=b.work_order and b.jobpn='" & TempJobPn & "'and closedflag='N' "
    Else ''(1035)
        str = "select distinct GroupID from QSMS_WOGroup a,QSMS_JobBOM b  where a.WO_TransDateTime between '" & BeginDate & "' and '" & EndDate & "' " & _
              "and a.line='" & CboLine & "' and a.work_order=b.work_order and b.jobpn='" & TempJobPn & "' and closedflag='N'"
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
Private Function GetJobPN()
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

If OptRelease.Value = True Then
   str = "select distinct JobPN from QSMS_WOGroup a,QSMS_JobBOM b   where a.WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' " & _
         "and a.line='" & CboLine & "' and A.work_order=b.work_order"
Else
    str = "select distinct JobPN from QSMS_WOGroup a,QSMS_JobBOM b   where a.WO_TransDateTime between '" & BeginDate & "' and '" & EndDate & "' " & _
          "and a.line='" & CboLine & "' and a.work_Order=b.work_Order"
End If

Set Rs = Conn.Execute(str)
CboJobPN.Clear
If Rs.EOF Then MsgBox "No data"
While Not Rs.EOF
      CboJobPN.AddItem Trim(Rs!Jobpn)
      Rs.MoveNext
Wend
End Function

Private Function GetGroupWO(ByVal GroupID As String)
Dim str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset
Dim TempJobPn As String
TempJobPn = ""
If Trim(CboJobPN) <> "" Then
   TempJobPn = Mid(CboJobPN, 1, InStr(1, CboJobPN, "-") - 1)
End If
str = "select distinct a.Work_Order from QSMS_WOGroup a,QSMS_JobBOM b  where a.GroupID= '" & GroupID & "' and a.work_Order=b.work_order and b.jobpn like '" & TempJobPn & "%'"

Set Rs = Conn.Execute(str)
ListWoall.Clear
CboWo.Clear
CboNotChkBOM.Clear
While Not Rs.EOF
          If ChkQSMS_WO(Trim(Rs!Work_Order)) = False Then
             CboNotChkBOM.AddItem Trim(Rs!Work_Order)
          Else
             ListWoall.AddItem Trim(Rs!Work_Order)
             CboWo.AddItem Trim(Rs!Work_Order)
          End If
          Rs.MoveNext
Wend
End Function


Private Function GetWoinfo(ByVal WO As String)
Dim str As String
Dim Rs As ADODB.Recordset
str = "select PN, Qty ,MB_Rev,Line,BuildType from Sap_Wo_List where WO='" & Trim(WO) & "'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   TxtMBPN = Rs!PN
   TxtWOQty = Rs!Qty
   TxtRev = Trim(Rs!Mb_Rev)
   CboLine.text = Trim(Rs!Line)
End If
str = "select Customer from ModelName where PN='" & TxtMBPN & "'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   TxtCustomer = Trim(Rs!Customer)
End If
CboJobPN.Clear
str = "select jobPn from QSMS_JobBOM where work_Order='" & WO & "'"
Set Rs = Conn.Execute(str)
While Not Rs.EOF
     CboJobPN.AddItem Trim(Rs!Jobpn)
     Rs.MoveNext
Wend
End Function


Private Function GetMachine(ByVal WO As String)
Dim str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset
Dim woGroup As String
str = "select [group] from sap_wo_list where wo='" & WO & "'"
Set Rs = Conn.Execute(str)
If Rs.EOF Then
    MsgBox "This WO is not exist! Please check!"
    Exit Function
End If
woGroup = Trim(Rs("Group"))

'Str = "select distinct Machine,MachinefinishedFlag from QSMS_WO where Work_Order= '" & WO & "' "
''(1035) a.machine like c.line+'%' -->a.line=c.line
str = "Select distinct machine from qsms_mebom a,qsms_jobbom b,sap_wo_list c where b.work_order='" & WO & "' and " & _
      "b.work_order=c.wo and a.line=c.line and c.mb_rev=a.version and b.jobpn=a.jobpn " & _
      "and jobgroup in " & GetJobGroup(woGroup)
Set Rs = Conn.Execute(str)
CboMachine.Clear
CboMachine.AddItem "ALL"
While Not Rs.EOF
    
        CboMachine.AddItem Trim(Rs!Machine)
    
     Rs.MoveNext
Wend

End Function
Private Function GetSlot(ByVal Machine As String)
Dim str As String
Dim Rs As New ADODB.Recordset
    
str = "select distinct slot from qsms_mebom where machine='" & Machine & "'  and line = '" & CboLine.text & "'order by slot"  ''(1035)
Set Rs = Conn.Execute(str)
Me.cboslot.Clear
While Not Rs.EOF
    Me.cboslot.AddItem Trim(Rs!Slot)
    Rs.MoveNext
Wend
End Function

Private Function GetComp(ByVal Machine As String, ByVal MBPN As String, ByVal WO As String, ByVal Line As String)


Dim str As String
Dim Rs As ADODB.Recordset

If Machine <> "ALL" Then
    str = "select CompPN from QSMS_WO  where Work_Order like '" & Trim(WO) & "%'  and Machine='" & Trim(Machine) & "'"
Else
    str = "select CompPN from QSMS_WO  where Work_Order like '" & Trim(WO) & "%'"
End If
Set Rs = Conn.Execute(str)
CboComp.Clear
While Not Rs.EOF
  
       CboComp.AddItem Trim(Rs!COMPPN)
       

    Rs.MoveNext
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

Private Function DeleteME_BOM(ByVal MBPN As String, ByVal Rev As String, Machine As String)
Dim str As String
Dim Rs As ADODB.Recordset
Dim jobgroup As String
Dim Line As String
Dim Rst As ADODB.Recordset
If Trim(MBPN) = "" Then 'Or Trim(Rev) = "" Then
   MsgBox "Please input MBPN"
   Exit Function
End If
If Trim(Machine) = "" Then
   MsgBox "Please input Machine (set Machine=All to delete all)"
   Exit Function
End If

jobgroup = GetSelectingJobGroup()
If jobgroup = "" Then
    If StrBU = "PO" Then
        MsgBox "叫匡拒JobGroup"
    Else
        MsgBox "请选择 JobGroup"
    End If
   Exit Function
End If


If UCase(Machine) = "ALL" Then
    'If MsgBox(IIf(StrBU = "PO", "琌虫絬埃?", "是否按工单线别删除？"), vbYesNo, "Message") = vbYes Then ''' add Dailog form to choose by line delete --Steven   (1036)
    '   machine = CboLine & "%"     (1036)
    ' Else  (1036)
       Machine = "%"
     'End If  (1036)
Else
    Machine = Machine & "%"
End If

'If JobGroup = "All" Then
'   JobGroup = "%"
'Else
'   JobGroup = JobGroup & "%"
'End If

str = "select *  FROM QSMS_MEBOM where (jobpn ='" & MBPN & "' or jobpn in (select jobpn from qsms_jobbom where mbpn='" & MBPN & "'))  and JobGroup in  " & jobgroup & " and version='" & Rev & "' and Machine like '" & Machine & "' and line like '" & Me.CboLine.text & "%'"  ''(1035)
Set Rs = Conn.Execute(str)
If Rs.EOF Then
  MsgBox "can not find ME BOM ,Please check the MBPN or Rev"
  Exit Function
End If

str = "delete  FROM QSMS_MEBOM where (jobpn ='" & MBPN & "' or jobpn in (select jobpn from qsms_jobbom where mbpn='" & MBPN & "')) and JobGroup in " & jobgroup & "  and version='" & Rev & "' and Machine like '" & Machine & "' and line like '" & Me.CboLine.text & "%' "  ''(1035)
Conn.Execute str
'''''''''''''''1073
'str = "select * from WO_MultiLine where WO='" & Me.TxtWO.Text & "'"
''1164
str = "select a.Line from WO_MultiLine a, Sap_Wo_List b where a.WO=b.WO and B.[Group] in(select [Group] from Sap_Wo_List where BuildType='4' and WO='" & Trim(TxtWO) & "')"
Set Rst = Conn.Execute(str)
If Not Rst.EOF Then
    Line = Rst("Line")
    str = "delete  FROM QSMS_MEBOM where (jobpn ='" & MBPN & "' or jobpn in (select jobpn from qsms_jobbom where mbpn='" & MBPN & "')) and JobGroup in " & jobgroup & "  and version='" & Rev & "' and Machine like '" & Machine & "' and line like '" & Line & "%' "  '()
    Conn.Execute str
End If
''''''''''''''1073
'Record who delete which MEBOM
strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Delete ME_BOM','" & Trim(MBPN) & "+" & Trim(Rev) & "+" & Machine & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"

Conn.Execute (strSQL)

MsgBox "Delete ME bom OK"

End Function

Private Function GetWO(ByVal Ctype As String, cOutPut As String) As String
Dim str As String, i As Integer
Dim Rs As ADODB.Recordset
Dim WO As String
Dim WoOutPut As String
WO = ""
WoOutPut = ""



Select Case UCase(Ctype)
       Case "BY_WORKORDERS"
            For i = 0 To ListWoSelecting.ListCount - 1
                              ListWoSelecting.ListIndex = i
                              WO = WO & Trim(ListWoSelecting.text) & ","
                              WoOutPut = WoOutPut & ListWoSelecting.text & Chr(10) & Chr(13)
            Next i
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
            Set Rs = Conn.Execute(str)
            While Not Rs.EOF
                 WO = WO & Trim(Rs!Work_Order) & ","
                 WoOutPut = WoOutPut & Trim(Rs!Work_Order) & Chr(10) & Chr(13)
                 Rs.MoveNext
            Wend
             WO = Mid(WO, 1, Len(WO) - 1)
       Case "BY_WORKORDER"
              WO = Trim(TxtWO)
              WoOutPut = WO
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
Dim Rs As ADODB.Recordset
If Len(Machine) > 6 Then ''(1035)
    str = "Select distinct jobgroup from qsms_mebom where (jobpn='" & Jobpn & "' or jobpn in (select distinct jobpn from QSMS_JObBom where MBPN='" & Jobpn & "')) and version='" & Version & "'" & _
          " and line ='" & CboLine.text & "' and Machine like '" & Machine & "%' " & " order by jobgroup"
Else
    str = "Select distinct jobgroup from qsms_mebom where (jobpn='" & Jobpn & "' or jobpn in (select distinct jobpn from QSMS_JObBom where MBPN='" & Jobpn & "')) and version='" & Version & "' and line = '" & CboLine.text & "'" & " order by jobgroup"
End If
Set Rs = Conn.Execute(str)
ListAllJobGroup.Clear
ListselectingJobGroup.Clear
While Not Rs.EOF
     ListAllJobGroup.AddItem Trim(Rs!jobgroup)
     Rs.MoveNext
     
Wend

End Function
Private Function GetSelectingJobGroup() As String
Dim SelectingJobGroup As String, i As Integer

    If ListselectingJobGroup.ListCount <= 0 Then
           If Len(Trim(TxtJobGroup)) > 0 And InStr(1, Trim(TxtJobGroup), "-") > 0 Then
               GetSelectingJobGroup = "('" & Trim(TxtJobGroup) & "')"
           Else
               SelectingJobGroup = ""
           End If
           Exit Function
    End If
    
    For i = 1 To ListselectingJobGroup.ListCount
        ListselectingJobGroup.ListIndex = i - 1
        SelectingJobGroup = SelectingJobGroup + "'" + Trim(ListselectingJobGroup.text) + "'" + ","
    
    Next i
    
    GetSelectingJobGroup = "(" + Mid(SelectingJobGroup, 1, Len(SelectingJobGroup) - 1) + ")"
End Function

Private Sub TxtMBPN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   TxtRev.Enabled = True
   TxtRev.SetFocus
End If
End Sub

Private Sub TxtRev_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call GetJobGroupByJobRev(Trim(CboMachine), Trim(TxtMBPN), Trim(TxtRev))
End If
End Sub
