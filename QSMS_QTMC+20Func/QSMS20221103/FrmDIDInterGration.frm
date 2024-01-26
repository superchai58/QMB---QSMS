VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmDIDInteGration 
   Caption         =   "DIDIntegration"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   14985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Comp Port data maintain"
      Height          =   1455
      Left            =   120
      TabIndex        =   35
      Top             =   0
      Width           =   9495
      Begin VB.Frame Frame2 
         Caption         =   "Printer"
         Height          =   615
         Left            =   1440
         TabIndex        =   45
         Top             =   720
         Width           =   4575
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
            Left            =   1320
            TabIndex        =   48
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
            Left            =   3000
            TabIndex        =   47
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox chkDelay 
            Caption         =   "Delay"
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
            Left            =   240
            TabIndex        =   46
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin VB.OptionButton OptPrint 
         BackColor       =   &H80000013&
         Caption         =   "Print Port"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   975
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
         TabIndex        =   43
         Text            =   "1"
         Top             =   240
         Width           =   495
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
         Left            =   4680
         TabIndex        =   42
         Text            =   "9600,N,8,1"
         Top             =   240
         Width           =   1455
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
         Height          =   375
         Left            =   6360
         Picture         =   "FrmDIDInterGration.frx":0000
         TabIndex        =   41
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptComp 
         BackColor       =   &H80000013&
         Caption         =   "Comp Port"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Caption         =   "Label"
         Height          =   615
         Left            =   6120
         TabIndex        =   37
         Top             =   720
         Width           =   1815
         Begin VB.OptionButton opOldLabel 
            Caption         =   "old"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton opNewLabel 
            Caption         =   "new"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   38
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.OptionButton OptNetwork 
         BackColor       =   &H80000013&
         Caption         =   "NetWork"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
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
         TabIndex        =   50
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
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
         TabIndex        =   49
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame FraConnection 
      BackColor       =   &H80000013&
      Caption         =   "DIDInfo"
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   14415
      Begin VB.ComboBox CmbFactory 
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
         Left            =   13200
         TabIndex        =   56
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtEndTime 
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
         Left            =   8640
         TabIndex        =   54
         Text            =   "0800"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtBeginTime 
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
         Left            =   4200
         TabIndex        =   51
         Text            =   "0500"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtSide 
         Enabled         =   0   'False
         Height          =   375
         Left            =   7080
         TabIndex        =   32
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtLine 
         Enabled         =   0   'False
         Height          =   375
         Left            =   5160
         TabIndex        =   31
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton CmdReprint 
         Caption         =   "Reprint"
         DragIcon        =   "FrmDIDInterGration.frx":030A
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
         Left            =   3960
         Picture         =   "FrmDIDInterGration.frx":5F1C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2160
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
         Left            =   1080
         Picture         =   "FrmDIDInterGration.frx":BB2E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2160
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
         Left            =   1800
         Picture         =   "FrmDIDInterGration.frx":BE38
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2160
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
         Left            =   360
         Picture         =   "FrmDIDInterGration.frx":C27A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2160
         Width           =   735
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
         Height          =   855
         Left            =   3240
         Picture         =   "FrmDIDInterGration.frx":C6BC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2160
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
         Left            =   4680
         Picture         =   "FrmDIDInterGration.frx":C9C6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2160
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
         Left            =   2520
         Picture         =   "FrmDIDInterGration.frx":CCD0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2160
         Width           =   735
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
         Left            =   1680
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox CboVendorCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   6120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox CboDateCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   10320
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox CboLotCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   1680
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox TxtQty 
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
         Left            =   8880
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox CboLine 
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
         Left            =   9840
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox CboSide 
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
         ItemData        =   "FrmDIDInterGration.frx":D6D2
         Left            =   11520
         List            =   "FrmDIDInterGration.frx":D6E2
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox CboDID 
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
         TabStop         =   0   'False
         Top             =   1680
         Width           =   4695
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   240
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
         Format          =   134348803
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   5880
         TabIndex        =   18
         Top             =   240
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
      Begin VB.Label LblFactory 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Factory"
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
         Left            =   12120
         TabIndex        =   55
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "EndTime"
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
         Left            =   7680
         TabIndex        =   53
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "BeginTime"
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
         Left            =   3000
         TabIndex        =   52
         Top             =   240
         Width           =   1215
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
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   1095
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
         Left            =   8760
         TabIndex        =   29
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H0000FF00&
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
         Left            =   120
         TabIndex        =   28
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
         Left            =   4440
         TabIndex        =   27
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
         Left            =   120
         TabIndex        =   26
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
         Left            =   8040
         TabIndex        =   25
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Line"
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
         Left            =   9240
         TabIndex        =   24
         Top             =   240
         Width           =   615
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
         Index           =   3
         Left            =   4800
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Side"
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
         Index           =   5
         Left            =   10800
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Side"
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
         Index           =   6
         Left            =   6240
         TabIndex        =   20
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Line"
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
         Index           =   7
         Left            =   4440
         TabIndex        =   19
         Top             =   1200
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid DG1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   6165
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
   Begin MSFlexGridLib.MSFlexGrid flexGridDemandMaterial 
      Height          =   1095
      Left            =   120
      TabIndex        =   33
      Top             =   5280
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   1931
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   10680
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label LblStatus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Top             =   4560
      Width           =   14415
   End
End
Attribute VB_Name = "FrmDIDInteGration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strDID As String
Private DID() As String
Dim PrintData As PtData
Dim CHKAutoDispatchForAnotherBU As Boolean
Dim PrinterPort As String
Dim PrinterType As String
Dim PrinterVendor As String
Dim PrinterSetting As String
Dim RSDG1 As New ADODB.Recordset
Dim DIDCount As Integer
Dim NoEntireCompPNQty  As Integer ''''散料数量
Dim isZebra As Boolean ''(1080)

Private Sub CboDID_KeyPress(KeyAscii As Integer)
Dim strSQL As String, Rs As New ADODB.Recordset
If KeyAscii = 13 And Trim(CboDID.text) <> "" Then
    If strDID = "" Then
        LblStatus.Caption = ""
        strSQL = "select * from QSMS_DID where DID = '" & Trim(CboDID.text) & "'"
        Set Rs = Conn.Execute(strSQL)
        If Rs.EOF = False Then ''''
            CboCompPN.text = Trim(Rs("CompPN"))
            CboVendorCode.text = Trim(Rs("VendorCode"))
            CboDateCode.text = Trim(Rs("DateCode"))
            CboLotCode.text = Trim(Rs("LotCode"))
            txtLine.text = Trim(Rs("Line"))
            txtSide.text = Trim(Rs("Side"))
            Set Rs = Nothing
            strSQL = "select Qty  from QSMS_NoEntireCompPNSetting  where   " & sq(Trim(CboCompPN.text)) & " like   PrefixPN+'%'"
            Set Rs = Conn.Execute(strSQL)
            If Rs.EOF = False Then ''''
                NoEntireCompPNQty = Val(Rs("Qty"))
            Else
                MsgBox "Please Maintain Data in QSMS_NoEntireCompPNSetting(Mainmenu-->UpLoadBasicData-->QSMS_NoEntireCompPNSetting)!"
                Call ClearData
                Exit Sub
            End If
        End If
    End If
    strSQL = "select * from QSMS_DID WITH(NOLOCK) where DID = " & sq(CboDID.text)
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = True Then
        MsgBox "This DID is not exist , and Please MaintainDID at first !"
        Call ClearData
        Exit Sub
    Else
        If CheckValid(UCase(Trim(CboDID.text))) = False Then
            Call ClearData
            Exit Sub
        End If
        TxtQty.text = str(Val(TxtQty.text) + Val(Rs("Qty")))
    End If
    strDID = strDID & UCase(Trim(CboDID.text)) & ";"
    CboDID.text = ""
    LblStatus.Caption = strDID
    CboDID.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
Call ClearData
End Sub

 

Private Sub CmdExcel_Click()
If RSDG1.EOF = False Then
    Call CopyToExcel(RSDG1)
Else
    MsgBox "No Data!"
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Dim strSQL As String, Rs As New ADODB.Recordset
Dim sDate As String, eDate As String
Set DG1.DataSource = Nothing
sDate = Format(dtpSDate, "YYYYMMDD")
eDate = Format(dtpEDate, "YYYYMMDD")
If Len(Trim(txtBeginTime.text)) <> 4 And IsNumeric(Trim(txtBeginTime.text)) = True Then
    MsgBox "请输入4位有效时间(HHMM)"
    Exit Sub
End If
If Len(Trim(txtEndTime.text)) <> 4 And IsNumeric(Trim(txtEndTime.text)) = True Then
    MsgBox "请输入4位有效时间(HHMM)"
    Exit Sub
End If
If CmbFactory.Visible = True Then
    If Trim(CmbFactory.text) = "" Then
        MsgBox "请选择Factory"
        Exit Sub
    End If
End If
strSQL = "QSMS_DIDIntegration @Item =" & sq("QUERY") & ", @CompPN  =" & sq(CboCompPN.text) & ", @Line = " & sq(CboLine.text) & ", @Side  = " & sq(CboSide.text) & ", @BeginTime =" & sq(sDate & Trim(txtBeginTime.text)) & ", @EndTime =" & sq(eDate & Trim(txtEndTime.text)) & ",@Factory=" & sq(Trim(CmbFactory.text))
Set RSDG1 = Conn.Execute(strSQL)
If RSDG1.EOF = False Then
    Set DG1.DataSource = RSDG1
    DG1.Refresh
Else
    MsgBox "No Data!"
End If
End Sub

 

 

Private Sub CmdRefresh_Click()
Call cmdFind_Click
End Sub

Private Sub cmdReprint_Click()
Dim str As String
Dim Rs As ADODB.Recordset

On Error GoTo errhandle
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
    
    str = "Select * from QSMS_DID Where DID='" & Trim(CboDID) & "'"
    Set Rs = Conn.Execute(str)
    If Rs.EOF Then
       MsgBox "can not find the DID,Please check"
       CboDID.SetFocus
       Exit Sub
    ElseIf UCase(Trim(Rs!FirstMachine)) = "RETURN" Or UCase(Trim(Rs!FirstMachine)) = "CALLBACK" Then
       MsgBox "Can not do reprint, this DID has been " + Trim(Rs!FirstMachine) + "ed !"
       CboDID.SetFocus
       Exit Sub
    Else
        PrintData.Line = Trim(Rs!Line)
        PrintData.Machine = Trim(Rs!FirstMachine)
        PrintData.Side = Trim(Rs!Side)
        PrintData.DIDWOGROUP = Trim(Rs!woGroup)
        TxtQty.text = Trim(Rs!Qty)
    End If

    str = "insert into QSMS_Log(system_name,event_no,did,[user_name],returnQty,Trans_Date) values('MaintainDID_Reprint','1','" & Trim(CboDID) & "','" & g_userName & "','0',dbo.FormatDate(getdate(),'yyyymmddhhnnss'))"
    Conn.Execute (str)
    
    Call PrintLabel(Trim(CboDID.text), TxtQty)
    Exit Sub
errhandle:
    MsgBox Err.Description
End Sub
 

Private Sub cmdSave_Click()
On Error GoTo ErrorHandle:
Dim i As Integer, strGroupDIDQty   As Long, TransDate As String, TempDID As String
Dim strSQL As String, Rs As New ADODB.Recordset, strAnotherQSMSIP As String
Dim InsertDIDOk As Boolean
Dim strInputDID() As String
strDID = UCase(strDID)
If Right(Trim(strDID), 1) = ";" Then
    strDID = Left(strDID, Len(strDID) - 1)
End If


If strDID = "" Then
    MsgBox "Please Input DID"
    Call ClearData
    Exit Sub
End If
If Trim(txtLine.text) = "" Then
    MsgBox ("Line can not  null")
    Call ClearData
    Exit Sub
End If
If Trim(txtSide.text) = "" Then
    MsgBox ("Side can not  null")
    Call ClearData
    Exit Sub
End If
'If CheckValid(strDID) = False Then
'    Call ClearData
'    Exit Sub
'End If

strInputDID = Split(strDID, ";")
If UBound(strInputDID) < 1 Then
    MsgBox "必须输入两个及其以上DID"
    Call ClearData
    Exit Sub
End If

LockTheForm (False)
InsertDIDOk = False


'strSql = "select dbo.FormatDate (GETDATE() ,'yyyymddhhnnss')"
strSQL = "select dbo.FormatDate (GETDATE() ,'yyyymmddhhnnss')"  '(1091)
Set Rs = Conn.Execute(strSQL)
TransDate = Trim(Rs(0))



LockTheForm (False)
'If CheckValid(strDID) = False Then
'    Call ClearData
'    LockTheForm (True)
'    Exit Sub
'End If
TempDID = Trim(GetDID(Trim(CboCompPN.text), TransDate))

strSQL = "EXEC QSMS_DIDIntegration @Item='DIDIngration', " & "  @CompPN =" & sq(CboCompPN.text) & ", @NEWDID=" & sq(TempDID) & ", @Qty =" & sq(Trim(TxtQty.text)) & ", @RemainQty=" & sq(Trim(TxtQty.text)) & ",@VendorCode=" & sq(Left(Trim(CboVendorCode), 7)) & ",@DateCode=" & sq(Trim(CboDateCode.text)) & ",@LotCode=" & sq(Trim(CboLotCode.text)) & ",      @Line=" & sq(Trim(txtLine.text)) & ",@Side=" & sq(Trim(txtSide.text)) & ",@ALLDID = " & sq(Trim(strDID)) & ",@Factory=" & sq(Factory) & ",@UID =" & sq(g_userName)
Set Rs = Conn.Execute(strSQL)
If Rs.EOF = False Then
    If Trim(Rs("result")) <> "1" Then
        LblStatus.Caption = Trim(Rs!ErrDesc)
      '  strsql = "EXEC QSMS_IntegrationMaterial @Item='RESTOREDISPATCH', @NEWDID=" & sq(TempDID) & ",@Line=" & sq(Trim(txtLine.Text)) & ",@Side=" & sq(Trim(txtSide.Text)) & ",@ALLDID = " & sq(Trim(strDID)) & ",@Factory=" & sq(Factory) & ",@UID =" & sq(g_userName)
        strSQL = "EXEC QSMS_DIDIntegration @Item='RESTOREDISPATCH', @NEWDID=" & sq(TempDID) & ",@Line=" & sq(Trim(txtLine.text)) & ",@Side=" & sq(Trim(txtSide.text)) & ",@ALLDID = " & sq(Trim(strDID)) & ",@Factory=" & sq(Factory) & ",@UID =" & sq(g_userName)
        Set Rs = Conn.Execute(strSQL)
        If Rs.EOF = False Then
            If Trim(Rs("result")) = "1" Then
                 LockTheForm (True)
                 Call ClearData
                 LblStatus.Caption = LblStatus.Caption & Trim(Rs!ErrDesc)
                 'LblStatus.Caption = " DID  " & TempDID & "  Dispatch  Fail "
                 Exit Sub
            Else
                 LockTheForm (True)
                 Call ClearData
                 LblStatus.Caption = LblStatus.Caption & Trim(Rs!ErrDesc)
                 'LblStatus.Caption = " DID  " & TempDID & "  Dispatch  Fail "
                 Exit Sub
            End If
        Else
             LockTheForm (True)
             Call ClearData
             LblStatus.Caption = LblStatus.Caption & "    " & "Please  Call QMS !"
             'LblStatus.Caption = " DID  " & TempDID & "  Dispatch  Fail "
             Exit Sub
        End If
    Else
        strSQL = "delete A from  QSMS_DID  A , QSMS_DIDIntergration_Log B where A.DID = B.OldDID and B.NewDID = " & sq(TempDID)
        Conn.Execute (strSQL)
        LblStatus.Caption = " DID  " & TempDID & "  Dispatch success !"
    End If
Else
    LockTheForm (True)
    Call ClearData
    LblStatus.Caption = " DID:  " & TempDID & "  Dispatch  Fail ," & "Please  Call QMS !"
    Exit Sub
End If
 
' strsql = "exec XL_DIDAutoDispatch  @DID=" & sq(TempDID) & ", @CompPN =" & sq(CboCompPN.Text) & ", @Qty =" & sq(TxtQty.Text) & ", @RemainQty=" & sq(TxtQty.Text) & _
'        ",@VendorCode=" & sq(Left(Trim(CboVendorCode), 7)) & ",@DateCode=" & sq(Trim(CboDateCode.Text)) & ",@LotCode=" & sq(Trim(CboLotCode.Text)) & _
'        ",@DIDLoc='' , @DIDMEM='',@UID=" & sq(g_userName) & " ,@Type='3', @WOList='',@extraWOGroup ='', @extraWO='', @extraLine = " & sq(txtLine.Text) & _
'        ", @extraSide =" & sq(txtSide.Text) & ",@extraMachine ='',@extraSlot ='', @extraLR ='',@extraQty =0,@Factory = " & sq(Factory) & " ,  @MSD =''"
'Set RS = Conn.Execute(strsql)
'LblStatus = Trim(RS!ErrDesc)
'If Trim(RS!result) <> "1" Then
'    'Call RefreshDg("", BeginDID, EndDID)
'    strsql = "EXEC QSMS_IntegrationMaterial @Item='RESTOREDISPATCH', @NEWDID=" & sq(TempDID) & ",@Line=" & sq(Trim(txtLine.Text)) & ",@Side=" & sq(Trim(txtSide.Text)) & ",@ALLDID = " & sq(Trim(strDID)) & ",@Factory=" & sq(Factory) & ",@UID =" & sq(g_userName)
'    Set RS = Conn.Execute(strsql)
'    If RS.EOF = False Then
'        If Trim(RS("result")) = "0" Then
'             LockTheForm (True)
'             Call ClearData
'             LblStatus.Caption = "Dispatch fail "
'             Exit Sub
'        Else
'             LockTheForm (True)
'             Call ClearData
'             LblStatus.Caption = " RestoreDispatch Data is fail "
'             Exit Sub
'        End If
'    Else
'        LockTheForm (True)
'        Call ClearData
'        LblStatus.Caption = " RestoreDispatch Data is fail "
'        Exit Sub
'    End If
'    'if
'    strStep = "6"
'    LockTheForm (True)
'    Call ClearData
'    strStep = "7"
'    Call cmdCancel_Click
'    cmdExcel.Enabled = False
'    Exit Sub
'End If

strSQL = "exec XL_GetDidPrintInfo  @DID='" & Trim(TempDID) & "',@Factory='" & Trim(Factory) & "'"
Set Rs = Conn.Execute(strSQL)
If Rs.EOF Then
    MsgBox "can not find the DID,Please check"
    LockTheForm (True)
    Call ClearData
   CboDID.SetFocus
   Exit Sub
Else
    
    PrintData.Line = Trim(Rs!Line)
    PrintData.Machine = Trim(Rs!FirstMachine)
    PrintData.Side = Trim(Rs!Side)
    PrintData.DIDWOGROUP = Trim(Rs!woGroup)
    PrintData.UniqueID = Trim(Rs!UniqueID) 'Stephen add Quanta Delivery Label 1286
    WOType = Trim(Rs!WOType)
    
    If CHKAutoDispatchForAnotherBU = True Then
        Set Rs = Rs.NextRecordset
        If Rs.EOF = False Then
            strAnotherQSMSIP = Trim(Rs!AnotherQSMSIP)
        End If
    End If
End If
InsertDIDOk = True

CboDID = TempDID
If PreDIDPrinted <> CboDID And InsertDIDOk = True Then
    Call PrintLabel(TempDID, TxtQty, strAnotherQSMSIP)
End If
PreDIDPrinted = CboDID
If OptZebra.Value = True And (chkDelay.Value = Checked) Then
    Call Sleep(900)
End If
LblStatus.Caption = " DID  " & TempDID & "  Dispatch  and print label success !"
LockTheForm (True)

Call ClearData
Exit Sub
ErrorHandle:
    LockTheForm (True)
    Call ClearData
    MsgBox Err.Description
End Sub
Private Function CheckValid(DID As String) As Boolean
On Error GoTo ErrorHandle:
Dim strSQL As String, Rs As New ADODB.Recordset
Dim i As Integer
Dim strInputDID() As String, strCompPN As String

If Right(Trim(DID), 1) = ";" Then
    DID = Left(DID, Len(DID) - 1)
End If
If InStr(1, strDID, DID) > 0 Then
    LblStatus.Caption = "Please do not input the same DID : " & DID
    CheckValid = False
    Exit Function
End If
strInputDID = Split(DID, ";")
strSQL = "select  0 from XL_ReelBaseQty WITH(NOLOCK) where CompPN =" & sq(CboCompPN.text) & " and BaseReelQty >" & str(NoEntireCompPNQty)
Set Rs = Conn.Execute(strSQL)
If Rs.EOF = True Then
    LblStatus.Caption = CboCompPN.text & " realBaseQty must be greater than the system-defined bulk material quantity:" & str(NoEntireCompPNQty)
    CheckValid = False
    Exit Function
End If
For i = 0 To UBound(strInputDID)
    strSQL = "select *  from QSMS_DID WITH(NOLOCK) where DID=" & sq(strInputDID(i))
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = True Then
        LblStatus.Caption = strInputDID(i) & " does not exist!"
        CheckValid = False
        Exit Function
    Else
        If UCase(Trim((Rs("CompPN")))) <> UCase(Trim(CboCompPN.text)) Then
            LblStatus.Caption = strInputDID(i) & " and the CompPN of the first DID are inconsistent!"
            CheckValid = False
            Exit Function
        End If
        If Val(Rs("Qty")) >= NoEntireCompPNQty And UnChkBaseReelQty <> "Y" Then
            LblStatus.Caption = strInputDID(i) & " Qty > " & str(NoEntireCompPNQty) & "（the system-defined bulk material quantity） !"
            CheckValid = False
            Exit Function
        End If
        If UCase(Trim(Rs("Line"))) <> UCase(Trim(txtLine.text)) Then
            LblStatus.Caption = strInputDID(i) & " and the Line of the first DID are inconsistent!"
            CheckValid = False
            Exit Function
        End If
        If UCase(Trim(Rs("side"))) <> UCase(Trim(txtSide.text)) Then
            LblStatus.Caption = strInputDID(i) & " and the Side of the first DID are inconsistent!"
            CheckValid = False
            Exit Function
        End If
    End If
    
    strSQL = "SELECT  0  FROM QSMS_Verify  WITH(NOLOCK) WHERE  DID = " & sq(strInputDID(i))
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = False Then
        LblStatus.Caption = strInputDID(i) & " DID has been used!"
        CheckValid = False
        Exit Function
    End If
    strSQL = "SELECT  0  FROM QSMS_FeederDID_Current  WITH(NOLOCK)  WHERE  DID = " & sq(strInputDID(i))
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = False Then
        LblStatus.Caption = strInputDID(i) & " DID has been used!"
        CheckValid = False
        Exit Function
    End If
    strSQL = "SELECT  0  FROM QSMS_FeederDID_Buffer WITH(NOLOCK)   WHERE  DID = " & sq(strInputDID(i))
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = False Then
        LblStatus.Caption = strInputDID(i) & " DID has been used!"
        CheckValid = False
        Exit Function
    End If
Next i
CheckValid = True
Exit Function
ErrorHandle:
   MsgBox Err.Description
End Function

 

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
Dim str As String
Dim Rs As ADODB.Recordset
Dim i As Long
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
dtpSDate = Date
dtpEDate = Date
Call Init
Call LockTheForm(True)
Call GetPrinterSetting(FrmDIDInteGration)
''(1080)
If OptZebra.Value = True Then
    isZebra = True
Else
    isZebra = False
End If
End Sub
 
Private Sub Init()
Dim strSQL As String
Dim Rs As ADODB.Recordset
NoEntireCompPNQty = 0
strSQL = "select distinct Line from Machine order by Line asc"
Set Rs = Conn.Execute(strSQL)
CboLine.Clear
While Not Rs.EOF
    CboLine.AddItem Rs!Line
    Rs.MoveNext
Wend
strSQL = "select distinct factory from site"
Set Rs = Conn.Execute(strSQL)
CmbFactory.Clear
If Rs.EOF = False Then
    If Rs.RecordCount > 1 Then
        CmbFactory.Visible = True
        CmbFactory.Enabled = True
        LblFactory.Visible = True
    Else
        CmbFactory.Visible = False
        CmbFactory.Enabled = False
        LblFactory.Visible = False
    End If
    While Not Rs.EOF
        CmbFactory.AddItem Trim(Rs!Factory)
        Rs.MoveNext
    Wend
End If
End Sub
 
Private Sub ClearData()
    CboCompPN.Clear
    CboVendorCode.Clear
    CboDateCode.Clear
    CboLotCode.Clear
    CboDID.text = ""
    TxtQty.text = "0"
    txtLine.text = ""
    txtSide.text = ""
    strDID = ""
    DIDCount = 0
End Sub

Private Function sq(ByVal Field As String) As String
   sq = "'" & Field & "'"
End Function

 
Private Sub RefreshGroupDIDQty()
Dim Rs As New ADODB.Recordset
Dim i As Integer
strSQL = "select CompPN,Qty from XL_MaxDIDMaintainQty order by comppn"
Set Rs = Conn.Execute(strSQL)
If Rs.EOF = False Then
    TempArry = Rs.GetRows
    ReDim arryGroupDIDQty(2, UBound(TempArry, 2)) As String
    For i = 0 To UBound(TempArry, 2)
        arryGroupDIDQty(0, i) = TempArry(0, i)
        arryGroupDIDQty(1, i) = TempArry(1, i)
    Next i
End If
End Sub
Private Sub LockTheForm(lockCtl As Boolean)
 Dim ctl As Control
 
 CboCompPN.Enabled = lockCtl
 CboVendorCode.Enabled = lockCtl
 CboDateCode.Enabled = lockCtl
 CboLotCode.Enabled = lockCtl
 CboDID.Enabled = lockCtl
 cmdFind.Enabled = lockCtl
 cmdSave.Enabled = lockCtl
 cmdCancel.Enabled = lockCtl
 CmdRefresh.Enabled = lockCtl
 cmdExit.Enabled = lockCtl
 cmdReprint.Enabled = lockCtl
 DG1.Enabled = lockCtl
End Sub
Private Function PrintLabel(strDID As String, strQty As String, Optional strAnotherQSMSIP As String) As String
    Dim M As Integer
    Dim tmpPrintStr As String
    Dim hFile As Long
    Dim hString As String
    Dim tmpDID As String
    Dim strDay As String
    Dim LabelFile As String
    Dim strPrinterType As String
    Dim tmpStr As String
    Dim tmpRS As ADODB.Recordset
    Dim rsTime As ADODB.Recordset
    Dim lptPort As Integer
    
    On Error GoTo errHandler
'    strSql = "select dbo.FormatDate (GETDATE() ,'DD')"
'    Set rsTime = Conn.Execute(strSql)
'    strDay = Trim(rsTime(0))

    strSQL = "select getdate()"
    Set rsTime = Conn.Execute(strSQL)
    strDay = Format(rsTime(0), "YYYY/MM/DD h:mm")        '(1091)
    
    
    If CHKAutoDispatchForAnotherBU = True And strAnotherQSMSIP <> "" Then
        tmpStr = "Select DIDHead from " & Trim(strAnotherQSMSIP) & ".QSMS.dbo.site"
    Else
        tmpStr = "Select DIDHead from site"
    End If
    Set tmpRS = Conn.Execute(tmpStr)
    
    If tmpRS.EOF Then
       MsgBox "can not find the DIDHead in the Table,Please check"
       Exit Function
    Else
        PrintData.BU = Trim(tmpRS!DIDHead)
    End If
    
      LabelFile = GetDIDLabelFile(frmMaintainDIDAutoDispatch, IIf(opOldLabel.Value = True, "OLD", "NEW"))
    If Dir(LabelFile) = vbNullString Then
        ''''''Added by Jing 2008.01.10  (0019)''''''
        MsgBox ("Can not find label file:" & LabelFile & "!"), vbCritical
        PrintLabel = "PRN_FileNoExist"
        Exit Function
    End If
    
    If opNewLabel.Value = True Then
        Dim x As Integer
        For x = 0 To 4
            WO(x) = ""
            Model(x) = ""
            Machine(x) = ""           '(1091)
            DIDType(x) = ""
            ISCYL(x) = ""
        Next x
        
        tmpStr = "Exec QSMS_GetDIDPrintInfo @DID='" & Trim(strDID) & "',@AnotherQSMSIP='" & Trim(strAnotherQSMSIP) & "',@PrinterType='" & Trim(PrinterType) & "',@PrintDpm='" & Trim(PrintDpm) & "'"        '''(0056)
        Set tmpRS = Conn.Execute(tmpStr)
        If tmpRS.EOF = False Then
            Dim i As Integer, j As Integer, ff As Integer
            j = tmpRS.RecordCount
            If j > 5 Then j = 5
            For i = 0 To j - 1
                WO(i) = tmpRS("Machine") + " " + tmpRS("Slot") + "-" + tmpRS("LR")
                Model(i) = tmpRS("model")
                Machine(i) = Mid(tmpRS("Machine"), 2, 1) + "-" + Mid(tmpRS("Slot"), 1, 1) + "-" + Mid(tmpRS("Machine"), 6, 1)     '(1091)
                Work_Order(i) = tmpRS("Work_Order")         '''1093
                DIDType(i) = tmpRS("DIDType")
                
                MachineCH(i) = tmpRS("MachineCH")         ''1044
                SideCH(i) = tmpRS("SideCH")         ''1044
                LRCH(i) = tmpRS("LRCH")         ''1044
                SlotCH(i) = tmpRS("Slot")         ''1044
                PN(i) = tmpRS("PN")
                
                If PrintedSeqID = "Y" Then  '(1147)
                    SeqID(i) = tmpRS("SeqID")
                End If
                If PrintedVenderCode = "Y" Then    ''1223
                    VenderCode(i) = tmpRS("VenderCode")
                    LR(i) = tmpRS("SLR")
                End If
                
                For ff = 0 To tmpRS.Fields.Count - 1    ''(1109)
                    If UCase(tmpRS.Fields(ff).Name) = "ISCYL" Then
                        ISCYL(i) = tmpRS("ISCYL")
                    End If
                Next ff
                tmpRS.MoveNext
            Next i
        End If
    End If
  ''(1080) replace by (1080)
'    If frmPrinterSetting.OptZebra.Value = True Then
'        isZebra = True
'        If opOldLabel.Value = True Then
'            LabelFile = Settings.AutoDispatchLabel
'        Else
'            LabelFile = Settings.AutoDispatchNewLabel
'        End If
'        strPrinterType = "Zebra"
'    Else
'        isZebra = False
'
'        ''''''updated by Jing   (0032)''''''
'        If opOldLabel.Value = True Then
'            LabelFile = Settings.AutoDispatchSatoLabel
'        Else
'            LabelFile = Settings.AutoDispatchSatoNewLabel
'        End If
'        strPrinterType = "SATO"
'    End If
    
'    LabelFile = GetDIDLabelFile(FrmDIDInteGration, IIf(opOldLabel.Value = True, "OLD", "NEW")) ''(1080) Get labelfile
'
'    If Dir(LabelFile) = vbNullString Then
'        MsgBox ("Can not find label file !"), vbCritical
'        PrintLabel = "PRN_FileNoExist"
'        Exit Function
'    End If
    
    If OptComp.Value = True Then
        MSComm.CommPort = TxtCompPort
        MSComm.Settings = TxtComm
        MSComm.OutBufferCount = 0
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
    If FileReadAll(tmpPrintStr, LabelFile) <= 0 Then
        MsgBox "Open file:" & LabelFile & " fail!!", vbCritical
    Exit Function
    End If

    
     tmpDID = Trim(strDID)
     If InStr(tmpPrintStr, "<DID_CODE>") > 0 Then
         If isZebra Then
             tmpDID = Replace(strDID, "^", "><")
         End If
        tmpPrintStr = Replace(tmpPrintStr, "<DID_CODE>", tmpDID)
     End If
     If InStr(tmpPrintStr, "<DID_TEXT>") > 0 Then
         If isZebra Then
            tmpDID = Replace(strDID, "^", "_5E")
         End If
        tmpPrintStr = Replace(tmpPrintStr, "<DID_TEXT>", tmpDID)
     End If
    
     tmpPrintStr = Replace(tmpPrintStr, "<UID>", uId)
     tmpPrintStr = Replace(tmpPrintStr, "<DATE>", strDay)
     tmpPrintStr = Replace(tmpPrintStr, "<QTY>", strQty)
     If opNewLabel.Value = True Then
        tmpPrintStr = Replace(tmpPrintStr, "<BU>", Trim(PrintData.Line) & "*")
     Else
        tmpPrintStr = Replace(tmpPrintStr, "<LINE>", Trim(PrintData.Line) & "*")
     End If
     
     tmpPrintStr = Replace(tmpPrintStr, "<SIDE>", PrintData.Side)
     tmpPrintStr = Replace(tmpPrintStr, "<MACHINE>", PrintData.Machine)
     tmpPrintStr = Replace(tmpPrintStr, "<DIDWOGROUP>", PrintData.DIDWOGROUP)
     tmpPrintStr = Replace(tmpPrintStr, "<WOTYPE>", WOType)
    If opNewLabel.Value = True Then
        tmpPrintStr = Replace(tmpPrintStr, "<WO1>", WO(0))
        tmpPrintStr = Replace(tmpPrintStr, "<WO2>", WO(1))
        tmpPrintStr = Replace(tmpPrintStr, "<WO3>", WO(2))
        tmpPrintStr = Replace(tmpPrintStr, "<WO4>", WO(3))
        tmpPrintStr = Replace(tmpPrintStr, "<WO5>", WO(4))
                
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINE1>", Machine(0))            '(1091)
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINE2>", Machine(1))
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINE3>", Machine(2))
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINE4>", Machine(3))
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINE5>", Machine(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<COUNT1>", SeqID(0))  '(1147)
        tmpPrintStr = Replace(tmpPrintStr, "<COUNT2>", SeqID(1))
        tmpPrintStr = Replace(tmpPrintStr, "<COUNT3>", SeqID(2))
        tmpPrintStr = Replace(tmpPrintStr, "<COUNT4>", SeqID(3))
        tmpPrintStr = Replace(tmpPrintStr, "<COUNT5>", SeqID(4))
        
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
        
         tmpPrintStr = Replace(tmpPrintStr, "<MachineCH1>", MachineCH(0))                    '1044
        tmpPrintStr = Replace(tmpPrintStr, "<MachineCH2>", MachineCH(1))
        tmpPrintStr = Replace(tmpPrintStr, "<MachineCH3>", MachineCH(2))
        tmpPrintStr = Replace(tmpPrintStr, "<MachineCH4>", MachineCH(3))
        tmpPrintStr = Replace(tmpPrintStr, "<MachineCH5>", MachineCH(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH1>", SideCH(0))                    '1044
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH2>", SideCH(1))
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH3>", SideCH(2))
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH4>", SideCH(3))
        tmpPrintStr = Replace(tmpPrintStr, "<SideCH5>", SideCH(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH1>", LRCH(0))                    '1044
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH2>", LRCH(1))
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH3>", LRCH(2))
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH4>", LRCH(3))
        tmpPrintStr = Replace(tmpPrintStr, "<LRCH5>", LRCH(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH1>", SlotCH(0))                    '1044
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH2>", SlotCH(1))
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH3>", SlotCH(2))
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH4>", SlotCH(3))
        tmpPrintStr = Replace(tmpPrintStr, "<SlotCH5>", SlotCH(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<LR1>", LR(0))                    '1223
        tmpPrintStr = Replace(tmpPrintStr, "<LR2>", LR(1))
        tmpPrintStr = Replace(tmpPrintStr, "<LR3>", LR(2))
        tmpPrintStr = Replace(tmpPrintStr, "<LR4>", LR(3))
        tmpPrintStr = Replace(tmpPrintStr, "<LR5>", LR(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<PN1>", PN(0))                    '1044
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
            If OptComp.Value = True Then
                For M = 1 To Len(tmpPrintStr) Step 100
                    MSComm.Output = Mid(tmpPrintStr, M, 100)
                    DoEvents
                Next M
                MSComm.PortOpen = False
            ElseIf OptPrint.Value = True Then
                For M = 1 To Len(tmpPrintStr) Step 50
                    Print #lptPort, Mid(tmpPrintStr, M, 50)
                    DoEvents
                Next M
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
Private Sub GetOneByOneMaterialDemand(COMPPN As String, VendorCode As String, DateCode As String, LotCode As String)
    Dim Rs As New ADODB.Recordset
    On Error GoTo Err_Handler
    
    LblStatus = ""
    strSQL = "exec [XL_Dispatch_MaterialPrompt_superchai]  @CompPN=" & sq(COMPPN) & ",@VendorCode=" & sq(VendorCode) & ",@DateCode=" & sq(DateCode) & ",@LotCode=" & sq(LotCode) & ",@Factory=" & sq(Factory)   'Add superchai 20240125
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = False Then
        If Rs("Result") = 0 Then
            Set Rs = Rs.NextRecordset
            Call FillFlexData(Rs, flexGridDemandMaterial)
        Else
            Call InitFlex(flexGridDemandMaterial)
            LblStatus = Rs("Description")
        End If
    End If
    
    Exit Sub
Err_Handler:
    
    MsgBox Err.Number & "," & Err.Description
End Sub
Private Sub InitFlex(Flex As MSFlexGrid)
    Dim intCol As Integer
    
    With Flex
        .Rows = 1
        .Rows = 20
        .Cols = 24
        
        .FormatString = "|GroupID|WO|WOqty|Machine|Slot|LR|CompPN|Item|BaseQty|NeedQty|DispatchQty|BalanceQty|PlanQty|PlanNeedQty|PlanBalanceQty|WorkDate|Shift|WOSeqID|SAPPercentage|Jobgroup|Jobpn|Line|Side"
        
        .ColWidth(0) = 300
        .ColWidth(1) = 1800
        .ColWidth(2) = 1000
        .ColWidth(3) = 600
        .ColWidth(4) = 800
        .ColWidth(5) = 600
        .ColWidth(6) = 420
        .ColWidth(7) = 1260
        .ColWidth(8) = 420
        .ColWidth(9) = 800
        .ColWidth(10) = 900
        .ColWidth(11) = 1000
        .ColWidth(12) = 1000
        .ColWidth(13) = 800
        .ColWidth(14) = 1100
        .ColWidth(15) = 1200
        .ColWidth(16) = 900
        .ColWidth(17) = 600
        .ColWidth(18) = 1000
        .ColWidth(19) = 1260
        .ColWidth(20) = 1000
        .ColWidth(21) = 1000
        .ColWidth(22) = 600
        .ColWidth(23) = 600
        
        .Col = 1
        .CellAlignment = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        For intCol = 2 To .Cols - 1
            .row = 0
            .Col = intCol
            .CellAlignment = flexAlignCenterCenter '4
            .ColAlignment(intCol) = flexAlignLeftCenter  '1
        Next intCol
    
    End With
    
End Sub

Private Sub FillFlexData(Rst As ADODB.Recordset, Flex As MSFlexGrid)
    Dim IntRow As Integer
    Dim intCol As Integer
    With Flex
        .Rows = 1
        .Rows = Rst.RecordCount + 1
        Do While Rst.EOF = False
            IntRow = IntRow + 1
             
            
            For intCol = 1 To .Cols - 1
                .TextMatrix(IntRow, intCol) = Trim(Rst.Fields(intCol - 1) & "")
                
                If intCol = 12 Or intCol = 15 Then
                    .Col = intCol
                    .row = IntRow
                    flexGridDemandMaterial.CellBackColor = vbRed
                End If
            Next intCol
            Rst.MoveNext
            
        Loop
        
        If .Rows = 1 Then
            .Rows = 20
        End If
        
    End With
End Sub
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
Private Function GetPrinterSetting(frm As Form)
On Error GoTo errhandle:
    
    If GetSetting("SMT", "QSMS", "Printer") = "Zebra" Then
        frm.OptZebra.Value = True
    Else
        frm.OptSATO.Value = True
    End If
    
    If GetSetting("SMT", "QSMS", "Port") = "COM" Then
        frm.OptComp.Value = True
    ElseIf GetSetting("SMT", "QSMS", "Port") = "LPT" Then
        frm.OptPrint.Value = True
    Else
        frm.optNetwork.Value = True
    End If
        
    If GetSetting("SMT", "QSMS", "CommPort") <> "" Then
        frm.TxtCompPort.text = GetSetting("SMT", "QSMS", "CommPort")
    Else
        frm.TxtCompPort.text = "1"
    End If
    
    If GetSetting("SMT", "QSMS", "Comm") <> "" Then
        frm.TxtComm.text = GetSetting("SMT", "QSMS", "Comm")
    Else
        frm.TxtComm.text = "9600,N,8,1"
    End If
    
    frm.OptZebra.Enabled = False
    frm.OptSATO.Enabled = False
    frm.OptComp.Enabled = False
    frm.OptPrint.Enabled = False
    frm.optNetwork.Enabled = False
    frm.TxtCompPort.Enabled = False
    frm.TxtComm.Enabled = False
    frm.CmdCommSave.Visible = False
    
Exit Function

errhandle:
    MsgBox Err.Description
End Function

 
