VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmMaintainDIDAutoDispatch 
   Caption         =   "Maintain DID & Auto dispatch  20170314A"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   15660
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl wave_control 
      Height          =   495
      Left            =   2880
      TabIndex        =   72
      Top             =   4800
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid flexGridDemandMaterial 
      Height          =   2055
      Left            =   1440
      TabIndex        =   70
      Top             =   5280
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   3625
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin VB.Frame FraConnection 
      BackColor       =   &H80000013&
      Caption         =   "DID maintain "
      Height          =   3135
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   15375
      Begin VB.TextBox txtImgVersion 
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
         Left            =   8160
         TabIndex        =   74
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox CboCompPN 
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
         Left            =   1920
         TabIndex        =   71
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox txtLotCode 
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
         Left            =   6720
         TabIndex        =   66
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtDateCode 
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
         Left            =   1920
         TabIndex        =   65
         Top             =   720
         Width           =   3015
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
         Left            =   10320
         TabIndex        =   64
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
      End
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
         Left            =   13680
         TabIndex        =   62
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cmbLine 
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
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtExtraQty 
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
         Left            =   13440
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ComboBox cmbLR 
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
         ItemData        =   "frmMaintainDIDAutoDispatch.frx":0000
         Left            =   11520
         List            =   "frmMaintainDIDAutoDispatch.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "DispatchType"
         Height          =   615
         Left            =   9720
         TabIndex        =   38
         Top             =   120
         Width           =   5535
         Begin VB.OptionButton optToWO 
            Caption         =   "ToWO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3600
            TabIndex        =   60
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optExtra 
            Caption         =   "Extra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1320
            TabIndex        =   58
            Top             =   200
            Width           =   855
         End
         Begin VB.OptionButton OptNormal 
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optSpecial 
            Caption         =   "Special"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2400
            TabIndex        =   56
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbSlot 
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
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ComboBox cmbSide 
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
         ItemData        =   "frmMaintainDIDAutoDispatch.frx":0004
         Left            =   7440
         List            =   "frmMaintainDIDAutoDispatch.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1095
      End
      Begin VB.ComboBox cmbMachine 
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
         Left            =   12720
         Style           =   2  'Dropdown List
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2535
      End
      Begin VB.ComboBox cmbWO 
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
         Left            =   9360
         Style           =   2  'Dropdown List
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1935
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
         Left            =   10440
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   720
         Width           =   1335
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
         Left            =   6720
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
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
         Left            =   840
         TabIndex        =   19
         Top             =   1200
         Width           =   5415
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
         Left            =   13080
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "1"
         Top             =   720
         Width           =   1335
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
         Left            =   2640
         Picture         =   "frmMaintainDIDAutoDispatch.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1680
         Width           =   855
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
         Left            =   3480
         Picture         =   "frmMaintainDIDAutoDispatch.frx":0A0A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1680
         Width           =   855
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
         Left            =   4320
         Picture         =   "frmMaintainDIDAutoDispatch.frx":0D14
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1680
         Width           =   855
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
         Left            =   120
         Picture         =   "frmMaintainDIDAutoDispatch.frx":101E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1680
         Width           =   855
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
         Picture         =   "frmMaintainDIDAutoDispatch.frx":1460
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1680
         Width           =   855
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
         Left            =   960
         Picture         =   "frmMaintainDIDAutoDispatch.frx":18A2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1680
         Width           =   855
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
         Left            =   15360
         Picture         =   "frmMaintainDIDAutoDispatch.frx":1BAC
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdReprint 
         Caption         =   "Reprint"
         DragIcon        =   "frmMaintainDIDAutoDispatch.frx":1FEE
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
         Picture         =   "frmMaintainDIDAutoDispatch.frx":7C00
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1680
         Width           =   735
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
         Left            =   5160
         Picture         =   "frmMaintainDIDAutoDispatch.frx":D812
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton cmdForceDel 
         Caption         =   "Command1"
         Height          =   255
         Left            =   14520
         TabIndex        =   8
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label LabelImgVersion 
         BackColor       =   &H0000FF00&
         Caption         =   "ImageVserion"
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
         Left            =   6360
         TabIndex        =   73
         Top             =   1200
         Width           =   1695
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
         Left            =   9480
         TabIndex        =   63
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
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
         Left            =   12000
         TabIndex        =   61
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblCount 
         BackColor       =   &H80000003&
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
         Left            =   12240
         TabIndex        =   59
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label7 
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
         Left            =   6840
         TabIndex        =   44
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Index           =   5
         Left            =   12720
         TabIndex        =   41
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
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
         Index           =   4
         Left            =   10800
         TabIndex        =   40
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
         Caption         =   "Slot"
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
         Index           =   3
         Left            =   8520
         TabIndex        =   36
         Top             =   2160
         Width           =   735
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
         Index           =   2
         Left            =   6840
         TabIndex        =   34
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FF00&
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
         Left            =   11520
         TabIndex        =   31
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         Caption         =   "WO"
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
         Left            =   8520
         TabIndex        =   30
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFF80&
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2640
         Width           =   12015
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
         Left            =   9720
         TabIndex        =   28
         Top             =   720
         Width           =   735
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
         Left            =   4920
         TabIndex        =   27
         Top             =   720
         Width           =   1815
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
         Left            =   4920
         TabIndex        =   26
         Top             =   240
         Width           =   1815
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
         TabIndex        =   25
         Top             =   240
         Width           =   1815
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
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
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
         TabIndex        =   23
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
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
         Left            =   11760
         TabIndex        =   22
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Comp Port data maintain"
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8055
      Begin VB.CheckBox IsTrayComp 
         Caption         =   "IsTrayComp"
         Height          =   330
         Left            =   1440
         TabIndex        =   75
         Top             =   1290
         Width           =   1575
      End
      Begin VB.OptionButton OptNetwork 
         BackColor       =   &H80000013&
         Caption         =   "NetWork"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   1320
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "Label"
         Height          =   550
         Left            =   6120
         TabIndex        =   53
         Top             =   720
         Width           =   1815
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
            TabIndex        =   55
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
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
            TabIndex        =   54
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.OptionButton OptComp 
         BackColor       =   &H80000013&
         Caption         =   "Comp Port"
         Height          =   255
         Left            =   120
         TabIndex        =   50
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
         Height          =   375
         Left            =   6360
         Picture         =   "frmMaintainDIDAutoDispatch.frx":DB1C
         TabIndex        =   49
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
         Left            =   4680
         TabIndex        =   48
         Text            =   "9600,N,8,1"
         Top             =   240
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
         TabIndex        =   47
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton OptPrint 
         BackColor       =   &H80000013&
         Caption         =   "Print Port"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Printer"
         Height          =   550
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   4575
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
            TabIndex        =   4
            Top             =   240
            Width           =   1455
         End
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
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
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
         TabIndex        =   52
         Top             =   240
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
         TabIndex        =   51
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frmaVendorDesc 
      BackColor       =   &H80000013&
      Caption         =   "Vendor description"
      Height          =   1695
      Left            =   8160
      TabIndex        =   0
      Top             =   0
      Width           =   7335
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
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   240
         Width           =   7095
      End
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   14160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSDataGridLib.DataGrid DG1 
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   7320
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   3625
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
   Begin MSDataGridLib.DataGrid DGOneByOne 
      Height          =   855
      Left            =   2280
      TabIndex        =   67
      Top             =   6240
      Visible         =   0   'False
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   1508
      _Version        =   393216
      BackColor       =   16777215
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
            Type            =   5
            Format          =   "0%"
            HaveTrueFalseNull=   1
            TrueValue       =   "True"
            FalseValue      =   "False"
            NullValue       =   ""
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
   Begin VB.Label lblOneByOne 
      Caption         =   "Demand Material:"
      Height          =   225
      Left            =   120
      TabIndex        =   68
      Top             =   4920
      Width           =   1935
   End
End
Attribute VB_Name = "frmMaintainDIDAutoDispatch"
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
'EQMS_ID             修 改 人     修改日期        描    述
'------------------------------------------------------------------------------------------------------------------------
'                    Jeanson      2007.10.15     update the condition of checking DID is using or not (0001)
'                    Jing         2007.10.24     check the printer is Zebra or Sato for print barcode (0002)
'                    Sandy        2007.11.05     if SATO printer will not delay time                  (0003)
'                    Jeanson      2007.11.07     update realqty/lastupdatetime instead of transdatetime when updating (0004)
'                    Sandy        2007.11.16     update the DID deleted function. (delete DID both in QSMS and in FUJI DB)(0005)
'                    Jing         2007.11.26     control the delete and reprint (0006)
'                    Sandy        2007.11.26     update to drive out the blank space in DID.---------(0007)
'                    Sandy        2007.11.30     added that don't allow to delete dispatched DID in MaintainDID----(0008)
'                    Sandy        2007.12.13     set default Printer according to each BU whether have FUJITrax ----(0009)
'                    Jing         2007.12.18     add the parameter Comppn to shrink the return data     (0010)
'                    Kane         2007.12.19     Add a dispatch type for extra dispatching(0011)
'                    Jing         2007.12.21     Add a Line in MaintainDIDAutoDispatch    (0012)
'                    Jing         2007.12.25     Check whether need dispatch materail after inputting the comppn  (0013)
'                    Jing         2007.12.26     Add an option of DelayTime for Zebra printer     (0014)
'                    Sandy        2007.12.28     Add DIDWoGoup into DIDAutoDispatck label     (0015)
'                    Kane         2008.01.02     Send 50 code to printer when sato printer    (0016)
'                    Jing         2008.01.02     Add DelayTime(0.1s) for Zebra and Sato printer     (0017)
'                    Jing         2008.01.03     Add reprint record in QMS_LOG    (0018)
'                    Jing         2008.01.10     Add error alarm when can not find label file     (0019)
'                    Jing         2008.01.10     Changed from 'NB25' to 'NB5'     (0020)
'                    Jing         2008.01.12     check previous printed DID whether it is the same as the current     (0021)
'                    Alex         2008.01.12     Move Code from CboCompPN_Click to  CboCompPN_LostFocus   (0022)
'                    Lynn         2008.01.22     Show DID Created just now  (0023)
'                    Salon        2008.01.25     not set the datecode/lotcode/VendorCode value automatically (0024)
'                    Jeanson      2008.01.25     save the DID print log (0025)
'                    Jeanson      2008.02.02     dispatch DID accordingto wo total need (0026)
'                    Jeanson      2008.02.20     trim the DID total value's space  (0027)
'                    Jing         2008.03.05     check the value(max value from set.ini) of txtGroupQty    (0028)
'                    Jeanson      2008.03.18     to check whether the vendor length is less than 8  (0029)
'                    Jeanson      2008.03.18     add the log of DID machine data into qsms_error_log   (0030)
'                    Sandy        2008.03.26     check ' in the datecode/lotcode/VendorCode value---(0031)
'                    Jing         2008.04.05     add new label format for Print DID label     (0032)
'                    Udall        2008.04.13     Add Factory  (0033)
'                    Jing         2008.04.22     Update Error Handle  (0034)
'                    Sandy        2008.04.25     Check whether need dispatch materail after inputting the comppn  (0035)
'                    Sandy        2008.04.26     add LockTheForm (True)when programe exit abnormal  (0036)
'                    Sandy        2008.05.06     update the Auto dispatch label  (0037)
'                    Jing         2008.05.30     Cancel:save the DID print log (0038)
'                    Kane         2008.06.11     Add wotype on did label,when did dispatched to wo which pilot is new then new,
'                                                  when pilot is eol and not exists pilot is new then eol else is empty (0039)
'                    Kane         2008.06.23     Add a constraint before print did label (0040)
'                    Lynn         2008.08.19     if the DID has been return or callback, PD can not do reprint (0041)
'                    Kane         2008.08.21     Add function to check groupqty can not more then definition in XL_MaxDIDMaintainQty table by comppn '(0042)
'                    Jeanson      2008.09.09     not to display the DID which Qty=0  (0043)
'                    Kane         2008.11.06     Add new dispatch type 5 dispatch to fix wo '(0044)
'                    Udall        2009.02.19     Add new function for check the DateCode and Inspection,save the Inspection (0045)
'                    Sandy        2009.04.21     add tmpRS("LR") in new auto diapach (0046)
'RQ09052213          Salon        2009.05.26    Change all character into upper case  (0047)
'QMS                 Sandy        2009.07.14    add fuction IsInteger to check integer(0048)
'EQMS                Kane         2009.07.15    check vendor code can not be nemeric '(0049)
'QMS                 Sandy        2009.07.22    GetDIDQty & chkNumber are replace by IsInteger(0050)
'RQ09070915          Udall        2009.08.10    校验几种特殊的CompPN的DateCode（For AP SMT） (0051)
'QMS                 Archer       2009.08.19    AutoDispatchForAnotherBU-->NB2/3物料不分,导致前台产生的DID不一定就是系统真正使用的DID (0052)
'QMS                 Kane         2009.08.26    更改LPT打印方式替换模板方式与Com打印方式相同   '(0053)
'QMS                 udall        2009.09.01    核对输入的QTY数量大小<100000，但不卡字符长度   '(0054)
'RQ09091301          Archer       2009/09/20    输入完Group Qty栏位内数量后按Enter后系统就可以列印出DID号码(0055)
'RQ09091804          Archer       2009/10/08    Modify DIDMaintain program to print the slot according to PMC WO scheduling(0056)
'QMS                 udall        2009/10/23    针对AP需刷入InspectionNo，因此操作动作与其他BU不同 (0057)
'QMS                 Kane         2009/10/30    打印时把线别改为机台的第一码，并纪录线别和机器对应的线别不匹配的情况。'(0058)
'RQ09122849          Archer       2010/03/06    Modify program to use new label format as default option(0059)
'QMS                 Archer       2010/03/09    Add Factory for SP:XL_GetDidPrintInfo  and XL_CheckNeedDispatch (0060)
'RQ10042606          Lynn         2010/05/26    For NB6, extra define the ChkOldDIDLabelQty individual. (0061)
'RQ10051758          Kevin        2010/06/04    MaintainDidAutoDispatch的时候不进行匹配查询   (0062)
'QMS                 Austin       2010/06/09    Check if is NeedMSD   (0063)
'QMS                 Austin       2010/0623     Modify On 0062(Kevin) (0064)
'QMS                 Austin       2010/07/07    not set the CompPN automatically  (0065)
'RQ10071501          Denver       2010/07/23    For One By One Material,it need show Current/Next Shift Material Demand  (0076)
'RQ10071501          Denver       2010/07/26    非OneByOne 材料也需要按此方式显示，并将 BalanceQty/PlanBalanceQty 突出显示  (0076)
'QMS                 Kyle         2010/08/10    Change the way to read template file in order to print label by the network printer. (0077)
'QMS                 Scofield     2012/06/21    核对输入的QTY数量大小由100000改为120000 (0078)
'QMS                 Rain         2019/07/17    核对DID对应DateCode 是否大于定义DateCode （0079）

'***********************************************************************************/
Dim Rs2 As ADODB.Recordset
Dim CommandType As Long
Dim isZebra As Boolean
Dim ExtraData As Extra
Dim PrintData As PtData
Dim strAnotherQSMSIP As String
Dim strLine As String
Dim strSide As String, strSQL As String
Dim AutoDispatchLabelFile As String
Dim AutoDispatchSatoLabelFile As String
Dim WOType As String
Dim arryGroupDIDQty() As String
Dim TempArry
Dim CHKAutoDispatchForAnotherBU As Boolean
Dim ReturnDID As Boolean ''判读是否为returnDID 的flag
Dim NewcompFlag As Boolean  '’判断是否为2DBarcode 的flag
Dim strDelaytime As Long
Dim strCheckScaner As String
Dim strVendorCode As String
Dim OldReturnDID As String  '1191 记录callback退仓后生成的DID


Private Sub CboCompPN_Click()
SendKeys "{HOME}+{END}"
End Sub

Private Sub CboCompPN_KeyPress(KeyAscii As Integer)

Dim NewComp() As String, index As Integer
Dim strSQL As String
Dim RS As New ADODB.Recordset
Dim str As String, I As Integer, IsBSMaterial As String
NewcompFlag = False
ReturnDID = False
OldReturnDID = "" '1191

If StrBU = "NB6" Then
    If strKeyInPNByManual = True Then
        strCheckScaner = "N"
    Else
        strCheckScaner = "Y"
    End If
Else
    If strKeyInPNByManual = True Then
        strCheckScaner = "N"
    End If
End If
If strCheckScaner = "Y" Then
    If Len(Trim(CboCompPN.Text)) < 1 Then strDelaytime = 0
        If strDelaytime <> 0 Then
            If GetTickCount - strDelaytime > 100 Then
                MsgBox "Please use scaner!"
                CboCompPN.Text = ""
                strDelaytime = 0
                Exit Sub
            End If
    End If
    strDelaytime = GetTickCount
End If
If strCheckScaner = "Y" And (KeyAscii = 13) Then
    strDelaytime = 0
End If


If (KeyAscii = 13 Or KeyAscii = 9) And CboCompPN <> "" Then
    CboCompPN.Text = UCase(Replace(Replace(Replace(Replace(CboCompPN.Text, " ", ""), vbCr, ""), vbLf, ""), ":", ";")) ''1247
    
    If InStr(1, Trim(CboCompPN.Text), ";") > 0 Then  '(1020)
        ''(1018) Start ---------------------''
        NewcompFlag = True
        NewComp = Split(Trim(CboCompPN.Text), ";")
        For index = 0 To UBound(NewComp)
            If index = 0 Then
                CboCompPN.Text = Trim(NewComp(index))
                
                ''If StrBU = "NB5" Then     ''1266
                    ''strSql = "SELECT * FROM HPLink_QuantaPN WHERE HPPN='" & Trim(CboCompPN.Text) & "'"
                    ''Set rs = Conn.Execute(strSql)
                    ''If rs.EOF = False Then
                       ''CboCompPN.Text = Trim(rs!QuantaPN)
                    ''End If
                ''End If                    ''1266
                
            ElseIf index = 1 Then
                txtDateCode.Text = Trim(NewComp(index))
            ElseIf index = 2 Then
                CboVendorCode.Text = Trim(NewComp(index))
            ElseIf index = 3 Then
                txtLotCode.Text = Trim(NewComp(index))
            ElseIf index = 4 Then
                TxtQty.Text = Trim(NewComp(index))
            End If
                
        Next index
        If CheckBSMaterial = "Y" Then   ''(1213)
            IsBSMaterial = "N"
            If txtLotCode = "@@@@" Then
                IsBSMaterial = "Y"
            Else
                strSQL = "select 1 from Component_Data where comppn='" + CboCompPN.Text + "' and [Functype]='BSMaterial'"
                Set RS = Conn.Execute(strSQL)
                If RS.EOF = False Then
                    IsBSMaterial = "Y"
                End If
            End If
        End If
        If IsBSMaterial = "Y" Then   ''1208
            txtDateCode.Text = ""
            txtLotCode.Text = ""
            txtDateCode.Text = UCase(Trim(InputBox("请刷入Datecode:", "Input Datecode")))
            txtLotCode.Text = UCase(Trim(InputBox("请刷入Lotcode:", "Input Lotcode")))
        End If
        
    ''(1018) End -----------------------''
    ElseIf InStr(1, Trim(CboCompPN.Text), "-") > 0 And Len(CboCompPN.Text) > 15 Then  '(1020)
        ReturnDID = True
        strSQL = "select CompPN,VendorCode,DateCode ,LotCode ,Qty from QSMS_DID_ToWH where DID = '" & Trim(CboCompPN.Text) & "' "
        Set RS = Conn.Execute(strSQL)
        If RS.EOF = False Then
            OldReturnDID = Trim(CboCompPN.Text) ''1191
            CboCompPN.Text = Trim(RS!compPN)
            CboVendorCode.Text = Trim(RS!VendorCode)
            txtDateCode.Text = Trim(RS!DateCode)
            txtLotCode.Text = Trim(RS!LotCode)
            TxtQty.Text = Trim(RS!Qty)
        Else
            MsgBox ("Can't find the information of this returnDID---" & Trim(CboCompPN.Text))
            CboCompPN_Click
            Call Warning_Sound
            Call ClearData
            Exit Sub
        End If
    End If
    strErrMessage = ""
    strErrMessage = FunPartNumberCheck(CboCompPN.Text)
    If strErrMessage <> "PASS" Then
        MsgBox strErrMessage
        CboCompPN_Click
        Call Warning_Sound
        Call ClearData
        Exit Sub
    End If
    If Check_AVL = "Y" Then
         str = "Select distinct VendorCode from QSMS_AVL where CompPN='" & Trim(CboCompPN) & "'"
         Set RS = Conn.Execute(str)
         I = 0
         CboVendorCode.Clear
         While Not RS.EOF
           CboVendorCode.AddItem Trim(RS!VendorCode)
           RS.MoveNext
           I = I + 1
         Wend
    End If
    '''(0060)
    
 If IsTrayComp = 1 Then '''1251
    str = "exec CompPN_Check" & Trim(sq(CboCompPN.Text))
    Set RS = Conn.Execute(str)
    If Trim(RS!result) <> "1" Then
        MsgBox ("Message: " & RS!ErrDesc), vbCritical
        CboCompPN_Click
        Call EffectSound("OO.wav")
        Call ClearData
        Exit Sub
    End If
End If

    str = "Exec XL_CheckNeedDispatch @Type='" & Trim(OptNormal.Value) & "',@CompPN='" & Trim(CboCompPN) & "',@Factory='" & Trim(Factory) & "'"
    Set RS = Conn.Execute(str)
    If Trim(RS!result) <> "1" Then
        MsgBox ("Message: " & RS!ErrDesc), vbCritical
        CboCompPN_Click
        'Call Warning_Sound
        Call EffectSound("OO.wav")
        Call ClearData
        Exit Sub
    End If
 

    
    If NewcompFlag = False And ReturnDID = False Then    ''(1018)'(1020)'--remove from CboCompPN_keypress by kaitlyn
        CboVendorCode.SetFocus
    ElseIf ChkOneByOneMaterial = "Y" Then
       Call TxtLotCode_KeyPress(13)
    Else
'        TxtQty.SetFocus
'        Call TxtQty_Click
        TxtGroupQty.SetFocus
        Call TxtGroupQty_Click
    End If
End If



End Sub

Private Sub CboCompPN_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If strKeyInPNByManual = True Then
    strCheckScaner = "N"
End If
If strCheckScaner = "Y" Then
    If App.Title <> App.EXEName Then
        If Shift = 2 Then
            MsgBox "Can't use Ctrl+V and Ctrl+C,input is void!", vbCritical
            CboCompPN.Text = ""
        End If
    End If
End If
End Sub

'Private Sub CboDateCode_Click()
'CboLotCode.SetFocus
'End Sub

Private Sub TxtDateCode_KeyPress(KeyAscii As Integer)

'(1100) begin
If strKeyInPNByManual = True Then
    strCheckScaner = "N"
End If

If strCheckScaner = "Y" Then
    If Len(Trim(txtDateCode.Text)) < 1 Then strDelaytime = 0
        If strDelaytime <> 0 Then
            If GetTickCount - strDelaytime > 100 Then
                MsgBox "Please use scaner!"
                txtDateCode.Text = ""
                strDelaytime = 0
                Exit Sub
            End If
    End If
    strDelaytime = GetTickCount
End If

If strCheckScaner = "Y" And (KeyAscii = 13 Or KeyAscii = 9) Then
    strDelaytime = 0
End If
'(1100) end

If KeyAscii = 13 Or KeyAscii = 9 Then
    txtLotCode.SetFocus
   ''Call CboDateCode_Click
End If
End Sub

Private Sub CboDID_Click()
Call cmdFind_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call UnHook  '(0011)撤消钩子
End Sub

'Private Sub CboLotCode_Click()
'TxtQty.SetFocus
'End Sub

Private Sub TxtLotCode_KeyPress(KeyAscii As Integer)
'(1100) begin
If strKeyInPNByManual = True Then
    strCheckScaner = "N"
End If

If strCheckScaner = "Y" Then
    If Len(Trim(txtLotCode.Text)) < 1 Then strDelaytime = 0
        If strDelaytime <> 0 Then
            If GetTickCount - strDelaytime > 100 Then
                MsgBox "Please use scaner!"
                txtLotCode.Text = ""
                strDelaytime = 0
                Exit Sub
            End If
    End If
    strDelaytime = GetTickCount
End If

If strCheckScaner = "Y" And (KeyAscii = 13 Or KeyAscii = 9) Then
    strDelaytime = 0
End If
'(1100) end
If UCase(Trim(CheckNeedMSD)) = "Y" Then   ''(1228)
    If IsNeedMSD(CboCompPN) = True Then
        MsgBox "这是MSD材料，请先烘烤！"
    End If
End If
If KeyAscii = 13 Or KeyAscii = 9 Then

     'RQ10071501          Denver       2010/07/23    For One By One Material,it need show Current/Next Shift Material Demand  (0076)
     If ChkOneByOneMaterial = "Y" And Trim(CboCompPN) <> "" And Trim(CboVendorCode) <> "" And Trim(txtDateCode) <> "" And Trim(txtLotCode) <> "" Then
         Call GetOneByOneMaterialDemand(Trim(CboCompPN), Trim(CboVendorCode), Trim(txtDateCode), Trim(txtLotCode))
           
     End If
    
     TxtQty.SetFocus
End If
    
End Sub

Private Sub CboVendorCode_Click()
Dim str As String
Dim RS As ADODB.Recordset
If Check_AVL = "Y" Then
    str = "Select Desc1 from QSMS_AVL where CompPN='" & Trim(CboCompPN) & "' and VendorCode='" & Trim(CboVendorCode) & "'"
    Set RS = Conn.Execute(str)
    txtDesc.Text = Trim(RS!Desc1)
End If
txtDateCode.SetFocus
'If ChkAVL(Trim(CboCompPN), Trim(CboVendorCode)) = True Then
'    CboDateCode.SetFocus
'Else
'    CboVendorCode.Text = ""
'    Exit Sub
'End If
End Sub

Private Sub CboVendorCode_KeyPress(KeyAscii As Integer)

'(1100) begin
If strKeyInPNByManual = True Then
    strCheckScaner = "N"
End If

If strCheckScaner = "Y" Then
    If Len(Trim(CboVendorCode.Text)) < 1 Then strDelaytime = 0
        If strDelaytime <> 0 Then
            If GetTickCount - strDelaytime > 100 Then
                MsgBox "Please use scaner!"
                CboVendorCode.Text = ""
                strDelaytime = 0
                Exit Sub
            End If
    End If
    strDelaytime = GetTickCount
End If

If strCheckScaner = "Y" And (KeyAscii = 13 Or KeyAscii = 9) Then
    strDelaytime = 0
End If
'(1100) end

If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboVendorCode_Click
End If
End Sub

'''''''''''''''''''''''''''''''''Added by Jing 2007.12.21 (0012)''''''''''''''''''''''''''''''''''
Private Sub cmbLine_Click()
Dim RS As New ADODB.Recordset

On Error GoTo errHandler
    If cmbLine = "" Then Exit Sub
    cmbWO.Clear
    cmbMachine.Clear
    cmbSide.Clear
    cmbSlot.Clear
    cmbLR.Clear
    
    strSQL = "Exec XL_GetAllWOInfoList 'WO','','','','','','" & Trim(CboCompPN.Text) & "','" & Trim(cmbLine.Text) & "'"
    Set RS = Conn.Execute(strSQL)
    If RS.EOF = False Then
        While Not RS.EOF
            cmbWO.AddItem (Trim(RS!GroupValue))
            RS.MoveNext
        Wend
    End If
    cmbWO.SetFocus
    Exit Sub
errHandler:
        MsgBox Err.Description
End Sub

Private Sub cmbLR_Click()
    If cmbLR = "" Then Exit Sub
    txtExtraQty.SetFocus
    Call txtExtraQty_Click
End Sub

Private Sub cmbMachine_Click()
Dim RS As ADODB.Recordset
On Error GoTo errHandler
    If cmbWO = "" Or cmbMachine = "" Then Exit Sub
    cmbSide.Clear
    cmbSlot.Clear
    cmbLR.Clear
    ''''''''''''''add by Jing 2007.12.18    (0010)''''''''''''''
    ''''''''''''''Updated by Jing 2007.12.21    (0012)'''''''''''
    
    strSQL = "Exec XL_GetAllWOInfoList 'Side','" & Trim(cmbWO) & "','" & Trim(cmbMachine) & "','','','','" & Trim(CboCompPN.Text) & "','" & Trim(cmbLine) & "'"
    Set RS = Conn.Execute(strSQL)
    If RS.EOF = False Then
        While Not RS.EOF
            cmbSide.AddItem (Trim(RS!GroupValue))
            RS.MoveNext
        Wend
    End If
    cmbSide.SetFocus
    Exit Sub
errHandler:
        MsgBox Err.Description
End Sub
Private Sub cmbSide_Click()
Dim RS As ADODB.Recordset
On Error GoTo errHandler
    If cmbWO = "" Or cmbMachine = "" Then Exit Sub
    cmbSlot.Clear
    cmbLR.Clear
    ''''''''''''''add by Jing 2007.12.18    (0010)''''''''''''''
    ''''''''''''''Updated by Jing 2007.12.21    (0012)'''''''''''
    
    strSQL = "Exec XL_GetAllWOInfoList 'Slot','" & Trim(cmbWO) & "','" & Trim(cmbMachine) & "','" & Trim(cmbSide) & "','','','" & Trim(CboCompPN.Text) & "','" & Trim(cmbLine) & "'"
    Set RS = Conn.Execute(strSQL)
    If RS.EOF = False Then
        While Not RS.EOF
            cmbSlot.AddItem (Trim(RS!GroupValue))
            RS.MoveNext
        Wend
    End If
    cmbSlot.SetFocus
    Exit Sub
errHandler:
        MsgBox Err.Description
End Sub

Private Sub cmbSlot_Click()
Dim RS As ADODB.Recordset
On Error GoTo errHandler
    If cmbWO = "" Or cmbMachine = "" Or cmbSide = "" Or cmbSlot = "" Then Exit Sub
    cmbLR.Clear
    
    ''''''''''''''add by Jing 2007.12.18    (0010)''''''''''''''
    ''''''''''''''Updated by Jing 2007.12.21    (0012)'''''''''''
    
    strSQL = "Exec XL_GetAllWOInfoList 'LR','" & Trim(cmbWO) & "','" & Trim(cmbMachine) & "','" & Trim(cmbSide) & "','" & Trim(cmbSlot) & "','','" & Trim(CboCompPN.Text) & "','" & Trim(cmbLine) & "'"
    Set RS = Conn.Execute(strSQL)
    If RS.EOF = False Then
        While Not RS.EOF
            cmbLR.AddItem (Trim(RS!GroupValue))
            RS.MoveNext
        Wend
    End If
    cmbLR.SetFocus
    Exit Sub
errHandler:
        MsgBox Err.Description
End Sub

Private Sub cmbWO_Click()
Dim RS As New ADODB.Recordset

On Error GoTo errHandler
    If cmbWO = "" Then Exit Sub
    cmbMachine.Clear
    cmbSide.Clear
    cmbSlot.Clear
    cmbLR.Clear
    ExtraData.Group = ""
    ExtraData.WO = cmbWO
    
    ''''''''''''''add by Jing 2007.12.18    (0010)''''''''''''''
    ''''''''''''''Updated by Jing 2007.12.21    (0012)'''''''''''
    strSQL = "Exec XL_GetAllWOInfoList 'Machine','" & Trim(cmbWO) & "','','','','','" & Trim(CboCompPN.Text) & "','" & Trim(cmbLine) & "'"
    Set RS = Conn.Execute(strSQL)
    If RS.EOF = False Then
        While Not RS.EOF
            cmbMachine.AddItem (Trim(RS!GroupValue))
            RS.MoveNext
        Wend
    End If
    Set RS = RS.NextRecordset
    If RS.EOF = False Then
        ExtraData.Group = Trim(RS!GroupID)
        ExtraData.Line = Trim(RS!Line)
    End If
    cmbMachine.SetFocus
    Exit Sub
errHandler:
        MsgBox Err.Description
End Sub

'Private Sub cmdAdd_Click()
'    If IsInteger(TxtQty) = False Then
'        MsgBox ("please check the qty of print, Qty>0!"), vbCritical
'        TxtQty.SetFocus
'        Exit Sub
'    End If
'
'    cmdAdd.Enabled = False
'    cmdDelete.Enabled = True
'    cmdSave.Enabled = True
'    cmdCancel.Enabled = True
'    cmdExit.Enabled = True
'    cmdFind.Enabled = True
'    CboCompPN.Enabled = True
'    CboVendorCode.Enabled = True
'    CboDateCode.Enabled = True
'    CboLotCode.Enabled = True
'    TxtQty.Enabled = True
'    CommandType = 1
'    cmdSave.SetFocus
'End Sub

Private Sub cmdCancel_Click()
    CboCompPN.Text = ""
    CboVendorCode.Text = ""
    txtDateCode.Text = ""
    txtLotCode.Text = ""
    TxtQty.Text = ""
    CboDID.Text = ""
    txtInspection = ""
    txtMSD = ""
    CboCompPN.SetFocus
    TxtGroupQty.Text = "1"
    
    Set DGOneByOne.DataSource = Nothing
    DGOneByOne.Refresh
End Sub

Private Sub CmdCommSave_Click()
SaveSetting "SMT", "QSMS", "CommPort", TxtCompPort
SaveSetting "SMT", "QSMS", "Comm", TxtComm

End Sub

Private Sub CmdExcel_Click()
Dim str As String
'Dim Rs As ADODB.Recordset
If Not Rs2.EOF Then
       Call CopyToExcel(Rs2)
    Else
       MsgBox ("No Data"), vbCritical
End If
End Sub

Private Sub cmdFind_Click()
     ''''0062
    Dim strPN As String   ''0064
     
    strPN = CboCompPN
     
    If Trim(CboCompPN) = "" And CboDID <> "" Then
        strPN = "%"
    End If
    
   If AutoDispatchForAnotherBU <> "" Then ''1269
       strSQL = "SELECT '['+QSMS_Server+'].'+QSMS_DB+'.DBO.'AS AutoDispatchForAnotherBU FROM QSMS_SMT_DB WHERE BU IN (Select Value from QSMS_ProConfig Where [key]='AutoDispatchForAnotherBU')"
       Set Rs2 = Conn.Execute(strSQL)
       
       strSQL = "Select * From (Select top 500 DID,CompPN,VendorCode,DateCode,LotCOde,Qty,UID,remainQty,TransDateTime,Line,Side,FirstMachine,UsedFlag From QSMS_DID Where CompPN like '" & Trim(strPN) & "' and DID like '" & Trim(CboDID) & "%' Order by TransDateTime desc UNION ALL Select top 500 DID,CompPN,VendorCode,DateCode,LotCOde,Qty,UID,remainQty,TransDateTime,Line,Side,FirstMachine,UsedFlag From " & Trim(Rs2!AutoDispatchForAnotherBU) & "QSMS_DID Where CompPN like '" & Trim(strPN) & "' and DID like '" & Trim(CboDID) & "%' Order by TransDateTime desc " & _
            " ) A Order by TransDateTime desc "
       Set Rs2 = Conn.Execute(strSQL)
       If Not Rs2.EOF Then Set DG1.DataSource = Rs2
    Else
        strSQL = "Select top 150 DID,CompPN,VendorCode,DateCode,LotCOde,Qty,UID,remainQty,TransDateTime,Line,Side,FirstMachine,UsedFlag From QSMS_DID Where CompPN like '" & Trim(strPN) & "' and DID like '" & Trim(CboDID) & "%' " & _
            " Order by TransDateTime desc "
        Set Rs2 = Conn.Execute(strSQL)
        If Not Rs2.EOF Then Set DG1.DataSource = Rs2
    End If
    cmdDelete.Enabled = True
    cmdSave.Enabled = True
    cmdExcel.Enabled = True
End Sub
Private Sub CmdRefresh_Click()
Call RefreshDg("")
End Sub

Private Sub cmdReprint_Click()
 Dim str As String
Dim RS As ADODB.Recordset

On Error GoTo errhandle

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
    
    'str = "Select * from QSMS_DID Where DID='" & Trim(CboDID) & "'"
    str = "exec [XL_GetDidPrintInfo] @DID='" & Trim(CboDID) & "'" '1242
    Set RS = Conn.Execute(str)
    If RS.EOF Then
       MsgBox "can not find the DID,Please check"
       CboDID.SetFocus
       Exit Sub
     ElseIf UCase(Trim(RS!FirstMachine)) = "DATECODE_CHECK" Then  ''(0079)
       MsgBox "DID DateCode 大于定义DateCode,请检查!"
       CboDID.SetFocus
       Exit Sub
    ElseIf UCase(Trim(RS!FirstMachine)) = "RETURN" Or UCase(Trim(RS!FirstMachine)) = "CALLBACK" Then  ''(0041)
       MsgBox "Can not do reprint, this DID has been " + Trim(RS!FirstMachine) + "ed !"
       CboDID.SetFocus
       Exit Sub
    ''ElseIf UCase(Trim(rs!Qty)) <> Trim(TxtQty) Then  ''(1016)
    ElseIf UCase(Trim(RS!Qty)) <> Replace(Replace(Replace(Replace(TxtQty, " ", ""), vbCr, ""), vbLf, ""), "pcs", "") Then  ''1247
       MsgBox "Reprint Qty <> DID Qty,please check !"
       CboDID.SetFocus
       Exit Sub
    Else
        PrintData.Line = Trim(RS!Line)
        PrintData.Machine = Trim(RS!FirstMachine)
        PrintData.Side = Trim(RS!Side)
        PrintData.DIDWOGROUP = Trim(RS!woGroup) '(0015)
        PrintData.location = Trim(RS!location) '1242
        PrintData.Mark = Trim(RS!Mark) '1255
        PrintData.jobgroup = Trim(RS!jobgroup) '1277
        strVendorCode = Trim(RS!VendorCode) '1277
        If AutoDispatchForAnotherBU <> "" Then
        Set RS = RS.NextRecordset       '''(0052)
        If RS.EOF = False Then
           strAnotherQSMSIP = Trim(RS!AnotherQSMSIP)
         End If
        End If
        
    End If
    
    ''''''''''''Added by Jing 2008.01.03    (0018)'''''''''''1095
    str = "insert into QSMS_Log(system_name,event_no,did,[user_name],returnQty,Trans_Date) values('MaintainDID_Reprint','1','" & Trim(CboDID) & "','" & g_delrightUser & "','0',dbo.FormatDate(getdate(),'yyyymmddhhnnss'))"
    Conn.Execute (str)
    
    Call PrintLabel(Trim(CboDID), TxtQty, strAnotherQSMSIP)
    Exit Sub
errhandle:
    MsgBox Err.Description
End Sub

Private Sub cmdSave_Click()
Dim strSQL As String
Dim str As String
Dim RS As ADODB.Recordset
Dim RS_EMMC As ADODB.Recordset  '(1120)
Dim TempDID As String
Dim TransDate As String
Dim I As Long, RetryCnt As Integer, j As Integer
Dim InsertDIDOk As Boolean
Dim BeginDID As String, EndDID As String, strStep As String
Dim GroupDIDQty As Integer
Dim DispatchType As String '(0044)
Dim ChkVendorCode As Boolean
Dim strAnotherQSMSIP As String
Dim MBPN As String  '(1120)
Dim strBatch As String
Dim str09Code As String '(1254)
Dim Chk09Code As String '(1254)
strBatch = ""
On Error GoTo errHandler:
    lblCount.Caption = ""
    lblStatus = ""
    strStep = "0"
'    cmdAdd.Enabled = True
    cmdFind.Enabled = True
    
    cmdDelete.Enabled = True
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdExit.Enabled = True
    PreDIDPrinted = ""      ' (0021)
    CHKAutoDispatchForAnotherBU = False '''(0052)
    DispatchType = IIf(optExtra = True, "2", IIf(optSpecial = True, "3", IIf(optToWO = True, "5", "0"))) '(0044)
    CboCompPN.Text = Replace(Replace(Replace(CboCompPN.Text, " ", ""), vbCr, ""), vbLf, "")     '--------(7)
     
    TxtQty = Replace(Replace(Replace(Replace(TxtQty, " ", ""), vbCr, ""), vbLf, ""), "pcs", "")  ''1247
    
    ''TxtQty = Trim(TxtQty)   'trim the DID total value's space  (0027)
    '--------------------------------------------------------------------
    
    CboVendorCode = Trim(CboVendorCode)
    If Len(CboVendorCode) > 7 Then      'to check whether the vendor length is less than 8  (0029)
            MsgBox ("Vendor Code must be less than 8"), vbCritical
            Call Warning_Sound
            Call ClearData
            Exit Sub
    End If
    
    If Trim(CboVendorCode) = "" Or Trim(txtDateCode) = "" Or Trim(txtLotCode) = "" Then
        MsgBox ("Vendorcode or datecode or lotcode can't be empty!"), vbCritical
        CboCompPN.Enabled = True
        CboCompPN.SetFocus
        Call Warning_Sound
        Call ClearData
        Exit Sub
    End If
    
    If IsInteger(Trim(CboVendorCode)) = True Then '(0049)
        MsgBox "Vendor code can not be numeric", vbCritical
        Call Warning_Sound
        Call ClearData
        Exit Sub
    End If

    If IsInteger(TxtQty) = False Then
        MsgBox ("please check the qty of Qty, Qty>0!!"), vbCritical
        Call Warning_Sound
        Call ClearData
        Exit Sub
    End If
    
    If CDbl(TxtQty) > 1200000 Then '(1050)(1099)
        MsgBox ("DID Qty can't over 1200000 ！!"), vbCritical '(1050)(1099)
        Call Warning_Sound
        Call ClearData
        Exit Sub
    End If
    
    If IsInteger(TxtGroupQty) = False Then
        MsgBox ("please check the qty of GroupQty, GroupQty>0!!"), vbCritical
        Call Warning_Sound
        Call ClearData
        Exit Sub
    End If
    '--------------------------------------------------------------------
    '(0042)
    GroupDIDQty = GetGroupDIDQty(Trim(CboCompPN))
    If GroupDIDQty > 0 And CInt(TxtGroupQty) > GroupDIDQty Then
        MsgBox "Group qty can not more then " & GroupDIDQty, vbCritical
        Call Warning_Sound
        Call ClearData
        Exit Sub
    End If
    
    
    If chkQty = "" Then
        If CInt(Trim(TxtGroupQty)) > 10 Then
            MsgBox ("Max reel Qty must be less than 10 !"), vbCritical
            Call Warning_Sound
            Call ClearData
            Exit Sub
        End If
    Else ''(0061)
        If ChkOldDIDLabelQty <> "" And opOldLabel.Value = True Then
            If CInt(Trim(TxtGroupQty)) > CInt(ChkOldDIDLabelQty) Then
                MsgBox ("Max reel Qty must be less than " & ChkOldDIDLabelQty & " for old label !"), vbCritical
                Call Warning_Sound
                Call ClearData
                Exit Sub
            End If
        Else
            If CInt(Trim(TxtGroupQty)) > CInt(chkQty) Then
                MsgBox ("Max reel Qty must be less than " & chkQty & " !"), vbCritical
                Call Warning_Sound
                Call ClearData
                Exit Sub
            End If
        End If
    End If
    
    '(1254) begin
     strSQL = "select 0 from QSMS_ProConfig where Station = 'QSMS' and [Key] = 'Check09Code' and Value = 'Y'"
     Set RS = Conn.Execute(strSQL)
     If Not RS.EOF Then
        Chk09Code = "Y"
     End If
     '(1254) end
     
     '(1120) begin
    strSQL = "select 0 from QSMS_ProConfig where Station = 'QSMS' and [Key] = 'CheckEMMCImageVersion' and Value = 'Y'"
    Set RS_EMMC = Conn.Execute(strSQL)
    If Not RS_EMMC.EOF Then
        strSQL = "select 0 from EMMC where EMMCPN ='" & CboCompPN.Text & "'"
        Set RS_EMMC = Conn.Execute(strSQL)
        If Not RS_EMMC.EOF Then
           If Trim(txtImgVersion) = "" Then
                MsgBox ("This CompPN is EMMC,please scan ImageVersion!"), vbCritical
                txtImgVersion.SetFocus
                Call Warning_Sound
                Call ClearData
                Exit Sub
           End If
           If optToWO = False Then
                MsgBox ("This CompPN is EMMC,please choose ToWO dispatchtype!"), vbCritical
                Call Warning_Sound
                Call ClearData
                Exit Sub
           End If
           strSQL = "select PN from SAP_WO_LIST where WO='" & Trim(cmbWO) & "' "
           Set RS_EMMC = Conn.Execute(strSQL)
           If Not RS_EMMC.EOF Then
              MBPN = Trim(RS_EMMC!PN)
              strSQL = "select 0 from EMMC where MBPN='" & MBPN & "' and ImageVersion= '" & Trim(txtImgVersion) & "'"
              Set RS_EMMC = Conn.Execute(strSQL)
              If RS_EMMC.EOF Then
                 MsgBox ("This imageversion does not match the MBPN!"), vbCritical
                 txtImgVersion.SetFocus
                 Call Warning_Sound
                 Call ClearData
                 Exit Sub
              End If
           Else
              MsgBox ("Please choose WO!"), vbCritical
              Call Warning_Sound
              Call ClearData
              Exit Sub
           End If
        End If
    End If
    '(1120) end
    
    If optExtra = True Or optSpecial = True Or optToWO = True Then
        If (IsInteger(txtExtraQty) = False Or CLng(txtExtraQty) > CLng(TxtQty)) Then '0048 ''(1094) CInt change to CLng
            MsgBox "Extra dispatch qty is not numeric or extra/special dispatch qty bigger then did qty!", vbCritical
            Call Warning_Sound
            Call ClearData
            Exit Sub
        End If
        If cmbWO = "" Or cmbMachine = "" Or cmbSide = "" Or cmbSlot = "" Or cmbLR = "" Then
            MsgBox "Please check if all extra/sepcial dispatch data were correct", vbCritical
            Call Warning_Sound
            Call ClearData
            Exit Sub
        End If
        If ExtraData.Group = "" Or ExtraData.Line = "" Then
            MsgBox "Can not get wo group or line information!", vbCritical
            Call Warning_Sound
            Call ClearData
            Exit Sub
        End If
        ExtraData.Qty = Trim(txtExtraQty)
        ExtraData.Machine = Trim(cmbMachine)
        ExtraData.Side = Trim(cmbSide)
        ExtraData.Slot = Trim(cmbSlot)
        ExtraData.LR = Trim(cmbLR)
    End If

    '******************************
    '****add by jeanson 2007/09/03
    strErrMessage = ""
    strErrMessage = FunPartNumberCheck(CboCompPN.Text)
    If strErrMessage <> "PASS" Then
        MsgBox strErrMessage
        CboCompPN.SetFocus
        Call Warning_Sound
        Call ClearData
        Exit Sub
    End If
    '******************************

    
    If ChkAVL(Trim(CboCompPN), Trim(CboVendorCode)) = False Then
        CboVendorCode.Text = ""
        Exit Sub
    End If
    
    '**Denver       2010.03.19     Add IC comp check function  （0068）
    If IC_CompChk = "Y" Then
        If IC_CompNeedBurn(Trim(CboCompPN)) = True Then
            Exit Sub
        End If
    End If
    
    ''''''''''''''''''''''''''Add by Jing  20071126   (0006) ''''''''''''''''''''''''''''
    strSQL = "select DID from QSMS_DID where did='" & CboDID.Text & "'"
    Set RS = Conn.Execute(strSQL)
    If Not RS.EOF Then
        MsgBox ("This DID exists !")
        Exit Sub
    End If
    '----------------------add by Sandy 20080326-(0031)----------------
    If InStr(1, Trim(CboVendorCode), "'") <> 0 Then
        MsgBox ("Please check VendorCode,Don't allow to include '")
        Call Warning_Sound
        Call ClearData
        Exit Sub
    End If
    If InStr(1, Trim(txtDateCode), "'") <> 0 Then
        MsgBox ("Please check DateCode,Don't allow to include '")
        Call Warning_Sound
        Call ClearData
        Exit Sub
    End If
    If InStr(1, Trim(txtLotCode), "'") <> 0 Then
        MsgBox ("Please check LotCode,Don't allow to include '")
        Call Warning_Sound
        Call ClearData
        Exit Sub
    End If
    
    If BatchControl = "Y" Then
        strSQL = "Exec QSMS_ProcessCompBatch @CompPN='" & Trim(CboCompPN) & "',@Type='CHECKBATCH'"
        Set RS = Conn.Execute(strSQL)
        If Not RS.EOF Then
            If UCase(RS!result) = "NEEDBATCH" Then
                strBatch = UCase(InputBox("Please Input the Batch of this CompPN", "Input Batch"))
                strSQL = "Exec QSMS_ProcessCompBatch @CompPN='" & Trim(CboCompPN) & "',@Type='CHECKBATCHVALUE',@Batch='" & Trim(strBatch) & "'"
                Set RS = Conn.Execute(strSQL)
                If UCase(RS!result) = "CHECKFAIL" Then
                    MsgBox ("Input Batch error, the Batch value must match with be defined! ")
                    Call Warning_Sound
                    Call ClearData
                    Exit Sub
                End If
            End If
        End If
    End If
    
    strStep = "1"
    If UCase(Trim(StrBU)) = "AS" Then           ''(0045)(0051)
        If CheckDataCode(txtInspection, CboVendorCode, txtDateCode) = False Then
            Exit Sub
        End If
'        If ChkDateCodeSpecial(Trim(CboVendorCode), Trim(CboCompPN), Trim(txtDateCode)) = False Then
'            Exit Sub
'        End If
    Else
        txtInspection = ""
    End If
    
    If UCase(Trim(ChkDateCode)) = "Y" Then ''1222
       If ChkDateCodeSpecial(Trim(CboVendorCode), Trim(CboCompPN), Trim(txtDateCode)) = False Then
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
    

   If Chk09Code = "Y" Then     ''1254
      strSQL = "select 0 from ConsignPN_HUA where [TYPE] = 'HUACustomerPN' and CompPN='" & CboCompPN.Text & "'"
      Set RS = Conn.Execute(strSQL)
      If Not RS.EOF Then
          str09Code = InputBox("Please Input 09Code:     ", "Input 09Code")
          If Len(str09Code) < 20 Then
            str09Code = ""
            MsgBox ("The Lenghth of 09Code Should Be Greater Than 20"), vbCritical
            Call Warning_Sound
            Call ClearData
            Exit Sub
          End If
      End If
   End If
    strStep = "2"
    '----------------------add by Sandy 20080326-(0031)----------------
    strSQL = "select getdate()"
    Set RS = Conn.Execute(strSQL)
    TransDate = Format(RS(0), "YYYYMMDDHHNNSS") '1101
    Select Case CommandType
        Case 1
            strStep = "3"
            LockTheForm (False)
            For I = 1 To CLng(Trim(TxtGroupQty))
                RetryCnt = 0
                InsertDIDOk = False
                'If there are more than one person add same component DID, the DID# may conflict
                On Error GoTo Retry
                Do While Not InsertDIDOk And RetryCnt < 10
                    '**Sandy        2007.12.2  update to drive out the blank space in DID.---------(0007)
                    TempDID = Trim(GetDID(Trim(CboCompPN), TransDate))  '' --------------(0023)
                    If BeginDID = "" Then BeginDID = TempDID
                    EndDID = TempDID
                    ''1191
                    If ReturnDID = True And CheckMSDCallBack = "Y" Then
                        strSQL = "exec [PD_MSD_LinkDIDAuto] @DID='" & TempDID & "',@ReturnDID='" & OldReturnDID & "',@CompPN='" & CboCompPN & "',@Inherit_WO='',@ReturnFlag='Y',@UID='" & g_userName & "' "
                        Set RS = Conn.Execute(strSQL)
                        
                        If UCase(RS!result) = "CHECKFAIL" Then ''1200
                            MsgBox ("Message: " & RS!ErrDesc), vbCritical
                            Call Warning_Sound
                            Call ClearData
                            Exit Sub
                        End If
                        
                    End If
                    
                    strStep = "4"
                    '---------------------0003--IIf(optExtra = True, "2", IIf(optSpecial = True, "3", "0"))-----------Add Factory (0033)--- '(0044)
                    If Chk09Code = "Y" And str09Code <> "" Then ''1254
                    strSQL = "exec XL_DIDAutoDispatch " & _
                             "'" & TempDID & "'," & _
                             "'" & Trim(CboCompPN) & "'," & _
                             "" & TxtQty & "," & _
                             "" & TxtQty & "," & _
                             "'" & Left(Trim(CboVendorCode), 7) & "'," & _
                             "'" & Trim(txtDateCode) & "'," & _
                             "'" & Trim(txtLotCode) & "'," & _
                             "'" & Trim(txtInspection) & "'," & _
                             "''," & _
                             "'" & g_userName & "'," & _
                             "'" & DispatchType & "'," & _
                             "''," & _
                             "'" & ExtraData.Group & "'," & _
                             "'" & ExtraData.WO & "'," & _
                             "'" & ExtraData.Line & "'," & _
                             "'" & ExtraData.Side & "'," & _
                             "'" & ExtraData.Machine & "'," & _
                             "'" & ExtraData.Slot & "'," & _
                             "'" & ExtraData.LR & "'," & _
                             "'" & ExtraData.Qty & "', " & _
                             "'" & Trim(Factory) & "', " & _
                             "'" & Trim(txtMSD) & "', " & _
                             "'" & Trim(OldReturnDID) & "'," & _
                             "'" & Trim(str09Code) & "'"
                    Else
                    strSQL = "exec XL_DIDAutoDispatch " & _
                             "'" & TempDID & "'," & _
                             "'" & Trim(CboCompPN) & "'," & _
                             "" & TxtQty & "," & _
                             "" & TxtQty & "," & _
                             "'" & Left(Trim(CboVendorCode), 7) & "'," & _
                             "'" & Trim(txtDateCode) & "'," & _
                             "'" & Trim(txtLotCode) & "'," & _
                             "'" & Trim(txtInspection) & "'," & _
                             "''," & _
                             "'" & g_userName & "'," & _
                             "'" & DispatchType & "'," & _
                             "''," & _
                             "'" & ExtraData.Group & "'," & _
                             "'" & ExtraData.WO & "'," & _
                             "'" & ExtraData.Line & "'," & _
                             "'" & ExtraData.Side & "'," & _
                             "'" & ExtraData.Machine & "'," & _
                             "'" & ExtraData.Slot & "'," & _
                             "'" & ExtraData.LR & "'," & _
                             "'" & ExtraData.Qty & "', " & _
                             "'" & Trim(Factory) & "', " & _
                             "'" & Trim(txtMSD) & "', " & _
                             "'" & Trim(OldReturnDID) & "'"    ''1233
                    End If
                    Set RS = Conn.Execute(strSQL)
                    lblStatus = Trim(RS!ErrDesc)
                    strStep = "5"
                    ''''''update by Jing (0034)''''''
                    If Trim(RS!result) <> "1" Then
                        Call RefreshDg("", BeginDID, EndDID)
                        strStep = "6"
                        LockTheForm (True) '(0036)
                        strStep = "7"
                        Call cmdCancel_Click
                        cmdExcel.Enabled = False
                        Exit Sub
                    End If

                    ''''NB2/3物料不分,导致前台产生的DID不一定就是系统真正使用的DID,仅作用于Normal Dispatch  ''''(0052)
                    If DispatchType = "0" Then
                        If Trim(RS!DID) <> Trim(TempDID) Then
                            TempDID = Trim(RS!DID)
                            CHKAutoDispatchForAnotherBU = True
                        End If
                    End If
                    
                    ''''StrSQL = "exec [XL_GetDidPrintInfo] @DID='" & Trim(TempDID) & "'"  ' (0039)
                    ''''(0060)
                    strSQL = "exec [XL_GetDidPrintInfo] @DID='" & Trim(TempDID) & "',@Factory='" & Trim(Factory) & "'"
                    Set RS = Conn.Execute(strSQL)
                    If RS.EOF Then
                       MsgBox "can not find the DID,Please check"
                       CboDID.SetFocus
                       Call Warning_Sound
                       Call ClearData
                       Exit Sub
                    Else
                        PrintData.Line = Trim(RS!Line)
                        PrintData.Machine = Trim(RS!FirstMachine)
                        PrintData.Side = Trim(RS!Side)
                        PrintData.DIDWOGROUP = Trim(RS!woGroup) '(0015)
                        PrintData.location = Trim(RS!location) '1242
                        PrintData.Mark = Trim(RS!Mark)  '1255
                        PrintData.jobgroup = Trim(RS!jobgroup) '1277
                        WOType = Trim(RS!WOType) ' (0039)
                        
                        If CHKAutoDispatchForAnotherBU = True Then
                            Set RS = RS.NextRecordset       '''(0052)
                            If RS.EOF = False Then
                                strAnotherQSMSIP = Trim(RS!AnotherQSMSIP)
                            End If
                        End If
                    End If
                    InsertDIDOk = True
Retry:
                    RetryCnt = RetryCnt + 1
                    DoEvents
                Loop
                
                strStep = "8"
                CboDID = TempDID
                If PreDIDPrinted <> CboDID And InsertDIDOk = True Then    '(0021)  'add Insertdidok=true (0040)
                    Call PrintLabel(TempDID, TxtQty, strAnotherQSMSIP)
                End If
                
                PreDIDPrinted = CboDID
                
                If BatchControl = "Y" And strBatch <> "" Then
                    strSQL = "Exec QSMS_ProcessCompBatch @DID='" & Trim(TempDID) & "',@Batch='" & Trim(strBatch) & "',@UID='" & g_userName & "',@Type='SAVEBATCH'"
                    Set RS = Conn.Execute(strSQL)
                End If
                
                If DIDAutoOpen = "Y" Then       '(1268)
                    strSQL = "Exec QSMS_DIDAutoOpen @DID='" & Trim(TempDID) & "'"
                    Set RS = Conn.Execute(strSQL)
                End If
                '''''if SATO printer will not delay time    (0003)

                '''''Updated by Jing 2007.12.26 (0014)

                If OptZebra.Value = True And (chkDelay.Value = Checked) Then
                    Call Sleep(900)
                End If
                
                Call Sleep(100)     '''Added by Jing 2008.01.02    (0017)'''
                
            Next I
            
            Call cmdFind_Click
           
            LockTheForm (True)
            cmdExcel.Enabled = False
    End Select
    strStep = "9"
    Call RefreshDg("", BeginDID, EndDID)     ''-----(0023)
 
    CommandType = 0
    TxtGroupQty = 1
    Call cmdCancel_Click
    Exit Sub
errHandler:
    MsgBox Err.Description & strStep
End Sub
Private Sub CmdVendorCode_Click()
Dim str As String
Dim RS As ADODB.Recordset

str = "select * from QSMS_DID where vendorcode = '" & CboVendorCode & "'order by CompPN"

Set RS = Conn.Execute(str)
 If Not RS.EOF Then
       Call CopyToExcel(RS)
    Else
       MsgBox ("No Data"), vbCritical
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DG1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    On Error Resume Next
    ''''''''''''''''''''''''''''''''''''(0024)
    With DG1
        CboDID = ""
        'CboCompPN = .Columns(1).Value    '''0065
        TxtGroupQty = ""
        CboVendorCode = ""
        txtDateCode = ""
        txtLotCode = ""
        TxtQty = ""
        PrintData.Side = ""
        PrintData.Line = ""
        PrintData.Machine = ""
    End With
    
    cmdDelete.Enabled = True
    cmdCancel.Enabled = True
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Function ChkInspectionNo(InspectionNo As String, DateCode As String) As Boolean
Dim strSQL As String
Dim RS As ADODB.Recordset
    ChkInspectionNo = True
    If Trim(InspectionNo) <> "" Then
        strSQL = "Exec QSMS_ChkInspectionNo '" & Trim(InspectionNo) & "','" & Trim(DateCode) & "'"
        Set RS = Conn.Execute(strSQL)
        If UCase(Trim(RS.Fields("Result"))) <> "PASS" Then
            MsgBox RS.Fields("iMessage"), vbCritical, "ErrMessage"
             ChkInspectionNo = False
        End If
    End If
End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
'''''(0047)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

Dim str As String
Dim RS As ADODB.Recordset

If UCase(Trim(StrBU)) = "AS" Then
    labInpectionNo.Visible = True
    txtInspection.Visible = True
End If

If UCase(ScanMSD) = "Y" Then
    labelMSD.Visible = True
    txtMSD.Visible = True
End If


''(0076)
If ChkOneByOneMaterial = "N" Then
    lblOneByOne.Visible = False
    flexGridDemandMaterial.Visible = False
'    DG1.Top = DGOneByOne.Top
'    DG1.Height = DG1.Height + DGOneByOne.Height
    DG1.Top = flexGridDemandMaterial.Top
    DG1.Height = DG1.Height + flexGridDemandMaterial.Height
Else
    Call InitFlex(flexGridDemandMaterial)
End If

''**Sandy        2007.12.13     set default Printer according to each BU whether have FUJITrax ----(0009)
'str = "SELECT FUJI_Server as FUJI FROM QSMS_SMT_DB WHERE BU IN (SELECT SITE FROM SITE)"
'Set rs = Conn.Execute(str)
'If Trim(rs!FUJI) <> "" Then OptSATO.Value = True
''(0042)
'TxtCompPort.Text = GetSetting("SMT", "QSMS", "CommPort", "1")
'TxtComm.Text = GetSetting("SMT", "QSMS", "Comm", "9600,N,8,1")

'20101115 Maggie Save Printer setting in local Registry (1019)
Call GetPrinterSetting(frmMaintainDIDAutoDispatch)

''(1080) printlabel中判断是否为zebra的移到此
If OptZebra.Value = True Then
    isZebra = True
Else
    isZebra = False
End If

Call RefreshGroupDIDQty '(0042)



Call Hook(CboCompPN.hWnd)  '(1100)
Call Hook(CboVendorCode.hWnd) '(1100)
Call Hook(txtDateCode.hWnd)   '启动钩子
Call Hook(txtLotCode.hWnd)
Call Hook(TxtQty.hWnd)

strCheckScaner = ReadIniFile("QSMS", "DIDScan", App.Path & "\set.ini")

End Sub

Private Function RefreshCompPN()
Dim str As String
Dim RS As ADODB.Recordset
str = "select CompPn from QSMS_RackID order by CompPN"
Set RS = Conn.Execute(str)
CboCompPN.Clear
While Not RS.EOF
      CboCompPN.AddItem Trim(RS!compPN) & vbNullString
      
      RS.MoveNext
Wend

End Function

Private Function RefreshDg(ByVal compPN As String, Optional ByVal BeginDID As String, Optional ByVal EndDID As String)
Dim str As String
On Err GoTo errHandler:
'not to display the DID which Qty=0  (0043)
If BeginDID = "" Or EndDID = "" Then
    str = "select top 150 DID,CompPN,VendorCode,DateCode,LotCode,Qty,RemainQty,UID,TransDateTime,Line,Side,FirstMachine from QSMS_DID where CompPN like '" & compPN & "%' and Qty<>0 order by TransDateTime desc"
Else
    str = "select TOP 150 DID,CompPN,VendorCode,DateCode,LotCode,Qty,RemainQty,UID,TransDateTime,Line,Side,FirstMachine from QSMS_DID where did between '" & BeginDID & "' and '" & EndDID & "' and DIDHostName=left(Host_Name(),20) and Qty<>0 order by TransDateTime desc"
End If
Set Rs2 = Conn.Execute(str)
If Rs2.EOF = False Then
    Set DG1.DataSource = Rs2   '''(0023)
    'If lblStatus.Caption = "" Then lblStatus.Caption = "DID Count:  " + CStr(Rs2.RecordCount)
    lblCount.Caption = "DIDQty: " + CStr(Rs2.RecordCount)
Else
    lblCount.Caption = "DIDQty: 0 "
End If
DG1.Refresh
Exit Function
errHandler:
    MsgBox Err.Description
End Function

'Private Function PrintLabel(strDID As String, strQty As String, Optional strAnotherQSMSIP As String) As String
'If OptComp.Value = True Then
   'Call PrintLabelCompPort(strDID, strQty, strAnotherQSMSIP)
'End If
'If OptPrint.Value = True Then
   'Call PrintLabelPrintPort(strDID, strQty, strAnotherQSMSIP)
'End If
'End Function
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
    Dim TmpDIDText As String
    Dim strMainPN As String ''(1226)
    Dim strREELWIDTH As String
  
    
    On Error GoTo errHandler
'    strDay = Format(Now, "YYYY/MM/DD")
    ''(1016)
    strSQL = "select getdate()"
    Set rsTime = Conn.Execute(strSQL)
    strDay = Format(rsTime(0), "YYMMDDHHNNSS")    '1101
    If StrBU = "NB4" Then                         '1148
        strDay = Format(rsTime(0), "YYYYMMDD")    '
    End If
    If CHKAutoDispatchForAnotherBU = True And strAnotherQSMSIP <> "" Then    '''(0052)
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
    
    ''''''added by Jing 2008.04.05  (0032)''''''
    If opNewLabel.Value = True Then
        Dim x As Integer
        For x = 0 To 4
            WO(x) = ""
            Model(x) = ""
            Machine(x) = ""
            Slot(x) = ""          '(1086)
            MachineUnit(x) = ""
            Work_Order(x) = ""
            DIDType(x) = ""
            ISCYL(x) = ""
            SeqID(x) = ""           '(1147)
            VenderCode(x) = ""      '1223
            LR(x) = ""               '1223

        Next x
           
        tmpStr = "Exec QSMS_GetDIDPrintInfo @DID='" & Trim(strDID) & "',@AnotherQSMSIP='" & Trim(strAnotherQSMSIP) & "',@PrinterType='" & Trim(PrinterType) & "',@PrintDpm='" & Trim(PrintDpm) & "'"        '''(0056)
        Set tmpRS = Conn.Execute(tmpStr)
        If tmpRS.EOF = False Then
            If StrBU = "NB5" Then   ''(1226)
                strREELWIDTH = tmpRS("ReelWidth")
                strMainPN = tmpRS("MainPN")
            End If
            Dim I As Integer, j As Integer, ff As Integer
            j = tmpRS.RecordCount
            If j > 5 Then j = 5
            For I = 0 To j - 1
                WO(I) = tmpRS("Machine") + " " + tmpRS("Slot") + "-" + tmpRS("LR") '0046
                Model(I) = tmpRS("model")
                Machine(I) = Mid(tmpRS("Machine"), 2, 1) + "-" + Mid(tmpRS("Slot"), 1, 1) + "-" + Mid(tmpRS("Machine"), 6, 1)
                Slot(I) = tmpRS("Slot") + "-" + tmpRS("LR")                        '(1086)
                MachineUnit(I) = tmpRS("Machine") + "-" + Mid(tmpRS("Slot"), 1, 1)
                Work_Order(I) = tmpRS("Work_Order")         ''1093
                DIDType(I) = tmpRS("DIDType")   '''(1105)
                
                MachineCH(I) = tmpRS("MachineCH")         ''1044
                SideCH(I) = tmpRS("SideCH")         ''1044
                LRCH(I) = tmpRS("LRCH")         ''1044
                SlotCH(I) = tmpRS("Slot")         ''1044
                PN(I) = tmpRS("PN")
                
                If PrintedSeqID = "Y" Then
                    SeqID(I) = tmpRS("SeqID")           '(1147)
                End If
                If PrintedVenderCode = "Y" Then    ''1223
                    VenderCode(I) = tmpRS("VenderCode")
                    LR(I) = tmpRS("SLR")
                End If
                For ff = 0 To tmpRS.Fields.Count - 1    ''(1109)
                    If UCase(tmpRS.Fields(ff).Name) = "ISCYL" Then
                        ISCYL(I) = tmpRS("ISCYL")
                    End If
                Next ff
                tmpRS.MoveNext
            Next I
        End If
    End If
  ''(1080) 20111219 repace by 1080
'    If OptZebra.Value = True Then
'        isZebra = True
'
'        ''''''updated by Jing   (0032)''''''
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
    ''(1080)
'   LabelFile = GetDIDLabelFile(frmMaintainDIDAutoDispatch, IIf(opOldLabel.Value = True, "OLD", "NEW"))
'
'''    LabelFile = Settings.DIDLabelPath & "Zebra_200_new_test.txt"  ''(only for test)
'    If Dir(LabelFile) = vbNullString Then
'        ''''''Added by Jing 2008.01.10  (0019)''''''
'        MsgBox ("Can not find label file:" & LabelFile & "!"), vbCritical
'        PrintLabel = "PRN_FileNoExist"
'        Exit Function
'    End If
    
    ''''for Com Port
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
        ''********************************updated by jing 20071024 (0002) ***********
     ''''  1196
    If InStr(tmpPrintStr, "<DID_2D>") > 0 Then
      
         If isZebra Then
             tmpDID = Trim(Replace(strDID, "^", "_5E"))
         End If
        tmpPrintStr = Replace(tmpPrintStr, "<DID_2D>", tmpDID)
     End If
     
     'for text ^, must be use ^FH_ and the use _5E (the ascii of ^)
     If InStr(tmpPrintStr, "<DID_TEXT>") > 0 Then
         ''********************************updated by jing 20071024 (0002) ***********
        If CheckBurnDID = "Y" Then    ''''''(1192)
            tmpStr = "Exec QSMS_GetFWRev @DID='" & Trim(strDID) & "',@AnotherQSMSIP='" & Trim(strAnotherQSMSIP) & "'"        '''(0056)
            Set tmpRS = Conn.Execute(tmpStr)
            TmpDIDText = Trim(tmpRS!DIDtext)
            
            If isZebra Then
                tmpDID = Replace(TmpDIDText, "^", "_5E")
            Else
                tmpDID = Trim(TmpDIDText)
            End If
        Else
            
            If isZebra Then
                tmpDID = Replace(strDID, "^", "_5E")
            End If
        End If
        tmpPrintStr = Replace(tmpPrintStr, "<DID_TEXT>", tmpDID)
        
     End If
     
    
     tmpPrintStr = Replace(tmpPrintStr, "<UID>", UID)
     tmpPrintStr = Replace(tmpPrintStr, "<DATE>", strDay)
     tmpPrintStr = Replace(tmpPrintStr, "<QTY>", strQty)
     tmpPrintStr = Replace(tmpPrintStr, "<Location>", PrintData.location) '1242
     tmpPrintStr = Replace(tmpPrintStr, "<MARK>", PrintData.Mark) '1255
     tmpPrintStr = Replace(tmpPrintStr, "<JOBGROUP>", PrintData.jobgroup) '1277
     tmpPrintStr = Replace(tmpPrintStr, "<VendorCode>", strVendorCode)     ''1282
     ''''''updated by Jing (0032)''''''
     If opNewLabel.Value = True Then
        tmpPrintStr = Replace(tmpPrintStr, "<BU>", PrintData.Line) '(0058)
     Else
        tmpPrintStr = Replace(tmpPrintStr, "<LINE>", PrintData.Line) '(0037) '(0058)
     End If
     
'     If PrintData.Line <> Left(PrintData.machine, 1) Then '(0058)
'        Conn.Execute ("Insert into QSMS_Error_log(Appname,SubFunction,SubID,Col1,Col2,Col3,DetailDesc,TransDateTime) values(" & _
'                     "'QSMS','PrintDID','Log','" & tmpDID & "','" & PrintData.Line & "','" & PrintData.machine & "'," & _
'                     "'Line and Machine did not match',dbo.formatdate(getdate(),'yyyymmddhhnnss'))")
'     End If
     tmpPrintStr = Replace(tmpPrintStr, "<SIDE>", PrintData.Side)
     tmpPrintStr = Replace(tmpPrintStr, "<MACHINE>", PrintData.Machine)
     tmpPrintStr = Replace(tmpPrintStr, "<DIDWOGROUP>", PrintData.DIDWOGROUP) '(0015)
     tmpPrintStr = Replace(tmpPrintStr, "<WOTYPE>", WOType)  ' (0039)
     tmpPrintStr = Replace(tmpPrintStr, "<MainPN>", strMainPN)  ''(1226)
     tmpPrintStr = Replace(tmpPrintStr, "<REELWIDTH>", strREELWIDTH)
     
    ''''''added by Jing 2008.04.05  (0032)''''''
    If opNewLabel.Value = True Then
        tmpPrintStr = Replace(tmpPrintStr, "<WO1>", WO(0))
        tmpPrintStr = Replace(tmpPrintStr, "<WO2>", WO(1))
        tmpPrintStr = Replace(tmpPrintStr, "<WO3>", WO(2))
        tmpPrintStr = Replace(tmpPrintStr, "<WO4>", WO(3))
        tmpPrintStr = Replace(tmpPrintStr, "<WO5>", WO(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINE1>", Machine(0))
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
        
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINEUNIT1>", MachineUnit(0))
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINEUNIT2>", MachineUnit(1))
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINEUNIT3>", MachineUnit(2))
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINEUNIT4>", MachineUnit(3))
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINEUNIT5>", MachineUnit(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<Work_Order1>", Work_Order(0))  ''''1093
        tmpPrintStr = Replace(tmpPrintStr, "<Work_Order2>", Work_Order(1))
        tmpPrintStr = Replace(tmpPrintStr, "<Work_Order3>", Work_Order(2))
        tmpPrintStr = Replace(tmpPrintStr, "<Work_Order4>", Work_Order(3))
        tmpPrintStr = Replace(tmpPrintStr, "<Work_Order5>", Work_Order(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<DIDType1>", DIDType(0)) ''''1105
        tmpPrintStr = Replace(tmpPrintStr, "<DIDType2>", DIDType(1))
        tmpPrintStr = Replace(tmpPrintStr, "<DIDType3>", DIDType(2))
        tmpPrintStr = Replace(tmpPrintStr, "<DIDType4>", DIDType(3))
        tmpPrintStr = Replace(tmpPrintStr, "<DIDType5>", DIDType(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<CYL1>", ISCYL(0))  ''(1109）
        tmpPrintStr = Replace(tmpPrintStr, "<CYL2>", ISCYL(1))
        tmpPrintStr = Replace(tmpPrintStr, "<CYL3>", ISCYL(2))
        tmpPrintStr = Replace(tmpPrintStr, "<CYL4>", ISCYL(3))
        tmpPrintStr = Replace(tmpPrintStr, "<CYL5>", ISCYL(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<COUNT1>", SeqID(0))  '(1147)
        tmpPrintStr = Replace(tmpPrintStr, "<COUNT2>", SeqID(1))
        tmpPrintStr = Replace(tmpPrintStr, "<COUNT3>", SeqID(2))
        tmpPrintStr = Replace(tmpPrintStr, "<COUNT4>", SeqID(3))
        tmpPrintStr = Replace(tmpPrintStr, "<COUNT5>", SeqID(4))
        
        tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE1>", VenderCode(0))   '1223
        tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE2>", VenderCode(1))
        tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE3>", VenderCode(2))
        tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE4>", VenderCode(3))
        tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE5>", VenderCode(4))
         
        tmpPrintStr = Replace(tmpPrintStr, "<LR1>", LR(0))                    '1223
        tmpPrintStr = Replace(tmpPrintStr, "<LR2>", LR(1))
        tmpPrintStr = Replace(tmpPrintStr, "<LR3>", LR(2))
        tmpPrintStr = Replace(tmpPrintStr, "<LR4>", LR(3))
        tmpPrintStr = Replace(tmpPrintStr, "<LR5>", LR(4))
        
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
        
        tmpPrintStr = Replace(tmpPrintStr, "<PN1>", PN(0))                    '1044
        tmpPrintStr = Replace(tmpPrintStr, "<PN2>", PN(1))
        tmpPrintStr = Replace(tmpPrintStr, "<PN3>", PN(2))
        tmpPrintStr = Replace(tmpPrintStr, "<PN4>", PN(3))
        tmpPrintStr = Replace(tmpPrintStr, "<PN5>", PN(4))

        
        'MachineUnit
        '(1063)
        'tmpPrintStr = Replace(tmpPrintStr, "<MACHINETYPE>", Mid(PrintData.machine, Len(PrintData.machine) - 3, 3))
        'tmpPrintStr = Replace(tmpPrintStr, "<MACHINECODE>", Right(PrintData.machine, 1))
        
        Dim tmpMachine As String
        tmpMachine = ""
        If InStr(WO(0), " ") > 1 Then
            tmpMachine = Mid(WO(0), 1, InStr(WO(0), " ") - 1)
            tmpPrintStr = Replace(tmpPrintStr, "<SLOT>", Mid(WO(0), InStr(WO(0), " ") + 1, Len(WO(0)) - InStr(WO(0), " ")))
        End If
        
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINETYPE>", Mid(tmpMachine, Len(tmpMachine) - 3, 3))
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINECODE>", Right(tmpMachine, 1))
    End If
    
    Select Case Trim(tmpPrintStr)
       Case vbNullString
       Case Else
            If OptComp.Value = True Then
                For M = 1 To Len(tmpPrintStr) Step 100
                    MSComm.Output = Mid(tmpPrintStr, M, 100)
                    'Debug.Print Mid(hString, m, 50)
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
    Call OK_Sound
    
    If DIDType(0) = "*Assigned*" Or DIDType(0) = "*Assigned*Residual*" Then  '''(1106)
        Call Assigned_Sound
    End If
    Exit Function
errHandler:
    MsgBox Err.Description
    If MSComm.PortOpen = True Then
        MSComm.PortOpen = False
    End If
End Function

Private Function PrintLabelCompPort(strDID As String, strQty As String, Optional strAnotherQSMSIP As String) As String
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

        On Error GoTo errHandler
        
        tmpStr = "select getdate()"
        Set rsTime = Conn.Execute(tmpStr)
        strDay = Format(rsTime(0), "YYMMDDHHNNSS") '1101
        
        If CHKAutoDispatchForAnotherBU = True And strAnotherQSMSIP <> "" Then    '''(0052)
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
'    If Dir(LabelFile) = vbNullString Then
'        ''''''Added by Jing 2008.01.10  (0019)''''''
'        MsgBox ("Can not find label file:" & LabelFile & "!"), vbCritical
'        PrintLabel = "PRN_FileNoExist"
'        Exit Function
'    End If
        ''''''added by Jing 2008.04.05  (0032)''''''
        If opNewLabel.Value = True Then
            Dim x As Integer
            For x = 0 To 4
                WO(x) = ""
                Model(x) = ""
                Work_Order(x) = ""
                DIDType(x) = ""
                
            Next x
            
'            tmpStr = "select distinct a.Machine,a.Slot,A.LR,substring(b.PN,3,3) as Model from QSMS_Dispatch a,sap_wo_list b where a.did='" & Trim(strDID) & "' and a.work_order=b.wo"
'            tmpStr = "Exec XL_GetDidDispatchInfo @DID='" & Trim(strDID) & "'"       '''(0052)
'            If CHKAutoDispatchForAnotherBU = True And strAnotherQSMSDB <> "" Then    '''(0052)
'                tmpStr = "select distinct a.Machine,a.Slot,A.LR,substring(b.PN,3,3) as Model from " & Trim(strAnotherQSMSDB) & "QSMS_Dispatch a," & strAnotherQSMSDB & "sap_wo_list b where a.did='" & Trim(strDID) & "' and a.work_order=b.wo"
'            Else
'                tmpStr = "select distinct a.Machine,a.Slot,A.LR,substring(b.PN,3,3) as Model from QSMS_Dispatch a,sap_wo_list b where a.did='" & Trim(strDID) & "' and a.work_order=b.wo"
'            End If
            tmpStr = "Exec QSMS_GetDIDPrintInfo @DID='" & Trim(strDID) & "',@AnotherQSMSIP='" & Trim(strAnotherQSMSIP) & "',@PrinterType='" & Trim(PrinterType) & "',@PrintDpm='" & Trim(PrintDpm) & "'"        '''(0056)
            Set tmpRS = Conn.Execute(tmpStr)
            If tmpRS.EOF = False Then
                Dim I As Integer, j As Integer
                j = tmpRS.RecordCount
                If j > 5 Then j = 5
                For I = 0 To j - 1
                    WO(I) = tmpRS("Machine") + " " + tmpRS("Slot") + "-" + tmpRS("LR") '0046
                    Model(I) = tmpRS("model")
                    Work_Order(I) = tmpRS("Work_Order")  ''''1093
                    DIDType(I) = tmpRS("DIDType")
                    
                MachineCH(I) = tmpRS("MachineCH")         ''1044
                SideCH(I) = tmpRS("SideCH")         ''1044
                LRCH(I) = tmpRS("LRCH")         ''1044
                SlotCH(I) = tmpRS("Slot")         ''1044
                PN(I) = tmpRS("PN")

                    tmpRS.MoveNext
                Next I
            End If
        End If
        If OptZebra.Value = True Then
            isZebra = True
            
            ''''''updated by Jing   (0032)''''''
            If opOldLabel.Value = True Then
                LabelFile = Settings.AutoDispatchLabel
            Else
                LabelFile = Settings.AutoDispatchNewLabel
            End If
            strPrinterType = "Zebra"
        Else
            isZebra = False
            
            ''''''updated by Jing   (0032)''''''
            If opOldLabel.Value = True Then
                LabelFile = Settings.AutoDispatchSatoLabel
            Else
                LabelFile = Settings.AutoDispatchSatoNewLabel
            End If
            strPrinterType = "SATO"
        End If
        
        If Dir(LabelFile) = vbNullString Then
            ''''''Added by Jing 2008.01.10  (0019)''''''
            MsgBox ("Can not find label file !"), vbCritical
            PrintLabelCompPort = "PRN_FileNoExist"
            Exit Function
        End If
        
        MSComm.CommPort = TxtCompPort 'Settings.PRNa_Port
        MSComm.Settings = TxtComm 'Settings.PRNa_Settings
        MSComm.OutBufferCount = 0 '清空输出缓存
        If MSComm.PortOpen = False Then MSComm.PortOpen = True
        tmpPrintStr = ""
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
                tmpPrintStr = tmpPrintStr & Trim(hString)
          End Select
        Loop
        Close #hFile
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
         tmpPrintStr = Replace(tmpPrintStr, "<DATE>", strDay)
         tmpPrintStr = Replace(tmpPrintStr, "<QTY>", strQty)
         
         ''''''updated by Jing (0032)''''''
         If opNewLabel.Value = True Then
            tmpPrintStr = Replace(tmpPrintStr, "<BU>", PrintData.Line) '(0037)
'            tmpPrintStr = Replace(tmpPrintStr, "<BU>", Left(PrintData.machine, 1)) '(0058)
         Else
            tmpPrintStr = Replace(tmpPrintStr, "<LINE>", PrintData.Line)
'            tmpPrintStr = Replace(tmpPrintStr, "<LINE>", Left(PrintData.machine, 1)) '(0037) '(0058)
         End If
'         If PrintData.Line <> Left(PrintData.machine, 1) Then '(0058)
'            Conn.Execute ("Insert into QSMS_Error_log(Appname,SubFunction,SubID,Col1,Col2,Col3,DetailDesc,TransDateTime) values(" & _
'                         "'QSMS','PrintDID','Log','" & tmpDID & "','" & PrintData.Line & "','" & PrintData.machine & "'," & _
'                         "'Line and Machine did not match',dbo.formatdate(getdate(),'yyyymmddhhnnss'))")
'         End If
         tmpPrintStr = Replace(tmpPrintStr, "<SIDE>", PrintData.Side)
         tmpPrintStr = Replace(tmpPrintStr, "<MACHINE>", PrintData.Machine)
         tmpPrintStr = Replace(tmpPrintStr, "<DIDWOGROUP>", PrintData.DIDWOGROUP) '(0015)
         tmpPrintStr = Replace(tmpPrintStr, "<WOTYPE>", WOType)  ' (0039)
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
            
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order1>", Work_Order(0))       '''1093
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order2>", Work_Order(1))
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order3>", Work_Order(2))
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order4>", Work_Order(3))
            tmpPrintStr = Replace(tmpPrintStr, "<Work_Order5>", Work_Order(4))
            
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType1>", DIDType(0))
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType2>", DIDType(1))
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType3>", DIDType(2))
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType4>", DIDType(3))
            tmpPrintStr = Replace(tmpPrintStr, "<DIDType5>", DIDType(4))
            
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
        
        tmpPrintStr = Replace(tmpPrintStr, "<PN1>", PN(0))                    '1044
        tmpPrintStr = Replace(tmpPrintStr, "<PN2>", PN(1))
        tmpPrintStr = Replace(tmpPrintStr, "<PN3>", PN(2))
        tmpPrintStr = Replace(tmpPrintStr, "<PN4>", PN(3))
        tmpPrintStr = Replace(tmpPrintStr, "<PN5>", PN(4))
            
        End If
        Select Case Trim(tmpPrintStr)
           Case vbNullString
           Case Else
                 For M = 1 To Len(tmpPrintStr) Step 100
                     MSComm.Output = Mid(tmpPrintStr, M, 100)
                     'Debug.Print Mid(hString, m, 50)
                     DoEvents
                 Next M
        End Select
        MSComm.PortOpen = False
        
        '''(0038)'''
'        '**********************save the DID print log (0025)**********************
'        'add the log of DID machine data into qsms_error_log   (0030)
'        strSQL = "insert  into QSMS_Error_Log(AppName,SubFunction,SubID,DetailDesc,Col1,Col2,Col3,Col4,Col5,TransDateTime) values('QSMS_XL','PrintDID','Print DID Log','" & Trim(strDID) & "','" & Trim(strPrinterType) & "','" & g_userName & "','Com','" & CStr(Len(tmpPrintStr)) & "','" & Trim(PrintData.Machine) & "',dbo.FormatDate(getdate(),'yyyymmddhhnnss'))"
'        Conn.Execute strSQL
'        '**********************save the DID print log (0025)**********************
        Exit Function
errHandler:
        MsgBox Err.Description
        If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
        End If
End Function

Private Function PrintLabelPrintPort(strDID As String, strQty As String, Optional strAnotherQSMSIP As String) As String
Dim M As Integer
Dim tmpPrintStr As String
Dim hFile As Long
Dim hString As String
Dim tmpDID As String
Dim FileNum As Integer, lptPort As Integer
Dim strDay As String
Dim LabelFile, strLabelFileContent As String
Dim strPort As String
Dim strPrinterType As String
Dim tmpStr As String
Dim tmpRS As ADODB.Recordset
Dim strSQL As String
Dim rsTime As ADODB.Recordset

On Error GoTo errHandler
        
        '1112
        strSQL = "select getdate()"
        Set rsTime = Conn.Execute(strSQL)
        strDay = Format(rsTime(0), "YYMMDDHHNNSS") '1101
        strDID = UCase(Trim(CboDID))
        strQty = Replace(Replace(Replace(Replace(TxtQty, " ", ""), vbCr, ""), vbLf, ""), "pcs", "") ''1247
        
     LabelFile = GetDIDLabelFile(frmMaintainDIDAutoDispatch, IIf(opOldLabel.Value = True, "OLD", "NEW"))
'    If Dir(LabelFile) = vbNullString Then
'        ''''''Added by Jing 2008.01.10  (0019)''''''
'        MsgBox ("Can not find label file:" & LabelFile & "!"), vbCritical
'        PrintLabel = "PRN_FileNoExist"
'        Exit Function
'    End If
    
        ''''''added by Jing 2008.04.05  (0032)''''''
        If opNewLabel.Value = True Then
            Dim x As Integer
            For x = 0 To 4
                WO(x) = ""
                Model(x) = ""
                Work_Order(x) = ""
                DIDType(x) = ""
               
            Next x

            tmpStr = "Exec QSMS_GetDIDPrintInfo @DID='" & Trim(strDID) & "',@AnotherQSMSIP='" & Trim(strAnotherQSMSIP) & "',@PrinterType='" & Trim(PrinterType) & "',@PrintDpm='" & Trim(PrintDpm) & "'"        '''(0056)
            Set tmpRS = Conn.Execute(tmpStr)
            
            If tmpRS.EOF = False Then
                Dim I As Integer, j As Integer
                j = tmpRS.RecordCount
                If j > 5 Then j = 5
                For I = 0 To j - 1
                    WO(I) = tmpRS("Machine") + " " + tmpRS("Slot") + "-" + tmpRS("LR") '0046
                    Model(I) = tmpRS("model")
                    Work_Order(I) = tmpRS("Work_Order")  '''1093
                    DIDType(I) = tmpRS("DIDType")
                    
                    MachineCH(I) = tmpRS("MachineCH")         ''1044
                    SideCH(I) = tmpRS("SideCH")         ''1044
                    LRCH(I) = tmpRS("LRCH")         ''1044
                    SlotCH(I) = tmpRS("Slot")         ''1044
                    PN(I) = tmpRS("PN")
                    tmpRS.MoveNext
                Next I
            End If
        End If
        
        If OptZebra.Value = True Then
            isZebra = True
            
            ''''''updated by Jing   (0032)''''''
            If opOldLabel.Value = True Then
                LabelFile = Settings.AutoDispatchLabel
            Else
                LabelFile = Settings.AutoDispatchNewLabel
            End If
            
            strPrinterType = "Zebra"
        Else
            isZebra = False
            
            ''''''updated by Jing   (0032)''''''
            If opOldLabel.Value = True Then
                LabelFile = Settings.AutoDispatchSatoLabel
            Else
                LabelFile = Settings.AutoDispatchSatoNewLabel
            End If
            
            strPrinterType = "SATO"
        End If
        
        
'        strLabelFileContent = funGetTxtFileContent(LabelFile)
        If Dir(LabelFile) = vbNullString Then
            ''''''Added by Jing 2008.01.10  (0019)''''''
            MsgBox ("Can not find label file !"), vbCritical
            Exit Function
        End If
        lptPort = OpenOutputFile("LPT1")
        If lptPort = 0 Then
            MsgBox "Open print port LPT1 error!"
            Exit Function
        End If
        tmpPrintStr = ""
        FileNum = FreeFile()
        Open LabelFile For Input As #FileNum
        While Not EOF(FileNum)
           Line Input #FileNum, hString
                hString = Trim(hString)
                tmpPrintStr = tmpPrintStr & Trim(hString)
        Wend
        tmpDID = Trim(strDID)  '***************add by jeanson 20070814******
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
        
        ''''''updated by Jing (0032)''''''
        If opNewLabel.Value = True Then  '(0053)
            tmpPrintStr = Replace(tmpPrintStr, "<BU>", PrintData.Line)
'            tmpPrintStr = Replace(tmpPrintStr, "<BU>", Left(PrintData.machine, 1)) '(0058)
        Else
            tmpPrintStr = Replace(tmpPrintStr, "<LINE>", PrintData.Line)
'            tmpPrintStr = Replace(tmpPrintStr, "<LINE>", Left(PrintData.machine, 1)) '(0058)
        End If
        
'        If PrintData.Line <> Left(PrintData.machine, 1) Then '(0058)
'            Conn.Execute ("Insert into QSMS_Error_log(Appname,SubFunction,SubID,Col1,Col2,Col3,DetailDesc,TransDateTime) values(" & _
'                         "'QSMS','PrintDID','Log','" & tmpDID & "','" & PrintData.Line & "','" & PrintData.machine & "'," & _
'                         "'Line and Machine did not match',dbo.formatdate(getdate(),'yyyymmddhhnnss'))")
'        End If
        tmpPrintStr = Replace(tmpPrintStr, "<SIDE>", PrintData.Side)
        tmpPrintStr = Replace(tmpPrintStr, "<MACHINE>", PrintData.Machine)
        tmpPrintStr = Replace(tmpPrintStr, "<DIDWOGROUP>", PrintData.DIDWOGROUP) '(0015)
        tmpPrintStr = Replace(tmpPrintStr, "<WOTYPE>", WOType)  ' (0039)
        
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
        
        tmpPrintStr = Replace(tmpPrintStr, "<PN1>", PN(0))                    '1044
        tmpPrintStr = Replace(tmpPrintStr, "<PN2>", PN(1))
        tmpPrintStr = Replace(tmpPrintStr, "<PN3>", PN(2))
        tmpPrintStr = Replace(tmpPrintStr, "<PN4>", PN(3))
        tmpPrintStr = Replace(tmpPrintStr, "<PN5>", PN(4))
            
        End If
        
        For M = 1 To Len(tmpPrintStr) Step 50
            Print #lptPort, Mid(tmpPrintStr, M, 50)
            DoEvents
        Next M

'        Open strPort For Output As #FileNum
'        Print #FileNum, strLabelFileContent
        Close #FileNum
        Close #lptPort
        
        '''(0038)'''
'        '**********************save the DID print log (0025)**********************
'        strSQL = "insert  into QSMS_Error_Log(AppName,SubFunction,SubID,DetailDesc,Col1,Col2,Col3,Col4,TransDateTime) values('QSMS_XL','PrintDID','Print DID Log','" & Trim(strDID) & "','" & Trim(strPrinterType) & "','" & g_userName & "','Com','" & CStr(Len(tmpPrintStr)) & "',dbo.FormatDate(getdate(),'yyyymmddhhnnss'))"
'        Conn.Execute strSQL
'        '**********************save the DID print log (0025)**********************
        Exit Function
errHandler:
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


''''''''''''''''''''''''''''''Updated by Jing 2007.12.21    (0012)''''''''''''''''''''''''''
Private Sub optExtra_Click()
Dim RS As New ADODB.Recordset

On Error GoTo errHandler

    Call clearExtraData
    If optExtra = True Then
        strSQL = "Exec XL_GetAllWOInfoList 'Line','','','','','','" & Trim(CboCompPN.Text) & "',''"
        Set RS = Conn.Execute(strSQL)
        If RS.EOF = False Then
            While Not RS.EOF
                cmbLine.AddItem (Trim(RS!GroupValue))
                RS.MoveNext
            Wend
        End If
    End If
    Exit Sub
    
errHandler:
        MsgBox Err.Description
        
'Dim Rs As New ADODB.Recordset
'Call clearExtraData
'If optExtra = True Then
'    strsql = "Exec XL_GetAllWOInfoList 'WO'"
'    Set Rs = Conn.Execute(strsql)
'    If Rs.EOF = False Then
'        While Not Rs.EOF
'            cmbWO.AddItem (Trim(Rs!GroupValue))
'            Rs.MoveNext
'        Wend
'    End If
'End If
End Sub
Private Sub clearExtraData()
    cmbLine.Clear
    cmbWO.Clear
    cmbMachine.Clear
    cmbSlot.Clear
    cmbLR.Clear
    cmbSide.Clear
End Sub

Private Sub OptNormal_Click()
If OptNormal = True Then
    Call clearExtraData
End If
End Sub
'-------------0003----------------------------

''''''''''''''''updated by Jing 2007.12.21  (0012)''''''''''''''''''
Private Sub optSpecial_Click()
Dim RS As New ADODB.Recordset

On Error GoTo errHandler

    Call clearExtraData
    If optSpecial = True Then
        strSQL = "Exec XL_GetAllWOInfoList 'Line','','','','','','" & Trim(CboCompPN.Text) & "',''"
        Set RS = Conn.Execute(strSQL)
        If RS.EOF = False Then
            While Not RS.EOF
                cmbLine.AddItem (Trim(RS!GroupValue))
                RS.MoveNext
            Wend
        End If
    End If
    Exit Sub
    
errHandler:
        MsgBox Err.Description


'Dim Rs As New ADODB.Recordset
'
'On Error GoTo errHandler
'    Call clearExtraData
'    If optSpecial = True Then
'        strsql = "Exec XL_GetAllWOInfoList 'WO'"
'        Set Rs = Conn.Execute(strsql)
'        If Rs.EOF = False Then
'            While Not Rs.EOF
'                cmbWO.AddItem (Trim(Rs!GroupValue))
'                Rs.MoveNext
'            Wend
'        End If
'    End If
'    Exit Sub
'errHandler:
'        MsgBox Err.Description
End Sub

Private Sub optToWO_Click() '(0044)
Dim RS As New ADODB.Recordset

On Error GoTo errHandler

    Call clearExtraData
    If optToWO = True Then
        strSQL = "Exec XL_GetAllWOInfoList 'Line','','','','','','" & Trim(CboCompPN.Text) & "',''"
        Set RS = Conn.Execute(strSQL)
        If RS.EOF = False Then
            While Not RS.EOF
                cmbLine.AddItem (Trim(RS!GroupValue))
                RS.MoveNext
            Wend
        End If
    End If
    Exit Sub
    
errHandler:
        MsgBox Err.Description
End Sub

Private Sub txtExtraQty_Click()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtExtraQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And txtExtraQty.Text <> "" Then
'    cmdAdd.SetFocus    '''(0055)
    If IsInteger(TxtQty) = False Then
        MsgBox ("please check the qty of print, Qty>0!"), vbCritical
        TxtQty.SetFocus
        Exit Sub
    End If
    cmdDelete.Enabled = True
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdExit.Enabled = True
    cmdFind.Enabled = True
    CboCompPN.Enabled = True
    CboVendorCode.Enabled = True
    txtDateCode.Enabled = True
    txtLotCode.Enabled = True
    TxtQty.Enabled = True
    CommandType = 1
    cmdSave.SetFocus
End If
End Sub

Private Sub TxtGroupQty_Click()
    SendKeys "{Home}+{End}"
End Sub

Private Sub TxtGroupQty_KeyPress(KeyAscii As Integer)
Dim I As Integer

On Err GoTo errHandler:
If KeyAscii = 13 And TxtGroupQty <> "" Then

    ''''''added by Jing (0028)''''''
    If IsInteger(TxtGroupQty) = False Then '00050
        MsgBox ("Please input number !"), vbCritical
        Call Warning_Sound
        Call ClearData
        Exit Sub
    End If
    
    If chkQty = "" Then
        If CInt(Trim(TxtGroupQty)) > 10 Then
            MsgBox ("Max reel Qty must be less than 10 !"), vbCritical
            Call Warning_Sound
            Call ClearData
            Exit Sub
        End If
    Else ''(0061)
        If ChkOldDIDLabelQty <> "" And opOldLabel.Value = True Then
            If CInt(Trim(TxtGroupQty)) > CInt(ChkOldDIDLabelQty) Then
                MsgBox ("Max reel Qty must be less than " & ChkOldDIDLabelQty & " for old label !"), vbCritical
                Call Warning_Sound
                Call ClearData
                Exit Sub
            End If
        Else
            If CInt(Trim(TxtGroupQty)) > CInt(chkQty) Then
                MsgBox ("Max reel Qty must be less than " & chkQty & " !"), vbCritical
                Call Warning_Sound
                Call ClearData
                Exit Sub
            End If
        End If
    End If

'    cmdAdd.SetFocus        ''''(0055)
    If IsInteger(TxtQty) = False Then
        MsgBox ("please check the qty of print, Qty>0!"), vbCritical
        TxtQty.SetFocus
        Call Warning_Sound
        Call ClearData
        Exit Sub
    End If
    cmdDelete.Enabled = True
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdExit.Enabled = True
    cmdFind.Enabled = True
    CboCompPN.Enabled = True
    CboVendorCode.Enabled = True
    txtDateCode.Enabled = True
    txtLotCode.Enabled = True
    TxtQty.Enabled = True
    CommandType = 1
    NeedMSD = False
    
    ''''(0057)
    
    If ScanMSD = "Y" Then
        NeedMSD = IsNeedMSD(Trim(CboCompPN))
    End If
    
    If ScanMSD = "Y" And NeedMSD = True Then
        txtMSD.SetFocus
        txtMSD_Click
    ElseIf UCase(Trim(StrBU)) = "AS" Then
        txtInspection.SetFocus
        txtInspection_Click
    Else
        cmdSave.SetFocus
        Call cmdSave_Click
    End If
    
End If
Exit Sub
errHandler:
    MsgBox Err.Description
End Sub


Private Sub txtInspection_Click()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtInspection_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13 Or KeyAscii = 9) And UCase(Trim(StrBU)) = "AS" Then
        cmdSave.SetFocus
        Call cmdSave_Click
    End If
End Sub
Private Sub txtMSD_Click()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtMSD_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13 Or KeyAscii = 9) Then
        '''1022
        If txtInspection.Visible = True Then
            txtInspection.SetFocus
        ElseIf cmdSave.Enabled Then
            cmdSave.SetFocus
        End If
        
'        labInpectionNo.Visible = True
'        txtInspection.Visible = True
'        txtInspection.SetFocus
        'If cmdSave.Enabled Then
            'cmdSave.SetFocus
        'End If
    End If
End Sub

Private Sub TxtQty_Click()
    SendKeys "{HOME}+{END}"
End Sub

'****************************
'*****add Qty check by jeanson 2006/09/07
Private Sub TxtQty_KeyPress(KeyAscii As Integer)
Dim str As String
If (KeyAscii = 13 Or KeyAscii = 9) Then
    '' TxtQty = Trim(TxtQty)   'trim the DID total value's space  (0027)
    TxtQty = Replace(Replace(Replace(Replace(TxtQty, " ", ""), vbCr, ""), vbLf, ""), "pcs", "") ''1247
    If IsInteger(TxtQty) = False Then '00050
        MsgBox ("Check the qty of print please !"), vbCritical
        TxtQty.SetFocus
        Call TxtQty_Click
        Exit Sub
    End If
    If TxtQty >= 120000 Then '(0054)  (0078)
        MsgBox ("The Qty >120000,please check !"), vbCritical
        TxtQty.SetFocus
        Call TxtQty_Click
        Exit Sub
    End If
    'cmdAdd.SetFocus
    TxtGroupQty.SetFocus
    Call TxtGroupQty_Click
End If
End Sub
'Function GetDIDQty(Qty As String) As Boolean
'Dim i As Long
'GetDIDQty = True
'If Qty = "" Then
'    GetDIDQty = False
'    Exit Function
'End If
'If IsInteger(Qty) = False Then '0048
'    GetDIDQty = False
'    Exit Function
'End If
'
'If Qty <= 0 Then
'    GetDIDQty = False
'End If
'End Function
''******************************

Private Sub LockTheForm(lockCtl As Boolean)
' On Error Resume Next
 Dim ctl As Control
 
' For Each ctl In Me.Controls
'
'   Debug.Print ctl
'   If ctl <> False Or ctl <> True Then
'   ctl.Enabled = lockCtl
'   End If
'
' Next ctl
 Frame4.Enabled = lockCtl
 OptComp.Enabled = lockCtl
 OptPrint.Enabled = lockCtl
 CmdCommSave.Enabled = lockCtl
 TxtGroupQty.Enabled = lockCtl
 CboCompPN.Enabled = lockCtl
 CboVendorCode.Enabled = lockCtl
 txtDateCode.Enabled = lockCtl
 txtLotCode.Enabled = lockCtl
 TxtQty.Enabled = lockCtl
 CboDID.Enabled = lockCtl
 cmdFind.Enabled = lockCtl
' cmdAdd.Enabled = lockCtl
 cmdDelete.Enabled = lockCtl
 cmdSave.Enabled = lockCtl
 cmdCancel.Enabled = lockCtl
 CmdRefresh.Enabled = lockCtl
 cmdExit.Enabled = lockCtl
 CmdReprint.Enabled = lockCtl
 DG1.Enabled = lockCtl
 End Sub

Private Function ChkAVL(ByVal compPN As String, ByVal VendorCode As String) As Boolean
Dim strSQL As String
Dim RS As ADODB.Recordset
ChkAVL = True
If Check_AVL <> "Y" Then
    ChkAVL = True 'add by Giant  --20070618
Else
    strSQL = "Select TOP 1 * from QSMS_AVL where CompPN='" & Trim(CboCompPN) & "' and VendorCode='" & Trim(CboVendorCode) & "' "
    Set RS = Conn.Execute(strSQL)
    If RS.EOF Then
        ChkAVL = False
        MsgBox "CompPN and VendorCode not match!! please check "
    End If
End If

End Function

'''''''added by Jing 2008.03.05  (0028)''''''
'Private Function ChkNumber(tmpStr As String) As Boolean
'Dim i As Integer
'
'tmpStr = Trim(tmpStr)
'If tmpStr <> "" Then
''    For i = 1 To Len(tmpStr)
''        If IsNumeric(Mid(tmpStr, i, 1)) = False Then
''            ChkNumber = False
''            Exit Function
''        End If
''    Next i
'    If IsInteger(tmpStr) = False Then '0048
'        ChkNumber = False
'        Exit Function
'    End If
'Else
'    ChkNumber = False
'    Exit Function
'End If
'ChkNumber = True
'End Function
Private Function GetGroupDIDQty(compPN As String) As Integer  '(0042)
Dim I As Integer
On Error GoTo ErrHdl:
GetGroupDIDQty = 0
For I = 0 To UBound(arryGroupDIDQty, 2) - 1
    If arryGroupDIDQty(0, I) = compPN Then
        GetGroupDIDQty = arryGroupDIDQty(1, I)
    End If
Next I
Exit Function
ErrHdl:
    GetGroupDIDQty = 0
End Function
Private Sub RefreshGroupDIDQty()  '(0042)
Dim RS As New ADODB.Recordset
Dim I As Integer

''20080827 Denver Yang  如果下面语句不能正常执行，将发生循环调用，呈现死机状态。
''On Error Resume Next

strSQL = "select CompPN,Qty from XL_MaxDIDMaintainQty order by comppn"
Set RS = Conn.Execute(strSQL)
If RS.EOF = False Then
    TempArry = RS.GetRows
    ReDim arryGroupDIDQty(2, UBound(TempArry, 2)) As String
    For I = 0 To UBound(TempArry, 2)
        arryGroupDIDQty(0, I) = TempArry(0, I)
        arryGroupDIDQty(1, I) = TempArry(1, I)
    Next I
End If
End Sub

Private Function IC_CompNeedBurn(compPN As String) As Boolean
On Error GoTo errHandler:
    
    Dim RS As New ADODB.Recordset
    Dim NeedBurn As Boolean
    
    NeedBurn = False
    
    strSQL = "exec IC_CompNeedBurn " & sq(compPN)
    Set RS = Conn.Execute(strSQL)
    If RS.EOF = False Then
        If RS("Result") = 0 Then
            If MsgBox(RS("Description") & " DO you burn IC for it firstly!", vbYesNo) = vbYes Then
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
            
            
Private Sub InitFlex(Flex As MSFlexGrid)
    Dim intCol As Integer
    
    With Flex
        .Rows = 1
        .Rows = 20
        .Cols = 24
        
        ''Add GroupID, 以当前是否有需求来显示 PlanBalanceQty
        ''.FormatString = "|WO|WOqty|Line|Side|Machine|Slot|LR|CompPN|Item|BaseQty|NeedQty|DispatchQty|BalanceQty|PlanQty|PlanNeedQty|PlanBalanceQty|WorkDate|Shift|WOSeqID|SAPPercentage|Jobgroup|Jobpn"
        .FormatString = "|GroupID|WO|WOqty|Machine|Slot|LR|CompPN|Item|BaseQty|NeedQty|DispatchQty|BalanceQty|PlanQty|PlanNeedQty|PlanBalanceQty|WorkDate|Shift|WOSeqID|SAPPercentage|Jobgroup|Jobpn|Line|Side"
        
        .ColWidth(0) = 300
'        .ColWidth(1) = 1360
        .ColWidth(1) = 1800  '20110104 Maggie 增长显示GroupID的格子长度 ’20110105   Denver 使用户可调整
        .ColWidth(2) = 1000
        .ColWidth(3) = 600
        .ColWidth(4) = 800
        .ColWidth(5) = 600
        .ColWidth(6) = 420
        .ColWidth(7) = 1260  'CompPN
        .ColWidth(8) = 420
        .ColWidth(9) = 800
        .ColWidth(10) = 900
        .ColWidth(11) = 1000
        .ColWidth(12) = 1000   'BalanceQty
        .ColWidth(13) = 800
        .ColWidth(14) = 1100
        .ColWidth(15) = 1200   'PlanBalanceQty
        .ColWidth(16) = 900
        .ColWidth(17) = 600
        .ColWidth(18) = 1000  'WOSeqID
        .ColWidth(19) = 1260  'SAPPercentage
        .ColWidth(20) = 1000  'Jobgroup
        .ColWidth(21) = 1000  'JobPN
        .ColWidth(22) = 600  'Line
        .ColWidth(23) = 600   'Side
        
        .Col = 1
'        .TextStyle = flexTextRaised
        .CellAlignment = flexAlignCenterCenter '4
        .ColAlignment(1) = flexAlignCenterCenter '4
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
            
''RQ10071501          Denver       2010/07/23    For One By One Material,it need show Current/Next Shift Material Demand  (0076)
''20100726            Denver       2010/07/26 非OneByOne 材料也需要按此方式显示，并将 BalanceQty/PlanBalanceQty 突出显示
Private Sub GetOneByOneMaterialDemand(compPN As String, VendorCode As String, DateCode As String, LotCode As String)
    Dim RS As New ADODB.Recordset
    On Error GoTo Err_Handler
    
    lblStatus = ""
    strSQL = "exec [XL_Dispatch_MaterialPrompt]  @CompPN=" & sq(compPN) & ",@VendorCode=" & sq(VendorCode) & ",@DateCode=" & sq(DateCode) & ",@LotCode=" & sq(LotCode) & ",@Factory=" & sq(Factory)
    Set RS = Conn.Execute(strSQL)
    If RS.EOF = False Then
        If RS("Result") = 0 Then
            Set RS = RS.NextRecordset
'            Set DGOneByOne.DataSource = rs
'            ''调整列宽度,方便OP查看信息
'            DGOneByOne.Columns("WO").Width = 1000
'            DGOneByOne.Columns("WOQty").Width = 600
'            DGOneByOne.Columns("Line").Width = 420
'            DGOneByOne.Columns("Side").Width = 420
'            DGOneByOne.Columns("Machine").Width = 800
'            DGOneByOne.Columns("Slot").Width = 420
'            DGOneByOne.Columns("LR").Width = 420
'            DGOneByOne.Columns("CompPN").Width = 1200
'            DGOneByOne.Columns("Item").Width = 420
            
            Call FillFlexData(RS, flexGridDemandMaterial)
            
        Else
'            Set DGOneByOne.DataSource = Nothing
            
            Call InitFlex(flexGridDemandMaterial)
            '''''''''1265 Begin'''''''
            If StrBU = "NB5" Then
               If RS("Result") = 2 Then
                  MsgBox Trim(RS.Fields("Description")), vbOKOnly Or vbInformation, "系统提示"
               End If
            End If
            '''''''''1265 End'''''''''
            lblStatus = RS("Description")
        End If
'        DGOneByOne.Refresh
    End If
    
    Exit Sub
Err_Handler:
    
    MsgBox Err.Number & "," & Err.Description
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
Private Sub Assigned_Sound()
    Call EffectSound("Assigned.wav")
End Sub
Private Sub ClearData()
    CboCompPN.Text = ""
    CboVendorCode.Text = ""
    txtDateCode.Text = ""
    txtLotCode.Text = ""
    CboDID.Text = ""
    TxtQty.Text = ""
    txtMSD.Text = ""
End Sub
