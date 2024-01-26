VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmInSpection 
   BackColor       =   &H80000013&
   Caption         =   "InSpection[20150603]"
   ClientHeight    =   10245
   ClientLeft      =   675
   ClientTop       =   255
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   10245
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H80000013&
      Height          =   855
      Left            =   6210
      TabIndex        =   35
      Top             =   -30
      Width           =   10575
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   7200
         Top             =   0
      End
      Begin VB.Label UserID 
         BackColor       =   &H00FFC0C0&
         Caption         =   "UserID"
         DataMember      =   "&H00FFC0C0&"
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
         TabIndex        =   41
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbluid 
         Caption         =   "Label9"
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
         Left            =   1080
         TabIndex        =   40
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblDatetime 
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
         Left            =   6120
         TabIndex        =   39
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblVersion 
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
         Left            =   3600
         TabIndex        =   38
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Version"
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
         Left            =   2640
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Time"
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
         Left            =   5280
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
   End
   Begin MCI.MMControl wave_control 
      Height          =   495
      Left            =   14880
      TabIndex        =   32
      Top             =   3960
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   11550
      TabIndex        =   25
      Top             =   9240
      Width           =   3345
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11550
      TabIndex        =   24
      Top             =   8760
      Width           =   3345
   End
   Begin VB.TextBox txtimagePath 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7950
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   7620
      Width           =   8535
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   675
      Left            =   14880
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      WhatsThisHelpID =   20
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   1191
      _Version        =   327682
      BorderStyle     =   1
      Max             =   100
      SelStart        =   50
      TickStyle       =   2
      Value           =   50
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8925
      Left            =   150
      TabIndex        =   0
      Top             =   -210
      Width           =   6015
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         Caption         =   "Port"
         Height          =   1185
         Left            =   90
         TabIndex        =   29
         Top             =   7080
         Width           =   5775
         Begin VB.ComboBox cboEquipType 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            ItemData        =   "FrmInSpection.frx":0000
            Left            =   3120
            List            =   "FrmInSpection.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   660
            Width           =   2535
         End
         Begin VB.OptionButton OptRS232 
            BackColor       =   &H000080FF&
            Caption         =   "RS232( COM6)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   31
            Top             =   240
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton OptGPIB 
            BackColor       =   &H000080FF&
            Caption         =   "GPIB(22)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label12 
            BackColor       =   &H000080FF&
            Caption         =   "Equipment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   660
            Width           =   2775
         End
      End
      Begin VB.OptionButton optIC 
         BackColor       =   &H0000FF00&
         Caption         =   "IC&&connecter(芯片,接口)"
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
         Left            =   3000
         TabIndex        =   28
         Top             =   7560
         Width           =   2895
      End
      Begin VB.OptionButton OptInduc 
         BackColor       =   &H0000FF00&
         Caption         =   "&Inductance(电感)"
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
         TabIndex        =   27
         Top             =   7560
         Width           =   2775
      End
      Begin VB.TextBox txtDelaytime 
         Height          =   495
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   8310
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H000000FF&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   90
         TabIndex        =   8
         Top             =   8310
         Width           =   1575
      End
      Begin VB.TextBox txtDID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label LotNo 
         BackColor       =   &H00FFC0C0&
         Caption         =   "LotNo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   57
         Top             =   6480
         Width           =   2055
      End
      Begin VB.Label lblLotNo 
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
         Left            =   2280
         TabIndex        =   56
         Top             =   6480
         Width           =   3615
      End
      Begin VB.Label lblcurrent 
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
         Left            =   2280
         TabIndex        =   53
         Top             =   6060
         Width           =   3615
      End
      Begin VB.Label lblVoltage 
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
         Left            =   2280
         TabIndex        =   52
         Top             =   5590
         Width           =   3615
      End
      Begin VB.Label lblFrequency 
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
         Left            =   2280
         TabIndex        =   51
         Top             =   5121
         Width           =   3615
      End
      Begin VB.Label Current 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Current(A)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   6060
         Width           =   2055
      End
      Begin VB.Label Voltage 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Voltage(V)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   5590
         Width           =   2055
      End
      Begin VB.Label Frequency 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Frequency(Hz)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   5121
         Width           =   2055
      End
      Begin VB.Label lblSpec 
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
         Left            =   2280
         TabIndex        =   47
         Top             =   3245
         Width           =   3615
      End
      Begin VB.Label Spec 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Spec.(+/- % ) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   3245
         Width           =   2055
      End
      Begin VB.Label lblunit 
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
         Left            =   2280
         TabIndex        =   45
         Top             =   4652
         Width           =   3615
      End
      Begin VB.Label Unit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   4652
         Width           =   2055
      End
      Begin VB.Label lblDIDUID 
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
         Left            =   2280
         TabIndex        =   43
         Top             =   4183
         Width           =   3615
      End
      Begin VB.Label DIDUID 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DIDUID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   4183
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DelaytTime"
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
         Left            =   1740
         TabIndex        =   26
         Top             =   8310
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblVendorPN 
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
         Left            =   6120
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "VendorPN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   20
         Top             =   1680
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblDIDQty 
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
         Left            =   2280
         TabIndex        =   17
         Top             =   2776
         Width           =   3615
      End
      Begin VB.Label sds 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DIDQty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2776
         Width           =   2055
      End
      Begin VB.Label lblChkNum 
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
         Left            =   2280
         TabIndex        =   14
         Top             =   3714
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "ChkNum:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   3714
         Width           =   2055
      End
      Begin VB.Label lblLotCode 
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
         Left            =   2280
         TabIndex        =   11
         Top             =   2307
         Width           =   3615
      End
      Begin VB.Label lblDateCode 
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
         Left            =   2280
         TabIndex        =   10
         Top             =   1838
         Width           =   3615
      End
      Begin VB.Label lblVendor 
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
         Left            =   2280
         TabIndex        =   9
         Top             =   1369
         Width           =   3615
      End
      Begin VB.Label lblcomppn 
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
         Left            =   2280
         TabIndex        =   7
         Top             =   900
         Width           =   3615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "LotCode:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2307
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DateCode:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1838
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Vendor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1369
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "QuantaPN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   900
         Width           =   2055
      End
      Begin VB.Label DID 
         BackColor       =   &H0000FF00&
         Caption         =   "DID\CompPN:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1485
      Left            =   120
      TabIndex        =   18
      Top             =   8760
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   2619
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   23
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
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
   Begin MSCommLib.MSComm MSComm1 
      Left            =   -150
      Top             =   -480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox txtSleepTime 
      Height          =   345
      Left            =   60
      TabIndex        =   54
      Text            =   "200"
      Top             =   7680
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblMsg 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6270
      TabIndex        =   55
      Top             =   7020
      Width           =   8595
   End
   Begin VB.Image Image1 
      Height          =   4800
      Left            =   6960
      Top             =   1380
      Width           =   6000
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ImagePath"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6270
      TabIndex        =   22
      Top             =   7620
      Width           =   1575
   End
   Begin VB.Label lblMsg1 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   19
      Top             =   11880
      Width           =   16455
   End
   Begin VB.Menu munIPQCRetest 
      Caption         =   "IPQC_ReTest"
   End
End
Attribute VB_Name = "FrmInSpection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''***********************************************************************************************
'''**EQMS           修 改 人     修改日期        描    述
'''**RQ09102710     Denver      2009.10.27      Add 测试LCR 型号为4300 的仪器  （0063）
'''**QMS            Austin      2010.04.19      Add 测试LCR 型号为6420 的仪器  （0064）
'''**QMS            Austin      2010.05.11      Change Flag:ScanDID=>Flag:ScanCompPN
'''**QMS            Austin      2010.05.24      生成LotNO=Site+yyyyMMddHHmmss (0065)
'''**QMS            Van         2013.05.16      DID数量过大，导致超出了函数CInt范围,改为CLng（0066）
'''**QMS            Feix        2017.06.14      判定测试次数是否超过四次（1258）
'''***********************************************************************************************
Private PreCompPN As String
Private DIDQty As String
Private strDID As String
Private strLotNo As String
Const EMPTYSTR = ""


Public Sub Delay_Time(MilliSecond As Long)
  Dim i As Long
  i = GetTickCount()
  Do While Abs(GetTickCount() - i) < MilliSecond
     DoEvents
  Loop

End Sub
'''0065
Private Function GetLotNo() As String
On Error Resume Next

    Dim strSQL As String
    Dim Rs As New ADODB.Recordset
    Dim strSite As String
    Dim transdatetime As String
    
    strSQL = "exec GetLotNo"
    Set Rs = Conn.Execute(strSQL)
    
    GetLotNo = Rs!LotNo
    
End Function

Private Function RefreshDataGrid()
DataGrid1.Columns(1).Width = 2600
DataGrid1.Columns(1).Width = 2000
DataGrid1.Columns(2).Width = 1000
DataGrid1.Columns(3).Width = 1000
DataGrid1.Columns(4).Width = 800
DataGrid1.Columns(5).Width = 2000

DataGrid1.Refresh


End Function


Private Sub cmdStart_Click()
  Dim strSQL As String, TestType  As String
  Dim Rs As New ADODB.Recordset
  Dim sGetStr As Long
  Dim TestStr As String
  Dim TestValue As Double
  Dim Errcode As String, ErrDesc As String
  Dim i As Integer
  Dim Unit As String
  Dim LCKNum As Integer '00001
'  Dim TestType As srting
  'If txtDID = "" Then Exit Sub
 On Error GoTo errHandler
  If ScanCompPN = "Y" Then
    TxtDID.Enabled = False
  End If
  cmdStart.Enabled = False
  
  If cboEquipType = "" Then
     lblmsg = "Please select EquipType!"
     lblmsg.BackColor = &H80FFFF
     Call Warning_Sound
     cmdStart.Enabled = True
     Exit Sub
  End If
  
 '''==============重测并清除以前测试数据================
    If ScanCompPN = "Y" Then
        lblmsg = TxtDID & " Test allowed"
    Else
'        If Factory = "T2" Then
            strSQL = "select IPQCFlag from QSMS_DID_InSpect where DID='" & Trim(TxtDID) & "'"  ''''判断是否以测试过
            If Rs.State Then Rs.Close
            Rs.Open strSQL, Conn
            If Rs.EOF = False Then
                If Rs("IPQCFlag") = "Y" Then
                     If chkIPQCRetest(lbluid, TxtDID) = False Then
                        lblmsg = TxtDID & " Pass has been tested, no permission to retest!"
                        lblmsg.BackColor = &HFF&
                        TxtDID = ""
                        TxtDID.Enabled = True
                        Exit Sub
                     End If
                ElseIf Rs("IPQCFlag") = "N" Then
                     If chkIPQCRetest(lbluid, TxtDID) = False Then
                        lblmsg = TxtDID & " Fail has been tested, no permission to retest！"
                        lblmsg.BackColor = &HFF&
                        TxtDID = ""
                        TxtDID.Enabled = True
                        Exit Sub
                     End If
                Else
                    lblmsg = TxtDID & " Test allowed"
                End If
            Else
               lblmsg = TxtDID & " Test allowed"
            End If
'        Else
'            strSQL = "selecT IPQCFlag from qsms_DID where DID='" & Trim(txtDID) & "'"  ''''判断是否以测试过
'            If RS.State Then RS.Close
'            RS.Open strSQL, Conn
'            If RS("IPQCFlag") = "Y" Then
'                 If chkIPQCRetest(lbluid, txtDID) = False Then
'                    lblMsg = txtDID & "已测试PASS，没有权限重测！"
'                    lblMsg.BackColor = &HFF&
'                    txtDID = ""
'                    txtDID.Enabled = True
'                    Exit Sub
'                 End If
'            ElseIf RS("IPQCFlag") = "N" Then
'                 If chkIPQCRetest(lbluid, txtDID) = False Then
'                    lblMsg = txtDID & "已测试FAIL，没有权限重测！"
'                    lblMsg.BackColor = &HFF&
'                    txtDID = ""
'                    txtDID.Enabled = True
'                    Exit Sub
'                 End If
'            Else
'               lblMsg = txtDID & "允许测试"
'            End If
'        End If
    End If
    ''1258
    If StrBU = "ESBU" Then
        strSQL = "select top 1 TestNum  from qsms_DID_inspect where DID='" & Trim(TxtDID) & "'and TestResult='FAIL'  order by transdatetime desc"
        If Rs.State Then Rs.Close
        Rs.Open strSQL, Conn
        If Rs.EOF = False Then
           If Rs("TestNum") >= 4 Then
           MsgBox ("It has exceeded the limit of LCK test times, please contact IPQC for confirmation.")
           Exit Sub
           End If
        End If
    End If
    ''1258
  '''========================================================

  If UCase(Trim(lblunit)) = "IC" Then
        If MsgBox("Is the IC information correct?", vbYesNo, "IC Test") = vbYes Then
            TestStr = 0   ''''0--> pass  ,1---> fail
            Errcode = "PASS"
        Else
            TestStr = 1
            Errcode = InputBox("Please enter the Error Code：", "IC Test", "")
            strSQL = "selecT * from IPQC_ErrCode where ErrCode='" & Errcode & "'   "
            If Rs.State Then Rs.Close
            Rs.Open strSQL, Conn
            If Rs.EOF Then
               lblmsg = "Please enter the correct Error Code！"
               lblmsg.BackColor = &H80FFFF
               Call Warning_Sound
               cmdStart.Enabled = True
               Exit Sub
            End If
        End If
  ElseIf UCase(Trim(lblunit)) = "CON" Then
     If MsgBox("Connecter PIN:" & lblSpec & " Is it consistent with the real thing?", vbYesNo, "CON Test") = vbYes Then
            TestStr = 0   ''''0--> pass  ,1---> fail
            Errcode = "PASS"
     Else
            TestStr = 1
            Errcode = InputBox("Please enter the Error Code：", "CON Test", "")
            strSQL = "selecT * from IPQC_ErrCode where ErrCode='" & Errcode & "'   "
            If Rs.State Then Rs.Close
            Rs.Open strSQL, Conn
            If Rs.EOF Then
               lblmsg = "Please enter the correct Error Code！"
               lblmsg.BackColor = &H80FFFF
               Call Warning_Sound
               cmdStart.Enabled = True
               Exit Sub
            Else
               ErrDesc = Rs("CHErrDesc")
            End If
      End If

  Else
        Delay_Time (txtDelaytime)
        Unit = lblunit
        If Unit = "TRIODE" Then Unit = "DIODE"
        If OptGPIB.Value = True Then
           sGetStr = MeasureE_GPIB(cboEquipType, 22, UCase(Trim(Unit)), UCase(Trim(lblFrequency)), UCase(Trim(lblVoltage)), UCase(Trim(lblcurrent))) ''' 仪器端口22, 'D'测试用参数 ,'R'电阻
        ElseIf OptRS232.Value = True Then
            '''**RQ09102710  Denver      2009.10.27    Add 测试LCR 型号为4300 的仪器  （0063）
            If Trim(cboEquipType) = "4300" Then
                TestStr = MeasureE4300_RS232(cboEquipType, UCase(Trim(Unit)), UCase(Trim(lblFrequency)), UCase(Trim(lblVoltage)))
                TestStr = Val(TestStr)
            ElseIf Trim(cboEquipType) = "6420" Then
                ''0064
                TestStr = MeasureE6420_RS232(cboEquipType, UCase(Trim(Unit)), UCase(Trim(lblFrequency)), UCase(Trim(lblVoltage)))
                If UCase(Unit) = "C" Then
                    TestStr = Val(TestStr) * 1000000000000#
                Else
                    TestStr = Val(TestStr)
                End If
            ElseIf Trim(cboEquipType) = "3523" Then
                TestStr = Measure3523_RS232(cboEquipType, UCase(Trim(Unit)), UCase(Trim(lblFrequency)), UCase(Trim(lblVoltage)))
                TestStr = Val(TestStr)
            ElseIf Trim(cboEquipType) = "8110G" Then
                ''(1024)
                TestStr = MeasureE8110G_RS232(cboEquipType, UCase(Trim(Unit)), UCase(Trim(lblFrequency)), UCase(Trim(lblVoltage)))
                TestStr = Val(TestStr)
            Else
                sGetStr = MeasureE_RS232(cboEquipType, 6, UCase(Trim(Unit)), UCase(Trim(lblFrequency)), UCase(Trim(lblVoltage)), UCase(Trim(lblcurrent)))
            End If
        Else
           lblmsg = "Please select the device port！"
           lblmsg.BackColor = &H80FFFF
           Call Warning_Sound
           cmdStart.Enabled = True
           Exit Sub
        End If
    
        '''**RQ09102710  Denver      2009.10.27    Add 测试LCR 型号为4300 的仪器  （0063）
        If Trim(cboEquipType) <> "4300" And Trim(cboEquipType) <> "6420" And Trim(cboEquipType) <> "3523" And Trim(cboEquipType) <> "8110G" Then
            TestStr = GetStringFromPointer(sGetStr)
            ''MsgBox TestStr
            If Asc(Left(TestStr, 1)) = Asc("E") Then    '''the first character of error code is "E"
               lblmsg = "Device abnormality：" & TestStr
               lblmsg.BackColor = &H80FFFF
               Call Warning_Sound
               cmdStart.Enabled = True
               TxtDID.Enabled = True
               Call txtDID_Click
               Exit Sub
            End If
        End If
        
    End If
   
  TestValue = CDbl(TestStr)
  strSQL = "exec QSMSDIDInSpect '" & Trim(strDID) & "','" & Trim(lblCompPN) & "','" & Trim(lblVendor) & "','" & Trim(TestValue) & "' ,'" & Trim(Errcode) & "','" & g_userName & "','" & ScanCompPN & "'"
  If Rs.State Then Rs.Close
  Rs.Open strSQL, Conn
            Dim Rsdata As New ADODB.Recordset
            
            IPQCFlag = UCase(Rs("result"))
            Set Rsdata = Rs.NextRecordset
            DataGrid1.Refresh
            Set DataGrid1.DataSource = Rsdata
            
            Call RefreshDataGrid
            Rsdata.MoveLast
            
            If Rsdata("testorder") <> lblChkNum Then    ''''当完成测试时退出处理
                ''lblMsg = "该DID第" & Rsdata("testorder") & " 颗测试结果为:" & Rsdata("testresult") & " " & ErrDesc
                lblmsg = "The " & Rsdata("testorder") & " test result of this DID is" & Rsdata("testresult") & " " & ErrDesc
                If Rsdata("testresult") = "PASS" Then
                   lblmsg.BackColor = &HFF00&
                Else
                   lblmsg.BackColor = &HFF&
                End If
                
                If ScanCompPN = "Y" Then
                    TxtDID.Enabled = True
                End If
                cmdStart.Enabled = True
                cmdStart.SetFocus
                Exit Sub
            Else
                If IPQCFlag = "PASS" Then
                    ''''update qsms_did
                     Conn.Execute ("update QSMS_DID set IPQCFlag='Y' where DID='" & Trim(TxtDID) & "';update qsms_DID_inspect set IPQCFlag='Y' where DID='" & Trim(TxtDID) & "'  ")
                     lblmsg.BackColor = &HFF00&
                      Call OK_Sound
                Else
                      ''''update qsms_did
                     Conn.Execute ("update QSMS_DID set IPQCFlag='N' where DID='" & Trim(TxtDID) & "';update qsms_DID_inspect set IPQCFlag='N' where DID='" & Trim(TxtDID) & "'  ")
                     lblmsg.BackColor = &HFF&
                     Call Warning_Sound
                End If
                  If DataGrid1.Columns(4) = "CON" Or DataGrid1.Columns(4) = "IC" Then
                      lblmsg = TxtDID & "The result of the visual inspection is " & IPQCFlag & ", visual inspection completed. "
                  Else
                      lblmsg = TxtDID & "The result of the inspection is " & IPQCFlag & ", inspection completed."
                  End If
                  TxtDID.Enabled = True
                  TxtDID = ""
                  cmdStart.Enabled = False
                  TxtDID.SetFocus
                  Exit Sub
            End If
     Exit Sub
errHandler:
    lblmsg = Err.Description
    TxtDID.Enabled = True
    TxtDID.SetFocus
    Call txtDID_Click
End Sub

 

Private Sub Dir1_Change()
imagePath = Dir1.path
txtimagePath = imagePath
 SaveSetting "QSMS_InSpection", "imagepath", "imagepath", imagePath
End Sub

Private Sub Drive1_Change()
Dir1 = Drive1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 
ValKeyCode = 32 '''space
ValKeyCode = 32 '''space
If KeyCode = ValKeyCode Then Call cmdStart_Click
End Sub

Private Sub Form_Load()
cmdStart.Enabled = False

''Image1.Picture = LoadPicture(App.Path & "\apple.jpg", vbLPLarge, vbLPColor)
Image1.Stretch = True
Image1.Height = 4800
Image1.Width = 6000

txtDelaytime = "3000"

lblVersion = ProgramDescription
lbluid = g_userName


imagePath = GetSetting("QSMS_InSpection", "imagepath", "imagepath")
txtimagePath = imagePath

lblVersion = App.EXEName


cboEquipType.AddItem "34401", 0
cboEquipType.AddItem "3302", 1
cboEquipType.AddItem "4235", 2
cboEquipType.AddItem "4300", 3
cboEquipType.AddItem "6420", 4
cboEquipType.AddItem "8110G", 5
cboEquipType.AddItem "3523", 6

If ScanCompPN = "Y" Then
    lblLotNo.Visible = True
    LotNo.Visible = True
End If


End Sub



Private Sub Slider1_Change()
Image1.Width = 3000 * (1 + Slider1.Value / 50)
Image1.Height = 3000 * (1 + Slider1.Value / 50)
'Call ShowPicture(Trim(lblcomppn), Trim(lblVendor))
End Sub

Private Sub Timer1_Timer()
 cmdStart.Enabled = False
End Sub
Private Sub ProcessTestfile()
Dim strSQL As String
Dim Rs As New ADODB.Recordset
Dim Rsdata As New ADODB.Recordset
Dim strFile As String

strFile = Dir(TestFilepath & "\", vbNormal)
 Do While strFile <> ""
      lblmsg.Caption = "Processing Request " & strFile
      ReadTestfile TestFilepath & "\" & strFile
      If strFileContents <> "" Then
         strSQL = "exec QSMSDIDInSpect '" & Trim(TxtDID) & "','" & Trim(lblCompPN) & "','" & Trim(lblVendor) & "','" & Trim(strFileContents) & "'  "
         If Rs.State Then Rs.Close
         Rs.Open strSQL, Conn
         If UCase(Left(Rs("result"), 4)) = "PASS" Then
            lblmsg = Mid(Rs("result"), 6, 50)

            Set Rsdata = Rs.NextRecordset
            DataGrid1.Refresh
            Set DataGrid1.DataSource = Rsdata
            Call RefreshDataGrid
            Rsdata.MoveLast
           '' TestDIDFlag = UCase(Trim(Rsdata("Testresult")))
            If UCase(Trim(Rsdata("Testresult"))) = "FAIL" Then   ''''当测试FAIL时退出处理
               lblmsg = TxtDID & " Test is bad."
               TxtDID.Enabled = True
               cmdStart.Enabled = True
               Timer1.Enabled = False
               DelFile TestFilepath & "\" & strFile
               Exit Sub
            End If
            If Rsdata("testorder") = lblChkNum Then    ''''当完成测试时退出处理
               lblmsg = TxtDID & " Test completed."
                  ''''update qsms_did
               TxtDID.Enabled = True
               cmdStart.Enabled = True
               Timer1.Enabled = False
               DelFile TestFilepath & "\" & strFile
            End If
            DelFile TestFilepath & "\" & strFile
            strFile = Dir(TestFilepath & "\", vbNormal)
         End If
      Else
        MsgBox "The test result file is empty."
      End If
 Loop

End Sub

Private Sub Timer2_Timer()
lblDatetime = Now
End Sub


Private Sub txtDID_Click()
    Sendkeys "{HOME}+{END}"
End Sub
Private Sub GetBaseInfoByDID(strDID As String)
    Dim strSQL As String
    Dim Rs As New ADODB.Recordset
    
'    If Factory = "T2" Then
        strSQL = "select * from QSMS_DID where DID='" & Trim(strDID) & "'"
        If Rs.State Then Rs.Close
        Rs.Open strSQL, Conn
        
        If Rs.EOF = False Then
            lblCompPN = Rs("comppn")
            lblVendor = Rs("VendorCode")
            lblLotCode = Rs("LotCode")
            lblDateCode = Rs("DateCode")
            lblDIDQty = Rs("Qty")
            DIDQty = Rs("Qty")
            lblDIDUID = Rs("UID")
        Else
            strSQL = "select * from QSMS_DID_ToWH where DID='" & Trim(strDID) & "'"
            If Rs.State Then Rs.Close
            Rs.Open strSQL, Conn
            If Rs.EOF = False Then
                lblCompPN = Rs("comppn")
                lblVendor = Rs("VendorCode")
                lblLotCode = Rs("LotCode")
                lblDateCode = Rs("DateCode")
                lblDIDQty = Rs("Qty")
                DIDQty = Rs("Qty")
                lblDIDUID = Rs("UID")
            Else
                lblLotNo = GetLotNo
                Exit Sub
            End If
        End If
'    Else
'        strSQL = "select * from QSMS_DID where DID='" & Trim(strDID) & "'"
'        If RS.State Then RS.Close
'        RS.Open strSQL, Conn
'
'        If RS.EOF = False Then
'            lblcomppn = RS("comppn")
'            lblVendor = RS("VendorCode")
'            lblLotCode = RS("LotCode")
'            lblDateCode = RS("DateCode")
'            lblDIDQty = RS("Qty")
'            DIDQty = RS("Qty")
'            lblDIDUID = RS("UID")
'        Else
'            lblMsg = txtDID & "该DID不存在！"
'            lblMsg.BackColor = &HFF&
'            txtDID.Enabled = True
'            cmdStart.Enabled = False
'            Exit Sub
'        End If
'    End If
    
    If IPQC_ChkVendorPN = "Y" Then
        strSQL = "selecT * from Vendor_PN where comppn='" & Trim(lblCompPN) & "'  and vendor='" & Trim(lblVendor) & "'    "
        If Rs.State Then Rs.Close
        Rs.Open strSQL, Conn
        If Rs.EOF = False Then
           lblVendorPN = Trim(Rs("VendorPN"))
        Else
           lblmsg = "Vendor PN does not exist."
           lblmsg.BackColor = &HFF&
        End If
    End If
    
End Sub
Private Sub GetRuleByCompPN(ByVal StrPN As String)
    Dim strSQL As String
    Dim Rs As New ADODB.Recordset
    '' 20100430  Denver  重复定义，导致DIDQty=0
    Dim Upper As Double, Lower As Double, Unit As String   '', DIDQty As Long
    
    strSQL = "selecT * from QSMS_InSpect_Rule  where comppn='" & Trim(StrPN) & "' "
    If Rs.State Then Rs.Close
    Rs.Open strSQL, Conn
    If Rs.EOF = False Then
       Upper = Rs("Upper")
       Lower = Rs("Lower")
       lblFrequency = Rs("Hz")
       lblVoltage = Rs("Volt")
       lblcurrent = Rs("Ampere")
       lblunit = Rs("Unit")
       lblChkNum = Rs("CHKNum")
       
       If UCase(Trim(lblunit)) = "IC" Or UCase(Trim(lblunit)) = "CON" Then
          lblChkNum = 1
       End If
       
       If UCase(Trim(lblunit)) = "CON" Then
          lblSpec = Lower
       Else
          lblSpec = Lower & "-" & Upper
       End If
       
       'If ScanCompPN <> "Y" And Len(lblcomppn) > 12 Then       '（1085）
       If ScanCompPN <> "Y" And DIDQty <> "" Then         '（1132）
'            If CInt(DIDQty) < CInt(RS("BaseQty")) Then         '(0066)
        If CLng(DIDQty) < CLng(Rs("BaseQty")) Then
                lblChkNum = Rs("CHKNum") - 1
            End If
       Else
            lblChkNum = Rs("CHKNum")
       End If
    Else
       lblmsg = StrPN & "This CompPN test standard does not exist."
       lblmsg.BackColor = &HFF&
       TxtDID.Enabled = True
       cmdStart.Enabled = False
       Exit Sub
    End If
End Sub
Private Sub txtDID_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    If Len(Trim(TxtDID)) = 0 Then
        TxtDID.SetFocus
        Exit Sub
    End If
    
    TxtDID = Replace(TxtDID, Chr(10), "")
    TxtDID = Replace(TxtDID, Chr(13), "")
    
    TxtDID = Trim(TxtDID)
    
    Dim Rs As New ADODB.Recordset
    Dim strSQL As String
    Dim Upper As Double, Lower As Double, Unit As String, DIDQty As Long
       
    Call C_Interface
    
 ''''''clear  DataGrid
    Set DataGrid1.DataSource = Nothing

   ''''get DID basc info
   
    If ScanCompPN <> "Y" Then
        strDID = Trim(TxtDID)
        GetBaseInfoByDID (strDID)
    Else
        lblCompPN = TxtDID
        strLotNo = GetLotNo
        strDID = strLotNo
        lblLotNo = strLotNo
    End If
    
    If Len(Trim(TxtDID)) < 12 Then        '（1085）
        lblCompPN = TxtDID
    ElseIf Len(Trim(TxtDID)) < 21 Then
        lblCompPN = Left(TxtDID, 11)        '(1216)
    Else
        lblCompPN = Left(TxtDID, 15)
    End If
        
    
 ''''get rule of comppn
    GetRuleByCompPN (lblCompPN)
    
    Spec = ChangSpec(lblSpec, lblunit)
    Call ShowPicture(Trim(lblCompPN), Trim(lblVendor))
    If UCase(Trim(lblunit)) = "R" Or UCase(Trim(lblunit)) = "C" Or UCase(Trim(lblunit)) = "L" Or UCase(Trim(lblunit)) = "Z" Then
        If MsgBox("Is the information of DID or CompPN correct?", vbYesNo, "System Prompt") <> vbYes Then
           Call txtDID_Click
           Exit Sub
        End If
    End If
    
    cmdStart.Enabled = True
    cmdStart.SetFocus
    lblmsg = "Ready to test."
End If
End Sub
Private Function ShowPicture(COMPPN As String, Vendor As String)
Dim imagepatch As String
Dim PicFilename As String

'imagePath = "c:\pic\"
'imagePath = "\\172.17.0.12\SMT_Load\"
If imagePath = "" Then
    imagePath = App.path
End If

If COMPPN <> "" And Vendor <> "" Then
    PicFilename = COMPPN & "-" & Vendor & ".jpg"

    If Dir(imagePath & "\" & PicFilename) = "" Then
        If IPQC_ChkVendorPN = "Y" Then
            MsgBox "Path or file not found！"
            Exit Function
        Else
             Exit Function
           ' PicFilename = "IPQC_NoPicture.jpg"
        End If
    End If
    
    Image1.Picture = LoadPicture(imagePath & "\" & PicFilename, vbLPLarge, vbLPColor)
    Image1.Stretch = True
  
   
 ' Image1.Height = 6000
 ' Image1.Width = 8000
End If


End Function

Private Function chkIPQCRetest(UID As String, DID As String) As Boolean
   Dim Rs As New ADODB.Recordset
   Dim strSQL As String
     
     strSQL = "selecT * from userright where username='" & lbluid & "' and Userright='IPQCRetest' "
     If Rs.State Then Rs.Close
     Rs.Open strSQL, Conn
     If Rs.EOF = False Then
        If MsgBox("Are you sure you want to retest the DID?", vbYesNo, "Message") = vbYes Then
        ''   Conn.Execute ("delete  from qsms_DID_inspect where DID='" & DID & "'  ;update qsms_did set IPQCFlag='' where   DID='" & DID & "'")
         ''  Conn.Execute ("delete  from qsms_DID_inspect where DID='" & DID & "' ")
          chkIPQCRetest = True
        End If
     Else
         chkIPQCRetest = False
     End If
     
End Function


Private Function C_Interface()
       lblCompPN = ""
       lblVendor = ""
       lblVendorPN = ""
       lblLotCode = ""
       lblDateCode = ""
       lblDIDQty = ""
       lblChkNum = ""
       lblDIDUID = ""
       lblSpec = ""
       lblunit = ""
End Function

Private Function ChangSpec(Spec As String, Unit As String) As String
     Select Case Unit
          Case "R"
            Spec = "Spec(O)"
          Case "C"
            Spec = "Spec(F)"
          Case "L"
            Spec = "Spec(H)"
          Case "CON"
            Spec = "Spec(PIN)"
          Case "IC"
            Spec = "Spec(PIN)"
          Case Else
            Spec = "Spec.(+/- % ) :"
    End Select
    ChangSpec = Spec
    
    
End Function



Public Sub ReadTestfile(strFilePath As String)
 Dim FileNo As Integer
 FileNo = FreeFile()
 Open strFilePath For Input As #FileNo
    strFileContents = Input$(LOF(FileNo), #FileNo)
 Close #FileNo
 
End Sub

Public Sub DelFile(strFilePath As String)
Dim fso As New FileSystemObject
If Dir(strFilePath) <> "" Then
  fso.DeleteFile strFilePath, True
End If
End Sub

Private Sub Warning_Sound()
      wave_control.FileName = App.path & "\OO.wav"
      wave_control.Command = "open"
      wave_control.Command = "play"
      Do While wave_control.Mode = mciModePlay
      Loop
      wave_control.Command = "close"
End Sub
Private Sub OK_Sound()
    wave_control.FileName = App.path & "\OK.wav"
    wave_control.Command = "open"
    wave_control.Command = "play"
    Do While wave_control.Mode = mciModePlay
    Loop
    wave_control.Command = "close"
End Sub
Private Function MeasureE6420_RS232(sEquipType As String, sUnit As String, dblFrequency As Double, dblVoltage As Double)
    Dim strCommand(5) As String
    Dim strResult As String
    
    strCommand(0) = ":MEAS"
    strCommand(1) = ":MEAS:TEST:AC"
    strCommand(2) = ":MEAS:FUNC:" & sUnit
    strCommand(3) = ":MEAS:LEVEL " & dblVoltage & ";FREQ " & dblFrequency & ";"
    strCommand(4) = ":MEAS:RANGE AUTO"
    strCommand(5) = ":MEAS:TRIG"
    
    Call NewRs232.SetupSerialPort(MSComm1, 1)
    Call NewRs232.OpenSerialPort(MSComm1)
    
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(0))
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(1))
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(2))
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(3))
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(4))
    
    Sleep (50)
    
    Call NewRs232.SendCommand(MSComm1, 1, strCommand(5))
    Sleep (150)
    strResult = SplitString(NewRs232.strResult)
    MeasureE6420_RS232 = strResult
End Function
Private Function MeasureE8110G_RS232(sEquipType As String, sUnit As String, dblFrequency As Double, dblVoltage As Double)   ''(1024)
    Dim strCommand(5) As String
    Dim strResult As String
    
    strCommand(0) = ":MEAS:SPEED SLOW"
    strCommand(1) = ":MEAS:TEST:AC"
    strCommand(2) = ":MEAS:FUNC:C"
    strCommand(3) = ":MEAS:LEVEL " & dblVoltage & ";FREQ " & dblFrequency & ";"
    strCommand(4) = ":MEAS:RANGE AUTO"
    strCommand(5) = ":MEAS:TRIG"
    
    Call NewRs232.SetupSerialPort(MSComm1, 1)
    Call NewRs232.OpenSerialPort(MSComm1)
    
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(0))
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(1))
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(2))
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(3))
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(4))
    
    Sleep (50)
    
    Call NewRs232.SendCommand(MSComm1, 1, strCommand(5))
    Sleep (150)
    If sUnit = "C" Then
        strResult = SplitStringNum(NewRs232.strResult, 0)
    Else
        strResult = SplitStringNum(NewRs232.strResult, 1)
    End If
    MeasureE8110G_RS232 = strResult
End Function

Private Function Measure3523_RS232(sEquipType As String, sUnit As String, dblFrequency As Double, dblVoltage As Double)
    Dim strCommand(5) As String
    Dim strResult As String
    
     strCommand(0) = "*RST" + vbCrLf
     strCommand(1) = ":TRIGger EXTernal" + vbCrLf
     strCommand(2) = ":FREQuency 120" + vbCrLf
     If (sUnit = "C") Then
     strCommand(3) = ":PARameter1 CS;:PARameter2 RS" + vbCrLf
     ElseIf (sUnit = "R") Then
     strCommand(3) = ":PARameter1 RS;:PARameter2 CS" + vbCrLf
     ElseIf (sUnit = "L") Then '''(1236)
     strCommand(3) = ":PARameter1 LS;:PARameter2 RS" + vbCrLf '''(1236)
     End If
     
     strCommand(4) = "*TRG" + vbCrLf
     strCommand(5) = ":MEASure?" + vbCrLf

    
    Call NewRs232.SetupSerialPort(MSComm1, 4)
    Call NewRs232.OpenSerialPort(MSComm1)
    
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(0))
     
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(1))
     
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(2))
    
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(3))
    
    Call NewRs232.SendCommand(MSComm1, 0, strCommand(4))
     
    
    Sleep (50)
    
    Call NewRs232.SendCommand(MSComm1, 1, strCommand(5))
     
    
    Sleep (150)
    strResult = SplitString(NewRs232.strResult)
    ' MsgBox (strResult)
    Measure3523_RS232 = strResult
End Function


'''**RQ09102710  Denver      2009.10.27    Add 测试LCR 型号为4300 的仪器  （0063）
Private Function MeasureE4300_RS232(sEquipType As String, sUnit As String, dblFrequency As Double, dblVoltage As Double) As String
    Dim strCommand(3) As String
    Dim strResult As String
    
    strCommand(0) = ":MEAS:FUNC1 " & sUnit
    strCommand(1) = ":MEAS:FREQ " & dblFrequency
    strCommand(2) = ":MEAS:LEV " & dblVoltage
    strCommand(3) = ":MEAS:TRIG"
    
    If PreCompPN = "" Or PreCompPN <> UCase(lblCompPN) Then
        PreCompPN = UCase(lblCompPN)
        
        Call NewRs232.SetupSerialPort(MSComm1, 1)
        Call NewRs232.OpenSerialPort(MSComm1)

        Call NewRs232.SendCommand(MSComm1, 0, strCommand(0))
'        Sleep (Val(txtSleepTime))
        Call NewRs232.SendCommand(MSComm1, 0, strCommand(1))
'        Sleep (Val(txtSleepTime))
        Call NewRs232.SendCommand(MSComm1, 0, strCommand(2))
        
        
'        Call NewRs232.SetupSerialPort(frmConnection.MSComm1, 1)
'        Call NewRs232.OpenSerialPort(frmConnection.MSComm1)
'
'        Call NewRs232.SendCommand(frmConnection.MSComm1, 0, strCommand(0))
''        Sleep (Val(txtSleepTime))
'        Call NewRs232.SendCommand(frmConnection.MSComm1, 0, strCommand(1))
''        Sleep (Val(txtSleepTime))
'        Call NewRs232.SendCommand(frmConnection.MSComm1, 0, strCommand(2))
        
    End If

    Sleep (Val(txtSleepTime))
    Call NewRs232.SendCommand(MSComm1, 1, strCommand(3))     ''trigger
'    Call NewRs232.SendCommand(frmConnection.MSComm1, 1, strCommand(3))     ''trigger
    Sleep (Val(txtSleepTime))
    strResult = SplitString(NewRs232.strResult)
    
    MeasureE4300_RS232 = strResult
    
End Function


Function SplitString(strResult As String) As String
Dim a() As String
    If strResult = "" Then
        SplitString = ""
        Exit Function
    End If
    a = Split(strResult, ",")
    SplitString = a(0)

End Function
Function SplitStringNum(strResult As String, Num As Integer) As String
Dim a() As String
    If strResult = "" Then
        SplitStringNum = ""
        Exit Function
    End If
    a = Split(strResult, ",")
    SplitStringNum = a(Num)

End Function


 
Private Sub MSComm1_OnComm()
Dim ErrorMessage As String
    Dim COMPortEvent As Variant
    Dim COMPortError As Variant
    
    ErrorMessage = EMPTYSTR
    COMPortEvent = MSComm1.CommEvent
    
    Select Case MSComm1.CommEvent
        ' Errors
        Case comEventBreak  ' A Break was received.
            ErrorMessage = "COM Port Error - A Break was Received"
        Case comEventCDTO   ' CD (RLSD) Timeout.
            ErrorMessage = "COM Port Error - CD (RLSD) Timeout"
        Case comEventCTSTO  ' CTS Timeout.
            ErrorMessage = "COM Port Error - CTS Timeout"
        Case comEventDSRTO  ' DSR Timeout.
            ErrorMessage = "COM Port Error - DST Timeout"
        Case comEventFrame  ' Framing Error
            ErrorMessage = "COM Port Error - Framing Error"
        Case comEventOverrun    ' Data Lost.
            ErrorMessage = "COM Port Error - Data Lost"
        Case comEventRxOver ' Receive buffer overflow.
            ErrorMessage = "COM Port Error - Receive Buffer Overflow"
        Case comEventRxParity   ' Parity Error.
            ErrorMessage = "COM Port Error - Receive Buffer Overflow"
        Case comEventTxFull ' Transmit buffer full.
            ErrorMessage = "COM Port Error - Transmit Buffer Full"
        Case comEventDCB    ' Unexpected error retrieving DCB]
            ErrorMessage = "COM Port Error - Unexpected Error Retrieving DCB"
        ' Events
        Case comEvCD        ' Change in the CD line.
            COMPortEvent = comEvCD
        Case comEvCTS       ' Change in the CTS line.
            COMPortEvent = comEvCTS
        Case comEvDSR       ' Change in the DSR line.
            COMPortEvent = comEvDSR
        Case comEvRing      ' Change in the Ring Indicator.
            COMPortEvent = comEvRing
        Case comEvReceive   ' Received RThreshold # of chars.
            Sleep (80)
            
            ''Call DealWithCommResult(MSComm1.Input)
            ''-------------------------------------------------------------------------
            ''20100430    Denver   这部分代码不能提取出去，否则MSComm将监测不到信息
            NewRs232.strResult = ""
            Do
            NewRs232.strResult = NewRs232.strResult & MSComm1.Input
            DoEvents
            'DoEvents
            Loop Until (Right(NewRs232.strResult, 1) = vbLf)
        
            If cboEquipType.text = "4300" Then
                While Asc(Right(NewRs232.strResult, 1)) < 20
                    NewRs232.strResult = Left(NewRs232.strResult, Len(NewRs232.strResult) - 1)
                Wend
            End If
            
            NewRs232.COMPortEvent = comEvReceive
            
            ''-------------------------------------------------------------------------
             
        Case comEvSend      ' There are SThreshold number of characters in the transmit buffer.
            COMPortEvent = comEvSend
        Case comEvEOF       ' An EOF charater was found in the input stream
            COMPortEvent = comEvEOF
    End Select
    
    If ErrorMessage <> EMPTYSTR Then
        MsgBox ErrorMessage, vbOKOnly + vbCritical, "OnComm"
        COMPortError = True
    End If
End Sub


'Private Sub DealWithCommResult(result As String)
'    NewRs232.strResult = ""
'    Do
'        NewRs232.strResult = NewRs232.strResult & result
'        DoEvents
'    Loop Until (Right(NewRs232.strResult, 1) = vbLf)
'
'    If cboEquipType.Text = "4300" Then
'        While Asc(Right(NewRs232.strResult, 1) < 20)
'            NewRs232.strResult = Left(NewRs232.strResult, Len(NewRs232.strResult) - 1)
'        Wend
'    End If
'
'End Sub


