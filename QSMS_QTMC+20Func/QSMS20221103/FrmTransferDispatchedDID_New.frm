VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTransferDispatchedDID_New 
   Caption         =   "FrmTransferDispatchedDID"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      Height          =   2055
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   9615
      Begin VB.Frame FraSB 
         Caption         =   "Small Board WO"
         Height          =   615
         Left            =   6600
         TabIndex        =   57
         Top             =   2280
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
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
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
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2655
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "Group"
         Height          =   255
         Left            =   1680
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton OptRelease 
         Caption         =   "Release"
         Height          =   255
         Left            =   120
         TabIndex        =   44
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
         Picture         =   "FrmTransferDispatchedDID_New.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   600
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
         Left            =   6840
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   120
         Width           =   2655
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
         Left            =   6840
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2655
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1695
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
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   840
         Width           =   2655
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
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   480
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   47
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
         Format          =   160169987
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   48
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
         Format          =   160169987
         CurrentDate     =   36482
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
         TabIndex        =   56
         Top             =   600
         Width           =   1455
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
         Left            =   4680
         TabIndex        =   55
         Top             =   1200
         Width           =   2175
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
         Index           =   15
         Left            =   120
         TabIndex        =   54
         Top             =   1560
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
         Left            =   4680
         TabIndex        =   53
         Top             =   120
         Width           =   2175
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
         Left            =   4680
         TabIndex        =   52
         Top             =   1560
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
         TabIndex        =   51
         Top             =   1080
         Width           =   1455
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
         Left            =   4680
         TabIndex        =   50
         Top             =   840
         Width           =   2175
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
         Left            =   4680
         TabIndex        =   49
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame FraDispatchDID 
      BackColor       =   &H00FF80FF&
      Caption         =   "DispatchDID"
      Height          =   1335
      Left            =   0
      TabIndex        =   21
      Top             =   2040
      Width           =   9615
      Begin VB.TextBox TxtVersion 
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
         Left            =   8280
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   240
         Width           =   1215
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
         Left            =   600
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox TxtSide 
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
         Left            =   4680
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   720
         Width           =   1095
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
         Left            =   3240
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   720
         Width           =   855
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox TxtDispatchQty 
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
         Left            =   5880
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtJobPN 
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FF00&
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
         ForeColor       =   &H00008000&
         Height          =   375
         Index           =   1
         Left            =   7440
         TabIndex        =   62
         Top             =   240
         Width           =   855
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
         TabIndex        =   60
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Side"
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
         Left            =   4200
         TabIndex        =   33
         Top             =   720
         Width           =   495
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
         TabIndex        =   28
         Top             =   240
         Width           =   855
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
         Left            =   4440
         TabIndex        =   27
         Top             =   240
         Width           =   1455
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
         Index           =   1
         Left            =   2760
         TabIndex        =   26
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
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
         Index           =   30
         Left            =   6000
         TabIndex        =   25
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000000&
      Caption         =   "DID information "
      Height          =   1215
      Left            =   0
      TabIndex        =   10
      Top             =   3360
      Width           =   9615
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
         Left            =   8400
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
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
         Left            =   8040
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   720
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
         Left            =   4920
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   720
         Width           =   1815
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
         Left            =   1560
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
         Width           =   2055
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   2415
      End
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
         Left            =   4920
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   2775
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
         Index           =   17
         Left            =   7920
         TabIndex        =   31
         Top             =   240
         Width           =   495
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
         Left            =   6840
         TabIndex        =   20
         Top             =   720
         Width           =   1215
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
         Left            =   3720
         TabIndex        =   19
         Top             =   720
         Width           =   1215
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
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1455
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
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1095
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
         Left            =   3720
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Mapping Machine-Slot"
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   9615
      Begin VB.TextBox TxtLR 
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
         Left            =   1920
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox TxtSlot 
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
         Left            =   1920
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1695
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
         Left            =   1920
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cboNewmachine 
         Height          =   315
         Left            =   4200
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox CboNewSlot 
         Height          =   315
         Left            =   4200
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox CboNewLR 
         Height          =   315
         Left            =   4200
         TabIndex        =   2
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H0000FF00&
         Caption         =   "OK"
         Height          =   1215
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   1455
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
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "SLot"
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
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "LR"
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
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Source"
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
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Detination"
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
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmTransferDispatchedDID_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strWO As String

Private Sub CboGroupID_Click()
Call GetGroupWO(CboGroupID)
End Sub

Private Sub CboGroupID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboGroupID_Click
End If
End Sub

Private Sub cboNewmachine_Click()
Dim Str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset

CboNewSlot.Clear
CboNewLR.Clear

Str = "select distinct Slot from QSMS_WO where work_order ='" & Trim(TxtWO) & "' and Machine='" & Trim(cboNewmachine) & "'"

Set Rs = Conn.Execute(Str)

While Not Rs.EOF
        CboNewSlot.AddItem Trim(Rs!Slot)
        Rs.MoveNext
Wend

CboNewLR.AddItem "0"
CboNewLR.AddItem "1"
CboNewLR.AddItem "2"

End Sub

Private Sub CboNotChkBOM_Click()
TxtWO = Trim(CboNotChkBOM)
Call GetSBWO(TxtWO)
End Sub

Private Sub CboNotFinishedWO_Click()
Dim wostr As String
    TxtWO = Trim(CboNotFinishedWO)
    Call GetSBWO(TxtWO)
    wostr = GetWoArray
    Call GetMachine(TxtWO, wostr)
End Sub

Private Sub cboWO_Click()
Dim wostr As String

TxtWO = Trim(cboWO)
Call GetSBWO(TxtWO)
wostr = GetWoArray
Call GetMachine(TxtWO, wostr)
End Sub

Private Sub cmdOK_Click()
On Error GoTo TransferError
    If TxtDID.Text = "" Then
        MsgBox "Please input DID!", vbCritical, "Error"
        Exit Sub
    End If
    
    If TxtWO = "" Or TxtLine = "" Or cboNewmachine = "" Or CboNewSlot = "" Or CboNewLR = "" Or TxtCompPN = "" Or TxtMachine = "" Or TxtSlot = "" Or TxtDispatchQty = "" Then
        MsgBox "Please input the machine & slot infomation", vbCritical
        Exit Sub
    End If
    Conn.Execute "EXEC QSMS_TransferDispatchDID '" & Trim(TxtMachine) & "','" & Trim(TxtSlot) & "','" & Trim(TxtLR) & "','" & Trim(TxtWO) & "','" & Trim(TxtCompPN) & "','" & Trim(TxtJobPN) & "','" & Trim(TxtDID) & "','" & Trim(TxtDispatchQty) & "','" & Trim(cboNewmachine) & "','" & Trim(CboNewSlot) & "','" & Trim(CboNewLR) & "','" & Trim(TxtVersion) & "'"
    
    Call UpdateMachineFlagByWO(TxtWO)
    MsgBox "OK ! "
    Call ClearData
    Exit Sub
TransferError:
    MsgBox Err.Description + ",Please contact QSMS SMT Staff"
End Sub

Private Sub CmdQuery_Click()
If Trim(CboLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If
Call GetGroupID
End Sub

Private Sub Form_Load()
Dim Str As String
Dim Rs As ADODB.Recordset
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

dtpSDate = Date
dtpEDate = Date
Call GetLine
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtDID_Click()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub TxtDID_KeyPress(KeyAscii As Integer)
Dim StrPN As String
If KeyAscii = 13 Or KeyAscii = 9 Then
    If Trim(TxtDID) = "" Then Exit Sub
    TxtDID = Trim(TxtDID)
    If GetDIDInfo(Trim(TxtDID)) = False Then
        TxtDID.Text = ""
        TxtDID.SetFocus
    Else
        cboNewmachine.SetFocus
    End If
End If
End Sub

Private Function GetDIDInfo(ByVal DID As String) As Boolean
Dim Str As String
Dim Rs As ADODB.Recordset
Dim Used_Flag As String


TxtCompPN = ""
TxtVendorCode = ""
TxtDateCode = ""
TxtLotCode = ""
TxtDIDDateTime = ""
TxtDIDTotalQty = ""
TxtDispatchQty = ""
TxtLine = ""
TxtSide = ""
TxtJobPN = ""
TxtMachine = ""
TxtSlot = ""
TxtLR = ""
GetDIDInfo = True

'Get DID information
Str = "select CompPN,VendorCode,DateCode,LotCode,Qty,RemainQty,UsedFlag,TransDateTime from QSMS_DID where DID='" & Trim(TxtDID) & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
        TxtCompPN = Trim(Rs!CompPN)
        TxtVendorCode = Trim(Rs!VendorCode)
        TxtDateCode = Trim(Rs!DateCode)
        TxtLotCode = Trim(Rs!LotCode)
        TxtDIDTotalQty = Trim(Rs!Qty)
        TxtDIDDateTime = Trim(Rs!TransDateTime)
Else
    Str = "select CompPN,VendorCode,DateCode,LotCode,Qty,RemainQty,TransDateTime from QSMS_DID_log where DID='" & Trim(TxtDID) & "'"
    Set Rs = Conn.Execute(Str)
    If Not Rs.EOF Then
        TxtCompPN = Trim(Rs!CompPN)
        TxtVendorCode = Trim(Rs!VendorCode)
        TxtDateCode = Trim(Rs!DateCode)
        TxtLotCode = Trim(Rs!LotCode)
        TxtDIDTotalQty = Trim(Rs!Qty)
        TxtDIDDateTime = Trim(Rs!TransDateTime)
    Else
        MsgBox "Can't find this DID,please check!", vbCritical, "Error"
        GetDIDInfo = False
        Exit Function
    End If
End If

'Get DID dispatch infromation
Str = "select Work_Order,Line,JobPN,Machine,Slot,LR,NeedQty,DIDQty,Side,DeletedFlag from QSMS_Dispatch where DID='" & Trim(TxtDID) & "' and work_order='" & Trim(TxtWO.Text) & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
        TxtDispatchQty = Trim(Rs!DIDQty)
        TxtLine = Trim(Rs!Line)
        TxtSide = Trim(Rs!Side)
        TxtJobPN = Trim(Rs!Jobpn)
        TxtMachine = Trim(Rs!Machine)
        TxtSlot = Trim(Rs!Slot)
        TxtLR = Trim(Rs!LR)
Else
    MsgBox "Can't find this DID,please check!", vbCritical, "Error"
    GetDIDInfo = False
    Exit Function
End If

End Function

Private Function GetGroupID()
Dim Str As String
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
       Str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )"
    Else
        Str = "select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N' AND Work_Order IN (SELECT WO FROM Sap_Wo_List )"
    End If
Else
    If OptRelease.Value = True Then
       Str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N'"
    Else
        Str = "select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag='N'"
    End If
End If

Set Rs = Conn.Execute(Str)
i = 0
CboGroupID.Clear
While Not Rs.EOF
      CboGroupID.AddItem Trim(Rs!GroupID)
      Rs.MoveNext
      i = i + 1
Wend
If i = 0 Then
   MsgBox "No data"
   
End If
End Function

Private Function GetGroupWO(ByVal GroupID As String)
Dim Str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset
CboClosed.Clear
Str = "select Work_Order,ClosedFlag from QSMS_WOGroup  where GroupID= '" & GroupID & "' order by Seq_NO"

Set Rs = Conn.Execute(Str)
cboWO.Clear
CboNotFinishedWO.Clear
CboNotChkBOM.Clear
While Not Rs.EOF
      If UCase(Trim(Rs!ClosedFlag)) = "Y" Then
          CboClosed.AddItem Trim(Rs!Work_Order)
      Else
'          If ChkMBWo(Rs!Work_Order) = True Then
        If ChkQSMS_WO(Trim(Rs!Work_Order)) = False Then
            CboNotChkBOM.AddItem Trim(Rs!Work_Order)
        Else
            If ChkWoFinished(Rs!Work_Order) = True Then
                cboWO.AddItem Trim(Rs!Work_Order)
            Else
                 CboNotFinishedWO.AddItem Trim(Rs!Work_Order)
            End If
        End If
'          End If
      End If
      Rs.MoveNext
Wend
End Function

Private Function GetSBWO(ByVal WO As String)
Dim Str As String
Dim Rs As ADODB.Recordset
Dim i As Long
Dim Group As String
i = 0
CboSBWO.Clear
FraSB.Visible = False
Str = "Select MB_Rev,[Group] from Sap_Wo_List where wo='" & WO & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
    TxtVersion = Trim(Rs!MB_Rev)
    Group = Trim(Rs!Group)
End If
Str = "select Wo from Sap_Wo_list where [Group] ='" & Group & "' and wo<>'" & WO & "' order by wo"
Set Rs = Conn.Execute(Str)
While Not Rs.EOF
     CboSBWO.AddItem Trim(Rs!WO)
     Rs.MoveNext
     i = i + 1
Wend
If i > 0 Then
    FraSB.Visible = True

End If
End Function

Private Function GetWoArray() As String
Dim WoArray As String
Dim Str As String
Dim Rs As ADODB.Recordset
Dim i As Long

    Str = "select wo from Sap_WO_List where [Group] in (select [group] from sap_wo_list where wo='" & Trim(TxtWO) & "')"
        Set Rs = Conn.Execute(Str)
        While Not Rs.EOF
               WoArray = WoArray + "'" + Trim(Rs!WO) + "'" + ","
               Rs.MoveNext
        Wend
    WoArray = Mid(WoArray, 1, Len(WoArray) - 1)
    WoArray = "(" + WoArray + ")"
    GetWoArray = WoArray
End Function

Private Function GetMachine(ByVal WO As String, ByVal wostr As String)
Dim Str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset
Dim rsMachine As ADODB.Recordset
Dim tempmachine As String
Dim Rs1 As ADODB.Recordset

cboNewmachine.Clear
CboNewSlot.Clear
CboNewLR.Clear

'Str = "select distinct Machine,Slot,LR,MachinefinishedFlag from QSMS_WO where work_order ='" & Trim(TxtWO) & "'" & _
'      "order by machine,MachinefinishedFlag"
'
'Set Rs = Conn.Execute(Str)
'
'While Not Rs.EOF
'     If tempmachine = "" Or tempmachine <> UCase(Trim(Rs!Machine)) Then
'        If UCase(Trim(Rs!MachinefinishedFlag)) = "N" Then
'            cboNewmachine.AddItem Trim(Rs!Machine)
'            CboNewSlot.AddItem Trim(Rs!Slot)
'            CboNewLR.AddItem Trim(Rs!LR)
'        End If
'     End If
'     tempmachine = UCase(Trim(Rs!Machine))
'     Rs.MoveNext
'Wend
Str = "select distinct Machine,MachinefinishedFlag from QSMS_WO where work_order ='" & Trim(TxtWO) & "'" & _
      "order by machine,MachinefinishedFlag"

Set Rs = Conn.Execute(Str)

While Not Rs.EOF
        If UCase(Trim(Rs!MachinefinishedFlag)) = "N" Then
            cboNewmachine.AddItem Trim(Rs!Machine)
        End If
     Rs.MoveNext
Wend

End Function

Private Function GetLine()
Dim Str As String
Dim Rs As ADODB.Recordset
Str = "select distinct Line from QSMS_woGroup order by line"
Set Rs = Conn.Execute(Str)
CboLine.Clear
While Not Rs.EOF
    CboLine.AddItem Rs!Line
    Rs.MoveNext
Wend
End Function

Private Function ClearData()
TxtDID = ""
TxtDispatchQty = ""
TxtCompPN = ""
TxtVendorCode = ""
TxtDateCode = ""
TxtLotCode = ""
TxtDIDDateTime = ""
TxtDIDTotalQty = ""
TxtLine = ""
TxtSide = ""
TxtJobPN = ""
TxtMachine = ""
TxtSlot = ""
TxtLR = ""
End Function
