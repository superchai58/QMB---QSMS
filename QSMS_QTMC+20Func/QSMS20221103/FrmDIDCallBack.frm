VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmDIDCallBack 
   Caption         =   "DID Call Back"
   ClientHeight    =   10605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   ScaleHeight     =   10605
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      Height          =   2775
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11775
      Begin VB.Frame FraSB 
         Caption         =   "Small Board WO"
         Height          =   615
         Left            =   6600
         TabIndex        =   44
         Top             =   1080
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
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   240
            Width           =   2415
         End
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
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1800
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
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1800
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
         TabStop         =   0   'False
         Top             =   240
         Width           =   2655
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
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptRelease 
         Caption         =   "Release"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optGroup 
         Caption         =   "Group"
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2280
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2280
         Width           =   2655
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   41
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
         Format          =   25755651
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   42
         TabStop         =   0   'False
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
         Format          =   25755651
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
         TabIndex        =   43
         Top             =   480
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
         TabIndex        =   36
         Top             =   1800
         Width           =   1455
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
         TabIndex        =   35
         Top             =   1800
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
         TabIndex        =   34
         Top             =   240
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
         TabIndex        =   33
         Top             =   2280
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
         TabIndex        =   32
         Top             =   1440
         Width           =   1455
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
         TabIndex        =   31
         Top             =   2280
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
         TabIndex        =   30
         Top             =   720
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
         TabIndex        =   29
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "DID information"
      Height          =   4215
      Left            =   0
      TabIndex        =   14
      Top             =   2760
      Width           =   11775
      Begin MSDataGridLib.DataGrid DGDIDReturned 
         Height          =   1695
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2990
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
      Begin VB.TextBox TxtChkCompPN 
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3720
         Visible         =   0   'False
         Width           =   3495
      End
      Begin MSDataGridLib.DataGrid DGDIDNotReturned 
         Height          =   1695
         Left            =   3360
         TabIndex        =   38
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2990
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
      Begin MSDataGridLib.DataGrid DGDIDInfo 
         Height          =   1695
         Left            =   6600
         TabIndex        =   39
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2990
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
      Begin MSDataGridLib.DataGrid DGDispatch 
         Height          =   1215
         Left            =   120
         TabIndex        =   40
         Top             =   2280
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   2143
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
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "CompPN:"
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
         TabIndex        =   17
         Top             =   3720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label LblChk 
         BackColor       =   &H80000000&
         Height          =   495
         Left            =   5280
         TabIndex        =   16
         Top             =   3600
         Width           =   6495
      End
   End
   Begin VB.Frame FraReturnDID 
      BackColor       =   &H80000013&
      Caption         =   "Return DID"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   6960
      Width           =   9615
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
         Left            =   1920
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   840
         Width           =   4095
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
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1575
      End
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
         Left            =   1800
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtDIDReturnedQty 
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
         Left            =   5040
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1095
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1320
         Width           =   975
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
         Left            =   7440
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   8520
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   855
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
         TabIndex        =   13
         Top             =   840
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
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1575
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
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DID Called Back Qty"
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
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   2175
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
         Left            =   6240
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LblMessage 
         BackColor       =   &H80000000&
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   6495
      End
   End
End
Attribute VB_Name = "FrmDIDCallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wostr As String


Private Sub CboGroupID_Click()

If ChkGroupClosed(Trim(CboGroupID)) = True Then
   MsgBox "The Group has been closed,can not return DID"
   Exit Sub
End If
Call GetGroupWO(CboGroupID)

End Sub
Private Function ChkGroupClosed(ByVal GroupID As String) As Boolean
Dim Str As String
Dim Rs As ADODB.Recordset
ChkGroupClosed = False
Str = "select * from QSMS_WoGroup where GroupID='" & Trim(GroupID) & "' and ClosedFlag<>'Y'"
Set Rs = Conn.Execute(Str)
If Rs.EOF Then
   ChkGroupClosed = True
End If
End Function

Private Sub CboGroupID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboGroupID_Click
End If
End Sub

Private Sub CboWo_Click()
Dim Str As String
Dim Rs As ADODB.Recordset
'Dim wostr As String
TxtWO = Trim(cboWO)
Str = "Select GroupID from QSMS_WOGroup where Work_Order='" & Trim(TxtWO) & "'"
Set Rs = Conn.Execute(Str)
If Rs.EOF Then
   MsgBox "Can not find the GroupID for the Work Order:" & Trim(TxtWO)
   Exit Sub
Else
   If ChkGroupClosed(Trim(Rs!GroupID)) = True Then
      MsgBox "The Group has been closed,can not return DID"
   Exit Sub
End If
End If
Call GetSBWO(TxtWO)
Call GetWoinfo(TxtWO)
Str = "Exec QSMSDIDCallBack '" & Trim(TxtWO) & "'"
Conn.Execute Str
wostr = GetWoArray
Call GetReturned_NotReturnDID(Trim(wostr))
End Sub

Private Sub cboWO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboWo_Click
End If
End Sub







Private Sub CmdQuery_Click()
If Trim(CboLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If
Call GetGroupID
End Sub

Private Sub DGDIDOK_Click()

End Sub
Private Function GetLine()
Dim Str As String
Dim Rs As ADODB.Recordset
Str = "select distinct Line from QSMS_woGroup"
Set Rs = Conn.Execute(Str)
CboLine.Clear
While Not Rs.EOF
    CboLine.AddItem Rs!Line
    Rs.MoveNext
Wend
End Function

Private Sub cmdReturnAll_Click()
Dim Str As String
Dim Rs As ADODB.Recordset
Dim TransDateTime As String
Str = "select GetDate()"
Set Rs = Conn.Execute(Str)
TransDateTime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")

Str = "Update QSMS_GroupDID set ReturnFlag='Y',TransDateTime='" & TransDateTime & "' ,UID='" & g_userName & "' where GroupID='" & Trim(CboGroupID) & "' and ReturnFlag<>'Y'"
Conn.Execute Str

Str = "Update QSMS_WoGroup Set ClosedType='Auto',ClosedFlag='Y' where GroupID='" & Trim(CboGroupID) & "' and ClosedFlag='N' "
Conn.Execute Str

Str = "exec QSMSSap2 '" & Trim(CboGroupID) & "'"
Conn.Execute Str
End Sub

Private Sub cmdSave_Click()
On Error GoTo EcmdSave_Click
If ChkErr = False Then
   Exit Sub
End If
Call UpdateReturnQty(Trim(wostr), Trim(TxtCompPN), Trim(TxtDID), Trim(TxtReturnQty))
LblMessage.Caption = "Call OK"
LblMessage.BackColor = &H80FF80
Call GetReturned_NotReturnDID(Trim(wostr))

Call GetDIDInfo(Trim(TxtDID), Trim(cboWO))
TxtDID.Text = ""
TxtReturnQty.Text = ""
TxtDID.SetFocus
Exit Sub
EcmdSave_Click:
    MsgBox Err.Description + ",Please contact QSMS SMT Staff"

End Sub

Private Sub Command1_Click()
'Call UpdateQSMSGroupCompQty("A200608150001")
End Sub

Private Sub DGDIDInfo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim Str As String
Dim Rs As ADODB.Recordset
On Error Resume Next
        With DGDIDInfo
             TxtDID = Trim(.Columns(0).Text)
             TxtDIDTotalQty.Text = Trim(.Columns(2).Text)
             TxtCompPN = Trim(.Columns(1).Text)
             TxtDIDReturnedQty.Text = Trim(.Columns(3).Text)
             Str = "select * from QSMS_Dispatch where DID='" & TxtDID & "' and work_order in (select work_order from qsms_WoGroup where  GroupID='" & CboGroupID & "')"
             Set Rs = Conn.Execute(Str)
             Set DGDispatch.DataSource = Rs
             If Err.Number <> 0 Then
                TxtDIDTotalQty.Text = vbNullString
                TxtDIDReturnedQty.Text = vbNullString
             End If

          
        End With

End Sub

Private Sub DGDIDNotReturned_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
        With DGDIDNotReturned
            
             TxtDID = Trim(.Columns(0).Text)
            
             Call TxtChkCompPN_KeyPress(13)
        End With
End Sub


Private Sub DGDIDReturned_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'On Error Resume Next
'        With DGDIDReturned
'
'             TxtChkCompPN = Trim(.Columns(0).Text)
'
'             Call TxtChkCompPN_KeyPress(13)
'        End With
'
        
        On Error Resume Next
        With DGDIDReturned
            
             TxtDID = Trim(.Columns(0).Text)
            
             Call TxtChkCompPN_KeyPress(13)
        End With
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
Dim Str As String
Dim Rs As ADODB.Recordset
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
Str = "select getdate()"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
    'Date = Rs(0)
    'Time = Rs(0)
End If
dtpSDate = Date
dtpEDate = Date
Call GetLine

End Sub

Private Function DeleteDID(ByVal GroupDID As String)
Dim Str As String
Dim Rs As ADODB.Recordset

Str = "Update QSMS_Dispatch Set DeletedFlag='Y' from QSMS_Dispatch A,QSMS_GroupDID B where B.GroupID='" & Trim(GroupDID) & "' and a.DID= b.DID and (b.Returnflag<>'Y' or b.ReturnQty=0)"
Conn.Execute Str

Str = "Delete   QSMS_DID  from QSMS_DID a,QSMS_GroupDID b where  B.GroupID='" & Trim(GroupDID) & "' and a.DID= b.DID and (b.Returnflag<>'Y' or b.ReturnQty=0)"
Conn.Execute Str

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

If OptRelease.Value = True Then
   Str = "select distinct GroupID from QSMS_WOGroup  where WO_TransDateTime between  '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag<>'Y'"
Else
    Str = "select distinct GroupID from QSMS_WOGroup  where substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' and line='" & CboLine & "' and closedflag<>'Y'"
End If

Set Rs = Conn.Execute(Str)
CboGroupID.Clear
If Rs.EOF Then MsgBox "No data"
While Not Rs.EOF
      CboGroupID.AddItem Trim(Rs!GroupID)
      Rs.MoveNext
Wend
End Function

Private Function GetGroupWO(ByVal GroupID As String)
Dim Str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset

Str = "select Work_Order,ClosedFlag from QSMS_WOGroup  where GroupID= '" & GroupID & "'"

Set Rs = Conn.Execute(Str)

cboWO.Clear
While Not Rs.EOF
     
          cboWO.AddItem Trim(Rs!Work_Order)
      
      Rs.MoveNext
Wend
End Function


Private Function GetWoinfo(ByVal WO As String)
Dim Str As String
Dim Rs As ADODB.Recordset
Str = "select PN, Qty from Sap_Wo_List where WO='" & Trim(WO) & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   TxtMBPN = Rs!PN
   TxtWOQty = Rs!Qty
End If
Str = "select Customer from ModelName where PN='" & TxtMBPN & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   TxtCustomer = Trim(Rs!Customer)
End If
End Function

Private Function GetReturned_NotReturnDID(ByVal Work_Order As String)
Dim Str As String
Dim Rs As ADODB.Recordset

'Str = "select DID,TotalQty,ReturnQty from QSMS_GroupDID where GroupID='" & GroupID & "' and ReturnFlag='Y'"
Str = "select distinct DID from QSMS_DIDCallBack where Work_Order in " & Work_Order & " and ReturnFlag='Y' order by DID"
Set Rs = Conn.Execute(Str)
Set DGDIDReturned.DataSource = Rs
DGDIDReturned.Caption = "(Call Back DID)  Total: " & Rs.RecordCount
'Str = "select DID,TotalQty,ReturnQty from QSMS_GroupDID where GroupID='" & GroupID & "' and ReturnFlag<>'Y'"
Str = "select Distinct DID from QSMS_DIDCallBack where work_order in " & Work_Order & " and ReturnFlag<>'Y' Order by DID"
Set Rs = Conn.Execute(Str)

Set DGDIDNotReturned.DataSource = Rs
DGDIDNotReturned.Caption = "(Not Call Back DID ) Total: " & Rs.RecordCount
LblChk.Caption = ""

End Function



Private Sub TxtChkCompPN_KeyPress(KeyAscii As Integer)
Dim Str As String
Dim Rs As ADODB.Recordset
Str = "select DID,ComppN,TotalQty,ReturnQty from QSMS_DIDCallBack where work_order in " & Trim(wostr) & " and DID='" & Trim(TxtDID) & "'"
Set Rs = Conn.Execute(Str)
If Rs.EOF Then
  LblChk.Caption = "The CompPN does not belong to the work Order"
Else
   TxtCompPN = Rs!CompPN
   Set DGDIDInfo.DataSource = Rs
   
End If

End Sub


Private Sub TxtDID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   If ChkDIDBelongToGroupID(Trim(CboGroupID), Trim(TxtDID)) = False Then
      
      Exit Sub
   End If
   Call GetDIDInfo(Trim(TxtDID), Trim(CboGroupID))
   TxtReturnQty.Text = ""
   TxtReturnQty.SetFocus
   LblMessage.BackColor = &H80000000
   LblMessage.Caption = ""
End If
End Sub

Private Function GetDIDInfo(ByVal DID As String, ByVal GroupID As String)
Dim Str As String
Dim Rs As ADODB.Recordset
Str = "Select Qty,RemainQty,CompPN from QSMS_DID where DID='" & DID & "' "
Set Rs = Conn.Execute(Str)
TxtDIDTotalQty = ""
TxtDIDReturnedQty = ""
TxtCompPN = ""

If Not Rs.EOF Then
   TxtDIDTotalQty = Trim(Rs!Qty)
   TxtCompPN = Trim(Rs!CompPN)
   TxtDIDReturnedQty = Trim(Rs!RemainQty)
End If



End Function


Private Function ChkDIDBelongToGroupID(ByVal GroupID As String, ByVal DID As String) As Boolean
Dim Str As String
Dim Rs As ADODB.Recordset
ChkDIDBelongToGroupID = True
Str = "Select DID,ReturnFlag from QSMS_DIDCallBack where work_order in " & Trim(wostr) & " and DID='" & Trim(DID) & "'"
Set Rs = Conn.Execute(Str)
If Rs.EOF Then
  ChkDIDBelongToGroupID = False
   MsgBox "The DID does not belong to the work order,Please check"
Else
'   If Trim(Rs!ReturnFlag) = "Y" Then
'       ChkDIDBelongToGroupID = False
'       MsgBox "The DID has been Call Back,Please check"
'   End If
End If

End Function

Private Function ChkErr() As Boolean
Dim Str As String
Dim Rs As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
ChkErr = True
If ChkDIDBelongToGroupID(Trim(CboGroupID), Trim(TxtDID)) = False Then
  
   ChkErr = False
End If
If Trim(TxtDIDTotalQty) = "" Or Trim(TxtCompPN) = "" Then
   MsgBox "DID total Qty or comppN Can not be empty,Please press enter key in DID txtbox"
   ChkErr = False
End If
If Trim(TxtReturnQty) = "" Or IsNumeric(TxtReturnQty) = False Then
   MsgBox "The Return Qty can not be empty or must be numeric"
   ChkErr = False
   Exit Function
End If
If CLng(TxtDIDTotalQty) < CLng(TxtReturnQty) Then
   MsgBox "Return Qty can not larger than total qty"
   ChkErr = False
End If
Str = "Select SUM(DIDQty) as PCBQty from QSMS_Dispatch where work_order in " & wostr & " and DID='" & Trim(TxtDID) & "'"
Set Rs = Conn.Execute(Str)

Str = "select * from qsms_dispatch where did='" & Trim(TxtDID) & "' and deletedFlag='N' and work_order not in " & wostr & ""
Set rsTemp = Conn.Execute(Str)

If IsNull(Trim(Rs.Fields(0))) = True Then
    MsgBox "This DID is not dispatching!"
    ChkErr = False
    Exit Function
End If

If rsTemp.EOF = False And CLng(TxtReturnQty) > Rs!PCBQty Then
    MsgBox "This DID has dispatched to more than one PCB,return Qty can not larger than the dispatched Qty : " & Rs!PCBQty & " of one PCB!", vbInformation
    ChkErr = False
End If




End Function

Private Function UpdateReturnQty(ByVal WO As String, CompPN As String, DID As String, ReturnQty As Long)
Dim Str As String
Dim Rs As ADODB.Recordset
Dim TransDateTime As String
Dim Qty, TempDIDQty, RemainQty, TempBalanceQty As Long
TempBalanceQty = 0
RemainQty = 0
TempDIDQty = 0
Qty = 0

Str = "select GetDate()"
Set Rs = Conn.Execute(Str)
TransDateTime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
Str = "Select Qty, RemainQty from QSMS_DID where DID='" & Trim(DID) & "'"
Set Rs = Conn.Execute(Str)

If Rs.EOF Then
   MsgBox "Can not find the DID,Please check"
   Exit Function
Else
    Qty = Rs!Qty
    RemainQty = Rs!RemainQty
End If
If ReturnQty = Qty Then  '''means call back the whole DID
    Str = "Insert into QSMS_Dispatch_bak(Work_Order,Line,WoQty,JobPN ,Machine,CompPN ,Slot,BaseQty,NeedQty, DID,DIDQty ,VendorCode,DateCode,LotCode ,UID,TransDateTime,DeletedFlag ) " & _
      "select Work_Order,Line,WoQty,JobPN ,Machine,CompPN ,Slot,BaseQty,NeedQty, DID,DIDQty ,VendorCode,DateCode,LotCode ,UID,TransDateTime,DeletedFlag  from QSMS_Dispatch" & _
      " where work_Order in " & wostr & " and did='" & Trim(DID) & "'"
    Conn.Execute Str
    Str = "Select work_order,Machine,Slot,LR ,NeedQty,DIDQty from QSMS_Dispatch where work_order in " & wostr & " and DID='" & DID & "'"
    Set Rs = Conn.Execute(Str)
    Do While Not Rs.EOF
               
        Str = "update qsms_wo set DispatchQty=DispatchQty-" & Rs!didqty & ",BalanceQty=DispatchQty-" & Rs!didqty & "-NeedQty,MachineFinishedFlag='N',WoFinishedFlag='N' where work_order = '" & Rs!Work_Order & "' " & _
              " and Machine='" & Trim(Rs!Machine) & "' and Slot='" & Rs!Slot & "' and LR='" & Rs!LR & "'"
        Conn.Execute Str
    Rs.MoveNext
    Loop
    
    
    Str = "Delete from QSMS_Dispatch where work_order in " & wostr & " and did='" & Trim(DID) & "'"
    Conn.Execute Str
    'maybe need add a table to record the did and delete the did from fuji database---20061009
    Str = "Delete from QSMS_DID where did='" & Trim(DID) & "'"
    Conn.Execute Str
    'total callback log
    Str = "Update QSMS_DIDCallBack set ReturnQty=ReturnQty+" & ReturnQty & ",ReturnFlag='Y',TransDateTime='" & TransDateTime & "' ,UID='" & g_userName & "' where work_order='" & cboWO & "' and DID='" & DID & "'"
    Conn.Execute Str
    
Else '''means call back some of the DID material
    
    If ReturnQty <= RemainQty Then  'means the did Qty is enough for the dispatched wo, doesn't need update the dispatched Qty and balanceqty,just only update the did total qty
       Str = "Update QSMS_DID set Qty=Qty-" & ReturnQty & ",RemainQty=RemainQty-" & ReturnQty & " where DID='" & Trim(DID) & "'"
       Conn.Execute Str
    '''need to update the related DID total Qty in table qsms_dispatch
       Str = "Update QSMS_Dispatch set TotalQty=TotalQty-" & ReturnQty & " from QSMS_dispatch a, QSMS_DID b where a.DID='" & Trim(DID) & "' and a.DID=b.DID and a.DIDDateTime=b.transdatetime"
       Conn.Execute Str
       'parts callback log
       Str = "Update QSMS_DIDCallBack set ReturnQty=ReturnQty+" & ReturnQty & ",ReturnFlag='Y',TransDateTime='" & TransDateTime & "' ,UID='" & g_userName & "' where work_order='" & cboWO & "' and DID='" & DID & "'"
       Conn.Execute Str
    Else  'means need update dispatched qty and balanceqty and need redispatch other DID to the work order
       TempBalanceQty = ReturnQty - RemainQty
       Str = "Select work_order,Machine,Slot,LR ,NeedQty,DIDQty from QSMS_Dispatch where work_order in " & wostr & " and DID='" & DID & "'"
       Set Rs = Conn.Execute(Str)
       If Not Rs.EOF Then
         Do While Not Rs.EOF
            If Rs!didqty > TempBalanceQty Then
                Str = "update qsms_wo set DispatchQty=DispatchQty-" & TempBalanceQty & ",BalanceQty=DispatchQty-" & TempBalanceQty & "-NeedQty,MachineFinishedFlag='N',WoFinishedFlag='N' where work_order='" & Trim(Rs!Work_Order) & "' " & _
                  " and Machine='" & Trim(Rs!Machine) & "' and Slot='" & Rs!Slot & "' and LR='" & Rs!LR & "'"
                Conn.Execute Str
            
                Str = "Update QSMS_Dispatch set DIDQty=didqty-" & TempBalanceQty & " where work_order='" & Trim(Rs!Work_Order) & "' and DID='" & DID & "'"
                Conn.Execute Str
                
                Str = "Update QSMS_DIDCallBack set ReturnQty=ReturnQty+" & TempBalanceQty & ",ReturnFlag='Y',TransDateTime='" & TransDateTime & "' ,UID='" & g_userName & "' where work_order='" & Rs!Work_Order & "' and DID='" & DID & "'"
                Conn.Execute Str
                
                Exit Do
            Else
               
                Str = "update qsms_wo set DispatchQty=DispatchQty-" & Rs!didqty & ",BalanceQty=DispatchQty-" & Rs!didqty & "-NeedQty,MachineFinishedFlag='N',WoFinishedFlag='N' where work_order='" & Trim(Rs!Work_Order) & "' " & _
                " and Machine='" & Trim(Rs!Machine) & "' and Slot='" & Rs!Slot & "' and LR='" & Rs!LR & "'"
                Conn.Execute Str
            
                Str = "Update QSMS_Dispatch set DIDQty=didqty-" & Rs!didqty & " where work_order='" & Trim(Rs!Work_Order) & "' and DID='" & DID & "'"
                Conn.Execute Str
                
                Str = "Update QSMS_DIDCallBack set ReturnQty=ReturnQty+" & Rs!didqty & ",ReturnFlag='Y',TransDateTime='" & TransDateTime & "' ,UID='" & g_userName & "' where work_order='" & Rs!Work_Order & "' and DID='" & DID & "'"
                Conn.Execute Str
                
                TempBalanceQty = TempBalanceQty - Rs!didqty
                Rs.MoveNext
               
            End If
         Loop
            Str = "Update QSMS_DID set Qty=Qty-" & ReturnQty & ",RemainQty=0,UsedFlag='Y' where DID='" & Trim(DID) & "'"
            Conn.Execute Str
            
            Str = "Update QSMS_Dispatch set TotalQty=TotalQty-" & ReturnQty & " from QSMS_dispatch a, QSMS_DID b where a.DID='" & Trim(DID) & "' and a.DID=b.DID and a.DIDDateTime=b.transdatetime"
            Conn.Execute Str
            
       End If
    
       
    End If
    
   
End If

'Str = "Update QSMS_DIDCallBack set ReturnQty=ReturnQty+" & ReturnQty & ",ReturnFlag='Y',TransDateTime='" & TransDateTime & "' ,UID='" & g_userName & "' where work_order='" & cboWO & "' and DID='" & DID & "'"
'Conn.Execute Str

End Function

Private Sub TxtDID_LostFocus()
Call TxtDID_KeyPress(13)
End Sub

Private Sub TxtReturnQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call cmdSave_Click
End If
End Sub

Private Function GetSBWO(ByVal WO As String)
Dim Str As String
Dim Rs As ADODB.Recordset
Dim i As Long
Dim Group As String
i = 0
CboSBWO.Clear
FraSB.Visible = False
Str = "Select [Group] from Sap_Wo_List where wo='" & WO & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   Group = Trim(Rs!Group)
   'TxtGroup = Group
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
    
'    For i = 1 To ListWoDispatching.ListCount
'        ListWoDispatching.ListIndex = i - 1
'        Str = "select wo from Sap_WO_List where [Group] in (select [group] from sap_wo_list where wo='" & ListWoDispatching.Text & "')"
'        Set Rs = Conn.Execute(Str)
'        While Not Rs.EOF
'               WoArray = WoArray + "'" + Trim(Rs!Wo) + "'" + ","
'               Rs.MoveNext
'        Wend
'        WoArray = WoArray + "'" + ListWoDispatching.Text + "'" + ","
        
'    Next i
    WoArray = Mid(WoArray, 1, Len(WoArray) - 1)
    WoArray = "(" + WoArray + ")"
    GetWoArray = WoArray
End Function

