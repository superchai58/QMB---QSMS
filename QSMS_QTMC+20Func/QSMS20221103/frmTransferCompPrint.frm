VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmTransferCompPrint 
   BackColor       =   &H8000000B&
   Caption         =   "TransferCompPrint 2023/11/04"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkPrint 
      Caption         =   "Change PN"
      Height          =   495
      Left            =   7080
      TabIndex        =   39
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      Caption         =   "Print Label"
      Height          =   2745
      Left            =   8160
      TabIndex        =   35
      Top             =   1800
      Width           =   1935
      Begin VB.CommandButton btpl 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   38
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtlp 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   960
         TabIndex        =   36
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "PrintQty:"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1665
      Left            =   3480
      TabIndex        =   19
      Top             =   30
      Width           =   6600
      Begin VB.TextBox txtModbusStatus 
         BackColor       =   &H000000FF&
         Height          =   300
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "Disconnected"
         Top             =   1280
         Width           =   2100
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   3480
         TabIndex        =   42
         Top             =   1080
         Width           =   1600
      End
      Begin VB.TextBox txtModbusIP 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1200
         TabIndex        =   41
         Text            =   "192.168.127.254"
         Top             =   920
         Width           =   2100
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   6000
         Top             =   960
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   6000
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemoteHost      =   "192.168.127.254"
         RemotePort      =   502
      End
      Begin MSCommLib.MSComm MSComm 
         Left            =   5280
         Top             =   600
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
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
         Height          =   735
         Left            =   3840
         Picture         =   "frmTransferCompPrint.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   1030
      End
      Begin VB.TextBox TxtComm 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1200
         TabIndex        =   25
         Text            =   "9600,N,8,1"
         Top             =   560
         Width           =   2100
      End
      Begin VB.TextBox TxtCompPort 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1200
         TabIndex        =   24
         Text            =   "1"
         Top             =   200
         Width           =   2100
      End
      Begin VB.Label Label14 
         Caption         =   "Status："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   43
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label13 
         Caption         =   "ModbusIP："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   40
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "Setting："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   23
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "CompPort："
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   210
         TabIndex        =   22
         Top             =   240
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1665
      Left            =   1690
      TabIndex        =   18
      Top             =   30
      Width           =   1790
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
         Height          =   300
         Left            =   210
         TabIndex        =   21
         Top             =   1080
         Width           =   1200
      End
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
         Height          =   300
         Left            =   210
         TabIndex        =   20
         Top             =   360
         Value           =   -1  'True
         Width           =   1200
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2115
      Left            =   0
      TabIndex        =   4
      Top             =   4680
      Width           =   10080
      _ExtentX        =   17780
      _ExtentY        =   3731
      _Version        =   393216
      BackColor       =   16777152
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
   Begin VB.Frame Frame2 
      Caption         =   "Information"
      Height          =   2745
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   8040
      Begin VB.TextBox TxtStandard 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2160
         Width           =   2100
      End
      Begin VB.TextBox TxtMark 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1080
         TabIndex        =   31
         Top             =   2160
         Width           =   2100
      End
      Begin VB.TextBox TxtUserID 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   300
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1680
         Width           =   2100
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   7080
         Picture         =   "frmTransferCompPrint.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   840
         Width           =   765
      End
      Begin VB.TextBox TxtLotCode 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   4560
         TabIndex        =   14
         Top             =   1200
         Width           =   2100
      End
      Begin VB.TextBox TxtVendorCode 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   4560
         TabIndex        =   13
         Top             =   720
         Width           =   2100
      End
      Begin VB.TextBox TxtQty 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1080
         TabIndex        =   10
         Top             =   1680
         Width           =   2100
      End
      Begin VB.TextBox TxtDateCode 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1080
         TabIndex        =   9
         Top             =   1200
         Width           =   2100
      End
      Begin VB.TextBox TxtCompPN 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1080
         TabIndex        =   8
         Top             =   720
         Width           =   2100
      End
      Begin VB.TextBox TxtDID 
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1080
         TabIndex        =   33
         Top             =   720
         Width           =   2100
      End
      Begin VB.Label lbCustomer 
         Caption         =   "lbCustomer"
         Height          =   255
         Left            =   1200
         TabIndex        =   46
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label15 
         Caption         =   "Customer:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label10 
         Caption         =   "Standard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3360
         TabIndex        =   29
         Top             =   2160
         Width           =   765
      End
      Begin VB.Label Label8 
         Caption         =   "Mark"
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
         TabIndex        =   30
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3405
         TabIndex        =   15
         Top             =   1680
         Width           =   765
      End
      Begin VB.Label Label7 
         Caption         =   "LotCode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3360
         TabIndex        =   12
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label Label6 
         Caption         =   "VendorCode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3360
         TabIndex        =   11
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label5 
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label Label4 
         Caption         =   "DateCode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "CompPN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label11 
         Caption         =   "DID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Printer"
      Height          =   1665
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   1695
      Begin VB.OptionButton optNetWork 
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
         Left            =   240
         TabIndex        =   28
         Top             =   1200
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
         Height          =   300
         Left            =   210
         TabIndex        =   2
         Top             =   750
         Width           =   1200
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
         Height          =   300
         Left            =   210
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1200
      End
   End
   Begin VB.Label LblMessage 
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   27
      Top             =   6960
      Width           =   9960
   End
End
Attribute VB_Name = "frmTransferCompPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim transdatetime As String
Dim strDelaytime As Long
Dim MbusQuery(11) As Byte
Public MbusResponse As String
Dim MbusByteArray(255) As Byte
Public MbusRead As Boolean
Public MbusWrite As Boolean
Dim ModbusTimeOut As Integer
Dim ModbusWait As Boolean
Dim ModbusData(64) As String
Dim CloseDateTime As Date


Private Sub btpl_Click()
  Dim LabelFile As String
  
  If Trim(txtlp) = "" Or IsNumeric(txtlp) = False Then
     LblMessage = "The printQty input format is wrong!"
     Exit Sub
  End If
  Call Printlp(txtlp)
  txtlp = ""
  
End Sub

Private Sub CmdCommSave_Click()
SaveSetting "SMT", "QSMS", "CommPort", TxtCompPort
SaveSetting "SMT", "QSMS", "Comm", TxtComm
End Sub

''1290
Private Sub cmdConnect_Click()
    Dim StartTime
    If (Winsock1.State <> sckClosed) Then
        Winsock1.Close
    End If
    Winsock1.RemoteHost = txtModbusIP.text
    Winsock1.Connect
    
    StartTime = Timer
    
    Do While ((Timer < StartTime + 2) And (Winsock1.State <> 7))
        DoEvents
    Loop
    If (Winsock1.State = 7) Then
       txtModbusStatus.text = "Connected"
       txtModbusStatus.BackColor = &HFF00&
       Timer1.Enabled = True
    Else
       txtModbusStatus.text = "Can't connect to " + txtModbusIP.text
       txtModbusStatus.BackColor = &HFF
    End If
End Sub

Private Sub CmdPrint_Click()

    Dim sSql As String
    Dim Rs As ADODB.Recordset
    Dim sCompPN As String
    'Dim intQty As Integer
    Dim intQty As Long   '(1180)
    
    sSql = "select getdate()"
    Set Rs = Conn.Execute(sSql)
    transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")

    If ValidateData = False Then
       Exit Sub
    End If

    LblMessage = "Message"
    
    If DateDiff("s", PrintTime, Now) <= CompPrintTimeSpan Then
        LblMessage.Caption = "Message: Please wait " & CompPrintTimeSpan & " sec before continuing to operate the Print."
        Exit Sub
    End If
    
'    If LabelPrintCheck = "Y" And ChkPrint.Value = 1 Then  ''1274
'        sSql = "exec QSMS_CompPrintCheck @CompPN='" & Trim(TxtCompPN.text) & "',@VendorCode='" & Trim(TxtVendorCode.text) & "',@DateCode='" & Trim(TxtDateCode.text) & "',@LotCode='" & Trim(TxtLotCode.text) & "',@Qty='" & Trim(TxtQty.text) & "',@UserID='" & Trim(TxtUserID.text) & "',@Mark='" & Trim(TxtMark.text) & "',@Standard='" & Trim(TxtStandard.text) & "'"
'        Set Rs = Conn.Execute(sSql)
'        If Not Rs.EOF Then
'            If Rs("Result") <> "" Then
'                TxtCompPN.text = Rs("Result")
'            End If
'        End If
'    Else
        sCompPN = Trim(TxtCompPN)
        intQty = Trim(TxtQty)
        
        sSql = "insert into CompPrintLog(CompPN,Qty,VendorCode,DateCode,LotCode,OPID,TransDateTime,Mark) " & _
            "values('" & UCase(Trim(sCompPN)) & "','" & UCase(Trim(intQty)) & "','" & UCase(Trim(TxtVendorCode)) & "','" & UCase(Trim(TxtDateCode)) & "','" & UCase(Trim(TxtLotCode)) & "','" & UCase(Trim(TxtUserID)) & "','" & UCase(Trim(transdatetime)) & "','" & UCase(Trim(TxtMark)) & "')"  '1117
        Set Rs = Conn.Execute(sSql)
'    End If
    
    Call PrintLabel

    Call reFreshData
    
    If StrBU = "NB5" Then ''1262
       TxtDID = ""
    End If
    
'    LblMessage.Caption = "Message: Please wait " & CompPrintTimeSpan & " sec before continuing to operate the Print."
'    Call Sleep(CompPrintTimeSpan)  ''1289
'    LblMessage.Caption = "Message: Print OK."
'
    TxtCompPN = ""
    TxtVendorCode = ""
    TxtDateCode = ""
    TxtLotCode = ""
    TxtQty = ""
    TxtMark = ""
    TxtStandard = ""
    
    TxtQty.Locked = False
    If StrBU = "NB5" Then      ''1262
       TxtDID.SetFocus
    Else
       TxtCompPN.SetFocus
    End If
    ''TxtCompPN.SetFocus
        
End Sub

''1290 送資料
Private Sub Timer1_Timer()
    Dim StartLow As Byte
    Dim StartHigh As Byte
    Dim LengthLow As Byte
    Dim LengthHigh As Byte
    If (Winsock1.State = 7) Then
        MbusQuery(0) = 0
        MbusQuery(1) = 0
        MbusQuery(2) = 0
        MbusQuery(3) = 0
        MbusQuery(4) = 0
        MbusQuery(5) = 6
        MbusQuery(6) = 1
        MbusQuery(7) = 2
        'MbusQuery(8) = StartHigh
        'MbusQuery(9) = StartLow
        'MbusQuery(10) = LengthHigh
        'MbusQuery(11) = LengthLow
        MbusQuery(8) = 0
        MbusQuery(9) = 0
        MbusQuery(10) = 0
        MbusQuery(11) = 8
        MbusRead = True
        'MbusQuery = Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(0) + Chr(6) + Chr(1) + Chr(3) + Chr(StartHigh) + Chr(StartLow) + Chr(LengtHigh) + Chr(LengthLow)
        Winsock1.SendData MbusQuery
        ModbusTimeOut = 0
    Else
        MsgBox ("Device not connected via TCP/IP")
        txtModbusStatus.text = "Disconnected"
        txtModbusStatus.BackColor = &HFF
        Timer1.Enabled = False
        TxtCompPN.Enabled = False
        TxtDID.Enabled = False
        TxtVendorCode.Enabled = False
        TxtDateCode.Enabled = False
        TxtLotCode.Enabled = False
        TxtQty.Enabled = False
        TxtUserID.Enabled = False
        TxtMark.Enabled = False
        TxtStandard.Enabled = False
        CmdPrint.Enabled = False
    End If
End Sub

''1290 取資料
Private Sub Winsock1_DataArrival(ByVal datalength As Long)
    Dim B As Byte
    Dim j As Byte
    Dim i As Integer
    For i = 1 To datalength
        Winsock1.GetData B
        MbusByteArray(i) = B
    Next
    j = 0
    If MbusRead Then
        For i = 10 To MbusByteArray(9) + 9 Step 2
            ModbusData(j) = str((MbusByteArray(i) * 256) + MbusByteArray(i + 1))
            j = j + 1
        Next i
    End If
    txtModbusStatus.text = "Discrete Read"
    If ModbusData(0) <> " 0" Then
        TxtCompPN.Enabled = False
        TxtDID.Enabled = False
        TxtVendorCode.Enabled = False
        TxtDateCode.Enabled = False
        TxtLotCode.Enabled = False
        TxtQty.Enabled = False
        TxtUserID.Enabled = False
        TxtMark.Enabled = False
        TxtStandard.Enabled = False
        CmdPrint.Enabled = False
    Else
        TxtCompPN.Enabled = True
        TxtDID.Enabled = True
        TxtVendorCode.Enabled = True
        TxtDateCode.Enabled = True
        TxtLotCode.Enabled = True
        TxtQty.Enabled = True
        TxtUserID.Enabled = True
        TxtMark.Enabled = True
        TxtStandard.Enabled = True
        CmdPrint.Enabled = True
    End If
End Sub

Private Function ValidateData() As Boolean
      
    Dim sSql As String
    Dim Rs As ADODB.Recordset
    ValidateData = False
    If Trim(TxtUserID) = "" Then
        LblMessage = "UserID is blank!!"
        Exit Function
    End If

    If Trim(TxtCompPN) = "" Then
        LblMessage = "CompPN is blank!!"
        Exit Function
    End If

    If Len(Trim(TxtCompPN)) < 11 Then
        LblMessage = "The CompPN's length must be > 11 !!"
        Exit Function
    End If
    
    If Trim(TxtVendorCode) = "" Then
        LblMessage = "VendorCode is blank!!"
        Exit Function
    End If
    
    If Trim(TxtDateCode) = "" Then
        LblMessage = "DateCode is blank!!"
        Exit Function
    End If
    
    If Trim(TxtLotCode) = "" Then
        LblMessage = "LotCode is blank!!"
        Exit Function
    End If

    If Trim(TxtQty) = "" Or IsNumeric(TxtQty) = False Then
       LblMessage = "The Qty can not be empty or must be numeric!"
       Exit Function
    End If
    
    TxtQty = Abs(Int(Trim(TxtQty)))
    
    If Trim(TxtQty) <= 0 Then
        LblMessage = "The Qty must be >0 !!"
        Exit Function
    End If
  
    If Trim(TxtCompPort) = "" Or Trim(TxtComm) = "" Then
        LblMessage = "Printer have not set!!"
        Exit Function
    End If

    ValidateData = True
    
End Function

Private Function Printlp(Qty As Integer)
Dim i As Integer
Dim j As Integer
Dim M As Integer
Dim tmpPrintStr As String
Dim LabelFile As String
Dim isZebra As Boolean
Dim lptPort As Integer
        
On Error GoTo errHandler

        LabelFile = Settings.KFLabel
        
        If Dir(LabelFile) = vbNullString Then
            MsgBox ("Can not find label file !"), vbCritical
            PrintLabel = "PRN_FileNoExist"
            Exit Function
        End If
        
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
        
        If FileReadAll(tmpPrintStr, LabelFile) <= 0 Then    '(1119)
            MsgBox "Open file:" & LabelFile & " fail!!", vbCritical
            Exit Function
        End If
        
        Select Case Trim(tmpPrintStr)
           Case vbNullString
           Case Else
                If OptComp.Value = True Then
                   For j = 1 To Qty
                      For i = 1 To Len(tmpPrintStr) Step 100
                          MSComm.Output = Mid(tmpPrintStr, i, 100)
                          DoEvents
                      Next i
                   Next j
                    MSComm.PortOpen = False
                ElseIf OptPrint.Value = True Then
                    For j = 1 To Qty
                        For i = 1 To Len(tmpPrintStr) Step 50
                           Print #lptPort, Mid(tmpPrintStr, i, 50)
                           DoEvents
                        Next i
                    Next j
                    Close #lptPort
                Else
                    For j = 1 To Qty
                       Printer.Print tmpPrintStr
                       Printer.EndDoc
                       Printer.KillDoc
                       For M = 1 To 5000
                       Next M
                    Next j
                End If
        End Select
        
        Exit Function
        
errHandler:
        MsgBox Err.Description
        If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
        End If
 
End Function
Private Function PrintLabel()
Dim hFile As Long
Dim i As Integer
Dim tmpPrintStr As String
Dim hString As String
Dim StrPN As String, tempPN As String, strVendor As String, strQty As String, strUserID As String, strMark As String
Dim strStandard As String   '添加新的栏位显示刷入非2D Barcode 信息 1249
Dim strDay As String
Dim strLot As String
Dim LabelFile As String
Dim isZebra As Boolean
Dim lptPort As Integer
Dim strSQL As String
Dim rsTime As ADODB.Recordset
Dim strDate As String '(1115)
Dim DIDTypeMSD As String        'superchai Add 20231104
Dim tmpStr As String        'superchai Add 20231104
Dim tmpRS As New Recordset  'superchai Add 20231104
        
On Error GoTo errHandler

        strSQL = "select getdate()"
        Set rsTime = Conn.Execute(strSQL)
        strDate = Format(rsTime(0), "YYMMDDHHNNSS") '(1115)

        If OptZebra.Value = True Then
            isZebra = True
            LabelFile = Settings.TransferCompPrintLabel
        Else
            isZebra = False
            Exit Function
        End If
        StrPN = UCase(Trim(TxtCompPN))
        strDay = UCase(Trim(TxtDateCode))
        strLot = UCase(Trim(TxtLotCode)) '1097
        strQty = UCase(Trim(TxtQty))
        strVendor = UCase(Trim(TxtVendorCode))
        strUserID = UCase(Trim(TxtUserID))
        strMark = UCase(Trim(TxtMark))  '1117
        strStandard = UCase(Trim(TxtStandard)) '1249
        
        If Dir(LabelFile) = vbNullString Then
            MsgBox ("Can not find label file !"), vbCritical
            PrintLabel = "PRN_FileNoExist"
            Exit Function
        End If
        
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
        
'        hFile = FreeFile
'        Open LabelFile For Input As #hFile
'        Do
'           Select Case EOF(hFile)
'              Case True
'                Close #hFile
'                PrintLabel = "PRN_Succeed"
'                Exit Do
'              Case False
'                Line Input #hFile, hString
'                hString = Trim(hString)
'                tmpPrintStr = tmpPrintStr & Trim(hString)
'
'          End Select
'        Loop
        If FileReadAll(tmpPrintStr, LabelFile) <= 0 Then    '(1119)
            MsgBox "Open file:" & LabelFile & " fail!!", vbCritical
            Exit Function
        End If
        
        tempPN = Trim(StrPN)
        If InStr(tmpPrintStr, "<PN_CODE>") > 0 Then
           tempPN = Replace(StrPN, "^", "><")
           tmpPrintStr = Replace(tmpPrintStr, "<PN_CODE>", tempPN)
        End If
        If InStr(tmpPrintStr, "<PN_TEXT>") > 0 Then
           tempPN = Replace(StrPN, "^", "_5E")
           tmpPrintStr = Replace(tmpPrintStr, "<PN_TEXT>", tempPN)
        End If
                ''---------------
        tmpPrintStr = Replace(tmpPrintStr, "<PN>", StrPN)
        tmpPrintStr = Replace(tmpPrintStr, "<DATE>", strDay)
        tmpPrintStr = Replace(tmpPrintStr, "<Lot>", strLot)  '1097
        tmpPrintStr = Replace(tmpPrintStr, "<Vendor>", strVendor)
        tmpPrintStr = Replace(tmpPrintStr, "<QTY>", strQty)
        tmpPrintStr = Replace(tmpPrintStr, "<Standard>", strStandard) '1249
        tmpPrintStr = Replace(tmpPrintStr, "<OPID>", strUserID)
        tmpPrintStr = Replace(tmpPrintStr, "<DateTime>", strDate) '(1115)
        tmpPrintStr = Replace(tmpPrintStr, "<Mark>", strMark)   '1117
        
        'superchai add 20231104 (B)
        tmpStr = "Exec QSMS_GetDIDMSD @DID='" & tempPN & "'"
        Set tmpRS = Conn.Execute(tmpStr)
        If tmpRS.EOF = False Then
            If tmpRS("result") = "1" Then
               DIDTypeMSD = "MSD"
            End If
        End If
        
        If DIDTypeMSD = "MSD" Then
           tmpPrintStr = Replace(tmpPrintStr, "<DIDType2>", DIDTypeMSD)
        Else
           tmpPrintStr = Replace(tmpPrintStr, "<DIDType2>", "")
        End If
        DIDTypeMSD = ""
        'superchai add 20231104 (E)
        
        Select Case Trim(tmpPrintStr)
           Case vbNullString
           Case Else
                If OptComp.Value = True Then
                    For i = 1 To Len(tmpPrintStr) Step 100
                        MSComm.Output = Mid(tmpPrintStr, i, 100)
                        DoEvents
                    Next i
                    MSComm.PortOpen = False
                ElseIf OptPrint.Value = True Then
                    For i = 1 To Len(tmpPrintStr) Step 50
                        Print #lptPort, Mid(tmpPrintStr, i, 50)
                        DoEvents
                    Next i
                    Close #lptPort
                Else
                    Printer.Print tmpPrintStr
                    Printer.EndDoc
                    Printer.KillDoc
                End If
        End Select
        ''___________________
'        Close #hFile

        PrintTime = Now
        Exit Function
        
errHandler:
        MsgBox Err.Description
        If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
        End If
End Function


Private Function PrintLabelCompPort() As String
Dim hFile As Long
Dim hString As String
Dim StrPN As String, strVendor As String, strQty As String, strUserID As String, strStandard As String
Dim strDay As String
Dim LabelFile As String
Dim isZebra As Boolean
        
On Error GoTo errHandler

        If OptZebra.Value = True Then
            isZebra = True
            LabelFile = Settings.CompPrintLabel
        Else
            isZebra = False
            Exit Function
        End If
        StrPN = UCase(Trim(TxtCompPN))
        strDay = UCase(Trim(TxtDateCode))
        strQty = UCase(Trim(TxtQty))
        strVendor = UCase(Trim(TxtVendorCode))
        strUserID = UCase(Trim(TxtUserID))
        strStandard = UCase(Trim(TxtStandard)) '1249
        
        If Dir(LabelFile) = vbNullString Then
            MsgBox ("Can not find label file !"), vbCritical
            PrintLabelCompPort = "PRN_FileNoExist"
            Exit Function
        End If
        
        MSComm.CommPort = TxtCompPort
        MSComm.Settings = TxtComm
        MSComm.OutBufferCount = 0
        
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
                hString = Replace(hString, "<PN>", StrPN)
                hString = Replace(hString, "<DATE>", strDay)
                hString = Replace(hString, "<Vendor>", strVendor)
                hString = Replace(hString, "<QTY>", strQty)
                hString = Replace(hString, "<standard>", strStandard) '1249
                hString = Replace(hString, "<OPID>", strUserID)
                
               Select Case Trim(hString)
                  Case vbNullString
                  Case Else
                    MSComm.Output = hString
                    Debug.Print hString
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

Private Function PrintLabelPrintPort() As String
Dim hFile As Long
Dim hString As String
Dim StrPN As String, strVendor As String, strQty As String, strUserID As String, strStandard As String
Dim FileNum As Integer, lptPort As Integer
Dim strDay As String
Dim LabelFile As String
Dim strPort As String, PrintLabel As String
Dim isZebra As Boolean
        
On Error GoTo errhandle
    strDay = UCase(Trim(TxtDateCode))
    StrPN = UCase(Trim(TxtCompPN))
    strQty = UCase(Trim(TxtQty))
    strVendor = UCase(Trim(TxtVendorCode))
    strUserID = UCase(Trim(TxtUserID))
    strStandard = UCase(Trim(TxtStandard))
        
    If OptZebra.Value = True Then
        isZebra = True
        LabelFile = Settings.CompPrintLabel
    Else
        isZebra = False
        Exit Function
    End If

    If Dir(LabelFile) = vbNullString Then
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
            hString = Replace(hString, "<PN>", StrPN)
            hString = Replace(hString, "<DATE>", strDay)
            hString = Replace(hString, "<Vendor>", strVendor)
            hString = Replace(hString, "<QTY>", strQty)
            hString = Replace(hString, "<Standard>", strStandard) '1249
            hString = Replace(hString, "<OPID>", strUserID)
            
            Print #lptPort, hString & Chr(13)
    Wend

    Close #FileNum
    Close #lptPort
    Exit Function
    
errhandle:
     MsgBox Err.Description
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


Private Sub Form_Load()
    Dim StartTime
    
    Call reFreshData
    
    If StrBU = "NB5" Then     ''1262
        Label11.Visible = True
        TxtDID.Visible = True
        Label11.Top = 360
        TxtDID.Top = 360
    End If
    
'    CompPrintModbus = "N"
    ''1290
    Timer1.Enabled = False
    If CompPrintModbus <> "Y" Then
        txtModbusIP.Visible = False
        txtModbusStatus.Visible = False
        cmdConnect.Visible = False
        Label13.Visible = False
        Label14.Visible = False
        Timer1.Enabled = False
    End If
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    PrintTime = DateAdd("s", -CompPrintTimeSpan, Now)
    CloseDateTime = DateAdd("n", 30, Now)
    
    
'    TxtCompPort.Text = GetSetting("SMT", "QSMS", "CommPort", "1")
'    TxtComm.Text = GetSetting("SMT", "QSMS", "Comm", "9600,N,8,1")
    
    '20101115 Maggie Save Printer setting in local Registry (1019)
    Call GetPrinterSetting(frmTransferCompPrint)
'    MSComm.CommPort = GetSetting("SMT", "QSMS", "CommPort") ''(1163)
'        MSComm.Settings = "9600,n,8,1"
'        MSComm.InputMode = comInputModeText
'        MSComm.DTREnable = True
'        MSComm.Handshaking = comRTS
'        MSComm.InputLen = 0
'        MSComm.PortOpen = True
'        MsgBox "(Comport=" & MSComm.CommPort & ";" & "Setting=" & MSComm.Settings & ")"
    TxtUserID = g_userName
    
    ''1290
    If (Winsock1.State <> sckClosed) Then
        Winsock1.Close
    End If
    Winsock1.RemoteHost = txtModbusIP.text
    Winsock1.Connect
    
    StartTime = Timer
    Do While ((Timer < StartTime + 2) And (Winsock1.State <> 7))
        DoEvents
    Loop
    
    If CompPrintModbus = "Y" Then
        Timer1.Enabled = True
        If (Winsock1.State = 7) Then
           txtModbusStatus.text = "Connected"
           txtModbusStatus.BackColor = &HFF00&
        Else
           txtModbusStatus.text = "Can't connect to " + txtModbusIP.text
           txtModbusStatus.BackColor = &HFF
        End If
    End If
End Sub

''1290
Private Sub Form_Unload(Cancel As Integer)
    If (Winsock1.State <> sckClosed) Then
        Winsock1.Close
    End If
    Do While (Winsock1.State <> sckClosed)
        DoEvents
    Loop
'    txtModbusStatus.text = "Disconnected"
'    txtModbusStatus.BackColor = &HFF
    Timer1.Enabled = False
End Sub


Private Sub reFreshData()
    Dim sSql As String
    Dim Rst As ADODB.Recordset
    
    sSql = "select top 800 * from CompPrintLog order by TransDateTime desc"
    Set Rst = Conn.Execute(sSql)
    Set DataGrid1.DataSource = Rst
    
     lbCustomer = Customer
End Sub

'Private Sub MSComm_OnComm()     ''(1163)
'Dim Instring As String, NewComp() As String
'    Call Delay_Time(0.5)
'    Do While MSComm.InBufferCount <> 0
'        Instring = Instring & MSComm.Input
'        DoEvents
'    Loop
'    NewComp = Split(Trim(Instring), """")
'    If UBound(NewComp) > 0 Then
'        If TxtCompPN.Text = "" Then
'            MsgBox "请先刷入料号！！"
'        Else
'            TxtQty.Text = Trim(NewComp(1))
'            TxtQty.Locked = True
'        End If
'    End If
'End Sub

Private Sub txtCompPN_Click()
'    Sendkeys "{home}+{end}"
End Sub


Private Sub txtDID_KeyPress(KeyAscii As Integer)    ''1262
Dim sql As String
Dim Rst As ADODB.Recordset
   If KeyAscii = 13 And Trim(TxtDID) <> "" Then
        sql = "Exec QSMS_PrintDID @DID='" & Trim(TxtDID.text) & "'"
        Set Rst = Conn.Execute(sql)
        If Not Rst.EOF Then
           TxtCompPN = Trim(Rst!COMPPN)
           TxtVendorCode = Trim(Rst!VendorCode)
           TxtDateCode = Trim(Rst!DateCode)
           TxtLotCode = Trim(Rst!LotCode)
        End If
           TxtQty.SetFocus
    End If
End Sub

Private Sub txtCompPN_KeyPress(KeyAscii As Integer)
Dim NewComp() As String, index As Integer
Dim COMPPN As New Recordset, qSql As String
Dim sSql As String
Dim Rs As ADODB.Recordset
'1097
If strDelaytime <> 0 Then
    If GetTickCount - strDelaytime > 100 Then
        MsgBox "Please use scaner!"
        TxtCompPN.text = ""
        strDelaytime = 0
        Call txtCompPN_Click
        Exit Sub
    End If
End If
strDelaytime = GetTickCount
If KeyAscii = 13 Or KeyAscii = 9 Then
     strDelaytime = 0
End If

If DateDiff("n", CloseDateTime, Now) > 0 Then
    MsgBox "Message: Please Open the TransferCompPrint Program again when you use over 30 min."
    Exit Sub
End If
    

If KeyAscii = 13 And Trim(TxtCompPN) <> "" Then
    strSQL = "select CompPN,VendorCode,DateCode,LotCode,Qty from QSMS_DID_ToWH with(nolock) where DID = '" & Trim(TxtCompPN.text) & "'"
    Set Rs = Conn.Execute(strSQL)
    If Rs.EOF = False Then
        TxtCompPN.text = Trim(Rs!COMPPN)
        TxtVendorCode.text = Trim(Rs!VendorCode)
        TxtDateCode.text = Trim(Rs!DateCode)
        TxtLotCode.text = Trim(Rs!LotCode)
        TxtQty.text = Trim(Rs!Qty)
        
        sSql = "select top 1 NewQPN, Vendor from QSMS_CompPNTransfer with(nolock) where customer = '" & Trim(Customer) & "' and OldQPN = '" & Trim(TxtCompPN.text) & "' and Vendor = '" & Trim(TxtVendorCode.text) & "'"
        Set Rs = Conn.Execute(sSql)
        If Not Rs.EOF Then
'            If Rs.RecordCount = 1 Then
                TxtCompPN.text = Rs("NewQPN")
'            Else
'                Set frmCompSelect.NewCompPNList = Rs
'                frmCompSelect.Show vbModal
'                txtCompPN.text = frmCompSelect.NewCompPN
'
''                LblMessage.Caption = "This function is not yet complete."
''                Exit Function
'            End If
        Else
            sSql = "select top 1 NewQPN, Vendor from QSMS_CompPNTransfer with(nolock) where customer = '" & Trim(Customer) & "' and OldQPN = '" & Trim(TxtCompPN.text) & "' and Vendor = 'All'"
            Set Rs = Conn.Execute(sSql)
            If Not Rs.EOF Then
                TxtCompPN.text = Rs("NewQPN")
            Else
                LblMessage.Caption = "OldPN and Vendor need to upload to UniUpload."
                MsgBox "OldPN and Vendor need to upload to UniUpload.", vbCritical
                
                TxtCompPN.text = ""
                TxtVendorCode.text = ""
                TxtDateCode.text = ""
                TxtLotCode.text = ""
                TxtQty.text = ""
                
                Exit Sub
            End If
        End If
        
        TxtQty.SetFocus
        Call TxtQty_Click
    ElseIf InStr(1, Trim(TxtCompPN.text), ";") > 0 Then
        NewComp = Split(Trim(TxtCompPN.text), ";")
        For index = 0 To UBound(NewComp)
            If index = 0 Then
                TxtCompPN.text = Trim(NewComp(index))
            ElseIf index = 1 Then
                TxtDateCode.text = Trim(NewComp(index))
            ElseIf index = 2 Then
                TxtVendorCode.text = Trim(NewComp(index))
            ElseIf index = 3 Then
                TxtLotCode.text = Trim(NewComp(index))
            ElseIf index = 4 Then                           '自动从2Dbarcode中获得QTY---（1114）
                TxtQty.text = Trim(NewComp(index))
            End If
        Next index
        
        sSql = "select top 1 NewQPN, Vendor from QSMS_CompPNTransfer with(nolock) where customer = '" & Trim(Customer) & "' and OldQPN = '" & Trim(TxtCompPN.text) & "' and Vendor = '" & Trim(TxtVendorCode.text) & "'"
        Set Rs = Conn.Execute(sSql)
        If Not Rs.EOF Then
'            If Rs.RecordCount = 1 Then
                TxtCompPN.text = Rs("NewQPN")
'            Else
'                Set frmCompSelect.NewCompPNList = Rs
'                frmCompSelect.Show vbModal
'                txtCompPN.text = frmCompSelect.NewCompPN
'
''                LblMessage.Caption = "This function is not yet complete."
''                Exit Function
'            End If
        Else
            sSql = "select top 1 NewQPN, Vendor from QSMS_CompPNTransfer with(nolock) where customer = '" & Trim(Customer) & "' and OldQPN = '" & Trim(TxtCompPN.text) & "' and Vendor = 'All'"
            Set Rs = Conn.Execute(sSql)
            If Not Rs.EOF Then
                TxtCompPN.text = Rs("NewQPN")
            Else
                LblMessage.Caption = "OldPN and Vendor need to upload to UniUpload."
                MsgBox "OldPN and Vendor need to upload to UniUpload.", vbCritical
                
                TxtCompPN.text = ""
                TxtVendorCode.text = ""
                TxtDateCode.text = ""
                TxtLotCode.text = ""
                TxtQty.text = ""
                
                Exit Sub
            End If
        End If
    
        If Len(Trim(TxtCompPN)) < 11 Then
            LblMessage.Caption = "The CompPN's length must be > 11 !!"
            Exit Sub
        End If
        
        TxtQty.SetFocus
        Call TxtQty_Click
    '(1261)（1265）'(1273)
'    ElseIf StrBU <> "NB6" And InStr(1, Trim(TxtCompPN.text), "-") > 0 And Len(TxtCompPN.text) > 15 Then
'        strSQL = "select CompPN,VendorCode,DateCode,LotCode,Qty from QSMS_DID_ToWH with(nolock) where DID = '" & Trim(TxtCompPN.text) & "'"
'        Set Rs = Conn.Execute(strSQL)
'        If Rs.EOF = False Then
'            TxtCompPN.text = Trim(Rs!COMPPN)
'            TxtVendorCode.text = Trim(Rs!VendorCode)
'            TxtDateCode.text = Trim(Rs!DateCode)
'            TxtLotCode.text = Trim(Rs!LotCode)
'            TxtQty.text = Trim(Rs!Qty)
'        End If
'        TxtQty.SetFocus
'        Call TxtQty_Click
    '(1261)（1265）
    Else
    '增加包装规格条码以及SAP规格条码的输入 -----(1249)
        If StrBU = "NB6" Then
            TxtStandard.text = TxtCompPN.text
            qSql = "select CompPN,VendorCode from CompPNPrint_SAPCompPNinfo with(nolock) where SAP_Size='" & Trim(TxtCompPN.text) & "' or Package_Size='" & Trim(TxtCompPN.text) & "'"
            Set COMPPN = Conn.Execute(qSql)
            If COMPPN.RecordCount > 0 Then
                TxtStandard.text = Trim(TxtCompPN.text)
                TxtCompPN.text = COMPPN.Fields("CompPN").Value
                TxtVendorCode.text = COMPPN.Fields("VendorCode").Value
                ''TxtVendorCode.SetFocus
                TxtDateCode.SetFocus
                Call TxtDateCode_Click
            Else
                MsgBox "Please Check  Uniupload --> 上传SAP_CompPN_Info ", vbCritical
                TxtCompPN.text = ""
                TxtCompPN.SetFocus
                Call txtCompPN_Click
            End If
        Else
'            sSql = "select NewQPN, Vendor from QSMS_CompPNTransfer with(nolock) where OldQPN = '" & Trim(txtCompPN.text) & "'"
'            Set Rs = Conn.Execute(sSql)
'            If Not Rs.EOF Then
'                If Rs.RecordCount = 1 Then
'                    txtCompPN.text = Rs("NewQPN")
'                Else
'                    Set frmCompSelect.NewCompPNList = Rs
'                    frmCompSelect.Show vbModal
'                    txtCompPN.text = frmCompSelect.NewCompPN
'                    txtVendorCode.text = frmCompSelect.NewCompPN
'
'    '                LblMessage.Caption = "This function is not yet complete."
'    '                Exit Function
'                End If
'            Else
'                LblMessage.Caption = "OldPN and Vendor need to upload to UniUpload."
'                MsgBox "OldPN and Vendor need to upload to UniUpload.", vbCritical
'
'                Exit Function
'            End If
'
'            If Len(Trim(txtCompPN)) < 11 Then
'                LblMessage.Caption = "The CompPN's length must be > 11 !!"
'                Exit Sub
'            End If
        
            TxtVendorCode.SetFocus
            Call TxtVendorCode_Click
        End If
    End If
End If
End Sub


Private Sub TxtDateCode_Click()
'   Sendkeys "{home}+{end}"
End Sub

Private Sub TxtDateCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(TxtDateCode) <> "" Then
      TxtLotCode.SetFocus
      Call TxtLotCode_Click
    End If
End Sub

Private Sub TxtLotCode_Click()
'    Sendkeys "{home}+{end}"
End Sub

Private Sub TxtLotCode_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 And Trim(TxtLotCode) <> "" Then
        TxtQty.SetFocus
        Call TxtQty_Click
     End If
End Sub

Private Sub TxtMark_Click()
'    Sendkeys "{home}+{end}"
End Sub

Private Sub TxtQty_Click()
'    Sendkeys "{home}+{end}"
End Sub

Private Sub TxtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(TxtQty) <> "" Then
        TxtMark.SetFocus
        Call TxtMark_Click
    End If
End Sub

Private Sub TxtMark_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(TxtMark) <> "" Then
        CmdPrint.SetFocus
    End If
End Sub

Private Sub TxtVendorCode_Click()
'   Sendkeys "{home}+{end}"
End Sub

Private Sub TxtVendorCode_KeyPress(KeyAscii As Integer)
    Dim sSql As String
    Dim Rs As ADODB.Recordset

    If KeyAscii = 13 And Trim(TxtVendorCode) <> "" Then
 
        sSql = "select top 1 NewQPN, Vendor from QSMS_CompPNTransfer with(nolock) where customer = '" & Trim(Customer) & "' and ((OldQPN = '" & Trim(TxtCompPN.text) & "' and Vendor = '" & Trim(TxtVendorCode.text) & "') or (NewQPN = '" & Trim(TxtCompPN.text) & "'))"
        Set Rs = Conn.Execute(sSql)
        If Not Rs.EOF Then
            TxtCompPN.text = Rs("NewQPN")
        Else
            sSql = "select top 1 NewQPN, Vendor from QSMS_CompPNTransfer with(nolock) where customer = '" & Trim(Customer) & "' and ((OldQPN = '" & Trim(TxtCompPN.text) & "' and Vendor = 'All') or (NewQPN = '" & Trim(TxtCompPN.text) & "'))"
            Set Rs = Conn.Execute(sSql)
            If Not Rs.EOF Then
                TxtCompPN.text = Rs("NewQPN")
            Else
                LblMessage.Caption = "OldPN and Vendor need to upload to UniUpload."
                MsgBox "OldPN and Vendor need to upload to UniUpload.", vbCritical
                
                TxtCompPN.text = ""
                TxtVendorCode.text = ""
                TxtDateCode.text = ""
                TxtLotCode.text = ""
                TxtQty.text = ""
                
                Exit Sub
            End If
        End If
     
        TxtDateCode.SetFocus
        Call TxtDateCode_Click
    End If
End Sub
