VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmCompPrint 
   BackColor       =   &H8000000B&
   Caption         =   "CompPrint 2019/04/12"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkPrint 
      Caption         =   "转料号"
      Height          =   495
      Left            =   7080
      TabIndex        =   39
      Top             =   3720
      Width           =   855
   End
   Begin VB.Frame Frame5 
      Caption         =   "开封标签打印"
      Height          =   2745
      Left            =   8160
      TabIndex        =   35
      Top             =   1800
      Width           =   1935
      Begin VB.CommandButton btpl 
         Caption         =   "打印"
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
         Caption         =   "打印数量:"
         Height          =   195
         Left            =   150
         TabIndex        =   37
         Top             =   720
         Width           =   795
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1665
      Left            =   3480
      TabIndex        =   19
      Top             =   30
      Width           =   6600
      Begin MSCommLib.MSComm MSComm 
         Left            =   5280
         Top             =   720
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
         Left            =   3930
         Picture         =   "frmCompPrint.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   600
         Width           =   1030
      End
      Begin VB.TextBox TxtComm 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   25
         Text            =   "9600,N,8,1"
         Top             =   990
         Width           =   2100
      End
      Begin VB.TextBox TxtCompPort 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   24
         Text            =   "1"
         Top             =   480
         Width           =   2100
      End
      Begin VB.Label Label2 
         Caption         =   "Settings："
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
         Top             =   990
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
         Top             =   480
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
         Picture         =   "frmCompPrint.frx":066A
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
Attribute VB_Name = "frmCompPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim transdatetime As String

Private Sub btpl_Click()
  Dim LabelFile As String
  
  If Trim(txtlp) = "" Or IsNumeric(txtlp) = False Then
     LblMessage = "打印数量输入错误!!"
     Exit Sub
  End If
  Call Printlp(txtlp)
  txtlp = ""
  
End Sub

Private Sub CmdCommSave_Click()
SaveSetting "SMT", "QSMS", "CommPort", TxtCompPort
SaveSetting "SMT", "QSMS", "Comm", TxtComm
End Sub

Private Sub CmdPrint_Click()

    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim sCompPN As String
    'Dim intQty As Integer
    Dim intQty As Long   '(1180)
    
    sSql = "select getdate()"
    Set rs = Conn.Execute(sSql)
    transdatetime = Format(rs.Fields(0), "YYYYMMDDHHMMSS")
    
    If ValidateData = False Then
       Exit Sub
    End If
    
    LblMessage = "Message"
    
    
    If LabelPrintCheck = "Y" And ChkPrint.Value = 1 Then  ''1274
      sSql = "exec QSMS_CompPrintCheck @CompPN='" & Trim(TxtCompPN.Text) & "',@VendorCode='" & Trim(TxtVendorCode.Text) & "',@DateCode='" & Trim(TxtDateCode.Text) & "',@LotCode='" & Trim(TxtLotCode.Text) & "',@Qty='" & Trim(TxtQty.Text) & "',@UserID='" & Trim(TxtUserID.Text) & "',@Mark='" & Trim(TxtMark.Text) & "',@Standard='" & Trim(TxtStandard.Text) & "'"
      Set rs = Conn.Execute(sSql)
      If Not rs.EOF Then
        If rs("Result") <> "" Then
            TxtCompPN.Text = rs("Result")
        End If
      End If
    Else
        sCompPN = Trim(TxtCompPN)
        intQty = Trim(TxtQty)
        
        sSql = "insert into CompPrintLog(CompPN,Qty,VendorCode,DateCode,LotCode,OPID,TransDateTime,Mark) " & _
            "values('" & UCase(Trim(sCompPN)) & "','" & UCase(Trim(intQty)) & "','" & UCase(Trim(TxtVendorCode)) & "','" & UCase(Trim(TxtDateCode)) & "','" & UCase(Trim(TxtLotCode)) & "','" & UCase(Trim(TxtUserID)) & "','" & UCase(Trim(transdatetime)) & "','" & UCase(Trim(TxtMark)) & "')"  '1117
        Set rs = Conn.Execute(sSql)
    End If
    
    Call PrintLabel

    Call reFreshData
    
    If StrBU = "NB5" Then      ''1262
       TxtDID = ""
    End If
    
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
Private Function ValidateData() As Boolean
      
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
        LblMessage = "The CompPN's length must be >11 !!"
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
Dim I As Integer
Dim J As Integer
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
                   For J = 1 To Qty
                      For I = 1 To Len(tmpPrintStr) Step 100
                          MSComm.Output = Mid(tmpPrintStr, I, 100)
                          DoEvents
                      Next I
                   Next J
                    MSComm.PortOpen = False
                ElseIf OptPrint.Value = True Then
                    For J = 1 To Qty
                        For I = 1 To Len(tmpPrintStr) Step 50
                           Print #lptPort, Mid(tmpPrintStr, I, 50)
                           DoEvents
                        Next I
                    Next J
                    Close #lptPort
                Else
                    For J = 1 To Qty
                       Printer.Print tmpPrintStr
                       Printer.EndDoc
                       Printer.KillDoc
                       For M = 1 To 5000
                       Next M
                    Next J
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
Dim I As Integer
Dim tmpPrintStr As String
Dim hString As String
Dim strPN As String, tempPN As String, strVendor As String, strQty As String, strUserID As String, strMark As String
Dim strStandard As String   '添加新的栏位显示刷入非2D Barcode 信息 1249
Dim strDay As String
Dim strLot As String
Dim LabelFile As String
Dim isZebra As Boolean
Dim lptPort As Integer
Dim strSql As String
Dim rsTime As ADODB.Recordset
Dim strDate As String '(1115)
        
On Error GoTo errHandler

        strSql = "select getdate()"
        Set rsTime = Conn.Execute(strSql)
        strDate = Format(rsTime(0), "YYMMDDHHNNSS") '(1115)

        If OptZebra.Value = True Then
            isZebra = True
            LabelFile = Settings.CompPrintLabel
        Else
            isZebra = False
            Exit Function
        End If
        strPN = UCase(Trim(TxtCompPN))
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
        
        tempPN = Trim(strPN)
        If InStr(tmpPrintStr, "<PN_CODE>") > 0 Then
           tempPN = Replace(strPN, "^", "><")
           tmpPrintStr = Replace(tmpPrintStr, "<PN_CODE>", tempPN)
        End If
        If InStr(tmpPrintStr, "<PN_TEXT>") > 0 Then
           tempPN = Replace(strPN, "^", "_5E")
           tmpPrintStr = Replace(tmpPrintStr, "<PN_TEXT>", tempPN)
        End If
                ''---------------
        tmpPrintStr = Replace(tmpPrintStr, "<PN>", strPN)
        tmpPrintStr = Replace(tmpPrintStr, "<DATE>", strDay)
        tmpPrintStr = Replace(tmpPrintStr, "<Lot>", strLot)  '1097
        tmpPrintStr = Replace(tmpPrintStr, "<Vendor>", strVendor)
        tmpPrintStr = Replace(tmpPrintStr, "<QTY>", strQty)
        tmpPrintStr = Replace(tmpPrintStr, "<Standard>", strStandard) '1249
        tmpPrintStr = Replace(tmpPrintStr, "<OPID>", strUserID)
        tmpPrintStr = Replace(tmpPrintStr, "<DateTime>", strDate) '(1115)
        tmpPrintStr = Replace(tmpPrintStr, "<Mark>", strMark)   '1117
       
        
        Select Case Trim(tmpPrintStr)
           Case vbNullString
           Case Else
                If OptComp.Value = True Then
                    For I = 1 To Len(tmpPrintStr) Step 100
                        MSComm.Output = Mid(tmpPrintStr, I, 100)
                        DoEvents
                    Next I
                    MSComm.PortOpen = False
                ElseIf OptPrint.Value = True Then
                    For I = 1 To Len(tmpPrintStr) Step 50
                        Print #lptPort, Mid(tmpPrintStr, I, 50)
                        DoEvents
                    Next I
                    Close #lptPort
                Else
                    Printer.Print tmpPrintStr
                    Printer.EndDoc
                    Printer.KillDoc
                End If
        End Select
        ''___________________
'        Close #hFile
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
Dim strPN As String, strVendor As String, strQty As String, strUserID As String, strStandard As String
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
        strPN = UCase(Trim(TxtCompPN))
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
                hString = Replace(hString, "<PN>", strPN)
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
Dim strPN As String, strVendor As String, strQty As String, strUserID As String, strStandard As String
Dim FileNum As Integer, lptPort As Integer
Dim strDay As String
Dim LabelFile As String
Dim strPort As String, PrintLabel As String
Dim isZebra As Boolean
        
On Error GoTo errhandle
    strDay = UCase(Trim(TxtDateCode))
    strPN = UCase(Trim(TxtCompPN))
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
            hString = Replace(hString, "<PN>", strPN)
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
    Call reFreshData
    
    If StrBU = "NB5" Then     ''1262
        Label11.Visible = True
        TxtDID.Visible = True
        Label11.Top = 360
        TxtDID.Top = 360
    End If
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

'    TxtCompPort.Text = GetSetting("SMT", "QSMS", "CommPort", "1")
'    TxtComm.Text = GetSetting("SMT", "QSMS", "Comm", "9600,N,8,1")
    
    '20101115 Maggie Save Printer setting in local Registry (1019)
    Call GetPrinterSetting(frmCompPrint)
'    MSComm.CommPort = GetSetting("SMT", "QSMS", "CommPort") ''(1163)
'        MSComm.Settings = "9600,n,8,1"
'        MSComm.InputMode = comInputModeText
'        MSComm.DTREnable = True
'        MSComm.Handshaking = comRTS
'        MSComm.InputLen = 0
'        MSComm.PortOpen = True
'        MsgBox "(Comport=" & MSComm.CommPort & ";" & "Setting=" & MSComm.Settings & ")"
    TxtUserID = g_userName
End Sub

Private Sub reFreshData()
    Dim sSql As String
    Dim Rst As ADODB.Recordset
    
    sSql = "select top 800 * from CompPrintLog order by TransDateTime desc"
    Set Rst = Conn.Execute(sSql)
    Set DataGrid1.DataSource = Rst
    
End Sub

Private Sub Text1_Change()

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
    SendKeys "{home}+{end}"
End Sub
Private Sub TxtDID_KeyPress(KeyAscii As Integer)    ''1262
Dim sql As String
Dim Rst As ADODB.Recordset
   If KeyAscii = 13 And Trim(TxtDID) <> "" Then
        sql = "Exec QSMS_PrintDID @DID='" & Trim(TxtDID.Text) & "'"
        Set Rst = Conn.Execute(sql)
        If Not Rst.EOF Then
           TxtCompPN = Trim(Rst!compPN)
           TxtVendorCode = Trim(Rst!VendorCode)
           TxtDateCode = Trim(Rst!DateCode)
           TxtLotCode = Trim(Rst!LotCode)
        End If
           TxtQty.SetFocus
    End If
End Sub

Private Sub txtCompPN_KeyPress(KeyAscii As Integer)
Dim NewComp() As String, index As Integer
Dim compPN As New Recordset, qSql As String
    '1097

If KeyAscii = 13 And Trim(TxtCompPN) <> "" Then
    If InStr(1, Trim(TxtCompPN.Text), ";") > 0 Then
        NewComp = Split(Trim(TxtCompPN.Text), ";")
        For index = 0 To UBound(NewComp)
            If index = 0 Then
                TxtCompPN.Text = Trim(NewComp(index))
            ElseIf index = 1 Then
                TxtDateCode.Text = Trim(NewComp(index))
            ElseIf index = 2 Then
                TxtVendorCode.Text = Trim(NewComp(index))
            ElseIf index = 3 Then
                TxtLotCode.Text = Trim(NewComp(index))
            ElseIf index = 4 Then                           '自动从2Dbarcode中获得QTY---（1114）
                TxtQty.Text = Trim(NewComp(index))
            End If
        Next index
        TxtQty.SetFocus
        Call TxtQty_Click
    '(1261)（1265）'(1273)
    ElseIf StrBU <> "NB6" And InStr(1, Trim(TxtCompPN.Text), "-") > 0 And Len(TxtCompPN.Text) > 15 Then
        strSql = "select CompPN,VendorCode,DateCode ,LotCode ,Qty from QSMS_DID_ToWH where DID = '" & Trim(TxtCompPN.Text) & "' "
        Set rs = Conn.Execute(strSql)
        If rs.EOF = False Then
            TxtCompPN.Text = Trim(rs!compPN)
            TxtVendorCode.Text = Trim(rs!VendorCode)
            TxtDateCode.Text = Trim(rs!DateCode)
            TxtLotCode.Text = Trim(rs!LotCode)
            TxtQty.Text = Trim(rs!Qty)
        End If
        TxtQty.SetFocus
        Call TxtQty_Click
    '(1261)（1265）
    Else
    '增加包装规格条码以及SAP规格条码的输入 -----(1249)
        If StrBU = "NB6" Then
            TxtStandard.Text = TxtCompPN.Text
            qSql = "select CompPN,VendorCode from CompPNPrint_SAPCompPNinfo where SAP_Size='" & Trim(TxtCompPN.Text) & "' or Package_Size='" & Trim(TxtCompPN.Text) & "'"
            Set compPN = Conn.Execute(qSql)
            If compPN.RecordCount > 0 Then
                TxtStandard.Text = Trim(TxtCompPN.Text)
                TxtCompPN.Text = compPN.Fields("CompPN").Value
                TxtVendorCode.Text = compPN.Fields("VendorCode").Value
                ''TxtVendorCode.SetFocus
                TxtDateCode.SetFocus
                Call TxtDateCode_Click
            Else
                MsgBox "Please Check  Uniupload --> 上传SAP_CompPN_Info ", vbCritical
                TxtCompPN.Text = ""
                TxtCompPN.SetFocus
                Call txtCompPN_Click
            End If
        Else
            TxtVendorCode.SetFocus
            Call TxtVendorCode_Click
        End If
    End If
End If
End Sub


Private Sub TxtDateCode_Click()
   SendKeys "{home}+{end}"
End Sub

Private Sub TxtDateCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(TxtDateCode) <> "" Then
      TxtLotCode.SetFocus
      Call TxtLotCode_Click
    End If
End Sub

Private Sub TxtLotCode_Click()
    SendKeys "{home}+{end}"
End Sub

Private Sub TxtLotCode_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 And Trim(TxtLotCode) <> "" Then
        TxtQty.SetFocus
        Call TxtQty_Click
     End If
End Sub

Private Sub TxtMark_Click()
    SendKeys "{home}+{end}"
End Sub

Private Sub TxtQty_Click()
    SendKeys "{home}+{end}"
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

Private Sub Txtsap_Change()

End Sub

Private Sub TxtVendorCode_Click()
   SendKeys "{home}+{end}"
End Sub

Private Sub TxtVendorCode_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 And Trim(TxtVendorCode) <> "" Then
      TxtDateCode.SetFocus
      Call TxtDateCode_Click
 End If
End Sub
