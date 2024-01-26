VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmCompPNPrint 
   BackColor       =   &H8000000B&
   Caption         =   "CompPNPrint[2010/11/15]"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   1665
      Left            =   3510
      TabIndex        =   13
      Top             =   30
      Width           =   5640
      Begin MSCommLib.MSComm MSComm 
         Left            =   5010
         Top             =   90
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
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   1030
      End
      Begin VB.TextBox TxtComm 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   19
         Text            =   "9600,N,8,1"
         Top             =   990
         Width           =   2100
      End
      Begin VB.TextBox TxtCompPort 
         BackColor       =   &H00FFFFC0&
         Height          =   300
         Left            =   1260
         TabIndex        =   18
         Text            =   "1"
         Top             =   480
         Width           =   2100
      End
      Begin VB.Label Label2 
         Caption         =   "Settings£º"
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
         TabIndex        =   17
         Top             =   990
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "CompPort£º"
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
         TabIndex        =   16
         Top             =   480
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1665
      Left            =   1725
      TabIndex        =   12
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   1200
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2265
      Left            =   30
      TabIndex        =   4
      Top             =   3720
      Width           =   9125
      _ExtentX        =   16087
      _ExtentY        =   3995
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
      Height          =   1935
      Left            =   30
      TabIndex        =   3
      Top             =   1740
      Width           =   9125
      Begin VB.TextBox TxtUserID 
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1215
         Width           =   2670
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
         Left            =   6870
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   765
      End
      Begin VB.TextBox TxtNewCompPN 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   5265
         TabIndex        =   8
         Top             =   435
         Width           =   2670
      End
      Begin VB.TextBox TxtCompPN 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   380
         Left            =   1095
         TabIndex        =   6
         Top             =   435
         Width           =   2670
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
         Left            =   270
         TabIndex        =   9
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label Label6 
         Caption         =   "NewCompPN"
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
         Left            =   4080
         TabIndex        =   7
         Top             =   525
         Width           =   1305
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
         Left            =   240
         TabIndex        =   5
         Top             =   525
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Printer"
      Height          =   1665
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1695
      Begin VB.OptionButton optNetWork 
         Caption         =   "Network"
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
         Height          =   255
         Left            =   210
         TabIndex        =   22
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton OptPrint 
         Caption         =   "Print Port"
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
         Height          =   300
         Left            =   210
         TabIndex        =   2
         Top             =   750
         Width           =   1200
      End
      Begin VB.OptionButton OptComp 
         Caption         =   "Comp Port"
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
      Left            =   30
      TabIndex        =   21
      Top             =   6030
      Width           =   9125
   End
End
Attribute VB_Name = "frmCompPNPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TransDateTime As String

'Private Sub CmdCommSave_Click()
'SaveSetting "SMT", "QSMS", "CommPort", TxtCompPort
'SaveSetting "SMT", "QSMS", "Comm", TxtComm
'End Sub

Private Sub CmdPrint_Click()

    If ValidateData = False Then
       Exit Sub
    End If

    Call PrintLabel

    Call RefreshData
    
    TxtCompPN = ""
    TxtNewCompPN = ""
    
    TxtCompPN.SetFocus
        
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
    
    If Trim(TxtNewCompPN) = "" Then
        LblMessage = "NewCompPN is blank!!"
        Exit Function
    End If
  
    If Trim(TxtCompPort) = "" Or Trim(TxtComm) = "" Then
        LblMessage = "Printer have not set!!"
        Exit Function
    End If

    ValidateData = True
    
End Function

Private Function PrintLabel()
Dim hFile As Long
Dim i As Integer
Dim tmpPrintStr As String
Dim hString As String
Dim strNewCompPN As String
Dim strDay As String
Dim LabelFile As String
Dim isZebra As Boolean
Dim lptPort As Integer
        
On Error GoTo errHandler

        If OptZebra.Value = True Then
            isZebra = True
            LabelFile = Settings.CompPNLabelPrint
        Else
            isZebra = False
            Exit Function
        End If
        strNewCompPN = UCase(Trim(TxtNewCompPN))
        
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
        
        hFile = FreeFile
        Open LabelFile For Input As #hFile
        Do
           Select Case EOF(hFile)
              Case True
                Close #hFile
                PrintLabel = "PRN_Succeed"
                Exit Do
              Case False
                Line Input #hFile, hString
                hString = Trim(hString)
                tmpPrintStr = tmpPrintStr & Trim(hString)
                
          End Select
        Loop
        
                ''---------------
        tmpPrintStr = Replace(tmpPrintStr, "<NewCompPN>", strNewCompPN)
        
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
        Close #hFile
        Exit Function
        
errHandler:
        MsgBox Err.Description
        If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
        End If
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
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

'    TxtCompPort.Text = GetSetting("SMT", "QSMS", "CommPort", "1")
'    TxtComm.Text = GetSetting("SMT", "QSMS", "Comm", "9600,N,8,1")

    '20101115 Maggie Save Printer setting in local Registry (1019)
    Call GetPrinterSetting(frmCompPNPrint)
    
    TxtUserID = g_userName
    
    Call RefreshData
    
End Sub

Private Sub RefreshData()
    Dim sSql As String
    Dim rst As adodb.Recordset
    
    sSql = "select * from CompPNChg_ForStock"
    Set rst = Conn.Execute(sSql)
    Set DataGrid1.DataSource = rst
    
End Sub

Private Sub TxtCompPN_Click()
    SendKeys "{home}+{end}"
End Sub

Private Sub TxtCompPN_KeyPress(KeyAscii As Integer)
Dim sSql As String
Dim rst As adodb.Recordset

On Error GoTo Err_Handler

    If Trim(TxtCompPN) = "" Or KeyAscii <> 13 Then Exit Sub

    sSql = "select * from CompPNChg_ForStock where OldCompPN=" & sq(UCase(Trim(TxtCompPN)))
    Set rst = Conn.Execute(sSql)
    If rst.EOF = False Then
        TxtNewCompPN = rst("NewCompPN")
        CmdPrint.SetFocus
    Else
        MsgBox "Can not find CompPN:" & Trim(TxtCompPN) & "!!", vbExclamation, "Prompt"
        TxtCompPN.SetFocus
    End If

    Exit Sub
    
Err_Handler:
    MsgBox Err.Number & ":" & Err.Description, vbExclamation, "ERROR"
End Sub
