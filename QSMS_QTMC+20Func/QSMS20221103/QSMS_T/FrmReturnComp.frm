VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form FrmReturnComp 
   Caption         =   "Return Component[20190808]"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   13260
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Chk_A8 
      Caption         =   "A8"
      Height          =   255
      Left            =   11880
      TabIndex        =   35
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox ChkHUA 
      Caption         =   "HUA"
      Height          =   255
      Left            =   11880
      TabIndex        =   34
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox ChkBGA 
      Caption         =   "重植球"
      Height          =   375
      Left            =   11880
      TabIndex        =   33
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   30
      Left            =   3360
      TabIndex        =   32
      Top             =   1200
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   53
      _Version        =   327682
   End
   Begin VB.CommandButton cmdGetRefID 
      Caption         =   "&GetRefID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1185
   End
   Begin VB.CommandButton cmdReprint 
      Caption         =   "&Reprint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   525
      Left            =   10080
      TabIndex        =   25
      Top             =   1860
      Width           =   3075
      Begin VB.OptionButton optBadMaterial 
         Caption         =   "Bad"
         Height          =   285
         Left            =   1590
         TabIndex        =   27
         Top             =   180
         Width           =   615
      End
      Begin VB.OptionButton optGoodMaterial 
         Caption         =   "Good"
         Height          =   345
         Left            =   480
         TabIndex        =   26
         Top             =   120
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin VB.Frame FraPrinter 
      BackColor       =   &H80000013&
      Caption         =   "Printer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   12915
      Begin VB.Frame Frame5 
         Height          =   615
         Left            =   3210
         TabIndex        =   7
         Top             =   240
         Width           =   3675
         Begin VB.OptionButton optNetwork 
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
            Left            =   2520
            TabIndex        =   31
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
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
            Height          =   255
            Left            =   60
            TabIndex        =   9
            Top             =   240
            Width           =   1215
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
            Height          =   255
            Left            =   1320
            TabIndex        =   8
            Top             =   240
            Width           =   1155
         End
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
         Left            =   8400
         TabIndex        =   6
         Text            =   "1"
         Top             =   300
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
         Left            =   9960
         TabIndex        =   5
         Text            =   "9600,N,8,1"
         Top             =   300
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
         Height          =   555
         Left            =   11520
         Picture         =   "FrmReturnComp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   1335
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   60
         TabIndex        =   1
         Top             =   210
         Width           =   3105
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
            Left            =   1500
            TabIndex        =   3
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
            Left            =   60
            TabIndex        =   2
            Top             =   210
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
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
         Index           =   2
         Left            =   6960
         TabIndex        =   11
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
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
         Index           =   6
         Left            =   8880
         TabIndex        =   10
         Top             =   300
         Width           =   1095
      End
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   9480
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5520
      TabIndex        =   23
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtLotCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7920
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtDateCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1320
      Width           =   4095
   End
   Begin VB.TextBox txtVendorCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7920
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox txtCompPN 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   960
      Width           =   4095
   End
   Begin MSDataGridLib.DataGrid gridReturnComp 
      Height          =   4335
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   7646
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.TextBox txtQty 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1920
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label lblFeedBack 
      BackColor       =   &H80000013&
      Caption         =   "Qty FeedBack: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   7440
      Width           =   12975
   End
   Begin VB.Label LblMessage 
      BackColor       =   &H80000000&
      Height          =   525
      Left            =   120
      TabIndex        =   24
      Top             =   6840
      Width           =   12975
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
      Left            =   240
      TabIndex        =   17
      Top             =   1320
      Width           =   1695
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
      Left            =   240
      TabIndex        =   16
      Top             =   960
      Width           =   1695
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
      Left            =   6120
      TabIndex        =   15
      Top             =   960
      Width           =   1815
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
      Left            =   6120
      TabIndex        =   14
      Top             =   1320
      Width           =   1815
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
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
   End
End
Attribute VB_Name = "FrmReturnComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/**********************************************************************************
'**文 件 名: FrmReturnComp.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: Denver Yang
'**日    期: 2008.03.19
'**描    述: QSMS Return CompPN
'
'**EQMSID       修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**             Jeanson      2008.03.28     to check the from:frmDIDChkStock exists or not (00001)
'**             Denver       2008.04.02     show return Qty and LineMC stock Qty   --(0002)
'**             Denver       2008.04.02     it will replace if there is "'"  --(0002)
'**             Austin       2009.12.18     Set  txtVendorCode.SetFocus   --0003
'**             Denver       2009.12.22     ESBU need not clear them    --0004
'*RQ10042001    Kane         2010/05/18     如果是禁用料则不能退料'(0005)
'***********************************************************************************/


Private Sub CmdCommSave_Click()
    SaveSetting "SMT", "QSMS", "CommPort", TxtCompPort
    SaveSetting "SMT", "QSMS", "Comm", TxtComm
End Sub

Private Sub cmdGetRefID_Click()
    Dim sSql As String
    Dim Rst As ADODB.Recordset
    Dim sCurrRefID As String
    Dim sMsg As String
    Dim UserName As String
    
    If Chk_A8.Value = 1 Then
        UserName = "A8_Job_Auto"
    Else
        UserName = g_userName
    End If
    
    sSql = "exec XL_DIDGetRefID @Type='Return', @IsGood=" & sq(IIf(optGoodMaterial.Value = True, "Y", "N")) & ",@UserName=" & sq(UserName) & ",@Factory=" & sq(Trim(Factory))
    Set Rst = Conn.Execute(sSql)
    If Rst.EOF = False Then
        
        If Rst("Result") <> 0 Then
            MsgBox Rst("Description"), vbExclamation, "Prompt"
            Exit Sub
        End If
        
        sMsg = Trim(Rst("Description") & "")
        sCurrRefID = DIDGetRefIDByResult(sMsg)
        
        ''打印Label for GetRefID
        With DIDInfo
            .DID = sCurrRefID
            .compPN = sCurrRefID
            .Qty = -10000
            .IsGood = IIf(optGoodMaterial.Value = True, "Y", "N")
            .DIDType = ""
        End With
        Call DIDPrintLabel(OptZebra.Value, CInt(Trim(TxtCompPort)), Trim(TxtComm))
        
        ''Check Stock qty By RefID
        sSql = "exec XL_DIDChkStockByRefID @Type='Auto',@RefID=" & sq(sCurrRefID) & ",@UserName=" & sq(UserName)
        Set Rst = Conn.Execute(sSql)
        If Rst.EOF = False Then
            If Rst("Result") <> 0 Then
                MsgBox Rst("Description"), vbExclamation, "Prompt"
                Exit Sub
            End If
            
            'to check the from:frmDIDChkStock exists or not (00001)
            Dim frm As Form
            For Each frm In Forms
                If frm.Name = "frmDIDChkStock" Then
                    Unload frm
                    Exit For
                End If
            Next frm
            'to check the from:frmDIDChkStock exists or not (00001)
            
            frmDIDChkStock.FuncType = "AutoChk"
            Set Rst = Rst.NextRecordset
            Set frmDIDChkStock.rstCompPN = Rst
            frmDIDChkStock.lblmsg = sMsg
            frmDIDChkStock.Show 1
            
        End If
    
    End If
End Sub

Private Sub cmdOK_Click()
On Error GoTo Err_Handler

    Dim sSql As String
    Dim Rst As ADODB.Recordset
    Dim sCompPN As String
    Dim intReturnQty As String
    ''1058
    cmdOK.Enabled = False
    
    LblMessage = ""
    lblFeedBack = "Qty FeedBack:"
    If ChkValidData = False Then
        GoTo Normal_Eixt
    End If

    ''20081230   Denver  Check OK后，运用中间变量保存，以防运行中修改
    sCompPN = Trim(TxtCompPN)
    intReturnQty = Trim(TxtQty)
    
    sSql = "select * from MSD_DATA where CompPN=" & sq(TxtCompPN.Text)    '''(1272)
    Set Rst = Conn.Execute(sSql)
        
    If BU = "ESBU" And Rst.EOF = False Then
        If MsgBox("MSD材料，请仔细检查！", vbYesNo) = vbNo Then
'            GoTo Normal_Eixt
            If Chk_A8.Value = 0 Then
                Exit Sub
            End If
        End If
    End If
    
    '(0005)
    If ChkBGA.Value = 1 Then    '' 1205
        sSql = "exec XL_ReturnComp " & sq(Trim(sCompPN)) & "," & sq(Trim(txtVendorCode)) & "," & sq(Trim(txtDateCode)) & "," & sq(Trim(txtLotCode)) & "," & sq(IIf(optGoodMaterial.Value = True, "YBGA", "NBGA")) & "," & Trim(intReturnQty) & "," & sq(g_userName) & "," & sq(Trim(Factory)) & "," & sq(Trim(CheckReturnForbiddenPN))
    ElseIf ChkHUA.Value = 1 Then  ''1245
        sSql = "exec XL_ReturnComp " & sq(Trim(sCompPN)) & "," & sq(Trim(txtVendorCode)) & "," & sq(Trim(txtDateCode)) & "," & sq(Trim(txtLotCode)) & "," & sq(IIf(optGoodMaterial.Value = True, "YHUA", "NHUA")) & "," & Trim(intReturnQty) & "," & sq(g_userName) & "," & sq(Trim(Factory)) & "," & sq(Trim(CheckReturnForbiddenPN))
    ElseIf Chk_A8.Value = 1 Then
        sSql = "UPDATE QSMS_DID_ToWH SET Status='0' WHERE CompPN ='" & TxtCompPN.Text & "' and VendorCode ='" & txtVendorCode.Text & "' and DateCode ='" & txtDateCode.Text & "' and LotCOde ='" & txtLotCode.Text & "' and Qty = '" & TxtQty.Text & "' and status ='A' "
        Set Rst = Conn.Execute(sSql)
        txtVendorCode = ""
        txtLotCode = ""
        txtDateCode = ""
        TxtQty = ""
        TxtCompPN = ""
        TxtCompPN.SetFocus
        
        Call reFreshData
        Exit Sub
    Else
        sSql = "exec XL_ReturnComp " & sq(Trim(sCompPN)) & "," & sq(Trim(txtVendorCode)) & "," & sq(Trim(txtDateCode)) & "," & sq(Trim(txtLotCode)) & "," & sq(IIf(optGoodMaterial.Value = True, "Y", "N")) & "," & Trim(intReturnQty) & "," & sq(g_userName) & "," & sq(Trim(Factory)) & "," & sq(Trim(CheckReturnForbiddenPN))
    End If
    Set Rst = Conn.Execute(sSql)
    If Rst("Result") <> 0 Then
        LblMessage.Caption = Rst("Description")
    Else
        LblMessage.BackColor = &H80FF80
        Set Rst = Rst.NextRecordset
        'PN/Qty/PU/NG/UID/Date   (bad)   'DID/Qty/PU/UID/Date (Good)
        If Rst.EOF = True Then
            LblMessage.Caption = "Get DID information fail,print DID fail!!"
            GoTo Normal_Eixt
        End If
        
        ''2008/04/02 denver    show return Qty and LineMC stock Qty   --(0002)
        lblFeedBack = Trim(Rst("QtyFeedback") & "")
        lblFeedBack = Mid(lblFeedBack, InStr(lblFeedBack, "##") + 2, 600)
            
        With DIDInfo
            .DID = Trim(Rst("DID") & "")
            .compPN = Trim(Rst("CompPN") & "")
            .Qty = Rst("Qty")
            .IsGood = Trim(Rst("IsGood") & "")
            .VendorCode = Trim(Rst("VendorCode"))
            .DateCode = Trim(Rst("DateCode"))
            .LotCode = Trim(Rst("LotCode"))
            If BU = "NB5" Then
                .WareHouseID = Trim(Rst("WareHouseID"))        '(1252)
            End If
            If ChkPrintDIDType = "Y" Then
                .DIDType = Trim(Rst("DIDType"))
            Else
                .DIDType = ""
            End If
        End With
        
        Call DIDPrintLabel(OptZebra.Value, CInt(Trim(TxtCompPort)), Trim(TxtComm))
        
        Call reFreshData
    End If
    
    
Normal_Eixt:
    ''20091222    Denver   ESBU need not clear them  ---0004
    If UCase(BU) <> "ESBU" And UCase(BU) <> "CC" Then
        txtVendorCode = ""
        txtLotCode = ""
        txtDateCode = ""
        TxtQty = ""
        TxtCompPN = ""
        TxtCompPN.SetFocus                 '---0003
    Else
        ''20101231 Maggie 清空txtCompPN,txtQty (1038)
        TxtCompPN = ""
        TxtQty = ""
        TxtCompPN.SetFocus
    End If
    
    ''1058
    cmdOK.Enabled = True
    
    Exit Sub
    
Err_Handler:
    LblMessage = Err.Number & ":" & Err.Description
    cmdOK.Enabled = True  ''1058
End Sub

Private Sub cmdReprint_Click()
      ''check printer
    If Trim(TxtCompPort) = "" Or Trim(TxtComm) = "" Then
        MsgBox "Printer have not set!!", vbInformation
        Exit Sub
    End If

    
    If gridReturnComp.row < 0 Then Exit Sub
    If Trim(gridReturnComp.Columns(0).Text) <> "" Then
        With DIDInfo
            .DID = Trim(gridReturnComp.Columns(0).Text)
            .compPN = Trim(gridReturnComp.Columns(1).Text)
            .Qty = Trim(gridReturnComp.Columns(2).Text)
            .IsGood = Trim(gridReturnComp.Columns(3).Text)
            If BU = "NB5" Then
                .WareHouseID = Trim(gridReturnComp.Columns(20).Text)        '(1252)
            End If
            If ChkPrintDIDType = "Y" Then
                .DIDType = Trim(gridReturnComp.Columns(8).Text)
            Else
                .DIDType = ""
            End If
        End With
        
        Call DIDPrintLabel(OptZebra.Value, CInt(Trim(TxtCompPort)), Trim(TxtComm))
    End If
End Sub

Private Sub Form_Load()
    Call reFreshData
    
    '20101115 Maggie Save Printer setting in local Registry (1019)
    Call GetPrinterSetting(FrmReturnComp)
    If BGAWarehouse = "Y" Then
        ChkBGA.Visible = True
    End If
     ''20100507    Denver     好坏料不让user 选择
    optGoodMaterial.Enabled = False
    optBadMaterial.Enabled = False
    
End Sub

Private Sub reFreshData()
    Dim sSql As String
    Dim Rst As ADODB.Recordset
    sSql = "exec XL_ReturnCompRefresh_T " & sq(Trim(Factory))
    Set Rst = Conn.Execute(sSql)
    Set gridReturnComp.DataSource = Rst
    
End Sub
  

Private Sub txtCompPN_Click()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtCompPN_KeyPress(KeyAscii As Integer)
Dim NewComp() As String, index As Integer
Dim strSQL As String
Dim RS As New ADODB.Recordset
    If KeyAscii = 13 And Trim(TxtCompPN) <> "" Then
        If InStr(1, Trim(TxtCompPN.Text), ";") > 0 Then
            NewComp = Split(Trim(TxtCompPN.Text), ";")
            For index = 0 To UBound(NewComp)
                If index = 0 Then
                    TxtCompPN.Text = Trim(NewComp(index))
                ElseIf index = 1 Then
                    txtDateCode.Text = Trim(NewComp(index))
                ElseIf index = 2 Then
                    txtVendorCode.Text = Trim(NewComp(index))
                ElseIf index = 3 Then
                    txtLotCode.Text = Trim(NewComp(index))
                ElseIf index = 4 Then
                    TxtQty.Text = Trim(NewComp(index))
                End If
            Next index
            'Call cmdOK_Click 'Jocelyn could change qty
        ElseIf InStr(1, Trim(TxtCompPN.Text), "-") > 0 And Len(TxtCompPN.Text) > 15 Then  '(1020)
            strSQL = "select CompPN,VendorCode,DateCode ,LotCode ,Qty from QSMS_DID_ToWH where DID = '" & Trim(TxtCompPN.Text) & "' "
            Set RS = Conn.Execute(strSQL)
            If RS.EOF = False Then
                TxtCompPN.Text = Trim(RS!compPN)
                txtVendorCode.Text = Trim(RS!VendorCode)
                txtDateCode.Text = Trim(RS!DateCode)
                txtLotCode.Text = Trim(RS!LotCode)
                TxtQty.Text = Trim(RS!Qty)
                Call cmdOK_Click
            Else
                MsgBox ("Can't find the information of this returnDID---" & Trim(TxtCompPN.Text))
                TxtCompPN.SetFocus
                Exit Sub
            End If
        End If
        If Trim(TxtCompPN) <> "" Then
            txtVendorCode.SetFocus
            Call TxtVendorCode_Click
        End If

    End If
End Sub

Private Sub TxtDateCode_Click()
    SendKeys "{home}+{end}"
End Sub

Private Sub TxtDateCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtDateCode) <> "" Then
        txtLotCode.SetFocus
        Call TxtLotCode_Click

    End If
End Sub

Private Sub TxtLotCode_Click()
    SendKeys "{home}+{end}"
End Sub

Private Sub TxtLotCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtLotCode) <> "" Then
        TxtQty.SetFocus
        Call TxtQty_Click
    End If
End Sub
 
Private Sub TxtQty_Click()
    SendKeys "{home}+{end}"
End Sub

Private Sub TxtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(TxtQty) <> "" Then
        If IsNumeric(Trim(TxtQty)) = True Then
            TxtQty = Abs(Trim(TxtQty))
            Call cmdOK_Click
        Else
            TxtQty.SetFocus
            Call TxtQty_Click
        End If
    End If
End Sub

Private Sub TxtVendorCode_Click()
    SendKeys "{home}+{end}"
End Sub

Private Sub TxtVendorCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtVendorCode) <> "" Then
        txtDateCode.SetFocus
        Call TxtDateCode_Click
    
    End If
End Sub
 
Private Function ChkValidData() As Boolean
    ChkValidData = False
    If Trim(txtVendorCode) = "" Then
        LblMessage = "Vendor Code is blank!!"
        Exit Function
    End If
    
    If Trim(txtLotCode) = "" Then
        LblMessage = "Lot Code is blank!!"
        Exit Function
    End If
    
    If Trim(txtDateCode) = "" Then
        LblMessage = "Date Code is blank!!"
        Exit Function
    End If
    
    If Trim(TxtCompPN) = "" Then
        LblMessage = "CompPN is blank!!"
        Exit Function
    End If
    If Trim(TxtQty) = "" Then
        LblMessage = "Qty is blank!!"
        Exit Function
    End If
    
    If IsNumeric(Trim(TxtQty)) = False Then
        LblMessage = "Please input numeric!!"
        Exit Function
    End If
    
    TxtQty = Abs(Int(Trim(TxtQty)))
     ''20081230  Denver   其实此处检查数量不为0
    If Trim(TxtQty) <= 0 Then
        LblMessage = "The Return Qty must be >0 !!"
        Exit Function
    End If
    
    ''2008/04/02  Denver  change to upper case  --(0002)
    ''2008/04/02  Denver  it will replace if there is "'"  --(0002)

    txtVendorCode = Replace(UCase(txtVendorCode), "'", "")
    txtLotCode = Replace(UCase(txtLotCode), "'", "")
    txtDateCode = Replace(UCase(txtDateCode), "'", "")
    TxtCompPN = Replace(UCase(TxtCompPN), "'", "")
    
    ChkValidData = True
    
End Function
        
      


''20071226 Denver Print DID for CallBack
''===================================================
'Private Function DIDPrintLabel(blnCompPort As Boolean, blnZebra As Boolean, intCompPort As Integer, sCommString As String)

    'If blnCompPort = True Then
      ' Call PrintLabelCompPort(blnZebra, intCompPort, sCommString)
    'Else
      ' Call PrintLabelPrintPort(blnZebra)
    'End If
'End Function

Private Function DIDPrintLabel(blnZebra As Boolean, intCompPort As Integer, sCommString As String)
    Dim hFile As Long
    Dim hString As String
    Dim strDID As String, tmpDID As String, strQty As String, strDIDType As String
    Dim strDay As String
    Dim LabelFile As String
    Dim lptPort As Integer
    Dim m As Integer
    Dim tmpPrintStr As String
    Dim strSQL As String
    Dim rsTime As ADODB.Recordset
    
        On Error GoTo errHandler
        '1112
        strSQL = "select getdate()"
        Set rsTime = Conn.Execute(strSQL)
        strDay = Format(rsTime(0), "YYMMDDHHNNSS")  '1101
        If blnZebra = True Then
'            LabelFile = Settings.AutoDispatchLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
'                LabelFile = Settings.DIDLabelGood
                strDID = DIDInfo.DID
            Else
'                LabelFile = Settings.DIDLabelBad
                strDID = DIDInfo.compPN
            End If
        Else
'            LabelFile = Settings.AutoDispatchSatoLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
'                LabelFile = Settings.DIDLabelSATOGood
                strDID = DIDInfo.DID
            Else
'                LabelFile = Settings.DIDLabelSATOBad
                strDID = DIDInfo.compPN
            End If
        End If
        strSQL = "select * from MSD_DATA where CompPN=" & sq(TxtCompPN.Text)    '''(1272)
        Set rsTime = Conn.Execute(strSQL)
        
        If BU = "ESBU" And rsTime.EOF = False Then
            LabelFile = GetDIDLabelFile(frmDIDCallBack_New, "GOOD_MSD")
        Else
            LabelFile = GetDIDLabelFile(FrmReturnDID, IIf(DIDInfo.IsGood = "Y", "GOOD", "BAD")) ''(1080) Get labelfile
        End If

        strDIDType = DIDInfo.DIDType
        ''特殊处理(打印RefID Label)
        If DIDInfo.Qty <= -10000 Then
            strQty = "RefID"
        Else
            strQty = CStr(DIDInfo.Qty)
        End If
     
        If (Dir(LabelFile) = vbNullString) Then
            ''''''Added by Jing 2008.01.10  (00003)'''''
            MsgBox ("Can not find Lable file !"), vbCritical
            DIDPrintLabel = "PRN_FileNoExist"
            Exit Function
        End If
        
        'TxtCompPort   TxtComm
        
        If OptComp.Value = True Then
            MSComm.CommPort = intCompPort
            MSComm.Settings = sCommString
            MSComm.OutBufferCount = 0 '清空输出缓存
            
            If MSComm.PortOpen = False Then MSComm.PortOpen = True
        ElseIf OptPrint.Value = True Then
            lptPort = OpenOutputFile("LPT1")
            If lptPort = 0 Then
                MsgBox "Open print port LPT1 error!"
                Exit Function
            End If
        End If
        
        ''(0077)
        hFile = FreeFile
        If FileReadAll(tmpPrintStr, LabelFile) <= 0 Then
            MsgBox "Open file:" & LabelFile & " fail!!", vbCritical
            Exit Function
        End If
'        Open LabelFile For Input As #hFile
'        Do
'           Select Case EOF(hFile)
'              Case True
'                Close #hFile
'                DIDPrintLabel = "PRN_Succeed"
'                Exit Do
'              Case False
'                Line Input #hFile, hString
'                hString = Trim(hString)
'                tmpPrintStr = tmpPrintStr & Trim(hString)
'            End Select
'        Loop
                
         tmpDID = Trim(strDID) '***************add by jeanson 20070814******
         'for Code 128 barcode, the ^ must be tranfer to ><
         If InStr(tmpPrintStr, "<DID_CODE>") > 0 Then
             ''********************************updated by jing 20071024 (0002) ***********
             If blnZebra Then
                 tmpDID = Replace(strDID, "^", "><")
             End If
             tmpPrintStr = Replace(tmpPrintStr, "<DID_CODE>", tmpDID)
         End If
         'for text ^, must be use ^FH_ and the use _5E (the ascii of ^)
         If InStr(tmpPrintStr, "<DID_TEXT>") > 0 Then
             ''********************************updated by jing 20071024 (0002) ***********
             If blnZebra Then
                tmpDID = Replace(strDID, "^", "_5E")
             End If
             tmpPrintStr = Replace(tmpPrintStr, "<DID_TEXT>", tmpDID)
         End If
        
         tmpPrintStr = Replace(tmpPrintStr, "<DIDType>", strDIDType)
         tmpPrintStr = Replace(tmpPrintStr, "<UID>", UID)
         tmpPrintStr = Replace(tmpPrintStr, "<DATE>", strDay)
         tmpPrintStr = Replace(tmpPrintStr, "<QTY>", strQty)
         tmpPrintStr = Replace(tmpPrintStr, "<LINE>", BUDIDShow)
         tmpPrintStr = Replace(tmpPrintStr, "<SIDE>", "")
'                hString = Replace(hString, "<MACHINE>", IIf(DIDInfo.IsGood = "Y", "", "NG"))
         tmpPrintStr = Replace(tmpPrintStr, "<MACHINE>", "NG")
         tmpPrintStr = Replace(tmpPrintStr, "<VENDORCODE1>", DIDInfo.VendorCode)    ''(1229)
         tmpPrintStr = Replace(tmpPrintStr, "<WHID>", DIDInfo.WareHouseID)          ''(1252)
         ''Debug.Print hString
         
'         MsgBox "step 1"
         
        Select Case Trim(tmpPrintStr)
           Case vbNullString
           Case Else
'                    MSComm.Output = hString
'                    Debug.Print hString
                If OptComp.Value = True Then
                    If blnZebra = True Then
'                        MsgBox "step 2"
                        '''1023
                        For m = 1 To Len(tmpPrintStr) Step 50
                            MSComm.Output = Mid(tmpPrintStr, m, 50)
                            
                        Next m
                    Else '   (0016)
'                        MsgBox "step 3"
                        For m = 1 To Len(tmpPrintStr) Step 50
                            MSComm.Output = Mid(tmpPrintStr, m, 50)
                            DoEvents
                            'Debug.Print Mid(hString, m, 50)
                        Next m
                    End If
                    MSComm.PortOpen = False
                ElseIf OptPrint.Value = True Then
                    If blnZebra = True Then
                        Print #lptPort, tmpPrintStr & Chr(13)
                    Else '   (0016)
                        ''(1027)
                        For m = 1 To Len(tmpPrintStr) Step 50
                            Print #lptPort, Mid(tmpPrintStr, m, 50)
                        Next m
                    End If
                    Close #lptPort
                Else
                    Printer.Print tmpPrintStr
                    Printer.EndDoc
                    Printer.KillDoc
                End If
             
        End Select
        
'       MsgBox "step 4"
       
        Close #hFile
        Exit Function
errHandler:
        MsgBox Err.Description
        If MSComm.PortOpen = True Then
            MSComm.PortOpen = False
        End If
End Function

Private Function PrintLabelCompPort(blnZebra As Boolean, intCompPort As Integer, sCommString As String) As String
    Dim hFile As Long
    Dim hString As String
    Dim strDID As String, tmpDID As String, strQty As String
    Dim strDay As String
    Dim LabelFile As String
    Dim m As Integer
    Dim strSQL As String
    Dim rsTime As ADODB.Recordset
        
        On Error GoTo errHandler
        '1112
        strSQL = "select getdate()"
        Set rsTime = Conn.Execute(strSQL)
        strDay = Format(rsTime(0), "YYMMDDHHNNSS") '1101
        If blnZebra = True Then
'            LabelFile = Settings.AutoDispatchLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
                LabelFile = Settings.DIDLabelGood
                strDID = DIDInfo.DID
            Else
                LabelFile = Settings.DIDLabelBad
                strDID = DIDInfo.compPN
            End If
        Else
'            LabelFile = Settings.AutoDispatchSatoLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
                LabelFile = Settings.DIDLabelSATOGood
                strDID = DIDInfo.DID
            Else
                LabelFile = Settings.DIDLabelSATOBad
                strDID = DIDInfo.compPN
            End If
        End If
        
        ''特殊处理(打印RefID Label)
        If DIDInfo.Qty <= -10000 Then
            strQty = "RefID"
        Else
            strQty = CStr(DIDInfo.Qty)
        End If
     
        If (Dir(LabelFile) = vbNullString) Then
            ''''''Added by Jing 2008.01.10  (00003)'''''
            MsgBox ("Can not find Lable file !"), vbCritical
            PrintLabelCompPort = "PRN_FileNoExist"
            Exit Function
        End If
        
        'TxtCompPort   TxtComm
        MSComm.CommPort = intCompPort
        MSComm.Settings = sCommString
        MSComm.OutBufferCount = 0 '清空输出缓存
        
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
                tmpDID = Trim(strDID) '***************add by jeanson 20070814******
                'for Code 128 barcode, the ^ must be tranfer to ><
                If InStr(hString, "<DID_CODE>") > 0 Then
                    ''********************************updated by jing 20071024 (0002) ***********
                    If blnZebra Then
                        tmpDID = Replace(strDID, "^", "><")
                    End If
                    hString = Replace(hString, "<DID_CODE>", tmpDID)
                End If
                'for text ^, must be use ^FH_ and the use _5E (the ascii of ^)
                If InStr(hString, "<DID_TEXT>") > 0 Then
                    ''********************************updated by jing 20071024 (0002) ***********
                    If blnZebra Then
                       tmpDID = Replace(strDID, "^", "_5E")
                    End If
                    hString = Replace(hString, "<DID_TEXT>", tmpDID)
                End If

                hString = Replace(hString, "<UID>", UID)
                
'                hString = Replace(hString, "<DATE>", strDay)
'                hString = Replace(hString, "<QTY>", strQTY)
'                hString = Replace(hString, "<LINE>", PrintData.Line)
'                hString = Replace(hString, "<SIDE>", PrintData.Side)
'                hString = Replace(hString, "<MACHINE>", PrintData.Machine)
                
                hString = Replace(hString, "<DATE>", strDay)
                hString = Replace(hString, "<QTY>", strQty)
                hString = Replace(hString, "<LINE>", BUDIDShow)
                hString = Replace(hString, "<SIDE>", "")
'                hString = Replace(hString, "<MACHINE>", IIf(DIDInfo.IsGood = "Y", "", "NG"))
                hString = Replace(hString, "<MACHINE>", "NG")
                ''Debug.Print hString
                
               Select Case Trim(hString)
                  Case vbNullString
                  Case Else
'                    MSComm.Output = hString
'                    Debug.Print hString

                    If blnZebra = True Then
                        MSComm.Output = hString
                    Else '   (0016)''(1027)
                        For m = 1 To Len(hString) Step 50
                            MSComm.Output = Mid(hString, m, 50)
                            'Debug.Print Mid(hString, m, 50)
                        Next m
                    End If
                    
                    
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

Private Function PrintLabelPrintPort(blnZebra As Boolean) As String
    Dim hFile As Long
    Dim hString As String
    Dim strDID As String, tmpDID As String, strQty As String
    Dim FileNum As Integer, lptPort As Integer
    Dim strDay As String
    Dim LabelFile, strLabelFileContent As String
    Dim strPort As String, PrintLabel As String
    Dim m As Integer
    Dim strSQL As String
    Dim rsTime As ADODB.Recordset
    
    On Error GoTo errHandler
             
        strSQL = "select getdate()"
        Set rsTime = Conn.Execute(strSQL)
        strDay = Format(rsTime(0), "YYMMDDHHNNSS") '1101
        
        ''特殊处理(打印RefID Label)
        If DIDInfo.Qty <= -10000 Then
            strQty = "RefID"
        Else
            strQty = CStr(DIDInfo.Qty)
        End If
        
        If blnZebra = True Then
'            LabelFile = Settings.AutoDispatchLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
                LabelFile = Settings.DIDLabelGood
                strDID = DIDInfo.DID
            Else
                LabelFile = Settings.DIDLabelBad
                strDID = DIDInfo.compPN
            End If
        Else
'            LabelFile = Settings.AutoDispatchSatoLabel
            If UCase(DIDInfo.IsGood) = "Y" Then
                LabelFile = Settings.DIDLabelSATOGood
                strDID = DIDInfo.DID
            Else
                LabelFile = Settings.DIDLabelSATOBad
                strDID = DIDInfo.compPN
                
            End If
            
        End If
'        strLabelFileContent = funGetTxtFileContent(LabelFile)
        If Dir(LabelFile) = vbNullString Then
            ''''''Added by Jing 2008.01.10  (00003)''''''
            MsgBox ("Can not find Label file !"), vbCritical
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
            tmpDID = Trim(strDID)  '***************add by jeanson 20070814******
            'for Code 128 barcode, the ^ must be tranfer to ><
            If InStr(hString, "<DID_CODE>") > 0 Then
                ''********************************updated by jing 20071024 (0002) ***********
                If blnZebra Then
                    tmpDID = Replace(strDID, "^", "><")
                End If
               hString = Replace(hString, "<DID_CODE>", tmpDID)
            End If
            'for text ^, must be use ^FH_ and the use _5E (the ascii of ^)
            If InStr(hString, "<DID_TEXT>") > 0 Then
                ''********************************updated by jing 20071024 (0002) ***********
                If blnZebra Then
                   tmpDID = Replace(strDID, "^", "_5E")
                End If
               hString = Replace(hString, "<DID_TEXT>", tmpDID)
            End If
            
            hString = Replace(hString, "<UID>", UID)
'                hString = Replace(hString, "<RACKID>", TxtRackID)
            hString = Replace(hString, "<DATE>", strDay)
            hString = Replace(hString, "<QTY>", strQty)
            hString = Replace(hString, "<LINE>", BUDIDShow)
            hString = Replace(hString, "<SIDE>", "")
            hString = Replace(hString, "<MACHINE>", IIf(DIDInfo.IsGood = "Y", "", "NG"))
            
'            Debug.Print hString
'            Print #lptPort, hString & Chr(13)
            If blnZebra = True Then
                Print #lptPort, hString & Chr(13)
            Else '   (0016) ''(1027)
                For m = 1 To Len(hString) Step 50
                    Print #lptPort, Mid(hString, m, 50)
                    'Debug.Print Mid(hString, m, 50)
                Next m
            End If
         
        Wend
'        Open strPort For Output As #FileNum
'        Print #FileNum, strLabelFileContent
        Close #FileNum
        Close #lptPort
        Exit Function
errHandler:
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
