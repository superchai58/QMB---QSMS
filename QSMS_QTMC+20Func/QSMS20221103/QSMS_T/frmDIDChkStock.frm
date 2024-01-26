VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmDIDChkStock 
   Caption         =   "Check DID Stock[20080-12-24]"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexcel 
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
      Height          =   615
      Left            =   10320
      Picture         =   "frmDIDChkStock.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1200
      Width           =   975
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
      Height          =   855
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox chkStock 
      Caption         =   "Check Qty with MC Stock"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1380
      Width           =   2535
   End
   Begin VB.TextBox txtRefID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   1320
      Width           =   5775
   End
   Begin VB.Frame frmCompPN 
      Caption         =   "CompPN Not OK"
      Height          =   4575
      Left            =   5760
      TabIndex        =   2
      Top             =   2880
      Width           =   5500
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   4215
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   5500
         _ExtentX        =   9710
         _ExtentY        =   7435
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
   End
   Begin VB.Frame fraCompPNOk 
      Caption         =   "CompPN OK"
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   5500
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4215
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   7435
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
   End
   Begin MSComCtl2.DTPicker dtpSDate 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   360
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
      Format          =   63045635
      CurrentDate     =   36482
   End
   Begin MSComCtl2.DTPicker dtpEDate 
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   360
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
      Format          =   63045635
      CurrentDate     =   36482
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
      Left            =   3960
      TabIndex        =   11
      Top             =   360
      Width           =   1455
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
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblRefID 
      Caption         =   "ReferenceID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblMsg 
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
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   10935
   End
End
Attribute VB_Name = "frmDIDChkStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/**********************************************************************************
'**文 件 名: frmDIDChkStock.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: Jeanson
'**日    期: 2007.10.01
'**描    述: Check DID Stock
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**Sandy        2008.01.25     check ReferrenceID by date,and check referenceID status---(2008012501)
'**Sandy        2008.06,05      when referenceID is empty; then query all unOK referenceID--(2008006061)
'***********************************************************************************/
Public FuncType As String
Public rstCompPN As New ADODB.Recordset
Private sSql As String


Private Sub cmdexcel_Click()
    Dim rst As ADODB.Recordset
    Dim BeginDate, EndDate As String
'    If Trim(txtRefID) <> "" Then'--(2008006061)
        BeginDate = Format(DTPsdate, "YYYY/MM/DD")
        BeginDate = Replace(BeginDate, "-", "")
        BeginDate = Replace(BeginDate, "/", "")
        EndDate = Format(DTPedate, "YYYY/MM/DD")
        EndDate = Replace(EndDate, "-", "")
        EndDate = Replace(EndDate, "/", "")
        If DTPsdate > DTPedate Then
            MsgBox ("The StartDate must be smaller than EndDate !"), vbCritical
            Exit Sub
        End If
        If chkStock.Value = 0 Then
            lblmsg = "ReferenceID:" & Trim(txtRefID) & " ToWH info is below:"
            sSql = "exec XL_DIDChkStockByRefID @Type='Query',@RefID='" & Trim(txtRefID) & "'"
            '",@BeginDate=" & sq(BeginDate) & ",@EndDate=" & sq(EndDate)
        Else
            'sSql = "exec XL_DIDChkStockByRefID @Type='Manual',@RefID=" & Trim(txtRefID) & ",@UserName=" & sq(g_userName) ' & ",@BeginDate=" & sq(BeginDate) & ",@EndDate=" & sq(EndDate)
            Exit Sub
        End If
        
        Set rstCompPN = Conn.Execute(sSql)
        If rstCompPN.EOF = False Then
            If rstCompPN("Result") <> 0 Then
                MsgBox rstCompPN("Description"), vbExclamation, "Prompt"
                Exit Sub
            End If
            Set rst = rstCompPN.NextRecordset
            Call CopyToExcel(rst)
            
        End If
        txtRefID = ""
'    End If
End Sub

Private Sub CmdQuery_Click()
    Dim rst As ADODB.Recordset
    Dim BeginDate, EndDate As String
    Dim rstOK As ADODB.Recordset
    Dim rstNotOK As ADODB.Recordset
    BeginDate = Format(DTPsdate, "YYYY/MM/DD")
    BeginDate = Replace(BeginDate, "-", "")
    BeginDate = Replace(BeginDate, "/", "")
    EndDate = Format(DTPedate, "YYYY/MM/DD")
    EndDate = Replace(EndDate, "-", "")
    EndDate = Replace(EndDate, "/", "")
    If DTPsdate > DTPedate Then
        MsgBox ("The StartDate must be smaller than EndDate !"), vbCritical
        Exit Sub
    End If
    sSql = "exec XL_CheckRefID @BeginDate=" & sq(BeginDate) & ",@EndDate=" & sq(EndDate)
    Set rstCompPN = Conn.Execute(sSql)
    If rstCompPN.EOF = False Then
        If rstCompPN("Result") <> 0 Then
            MsgBox rstCompPN("Description"), vbExclamation, "Prompt"
            Exit Sub
        End If
        Set rst = rstCompPN.NextRecordset
        frmCompPN.Caption = "ReferenceID Not OK"
        fraCompPNOk.Caption = "ReferenceID is OK"
        frmCompPN.Left = 5760
        frmCompPN.Width = 5500
        DataGrid2.Width = 5500
        Set rstOK = rst.Clone
        Set rstNotOK = rst.Clone
        rstOK.Filter = "IsPass='Y'"
        Set DataGrid1.DataSource = rstOK
        DataGrid1.Refresh
        rstNotOK.Filter = "IsPass<>'Y'"
        Set DataGrid2.DataSource = rstNotOK
        DataGrid2.Refresh
        DataGrid1.Columns(0).Width = 2500
        DataGrid2.Columns(0).Width = 2500
    End If
End Sub

Private Sub Form_Load()
    Dim rst As ADODB.Recordset
    
    If UCase(FuncType) = "AUTOCHK" Then
        chkStock.Enabled = False
        txtRefID.Enabled = False
        chkStock.Value = 1
        Call RefreshData(rstCompPN)
    Else
        chkStock.Enabled = True
        txtRefID.Enabled = True
        chkStock.Value = 0
        Set DataGrid1.DataSource = Nothing
        Set DataGrid2.DataSource = Nothing
        lblmsg = ""
    End If
    DTPsdate = Date
    DTPedate = Date
End Sub
Private Sub RefreshData(rst As ADODB.Recordset)
    Dim rstOK As ADODB.Recordset
    Dim rstNotOK As ADODB.Recordset
    
    If chkStock.Value = 0 Then
        frmCompPN.Caption = "DID ToWH Info"
        frmCompPN.Left = fraCompPNOk.Left
        frmCompPN.Width = 11000
        DataGrid2.Width = 11000
        Set DataGrid2.DataSource = rst
        DataGrid2.Refresh
    Else
        frmCompPN.Caption = "CompPN Not OK"
        frmCompPN.Left = 5760
        frmCompPN.Width = 5500
        DataGrid2.Width = 5500
        
        Set rstOK = rst.Clone
        Set rstNotOK = rst.Clone
        
'        lblMsg = "DID check qty with MC Stock is below:"
        rstOK.Filter = "IsToWH='Y'"
        Set DataGrid1.DataSource = rstOK
        DataGrid1.Refresh
        
        rstNotOK.Filter = "IsToWH<>'Y'"
        Set DataGrid2.DataSource = rstNotOK
        DataGrid2.Refresh
    End If
    
    
    
End Sub


Private Sub txtRefID_Click()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtRefID_KeyPress(KeyAscii As Integer)
    Dim rst As ADODB.Recordset
    Dim BeginDate, EndDate As String
    If KeyAscii = 13 And Trim(txtRefID) <> "" Then
        BeginDate = Format(DTPsdate, "YYYY/MM/DD")
        BeginDate = Replace(BeginDate, "-", "")
        BeginDate = Replace(BeginDate, "/", "")
        EndDate = Format(DTPedate, "YYYY/MM/DD")
        EndDate = Replace(EndDate, "-", "")
        EndDate = Replace(EndDate, "/", "")
        If DTPsdate > DTPedate Then
            MsgBox ("The StartDate must be smaller than EndDate !"), vbCritical
            Exit Sub
        End If
        If chkStock.Value = 0 Then
            lblmsg = "ReferenceID:" & Trim(txtRefID) & " ToWH info is below:"
            sSql = "exec XL_DIDChkStockByRefID @Type='Query',@RefID=" & Trim(txtRefID) '& ",@BeginDate=" & sq(BeginDate) & ",@EndDate=" & sq(EndDate)
        Else
            sSql = "exec XL_DIDChkStockByRefID @Type='Manual',@RefID=" & Trim(txtRefID) & ",@UserName=" & sq(g_userName) ' & ",@BeginDate=" & sq(BeginDate) & ",@EndDate=" & sq(EndDate)
        End If
        
        Set rstCompPN = Conn.Execute(sSql)
        If rstCompPN.EOF = False Then
            If rstCompPN("Result") <> 0 Then
                MsgBox rstCompPN("Description"), vbExclamation, "Prompt"
                Exit Sub
            End If
            lblmsg = rstCompPN("Description")
            Set rst = rstCompPN.NextRecordset
            Call RefreshData(rst)
            
        End If
        txtRefID = ""
    End If
    
End Sub
