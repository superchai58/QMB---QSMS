VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FrmQryReturnedDID 
   Caption         =   "QueryDispatchOfReturnedDID 20091222"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   12825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDQuery 
      BackColor       =   &H0000FF00&
      Caption         =   "Query"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4605
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   8123
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Query Dispatch Reprot By Returned DID"
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12645
      Begin VB.CommandButton cmdqry2 
         BackColor       =   &H0000FF00&
         Caption         =   "Query-A"
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
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TXTNEWDID 
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   960
         Width           =   4575
      End
      Begin VB.CommandButton CMDexcel 
         BackColor       =   &H000080FF&
         Caption         =   "&To excel"
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
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1335
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
         Left            =   1680
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "DID-A"
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
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
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
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmQryReturnedDID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************************************************************
'**文 件 名: FrmQryReturnedDID.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人:
'**日    期:
'**描    述:
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**Kane       2007.11.15     Show all qsms_groupdid columns on interface--------(0001)
'***********************************************************************************/
Option Explicit
Dim str As String
Dim rs As ADODB.Recordset

Dim sSql  As String

Private Sub cmdExcel_Click()

On Error GoTo errHandler:

    cmdExcel.Enabled = False
    

    
    If rs.State = 1 Then
       If rs.RecordCount <> 0 Then
            rs.MoveFirst
            Call CopyToExcel(rs)
           ' Call OutPutExcel(Rs, xlApplication, "Detail")
       Else
          MsgBox "No data to Excel!!", vbCritical
       End If
    End If
    
   
    Set rs = Nothing

 cmdExcel.Enabled = True
 
Exit Sub

errHandler:
    MsgBox "Please click the Query button first !"
    cmdExcel.Enabled = True
End Sub

Private Sub cmdqry2_Click()
    If Trim(TXTNEWDID) <> "" Then
        sSql = "exec  QSMS_ReturnQry  'NewDID', " & sq(Trim(TXTNEWDID))
        Set rs = Conn.Execute(sSql)
        If rs.EOF = True Then
            MsgBox "Query fail"
        
        Else
            If rs("Result") <> 0 Then
                MsgBox Trim(rs("Description") & "")
                 
            Else
                Set rs = rs.NextRecordset
                Set DataGrid1.DataSource = rs
            
            End If
        End If
    End If
End Sub

Private Sub cmdQuery_Click()
    If Trim(txtDID) <> "" Then
        sSql = "exec  QSMS_ReturnQry  'ReturnDID', " & sq(Trim(txtDID))
        Set rs = Conn.Execute(sSql)
        If rs.EOF = True Then
            MsgBox "Query fail"
        
        Else
            If rs("Result") <> 0 Then
                MsgBox Trim(rs("Description") & "")
                 
            Else
                Set rs = rs.NextRecordset
                Set DataGrid1.DataSource = rs
            
            End If
        End If
    End If
End Sub

'Private Sub cmdqry2_Click()
'Dim TEMPDID As String
'If TXTNEWDID.Text <> "" Then
'    Set DataGrid1.DataSource = Nothing
'
'    Str = "select * From qsms_groupdid where NEWDID='" & TXTNEWDID & "' AND ReturnFlag='Y'"  '--------------0001
'    Set Rs = Conn.Execute(Str)
'    If Rs.EOF Then
'        MsgBox "Please make sure the DID has been Returned, only Returned DID can be accept!"
'        Exit Sub
'    End If
'
'    Set DataGrid1.DataSource = Rs
'
'End If
'End Sub

'Private Sub CmdQuery_Click()
'Dim TEMPDID As String
'Dim tempTransdatetime As String
'If TxtDID.Text <> "" Then
'    Set DataGrid1.DataSource = Nothing
'    Str = "SELECT DID,transdatetime FROM QSMS_DId WHERE DID='" & Trim(TxtDID) & "'"
'    Set Rs = Conn.Execute(Str)
'    If Rs.EOF Then
'        MsgBox "Can't find the DID info!"
'        Exit Sub
'    End If
'    TEMPDID = Rs!DID
'    tempTransdatetime = Rs!TransDateTime
'    Str = "select NewDID,DIDDateTime From QSMS_GroupDID where DID='" & TEMPDID & "' AND transdatetime='" & tempTransdatetime & "' AND ReturnFlag='Y'"
'    Set Rs = Conn.Execute(Str)
'    If Rs.EOF Then
'        MsgBox "Please make sure the DID has been Returned, only Returned DID can be accept!"
'        Exit Sub
'    End If
'    TEMPDID = Rs!NewDID
'    tempTransdatetime = Rs!DIDDateTime
'    Str = "select * from qsms_dispatch where did='" & TEMPDID & "' and diddatetime ='" & tempTransdatetime & "'"
'
'    Set Rs = Conn.Execute(Str)
'    If Rs.EOF Then
'        MsgBox "No find this DID dispatch records !"
'        Exit Sub
'    End If
'    Set DataGrid1.DataSource = Rs
'
'End If
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub



Private Sub TxtDID_Click()
SendKeys "{HOME}+{END}"
End Sub



