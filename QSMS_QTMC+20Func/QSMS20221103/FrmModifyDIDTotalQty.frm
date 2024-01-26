VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmModifyDIDTotalQty 
   BackColor       =   &H0000FF00&
   Caption         =   "Modify DID Total Qty[20100826]"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   14100
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraConnection 
      BackColor       =   &H80000013&
      Caption         =   "DID maintain "
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   14415
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
         Left            =   1800
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1815
      End
      Begin VB.ComboBox CboLotCode 
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
         Left            =   8400
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5055
      End
      Begin VB.ComboBox CboDateCode 
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
         Left            =   1800
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5055
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
         Height          =   360
         Left            =   8400
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   720
         Width           =   5055
      End
      Begin VB.ComboBox CboCompPN 
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
         Left            =   1800
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   720
         Width           =   5055
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
         Left            =   8400
         TabIndex        =   9
         Top             =   1680
         Width           =   5055
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
         Left            =   1800
         TabIndex        =   8
         Text            =   "1"
         Top             =   240
         Width           =   1815
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2160
         Width           =   735
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
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2160
         Width           =   735
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
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
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   735
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2160
         Width           =   735
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
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2160
         Width           =   735
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
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2160
         Width           =   735
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
         TabIndex        =   21
         Top             =   1680
         Width           =   1575
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
         Left            =   6840
         TabIndex        =   20
         Top             =   1200
         Width           =   1575
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
         Left            =   6840
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080FF80&
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
         TabIndex        =   18
         Top             =   720
         Width           =   1575
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
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
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
         Left            =   6840
         TabIndex        =   16
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FF80&
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
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSDataGridLib.DataGrid DG1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   7223
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
End
Attribute VB_Name = "FrmModifyDIDTotalQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CommandType As Long

Private Sub CboCompPN_Click()
    CboVendorCode.SetFocus
End Sub

Private Sub CboCompPN_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13 Or KeyAscii = 9) And CboCompPN <> "" Then
'******************************
'****add by jeanson 2007/09/04
    CboCompPN.Text = Replace(Replace(Replace(CboCompPN.Text, " ", ""), vbCr, ""), vbLf, "")
    strErrMessage = ""
    strErrMessage = FunPartNumberCheck(CboCompPN.Text)
    If strErrMessage <> "PASS" Then
        MsgBox strErrMessage
        CboCompPN.SetFocus
        Exit Sub
    End If
'******************************
'******************************
   CboCompPN_Click
End If
End Sub

Private Sub CboDateCode_Click()
CboLotCode.SetFocus
End Sub

Private Sub CboDateCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboDateCode_Click
End If
End Sub

Private Sub CboDID_Click()
Call cmdFind_Click
End Sub

Private Sub CboLotCode_Click()
TxtQty.SetFocus
End Sub

Private Sub CboLotCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboLotCode_Click
End If
End Sub

Private Sub CboVendorCode_Click()
CboDateCode.SetFocus
End Sub

Private Sub CboVendorCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboVendorCode_Click
End If
End Sub



Private Sub cmdCancel_Click()
CboCompPN.Text = ""
CboVendorCode.Text = ""
CboDateCode.Text = ""
CboLotCode.Text = ""
TxtQty.Text = ""
CboDID.Text = ""

End Sub



Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
 Dim Rs As ADODB.Recordset
    strSql = "Select DID,CompPN,VendorCode,DateCode,LotCOde,Qty,UID,remainQty,TransDateTime,UsedFlag From QSMS_DID Where DID like '" & Trim(CboDID) & "%' " & _
              " Order by CompPN,DID  "
    Set Rs = Conn.Execute(strSql)
    Set DG1.DataSource = Rs
    
    cmdUpdate.Enabled = True
    
    cmdSave.Enabled = True
End Sub

Private Sub CmdRefresh_Click()
Call RefreshDg("")
End Sub



Private Sub cmdSave_Click()
    Dim strSql As String
    Dim Rs As ADODB.Recordset
    Dim TempDID As String
    Dim TransDate As String
    Dim i As Long
    Dim intDIDInitQty As Long    'add by jeanson 20071008
    Dim strlog As String
    
    cmdFind.Enabled = True
    cmdUpdate.Enabled = True
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdExit.Enabled = True
    If Trim(TxtQty) = "" Then
        MsgBox (" qty can't be empty!!"), vbCritical
        CboCompPN.Enabled = True
        CboCompPN.SetFocus
        Exit Sub
    End If
    strSql = "select getdate()"
    Set Rs = Conn.Execute(strSql)
    TransDate = Format(Rs(0), "YYYYMMDDHHMMSS")
    Select Case CommandType
        Case 2
            If Trim(CboDID) = "" Then
                MsgBox ("DID can't be empty!!"), vbCritical
                CboDID.Enabled = True
                CboDID.SetFocus
                Exit Sub
            End If
            TempDID = Trim(CboDID)
            strSql = "Select Qty,RemainQty from QSMS_DID where DID='" & Trim(TempDID) & "'"
            Set Rs = Conn.Execute(strSql)
            If Not Rs.EOF Then
                
                ''20080815  Denver  do not check update Qty
'                If Rs!Qty * 1.1 < CLng(Trim(TxtQty)) Then
'                   MsgBox "Total DID too large,Please check."
'                   Exit Sub
'                End If
                intDIDInitQty = Rs!Qty
            Else
                   MsgBox "Without DID Total Qty, please contact QMS."
                   Exit Sub
            End If
            
        
            
            
'            strSql = "Update QSMS_DID Set UID='" & g_userName & "',Qty='" & Trim(txtQty) & "' Where DID='" & Trim(CboDID) & "'"
            strSql = "Update QSMS_DID Set UID='" & g_userName & "',Qty='" & Trim(TxtQty) & "',RealQty='" & Trim(TxtQty) & "' Where DID='" & Trim(CboDID) & "'"
            Conn.Execute strSql
            
            strSql = "QTY: " + Trim(intDIDInitQty) + " -> " + Replace(strSql, "'", " ")
            strlog = "insert into qms_log(system_name,event_no,sn,user_name,desc1,trans_date) values('ModifyDIDTotalQty','1','" & Trim(CboDID) & "','" & g_userName & "','" & strSql & "',dbo.formatdate(getdate(),'YYYYMMDDHHNNSS'))"
            Conn.Execute strlog
            
             If Trim(CboDID) = "" Then
                      CboDID = TempDID
             End If
            
            Call cmdFind_Click
    End Select
    Call RefreshDg("")
 
    CommandType = 0
    TxtGroupQty = 1
    Call cmdCancel_Click
End Sub

Private Sub cmdUpdate_Click()
  
    cmdUpdate.Enabled = True
 
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdExit.Enabled = True
    cmdFind.Enabled = True
    
    CboCompPN.Enabled = True
    CboVendorCode.Enabled = True
    CboDateCode.Enabled = True
    CboLotCode.Enabled = True
    TxtQty.Enabled = True
    

    CommandType = 2
End Sub

Private Sub DG1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    On Error Resume Next
    With DG1
        CboDID = .Columns(0).Value
        CboCompPN = .Columns(1).Value
        CboVendorCode = .Columns(2).Value
        CboDateCode = .Columns(3).Value
        CboLotCode = .Columns(4).Value
        TxtQty = .Columns(5).Value
    End With
    cmdUpdate.Enabled = True
    'cmdDelete.Enabled = True
    cmdCancel.Enabled = True
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
Call RefreshDg("")
End Sub

Private Function RefreshDg(ByVal CompPN As String)
Dim Str As String
Dim Rs As ADODB.Recordset
'Str = "select DID,CompPN,VendorCode,DateCode,LotCode,Qty,RemainQty,UID,TransDateTime from QSMS_DID where CompPN like '" & CompPN & "%' order by DID"
Str = "select top 20 DID,CompPN,VendorCode,DateCode,LotCode,Qty,RemainQty,UID,TransDateTime from QSMS_DID where CompPN like '" & CompPN & "%' order by DID"     '******Modify by jeason 2007/08/06
Set Rs = Conn.Execute(Str)
Set DG1.DataSource = Rs
DG1.Refresh
End Function
Private Sub TxtQty_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13 Or KeyAscii = 9) And TxtQty <> "" Then
    TxtQty = GetDIDQty(TxtQty)
End If
End Sub
Function GetDIDQty(Qty As String) As String
Dim i As Long
Dim strQty As String
    For i = 1 To Len(Qty)
        If IsNumeric(Mid(Qty, i, 1)) Then
            strQty = strQty + Mid(Qty, i, 1)
        End If
    Next i
    GetDIDQty = strQty
End Function
'******************************


