VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDummyECN 
   Caption         =   "Dummy ECN[20110630]"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3585
      Left            =   90
      TabIndex        =   5
      Top             =   60
      Width           =   11445
      Begin VB.CommandButton cmdDeleteECN 
         Caption         =   "Delete ECN"
         Height          =   435
         Left            =   6480
         TabIndex        =   23
         Top             =   2820
         Width           =   1395
      End
      Begin VB.ComboBox cboModelName 
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
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   210
         Width           =   2625
      End
      Begin VB.TextBox txtPN 
         BackColor       =   &H8000000A&
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
         Height          =   375
         Left            =   1350
         TabIndex        =   14
         Top             =   630
         Width           =   2625
      End
      Begin VB.TextBox txtRevision 
         BackColor       =   &H8000000A&
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
         Height          =   375
         Left            =   1350
         TabIndex        =   13
         Top             =   1020
         Width           =   2625
      End
      Begin VB.ComboBox cboJobPN 
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
         Left            =   5610
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   210
         Width           =   2625
      End
      Begin VB.TextBox txtCompPN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1350
         TabIndex        =   11
         Top             =   1740
         Width           =   2625
      End
      Begin VB.TextBox txtNewCompPN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1350
         TabIndex        =   10
         Top             =   2220
         Width           =   2625
      End
      Begin VB.ListBox LstItem 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   4290
         TabIndex        =   9
         Top             =   1020
         Width           =   7005
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Item"
         Height          =   435
         Left            =   720
         TabIndex        =   8
         Top             =   2820
         Width           =   1395
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Item"
         Height          =   435
         Left            =   2160
         TabIndex        =   7
         Top             =   2820
         Width           =   1395
      End
      Begin VB.CommandButton cmdCreateECN 
         Caption         =   "Create ECN"
         Height          =   435
         Left            =   4740
         TabIndex        =   6
         Top             =   2820
         Width           =   1395
      End
      Begin VB.Label lblModelName 
         Caption         =   "ModelName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   22
         Top             =   210
         Width           =   1245
      End
      Begin VB.Label lblPN 
         Caption         =   "Part Num"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   21
         Top             =   660
         Width           =   1245
      End
      Begin VB.Label lblRevision 
         Caption         =   "Revision"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   20
         Top             =   1050
         Width           =   1245
      End
      Begin VB.Label lblJobPN 
         Caption         =   "Job PN"
         Height          =   345
         Left            =   4380
         TabIndex        =   19
         Top             =   210
         Width           =   1245
      End
      Begin VB.Label lblCompPN 
         Caption         =   "CompPN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   18
         Top             =   1770
         Width           =   1245
      End
      Begin VB.Label Label2 
         Caption         =   "New CompPN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   17
         Top             =   2250
         Width           =   1455
      End
      Begin VB.Label lblDummyECN 
         Caption         =   "Dummy ECN List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4320
         TabIndex        =   16
         Top             =   630
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8940
      TabIndex        =   4
      Top             =   5550
      Width           =   2445
   End
   Begin VB.TextBox txtFilter 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8940
      TabIndex        =   2
      Top             =   4140
      Width           =   2445
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Query"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8940
      TabIndex        =   0
      Top             =   5070
      Width           =   2445
   End
   Begin MSDataGridLib.DataGrid gridDummyECN 
      Height          =   3765
      Left            =   90
      TabIndex        =   1
      Top             =   3720
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   6641
      _Version        =   393216
      DefColWidth     =   100
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
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
   Begin VB.Label lblFilter 
      Alignment       =   2  'Center
      Caption         =   "Filter MBPN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8940
      TabIndex        =   3
      Top             =   3810
      Width           =   2445
   End
End
Attribute VB_Name = "frmDummyECN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/**********************************************************************************
'**文 件 名: DummyECN.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: Denver Yang
'**日    期: 2011.06.20
'**描    述: For Dummy ECN
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------


'***********************************************************************************

Dim rst As ADODB.Recordset
Dim rstDummyECN As ADODB.Recordset
Dim sSql As String

Private Sub cboModelName_Click()
    Call GetModelInfo(Trim(cboModelName))
    Call GetJobInfo
    LstItem.Clear
    txtCompPN = ""
    txtNewCompPN = ""
End Sub

Private Sub cmdAdd_Click()
Dim CompList As String

    If Trim(txtPN) = "" Or Trim(txtRevision) = "" Or Trim(cboJobPN) = "" Then
        MsgBox "Please confirm PN/Revision/JobPN information!"
        Exit Sub
    End If
    
    If Trim(txtCompPN) = "" Or Trim(txtNewCompPN) = "" Then
        MsgBox "Please Input use component firstly!"
        Exit Sub
    End If
    
    ''Add Replace Chr10/Chr13
    txtCompPN = Replace(Replace(Trim(txtCompPN), Chr(10), ""), Chr(13), "")
    txtNewCompPN = Replace(Replace(Trim(txtNewCompPN), Chr(10), ""), Chr(13), "")
    
    If UCase(Trim(txtCompPN)) = UCase(Trim(txtNewCompPN)) Then
        MsgBox "CompPN and NewCompPN can not be same!"
        Exit Sub
    End If
    
    If ChkExists = True Then Exit Sub
    
    CompList = Trim(cboJobPN) & "-->" & Trim(txtCompPN) & "-->" & Trim(txtNewCompPN)
    CompList = UCase(CompList)
    
    LstItem.AddItem CompList
    txtCompPN = ""
    txtNewCompPN = ""
    txtCompPN.SetFocus
        
End Sub

Private Sub cmdCreateECN_Click()
Dim ItemList As String

    If Trim(txtPN) = "" Or Trim(txtRevision) = "" Or Trim(cboJobPN) = "" Then
        MsgBox "Please confirm PN/Revision/JobPN information!"
        Exit Sub
    End If
    
    If LstItem.ListCount < 1 Then Exit Sub
    ItemList = GetItemList()
    
    sSql = "exec DummyECNSave  'Add'," & sq(Trim(txtPN)) & "," & sq(Trim(txtRevision)) & "," & sq(Trim(ItemList)) & "," & sq(g_userName)
    Set rst = Conn.Execute(sSql)
    If rst.EOF = False Then
        If rst("Result") = 0 Then
            LstItem.Clear
        End If
        MsgBox Trim(rst("Description") & "")
    Else
        MsgBox "Save Dummy ECN Fail!"
    End If
    
End Sub

Private Sub cmdDelete_Click()
    Dim I As Integer
    If LstItem.ListCount <= 0 Then Exit Sub
    If LstItem.ListIndex < 0 Then Exit Sub
    I = LstItem.ListIndex
    
    LstItem.RemoveItem I
    If LstItem.ListCount > 0 Then
        If LstItem.ListCount - 1 >= I Then
            LstItem.ListIndex = I
        Else
            LstItem.ListIndex = LstItem.ListCount - 1
        End If
    End If
    
End Sub

Private Sub cmdDeleteECN_Click()
    If Trim(txtPN) = "" Or Trim(txtRevision) = "" Then
        MsgBox "Please select Model or GroupID firstly!"
        Exit Sub
    End If
    
    If MsgBox("Do you want to delete the Dummy ECN?", vbYesNo) = vbYes Then
        sSql = "exec DummyECNSave  'Delete'," & sq(Trim(txtPN)) & "," & sq(Trim(txtRevision)) & ",''," & sq(g_userName)
        Set rst = Conn.Execute(sSql)
        If rst.EOF = False Then
            If rst("Result") = 0 Then
                LstItem.Clear
            End If
            MsgBox Trim(rst("Description") & "")
        Else
            MsgBox "Save Dummy ECN Fail!"
        End If
    
    End If
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQuery_Click()
    sSql = "Select * from DummyECN Order by TransDateTime desc "
    Set rstDummyECN = Conn.Execute(sSql)
    Set gridDummyECN.DataSource = rstDummyECN
    gridDummyECN.Refresh
    
    txtFilter = ""
End Sub

 

Private Sub Form_Load()
    Call GetModelInfo("")
    
    Call GetJobInfo
End Sub

Private Sub GetModelInfo(sModel As String)
    If sModel = "" Then
        sSql = "Select * from ModelName Order by ModelName "
        
        Set rst = Conn.Execute(sSql)
        cboModelName.Clear
        Do While rst.EOF = False
        
            cboModelName.AddItem Trim(rst("ModelName") & "")
            rst.MoveNext
        Loop
        If cboModelName.ListCount > 0 Then
            cboModelName.ListIndex = 0
        End If
        
    Else
        sSql = "Select * from ModelName where ModelName= " & sq(sModel)
        Set rst = Conn.Execute(sSql)
        If rst.EOF = False Then
            txtPN = Trim(rst("PN") & "")
            txtRevision = Trim(rst("Revision") & "")
        Else
            txtPN = ""
            txtRevision = ""
        End If
    End If
End Sub

Private Sub GetJobInfo()
    If Trim(txtPN) = "" Then
        MsgBox "Please select Model firstly!"
        Exit Sub
    End If
    '''1189
    sSql = "exec UniReport_QueryJobBom  " & sq(Trim(txtPN)) & ""
    
    Set rst = Conn.Execute(sSql)
    cboJobPN.Clear
    Do While rst.EOF = False
        cboJobPN.AddItem Trim(rst("JobPN") & "")
        rst.MoveNext
    Loop
     
End Sub

Private Function ChkExists() As Boolean
    Dim I As Integer
    Dim CompList As String
    
    ChkExists = False
    
    CompList = Trim(cboJobPN) & "-->" & Trim(txtCompPN) & "-->" & Trim(txtNewCompPN)
    CompList = UCase(CompList)
    
    For I = 0 To LstItem.ListCount - 1
        LstItem.ListIndex = I
        If CompList = LstItem.Text Then
            ChkExists = True
            MsgBox "the Compoment:" & CompList & " exists!"
            Exit For
        End If
    Next I
    
End Function

Private Function GetItemList() As String
    Dim I As Integer
    Dim CompList As String
    
    For I = 0 To LstItem.ListCount - 1
        'LstItem.ListIndex = I
        LstItem.Selected(I) = True
        CompList = LstItem.Text
        GetItemList = GetItemList & Replace(CompList, "-->", ",") & "@@"
         
    Next I
    
    If GetItemList <> "" Then
        GetItemList = Mid(GetItemList, 1, Len(GetItemList) - 2)
    End If
    
End Function
 

 

Private Sub gridDummyECN_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rstTemp As ADODB.Recordset
On Error GoTo errHandler

     With gridDummyECN
        If Not IsNumeric(LastRow) Then Exit Sub
        If .row < 0 Then Exit Sub
        If rstDummyECN.State = False Then Exit Sub
        If Trim(.Columns("0")) = "" Then Exit Sub
        
        LstItem.Clear
        Set rstTemp = rstDummyECN.Clone
        
        rstTemp.Filter = "MBPN=" & Trim(.Columns("MBPN").Text)
        
        Do While rstTemp.EOF = False
            LstItem.AddItem Trim(rstTemp("JobPN")) & "-->" & Trim(rstTemp("CompPN")) & "-->" & Trim(rstTemp("NewCompPN"))
            rstTemp.MoveNext
        Loop
         
         
        txtPN = Trim(.Columns("MBPN").Text)
        txtRevision = Trim(.Columns("Version").Text)
    End With
    
    Exit Sub
errHandler:
    Err.Clear

End Sub

Private Sub txtCompPN_Click()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtCompPN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtCompPN) > "" Then
        txtNewCompPN.SetFocus
    
    End If
End Sub


Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If rstDummyECN.State Then
            If Trim(txtFilter) = "" Then
                rstDummyECN.Filter = ""
            Else
                rstDummyECN.Filter = " MBPN like " & sq(Trim(txtFilter) & "%")
            End If
            Set gridDummyECN.DataSource = rstDummyECN
            gridDummyECN.Refresh
            
        End If
    End If
End Sub

Private Sub txtNewCompPN_Click()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtNewCompPN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtNewCompPN) > "" Then
        cmdAdd.SetFocus
    
    End If
End Sub
