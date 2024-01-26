VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmMaintainWOSeq 
   Caption         =   "Maintain WO Seq [20160927]"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "Select Work Order"
      Height          =   9015
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.ComboBox CboLine 
         Height          =   315
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Add"
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
         Left            =   6120
         Picture         =   "FrmMaintainWOSeq.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Add the Wo to selected Group"
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset"
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
         Left            =   6120
         Picture         =   "FrmMaintainWOSeq.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   6000
         Width           =   975
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
         Height          =   855
         Left            =   6120
         Picture         =   "FrmMaintainWOSeq.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3960
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Find By WO  Group ID"
         Height          =   3375
         Left            =   6840
         TabIndex        =   17
         Top             =   240
         Width           =   3975
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
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
            Picture         =   "FrmMaintainWOSeq.frx":0B8E
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   2400
            Width           =   975
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
            Left            =   840
            TabIndex        =   31
            Top             =   1920
            Width           =   3015
         End
         Begin VB.ComboBox CboWo 
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
            Left            =   840
            TabIndex        =   28
            Top             =   1320
            Width           =   3015
         End
         Begin VB.CommandButton CmdQueryID 
            Caption         =   "Query ID"
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
            Left            =   2160
            TabIndex        =   22
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton optGroup 
            Caption         =   "Group"
            Height          =   255
            Left            =   1320
            TabIndex        =   21
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton OptRelease 
            Caption         =   "Release"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   975
         End
         Begin VB.ComboBox CboGroupID 
            Height          =   315
            Left            =   840
            TabIndex        =   18
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
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
            Index           =   7
            Left            =   240
            TabIndex        =   30
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label4 
            BackColor       =   &H0000FF00&
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
            Index           =   5
            Left            =   240
            TabIndex        =   29
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label4 
            BackColor       =   &H0080FF80&
            Caption         =   "ID"
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
            Left            =   240
            TabIndex        =   19
            Top             =   720
            Width           =   495
         End
      End
      Begin VB.ListBox lstWO_SELECT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6060
         Left            =   3720
         MultiSelect     =   2  'Extended
         TabIndex        =   15
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CommandButton cmdDELALL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   4680
         Width           =   495
      End
      Begin VB.CommandButton cmdDEL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4200
         Width           =   495
      End
      Begin VB.CommandButton cmdADDALL 
         BackColor       =   &H00C0C0C0&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3720
         Width           =   495
      End
      Begin VB.CommandButton cmdADD 
         BackColor       =   &H00C0C0C0&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3240
         Width           =   495
      End
      Begin VB.ListBox lstWO_LIST 
         Height          =   5910
         Left            =   0
         TabIndex        =   9
         Top             =   2520
         Width           =   3015
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
         Left            =   1680
         TabIndex        =   3
         Top             =   840
         Width           =   2895
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
         TabIndex        =   2
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton CmdQuery 
         Caption         =   "Query WO"
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
         Left            =   4920
         TabIndex        =   1
         Top             =   1320
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpSDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
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
         Format          =   64749571
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker dtpEDate 
         Height          =   375
         Left            =   4920
         TabIndex        =   26
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
         Format          =   64749571
         CurrentDate     =   36482
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
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
         Index           =   1
         Left            =   3480
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "WO Selected Seq #"
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
         Left            =   3720
         TabIndex        =   16
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total WO---without Group ID"
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
         Left            =   0
         TabIndex        =   10
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
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
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
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
         Left            =   120
         TabIndex        =   7
         Top             =   840
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
         Left            =   4920
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   5
         Top             =   1320
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmMaintainWOSeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Wo_TransDate As String
Private strLine As String

Private Sub CboGroupID_Click()
Call GetGroupWO(CboGroupID)
End Sub

Private Sub CboGroupID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 9 Then
   Call CboGroupID_Click
End If
End Sub
Private Sub CboLine_DropDown()
    ''0069
    lstWO_SELECT.Clear
    lstWO_LIST.Clear
End Sub

Private Sub CboWo_Click()
txtWO.text = cboWO.text
Call GetWoinfo(txtWO.text)
End Sub

Private Sub CmdADD_Click()
    Dim Pointer As Long
    Dim Rs As New ADODB.Recordset
    Dim i As Integer
    
    If lstWO_LIST.ListCount <= 0 Then Exit Sub
    If lstWO_LIST.ListIndex < 0 Then Exit Sub
    
'''''''''(000011)
    strSQL = "select WO from Sap_Wo_List where [group] in (select [group] from sap_wo_list where wo ='" & Trim(lstWO_LIST.text) & "')"
    Set Rs = Conn.Execute(strSQL)

    While Not Rs.EOF
        i = 0
        While i < lstWO_LIST.ListCount
           If Trim(lstWO_LIST.List(i)) = Trim(Rs!WO) Then
              Pointer = i
              lstWO_SELECT.AddItem Trim(lstWO_LIST.List(i))
              lstWO_LIST.RemoveItem Pointer
           End If
           i = i + 1
        Wend
        
        Rs.MoveNext
    Wend
    
    If lstWO_LIST.ListCount > 0 Then
        lstWO_LIST.ListIndex = 0
    End If
    
'    Pointer = lstWO_LIST.ListIndex
'    lstWO_SELECT.AddItem Trim(lstWO_LIST.Text)
'    lstWO_LIST.RemoveItem Pointer
'    If lstWO_LIST.ListCount <> Pointer Then
'       lstWO_LIST.ListIndex = Pointer
'    End If
    
End Sub

Private Sub cmdADDALL_Click()

    If lstWO_LIST.ListCount <= 0 Then Exit Sub
    
    Do While lstWO_LIST.ListCount > 0
  
      lstWO_LIST.ListIndex = 0
      lstWO_SELECT.AddItem Trim(lstWO_LIST.text)
      lstWO_LIST.RemoveItem 0
   
    Loop
    
End Sub

Private Sub cmdDel_Click()
    Dim Pointer As Long
    If lstWO_SELECT.ListCount <= 0 Then Exit Sub
    If lstWO_SELECT.ListIndex < 0 Then Exit Sub
    Pointer = lstWO_SELECT.ListIndex
    If ChkDelete(Trim(lstWO_SELECT.text)) = True Then
        lstWO_LIST.AddItem Trim(lstWO_SELECT.text)
        lstWO_SELECT.RemoveItem Pointer
        If lstWO_SELECT.ListCount <> Pointer Then
           lstWO_SELECT.ListIndex = Pointer
        End If
        
    Else
        MsgBox "The Wo has dispatch or is dispathing"
    End If
End Sub

Private Sub cmdDELALL_Click()
'Dim I As Long
'Dim WoStr As String
'I = 0
    If lstWO_SELECT.ListCount <= 0 Then Exit Sub
    Do While lstWO_SELECT.ListCount > 0
        lstWO_SELECT.ListIndex = 0
        'If ChkDelete(Trim(lstWO_SELECT.Text)) = True Then
       
           lstWO_LIST.AddItem Trim(lstWO_SELECT.text)
           lstWO_SELECT.RemoveItem 0
       ' Else
           ' MsgBox "The Wo has dispatch or is dispathing: " & lstWO_SELECT.Text
       ' End If
        'I = I + 1
    Loop
    
End Sub

Private Function ListWO()
Dim str As String
Dim BeginDate As String
Dim EndDate As String
Dim Rs As ADODB.Recordset
BeginDate = Format(dtpSDate, "YYYY/MM/DD")
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
BeginDate = BeginDate + "000000"

EndDate = Format(dtpEDate, "YYYY/MM/DD")
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")
EndDate = EndDate + "240000"
str = "select  WO from Sap_Wo_List where Trans_Date between '" & BeginDate & "' and '" & EndDate & "' and line like '" & Trim(CboLine) & "%'"
If StrBU = "NB5" Or StrBU = "NB3" Then ''(1124)(1158)
    str = "select  WO from Sap_Wo_List where WO_Type='PP10' and Trans_Date between '" & BeginDate & "' and '" & EndDate & "' and line like '" & Trim(CboLine) & "%'"
End If
Set Rs = Conn.Execute(str)
lstWO_LIST.Clear
While Not Rs.EOF
      If GetGroupID(Rs!WO) = "" Then
          lstWO_LIST.AddItem Trim(Rs!WO)
      End If
      Rs.MoveNext
Wend
End Function


Private Function GetGroupIDByDate()
Dim str As String
Dim BeginDate As String
Dim EndDate As String
Dim Rs As ADODB.Recordset
BeginDate = Format(dtpSDate, "YYYY/MM/DD")
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
EndDate = Format(dtpEDate, "YYYY/MM/DD")
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")
If OptRelease.Value = True Then
   str = "select distinct GroupID from QSMS_WOGroup A  where wo_TransDateTime between '" & BeginDate & "' and '" & EndDate & "' " & _
         "and Line='" & Trim(CboLine) & "' and exists(select 0 from QSMS_WOGroup B where a.GroupID=b.GroupID and b.Closedflag='N') "
Else
    str = "select distinct GroupID from QSMS_WOGroup A where substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and " & _
          "'" & EndDate & "' and Line='" & Trim(CboLine) & "' and Line='" & Trim(CboLine) & "' and " & _
          "exists(select 0 from QSMS_WOGroup B where a.GroupID=b.GroupID and b.Closedflag='N') "
End If

Set Rs = Conn.Execute(str)
CboGroupID.Clear
While Not Rs.EOF
      CboGroupID.AddItem Trim(Rs!GroupID)
      Rs.MoveNext
Wend
End Function
Private Function GetGroupWO(ByVal GroupID As String)
Dim str As String
Dim TransDate As String
Dim Rs As ADODB.Recordset

str = "select Work_Order from QSMS_WOGroup  where GroupID= '" & GroupID & "' order by Seq_NO"

Set Rs = Conn.Execute(str)
cboWO.Clear
While Not Rs.EOF
      cboWO.AddItem Trim(Rs!Work_Order)
      Rs.MoveNext
Wend

End Function
Private Function GenGroupID() As String
Dim str As String
Dim TransDate As String
Dim TempGroupHead As String
Dim Rs As ADODB.Recordset
str = "select getdate()"
Set Rs = Conn.Execute(str)
TransDate = Format(Rs(0), "YYYYMMDD")
If NewGroupIDRule = "Y" Then
    TempGroupHead = UCase(Trim(CboLine)) & Mid(TransDate, 3, 2)
Else
    TempGroupHead = UCase(Trim(CboLine)) & TransDate
End If
str = "select top 1 GroupID  from QSMS_WOGroup  where GroupID like '" & TempGroupHead & "%' order by GroupID desc"
Set Rs = Conn.Execute(str)
If Rs.EOF Then
   GenGroupID = TempGroupHead & "0001"
Else
   GenGroupID = TempGroupHead & Format(CLng(Right(Trim(Rs!GroupID), 4)) + 1, "0000")
End If

End Function
Private Function DblChkLine(ByVal WO As String, ByVal Line As String) As Boolean
Dim str As String
Dim Rs As ADODB.Recordset

str = "select WO from SAP_WO_list Where WO='" & WO & "' and line='" & Line & "'"
Set Rs = Conn.Execute(str)
If Rs.EOF Then
  DblChkLine = False
Else
 DblChkLine = True
End If
End Function

Private Sub cmdDelete_Click()
Dim str As String
Dim Rs As ADODB.Recordset
Dim Seq As Long
Dim sMsg As String

If ChkDelete(Trim(cboWO)) = True Then ''1209
    str = "exec DeleteWOByGroup '" & Trim(CboGroupID) & "','" & Trim(cboWO) & "','" & Trim(g_userName) & "'"
    Set Rs = Conn.Execute(str)
    If Not Rs.EOF Then
            sMsg = Rs!Description
            MsgBox sMsg
    Else
        MsgBox "can not find the Wo :" & Trim(cboWO)
    End If
    
'   str = "select Work_Order,Seq_NO from QSMS_Wogroup where GroupID='" & Trim(CboGroupID) & "' andWork_Order='" & Trim(CboWo) & "'"
'   Set rs = Conn.Execute(str)
'   If Not rs.EOF Then
'      Seq = rs!Seq_No
'      str = "delete from QSMS_WOGroup where GroupID='" & Trim(CboGroupID) & "' and Work_Order='" & Trim(CboWo)& "'"
'      Conn.Execute str
'      str = "Update QSMS_WoGroup set Seq_NO=Seq_No-1 where GroupID='" & Trim(CboGroupID) & "' and Seq_NO> " &Seq& ""
'      Conn.Execute str
'      MsgBox "Delete the WO OK"
'   Else
'      MsgBox "can not find the Wo :" & Trim(CboWo)
'   End If
   
End If
End Sub

Private Sub cmdReset_Click()


    
lstWO_LIST.Clear
lstWO_SELECT.Clear
CboGroupID.Clear
End Sub

Private Sub CmdQuery_Click()
If Trim(CboLine) = "" Then
   MsgBox "Please input line"
   Exit Sub
End If

strLine = CboLine.text
Call ListWO
End Sub

Private Sub CmdQueryID_Click()
Call GetGroupIDByDate
End Sub

Private Sub cmdSave_Click()
Dim i As Long
Dim str As String
Dim Rs As ADODB.Recordset
Dim tempwo As String
Dim TempGroupID As String
Dim TempGroupDatetime, TempDateTime As String
Dim MBFlag As String
Dim WOList As String
Dim sMsg As String
Dim Response As String

If lstWO_SELECT.ListCount <= 0 Then Exit Sub

WOList = ""
'###########(1) Check the line match
For i = 0 To lstWO_SELECT.ListCount - 1
    lstWO_SELECT.ListIndex = i
    tempwo = lstWO_SELECT.text
    
    WOList = WOList & tempwo & ","
    If DblChkLine(tempwo, CboLine) = False Then
       MsgBox "Line doesn't match the wo,Please check"
       Exit Sub
    
    End If
Next i

'###########(3) Generate Group ID
TempGroupID = GenGroupID

'###########(4)Insert Group ID into QSMS_WOGroup
str = "select getdate()"
Set Rs = Conn.Execute(str)
TempGroupDatetime = Format(Rs(0), "YYYYMMDDHHMMSS")

    str = "Exec CHKMaintainWO " & sq(WOList) & ",'" & (CboLine) & "'," & sq(TempGroupID) '''1239
    Set Rs = Conn.Execute(str)
    If Rs("Result") = "Fail" Then
        MsgBox (Rs("Description")), vbInformation
        Exit Sub
    End If
    

''ESBU               **Denver       2009.08.04      Add upload PNGroup and check PNGroup when create WO group  (0058)
If CheckPNGroup = "Y" Then
    WOList = Mid(WOList, 1, Len(WOList) - 1)
    
    str = "exec ChkPNGroup " & sq(WOList)
    Set Rs = Conn.Execute(str)
    If Rs("Result") <> 0 Then
        sMsg = Trim(Rs("Description") & "") & Chr(13)
        Set Rs = Rs.NextRecordset
        Do While Rs.EOF = False
            sMsg = sMsg & Chr(13) & Trim(Rs("ErrInfo") & "")
            Rs.MoveNext
        Loop
        MsgBox sMsg
        Exit Sub
    End If
End If


For i = 0 To lstWO_SELECT.ListCount - 1
    lstWO_SELECT.ListIndex = i
    tempwo = lstWO_SELECT.text
    
'*****************************************************(0073)********************************
    str = "Exec CheckWOGroupID " & sq(tempwo) & "," & sq(TempGroupID)
    Set Rs = Conn.Execute(str)
    If UCase(Rs("Item")) = "N" Then
        MsgBox "Other work order which is in the same PCB has already in the system,GroupID is: " & Rs("GroupID"), vbInformation
        Exit Sub
    End If
    
'******************************************************************************************

    If ChkWOGroupID = "Y" Then  ''(1128)
        str = "Exec QSMS_ChkWOGroupID " & sq(tempwo) & "," & sq(TempGroupID)
        Set Rs = Conn.Execute(str)
        If Rs("Result") = "1" Then
            Response = MsgBox(Rs("Err") & Chr(13) & " Do you want to continue ? ", vbYesNo)
            If Response = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    Call GetWoinfo(tempwo)
    str = "select Work_Order ,GroupID from QSMS_WoGroup where Work_Order='" & tempwo & "' union all select Work_Order ,GroupID from QSMS_History.dbo.qsms_wogroup where Work_Order='" & tempwo & "'"    ''1139
    Set Rs = Conn.Execute(str)
    If Rs.EOF Then
       If ChkMBWo(Trim(tempwo)) = True Then
          MBFlag = 1   'means the work order is MB
       Else
          MBFlag = 0   'means the work order is small board
       End If
       str = "insert into QSMS_WoGroup(Work_Order,MBPN,MBFlag,Line,Seq_No,GroupID,WO_TransDateTime,Group_TransDateTime,Sap1Flag,ClosedFlag,ClosedType,UID)" & _
             " values ('" & Trim(tempwo) & "','" & TxtMBPN & "','" & MBFlag & "','" & strLine & "' " & _
         "," & i + 1 & ",'" & TempGroupID & "','" & Wo_TransDate & "','" & TempGroupDatetime & "','N','N','','" & g_userName & "')"

        Conn.Execute (str)

''''''''''''(0007)
        str = "EXEC XL_CheckWOGroupID '" & Trim(tempwo) & "','" & Trim(TempGroupID) & "'"
        If Rs.State = 1 Then Rs.Close
        Set Rs = Conn.Execute(str)
        If UCase(Rs!Item) = "N" Then
            MsgBox "SP:XL_CheckWOGroupID Warnning"
            Exit Sub
        End If
    
    Else
        MsgBox "The Work Order already In DB, GroupID is: " & Rs!GroupID
        
    End If
Next i
lstWO_SELECT.Clear
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdUpdate_Click()
Dim i, Seq_No As Long
Dim tempwo As String
Dim Rs As New ADODB.Recordset
Dim str As String
Dim TempDateTime, TempGroupDatetime As String, DTLimit As String
Dim Response As String

str = "select getdate()"
Set Rs = Conn.Execute(str)
TempGroupDatetime = Format(Rs(0), "YYYYMMDDHHMMSS")
str = "Select Max(Seq_NO) from QSMS_WoGroup where GroupID='" & Trim(CboGroupID) & "'"
Set Rs = Conn.Execute(str)
Seq_No = Rs.Fields(0)
'''''''''''''''''''''''''''''''''''''begin (0005)
str = "EXEC QSMS_CheckGroupID '" & Trim(CboGroupID) & "'"
If Rs.State = 1 Then Rs.Close
Set Rs = Conn.Execute(str)
If UCase(Rs!Item) = "N" Then
    MsgBox "The GroupID is over " & Rs!DTLimit & " week,can not insert"
    Exit Sub
End If

DTLimit = Rs!DTLimit


'''''''''''''''''''''''''''''''''''''end (0005)
For i = 0 To lstWO_SELECT.ListCount - 1
    lstWO_SELECT.ListIndex = i
    tempwo = lstWO_SELECT.text
'*****************************************************(0073)********************************
    str = "Exec CheckWOGroupID " & sq(tempwo) & "," & sq(CboGroupID)
    Set Rs = Conn.Execute(str)
    If UCase(Rs("Item")) = "N" Then
        MsgBox "Other work order which is in the same PCB has already in the system,GroupID is: " & Rs("GroupID"), vbInformation
        Exit Sub
    End If
'******************************************************************************************

'*****************************************************(1283)********************************
    str = "Exec QSMS_CheckAssignedComp " & sq(tempwo) & "," & sq(CboGroupID)
    Set Rs = Conn.Execute(str)
    If Rs("Result") = "Fail" Then
        MsgBox (Rs("Description")), vbInformation
        Exit Sub
    End If
'******************************************************************************************

    str = "Exec CHKMaintainWO " & sq(tempwo) & ",'" & (CboLine) & "'," & sq(CboGroupID) '''1239
    Set Rs = Conn.Execute(str)
    If Rs("Result") = "Fail" Then
        MsgBox (Rs("Description")), vbInformation
        Exit Sub
    End If
    
    If ChkWOGroupID = "Y" Then  ''(1128)
        str = "Exec QSMS_ChkWOGroupID " & sq(tempwo) & "," & sq(CboGroupID)
        Set Rs = Conn.Execute(str)
        If Rs("Result") = "1" Then
            Response = MsgBox(Rs("Err") & Chr(13) & " Do you want to continue ? ", vbYesNo)
            If Response = vbNo Then
                Exit Sub
            End If
        End If
    End If

    Call GetWoinfo(tempwo)
    Seq_No = Seq_No + 1
    str = "select Work_Order ,GroupID from QSMS_WoGroup where Work_Order='" & tempwo & "'"
    If Rs.State = 1 Then Rs.Close
    Set Rs = Conn.Execute(str)
    If Not Rs.EOF Then
       str = "Update QSMS_WoGroup set Seq_NO='" & Seq_No & "' where Work_Order='" & tempwo & "'"
       Conn.Execute str
    Else
        str = "insert into QSMS_WoGroup(Work_Order,Line,Seq_No,GroupID,WO_TransDateTime,Group_TransDateTime,Sap1Flag,ClosedFlag,ClosedType,UID) values ('" & Trim(tempwo) & "','" & Trim(CboLine) & "' " & _
            "," & Seq_No & ",'" & CboGroupID & "','" & Wo_TransDate & "','" & TempGroupDatetime & "','N','N','','" & g_userName & "')"
        Conn.Execute str
    End If
''''''''''''(0007)
    str = "EXEC XL_CheckWOGroupID '" & Trim(tempwo) & "','" & Trim(CboGroupID) & "'"
    If Rs.State = 1 Then Rs.Close
    Set Rs = Conn.Execute(str)
    If UCase(Rs!Item) = "N" Then
        MsgBox "Run SP Fail:XL_CheckWOGroupID"
        Exit Sub
    End If
       
Next i

str = "select top 1 Group_TransDateTime as BeginDateTime, dbo.FormatDate(DATEADD(wk," & DTLimit & ",dbo.Format_To_Date(Group_TransDateTime)),'YYYYMMDDHHNNSS') as EndDateTime from QSMS_WoGroup where GroupID = '" & Trim(CboGroupID) & "' order by Group_TransDateTime"
If Rs.State = 1 Then Rs.Close
Set Rs = Conn.Execute(str)

MsgBox "Update OK" & vbCrLf & vbCrLf & "BeginDateTime: " & Rs!begindatetime & " ~ EndDateTime: " & Rs!EndDateTime & vbCrLf & "If the WOGroup is about to expire, please prepare to create a new WOGroup."

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
Dim str As String
Dim Rs As ADODB.Recordset
Dim i As Long
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
str = "select getdate()"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
    'Date = Rs(0)
    'Time = Rs(0)
End If
dtpSDate = Date
dtpEDate = Date
Call GetLine
End Sub
Private Function GetLine()
Dim str As String
Dim Rs As ADODB.Recordset
''0072
str = "select distinct Line from Machine order by Line"
Set Rs = Conn.Execute(str)
CboLine.Clear
While Not Rs.EOF
    CboLine.AddItem Rs!Line
    Rs.MoveNext
Wend
End Function

Private Sub lstWO_LIST_Click()

Call GetWoinfo(lstWO_LIST.text)
End Sub
Private Function GetWoinfo(ByVal WO As String)
Dim str As String
Dim Rs As ADODB.Recordset

str = "select Line,PN ,Qty ,Trans_date from SAP_WO_list Where WO='" & WO & "' "
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   TxtMBPN = Trim(Rs!PN)
   TxtWOQty = Trim(Rs!Qty)
   Wo_TransDate = Left(Trim(Rs!Trans_Date), 8)
End If
End Function


Private Function ChkDelete(ByVal WO As String) As Boolean
'(1)if WO has dispatched can not delete from the group
'(2)if wo  dispatching can not delete from the group
Dim str As String
Dim Rs As ADODB.Recordset
Dim ForecastDate As Date
ChkDelete = True
str = "select * from QSMS_WoGroup where Work_Order='" & WO & "'"
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
   If UCase(Trim(Rs!DispatchFlag)) = "Y" Then
      ChkDelete = False
      MsgBox "can not delete!!!!The work Order has been dispatched:" & WO
   End If
   

End If
str = "Select Count(*) From QSMS_Dispatch where Work_Order='" & WO & "'"
Set Rs = Conn.Execute(str)
If Rs.Fields(0) > 1 Then
   ChkDelete = False
   MsgBox "can not delete!!!!! The word order is dispatching:" & WO
End If
End Function


