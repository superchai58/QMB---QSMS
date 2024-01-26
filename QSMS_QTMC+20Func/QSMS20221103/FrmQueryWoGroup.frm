VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmQueryWoGroup 
   Caption         =   "Frm Query WOGroup"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   14685
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtWO 
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton CmdClosed 
      Caption         =   "&Excel Close"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Excel_UnCLose"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ComboBox CboLine 
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DGNotFinished 
      Height          =   3135
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   5530
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
      Caption         =   "Un closed Wo List"
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
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpSDate 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   720
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
      Format          =   69926915
      CurrentDate     =   36482
   End
   Begin MSDataGridLib.DataGrid DGFinish 
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   6376
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
      Caption         =   "Closed WO List"
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
   Begin MSComCtl2.DTPicker dtpEDate 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1200
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
      Format          =   69926915
      CurrentDate     =   36482
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "Work_Order"
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
      TabIndex        =   11
      Top             =   1680
      Width           =   1455
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
      TabIndex        =   8
      Top             =   1200
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
      Left            =   120
      TabIndex        =   3
      Top             =   720
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
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "FrmQueryWoGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsOK As ADODB.Recordset
Dim RsNotOK As ADODB.Recordset

Private Sub CmdClosed_Click()

 If Not RsOK.EOF Then
       Call CopyToExcel(RsOK)
    Else
       MsgBox ("No Data"), vbCritical
End If

End Sub

Private Sub cmdExcel_Click()
 If Not RsNotOK.EOF Then
       Call CopyToExcel(RsNotOK)
    Else
       MsgBox ("No Data"), vbCritical
End If
End Sub

Private Sub CmdQuery_Click()
Dim Str As String
Dim BeginDate As String
Dim EndDate As String
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
BeginDate = Format(dtpSDate, "YYYY/MM/DD")
BeginDate = Replace(BeginDate, "-", "")
BeginDate = Replace(BeginDate, "/", "")
EndDate = Format(dtpEDate, "YYYY/MM/DD")
EndDate = Replace(EndDate, "-", "")
EndDate = Replace(EndDate, "/", "")
If Len(Trim(TxtWO)) = 9 Then
       Str = "select a.GroupID,a.Seq_NO,a.Work_Order,a.Line,b.PN,B.MB_Rev,B.Qty,a.Wo_TransDateTime,a.Group_TransDateTime,a.DispatchFlag,a.Sap1Flag,a.ClosedFlag,a.ClosedType,a.UID,a.CloseDateTime " & _
           "from qsms_WOGroup a,Sap_Wo_List b  where a.work_order=b.wo and a.groupID in (select Groupid from qsms_woGroup where work_order='" & Trim(TxtWO) & "') " & _
           "and ClosedFlag='N' and a.line like '" & Trim(CboLine) & "%' order by groupID"
        Set RsNotOK = Conn.Execute(Str)
        Set DGNotFinished.DataSource = RsNotOK
        
        Str = "select a.GroupID, a.Seq_NO,a.Work_Order,a.Line,b.PN,B.MB_Rev,B.Qty,a.Wo_TransDateTime,a.Group_TransDateTime,a.DispatchFlag,a.Sap1Flag,a.ClosedFlag,a.ClosedType,a.UID,a.CloseDateTime " & _
               "from qsms_WOGroup a,Sap_Wo_List b  where a.work_order=b.wo and a.groupID in (select Groupid from qsms_woGroup where work_order='" & Trim(TxtWO) & "') " & _
               "and ClosedFlag='Y' and a.line like '" & Trim(CboLine) & "%' order by GroupID"
        
        Set RsOK = Conn.Execute(Str)
        Set DGFinish.DataSource = RsOK
Else

        Str = "select a.GroupID,a.Seq_NO,a.Work_Order,a.Line,b.PN,B.MB_Rev,B.Qty,a.Wo_TransDateTime,a.Group_TransDateTime,a.DispatchFlag,a.Sap1Flag,a.ClosedFlag,a.ClosedType,a.UID,a.CloseDateTime " & _
               "from qsms_WOGroup a,Sap_Wo_List b  where a.Work_Order=b.Wo and " & _
                "substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' and ClosedFlag='N' and a.line like '" & Trim(CboLine) & "%' order by groupID"
        Set RsNotOK = Conn.Execute(Str)
        Set DGNotFinished.DataSource = RsNotOK
        
        Str = "select a.GroupID, a.Seq_NO,a.Work_Order,a.Line,b.PN,B.MB_Rev,B.Qty,a.Wo_TransDateTime,a.Group_TransDateTime,a.DispatchFlag,a.Sap1Flag,a.ClosedFlag,a.ClosedType,a.UID,a.CloseDateTime " & _
               "from qsms_WOGroup a,Sap_Wo_List b  where a.Work_Order=b.Wo and " & _
               "substring(Group_TransDateTime,1,8) between '" & BeginDate & "' and '" & EndDate & "' and ClosedFlag='Y' and a.line like '" & Trim(CboLine) & "%' order by GroupID"
        
        Set RsOK = Conn.Execute(Str)
        Set DGFinish.DataSource = RsOK
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim Str As String
Dim Rs As ADODB.Recordset
Dim i As Long
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
CboLine.Clear
Str = "select distinct line from QSMS_WoGroup  order by line"
Set Rs = Conn.Execute(Str)
While Not Rs.EOF
    CboLine.AddItem Trim(Rs!Line)
    Rs.MoveNext
Wend
End Sub
