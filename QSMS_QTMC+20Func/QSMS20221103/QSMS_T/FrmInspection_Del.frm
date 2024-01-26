VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmInspection_Del 
   Caption         =   "Delete Inspection Result [20100421]"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6960
      TabIndex        =   25
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4680
      TabIndex        =   24
      Top             =   3360
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3735
      Left            =   240
      TabIndex        =   23
      Top             =   3840
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   6588
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   11175
      Begin VB.TextBox txtTransdatetime 
         Enabled         =   0   'False
         Height          =   375
         Left            =   8640
         TabIndex        =   22
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtResult 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4680
         TabIndex        =   20
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtValue 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtLower 
         Enabled         =   0   'False
         Height          =   375
         Left            =   9120
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtUpper 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6360
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtCompPN 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000FF00&
         Caption         =   "TransDateTime:"
         Height          =   375
         Left            =   6840
         TabIndex        =   21
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H0000FF00&
         Caption         =   "TestResult:"
         Height          =   375
         Left            =   3240
         TabIndex        =   19
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H0000FF00&
         Caption         =   "TestValue:"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H0000FF00&
         Caption         =   "Lower:"
         Height          =   375
         Left            =   8160
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H0000FF00&
         Caption         =   "Upper:"
         Height          =   375
         Left            =   5400
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "CompPN£º"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin VB.CommandButton cmdQuery 
         Caption         =   "Query"
         Height          =   855
         Left            =   9360
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtETime 
         Height          =   375
         Left            =   7800
         TabIndex        =   8
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Format          =   70254594
         CurrentDate     =   .999305555555556
      End
      Begin MSComCtl2.DTPicker dtEDate 
         Height          =   375
         Left            =   6240
         TabIndex        =   7
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   70254593
         CurrentDate     =   40289
      End
      Begin MSComCtl2.DTPicker dtSTime 
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         Format          =   70254594
         CurrentDate     =   40289
      End
      Begin MSComCtl2.DTPicker dtSDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   70254593
         CurrentDate     =   40289
      End
      Begin VB.TextBox txtDID 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   7095
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FF00&
         Caption         =   "EndDate:"
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FF00&
         Caption         =   "StartDate:"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "DID\CompPN:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmInspection_Del"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    If MsgBox("Are you sure to Delete?", vbYesNo) = vbYes Then
        Dim str As String
        Dim rs As New ADODB.Recordset
        
        If txtCompPN = "" Or txtTransdatetime = "" Then
            MsgBox "Please Select the Data you want to Delete"
            Exit Sub
        End If
        str = "delete QSMS_DID_InSpect where compPn='" & txtCompPN & "' and Transdatetime='" & txtTransdatetime & "'"
        Conn.Execute (str)
        
        MsgBox "Delete OK"
        Call cmdQuery_Click
    End If
End Sub

Private Sub cmdQuery_Click()
    Dim str As String
    Dim sDate As String
    Dim eDate As String
    Dim rs As New ADODB.Recordset
    
    Set DataGrid1.DataSource = Nothing
    sDate = Format(dtSDate, "YYYYMMDD") & Format(dtSTime, "HHNNSS")
    eDate = Format(dtEDate, "YYYYMMDD") & Format(dtETime, "HHNNSS")
    
    str = "select * from QSMS_DID_InSpect where (DID like '%" & txtDID & "%' or CompPN like '%" & txtDID & "%') and transDatetime between '" & sDate & "' and '" & eDate & "' order by TransDateTime desc"

    Set rs = Conn.Execute(str)
    If Not rs.EOF Then
        Set DataGrid1.DataSource = rs
    End If
End Sub
Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If IsNull(LastRow) Then
        Exit Sub
    End If
    With DataGrid1
        txtCompPN = .Columns(2).Value
        txtUpper = .Columns(3).Value
        txtLower = .Columns(4).Value
        txtResult = .Columns(13).Value
        txtValue = .Columns(12).Value
        txtTransdatetime = .Columns(15).Value
    End With
End Sub
