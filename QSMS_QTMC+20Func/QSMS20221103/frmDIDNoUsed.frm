VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDIDNoUsed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DIDNoUsed"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   12240
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dgInfo 
      Height          =   5655
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   9975
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
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.TextBox txtTime2 
         Height          =   405
         Left            =   6720
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtTime1 
         Height          =   405
         Left            =   3480
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   7920
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "Query"
         Height          =   375
         Left            =   6240
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cbbDateRange 
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
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cbbLine 
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
         Left            =   960
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Greater than or equal to£º"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Less than£º"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DateRange:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2160
         TabIndex        =   3
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Line:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmDIDNoUsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset

Private Sub CmdExcel_Click()
If rs.State = 1 Then
    If rs.EOF Then
        MsgBox ("NO DATA !")
    Else
        Call CopyToExcel(rs)
    End If
End If
End Sub

Private Sub CmdQuery_Click()
Dim strSQL As String
Dim DateType As String

If Trim(cbbLine.Text) = "" Then
    MsgBox ("Please input Line !"), vbCritical
    Exit Sub
End If

If Trim(cbbDateRange.Text) = "" And txtTime1.Text = "" And txtTime2.Text = "" Then
    MsgBox ("Please input DateRange !"), vbCritical
    Exit Sub
End If
If BU = "NB5" And cbbDateRange.Text = "" And (txtTime1.Text <> "" And txtTime2.Text <> "") Then
    DateType = "4"
    strSQL = "Exec Query_NOUseDID '" & DateType & "','" & Trim(cbbLine.Text) & "','" & Trim(txtTime1.Text) & "','" & Trim(txtTime2.Text) & "'"
Else
    If cbbDateRange.Text = ">=3 and <5" Then
        DateType = "1"
    Else
        If cbbDateRange.Text = ">=5 and <10" Then
            DateType = "2"
        ElseIf (cbbDateRange.Text = ">=5 and <10") Then
            DateType = "3"
        End If
    End If
    strSQL = "Exec Query_NOUseDID '" & DateType & "','" & Trim(cbbLine.Text) & "'"
End If
'strSQL = "Exec Query_NOUseDID '" & DateType & "','" & Trim(cbbLine.Text) & "'"
Set rs = Conn.Execute(strSQL)

If rs.EOF Then
    Set dgInfo.DataSource = Nothing
    MsgBox ("NO DATA !")
Else
    Set dgInfo.DataSource = rs
End If

End Sub

Private Sub Form_Load()
Dim strSQL As String
Dim tmpRS As New ADODB.Recordset

strSQL = "select distinct line from QSMS_Dispatch"
Set tmpRS = Conn.Execute(strSQL)

cbbLine.Clear
While tmpRS.EOF = False
    cbbLine.AddItem (Trim(tmpRS("Line")))
    tmpRS.MoveNext
Wend

cbbDateRange.Clear
cbbDateRange.AddItem ">=3 and <5"
cbbDateRange.AddItem ">=5 and <10"
cbbDateRange.AddItem ">=10"
If BU = "NB5" Then
    txtTime1.Visible = True
    txtTime2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    
End If
End Sub

