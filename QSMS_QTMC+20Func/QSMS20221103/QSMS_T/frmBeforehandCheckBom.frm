VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBeforehandCheckBom 
   Caption         =   "Beforehand_CheckBom"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   12915
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox comFac 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2280
      TabIndex        =   10
      Top             =   960
      Width           =   3975
   End
   Begin VB.CommandButton ComExcel 
      Caption         =   "Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   8
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton ComCKB 
      Caption         =   "CheckBom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   7
      Top             =   360
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DGResult 
      Height          =   5775
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   10186
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
   Begin VB.TextBox txtCQty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.ComboBox CombLine 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8040
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.ComboBox CombModel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2280
      TabIndex        =   1
      Text            =   "CombModel"
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Factory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CombineQty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Line"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ModelName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmBeforehandCheckBom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset, Str As String
Private Sub ComCKB_Click()
Dim PN As String, Rev As String

 If Trim(CombLine) = "" Then
    MsgBox "Please select the Line at first!", vbCritical, "ErrMessage"
    CombLine.SetFocus
    Exit Sub
 End If
 If Trim(CombModel) = "" Then
    MsgBox "Please select the Model at first!", vbCritical, "ErrMessage"
    CombModel.SetFocus
    Exit Sub
 End If
 If Trim(comFac) = "" Then
    MsgBox "Please select the Factory at first!", vbCritical, "ErrMessage"
    comFac.SetFocus
    Exit Sub
 End If
 If Trim(txtCQty) = "" Then
    MsgBox "Please input the combine qty at first!", vbCritical, "ErrMessage"
    txtCQty.SetFocus
    Exit Sub
 End If
 If IsNumeric(txtCQty) = False Then
    MsgBox "Please input the qty!", vbCritical, "ErrMessage"
    txtCQty.SetFocus
    Exit Sub
 End If
 PN = Left(Trim(CombModel), InStr(Trim(CombModel), "-") - 1)
 Rev = Right(Trim(CombModel), Len(Trim(CombModel)) - Len(PN) - 1)
 
 Str = "select * from Sap_Wo_List where WO='VIRTUALWO' and Trans_Date>dbo.formatdate(dateadd(N,-8,getdate()),'YYYYMMDDHHNNSS')"
 Set rs = Conn.Execute(Str)
 If rs.EOF = False Then
    MsgBox "Somebody is checking Bom,please CheckBom after some minutes,thanks!"
    Exit Sub
 End If
 
 Str = "exec QSMS_BeforehandCheckBom '" & Trim(comFac) & "','" & Trim(CombLine) & "','" & PN & "','" & Rev & "','" & txtCQty & "'"
 Set rs = Conn.Execute(Str)
 
 If rs.EOF Then
    Set DGResult.DataSource = rs
    MsgBox "CheckBom OK!", vbOKOnly, "Message"
 Else
    Str = "select * from Sap_Bom_Fail where Work_Order='VIRTUALWO'"
    Set rs = Conn.Execute(Str)
    Set DGResult.DataSource = rs
    MsgBox "CheckBom Fail!", vbCritical, "Message"
 End If
End Sub

Private Sub ComExcel_Click()
If Trim(DGResult.Columns(0)) <> "" Then
    Call CopyToExcel(rs)
End If
End Sub

Private Sub Form_Load()
    Call GetLine
    Call GetModel
    With comFac
        .AddItem "F1"
        .AddItem "F2"
        .AddItem "F4"
        .AddItem "F6"
        .AddItem "F7"
        .AddItem "QB"
        .AddItem "QC"
    End With
End Sub

Private Sub GetLine()
Str = "select distinct Line from QSMS_woGroup order by line"
Set rs = Conn.Execute(Str)
CombLine.Clear
While Not rs.EOF
    CombLine.AddItem rs!Line
    rs.MoveNext
Wend
End Sub

Private Sub GetModel()

Str = "select distinct MBPN+'-'+MBRev as Model from SAPBOM order by Model"
Set rs = Conn.Execute(Str)
CombModel.Clear
While Not rs.EOF
    CombModel.AddItem rs.Fields("Model")
    rs.MoveNext
Wend
End Sub
