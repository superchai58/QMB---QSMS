VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTraceReport 
   BackColor       =   &H00E0E0E0&
   Caption         =   "TraceReport（20160614）"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14325
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmComPNbySN"
   ScaleHeight     =   9510
   ScaleWidth      =   14325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "DataFormat"
      Height          =   615
      Left            =   9000
      TabIndex        =   37
      Top             =   240
      Width           =   2415
      Begin VB.OptionButton opttxt 
         Caption         =   "Txt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1320
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optExcel 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame FraSN 
      BackColor       =   &H00808080&
      Caption         =   "FraSN"
      Height          =   8535
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   3975
      Begin MSDataGridLib.DataGrid DataGridSN 
         Height          =   6495
         Left            =   120
         TabIndex        =   34
         Top             =   1920
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   11456
         _Version        =   393216
         AllowUpdate     =   0   'False
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
      Begin VB.Frame FraSNBy 
         BackColor       =   &H0000FF00&
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   3735
         Begin VB.OptionButton OptBatch 
            BackColor       =   &H0000FF00&
            Caption         =   "By Batch"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   32
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton OptSN 
            BackColor       =   &H0000FF00&
            Caption         =   "By One"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   31
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtSN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   29
         Top             =   1440
         Width           =   3735
      End
      Begin VB.Label lblSNWO 
         BackColor       =   &H0000FF00&
         Caption         =   "SN/DID/WO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   3735
      End
   End
   Begin VB.Frame FraCompN 
      BackColor       =   &H00808080&
      Caption         =   "FraCompN"
      Height          =   3495
      Left            =   4320
      TabIndex        =   11
      Top             =   960
      Width           =   9495
      Begin VB.TextBox txtModel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   16
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtLotCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   15
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox txtDateCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   14
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtVendorCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox TxtCompPN 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1680
         TabIndex        =   12
         Top             =   360
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker DTPBeginTime 
         Height          =   495
         Left            =   4560
         TabIndex        =   17
         Top             =   2040
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   112918531
         UpDown          =   -1  'True
         CurrentDate     =   37678
      End
      Begin MSComCtl2.DTPicker DTPBeginDate 
         Height          =   495
         Left            =   1680
         TabIndex        =   18
         Top             =   2040
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   873
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
         Format          =   112918531
         CurrentDate     =   36482
      End
      Begin MSComCtl2.DTPicker DTPEndTime 
         Height          =   495
         Left            =   4560
         TabIndex        =   19
         Top             =   2640
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm:ss"
         Format          =   112918531
         UpDown          =   -1  'True
         CurrentDate     =   37678
      End
      Begin MSComCtl2.DTPicker DTPEndDate 
         Height          =   495
         Left            =   1680
         TabIndex        =   20
         Top             =   2640
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   873
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
         Format          =   112918531
         CurrentDate     =   36482
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Begin Date"
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
         Height          =   495
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label2 
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
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "CompPN"
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
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "VendorCode"
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
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Date Code"
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
         Left            =   4680
         TabIndex        =   23
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Lot Code"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         Caption         =   "Model"
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
         Left            =   4680
         TabIndex        =   21
         Top             =   1440
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdGetData 
      Caption         =   "GetData"
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
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "导入数据"
      Height          =   1935
      Left            =   4320
      TabIndex        =   2
      Top             =   6360
      Width           =   9615
      Begin VB.TextBox Txtpath 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   5295
      End
      Begin VB.CommandButton inputSN 
         BackColor       =   &H80000000&
         Caption         =   "导入数据"
         Height          =   375
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton CMDChosefile 
         BackColor       =   &H80000000&
         Caption         =   "选择文件"
         Height          =   375
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "DataFormat"
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
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   930
         Left            =   1560
         Picture         =   "FrmTraceReport.frx":0000
         Top             =   960
         Width           =   7695
      End
      Begin VB.Label Label1 
         Caption         =   "File Path\\"
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
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.ComboBox CbbDataType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6000
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   13440
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "如果选用Txt格式: 数据结果本地存放路径:C:\TraceReportData\         数据结果服务器共享路径:QSMS Server D:\TraceReportData\"
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
      Left            =   240
      TabIndex        =   36
      Top             =   8760
      Width           =   13695
   End
   Begin VB.Label labelInfor 
      BackColor       =   &H00C0C0FF&
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
      Left            =   12840
      TabIndex        =   10
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblInfor 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4320
      TabIndex        =   9
      Top             =   4560
      Width           =   9495
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Begin Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Index           =   2
      Left            =   6000
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "DataType"
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
      Left            =   4440
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmTraceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DB As String
Dim BatchOne As String
Dim Mode As String
Dim strTable As String
Dim SN As String
Dim strSQL As String
Dim RS As New ADODB.Recordset
Dim xlApp As New Excel.Application
Dim xlBook As New Excel.Workbook
Dim xlSheet As New Excel.Worksheet
Dim chkAutoPass As String

Private Sub cmdGetData_Click()
Dim str As String, strSQL As String, strSDateTime As String, strEDateTime As String, strDestPath As String, strSourceFile As String, strDeleteFile As String
Dim RS As New ADODB.Recordset, resultType As String

If Trim(CbbDataType) = "" Then
    MsgBox "Please select the Data Type!"
    Exit Sub
End If
    
    If optExcel.Value = True Then
        resultType = "EXCEL"
    Else
        resultType = "TEXT"
    End If
    

    labelInfor.Caption = ""
    lblInfor.Caption = ""
    strSDateTime = Trim(Format(DTPBeginDate & " " & DTPBeginTime.Value, "YYYYMMDDHHNNSS"))
    strEDateTime = Trim(Format(DTPEndDate & " " & DTPEndTime.Value, "YYYYMMDDHHNNSS"))
    strDestPath = "C:\TraceReportData"
    
    If Dir(strDestPath, vbDirectory) = "" Then
        CreateFolder (strDestPath)
    End If
    
    If UCase(Trim(CbbDataType.Text)) = "TRACE BY COMPPN" Then
        If strSDateTime > strEDateTime Then
            MsgBox "Please input right start date time and end date time!"
            Exit Sub
        End If
        If Trim(TxtCompPN) = "" Then
            MsgBox "Please input the CompPN!"
            TxtCompPN.SetFocus
            Exit Sub
        End If
        If DateDiff("d", DTPBeginDate.Value, DTPEndDate.Value) > 31 Then
            MsgBox "系统一次查询的时间跨度不能超过31天!请分多次查询."
            Exit Sub
        End If
    
        str = "Exec TraceReport_GetSNByComp '" & Trim(TxtCompPN) & "','" & Trim(txtVendorCode) & "' ,'" & Trim(txtDateCode) & "','" & Trim(txtLotCode) & "','" & strSDateTime & "','" & strEDateTime & "','" & Trim(txtModel) & "','" & resultType & "'"
    Else
        If OptSN.Value = True Then
            If Trim(TxtSN) = "" Then
                MsgBox "Please input the SN or DID or WO!"
                Exit Sub
            End If
            
            If UCase(Trim(CbbDataType.Text)) = "TRACE BY SN" Then
                str = "EXEC  TraceReport_GetCompBySN 'one', '" & Trim(TxtSN) & "','" & Trim(TxtCompPN) & "','" & resultType & "'"
            End If
            If UCase(Trim(CbbDataType.Text)) = "TRACE BY DID" Then
                str = "EXEC TraceReport_GetSNByDID 'one','" & Trim(TxtSN) & "','" & resultType & "'"
            End If
            If UCase(Trim(CbbDataType.Text)) = "TRACE BY WO" Then
                str = "EXEC TraceReport_GetCompByWO '" & Trim(TxtSN) & "','" & Trim(TxtCompPN) & "','" & resultType & "'"
            End If
        Else
            If UCase(Trim(CbbDataType.Text)) = "TRACE BY SN" Then
                str = "EXEC  TraceReport_GetCompBySN 'Batch', '','','" & resultType & "'"
            End If
            If UCase(Trim(CbbDataType.Text)) = "TRACE BY DID" Then
                str = "EXEC TraceReport_GetSNByDID 'Batch','','" & resultType & "'"
            End If
            If UCase(Trim(CbbDataType.Text)) = "TRACE BY WO" Then
                 str = "EXEC TraceReport_GetCompByWO_Batch  '','" & resultType & "'"   ''1206
            End If
        End If
    End If
    RS.CursorLocation = adUseClient
    Set RS = Conn.Execute(str)
    
    If UCase(resultType) = "EXCEL" Then
        Call CopyToExcel(RS)
    Else
        strDestPath = strDestPath & "\" & Trim(RS.Fields("FileName"))
        lblInfor.Caption = "从服务器:" & Trim(RS.Fields("FilePath")) & " 复制查询的 TraceReport 数据到本地目录:" & strDestPath & ",请核对!"
        strSourceFile = Trim(RS.Fields("FilePath"))
        FileCopy strSourceFile, strDestPath
        
        MsgBox "从服务器:" & Trim(RS.Fields("FilePath")) & " 复制查询的 TraceReport 数据到本地目录:" & strDestPath & "完成,请核对!", vbOKOnly, "Message"
        
        strDeleteFile = "D:\TraceReportData\" & Trim(RS.Fields("FileName"))
        str = "Exec TraceReport_DeleteFile '" & strDeleteFile & "'"
        Conn.Execute (str)
    End If
    labelInfor.Caption = "Data create OK!"
End Sub

Private Sub DataGridSN_SelChange(Cancel As Integer)
'    If DataGridSN.Columns.Count > 1 Then  ''20101111 update by kaitlyn for Bug
    If DataGridSN.ApproxCount > 1 Then
        TxtSN.Text = Trim(DataGridSN.Columns("SN").Value)
        TxtCompPN.Text = Trim(DataGridSN.Columns("CompPN").Value)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub CbbDataType_click()
      
      If Trim(CbbDataType.Text) = "Trace By SN" Then
          lblSNWO = "SN"
      ElseIf Trim(CbbDataType.Text) = "Trace By WO" Then
          lblSNWO = "WO"
      ElseIf Trim(CbbDataType.Text) = "Trace By DID" Then
          lblSNWO = "DID"
      End If
  
End Sub

Private Sub Form_Load()

    With CbbDataType
        .AddItem "Trace By SN"
        .AddItem "Trace By WO"
        .AddItem "Trace By DID"
        .AddItem "Trace By CompPN"
    End With
    DTPBeginDate = Date
    DTPEndDate = Date
    DTPBeginTime = Time
    DTPEndTime = Time
        
    Call FreshSN
    
    ''''Kyle    20110111
    StrBU = ReadIniFile("COMMON", "BU", App.Path & "\set.ini")
    If StrBU = "PO" Then
        Call LoadTradChinese
    End If
End Sub


Private Sub OptSN_Click()
     
    If Trim(CbbDataType.Text) = "Trace By SN" Then
       lblSNWO = "SN"
    End If
    If Trim(CbbDataType.Text) = "Trace By DID" Then
       lblSNWO = "DID"
    End If
End Sub
Private Sub inputSN_Click()
   If Trim(Txtpath.Text) = "" Then
      MsgBox ("please check it ")
        Exit Sub
    End If
Call UpLoadfile
End Sub


Private Sub UpLoadfile()
Dim strSQL As String
Dim Row_ID As Long
Dim tempSN As String, tempCompPN As String
Dim I As Long
Dim RS As New ADODB.Recordset
On Error GoTo ErrH

   Row_ID = 2
   xlApp.Workbooks.Open (Txtpath)

   strSQL = "delete TraceReport_TempSN where HostName=host_name()"  ''(1231)
   Conn.Execute (strSQL)
   
   tempCompPN = ""
   
   If xlApp.Cells(Row_ID, 1) = "" Then
        MsgBox ("wrong")
        xlApp.Workbooks.Close
        Set xlBook = Nothing
        Set xlSheet = Nothing
        Set xlApp = Nothing
    Exit Sub
    End If

    With xlApp
         Do While .Cells(Row_ID, 1) <> ""
            tempSN = .Cells(Row_ID, 1)
            tempCompPN = Trim(.Cells(Row_ID, 2))
            strSQL = "insert into TraceReport_TempSN(SN,CompPN,HostName) values ('" & Trim(tempSN) & "','" & Trim(tempCompPN) & "',host_name())"    ''(1231)
            strSQL = Replace(strSQL, Chr(10), "")
            strSQL = Replace(strSQL, Chr(13), "")
            Conn.Execute (strSQL)
            Row_ID = Row_ID + 1
            lblInfor.Caption = "Load Data count:" & Trim(Row_ID)
         Loop
         
    End With
    
    xlApp.Workbooks.Close
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Set xlApp = Nothing

    Call FreshSN
ErrH:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "ErrMeg"
    End If
End Sub

Private Sub FreshSN()
Dim RS As New ADODB.Recordset
Dim strSQL As String
On Error GoTo ErrH
    strSQL = "select distinct SN,CompPN from TraceReport_TempSN where HostName=host_name() order by SN" ''(1231)
    If RS.State Then RS.Close
    Set RS = Conn.Execute(strSQL)
    Set DataGridSN.DataSource = RS
    
ErrH:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "ErrMeg"
    End If
End Sub

Private Sub CMDChosefile_Click()
    dlg.Filter = "*.xls|*.xlsx"
    dlg.ShowOpen
    Txtpath.Text = dlg.FileName
End Sub

Private Sub LoadTradChinese()
    Frame1.Caption = "蹲戈"
    CMDChosefile.Caption = "匡拒"
    inputSN.Caption = "旧"
    Label6.Caption = "狦匡ノtxtΑ: 玻ネ郎セ隔畖C:\TraceReportData\      狝竟ㄉ隔畖QSMS Server D:\TraceReportData\"
End Sub
