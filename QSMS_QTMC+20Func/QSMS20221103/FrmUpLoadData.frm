VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmUpLoadData 
   Caption         =   "Upload Data [20230919]"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraFile 
      BackColor       =   &H80000013&
      Caption         =   "File & Sheet selection"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.CommandButton cmdExcel 
         Caption         =   "&Excel"
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
         Height          =   735
         Left            =   4200
         Picture         =   "FrmUpLoadData.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton Cmd_Load 
         Caption         =   "Load Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         Picture         =   "FrmUpLoadData.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Txt_RowCount 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5280
         TabIndex        =   5
         Top             =   2640
         Width           =   1335
      End
      Begin VB.ComboBox cboSheetName 
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
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   5055
      End
      Begin VB.TextBox txtFilePath 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   5055
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "Select File"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox CboFuncType 
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
         ItemData        =   "FrmUpLoadData.frx":0614
         Left            =   1560
         List            =   "FrmUpLoadData.frx":0616
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   1200
         Width           =   5055
      End
      Begin VB.Label lblTemplete 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Template"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblFileFormat 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1560
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   1680
         Width           =   5055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Row Count"
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
         Left            =   5280
         TabIndex        =   10
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sheet Name"
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
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Func Type"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmUpLoadData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'EQMS_ID             **修 改 人     修改日期        描    述
'-------------------------------------------------------------------------------------------------
'                    **Jing         2007.10.30      Upload DID splicing qty threshold by component header (0004)
'                    **Jing         2007.11.19      Add a function to upload excel to WOScheduleList (0006)
'                    **Jing         2007.11.26      Upload PMC 2 shift of next 12 WO sequence (0007)
'                    **Jing         2007.12.05      Upload PN to XL_ImplementPN  (0008)
'                    **Jing         2007.12.17      Upload data to XL_WOPN   (0009)
'                    **Jing         2008.01.07      Remove the time limit of upload the WOPlanSeq   (0010)
'                    **Jing         2008.01.08      Ignore the upload data that the value of TotalQty is ZERO   (0011)
'                    **Jeanson      2008.01.14      don't check the date when PMC uplaod WO Plan   (0012)
'                    **Sandy        2008.01.16      Upload data to MaterialToWHID   (0013)
'                    **Jing         2008.02.01      Update (upload the WOPlanSeq) for XL_WOPlanSeq(enddatetime) (0014)
'                    **Jing         2008.02.12      Update (enddatetime) (0015)
'                    **Jing         2008.02.14      Changed from CInt() to CDBL()   (0016)
'                    **Jing         2008.02.19      Upload FixPN to XL_PNOneByOne  (0017)
'                    **Steven       2008.02.27      add three items(Frequency,voltage,current) to QSMS_Inspect_rule (0018)
'                    **Jing         2008.03.02      Upload MBPN/PN/Interval to XL_PNInterval   (0019)
'                    **Udall        2008.03.03      Add a comppn "QJ" for upload UnCHKCompPN   (0020)
'                    **Kane         2008.03.17      Check if all WOs in one PCB check bom pass when upload PMC plan to xl_woplanseq(0021)
'                    **Jing         2008.03.23      Upload WOPlanEC to XL_EC_WOPlan  (0022)
'                    **Jing         2008.03.31      Check only one wo to upload XL-WOPlanSeq in one PCBGroup (0023)
'                    **Archer       2008.04.01      upload the DoubleTables setting    (0024)
'                    **Kane         2008.04.01      Show wos when upload replace pn maybe impacted (0025)
'                    **Archer       2008.04.24      Upload The Equipment For Process Of Work Hour System      (0026)
'                    **Archer       2008.04.24      Upload The Line Configuration Of Work Hour System (0027)
'                    **Jing         2008.05.09      update XL_WOPlanSeq (0028)
'                    **Sandy        2008.05.19      mark 'REPLACEPN' (0029)
'                    **Salon        2008.05.22      Add Upload JobGroup (0030)
'                    **Jing         2008.06.03      add UserRight for upload ReplacePN    (0031)
'                    **Jing         2008.06.06      add a flag for delete from ReplacePN    (0032)
'                    **Kevin        2008.07.04      add Factory for load_XL_WOPlanSeq()    (0033)
'                    **Giant        2008.07.07      add Factory for Load_QSMS_BOM()    (0034)
'                    **Kevin        2008.07.08      add condition to make sure wo suit to be the factory(0035)
'                    **Giant        2008.07.21      add user can upload Lost ReplacePN (0036)
'                    **Jing         2008.07.29      if the WOInput Qty is less than before when PMC upload scheduling  (0037)
'                    **Denver       2008.08.04      Add Factory  (0038)
'                    **Jing         2008.08.05      Show Fail Message when upload Multiple Same WO in Date/Shift/Line/Factory   (0039)
'                    **Giant        2008.07.07      add new function Load_LineFUJIServer to upload fuji server list by line    (0040)
'                    **Kane         2008.08.22      Add new function to upload XL_MaxDIDMaintainQty (0041)
'                    **Udall        2008.09.10      Add a new function for upload the QSMS_UnCheckCompPN table data (0042)
'                    **Kane         2008.09.26      remove begindatetime and enddatetime from xl_woplanseq excel '(0043)
'                    **Jeanson      2009.01.19      order by Datetime '(0044)
'                    **Sandy        2009.01.20      跟新MEBOM格式根据NB5新机器(0045)
'                    **Sandy        2009.02.06      add condition to make sure Single board diffent of Wave board(0046)
'                    **Sandy        2009.03.20      add upload plan line check (0047)
'                                  **Kevin        2009.04.07      show action row when upload XL_PNINTERVAL (0048)
'                                  **Sandy        2009.04.20      add upload plan line check (0049)
'                                  **Sandy        2009.05.04      add upload glue Dailyschedule plan (0050)
''EQMS:RQ09052607    **Kane         2009.06.04      add upload no check replacepn when splicing '(0051)
''QMS                         **Sandy        2009.06.08      在同一厂区的排成，上传到相应的BU(0052)
''QMS                         **Sandy        2009.06.16      在同一厂区的排成，上传到相应的BU,再次上传更新全部数据(0053)
''RQ09052928              **Sandy        2009.06.17      Me BOM时在上传*Other机台时，只需要上传“*”号线别复制到其他的线别(0054)
''QMS                         **Sandy        2009.07.22      add upload Machine function and delete FrmuploadMachineType;(0055)
''QMS                         **Sandy        2009.07.28      转换为字符型(0056)
''QMS                         **Lynn         2009.07.30      Add "Factory='F(x)'" in where command, or it will fail when two factory have same mebom(0057)
''ESBU                        **Denver       2009.08.04      Add upload PNGroup and check PNGroup when create WO group  (0058)
''QMS                        **Sandy        2009.08.12      check CompPN in JobPN (0059)
''QMS                        **Sandy        2009.08.18      更改排程上传提示（0060）
''QMS                        **Udall        2009.09.02      在合适的条件下覆盖旧有的Mebom（0061）
''                               **Richie       2009.10.19      上传机种对应的IC料号（0062）
''RQ09100707            **Udall        2009.10.22      上传在一个PCB(可能是多联板)中仅需要使用一个的CompPN（0063）
''RQ09052928             **Sandy        2009.10.27      Me BOM时在上传*Other机台时，只需要上传“*”号线别复制到其他的线别(0064)
''QMS                        **Lynn         2009.12.02      增加上传WO_AssignPN的功能。(0065)
''QMS                        **Lynn         2009.12.14      Check the unuseful data which CompPN=AssignPN and Vendor=''。(0066)
''RQ09121405             **Sandy        2009.12.30      Modify Upload machine bom function to check whether the slot is 0, if it is 0, show error message to user.(0067)
''QMS                        **Kevin        2010.04.28      增加记录上传XL排程时候的记录 (0068)
''RQ10051760            **Kevin        2010.05.24      UPLOAD BOM?MaterialToWHID需增加删除项  (0069)
''20100811                **Denver       2010.08.11      对于MachineType 删除操作，不需要完全检查信息有效性(0070)
''QMS                       **Walfan       2013.04.30      上传XL_WOPanSeq时增加Flag标识那些工单为CTO工单 (0071) --1135
'' QMS                      **Yan           2014.07.01       QSMS_MEBom数据太大，当前卡1W行
'***********************************************************************************/


Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private Const ReplacePN_MAX_Num = 30
Public ReplaceCutStr As String
Dim ReplacePNList(1 To ReplacePN_MAX_Num) As String
Public Total_Qty, Deleted_Qty, Insert_Qty As Long
Dim strtemp As String
Dim tempmachine, TempJobPn, tempVersion As String, TempJObGroup As String, TempFactory As String, TempLine As String
Dim UpdateBOM_Qty, InsertBOM_Qty As Long


Private Sub CboFuncType_Change()
If Trim(CboFuncType) <> "" Then
    CmdExcel.Enabled = True
End If
End Sub

Private Sub CboFuncType_Click()
If Trim(CboFuncType) <> "" Then
    CmdExcel.Enabled = True
    
    Select Case UCase(CboFuncType)
        Case "AVL"
              Me.lblFileFormat = App.Path & "\Template\AVL.xls"
        Case UCase("QSMS_CheckCompPN")
              Me.lblFileFormat = App.Path & "\Template\QSMS_CheckCompPN.xls"
        Case "MACHINETYPE"
              Me.lblFileFormat = App.Path & "\Template\MACHINETYPE.xls"
        Case "AVL-WIN"
              Me.lblFileFormat = App.Path & "\Template\AVL-WIN.xls"
        Case "CONTROLPARTS"
              Me.lblFileFormat = App.Path & "\Template\CONTROLPARTS.xls"
        Case "NEXTDEVICE"
              Me.lblFileFormat = App.Path & "\Template\NEXTDEVICE.xls"
        Case "NONAVL"
              Me.lblFileFormat = App.Path & "\Template\NONAVL.xls"
        Case "QSMS_MEBOM"
              Me.lblFileFormat = App.Path & "\Template\QSMS_MEBOM.xls"
        Case "SINGLESIDEBRD"
              Me.lblFileFormat = App.Path & "\Template\SINGLESIDEBRD.xls"
        Case "NEGATIVEBRD"
              Me.lblFileFormat = App.Path & "\Template\NEGATIVEBRD.xls"
        Case "REPLACEPN"
              Me.lblFileFormat = App.Path & "\Template\REPLACEPN.xls"
               Case "DOCUMENTCOMP"
                     Me.lblFileFormat = App.Path & "\Template\DOCUMENTCOMP.xls"
        Case "UNCHKCOMP"
              Me.lblFileFormat = App.Path & "\Template\UNCHKCOMP.xls"
        Case UCase("FujiBrdSeqMapping")
              Me.lblFileFormat = App.Path & "\Template\FujiBrdSeqMapping.xls"
        Case UCase("PhilipsBrdSeqMapping")
             Me.lblFileFormat = App.Path & "\Template\PhilipsBrdSeqMapping.xls"
        Case UCase("LineFUJIServer")
            Me.lblFileFormat = App.Path & "\Template\LineFUJIServer.xls"
        Case UCase("TraySlot")
            Me.lblFileFormat = App.Path & "\Template\TraySlot.xls"
        Case "UPDATEJOBPN"
            Me.lblFileFormat = App.Path & "\Template\AVL-WIN.xls"
        Case UCase("CTO_Model")
            Me.lblFileFormat = App.Path & "\Template\CTO_Model.xls"
        Case "CASTRATE"
            Me.lblFileFormat = App.Path & "\Template\CASTRATE.xls"
        Case "ONEBYONE"
            Me.lblFileFormat = App.Path & "\Template\ONEBYONE.xls"
        Case "JOBSIDE"
           Me.lblFileFormat = App.Path & "\Template\JOBSIDE.xls"
        Case "BRDCOMBINEQTY"
            Me.lblFileFormat = App.Path & "\Template\BRDCOMBINEQTY.xls"
        Case "DIO"
            Me.lblFileFormat = App.Path & "\Template\DIO.xls"
        Case "NOMACHINEDROPCOMPPN"
            Me.lblFileFormat = App.Path & "\Template\NOMACHINEDROPCOMPPN.xls"
        Case "LOSTREPLACEPN"
            Me.lblFileFormat = App.Path & "\Template\LOSTREPLACEPN.xls"
        Case "COMPPNINSPECTRULE"
            Me.lblFileFormat = App.Path & "\Template\COMPPNINSPECTRULE.xls"
        Case "PNALARMQTY"
            Me.lblFileFormat = App.Path & "\Template\PNALARMQTY.xls"
        Case "WOSCHEDULELIST"
            Me.lblFileFormat = App.Path & "\Template\WOSCHEDULELIST.xls"
        Case "XL_WOPLANSEQ"
            Me.lblFileFormat = App.Path & "\Template\XL_WOPLANSEQ.xls"
        Case "XL_WOPLANSEQSHIFTID"
            Me.lblFileFormat = App.Path & "\Template\XL_WOPLANSEQShiftID.xls"       '''(1121)
        Case "XL_WOPLANLINE"
            Me.lblFileFormat = App.Path & "\Template\XL_WOPlanLine.xls"
        Case UCase("Daily Schedule")
            Me.lblFileFormat = App.Path & "\Template\Daily Schedule.xls"
        Case "XL_IMPLEMENTPN"
            Me.lblFileFormat = App.Path & "\Template\XL_IMPLEMENTPN.xls"
        Case "XL_WOPN"
            Me.lblFileFormat = App.Path & "\Template\XL_WOPN.xls"
        Case "MATERIALTOWHID"
            Me.lblFileFormat = App.Path & "\Template\MATERIALTOWHID.xls"
        Case "XL_PNONEBYONE"
            Me.lblFileFormat = App.Path & "\Template\XL_PNONEBYONE.xls"
        Case "XL_PNINTERVAL"
            Me.lblFileFormat = App.Path & "\Template\XL_PNINTERVAL.xls"
        Case "XL_ECWOPLAN"
            Me.lblFileFormat = App.Path & "\Template\XL_ECWOPLAN.xls"
        Case "XL_DOUBLETABLES"
            Me.lblFileFormat = App.Path & "\Template\XL_DOUBLETABLES.xls"
        Case "UPLOAD_JOBGROUP"
            Me.lblFileFormat = App.Path & "\Template\UPLOAD_JOBGROUP.xls"
        Case "XL_MAXDIDMAINTAINQTY"
            Me.lblFileFormat = App.Path & "\Template\XL_MAXDIDMAINTAINQTY.xls"
        Case "NOCHECKREPLACEPNSPLICING"
            Me.lblFileFormat = App.Path & "\Template\NOCheckReplacePNSplicing.xls"
         Case "PNGROUP"         '**Denver       2009.08.04      Add upload PNGroup and check PNGroup when create WO group  (0058)
            Me.lblFileFormat = App.Path & "\Template\PNGroup.xls"
    '    Case "IC_COMPPN"        '**Richie        2009.10.19      （0062）
     '       Me.lblFileFormat = App.Path & "\Template\IC_CompPN.xls"
        Case "IC_SHEARPIN"        '**Richie        2009.10.19      （0062）
            Me.lblFileFormat = App.Path & "\Template\Shearpintemplate.xls"
        Case "2NDSOURCE_ASSIGNPN"
            Me.lblFileFormat = App.Path & "\Template\2NDSOURCE_ASSIGNPN.xls"
        Case "UPLOAD_TRAYCOMPPN"   '2010.08.17 add by kaitlyn (1001)
            Me.lblFileFormat = App.Path & "\Template\upload_traycompPN.xls"
        Case "COMPONENT_DATA"      '2010.12.08 add by kaitlyn (1024)
            Me.lblFileFormat = App.Path & "\Template\Component_data.xls"
        Case "MACHINE_DATA"                               '''(1141)
            Me.lblFileFormat = App.Path & "\Template\Machine_DATA.xls"
        Case "COMPPN_SPACER"                               '''(1154)
            Me.lblFileFormat = App.Path & "\Template\CompPN_Spacer.xls"
        Case "AVLC"
            Me.lblFileFormat = App.Path & "\Template\AVLC.xls"
        Case "A8_MANUAL"
            Me.lblFileFormat = App.Path & "\Template\A8_Manual.xls"
        Case "A8_DIDTYPE"
            Me.lblFileFormat = App.Path & "\Template\A8_DIDType.xls"
        Case Else
             MsgBox "Please check the right sheet name."
    End Select
End If
End Sub

Private Sub CboFuncType_KeyPress(KeyAscii As Integer)
If Trim(CboFuncType) <> "" Then
    CmdExcel.Enabled = True
End If

End Sub
Function upLoad_MachineType(ByVal xlsSheetName As String) As Boolean '0055
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rCount, Row_Count As Long
Dim Deleted_Qty As Integer
Dim Vendor  As String, Factory  As String, Machine As String, Unit As String, Qty As String, MaxSlotNum, LR As String, FujiData As String, Line As String, Side As String, DeletedFlag As String
Dim TempJobPn, tempVersion, TempLine As String
Dim Total_Qty, Update_Qty, Insert_Qty As Long
Dim i As Integer, SeqIDByLine As Integer
Dim SearchChar1, MyPos1
Dim MappingID, DIOCircuit As String
Dim M1 As String, M2 As String, N1 As String
Dim N2 As Integer, N3 As Integer, N4 As Integer
Dim strSQL As String
Dim PreLine As String
Dim LineArray As String
Dim Rs As ADODB.Recordset
PreLine = ""
If UCase(xlsSheetName) <> "MACHINETYPE" Then
    Exit Function
End If
Set xlApp = CreateObject("Excel.Application")
Let xlApp.Visible = False
Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.DisplayAlerts = False
upLoad_MachineType = False

    rCount = 2
    Total_Qty = 0
    Insert_Qty = 0
    Update_Qty = 0
    TempJobPn = ""
    SearchChar1 = "-"
    
    With xlsBook.Worksheets(Trim(xlsSheetName))
    
        While Trim(.Cells(rCount, 1)) <> ""
             Vendor = Trim(.Cells(rCount, 1) & vbNullString)
             Line = Trim(.Cells(rCount, 2) & vbNullString)
             '''''''''''(1029)'''''''''
             Side = Trim(.Cells(rCount, 3) & vbNullString)
             Factory = Trim(.Cells(rCount, 4) & vbNullString)
             Machine = Replace(Trim(.Cells(rCount, 5) & vbNullString), "'", " ")
             Unit = Trim(.Cells(rCount, 6) & vbNullString)
             Qty = Trim(.Cells(rCount, 7) & vbNullString)
             MaxSlotNum = Trim(.Cells(rCount, 8) & vbNullString)
             LR = Trim(.Cells(rCount, 9) & vbNullString)
             MappingID = Trim(.Cells(rCount, 10) & vbNullString)
             FujiData = Trim(.Cells(rCount, 11) & vbNullString)
             DIOCircuit = Trim(.Cells(rCount, 12) & vbNullString)
             DeletedFlag = Trim(.Cells(rCount, 13) & vbNullString)
 
             If IsNumeric(Trim(.Cells(rCount, 14))) Then
                SeqIDByLine = Trim(.Cells(rCount, 14))  '(1012)
             Else
                MsgBox "SeqIDByLine must be number in the file you upload"
             End If
             If Trim(Line) = "" Then
                  MsgBox "Line is empty,it must define ! Row:" & rCount
                  Exit Function
             End If
             If Trim(Side) = "" Then
                  MsgBox "Side is empty,it must define ! Row:" & rCount
                  Exit Function
             End If
             
             '''1040
'             If Left(UCase(machine), 1) <> UCase(Line) Then '---(0003)(1029)
'                  MsgBox "The Machine= " & machine & "  or line= " & Line & " is define wrong! Row:" & rCount
'                  Exit Function
'             End If
           
            ''20100811           **Denver       2010.08.11      对于MachineType 删除操作，不需要完全检查信息有效性(0070)'''''''''''(1029)'''''''''
            If UCase(DeletedFlag) = "Y" Then
                strSQL = "delete from Machine where Line='" & Trim(Line) & "' and Machine='" & Trim(Machine) & "' "
                Conn.Execute strSQL
                Call InsertIntoQSMSLog("SMT_QSMS", "Delete Machine Type", "machine=" & Machine & " and factory=" & Factory & "")
                Deleted_Qty = Deleted_Qty + 1
                
                GoTo DoNext
            End If
           
            '*********************JUDGE machine name Begin add by giant 061110*************************************************
            ''(1065) mark by kaitlyn
            
''            If InStr(UCase(machine), "NXT") = 0 Then
''                If InStr(UCase(machine), "OTHERS") = 0 Then
''                    'Non-NXT machine name is like ASCP7A, ASCP7B, ...
''                    If Not UCase(machine) Like "[A-Z,0-9][S,C]???[A-Z]" Then  '(1021)
''                        MsgBox "The Machine name is not correct:" & machine & vbCrLf & "Format should be [A-Z,0-9][S,C]???[A-Z]" & vbCrLf & "Row:" & rCount
''                        Exit Function
''                    End If
''                ElseIf InStr(UCase(machine), "OTHERS") > 0 Then
''                    'Non-NXT machine name is like ASOthers, ACOthers, ...
''                    If Not UCase(machine) Like "[A-Z,0-9][S,C,Q,W]OTHERS*" Then  '(1021)
''                        MsgBox "The Machine name is not correct:" & machine & vbCrLf & "Format should be [A-Z,0-9][S,C,Q,W]OTHERS*" & vbCrLf & "Row:" & rCount
''                        Exit Function
''                    End If
''                End If
''            ElseIf InStr(UCase(machine), "NXT") > 0 Then
''                'NXT machine name is like ASNXTA01, ASNXTA02, ...
''                If Not UCase(machine) Like "[A-Z,0-9][S,C]NXT[A-Z][0-9][0-9]" Then '(1021)
''                    MsgBox "The Machine name is not correct:" & machine & vbCrLf & "NXT format should be [A-Z,0-9][S,C]NXT[A-Z][0-9][0-9]" & vbCrLf & "Row:" & rCount
''                    Exit Function
''                End If
''            End If
            
            '*********************JUDGE format Begin *********************************************************
           If InStr(UCase(Machine), "OTHERS") = 0 Then  ''(1046)
            If UCase(Vendor) = "FUJI" Then
                 MyPos1 = InStr(MappingID, SearchChar1)
                 If MyPos1 = 0 Then
                     MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
                     Exit Function
                 Else
                     N3 = Mid(MappingID, MyPos1 + 1)
                     N4 = Mid(MappingID, 1, MyPos1 - 1)
                     If N3 > 20 Then
                         MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
                         Exit Function
                     End If
                     If N4 < 1 Or N4 > 5 Then
                         MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
                         Exit Function
                     End If
                 End If
             Else
                 '''''''''''''''(1039)''''''''''''''
                 If UCase(Vendor) = "PANA" Then
                     If Len(MappingID) <> 7 Then
                         MsgBox ("Wrong MappingID(MappingID length must be 7 characters) :" & MappingID & ", Row:" & rCount)
                         Exit Function
                     End If
                 End If
                 '''''''''''''''(1039)'''''''''''''
                 M1 = Left(Machine, 2)
                 M2 = Left(MappingID, 2)
                 N1 = Left(MappingID, 2)
                 If IsNumeric(Right(MappingID, 5)) = True Then
                     N2 = CInt(Right(MappingID, 5))
                 End If
                 
                 If N2 < 1 Then  '0002
                     MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
                     Exit Function
                 End If
                 
                 'If M1 <> M2 Or N1 <> "MC" Then
                     'MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
                     'Exit Function
                 'End If
             End If
            End If
           
                If IsNumeric(Mid(FujiData, 1)) = True Then
                    N2 = CInt(Mid(FujiData, 1))
                    If N2 > 9 Then
                        MsgBox ("Wrong FujiData :" & FujiData & ", Row:" & rCount)
                        Exit Function
                    End If
                Else
                        MsgBox ("Wrong FujiData :" & FujiData & ", Row:" & rCount)
                        Exit Function
                End If
                
                If IsNumeric(Mid(DIOCircuit, 1)) = True Then
                    N2 = CInt(Mid(DIOCircuit, 1))
                    If N2 > 9 Then
                        MsgBox ("Wrong DIOCircuit :" & DIOCircuit & ", Row:" & rCount)
                        Exit Function
                    End If
                Else
                        MsgBox ("Wrong DIOCircuit :" & DIOCircuit & ", Row:" & rCount)
                        Exit Function
                End If
            '*********************Delecte machine information by Line ***************************************************
'                If UCase(DeletedFlag) = "Y" Then
'                   strSQL = "delete from Machine where Machine='" & Trim(machine) & "' "
'                   Conn.Execute strSQL
'                   Call InsertIntoQSMSLog("SMT_QSMS", "Delete Machine Type", "machine=" & machine & " and factory=" & Factory & "")
'                   Deleted_Qty = Deleted_Qty + 1
'                Else
            '*********************insert or update machine information ***************************************************
            
            ''20100811           **Denver       2010.08.11      对于MachineType 删除操作，不需要完全检查信息有效性(0070)
            If UCase(DeletedFlag) <> "Y" Then
                '''''''''''(1029)'''''''''
               strSQL = "select * from Machine where Line='" & Trim(Line) & "' and Machine='" & Trim(Machine) & "'"
               Set Rs = Conn.Execute(strSQL)
               If Rs.EOF Then
                   strSQL = "Insert into Machine(Vendor,Factory,Line,Machine,Unit,SeqIDByLine,Qty,MaxSlotNum,LR,MappingID,FujiData,DIOCircuit,OPID,Side) " & _
                               " values('" & Trim(Vendor) & "','" & Factory & "','" & Trim(Line) & "','" & Trim(Machine) & "','" & Trim(Unit) & "','" & Trim(SeqIDByLine) & "','" & Trim(Qty) & "','" & Trim(MaxSlotNum) & "','" & Trim(LR) & "','" & Trim(MappingID) & "','" & Trim(FujiData) & "','" & Trim(DIOCircuit) & "','" & Trim(g_userName) & "','" & Trim(Side) & "')" '----(0001)
                   Conn.Execute strSQL
                   Insert_Qty = Insert_Qty + 1
               Else
                   strSQL = "Update Machine set Side='" & Trim(Side) & "',Vendor='" & Trim(Vendor) & "',Factory='" & Factory & "',Unit='" & Trim(Unit) & "',SeqIDByLine='" & Trim(SeqIDByLine) & "',Qty='" & Trim(Qty) & "',MaxSlotNum='" & Trim(MaxSlotNum) & "',LR='" & Trim(LR) & "',MappingID='" & Trim(MappingID) & "',FujiData='" & Trim(FujiData) & "',DIOCircuit='" & Trim(DIOCircuit) & "',OPID='" & Trim(g_userName) & "' where Line='" & Trim(Line) & "' and Machine='" & Trim(Machine) & "'" '----(0001)
                   Conn.Execute strSQL
                   Update_Qty = Update_Qty + 1
               End If
            End If
            '''''''''''(1029)'''''''''
            
            If PreLine <> Line Then '---(0003)
                strSQL = "select line,count(distinct vendor) from machine where line='" & Trim(Line) & "' and vendor<>'DIP' group by line  having count(distinct vendor)>1"
                Set Rs = Conn.Execute(strSQL)
                If Rs.EOF = False Then
                    strSQL = "select machine,vendor from machine where line='" & Trim(Line) & "' and vendor<>'DIP'"
                    If Rs.State Then Rs.Close
                    Set Rs = Conn.Execute(strSQL)
                    While Not Rs.EOF
                        LineArray = LineArray + Trim(Rs!Machine) + "-->" + Trim(Rs!Vendor) + vbCrLf
                        Rs.MoveNext
                    Wend
                    MsgBox ("There are multiple MachineVendors on the same line, please confirm whether they are correct?" & vbCrLf & "" & LineArray & "If it is not correct, please delete or update the wrong information later, and click OK to continue uploading the information.")
                End If
                PreLine = Line
            End If
            
DoNext:
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1


        Wend
    End With
upLoad_MachineType = True
strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','upLoad_MachineType','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSQL)

xlsBook.Close
xlApp.Quit
Set xlApp = Nothing
Set xlsBook = Nothing
Conn.Execute ("exec GenUpdateMachineType")     '*******************add by jeanson 10/10*******

MsgBox "*** Load  finish ! ***" & "   " & vbCrLf & _
             "Total Counter : " & Total_Qty & vbCrLf & _
             "Insert succeed : " & Insert_Qty & vbCrLf & _
             "Update succeed : " & Update_Qty & vbCrLf & _
             "Delete Qty : " & Deleted_Qty
              
End Function





Private Sub Cmd_Load_Click()
Dim position As Long


If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Or Trim(CboFuncType) = "" Then
    MsgBox ("SheetName and FilePath and FuncType can not be NULL !")
    Exit Sub
End If

'If UCase(Right(txtFilePath, 3)) = "XLS" Then '''(1028)
'
    If Trim(CboFuncType) <> "FujiBrdSeqMapping" And Trim(CboFuncType) <> "PhilipsBrdSeqMapping" Then
        If UCase(Trim(cboSheetName)) <> UCase(Trim(CboFuncType)) Then
            MsgBox "Function type  does not match the sheet name ,please check" & "sheet name should the same as function type"
            Exit Sub
        End If
    End If
'End If



Select Case UCase(CboFuncType)
    Case "AVL"
          Call Load_AVL(cboSheetName)
    Case UCase("QSMS_CheckCompPN")
          Call Load_QSMS_CheckCompPN(cboSheetName)
    Case UCase("MachineType") '0055
        If upLoad_MachineType(Trim(cboSheetName)) = False Then
            MsgBox ("Fail")
        Else
            MsgBox ("Finish")
        End If
    Case "AVL-WIN"
          Call Load_AVL_WIN(cboSheetName)
    Case "CONTROLPARTS"
          Call Load_AVL_ControlParts(cboSheetName)
    Case "NEXTDEVICE"
          Call Load_NextDevice(cboSheetName)
    Case "NONAVL"
          Call Load_NonAVL(cboSheetName)
    '       Case "DID"
    '             Call Load_DID(cboSheetName)
    Case "QSMS_MEBOM"
          Call Load_QSMS_BOM(Trim(cboSheetName))
    Case "SINGLESIDEBRD"
          Call Load_SingleSideBrd(Trim(cboSheetName))
    Case "NEGATIVEBRD"
          Call Load_NegativeBrd(Trim(cboSheetName))
    ' (0029)
    Case "REPLACEPN"
          Call Load_ReplacePN(cboSheetName)
           Case "DOCUMENTCOMP"
                 Call Load_DocumentComp(cboSheetName)
    Case "UNCHKCOMP"
          Call Load_UnChkComp(cboSheetName)
    Case UCase("FujiBrdSeqMapping")
          Call Load_FujiBrdSeqMapping(cboSheetName)
    Case UCase("PhilipsBrdSeqMapping")
          Call Load_PhilipsBrdSeqMapping(cboSheetName)
    '       Case "MACHINETYPE"
    '             Call Load_MachineType(cboSheetName)
    Case UCase("LineFUJIServer")
        Call Load_LineFUJIServer(cboSheetName)
    Case UCase("TraySlot")
        Call Load_TraySlot(cboSheetName)
    Case "UPDATEJOBPN"
        Call Load_UpdateJobPn(cboSheetName)
    Case UCase("CTO_Model")
        Call Upload_CTO_Model(cboSheetName)
    Case "CASTRATE"
        Call Load_CastRate(cboSheetName)
    Case "ONEBYONE"
        Call Load_OneByOne(cboSheetName)
    Case "JOBSIDE"
        Call Load_JobSide(cboSheetName)
    Case "BRDCOMBINEQTY"
        Call Load_BrdCombineQty(cboSheetName)
    Case "DIO"
        Call Load_coputername(cboSheetName)
    Case "NOMACHINEDROPCOMPPN"                  ''(0042)
        Call Load_NoMachineDropCompPN(cboSheetName)
    Case "PCB_SINGLECOMPPN"                  ''(0063)
        Call Load_PCB_SingleCompPN(cboSheetName)
    Case "LOSTREPLACEPN"
        Call Load_LostReplacePN(cboSheetName)   '(0036)
    Case "COMPPNINSPECTRULE"
        Call Load_InSpectRule(cboSheetName)
    Case "PNALARMQTY"        '''''''add by Jing 2007.10.30   (0004)'''''''
        Call Load_PNAlarmQty(cboSheetName)
    Case "WOSCHEDULELIST"    '''''''add by Jing 2007.11.19   (0006)'''''''
        Call Load_WOScheduleList(cboSheetName)
    Case "XL_WOPLANSEQ"      '''''''add by Jing 2007.11.26   (0007)'''''''
         Call load_XL_WOPlanSeq(cboSheetName)
    Case "XL_WOPLANSEQSHIFTID"      '''''''add by Newton 2013.01.25'''''(1121)
         Call load_XL_WOPlanSeqShiftID(cboSheetName)
    Case "XL_WOPLANLINE" '0047
         Call load_WOPlanLine(cboSheetName)
    Case UCase("Daily Schedule") '0050
         Call load_DailySchedule(cboSheetName)
    Case "XL_IMPLEMENTPN"    '''''''add by Jing 2007.12.05   (0008)'''''''
        Call load_XL_ImplementPN(cboSheetName)
    Case "XL_WOPN"           '''''''add by Jing 2007.12.17   (0009)'''''''
        Call Load_XL_WOPN(cboSheetName)
    Case "MATERIALTOWHID"    ''''''''add by Sandy 2008.01.17   (0013)......
        Call XL_MATERIALTOWHID(cboSheetName)
    Case "XL_PNONEBYONE"    ''''''add by Jing 2008.02.19   (0017)''''''
        Call load_XL_PNOneByOne(cboSheetName)
    Case "XL_PNINTERVAL"       ''''''add by Jing 2008.03.02    (0019)''''''
        Call load_XL_PNInterval(cboSheetName)
    Case "XL_ECWOPLAN"         ''''''added by Jing 2008.03.23   (0022)''''''
        Call load_XL_ECWOPlan(cboSheetName)
    Case "XL_DOUBLETABLES"      '''''''Added by Archer  2008.04.01   (0024)'''''''
        Call load_XL_DoubleTables(cboSheetName)
'    Case "WORKHS_EQUIPMENT"        ''''''Add by  Archer 2008/04/18     (0026)''''''
'        Call load_WorkHS_Equipment(cboSheetName)
'    Case "WORKHS_LINECONFIG"   '''''Add by Archer 2008/04/21    (0027) '''''
'        Call load_WorkHS_LineConfig(cboSheetName)
    Case "UPLOAD_JOBGROUP"      '''''''Added by Salon  2008.05.22   (0030)'''''''
        Call Upload_JobGroup(cboSheetName)
    Case "XL_MAXDIDMAINTAINQTY"  '(0041)
        Call XL_MaxDIDMaintainQty(cboSheetName)
    Case "NOCHECKREPLACEPNSPLICING" '(0051)
        Call Upload_NOCheckReplacePNSplicing(cboSheetName)
        
     Case "PNGROUP"      '''''''**Denver       2009.08.04      Add upload PNGroup and check PNGroup when create WO group  (0058)
        Call Upload_PNGroup(Trim(txtFilePath), cboSheetName)
'    Case "IC_COMPPN"    '''''add by Richie  2009.10.19  (0062)
 '       Call UploadIC_CompPN(cboSheetName)
    Case "IC_SHEARPIN"    '''''add by Richie  2009.10.19  (0062)
        Call UploadIC_ShearPin(cboSheetName)
    Case "2NDSOURCE_ASSIGNPN"
        Call Load_WO_AssignPN(Trim(cboSheetName)) ''(0065)
    Case "UPLOAD_TRAYCOMPPN"    '2010.08.17 add by kaitlyn '(1001)
        Call upload_traycompPN(Trim(cboSheetName))
    Case "COMPONENT_DATA"       '2010.12.08 add by kaitlyn (1024)
        Call upload_ComponentData(Trim(cboSheetName))
    Case "MACHINE_DATA"                              '(1129)
        Call Upload_MachineData(Trim(cboSheetName))
    Case "COMPPN_SPACER"                               '''(1154)
        Call Upload_CompPNSpacer(Trim(cboSheetName))
    Case "AVLC"                               '''(1195)
        Call Upload_AVLC(Trim(cboSheetName))
    Case "A8_MANUAL"                               '''(1195)
        Call Upload_A8_Manual(Trim(cboSheetName))
    Case "A8_DIDTYPE"                               '''(1195)
        Call Upload_A8_DIDType(Trim(cboSheetName))
    Case Else
         MsgBox "Please check the right sheet name."
End Select

End Sub

Private Sub CmdExcel_Click()
Dim str As String
Dim Rs As ADODB.Recordset
Select Case UCase(CboFuncType)
        Case "AVL"
            str = "select  * from QSMS_AVL order by Customer"
        Case "AVL-WIN"
            str = "select  * from QSMS_AVL order by Customer"
        Case "CONTROLPARTS"
            str = "select   * from QSMS_ControlPart order by Model"
        Case "NEXTDEVICE"
            str = "select   * from QSMS_MEBom_NextDevice order by NextDeviceID,Machine"
        Case "NONAVL"
            str = "select   * from QSMS_NonAVL order by CompPN"
        Case "NOMACHINEDROPCOMPPN"
            str = "select   * from QSMS_UnCheckCompPN where Type='NOMDrop' order by CompPN"
        Case "SINGLESIDEBRD"
            str = "select   * from QSMS_SingleSideBrd order by MBPN"
        Case "NEGATIVEBRD"
            str = "select   * from QSMS_NegativeBrd order by MBPN"
        Case "DID"
            str = "select   * from QSMS_DID order by DID"
        Case "QSMS_MEBOM"
            str = "select   * from QSMS_MEBom order by Machine,JObPN,Version,CompPN"
        Case "REPLACEPN"
            str = "select   * from QSMS_ReplacePN order by JobPN,Version,ID"
        Case "TRAYSLOT"
            str = "Select   * from TraySlot order by machine"
        Case "DOCUMENTCOMP"
            str = "select   * from QSMS_DocuComp order by transdatetime desc"
        Case "UNCHKCOMP"
            str = "select   * from QSMS_UnChkComp order by CompHead"
        Case "MACHINETYPE"
            str = "select   * from Machine order by Machine"
        Case "UPDATEJOBPN"
            str = "select   * from QSMS_UpdateJobPN order by Model"
        Case "CTO_MODEL"
            str = "select   * from CTO_Model order by Model"
        Case "CASTRATE"
            str = "select   * from QSMS_CastRate order by CompHead"
        Case "DAILY SCHEDULE"
            str = "select   * from XL_DailySchedule order by WORKDATE DESC, LINE"
        Case "ONEBYONE"
            str = "select   * from QSMS_OneByOne order by CompPN"
        Case UCase("FujiBrdSeqMapping")
            str = "select   * from FujiBrdSeqMapping"
        Case UCase("PhilipsBrdSeqMapping")
            str = "select   * from PhilipsBrdSeqMapping"
        Case UCase("ComppnInSpectRule")
            str = "select   * from QSMS_InSpect_Rule"
        Case UCase("PNALARMQTY")                 '''''''add by Jing 2007.10.30   (0004)'''''''
            str = "select   * from PNALARMQTY"
        Case UCase("WOSCHEDULELIST")             '''''''add by Jing 2007.11.19   (0006)'''''''
            str = "select   * from WOScheduleList"
        Case UCase("XL_WOPLANSEQ")               '''''''add by Jing 2007.11.26   (0007)'''''''
            str = "select top(200) Date,Shift,Line,WO,PlanQty,SeqID,TransDateTime,OPID,InputQty,Factory from XL_WOPlanSeq order by Date Desc,shift,Line,SeqID"  ''(0043)(0044)
        Case "XL_WOPLANSEQSHIFTID"
            str = "select top(200) Date,Shift,Line,WO,PlanQty,SeqID,TransDateTime,OPID,InputQty,Factory,shiftID from XL_WOPlanSeq order by Date Desc,shift,Line,SeqID,shiftid"  ''（1123）
        Case "XL_WOPLANLINE" '0047
            str = "select   * from XL_WOPlanLine"
        Case UCase("XL_IMPLEMENTPN")             '''''''add by Jing 2007.12.05   (0008)'''''''
            str = "select distinct PrefixPN,UID,TransDateTime from XL_ImplementPN"
        Case UCase("XL_WOPN")                    '''''''add by Jing 2007.12.17   (0009)'''''''
            str = "select   * from XL_WOPN"
        Case "MATERIALTOWHID"    ''''''''add by Sandy 2008.01.17   (0013)......
            str = "select   * from MATERIALTOWHID"
        Case "XL_PNONEBYONE"    ''''''add by Jing   2008.02.19  (0017)''''''
            str = "select   * from xl_PNOneByOne"
        Case "XL_PNINTERVAL"    ''''''add by Jing   2008.03.02  (0019)''''''
            str = "select   * from xl_PNInterval"
        Case "XL_ECWOPLAN"      ''''''added by Jing     2008.03.23  (0022)''''''
            str = "select   * from XL_EC_WOPlan"
        Case "XL_DOUBLETABLES"
            str = "select   * from DoubleTables"
'        Case "WORKHS_EQUIPMENT"        ''''''Add by  Archer 2008/04/18     (0026)''''''
'            Str = "select * from WorkHS_Equipment"
'        Case "WORKHS_LINECONFIG"       ''''''Add by Archer 2008/04/21  (0027)'''''
'            Str = "select * from WorkHS_LineConfig"
        Case "NOCHECKREPLACEPNSPLICING"
            str = "SELECT * FROM QSMS_NOCheckReplacePNSplicing"
        Case "PNGROUP"
            str = "SELECT * FROM PN_Group Order by PNGroup,PN "
    '    Case "IC_COMPPN"    '''''add by Richie  2009.10.19  (0062)
     '       str = "select * from IC_CompPN"
        Case "IC_SHEARPIN"
            str = "select * from IC_ShearPin"
        Case "BRDCOMBINEQTY"
            str = "select * from QSMS_BrdCombineQty"
        Case "2NDSOURCE_ASSIGNPN"
            str = "select TOP 1 WO,MBPN,Rev,CompPN,AssignedCompPN,VendorCode,'N/Y' DeleteFlag from WO_AssignPN_Vendor"
        Case "UPLOAD_TRAYCOMPPN"  '2010.08.17 add by kaitlyn (1001)
            str = "select * from Tray_PN_BaseQty "
        Case "XL_MAXDIDMAINTAINQTY"
            str = "select * from XL_MaxDIDMaintainQty" ''(1009)
        Case "COMPONENT_DATA"   '2010.12.08 add by kaitlyn (1024)
            str = "select * from Component_data"
        Case "MACHINE_DATA"                          '(1129)
            str = "select * from Machine_data"
        Case "COMPPN_SPACER"                          '(1154)
            str = "select CompPN,Value,UID,TransdateTime from CompPN_BaseData WHERE TYPE='Spacer'"
        Case "AVLC"                          '(1154)
            str = "select * from QSMS_CustomerData"
        Case "A8_MANUAL"
            str = "select * from Asce_ManualUpload"
        Case "A8_DIDTYPE"
            str = "select * from A8_DIDType"
        Case Else
            MsgBox "Please check the right sheet name."
            Exit Sub
End Select
Set Rs = Conn.Execute(str)
If Not Rs.EOF Then
       Call CopyToExcel(Rs)
    Else
       MsgBox ("No Data"), vbCritical
End If
End Sub

Private Sub cmdFile_Click()
 CommonDialog1.ShowOpen
  txtFilePath = CommonDialog1.FileName
  cboSheetName.Clear
  Call ReadAllSheetName(txtFilePath)
  cboSheetName.Enabled = True
  cboSheetName.AddItem "ALL"
End Sub
Private Sub ReadAllSheetName(FilePath As String)
    On Error GoTo ERRHEAR
    Dim TempStr As String
    Dim i As Long
    Workbooks.Open FilePath
    Worksheets(1).Activate
    i = 0
    Do
       cboSheetName.AddItem ActiveSheet.Name
       ' TempDim(I) = TempStr
        ActiveSheet.Next.Select
        i = i + 1
    Loop
    cboSheetName.AddItem "ALL"
No_Data:
    'AllNum = I
    Workbooks.Close
    GoTo PASS
ERRHEAR:
    If Err.Number = 91 Then
        Resume No_Data
    End If
PASS:
End Sub
Private Sub UpLoad_QSMS_BOM(ByVal Machine As String, ByVal Jobpn As String, ByVal Version As String, ByVal Slot As String, ByVal COMPPN As String, ByVal Qty As String, ByVal LR As String, ByVal jobgroup As String, ByVal BuildType As String, ByVal Side As String, ByVal Factory As String, ByVal Line As String, location As String)
  Dim strSQL As String
  Dim Rs As ADODB.Recordset
  Dim errMsg As String
    '******************************
    '             If Len(Jobpn) <> 11 Then
    '                MsgBox "The JobPN:" & Jobpn & ",length must be 11,please check the JobPN!"
    '                Exit Sub
    '             End If
    '
    If Len(Version) <> 3 And Len(Version) <> 2 And Len(Version) <> 0 Then   '(1150)
       MsgBox "The Version:" & Version & ",length must be 2 or 3 or 0,please check the Version!"
       Exit Sub
    End If
    
    If InStr(UCase(Machine), "QOTHER") > 0 And UCase(Side) <> "Q" Then
        MsgBox "The Machine name is " & Machine & ",the Side must be Q!Please check!", vbCritical, "Message!"
        Exit Sub
    End If
    
    If InStr(UCase(Machine), "SOTHER") > 0 And BuildType = "3" Then
        MsgBox "The Machine name is " & Machine & ",the BuildType is 3!They are not match,Please check!", vbCritical, "Message!"
        Exit Sub
    End If
         
    If InStr(UCase(Machine), "COTHER") > 0 And (BuildType = "2" Or BuildType = "4") Then
        MsgBox "The Machine name is " & Machine & ",the BuildType is 2 or 4!They are not match,Please check!", vbCritical, "Message!"
        Exit Sub
    End If
     
     ''(1068)****************************
'      If Mid(UCase(machine), 2, 1) <> UCase(side) And Mid(UCase(machine), 2, 1) <> "W" And BuildType = "1" Then
'        MsgBox "The Machine name is " & machine & ",the BuildType is 1!They are not match,Please check!", vbCritical, "Message!"
'        Exit Sub
'     End If
'      If Mid(UCase(machine), 2, 1) = "W" And UCase(side) <> "Q" And BuildType = "1" Then
'        MsgBox "The Machine name is " & machine & ",the BuildType is 1!They are not match,Please check!", vbCritical, "Message!"
'        Exit Sub
'     End If
     ''***************************************
     If Trim(LR) = "" Then
         LR = "0"
     End If
     
     If UCase(Side) = "S" Or UCase(Side) = "C" Then
        If Settings.UpdateJobSide = "Y" Then
           strSQL = "select * from QSMS_JobSide where JobPN='" & Trim(Jobpn) & "'"
           Set Rs = Conn.Execute(strSQL)
           If Rs.EOF Then
              MsgBox Trim(Jobpn) & ":Can't find the job side by the JobPN,please check!", vbCritical, "ErrMessage"
              Exit Sub
           Else
              If UCase(Trim(Rs("side"))) <> "S" And UCase(Trim(Rs("side"))) <> "C" Then
                 MsgBox Trim(Rs("Side")) & ":Job side's format is wrong ,the side must be S or C,please define it afresh!", vbCritical, "ErrMessage"
                 Exit Sub
               Else
                 Side = Trim(Rs("Side"))
               End If
           End If
        End If
     End If
     
    If Trim(Factory) = "" Then  ''(0034)
        MsgBox "No Factory information in Excel,Please check!", vbCritical, "Message!"
        Exit Sub
    End If
    If Trim(Line) = "" Then  ''(1025)
        MsgBox "No Line information in Excel,Please check!", vbCritical, "Message!"
        Exit Sub
    End If
    
    If ChkMEBOM_Location = "Y" And Trim(location) = "" Then  ''''(1250)
        MsgBox "No Location information in Excel! Please check!", vbCritical, "Message!"
        Exit Sub
    End If
        
    If ChkValid(Slot, LR, Machine, Jobpn, jobgroup, COMPPN, BuildType, Side, errMsg) = False Then
        MsgBox errMsg
        Exit Sub
     End If
     
     '（0061）
     If (TempJObGroup = "" Or TempJObGroup <> jobgroup) Or (TempJobPn = "" Or TempJobPn <> Jobpn) Or (tempVersion <> Version) Or (tempmachine = "" Or tempmachine <> Machine) Or (TempLine = "" Or TempLine <> Line) Then
'     If (TempJObGroup = "" Or TempJObGroup <> jobgroup) Or (TempJobPn = "" Or TempJobPn <> Jobpn) Or (tempVersion = "" Or tempVersion <> Version)     Or (tempmachine = "" Or tempmachine <> Machine) Or (TempLine = "" Or TempLine <> Line) Then
        strSQL = "delete from QSMS_MEBom where Jobgroup='" & jobgroup & "' and Jobpn='" & Jobpn & "' and machine='" & Trim(Machine) & "' and Version='" & Version & "' and BuildType='" & Trim(BuildType) & "' and Factory='" & Trim(Factory) & "'and Line='" & Trim(Line) & "'" ''(0034)(1025)
        Conn.Execute strSQL
     End If
     
     If Qty <> "0" Then
            strSQL = "select * from QSMS_MEBom where Jobgroup='" & jobgroup & "' and Machine='" & Trim(Machine) & "' and JobPN='" & Trim(Jobpn) & "'  and Version='" & Version & "'  and Slot='" & Slot & "' and LR='" & LR & "' and BuildType='" & BuildType & "' and Factory='" & Trim(Factory) & "'and Line='" & Trim(Line) & "'"
            Set Rs = Conn.Execute(strSQL)
            If Rs.EOF Then
               strSQL = "Insert into QSMS_MEBom(Machine,JobPN,Version,CompPN,LR,Slot,Qty,JobGroup,BuildType,Side,UID,Factory,line,Location) " & _
                        " values('" & Trim(Machine) & "','" & Trim(Jobpn) & "','" & Trim(Version) & "', " & _
                        " '" & COMPPN & "','" & LR & "','" & Slot & "'," & Qty & ",'" & jobgroup & "','" & BuildType & "','" & Side & "','" & g_userName & "','" & Trim(Factory) & "','" & Trim(Line) & "','" & Trim(location) & "')" ''(0034)  (1025)
               Conn.Execute strSQL
               InsertBOM_Qty = InsertBOM_Qty + 1
            Else
              strSQL = "Update QSMS_MEBom set CompPN='" & Trim(COMPPN) & "',Qty=" & Qty & ",BuildType='" & Trim(BuildType) & "',Side='" & Trim(Side) & "',[UID]='" & g_userName & "',TransDateTime=convert(char(8),getdate(),112) + left(replace(convert(char(8),getdate(),108), ':', ''),6),Jobgroup='" & jobgroup & "',Factory='" & Trim(Factory) & "',Line='" & Trim(Line) & "',Location='" & Trim(location) & "' " & _
                       " where Machine='" & Trim(Machine) & "' and JobGroup='" & Trim(jobgroup) & "' and JobPN='" & Trim(Jobpn) & "' and Version='" & Version & "' and LR='" & LR & "' and Slot='" & Slot & "' and BuildType='" & BuildType & "' AND Factory='" & Trim(Factory) & "' AND Line='" & Trim(Line) & "'"   '(0057)
              Conn.Execute strSQL   ''(0034) (1025)
              UpdateBOM_Qty = UpdateBOM_Qty + 1
            End If
     End If
     tempmachine = Machine
     TempJobPn = Jobpn
     tempVersion = Version
     TempJObGroup = jobgroup
     TempFactory = Factory  ''(0034)
     TempLine = Line
     
     DoEvents
     DoEvents
     DoEvents
     DoEvents
     DoEvents

End Sub
Private Sub Load_QSMS_BOM(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim Factory, Line, Machine, Jobpn, Version, Slot, UpcompPN, COMPPN, Qty, LR, jobgroup As String, BuildType As String, Side As String
  Dim Total_Qty As Long
  Dim TMachine As String
  Dim strSQL As String
  Dim Rs As ADODB.Recordset
  Dim location As String
  Dim RE As String
  
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False

  rCount = 2
  Total_Qty = 0
  InsertBOM_Qty = 0
  UpdateBOM_Qty = 0
  tempmachine = ""
  TempJobPn = ""
  tempVersion = ""
  TempJObGroup = ""
  TempFactory = ""
  TempLine = ""
  
  With xlsBook.Worksheets(Trim(Shift_Item))
  
''''(1256)取消排序
'    ' Macro3 Macro
'    ' 宏由 Administrator 录制，时间: 2006/9/22
'    '*************** Sort MEBom start ***********************************
        If BU = "ESBU" Then
            .Columns("A:A").Select
            .Range("A1:M430").Sort Key1:=.Range("A1"), Order1:=xlAscending, Header:= _
                xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                SortMethod:=xlPinYin ', DataOption1:=xlSortNormal
            .Columns("B:B").Select
            .Range("A1:M430").Sort Key1:=.Range("B1"), Order1:=xlAscending, Header:= _
                xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                SortMethod:=xlPinYin ', DataOption1:=xlSortNormal
            .Columns("C:C").Select
            .Range("A1:M430").Sort Key1:=.Range("C1"), Order1:=xlAscending, Header:= _
                xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                SortMethod:=xlPinYin ', DataOption1:=xlSortNormal
            .Columns("H:H").Select
            .Range("A1:M430").Sort Key1:=.Range("H1"), Order1:=xlAscending, Header:= _
                xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                SortMethod:=xlPinYin ', DataOption1:=xlSortNormal
        '  '  **************** Sort MEBom end *************************************
        End If
      If BU = "NB5" Then
             '*************** Sort MEBom start ***********************************
             .Columns("A:A").Select
             .Range("A1:J430").Sort Key1:=.Range("A1"), Order1:=xlAscending, Header:= _
                 xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                 SortMethod:=xlPinYin ', DataOption1:=xlSortNormal
             .Columns("B:B").Select
             .Range("A1:J430").Sort Key1:=.Range("B1"), Order1:=xlAscending, Header:= _
                 xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                 SortMethod:=xlPinYin ', DataOption1:=xlSortNormal
            .Columns("C:C").Select
             .Range("A1:J430").Sort Key1:=.Range("C1"), Order1:=xlAscending, Header:= _
                 xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                 SortMethod:=xlPinYin ', DataOption1:=xlSortNormal
             .Columns("H:H").Select
             .Range("A1:J430").Sort Key1:=.Range("H1"), Order1:=xlAscending, Header:= _
                 xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
                 SortMethod:=xlPinYin ', DataOption1:=xlSortNormal
          '  **************** Sort MEBom end *************************************
       End If
          
       While Trim(.Cells(rCount, 1)) <> ""
             Machine = Trim(.Cells(rCount, 1) & vbNullString)
             Jobpn = Trim(.Cells(rCount, 2) & vbNullString)
             Version = Replace(Trim(.Cells(rCount, 3) & vbNullString), "'", " ")
             Slot = Replace(Trim(.Cells(rCount, 4) & vbNullString), " ", "")
             
             COMPPN = Trim(.Cells(rCount, 5) & vbNullString)
             Qty = Trim(.Cells(rCount, 6) & vbNullString)
             LR = Trim(.Cells(rCount, 7) & vbNullString)
             jobgroup = Trim(.Cells(rCount, 8) & vbNullString)
             BuildType = Trim(.Cells(rCount, 9) & vbNullString)
             Side = Trim(.Cells(rCount, 10) & vbNullString)
             Factory = Trim(.Cells(rCount, 11) & vbNullString)  ''(0034)
             Line = Trim(.Cells(rCount, 12) & vbNullString)
             location = Trim(.Cells(rCount, 13) & vbNullString)
             
            '******************************
            '****add by jeanson 2007/09/03
            strErrMessage = ""
            strErrMessage = FunPartNumberCheck(Jobpn)
            If strErrMessage <> "PASS" Then
                MsgBox strErrMessage
                Exit Sub
            End If
            If Slot = "0" Then ''0067
                MsgBox ("Line: " & rCount & " Slot=0; it is not allow define slot equal zero,please check it!")
            Else
              If CheckMachine(Line, Machine, Side) = False Then '(1032)
                  Exit Sub
              Else
              
            '------------superchai comments 20230919 (B)------------------
              '''''''''''''''''''20180704   Rain  对Location进行排序
            'If BU = "ESBU" Then
            '  strSQL = "exec location_Sort '" & location & "','" & Qty & "'" '''''1272
            '  Set Rs = Conn.Execute(strSQL)
            '  location = Trim(Rs("location"))
            '  RE = Trim(Rs("RE"))
            '
            '   If RE <> "PASS" Then
            '     MsgBox RE
            '   Exit Sub
            '   End If
            ' End If
            '------------superchai comments 20230919 (E)------------------
               
              ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Call UpLoad_QSMS_BOM(Machine, Jobpn, Version, Slot, COMPPN, Qty, LR, jobgroup, BuildType, Side, Factory, Line, location)
              ''(1068)***S***************mark by kaitlyn******************************
''                  If InStr(UCase(machine), "OTHER") > 0 Then
''                      If Left(machine, 1) = "*" Then ''0054
''    '                    TMachine = machine
''    '                    strSQL = "select distinct Line from Machine order by line"
''    '                    Set rs = Conn.Execute(strSQL)
''    '                    While Not rs.EOF
''    '                        If Trim(rs!Line) <> "" Then
''    '                            machine = Trim(rs!Line) + Mid(TMachine, 2, 10)
''    '                            If Not UCase(machine) Like "[A-Z][S,C,W,Q]OTHERS*" Then
''    '                                MsgBox "The Machine name is not correct:" & machine & vbCrLf & "Format should be [A-Z][S,C,W,Q]OTHERS*" & vbCrLf & "Row:" & rCount
''    '                                Exit Sub
''    '                            End If
''    '                            Call UpLoad_QSMS_BOM(machine, Jobpn, Version, Slot, COMPPN, Qty, LR, jobgroup, BuildType, Side, Factory)
''    '                        End If
''    '                        rs.MoveNext
''    '                    Wend
''                             Call UpLoad_QSMS_BOM(machine, Jobpn, Version, Slot, compPN, Qty, LR, jobgroup, BuildType, side, Factory, Line)  '(1025)
''                       Else
''                          If Not UCase(machine) Like "[A-Z,0-9][S,C,W,Q]OTHERS*" Then  '(1021)
''                              MsgBox "The Machine name is not correct:" & machine & vbCrLf & "Format should be [A-Z,0-9][S,C,W,Q]OTHERS*" & vbCrLf & "Row:" & rCount
''                              Exit Sub
''                          End If
''                         Call UpLoad_QSMS_BOM(machine, Jobpn, Version, Slot, compPN, Qty, LR, jobgroup, BuildType, side, Factory, Line)  '(1025)
''                       End If
''                  Else
''                    Call UpLoad_QSMS_BOM(machine, Jobpn, Version, Slot, compPN, Qty, LR, jobgroup, BuildType, side, Factory, Line)
''                  End If
            ''(1068)*****E**********************************************************
               End If
           End If
                rCount = rCount + 1
                Total_Qty = Total_Qty + 1
                Txt_RowCount = Total_Qty
      Wend
      
      If Left(Machine, 1) = "*" And InStr(UCase(Machine), "OTHER") > 0 Then '0064
        strSQL = "EXEC QSMS_MEBOMToAllLine"
        Conn.Execute (strSQL)
     End If
     
End With
strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_QSMS_BOM','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSQL)

 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
  
Total_Qty = InsertBOM_Qty + UpdateBOM_Qty
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & InsertBOM_Qty & vbCrLf & _
               "Update succeed : " & UpdateBOM_Qty & vbCrLf
              
End Sub
Private Function ChkValid(ByVal Slot As String, ByVal LR As String, Machine, ByVal Job As String, ByVal jobgroup As String, ByVal COMPPN As String, ByVal BuildType As String, ByVal Side As String, Msg As String) As Boolean
Dim Head, Tail As Long
Dim i As Long
ChkValid = True


''''1>check slot

'I = InStr(1, Slot, "-")
'If I > 0 Then
'   Head = Mid(Slot, 1, I - 1)
'   Tail = Mid(Slot, I + 1)
'   If IsNumeric(Head) = False Or IsNumeric(Tail) = False Then
'      ChkValid = False
'      Exit Function
'   End If
'   If CLng(Head) > 4 Or CLng(Tail) > 40 Then
'      ChkValid = False
'   End If
'
'Else 0045
If InStr(UCase(Machine), "OTHERS") <> 0 Then
    If IsNumeric(Slot) = False Then
          Msg = "Slot:" & Slot & " must be numeric!"
          ChkValid = False
          Exit Function
       End If
       If CLng(Slot) > 240 Then
          Msg = "Slot:" & Slot & " must <= 240!"
          ChkValid = False
       End If
    'End If
    ''''2>check LR
    'If IsNumeric(LR) = False Then
    '    ChkValid = False
    '    Exit Function
    'End If
    If CLng(LR) <> 0 Then
       Msg = "LR:" & LR & " must be 0 for DIP!"
       ChkValid = False
    End If
    '''''3>check machine
'    If Len(Trim(machine)) <> 6 Then
'       Msg = "Machine:" & machine & " lenght must be 6 or like OTHERS!"
'       ChkValid = False
'    End If
End If

''''(4)Chk JobGroup
If InStr(1, jobgroup, "-") > 0 And (Len(Trim(jobgroup)) = 15 Or Len(Trim(jobgroup)) = 14 Or Len(Trim(jobgroup)) = 12) Then   ' (1150)
Else
   Msg = "JobGroup:" & jobgroup & " format must be PN-REV, length is must be 15/14!"
   ChkValid = False
End If
'''(5)Chk JOb
'******************************
'****add by jeanson 2007/09/03
strErrMessage = ""
strErrMessage = FunPartNumberCheck(Job)
If strErrMessage <> "PASS" Then
    MsgBox strErrMessage
Exit Function
End If
'******************************
'If Len(Job) <> 11 Then
'    Msg = "Job:" & Job & " length must be 11!"
'    ChkValid = False
'End If
'(6)chk ComppN
'******************************
'****add by jeanson 2007/09/03
strErrMessage = ""
strErrMessage = FunPartNumberCheck(COMPPN)
If strErrMessage <> "PASS" Then
    MsgBox strErrMessage

Exit Function
End If
'******************************
'If Len(Trim(CompPN)) <> 11 Then
'    Msg = "CompPN:" & CompPN & " length must be 11!"
'    ChkValid = False
'End If

If Trim(BuildType) <> "1" And Trim(BuildType) <> "2" And Trim(BuildType) <> "3" And Trim(BuildType) <> "4" Then
    Msg = "BuildType:" & BuildType & " must be 1,2,3 or 4!"
    ChkValid = False
End If

If Trim(Side) <> "S" And Trim(Side) <> "C" And Trim(Side) <> "Q" And Trim(Side) <> "W" Then
    Msg = "Side:" & Side & " must be S,C or Q ,W!"
    ChkValid = False
End If
'Sub Macro1()
'
' Macro1 Macro
' 宏由 Administrator 录制，时间: 2006/9/22
'

'
'    Columns("A:A").Select
'    Range("A1:H9").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:= _
'        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
'        SortMethod:=xlPinYin, DataOption1:=xlSortNormal
'    Columns("B:B").Select
'    Range("A1:H9").Sort Key1:=Range("B1"), Order1:=xlAscending, Header:= _
'        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
'        SortMethod:=xlPinYin, DataOption1:=xlSortNormal
'    Columns("C:C").Select
'    Range("A1:H9").Sort Key1:=Range("C1"), Order1:=xlAscending, Header:= _
'        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
'        SortMethod:=xlPinYin, DataOption1:=xlSortNormal
'    Columns("H:H").Select
'    Range("A1:H9").Sort Key1:=Range("H1"), Order1:=xlAscending, Header:= _
'        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
'        SortMethod:=xlPinYin, DataOption1:=xlSortNormal
'End Sub

End Function

Private Sub Load_LostReplacePN(Shift_Item As String)    'ADD by Giant 2008/07/21 (0036)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rCount, Row_Count As Long
Dim WO, Jobpn, Item, COMPPN, Version As String
Dim TempJobPn, tempVersion As String
Dim Total_Qty, Update_Qty, Insert_Qty As Long
Dim transdatetime As String
Dim strSQL As String
Dim Rs As ADODB.Recordset
Dim strDelete As String

If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
   Exit Sub
End If
Set xlApp = CreateObject("Excel.Application")
Let xlApp.Visible = False
Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.DisplayAlerts = False

rCount = 2
Total_Qty = 0
Insert_Qty = 0
Update_Qty = 0
TempJobPn = ""

strSQL = "select getdate() as TransDateTime"
Set Rs = Conn.Execute(strSQL)
transdatetime = Format(Rs(0), "yyyymmddhhnnss")
With xlsBook.Worksheets(Trim(Shift_Item))
    While Trim(.Cells(rCount, 1)) <> ""
        WO = Trim(.Cells(rCount, 1) & vbNullString)
        Jobpn = Trim(.Cells(rCount, 2) & vbNullString)
        COMPPN = Trim(.Cells(rCount, 3) & vbNullString)
        Version = Trim(.Cells(rCount, 4) & vbNullString)
        
        If Len(Jobpn) > 20 Or Len(Version) > 20 Or Len(COMPPN) > 20 Or Len(COMPPN) < 11 Then
            MsgBox "Excel file format error,please check: ROW:" & rCount + 1
            Exit Sub
        End If
    
        strSQL = "select * from QSMS_LostReplacePN where wo='" & Trim(WO) & "' and CompPN='" & Trim(COMPPN) & "' and JobPN='" & Jobpn & "' and version='" & Version & "' "
        Set Rs = Conn.Execute(strSQL)
        If Rs.EOF Then
            strSQL = "Insert into QSMS_LostReplacePN(WO,JObPN,CompPN,version,GetFlag,Transdatetime) " & _
            " values('" & Trim(WO) & "','" & Trim(Jobpn) & "','" & Trim(COMPPN) & "','" & Version & "','N','" & transdatetime & "')"
            Conn.Execute strSQL
            Insert_Qty = Insert_Qty + 1
        Else
            ''strSQL = "Update QSMS_ReplacePN set GetFlag='N',Transdatetime='" & TransDateTime & "' where wo='" & Trim(WO) & "' and CompPN='" & Trim(CompPN) & "' and JobPN='" & Jobpn & "' and version='" & Version & "'"
            strSQL = "Update QSMS_LostReplacePN set GetFlag='N',Transdatetime='" & transdatetime & "' where wo='" & Trim(WO) & "' and CompPN='" & Trim(COMPPN) & "' and JobPN='" & Jobpn & "' and version='" & Version & "'"     '''''Salon 2008-08-05
            Conn.Execute strSQL
            Update_Qty = Update_Qty + 1
        End If
    
    TempJobPn = Jobpn
    tempVersion = Version
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    
    rCount = rCount + 1
    Total_Qty = Total_Qty + 1
    Txt_RowCount = Total_Qty
    Wend
End With
strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_LostReplacePN','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSQL)


xlsBook.Close
xlApp.Quit
Set xlApp = Nothing
Set xlsBook = Nothing
MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
             "Total Counter : " & Total_Qty & vbCrLf & _
             "Insert succeed : " & Insert_Qty & vbCrLf & _
             "Update succeed : " & Update_Qty & vbCrLf
             
strSQL = "EXEC QSMS_GetLostReplacePN"
Conn.Execute (strSQL)
End Sub

Private Sub Load_ReplacePN(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rCount, Row_Count As Long
Dim Jobpn, Item, COMPPN, Version As String
Dim TempJobPn, tempVersion As String
Dim Total_Qty, Update_Qty, Insert_Qty As Long
Dim transdatetime As String
Dim strSQL As String
Dim Rs As ADODB.Recordset
Dim strDelete As String

If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
   Exit Sub
End If
Set xlApp = CreateObject("Excel.Application")
Let xlApp.Visible = False
Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.DisplayAlerts = False

rCount = 2
Total_Qty = 0
Insert_Qty = 0
Update_Qty = 0
TempJobPn = ""
'(0025)--在同一个人上传时得到同一个时间
strSQL = "select getdate() as TransDateTime"
Set Rs = Conn.Execute(strSQL)
transdatetime = Format(Rs(0), "yyyymmddhhnnss")
With xlsBook.Worksheets(Trim(Shift_Item))

     While Trim(.Cells(rCount, 1)) <> ""
     
            Jobpn = Trim(.Cells(rCount, 1) & vbNullString)
            Version = Trim(.Cells(rCount, 2) & vbNullString)
            Item = Replace(Trim(.Cells(rCount, 3) & vbNullString), "'", " ")
            COMPPN = Trim(.Cells(rCount, 4) & vbNullString)
            strDelete = UCase(Trim(.Cells(rCount, 5)))
            
            If Len(Jobpn) > 20 Or Len(Version) > 20 Or Len(Item) > 10 Or Len(COMPPN) > 20 Or Len(COMPPN) < 11 Then
                MsgBox "Excel file format error,please check: ROW:" & rCount + 1
                Exit Sub
            End If

            ''''''update by Jing (0032)
            
            If strDelete = "Y" Then
               strSQL = "Delete from Qsms_ReplacePN where Jobpn='" & Jobpn & "' and version='" & Version & "' and CompPN='" & COMPPN & "'"
               Conn.Execute (strSQL)
            Else
'                If (TempJobPn = "" Or TempJobPn <> Jobpn) Or (tempVersion = "" Or tempVersion <> Version) Then
'                   strSQL = "delete from Qsms_Replacepn where Jobpn='" & Jobpn & "' and version='" & Version & "'"
'                   Conn.Execute strSQL
'                   Call InsertIntoQSMSLog("SMT_QSMS", "UpLoad Replace_PN", "Jobpn=" & Jobpn & " and version=" & Version & "")
'                End If
                
                strSQL = "select * from QSMS_ReplacePN where CompPN='" & Trim(COMPPN) & "' and JobPN='" & Jobpn & "' and version='" & Version & "' "
                Set Rs = Conn.Execute(strSQL)
                If Rs.EOF Then
                   strSQL = "Insert into QSMS_ReplacePN(JObPN,version,ID,CompPN,UID,Transdatetime) " & _
                            " values('" & Trim(Jobpn) & "','" & Version & "','" & Trim(Item) & "','" & Trim(COMPPN) & "','" & g_userName & "','" & transdatetime & "')"
                   Conn.Execute strSQL
                   Insert_Qty = Insert_Qty + 1
                Else
                  strSQL = "Update QSMS_ReplacePN set ID='" & Trim(Item) & "',UID='" & g_userName & "',Transdatetime='" & transdatetime & "' where CompPN='" & Trim(COMPPN) & "' and JobPN='" & Jobpn & "' and version='" & Version & "' "
                  Conn.Execute strSQL
                  Update_Qty = Update_Qty + 1
                End If
                
                TempJobPn = Jobpn
                tempVersion = Version
                DoEvents
                DoEvents
                DoEvents
                DoEvents
                DoEvents
            End If
            
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
    Wend
End With
strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_ReplacePN','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSQL)


xlsBook.Close
xlApp.Quit
Set xlApp = Nothing
Set xlsBook = Nothing
MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
             "Total Counter : " & Total_Qty & vbCrLf & _
             "Insert succeed : " & Insert_Qty & vbCrLf & _
             "Update succeed : " & Update_Qty & vbCrLf
'===========(0025)根据上传人和时间得到可能影响到的工单
strSQL = "select work_order,'The replace PN upload just now maybe impact this wo!Please re-check bom this wo by manual!' as WarnningDesc " & _
         "from qsms_wogroup where work_order in (select work_order from qsms_jobbom a,sap_wo_list b,qsms_replacepn c " & _
         "where a.work_order=b.wo and a.jobpn=c.jobpn and b.mb_rev=c.version and c.uid='" & g_userName & "' " & _
         "and c.transdatetime='" & transdatetime & "' and b.inputdt<>'' and b.wo_finishqcdatetime='') and closedflag<>'Y'"
Set Rs = Conn.Execute(strSQL)
If Rs.EOF = False Then
   Call CopyToExcel(Rs)
End If
strSQL = "insert into qsms_error_log(Appname,SubFunction,SubID,Col1,DetailDesc,TransDateTime) " & _
         "select 'SMT_QSMS','ReplacePN',work_order,'','The replace PN upload just now maybe impact this wo!" & _
         "Please check bom by manual!','" & transdatetime & "' from qsms_wogroup where work_order in " & _
         "(select work_order from qsms_jobbom a,sap_wo_list b,qsms_replacepn c " & _
         "where a.work_order=b.wo and a.jobpn=c.jobpn and b.mb_rev=c.version and c.uid='" & g_userName & "' " & _
         "and c.transdatetime='" & transdatetime & "' and b.inputdt<>'' and b.wo_finishqcdatetime='') and closedflag<>'Y'"
Conn.Execute (strSQL)
End Sub

Private Sub Load_MachineType(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rCount, Row_Count As Long
Dim Vendor  As String, Factory  As String, Machine As String, Unit As String, Qty As String, MaxSlotNum, LR As String, FujiData As String, Line As String, DeletedFlag As String
Dim TempJobPn, tempVersion, TempLine As String
Dim Total_Qty, Update_Qty, Insert_Qty As Long
Dim i As Integer
Dim SearchChar1, MyPos1
Dim MappingID As String
Dim DIOCircuit As String
Dim M1 As String, M2 As String, N1 As String
Dim N2 As Integer, N3 As Integer, N4 As Integer
Dim strSQL As String
Dim Rs As ADODB.Recordset

If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
   Exit Sub
End If
Set xlApp = CreateObject("Excel.Application")
Let xlApp.Visible = False
Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.DisplayAlerts = False


'two phase, first phase is to check format, 2nd phase is insert data
'For i = 1 To 2
    rCount = 2
    Total_Qty = 0
    Insert_Qty = 0
    Update_Qty = 0
    Txt_RowCount = 0
    TempJobPn = ""
    SearchChar1 = "-"
    
    'Phase2, delete old data
'    If i = 2 Then
'        strSql = "Truncate table Machine"
'        Conn.Execute (strSql)
'    End If
        
    With xlsBook.Worksheets(Trim(Shift_Item))
    
          While Trim(.Cells(rCount, 1)) <> ""
               Vendor = Trim(.Cells(rCount, 1) & vbNullString)
               Line = Trim(.Cells(rCount, 2) & vbNullString)
               Factory = Trim(.Cells(rCount, 3) & vbNullString)
               Machine = Replace(Trim(.Cells(rCount, 4) & vbNullString), "'", " ")
               Unit = Trim(.Cells(rCount, 5) & vbNullString)
               Qty = Trim(.Cells(rCount, 6) & vbNullString)
               MaxSlotNum = Trim(.Cells(rCount, 7) & vbNullString)
               LR = Trim(.Cells(rCount, 8) & vbNullString)
               MappingID = Trim(.Cells(rCount, 9) & vbNullString)
               FujiData = Trim(.Cells(rCount, 10) & vbNullString)
               DIOCircuit = Trim(.Cells(rCount, 11) & vbNullString)
               DeletedFlag = Trim(.Cells(rCount, 12) & vbNullString)
                '*********************JUDGE machine name Begin add by giant 061110*************************************************
                If InStr(UCase(Machine), "NXT") = 0 Then
                    If InStr(UCase(Machine), "OTHERS") = 0 Then
                        'Non-NXT machine name is like ASCP7A, ASCP7B, ...
                        If Not UCase(Machine) Like "[A-Z,0-9][S,C]???[A-Z]" Then  '(1021)
                            MsgBox "The Machine name is not correct:" & Machine & vbCrLf & "Format should be [A-Z,0-9][S,C]???[A-Z]" & vbCrLf & "Row:" & rCount
                            Exit Sub
                        End If
                    ElseIf InStr(UCase(Machine), "OTHERS") > 0 Then
                        'Non-NXT machine name is like ASOthers, ACOthers, ...
                        If Not UCase(Machine) Like "[A-Z,0-9][S,C,Q,W]OTHERS*" Then  '(1021)
                            MsgBox "The Machine name is not correct:" & Machine & vbCrLf & "Format should be [A-Z,0-9][S,C,Q,W]OTHERS*" & vbCrLf & "Row:" & rCount
                            Exit Sub
                        End If
                    End If
                ElseIf InStr(UCase(Machine), "NXT") > 0 Then
                    'NXT machine name is like ASNXTA01, ASNXTA02, ...
                    If Not UCase(Machine) Like "[A-Z,0-9][S,C]NXT[A-Z][0-9][0-9]" Then '(1021)
                        MsgBox "The Machine name is not correct:" & Machine & vbCrLf & "NXT format should be [A-Z,0-9][S,C]NXT[A-Z][0-9][0-9]" & vbCrLf & "Row:" & rCount
                        Exit Sub
                    End If
                End If
                '*********************JUDGE format Begin *********************************************************
               If Vendor = "Fuji" Then
                    MyPos1 = InStr(MappingID, SearchChar1)
                    If MyPos1 = 0 Then
                        MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
                        Exit Sub
                    Else
                        N3 = Mid(MappingID, MyPos1 + 1)
                        N4 = Mid(MappingID, 1, MyPos1 - 1)
                        If N3 > 20 Then
                            MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
                            Exit Sub
                        End If
                        If N4 < 1 Or N4 > 5 Then
                            MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
                            Exit Sub
                        End If
                    End If
               Else
                    M1 = Left(Machine, 2)
                    M2 = Left(MappingID, 2)
                    N1 = Mid(MappingID, 3, 2)
                    If IsNumeric(Right(MappingID, 5)) = True Then
                        N2 = CInt(Right(MappingID, 5))
                    End If
                    
                    If N2 < 1 Or N2 > 9 Then
                        MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
                        Exit Sub
                    End If
                    
                    If M1 <> M2 Or N1 <> "MC" Then
                        MsgBox ("Wrong MappingID :" & MappingID & ", Row:" & rCount)
                        Exit Sub
                    End If
               End If
               
                    If IsNumeric(Mid(FujiData, 1)) = True Then
                        N2 = CInt(Mid(FujiData, 1))
                        If N2 > 9 Then
                            MsgBox ("Wrong FujiData :" & FujiData & ", Row:" & rCount)
                            Exit Sub
                        End If
                    Else
                            MsgBox ("Wrong FujiData :" & FujiData & ", Row:" & rCount)
                            Exit Sub
                    End If
                    
                    If IsNumeric(Mid(DIOCircuit, 1)) = True Then
                        N2 = CInt(Mid(DIOCircuit, 1))
                        If N2 > 9 Then
                            MsgBox ("Wrong DIOCircuit :" & DIOCircuit & ", Row:" & rCount)
                            Exit Sub
                        End If
                    Else
                            MsgBox ("Wrong DIOCircuit :" & DIOCircuit & ", Row:" & rCount)
                            Exit Sub
                    End If
                '*********************Delecte machine information by Line ***************************************************
                If UCase(DeletedFlag) = "Y" Then
                   strSQL = "delete from Machine where Machine='" & Trim(Machine) & "' "
                   Conn.Execute strSQL
                   Deleted_Qty = Deleted_Qty + 1
                Else
                '*********************insert or update machine information ***************************************************
                   strSQL = "select * from Machine where Machine='" & Trim(Machine) & "'"
                   Set Rs = Conn.Execute(strSQL)
                   If Rs.EOF Then
                       strSQL = "Insert into Machine(Vendor,Factory,Machine,Unit,Qty,MaxSlotNum,LR,MappingID,FujiData,DIOCircuit,OPID) " & _
                                   " values('" & Trim(Vendor) & "','" & Factory & "','" & Trim(Machine) & "','" & Trim(Unit) & "','" & Trim(Qty) & "','" & Trim(MaxSlotNum) & "','" & Trim(LR) & "','" & Trim(MappingID) & "','" & Trim(FujiData) & "','" & Trim(DIOCircuit) & "','" & Trim(g_userName) & "')"
                       Conn.Execute strSQL
                       Insert_Qty = Insert_Qty + 1
                   Else
                       strSQL = "Update Machine set Vendor='" & Trim(Vendor) & "',Factory='" & Factory & "',Unit='" & Trim(Unit) & "',Qty='" & Trim(Qty) & "',MaxSlotNum='" & Trim(MaxSlotNum) & "',LR='" & Trim(LR) & "',MappingID='" & Trim(MappingID) & "',FujiData='" & Trim(FujiData) & "',DIOCircuit='" & Trim(DIOCircuit) & "',OPID='" & Trim(g_userName) & "' where Machine='" & Trim(Machine) & "'"
                       Conn.Execute strSQL
                       Update_Qty = Update_Qty + 1
                   End If
                End If
                    rCount = rCount + 1
                    Total_Qty = Total_Qty + 1
                    Txt_RowCount = Total_Qty

        Wend
    End With
strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_MachineType','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSQL)

xlsBook.Close
xlApp.Quit
Set xlApp = Nothing
Set xlsBook = Nothing
Conn.Execute ("exec GenUpdateMachineType")     '*******************add by jeanson 10/10*******

MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
             "Total Counter : " & Total_Qty & vbCrLf & _
             "Insert succeed : " & Insert_Qty & vbCrLf & _
             "Update succeed : " & Update_Qty & vbCrLf
              
End Sub

Private Sub Load_DID(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim DID, COMPPN, Qty, RemainQty, VendorCode, DateCode, LotCode As String
  Dim transdatetime As String
  Dim Total_Qty, Update_Qty, Insert_Qty As Long
 
  Dim str As String
  Dim Rs As ADODB.Recordset
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False

  rCount = 2
  Total_Qty = 0
  Insert_Qty = 0
  Update_Qty = 0
  str = "Select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Format(Rs(0), "YYYYMMDDHHMMSS")
  With xlsBook.Worksheets(Trim(Shift_Item))
 
       While Trim(.Cells(rCount, 1)) <> ""
           
             DID = Trim(.Cells(rCount, 1) & vbNullString)
             COMPPN = Replace(Trim(.Cells(rCount, 3) & vbNullString), "'", " ")
        
             RemainQty = Trim(.Cells(rCount, 6) & vbNullString)
             Qty = RemainQty
             VendorCode = Trim(.Cells(rCount, 11) & vbNullString)
             LotCode = Trim(.Cells(rCount, 12) & vbNullString)
             DateCode = Trim(.Cells(rCount, 13) & vbNullString)
             If Qty = "" Then Qty = 0
             If RemainQty = "" Then RemainQty = 0
             str = "select * from QSMS_DID where DID='" & Trim(DID) & "'  "
             
             Set Rs = Conn.Execute(str)
             If Rs.EOF Then
                str = "Insert into QSMS_DID(DID,CompPN,Qty,RemainQty,VendorCode,DateCode,LotCode,UID,TransDateTime,UsedFlag) " & _
                         " values('" & Trim(DID) & "','" & Trim(COMPPN) & "'," & Trim(Qty) & "," & RemainQty & ",'" & VendorCode & "','" & DateCode & "','" & LotCode & "','" & g_userName & "','" & transdatetime & "','N')"
                Conn.Execute str
                Insert_Qty = Insert_Qty + 1
             Else
               str = "Update QSMS_DID set VendorCode='" & Trim(VendorCode) & "',Datecode='" & DateCode & "',LotCode='" & LotCode & "',Qty=" & Qty & ",RemainQty=" & RemainQty & ",TransDateTime='" & transdatetime & "'" & _
                        ",UsedFlag='N',UID='" & g_userName & "' where DID='" & Trim(DID) & "' "
               Conn.Execute str
               Update_Qty = Update_Qty + 1
               
             End If
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
      Wend
End With
 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Update succeed : " & Update_Qty & vbCrLf
              
End Sub
Private Sub Load_NonAVL(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim Work_Order, COMPPN, VendorCode, DateCode, LotCode, Customer, Model As String, DeletedFlag As String
  Dim Total_Qty, Deleted_Qty, Insert_Qty As Long
  Dim str As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  rCount = 2
  Total_Qty = 0

  Insert_Qty = 0
  Deleted_Qty = 0
  str = "select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Rs.Fields(0) 'Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
  
  With xlsBook.Worksheets(Trim(Shift_Item))
       
       While Trim(.Cells(rCount, 2)) <> ""
             
             Work_Order = Trim(.Cells(rCount, 1) & vbNullString)
             COMPPN = Trim(.Cells(rCount, 2) & vbNullString)
             VendorCode = Replace(Trim(.Cells(rCount, 3) & vbNullString), "'", " ")
             DateCode = Replace(Trim(.Cells(rCount, 4) & vbNullString), "'", " ")
             LotCode = Replace(Trim(.Cells(rCount, 5) & vbNullString), "'", " ")
             Customer = Trim(.Cells(rCount, 6) & vbNullString)
             Model = Trim(.Cells(rCount, 7) & vbNullString)
             DeletedFlag = Trim(.Cells(rCount, 8) & vbNullString)
             
             If Len(COMPPN) > 20 Or Len(VendorCode) > 50 Or Len(Customer) > 20 Or Len(Model) > 10 Then
                MsgBox "Excel file format error,please check: ROW:" & rCount + 1
                Exit Sub
             End If
             If UCase(DeletedFlag) = "Y" Then
                   str = "delete from QSMS_NonAVL where CompPN='" & Trim(COMPPN) & "' and VendorCode='" & VendorCode & "' and DateCode='" & DateCode & "' " & _
                         "and LotCode='" & LotCode & "' and Customer='" & Customer & "' and Model='" & Model & "' and Work_Order='" & Work_Order & "' "
                   Conn.Execute str
                   Deleted_Qty = Deleted_Qty + 1
             Else
                   str = "select * from QSMS_NonAVL where CompPN='" & Trim(COMPPN) & "' and VendorCode='" & VendorCode & "' and DateCode='" & DateCode & "' " & _
                         "and LotCode='" & LotCode & "' and Customer='" & Customer & "' and Work_Order='" & Work_Order & "'"
                   Set Rs = Conn.Execute(str)
                   If Rs.EOF Then
                       str = "Insert into QSMS_NonAVL(Work_Order,CompPN,VendorCode,DateCode,LotCOde,Customer,Model,TransDateTime) " & _
                         " values('" & Work_Order & "','" & Trim(COMPPN) & "','" & Trim(VendorCode) & "','" & DateCode & "','" & LotCode & "','" & Trim(Customer) & "','" & Model & "','" & transdatetime & "')"
                       Conn.Execute str
                      Insert_Qty = Insert_Qty + 1
                   End If
             End If
             
               
             
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
      Wend
End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_NonAVL','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)
 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Deleted succeed : " & Deleted_Qty & vbCrLf
              
End Sub

Private Sub Load_AVL(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount As Long
  Dim aryModel, aryVcode As Variant
  Dim i, j As Integer
  Dim COMPPN, Desc1, VendorCode, VendorCode1, VendorCode2, DateCode, LotCode, Customer, Model, Model1, Model2 As String, DeletedFlag As String
  Dim str, TempStr As String, delflag As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  Total_Qty = 0
  Insert_Qty = 0
  Deleted_Qty = 0
  str = "select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
  
'///////////////////////////////////////for PAL////////////////////////////////////////////////////
  rCount = 4
  With xlsBook.Worksheets(Trim(Shift_Item))
     Customer = Trim(Replace(Replace(.Cells(1, 2) & vbNullString, vbCr, ""), vbLf, "")) ' Trim(.Cells(1, 2) & vbNullString)
     Model1 = Trim(Replace(Replace(.Cells(2, 2) & vbNullString, vbCr, ""), vbLf, ""))  'Trim(.Cells(2, 2) & vbNullString)
     
        aryModel = Split(Model1, ";")
        For j = 0 To UBound(aryModel)
            Model2 = aryModel(j)
            If Len(Model2) = 3 Then
                Model = Model2
'*********************check model if defined*********************************
             str = "select * from modelname where substring(modelname,3,3)=" & sq(Model)
             Set Rs = Conn.Execute(str)
             If Rs.EOF Then
                MsgBox "This ModelName not defined !ROW:" & rCount
                Exit Sub
             End If
'*********************check model if defined*********************************
            End If
        
            While Trim(.Cells(rCount, 1)) <> ""
                COMPPN = Trim(Replace(Replace(.Cells(rCount, 1) & vbNullString, vbCr, ""), vbLf, ""))
                VendorCode1 = Trim(Replace(Replace(.Cells(rCount, 2) & vbNullString, vbCr, ""), vbLf, "")) ' Replace(Trim(.Cells(rCount, 2) & vbNullString), "'", " ")
                delflag = Trim(Replace(Replace(.Cells(rCount, 3) & vbNullString, vbCr, ""), vbLf, ""))
                If Len(Trim(VendorCode1)) < 4 Then
                    VendorCode = Trim(VendorCode1)
                    Call InsertAVL(Trim(COMPPN), Trim(VendorCode), Trim(Customer), Trim(Model), Trim(Desc1), Trim(rCount), Trim(transdatetime), delflag)
                Else
                    aryVcode = Split(VendorCode1, ";")
                    For i = 0 To UBound(aryVcode)
                        VendorCode2 = aryVcode(i)
                        If Len(Trim(VendorCode2)) = 3 Then
                            VendorCode = VendorCode2
                        End If
                        Call InsertAVL(Trim(COMPPN), Trim(VendorCode), Trim(Customer), Trim(Model), Trim(Desc1), Trim(rCount), Trim(transdatetime), delflag)
                    Next i
                End If
                rCount = rCount + 1
                DoEvents
            Wend
            rCount = 4
        Next j
End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_AVL','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)

xlsBook.Close
xlApp.Quit
Set xlApp = Nothing
Set xlsBook = Nothing

 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Deleted succeed : " & Deleted_Qty & vbCrLf
              
End Sub


Private Sub Load_UnChkComp(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim CompHead, DeletedFlag As String
  Dim Total_Qty, Deleted_Qty, Insert_Qty As Long
   Dim str As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  rCount = 2
  Total_Qty = 0

  Insert_Qty = 0
  Deleted_Qty = 0
  str = "select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
  
  With xlsBook.Worksheets(Trim(Shift_Item))
       
       While Trim(.Cells(rCount, 1)) <> ""
           
             CompHead = Trim(.Cells(rCount, 1) & vbNullString)
             DeletedFlag = Trim(.Cells(rCount, 2) & vbNullString)
             If Len(CompHead) > 11 Then
                MsgBox "Excel file format errot,Please check: row :" & rCount + 1
                Exit Sub
             End If
           
             Select Case UCase(Mid(CompHead, 1, 2))
                    Case "DA", "HC", "XX", "XY", "ZZ", "HF", "HA", "HE", "GA", "JX", "SA", "QJ"   ''(0020)
                    Case Else
                          MsgBox "Comp can not be upload,Please check"
                          Exit Sub
             End Select
             
             If UCase(DeletedFlag) = "Y" Then
                   str = "delete from QSMS_UnChkComp where CompHead='" & Trim(CompHead) & "' "
                   Conn.Execute str
                   Deleted_Qty = Deleted_Qty + 1
             Else
                   str = "select * from QSMS_UnChkComp where CompHead='" & Trim(CompHead) & "' "
                   Set Rs = Conn.Execute(str)
                   If Rs.EOF Then
                       str = "Insert into QSMS_UnChkComp(CompHead,TransDateTime,UID) " & _
                         " values('" & Trim(CompHead) & "','" & transdatetime & "','" & g_userName & "')"
                       Conn.Execute str
                      Insert_Qty = Insert_Qty + 1
                   End If
             End If
             
               
             
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
      Wend
End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_UnChkComp','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)


 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Deleted succeed : " & Deleted_Qty & vbCrLf
              
End Sub

Private Sub Load_DocumentComp(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim MBPN, Version, oldCompPN, NewCompPN, BeginDate, EndDate, DeletedFlag As String
  Dim Qty As Long
  Dim Total_Qty, Update_Qty, Insert_Qty, Delete_Qty As Long
  Dim str As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String
  
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  rCount = 2
  Total_Qty = 0

  Insert_Qty = 0
  Update_Qty = 0
  Delete_Qty = 0
  str = "select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
  
  With xlsBook.Worksheets(Trim(Shift_Item))
       
       While Trim(.Cells(rCount, 1)) <> ""
           
             MBPN = Trim(.Cells(rCount, 1) & vbNullString)
             Version = Trim(.Cells(rCount, 2) & vbNullString)
             oldCompPN = Trim(.Cells(rCount, 3) & vbNullString)
             NewCompPN = Trim(.Cells(rCount, 4) & vbNullString)
             Qty = CLng(Trim(.Cells(rCount, 5)))
             BeginDate = Trim(.Cells(rCount, 6) & vbNullString)
             EndDate = Trim(.Cells(rCount, 7) & vbNullString)
             DeletedFlag = Trim(.Cells(rCount, 8) & vbNullString)
             
             If Len(MBPN) > 13 Or Len(Version) > 5 Or Len(oldCompPN) > 11 Or Len(NewCompPN) > 11 Or Len(BeginDate) > 8 Or Len(EndDate) > 8 Then
                MsgBox "Excel file format error,please check: ROW:" & rCount + 1
                Exit Sub
             End If
             If BeginDate = "" Or Len(BeginDate) <> 8 Or IsNumeric(BeginDate) = False Then
                BeginDate = Mid(transdatetime, 1, 8)
                EndDate = Format(DateAdd("d", 14, Date), "YYYYMMDD")
                
             Else
                EndDate = Format(DateAdd("d", 14, Date), "YYYYMMDD")
             End If
             If UCase(Trim(DeletedFlag)) = "Y" Then
                    str = "delete from QSMS_DocuComp where  JobPN='" & Trim(MBPN) & "'   and OldCompPN='" & oldCompPN & "' and Version='" & Version & "' and NewCompPN='" & NewCompPN & "'" & _
                          " and begindate <='" & Mid(transdatetime, 1, 8) & "' and EndDate>='" & Mid(transdatetime, 1, 8) & "'"
                    Conn.Execute str
                    Delete_Qty = Delete_Qty + 1
            Else
                     If NewCompPN = "" Then  'the old component was forbided
                            str = "select * from QSMS_DocuComp where JObPN='" & Trim(MBPN) & "'   and OldCompPN='" & oldCompPN & "' and version='" & Version & "' "
                            Set Rs = Conn.Execute(str)
                            If Rs.EOF Then
                                str = "Insert into QSMS_DocuComp(JobPN,Version,OldCompPN,NewCompPN,BeginDate,EndDate,TransDateTime,UID) " & _
                                 " values('" & Trim(MBPN) & "','" & Version & "','" & oldCompPN & "','" & NewCompPN & "','" & BeginDate & "','" & EndDate & "','" & transdatetime & "','" & g_userName & "')"
                                Conn.Execute str
                                Insert_Qty = Insert_Qty + 1
                            End If
                        
                     Else
                            str = "select * from QSMS_DocuComp where JobpN='" & Trim(MBPN) & "'   and newcomppn='" & NewCompPN & "' and version='" & Version & "'"
                            
                            Set Rs = Conn.Execute(str)
                            If Rs.EOF Then
                                str = "Insert into QSMS_DocuComp(JobPN,Version,OldCompPN,NewCompPN,BeginDate,EndDate,TransDateTime,UID) " & _
                                  " values('" & Trim(MBPN) & "','" & Version & "','" & oldCompPN & "','" & NewCompPN & "','" & BeginDate & "','" & EndDate & "','" & transdatetime & "','" & g_userName & "')"
                                Conn.Execute str
                                Insert_Qty = Insert_Qty + 1
                            Else
                                str = "Update QSMS_DocuComp set OldComppN='" & oldCompPN & "',transdatetime='" & transdatetime & "' where JobPN='" & Trim(MBPN) & "' and Version='" & Version & "'  and newcomppn='" & NewCompPN & "' "
                                Conn.Execute str
                                Update_Qty = Update_Qty + 1
                            End If
                    
                     End If
            End If
               
             
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
      Wend
End With
 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
                "Delete succeed : " & Delete_Qty & vbCrLf & _
               "Update succeed : " & Update_Qty & vbCrLf
              
End Sub

Private Sub Upload_CTO_Model(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim Model, PN, TempModel As String
  Dim Total_Qty, Deleted_Qty, Insert_Qty, Update_Qty As Long
   Dim str As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  rCount = 2
  Total_Qty = 0
  Update_Qty = 0
  Insert_Qty = 0
  Deleted_Qty = 0
  str = "select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
  
  With xlsBook.Worksheets(Trim(Shift_Item))
       
       While Trim(.Cells(rCount, 1)) <> ""
           
             Model = Trim(.Cells(rCount, 1) & vbNullString)
             PN = Trim(.Cells(rCount, 2) & vbNullString)
             If TempModel = "" Or TempModel <> Model Then
'                Strsql = "delete from CTO_Model where PN='" & PN & "' and Model='" & Model & "' "
                str = "delete from CTO_Model where  Model='" & Model & "' "
                Conn.Execute str
             End If
            str = "select * from CTO_Model where  Model='" & Trim(Model) & "' and pN='" & PN & "'"
            Set Rs = Conn.Execute(str)
            If Rs.EOF Then
                str = "Insert into CTO_Model(model,PN,OPID,TransDateTime) " & _
                  " values('" & Trim(Model) & "','" & PN & "','" & g_userName & "','" & transdatetime & "')"
                Conn.Execute str
               Insert_Qty = Insert_Qty + 1
            Else
                str = "Update CTO_Model set Model='" & Trim(Model) & "',transDateTime='" & transdatetime & "',OPID='" & g_userName & "' where   pN='" & PN & "' "
                Conn.Execute str
                Update_Qty = Update_Qty + 1
            End If
            TempModel = Model
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
      Wend
End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','upload_CTOModel','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)


 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Update succeed : " & Update_Qty & vbCrLf & _
               "Deleted succeed : " & Deleted_Qty & vbCrLf
              
End Sub


Private Sub Load_UpdateJobPn(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim Model, COMPPN, SourceJobpn, DestJobpn, DeletedFlag As String
  Dim Total_Qty, Deleted_Qty, Insert_Qty, Update_Qty As Long
   Dim str As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  rCount = 2
  Total_Qty = 0
  Update_Qty = 0
  Insert_Qty = 0
  Deleted_Qty = 0
  str = "select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
  
  With xlsBook.Worksheets(Trim(Shift_Item))
       
       While Trim(.Cells(rCount, 1)) <> ""
           
             Model = Trim(.Cells(rCount, 1) & vbNullString)
             COMPPN = Trim(.Cells(rCount, 2) & vbNullString)
             SourceJobpn = Trim(.Cells(rCount, 3) & vbNullString)
             DestJobpn = Trim(.Cells(rCount, 4) & vbNullString)
             DeletedFlag = Trim(.Cells(rCount, 5) & vbNullString)
            '******************************
            '****add by jeanson 2007/09/03
            strErrMessage = ""
            strErrMessage = FunPartNumberCheck(COMPPN)
            If strErrMessage <> "PASS" Then
                MsgBox strErrMessage
            Exit Sub
            End If
            '******************************
'             If Len(Model) <> 3 Or Len(CompPN) <> 11 Or Len(SourceJobpn) <> 11 Or Len(DestJobpn) <> 11 Then
'             If Len(Model) <> 3 Or Len(CompPN) <> 11 Or Len(SourceJobpn) <> 11 Or Len(DestJobpn) <> 11 Then
'                MsgBox "Excel file format error,Please check: row :" & rCount + 1
'                Exit Sub
'             End If
             If UCase(DeletedFlag) = "Y" Then
                   str = "delete from QSMS_UpdateJobPN where Model='" & Trim(Model) & "' and ComppN='" & COMPPN & "' and SourceJobPN='" & SourceJobpn & "'"
                   Conn.Execute str
                   Deleted_Qty = Deleted_Qty + 1
             Else
                   str = "select * from QSMS_UpdateJobPN where  Model='" & Trim(Model) & "' and ComppN='" & COMPPN & "' and SourceJobPN='" & SourceJobpn & "'"
                   Set Rs = Conn.Execute(str)
                   If Rs.EOF Then
                       str = "Insert into QSMS_UpdateJobPN(model,CompPN,SourceJobpn,DestJobpn,UID,TransDateTime) " & _
                         " values('" & Trim(Model) & "','" & COMPPN & "','" & SourceJobpn & "' ,'" & DestJobpn & "','" & g_userName & "','" & transdatetime & "')"
                       Conn.Execute str
                      Insert_Qty = Insert_Qty + 1
                   Else
                       str = "Update QSMS_UpdateJobpn set DestJobPN='" & DestJobpn & "' where  Model='" & Trim(Model) & "' and ComppN='" & COMPPN & "' and SourceJobPN='" & SourceJobpn & "'"
                       Conn.Execute str
                       Update_Qty = Update_Qty + 1
                   End If
             End If
             
               
             
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
      Wend
End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_UpdateJobPn','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)


 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Update succeed : " & Update_Qty & vbCrLf & _
               "Deleted succeed : " & Deleted_Qty & vbCrLf
              
End Sub
Private Sub Load_FujiBrdSeqMapping(SheetName As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim Total_Qty, Update_Qty, Insert_Qty As Long
  Dim Jobpn As String, Rev As String, BrdSeq As String, arrBrdSeq, BrdPN As String, BrdRev As String
  Dim PNRev
  Dim i As Long
 
  Dim str As String
  Dim Rs As ADODB.Recordset
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  
  PNRev = Split(SheetName, "-")
  If UBound(PNRev) <> 1 Then
    MsgBox ("The sheet name must be PN-Rev!")
    Exit Sub
  End If
  
  Jobpn = Trim(PNRev(0))
  Rev = Trim(PNRev(1))
      '******************************
    '****add by jeanson 2007/09/03
    strErrMessage = ""
    strErrMessage = FunPartNumberCheck(Jobpn)
    If strErrMessage <> "PASS" Then
        MsgBox strErrMessage
    Exit Sub
    End If
    '******************************
'  If Len(Jobpn) <> 11 Then
'     MsgBox ("The JobPN:" & Jobpn & ",length must be 11,please check the JobPN!")
'     Exit Sub
'  End If
  If Len(Rev) <> 3 And Len(Rev) <> 2 Then
     MsgBox ("The Version:" & Rev & ",length must be 2 or 3,please check the Version!")
     Exit Sub
  End If
 
  'del old data
  str = "delete from FujiBrdSeqMapping where JobPN=" & sq(Jobpn) & " and Rev=" & sq(Rev)
  Conn.Execute (str)
  
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False

  rCount = 2
  Total_Qty = 0
  Insert_Qty = 0
  Update_Qty = 0
  With xlsBook.Worksheets(Trim(SheetName))
 
       While Trim(.Cells(rCount, 1)) <> ""
            BrdSeq = Trim(.Cells(rCount, 1) & vbNullString)
            PNRev = Split(Replace(Trim(.Cells(rCount, 2) & vbNullString), "'", " "), "-")
            If UBound(PNRev) <> 1 Then
                 MsgBox ("The PN-Rev:" & Trim(.Cells(rCount, 2)) & " format is not correct!")
                 Exit Sub
            End If
            BrdPN = Trim(PNRev(0))
            BrdRev = Trim(PNRev(1))
            '******************************
            '****add by jeanson 2007/09/03
            strErrMessage = ""
            strErrMessage = FunPartNumberCheck(Jobpn)
            If strErrMessage <> "PASS" Then
                MsgBox strErrMessage
            Exit Sub
            End If
            '******************************
'            If Len(Jobpn) <> 11 Then
'               MsgBox ("The BrdPN:" & BrdPN & ",length must be 11,please check the BrdPN!")
'               Exit Sub
'            End If
            If Len(BrdRev) <> 3 And Len(BrdRev) <> 2 Then
               MsgBox ("The BrdRev:" & Rev & ",length must be 2 or 3,please check the BrdRev!")
               Exit Sub
            End If
            arrBrdSeq = Split(BrdSeq, ",")
            For i = 0 To UBound(arrBrdSeq)
                If Trim(arrBrdSeq(i)) > "" Then
                    If Not IsNumeric(arrBrdSeq(i)) Then
                        MsgBox ("The BrdSeq:" & arrBrdSeq(i) & " must be numeric!")
                        Exit Sub
                    End If
                     If Len(Jobpn) > 11 Or Len(Rev) > 10 Or Len(Trim(arrBrdSeq(i))) > 10 Or Len(BrdPN) > 11 Or Len(BrdRev) > 11 Then
                        MsgBox "Excel file format error,Please check Row:" & rCount + 1
                        Exit Sub
                    End If
                    str = "insert FujiBrdSeqMapping (JobPN, Rev, BrdSeq, BrdPN, BrdRev) values (" & _
                        sq(Jobpn) & "," & sq(Rev) & "," & sq(Trim(arrBrdSeq(i))) & "," & sq(BrdPN) & "," & sq(BrdRev) & ")"
                    Conn.Execute str
                    Insert_Qty = Insert_Qty + 1
                End If
                DoEvents
            Next i
             
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
      Wend
 End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_FujiBrdSeqMappi','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)

 
 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & SheetName & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Update succeed : " & Update_Qty & vbCrLf
              
End Sub
Private Sub Load_PhilipsBrdSeqMapping(SheetName As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim Total_Qty, Update_Qty, Insert_Qty As Long
  Dim Jobpn As String, Rev As String, BrdSeq As String, arrBrdSeq, BrdPN As String, BrdRev As String
  Dim PNRev
  Dim i As Long
 
  Dim str As String
  Dim Rs As ADODB.Recordset
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  
  PNRev = Split(SheetName, "-")
  If UBound(PNRev) <> 1 Then
    MsgBox ("The sheet name must be PN-Rev!")
    Exit Sub
  End If
  
  Jobpn = Trim(PNRev(0))
  Rev = Trim(PNRev(1))
  '******************************
'****add by jeanson 2007/09/03
strErrMessage = ""
strErrMessage = FunPartNumberCheck(Jobpn)
If strErrMessage <> "PASS" Then
    MsgBox strErrMessage
    
Exit Sub
End If
'******************************
'  If Len(Jobpn) <> 11 Then
'     MsgBox ("The JobPN:" & Jobpn & ",length must be 11,please check the JobPN!")
'     Exit Sub
'  End If
  If Len(Rev) <> 3 And Len(Rev) <> 2 Then
     MsgBox ("The Version:" & Rev & ",length must be 2 or 3,please check the Version!")
     Exit Sub
  End If
 
  'del old data
  str = "delete from PhilipsBrdSeqMapping where JobPN=" & sq(Jobpn) & " and Rev=" & sq(Rev)
  Conn.Execute (str)
  
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False

  rCount = 2
  Total_Qty = 0
  Insert_Qty = 0
  Update_Qty = 0
  With xlsBook.Worksheets(Trim(SheetName))
 
       While Trim(.Cells(rCount, 1)) <> ""
            BrdSeq = Trim(.Cells(rCount, 1) & vbNullString)
            PNRev = Split(Replace(Trim(.Cells(rCount, 2) & vbNullString), "'", " "), "-")
            If UBound(PNRev) <> 1 Then
                 MsgBox ("The PNRev:" & PNRev & " format is not correct!")
                 Exit Sub
            End If
            BrdPN = Trim(PNRev(0))
            BrdRev = Trim(PNRev(1))
            '******************************
            '****add by jeanson 2007/09/03
            strErrMessage = ""
            strErrMessage = FunPartNumberCheck(Jobpn)
            If strErrMessage <> "PASS" Then
                MsgBox strErrMessage

            Exit Sub
            End If
            '******************************
'            If Len(Jobpn) <> 11 Then
'               MsgBox ("The BrdPN:" & BrdPN & ",length must be 11,please check the BrdPN!")
'               Exit Sub
'            End If
            If Len(BrdRev) <> 3 And Len(BrdRev) <> 2 Then
               MsgBox ("The BrdRev:" & Rev & ",length must be 2 or 3,please check the BrdRev!")
               Exit Sub
            End If
            arrBrdSeq = Split(BrdSeq, ",")
            For i = 0 To UBound(arrBrdSeq)
                If Trim(arrBrdSeq(i)) > "" Then
                    If Not IsNumeric(arrBrdSeq(i)) Then
                        MsgBox ("The BrdSeq:" & arrBrdSeq(i) & " must be numeric!")
                        Exit Sub
                    End If
                     If Len(Jobpn) > 11 Or Len(Rev) > 10 Or Len(Trim(arrBrdSeq(i))) > 10 Or Len(BrdPN) > 11 Or Len(BrdRev) > 11 Then
                        MsgBox "Excel file format error,Please check Row:" & rCount + 1
                        Exit Sub
                    End If
                    str = "insert PhilipsBrdSeqMapping (JobPN, Rev, BrdSeq, BrdPN, BrdRev) values (" & _
                        sq(Jobpn) & "," & sq(Rev) & "," & sq(Trim(arrBrdSeq(i))) & "," & sq(BrdPN) & "," & sq(BrdRev) & ")"
                    Conn.Execute str
                    Insert_Qty = Insert_Qty + 1
                End If
                DoEvents
            Next i
             
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
      Wend
 End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_FujiBrdSeqMappi','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)

 
 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & SheetName & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Update succeed : " & Update_Qty & vbCrLf
              
End Sub

Private Sub Load_SingleSideBrd(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim MBPN As String, DeletedFlag As String, BuildType As String
  Dim Total_Qty, Deleted_Qty, Insert_Qty As Long
  Dim str As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  rCount = 2
  Total_Qty = 0

  Insert_Qty = 0
  Deleted_Qty = 0
  str = "select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
  
  With xlsBook.Worksheets(Trim(Shift_Item))
       
    While Trim(.Cells(rCount, 1)) <> ""
       
        MBPN = Trim(.Cells(rCount, 1) & vbNullString)
        DeletedFlag = Trim(.Cells(rCount, 2) & vbNullString)
        BuildType = Trim(.Cells(rCount, 3) & vbNullString)
        
        If BuildType = "" Then BuildType = "1"
        '******************************
        '****add by jeanson 2007/09/03
        strErrMessage = ""
        strErrMessage = FunPartNumberCheck(MBPN)
        If strErrMessage <> "PASS" Then
            MsgBox strErrMessage
        Exit Sub
        End If
        '******************************
'        If Len(MBPN) <> 11 Then
'           MsgBox "Excel file format error,please check: ROW:" & rCount + 1
'           Exit Sub
'        End If
        If Trim(BuildType) <> "1" And Trim(BuildType) <> "2" And Trim(BuildType) <> "3" And Trim(BuildType) <> "4" Then
           MsgBox ("BuildType must be 1,2,3 or 4.")
           Exit Sub
        End If
        
        If UCase(DeletedFlag) = "Y" Then
             str = "delete from QSMS_SingleSideBrd where MBPN='" & Trim(MBPN) & "' "
             Conn.Execute str
             Deleted_Qty = Deleted_Qty + 1
        Else
             str = "select * from QSMS_WaveSideBrd where  MBPN='" & Trim(MBPN) & "' " '0046
             Set Rs = Conn.Execute(str)
             If Rs.EOF = True Then
                str = "select * from QSMS_SingleSideBrd where  MBPN='" & Trim(MBPN) & "' "
                Set Rs = Conn.Execute(str)
                If Rs.EOF Then
                    str = "Insert into QSMS_SingleSideBrd(MBPN,UID,TransDateTime,BuildType) " & _
                      " values('" & Trim(MBPN) & "','" & Trim(g_userName) & "','" & transdatetime & "','" & Trim(BuildType) & "')"
                    Conn.Execute str
                   Insert_Qty = Insert_Qty + 1
                End If
            Else
                MsgBox ("this PN in table QSMS_WaveSideBrd,upload fail!")
                Exit Sub
            End If
        End If
                                     
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
             
        rCount = rCount + 1
        Total_Qty = Total_Qty + 1
        Txt_RowCount = Total_Qty
    Wend
End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_SingleSideBrd','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)

 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Deleted succeed : " & Deleted_Qty & vbCrLf
              
End Sub


Private Sub Load_NegativeBrd(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim MBPN As String, DeletedFlag As String
  Dim Total_Qty, Deleted_Qty, Insert_Qty As Long
  Dim str As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  rCount = 2
  Total_Qty = 0

  Insert_Qty = 0
  Deleted_Qty = 0
  str = "select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
  
  With xlsBook.Worksheets(Trim(Shift_Item))
       
       While Trim(.Cells(rCount, 1)) <> ""
           
             MBPN = Trim(.Cells(rCount, 1) & vbNullString)
             DeletedFlag = Trim(.Cells(rCount, 2) & vbNullString)
            '******************************
            '****add by jeanson 2007/09/03
            strErrMessage = ""
            strErrMessage = FunPartNumberCheck(MBPN)
            If strErrMessage <> "PASS" Then
                MsgBox strErrMessage

            Exit Sub
            End If
            '******************************
'             If Len(MBPN) <> 11 Then
'                MsgBox "Excel file format error,please check: ROW:" & rCount + 1
'                Exit Sub
'             End If
             If UCase(DeletedFlag) = "Y" Then
                   str = "delete from QSMS_NegativeBrd where MBPN='" & Trim(MBPN) & "' "
                   Conn.Execute str
                   Deleted_Qty = Deleted_Qty + 1
             Else
                   str = "select * from QSMS_NegativeBrd where  MBPN='" & Trim(MBPN) & "' "
                   Set Rs = Conn.Execute(str)
                   If Rs.EOF Then
                       str = "Insert into QSMS_NegativeBrd(MBPN,UID,TransDateTime) " & _
                         " values('" & Trim(MBPN) & "','" & Trim(g_userName) & "','" & transdatetime & "')"
                       Conn.Execute str
                      Insert_Qty = Insert_Qty + 1
                   End If
             End If
             
               
             
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
      Wend
End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_NegativeBrd','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)

 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Deleted succeed : " & Deleted_Qty & vbCrLf
              
End Sub




Public Function Clear_ReplacePN_List()
    Dim i As Long
    For i = 1 To ReplacePN_MAX_Num
      ReplacePNList(i) = ""
    Next i
End Function

Private Sub Form_Load()
Dim tmpSQL As String
Dim tmpRS As New ADODB.Recordset

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
CboFuncType.AddItem "QSMS_MEBom"

'CboFuncType.AddItem "RACKID"

''''''Added by Jing (0031)
tmpSQL = "Select * from UserRight where UserName='" & Trim(g_userName) & "' and UserRight='UploadReplacePN'"
Set tmpRS = Conn.Execute(tmpSQL)
If tmpRS.EOF = False Then
    CboFuncType.AddItem "REPLACEPN"
End If
CboFuncType.AddItem "QSMS_CheckCompPN"
CboFuncType.AddItem "MachineType"
CboFuncType.AddItem "AVL"
CboFuncType.AddItem "AVL-WIN"
CboFuncType.AddItem "BrdCombineQty"
CboFuncType.AddItem "CastRate"
CboFuncType.AddItem "ComppnInSpectRule"
CboFuncType.AddItem "ControlParts"
CboFuncType.AddItem "CTO_Model"
CboFuncType.AddItem "Component_Data"  '(1024)
CboFuncType.AddItem "DIO"
CboFuncType.AddItem "DocumentComp"
CboFuncType.AddItem "FujiBrdSeqMapping"
CboFuncType.AddItem "JobSide"
CboFuncType.AddItem "LineFUJIServer"
CboFuncType.AddItem "LostReplacePN"
CboFuncType.AddItem "MaterialToWHID" '(0013)
CboFuncType.AddItem "NegativeBrd"
CboFuncType.AddItem "NextDevice"
CboFuncType.AddItem "NonAVL"
CboFuncType.AddItem "NoMachineDropCompPN"
CboFuncType.AddItem "OneByOne"
CboFuncType.AddItem "PhilipsBrdSeqMapping"
CboFuncType.AddItem "PNAlarmQty"
CboFuncType.AddItem "SingleSideBrd"
CboFuncType.AddItem "TraySlot"
'CboFuncType.AddItem "DID"
CboFuncType.AddItem "UNCHKCOMP"
CboFuncType.AddItem "PCB_SingleCompPN"
CboFuncType.AddItem "UpdateJobPN"
CboFuncType.AddItem "Upload_JobGroup"
CboFuncType.AddItem "WOSCHEDULELIST"
'CboFuncType.AddItem "MachineType"
CboFuncType.AddItem "XL_WOPlanSeq"
CboFuncType.AddItem "XL_WOPlanSeqShiftID"
CboFuncType.AddItem "XL_WOPlanLine"
CboFuncType.AddItem "Daily Schedule"
CboFuncType.AddItem "XL_ImplementPN"
CboFuncType.AddItem "XL_WOPN"
CboFuncType.AddItem "XL_PNOneByOne"
CboFuncType.AddItem "XL_PNInterval" '(0019)
CboFuncType.AddItem "XL_EcWOPlan"   '(0022)
CboFuncType.AddItem "XL_DoubleTables"  '(0024)
'CboFuncType.AddItem "WORKHS_EQUIPMENT"  '(0026)
'CboFuncType.AddItem "WORKHS_LINECONFIG"   '(0027)
CboFuncType.AddItem "XL_MaxDIDMaintainQty"
CboFuncType.AddItem "NOCheckReplacePNSplicing"
CboFuncType.AddItem "PNGroup"
'CboFuncType.AddItem "IC_CompPN"
CboFuncType.AddItem "IC_ShearPin"
CboFuncType.AddItem "2ndSource_AssignPN"
CboFuncType.AddItem "upload_traycompPN" '(1001)
CboFuncType.AddItem "Machine_Data"
CboFuncType.AddItem "CompPN_Spacer"   '''(1154)
CboFuncType.AddItem "AVLC"   '''
CboFuncType.AddItem "A8_Manual"   '''
CboFuncType.AddItem "A8_DIDType"   '''
End Sub
Private Sub Load_TraySlot(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim Machine As String, Slot As String, ErrRow As String, Line As String, Side As String
Dim Total_Qty, Insert_Qty As Long, rCount As Long
Dim strSQL As String, transdatetime As String
Dim Rs As New ADODB.Recordset


If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
   Exit Sub
End If
Set xlApp = CreateObject("Excel.Application")
Let xlApp.Visible = False
Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.DisplayAlerts = False

rCount = 2
Total_Qty = 0
Insert_Qty = 0
strSQL = "select getdate()"
Set Rs = Conn.Execute(strSQL)
transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
'delete old data
strSQL = "Truncate table TraySlot"
Conn.Execute (strSQL)
With xlsBook.Worksheets(Trim(Shift_Item))
      While Trim(.Cells(rCount, 1)) <> ""
           Machine = Replace(Trim(.Cells(rCount, 1) & vbNullString), "'", " ")
           Slot = Trim(.Cells(rCount, 2) & vbNullString)
           Line = Trim(.Cells(rCount, 3) & vbNullString)
           Side = Trim(.Cells(rCount, 4) & vbNullString)
           strSQL = "Select Machine from Machine Where Machine='" & Trim(Machine) & "'"
           If Rs.State Then Rs.Close
           Rs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
           If Rs.EOF Then
             MsgBox "Can't find Machine: " & Trim(Machine) & " in MachineType,Please check Machine Name!", vbCritical
             Exit Sub
           End If
           strSQL = "Select * from TraySlot Where Machine='" & Trim(Machine) & "' and Slot='" & Trim(Slot) & "'and line='" & Trim(Line) & "'and Side='" & Trim(Side) & "'"
           If Rs.State Then Rs.Close
           Rs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
           If Rs.EOF = False Then
                MsgBox Trim(Machine) & "," & Trim(Slot) & "," & Trim(Line) & "," & Trim(Side) & " is duplication in Tray,please check!", vbCritical
                ErrRow = ErrRow & "," & rCount
           Else
                strSQL = "Insert TraySlot(Machine,Slot,UID,TransDateTime,line,side) Values('" & Trim(Machine) & "','" & Trim(Slot) & "','" & g_userName & "','" & Trim(transdatetime) & "','" & Trim(Line) & "','" & Trim(Side) & "')"   ''''(1037)
                Conn.Execute strSQL
                Insert_Qty = Insert_Qty + 1
           End If
           Total_Qty = Total_Qty + 1
           rCount = rCount + 1
           Txt_RowCount = Total_Qty
     Wend
End With

If Len(ErrRow) > 0 Then
    ErrRow = Mid(ErrRow, 2)
End If
strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_TraySlot','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSQL)


xlsBook.Close
xlApp.Quit
Set xlApp = Nothing
Set xlsBook = Nothing
MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
             "Total Counter : " & Total_Qty & vbCrLf & _
             "Insert succeed : " & Insert_Qty & vbCrLf & _
             "Err Row:" & ErrRow
End Sub

Public Function InsertAVL(ByVal COMPPN As String, ByVal VendorCode As String, ByVal Customer As String, ByVal Model As String, ByVal Desc1 As String, ByVal rCount As Long, ByVal transdatetime As String, Optional DeletedFlag As String = "")
Dim str As String
Dim Rs As ADODB.Recordset
             If Len(COMPPN) > 11 Or Len(VendorCode) > 3 Or Len(Customer) > 10 Or Len(Model) > 20 Then
                MsgBox "Excel file format error,please check: ROW:" & rCount + 1
                Exit Function
             End If
             If UCase(DeletedFlag) = "Y" Then
                   str = "delete from QSMS_AVL where CompPN='" & Trim(COMPPN) & "' and VendorCode='" & VendorCode & "' and Customer='" & Customer & "' and Model='" & Model & "' "
                   Conn.Execute str
                   Deleted_Qty = Deleted_Qty + 1
             Else
                   str = "select * from QSMS_AVL where CompPN='" & Trim(COMPPN) & "' and VendorCode='" & VendorCode & "' and Customer='" & Customer & "' and Model='" & Model & "' "
                   Set Rs = Conn.Execute(str)
                   If Rs.EOF Then
                       str = "Insert into QSMS_AVL(CompPN,VendorCode,Customer,Model,Desc1,TransDateTime) " & _
                         " values('" & Trim(COMPPN) & "','" & Trim(VendorCode) & "','" & Trim(Customer) & "','" & Model & "','" & Desc1 & "','" & transdatetime & "')"
                       Conn.Execute str
                      Insert_Qty = Insert_Qty + 1
                   End If
             End If
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
End Function

Private Sub Load_AVL_WIN(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount As Long
  Dim aryModel, aryVcode As Variant
  Dim i, j As Integer
  Dim COMPPN, Desc1, VendorCode, VendorCode1, VendorCode2, DateCode, LotCode, Customer, Model, Model1, Model2 As String, DeletedFlag As String
  Dim str, TempStr As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String, delflag As String
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  Total_Qty = 0
  Insert_Qty = 0
  Deleted_Qty = 0
  str = "select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")

'///////////////////////////////////////for Dell(BOMList)////////////////////////////////////////////////////

rCount = 2
With xlsBook.Worksheets(Trim(Shift_Item))
    While Trim(.Cells(rCount, 1)) <> ""
        'Customer = Trim(.Cells(rCount, 1) & vbNullString)
        'Only for WIN
        Customer = "WIN"
        Model = Trim(Replace(Replace(.Cells(rCount, 1) & vbNullString, vbCr, ""), vbLf, ""))
        COMPPN = Trim(Replace(Replace(.Cells(rCount, 2) & vbNullString, vbCr, ""), vbLf, ""))
        VendorCode1 = Trim(Replace(Replace(.Cells(rCount, 3) & vbNullString, vbCr, ""), vbLf, ""))
        'Desc1 = Replace(Trim(.Cells(rCount, 4) & vbNullString), "'", " ")
        delflag = Trim(Replace(Replace(.Cells(rCount, 4) & vbNullString, vbCr, ""), vbLf, ""))
        
        '*********************check model if defined*********************************
        str = "select * from modelname where modelname=" & sq(Model)
        Set Rs = Conn.Execute(str)
        If Rs.EOF Then
           MsgBox "This ModelName not defined, Model=" & Model & ", PN=" & COMPPN & ", ROW:" & rCount
           Exit Sub
        End If
        
        'multiple vendor
        If InStr(1, VendorCode1, ";") > 0 Then
            aryVcode = Split(VendorCode1, ";")
            For i = 0 To UBound(aryVcode)
                TempStr = Trim(aryVcode(i))
                If InStr(1, VendorCode1, "1.") > 0 Then
                    VendorCode2 = StrBetween(TempStr, ".", "(", 1)
                    If Len(Trim(VendorCode2)) = 3 Then
                        VendorCode = VendorCode2
                    End If
                Else
                    VendorCode = Mid(TempStr, 1, 3)
                End If
                Call InsertAVL(Trim(COMPPN), Trim(VendorCode), Trim(Customer), Trim(Model), Trim(Desc1), Trim(rCount), Trim(transdatetime), delflag)
            Next i
        Else
            VendorCode = Mid(VendorCode1, 1, 3)
            Call InsertAVL(Trim(COMPPN), Trim(VendorCode), Trim(Customer), Trim(Model), Trim(Desc1), Trim(rCount), Trim(transdatetime), delflag)
        End If
        rCount = rCount + 1
        DoEvents
    Wend
End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_AVL_WIN','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)
xlsBook.Close
xlApp.Quit
Set xlApp = Nothing
Set xlsBook = Nothing

 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Deleted succeed : " & Deleted_Qty & vbCrLf
End Sub
Private Sub Load_AVL_ControlParts(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount As Long
  Dim aryModel, aryVcode As Variant
  Dim i, j As Integer
  Dim PNCheck As String
  Dim COMPPN, Desc1, VendorCode, VendorCode1, VendorCode2, DateCode, LotCode, Customer, Model, Model1, Model2 As String, DeletedFlag As String
  Dim str, TempStr As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  Total_Qty = 0
  Insert_Qty = 0
  Deleted_Qty = 0
  str = "select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")

'///////////////////////////////////////for ControlParts////////////////////////////////////////////////////
       rCount = 2
 With xlsBook.Worksheets(Trim(Shift_Item))
       While Trim(.Cells(rCount, 1)) <> ""
             'Customer = Trim(.Cells(rCount, 1) & vbNullString)
             Model = Trim(.Cells(rCount, 1) & vbNullString)
             PNCheck = Trim(.Cells(rCount, 2) & vbNullString)
             If Len(PNCheck) > 11 Then
                MsgBox "This CompPN format is error !ROW:" & rCount
                Exit Sub
             End If
             COMPPN = Mid(Trim(.Cells(rCount, 2) & vbNullString), 1, 11)
             VendorCode = Mid(Trim(Replace(Trim(.Cells(rCount, 3) & vbNullString), " ", "")), 1, 3)
'*********************check modelname if defined*********************************
             str = "select * from ModelName where modelname=" & sq(Model)
             Set Rs = Conn.Execute(str)
             If Rs.EOF Then
                MsgBox "This ModelName not defined !ROW:" & rCount
                Exit Sub
             End If
'*********************check modelname if defined*********************************
'             If Len(VendorCode) > 3 Then
'                MsgBox "The format of VendorCode is error!please check: ROW:" & rCount
'                Exit Sub
'             End If
             If Len(Customer) > 11 Or Len(Model) > 20 Then
                MsgBox "Excel file format error,please check: ROW:" & rCount
                Exit Sub
             End If
             If UCase(DeletedFlag) = "Y" Then
                   str = "delete from QSMS_ControlPart where CompPN='" & Trim(COMPPN) & "' and VendorCode='" & VendorCode & "' and Model='" & Model & "' "
                   Conn.Execute str
                   Deleted_Qty = Deleted_Qty + 1
             Else
                   str = "select * from QSMS_ControlPart where CompPN='" & Trim(COMPPN) & "' and VendorCode='" & VendorCode & "' and Model='" & Model & "' "
                   Set Rs = Conn.Execute(str)
                   If Rs.EOF Then
                       str = "Insert into QSMS_ControlPart(CompPN,VendorCode,Model,TransDateTime) " & _
                         " values('" & Trim(COMPPN) & "','" & Trim(VendorCode) & "','" & Model & "','" & transdatetime & "')"
                       Conn.Execute str
                      Insert_Qty = Insert_Qty + 1
                   End If
             End If
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
      Wend
End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_AVL_ControlPart','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)

  xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing

 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Deleted succeed : " & Deleted_Qty & vbCrLf
End Sub

Private Sub Load_CastRate(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim CompHead, CompHead1, Desc1, DeletedFlag, Rate1, UpLimit1 As String
  Dim aryVcode As Variant
  Dim Total_Qty, Deleted_Qty, Insert_Qty, Update_Qty As Long
  Dim UpLimit
  Dim Rate
  Dim i As Integer
  Dim str As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  rCount = 2
  Total_Qty = 0
  Update_Qty = 0
  Insert_Qty = 0
  Deleted_Qty = 0
  str = "select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
  
  With xlsBook.Worksheets(Trim(Shift_Item))
    While Trim(.Cells(rCount, 1)) <> ""
        CompHead1 = Trim(.Cells(rCount, 1) & vbNullString)
        Desc1 = Trim(.Cells(rCount, 2) & vbNullString)
        Rate1 = Trim(.Cells(rCount, 3) & vbNullString)
        UpLimit1 = Trim(.Cells(rCount, 4) & vbNullString)
        DeletedFlag = Trim(.Cells(rCount, 5) & vbNullString)
        Rate = CDec(Rate1)
           If UpLimit1 = "" Then
               UpLimit = CStr(UpLimit1)
           Else
               UpLimit = CInt(UpLimit1)
           End If
            aryVcode = Split(CompHead1, ",")
                For i = 0 To UBound(aryVcode)
                    CompHead = aryVcode(i)
'                    str = "select * from QSMS_CastRate where CompHead='" & Trim(CompHead) & "'"
'                    Set Rs = Conn.Execute(str)
'                    If Not Rs.EOF Then
'                        MsgBox "This PN already exist !PN:" & CompHead
'                    End If
                    If UCase(DeletedFlag) = "Y" Then
                        str = "delete from QSMS_CastRate where CompHead='" & Trim(CompHead) & "'"
                        Conn.Execute str
                        Deleted_Qty = Deleted_Qty + 1
                    Else
                        str = "select * from QSMS_CastRate where CompHead='" & Trim(CompHead) & "'"
                        Set Rs = Conn.Execute(str)
                        If Rs.EOF Then
                            str = "Insert into QSMS_CastRate(CompHead,Rate,UpLimit,Desc1,UID,TransDateTime) " & _
                                " values('" & Trim(CompHead) & "','" & Rate & "','" & UpLimit & "' ,'" & Desc1 & "','" & g_userName & "','" & transdatetime & "')"
                            Conn.Execute str
                            Insert_Qty = Insert_Qty + 1
                        Else
                            str = "Update QSMS_CastRate set Desc1='" & Desc1 & "', Rate='" & Rate & "', UpLimit='" & UpLimit & "', UID='" & g_userName & "', TransDateTime='" & transdatetime & "' where CompHead='" & Trim(CompHead) & "'"
                            Conn.Execute str
                            Update_Qty = Update_Qty + 1
                        End If
                    End If
                Next i
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
    Wend
         DoEvents
         DoEvents
         DoEvents
         DoEvents
         DoEvents
             
End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_CastRate','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)


 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Update succeed : " & Update_Qty & vbCrLf & _
               "Deleted succeed : " & Deleted_Qty & vbCrLf
              
End Sub



Private Sub Load_OneByOne(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim Item, COMPPN, OneByOne As String
  Dim Desc1 As String
  Dim position As Integer
  Dim aryVcode As Variant
  Dim Total_Qty, Deleted_Qty, Insert_Qty, Update_Qty As Long
  Dim TempDesc1 As String
  Dim UpLimit
  Dim Rate
  Dim i As Integer
  Dim PINQty As Integer
  Dim str As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String
  
'料号Header  规格             one by one 定义
'---------- -------------    --------------------------------------------------------------
'A          IC               Description 里()中,若是<=8p,就定义为N,若>8P,则为 Y
'L          IC               Description 里()中,若是<=8p,就定义为N,若>8P,则为 Y
'DF         connecter        compn 里,第5,6码若<=08 ,就定义为N,若>08 或为英文字母,则为 Y
'SF         connecter        compn 里,第5,6码若<=08 ,就定义为N,若>08 或为英文字母,则为 Y

  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  rCount = 2
  Total_Qty = 0
  Update_Qty = 0
  Insert_Qty = 0
  Deleted_Qty = 0
  str = "select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
  
  With xlsBook.Worksheets(Trim(Shift_Item))
    While Trim(.Cells(rCount, 1)) <> ""
        PINQty = 0
        Item = Trim(.Cells(rCount, 1) & vbNullString)
        COMPPN = Trim(.Cells(rCount, 2) & vbNullString)
        Desc1 = Trim(.Cells(rCount, 3) & vbNullString)
        
        ''for A or L head compPN  --add by Giant 20071006
        If UCase(Mid(COMPPN, 1, 1)) = "A" Or UCase(Mid(COMPPN, 1, 1)) = "L" Then
             If StrBetween(Desc1, "(", ")") = "" Or Val(StrBetween(Desc1, "(", ")")) = 0 Then
                 MsgBox "Desc format error,Please check,Line : " & rCount
                 Set xlApp = Nothing
                 Set xlsBook = Nothing
                 Exit Sub
             End If
             PINQty = Val(StrBetween(Desc1, "(", ")"))
        End If

        ''for DF or SF head compPN  --add by Giant 20071006
        If UCase(Mid(COMPPN, 1, 2)) = "DF" Or UCase(Mid(COMPPN, 1, 2)) = "SF" Then
           If IsNumeric(Mid(COMPPN, 5, 2)) = False Then
               If IsNumeric(Mid(COMPPN, 6, 1)) = False Then
                    MsgBox "CompPn format error,Please check,Line : " & rCount
                    Set xlApp = Nothing
                    Set xlsBook = Nothing
                    Exit Sub
               Else
                    PINQty = 100
               End If
           Else
               PINQty = CLng(Mid(COMPPN, 5, 2))
           End If
        End If
        
        If PINQty > 8 Then
           OneByOne = "Y"
        Else
           OneByOne = "N"
        End If
        
        str = "select * from QSMS_OneByOne where CompPN='" & Trim(COMPPN) & "'"
        Set Rs = Conn.Execute(str)
        If Not Rs.EOF Then
            str = "Update QSMS_OneByOne set Desc1='" & Desc1 & "', OneByOne='" & OneByOne & "', UID='" & g_userName & "', TransDateTime='" & transdatetime & "' where CompPN='" & Trim(COMPPN) & "'"
            Conn.Execute str
            Update_Qty = Update_Qty + 1
        Else
            str = "Insert into QSMS_OneByOne(CompPN,OneByOne,Desc1,UID,TransDateTime) " & _
                " values('" & Trim(COMPPN) & "','" & OneByOne & "','" & Desc1 & "' ,'" & g_userName & "','" & transdatetime & "')"
            Conn.Execute str
            Insert_Qty = Insert_Qty + 1
        End If
        
        rCount = rCount + 1
        Total_Qty = Total_Qty + 1
        Txt_RowCount = Total_Qty
    Wend
         DoEvents
         DoEvents
         DoEvents
         DoEvents
         DoEvents
             
End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_OneByOne','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)


 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Update succeed : " & Update_Qty & vbCrLf & _
               "Deleted succeed : " & Deleted_Qty & vbCrLf
              
End Sub

Private Sub Load_NextDevice(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim NextDeviceID, Machine, Jobpn, jobgroup, Version, COMPPN, LR, Slot, flag, DeletedFlag, ChkFlag As String
  Dim tempNextDeviceID, tempmachine, TempJobPn, TempJObGroup, tempVersion, tempCompPN, tempLR, tempSlot, tempFlag As String
  Dim Total_Qty, Deleted_Qty, Insert_Qty, Update_Qty As Long
  Dim strSQL As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  rCount = 2
  Total_Qty = 0
  Update_Qty = 0
  Insert_Qty = 0
  Deleted_Qty = 0
  strSQL = "select getdate()"
  Set Rs = Conn.Execute(strSQL)
  transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")

  With xlsBook.Worksheets(Trim(Shift_Item))

       While Trim(.Cells(rCount, 1)) <> ""

            NextDeviceID = Trim(.Cells(rCount, 1) & vbNullString)
            Machine = Trim(.Cells(rCount, 2) & vbNullString)
            Jobpn = Trim(.Cells(rCount, 3) & vbNullString)
            jobgroup = Trim(.Cells(rCount, 4) & vbNullString)
            Version = Trim(.Cells(rCount, 5) & vbNullString)
            COMPPN = Trim(.Cells(rCount, 6) & vbNullString)
            LR = Trim(.Cells(rCount, 7) & vbNullString)
            Slot = Trim(.Cells(rCount, 8) & vbNullString)
            flag = Trim(.Cells(rCount, 9) & vbNullString)
            DeletedFlag = Trim(.Cells(rCount, 10) & vbNullString)
            '******************************
            '****add by jeanson 2007/09/03
            strErrMessage = ""
            strErrMessage = FunPartNumberCheck(COMPPN)
            If strErrMessage <> "PASS" Then
                MsgBox strErrMessage

            Exit Sub
            End If
            '******************************
'            If Len(Version) <> 3 Or Len(Jobpn) <> 11 Or Len(CompPN) <> 11 Then
            If Len(Version) <> 3 Then
               MsgBox "Excel file format error,Please check: row :" & rCount + 1
               Exit Sub
            End If
            
            If UCase(DeletedFlag) = "Y" Then
                strSQL = "delete from QSMS_MEBom_NextDevice where Machine='" & Trim(Machine) & "' and Jobgroup='" & jobgroup & "'"
                Conn.Execute strSQL
                Deleted_Qty = Deleted_Qty + 1
            Else
                If (tempmachine = "" Or tempmachine <> Machine) Or (TempJObGroup = "" Or TempJObGroup <> jobgroup) Then
                   strSQL = "delete from QSMS_MEBom_NextDevice where  Jobgroup='" & jobgroup & "' and Machine='" & Machine & "' "
                   Conn.Execute strSQL
                   Deleted_Qty = Deleted_Qty + 1
                End If
               strSQL = "Insert into QSMS_MEBom_NextDevice(NextDeviceID,Machine,Jobpn,JobGroup,Version,CompPN,LR,Slot,Flag,UID,TransDateTime) " & _
               " values('" & Trim(NextDeviceID) & "','" & Machine & "','" & Jobpn & "' ,'" & jobgroup & "','" & Version & "','" & COMPPN & "','" & LR & "','" & Slot & "','" & flag & "','" & g_userName & "','" & transdatetime & "')"
                  Conn.Execute strSQL
                  Insert_Qty = Insert_Qty + 1
            End If

             tempmachine = Machine
             TempJobPn = Jobpn
             tempVersion = Version
             TempJObGroup = jobgroup
             tempCompPN = COMPPN
             
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
      Wend
End With
strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_NextDevice','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSQL)

 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Delete succeed : " & Deleted_Qty & vbCrLf
              
End Sub


Private Sub Load_JobSide(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim Jobpn As String, Side As String, DeletedFlag As String
  Dim Total_Qty, Deleted_Qty, Insert_Qty As Long
  Dim str As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  rCount = 2
  Total_Qty = 0

  Insert_Qty = 0
  Deleted_Qty = 0
  str = "select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
  
  With xlsBook.Worksheets(Trim(Shift_Item))
       
       While Trim(.Cells(rCount, 1)) <> ""
           
             Jobpn = Trim(.Cells(rCount, 1) & vbNullString)
             Side = Trim(.Cells(rCount, 2) & vbNullString)
             DeletedFlag = Trim(.Cells(rCount, 3) & vbNullString)
            '******************************
            '****add by jeanson 2007/09/03
            strErrMessage = ""
            strErrMessage = FunPartNumberCheck(Jobpn)
            If strErrMessage <> "PASS" Then
                MsgBox strErrMessage

            Exit Sub
            End If
            '******************************
'            If Len(Jobpn) <> 11 Or Len(Side) <> 1 Or UCase((Trim(Side)) <> "C" And UCase(Trim(Side)) <> "S" And UCase(Trim(Side)) <> "W") Then
             If Len(Side) <> 1 Or UCase((Trim(Side)) <> "C" And UCase(Trim(Side)) <> "S" And UCase(Trim(Side)) <> "W") Then
                MsgBox "Excel file format error,please check: ROW:" & rCount + 1
                Exit Sub
             End If
             If UCase(DeletedFlag) = "Y" Then
                   str = "delete from QSMS_JobSide where JobPN='" & Trim(Jobpn) & "' "
                   Conn.Execute str
                   Deleted_Qty = Deleted_Qty + 1
             Else
                   str = "select * from QSMS_JobSide where  JobPN='" & Trim(Jobpn) & "' "
                   Set Rs = Conn.Execute(str)
                   If Rs.EOF Then
                       str = "Insert into QSMS_JobSide(JobPN,Side,UID,TransDateTime) " & _
                         " values('" & Trim(Jobpn) & "','" & UCase(Trim(Side)) & "','" & Trim(g_userName) & "','" & transdatetime & "')"
                       Conn.Execute str
                      Insert_Qty = Insert_Qty + 1
                   End If
             End If
             
               
             
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
      Wend
End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_JobSide','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)

 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Deleted succeed : " & Deleted_Qty & vbCrLf
              
End Sub

Private Sub Load_BrdCombineQty(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim MBPN, DeletedFlag As String
  Dim CombineQty As Long
  Dim Total_Qty, Deleted_Qty, Insert_Qty, Update_Qty As Long
  Dim str As String
  Dim Rs As ADODB.Recordset
  Dim transdatetime As String
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.DisplayAlerts = False
  rCount = 2
  Total_Qty = 0

  Insert_Qty = 0
  Deleted_Qty = 0
  Update_Qty = 0
  str = "select getdate()"
  Set Rs = Conn.Execute(str)
  transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")
  
  With xlsBook.Worksheets(Trim(Shift_Item))
       
       While Trim(.Cells(rCount, 1)) <> ""
           
             MBPN = Trim(.Cells(rCount, 1) & vbNullString)
             CombineQty = Trim(.Cells(rCount, 2) & vbNullString)
             DeletedFlag = Trim(.Cells(rCount, 3) & vbNullString)
            '******************************
            '****add by jeanson 2007/09/03
            strErrMessage = ""
            strErrMessage = FunPartNumberCheck(MBPN)
            If strErrMessage <> "PASS" Then
                MsgBox strErrMessage

            Exit Sub
            End If
            '******************************
'             If Len(MBPN) <> 11 Or IsNumeric(CombineQty) = False Then
             If IsNumeric(CombineQty) = False Then
                MsgBox "Excel file format error,please check: ROW:" & rCount + 1
                Exit Sub
             End If
             If UCase(DeletedFlag) = "Y" Then
                   str = "delete from QSMS_BrdCombineQty where MBPN='" & Trim(MBPN) & "' "
                   Conn.Execute str
                   Deleted_Qty = Deleted_Qty + 1
             Else
                   str = "select * from QSMS_BrdCombineQty where MBPN='" & Trim(MBPN) & "' "
                   Set Rs = Conn.Execute(str)
                   If Rs.EOF Then
                       str = "Insert into QSMS_BrdCombineQty(MBPN,CombineQty,UID,TransDateTime) " & _
                         " values('" & Trim(MBPN) & "'," & CombineQty & ",'" & g_userName & "','" & transdatetime & "')"
                       Conn.Execute str
                      Insert_Qty = Insert_Qty + 1
                    Else
                       str = "Update QSMS_BrdCombineQty set CombineQty=" & CombineQty & ",UID='" & g_userName & "',TransDateTime='" & transdatetime & "' where MBPN='" & Trim(MBPN) & "' "
                       Conn.Execute str
                       Update_Qty = Update_Qty + 1
                   End If
             End If
             
               
             
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             DoEvents
             
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
      Wend
End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','BrdCombineQty','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)

 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Update succeed : " & Update_Qty & vbCrLf & _
               "Deleted succeed : " & Deleted_Qty & vbCrLf
              
End Sub

Private Sub Load_coputername(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim rCount, Row_Count As Long
  Dim str As String
  Dim Rs As New ADODB.Recordset
  Dim transdatetime As String
  Dim i As Integer
  Dim strSQL As String
  Const intField_1 As Integer = 1
  Const intField_2 As Integer = 2
  
  
  
  If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
     Exit Sub
  End If
  Set xlApp = CreateObject("Excel.Application")
  Let xlApp.Visible = False
  Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
  xlApp.UserControl = True
  xlApp.DisplayAlerts = False
  
  
  '''''''''''''————————初始化数据表————————''''''''
If Trim(xlApp.Worksheets(1).Cells(1, intField_1)) <> "Line" And Trim(xlApp.Worksheets(1).Cells(1, intField_2)) <> "PCName" Then
MsgBox "Excel Field Format Error, please tel QMS for help!", vbOKOnly + vbInformation, "Excel Field Format Error"

xlApp.Quit
Set xlApp = Nothing
Exit Sub
End If

i = 2
Do While Not (xlApp.Worksheets(1).Cells(i, 1) = "")
  
'''''''''''''————————导入数据，用插入方法导入数据————————''''''''
strSQL = "select * from QSMS_ProConfig where line='" & Trim(xlApp.Worksheets(1).Cells(i, intField_1)) & "' and station='DIO' AND session='BASE' AND [KEY]='HostName'"
If Rs.State = 1 Then Rs.Close
Set Rs = Conn.Execute(strSQL)
    If Not Rs.EOF Then
        strSQL = "update QSMS_ProConfig set value='" & Trim(xlApp.Worksheets(1).Cells(i, intField_2)) & "' where line='" & Trim(xlApp.Worksheets(1).Cells(i, intField_1)) & "' and station='DIO' AND session='BASE' AND [KEY]='HostName'"
        If Rs.State = 1 Then Rs.Close
        Set Rs = Conn.Execute(strSQL)
    Else
        strSQL = "Insert Into QSMS_ProConfig Values ('" & Trim(xlApp.Worksheets(1).Cells(i, intField_1)) & "','DIO','BASE','HostName','" & Trim(xlApp.Worksheets(1).Cells(i, intField_2)) & "')"
        If Rs.State = 1 Then Rs.Close
        Set Rs = Conn.Execute(strSQL)
    End If
   i = i + 1
Loop
Set Rs = Nothing

xlApp.Quit
Set xlApp = Nothing

MsgBox "Excel upload to SQL ok"
  
  
End Sub

Private Sub Load_InSpectRule(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim COMPPN As String, Unit As String, ErrRow As String
Dim Total_Qty, Insert_Qty As Long, rCount As Long, Upper As Double, Lower As Double, Hz As Single, Volt As Single, Ampere As Single, ChkNum As Integer, BaseQty As Long, delflag As String
Dim strSQL As String, transdatetime As String
Dim resistPn As String
Dim digit3 As String, digit4 As String, digit5 As String, digit6 As String, digit7 As String, digit8 As String
Dim resistanceValue As Double, CapacitanceValue As Double
Dim Rs As New ADODB.Recordset

If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
   Exit Sub
End If
Set xlApp = CreateObject("Excel.Application")
Let xlApp.Visible = False
Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.DisplayAlerts = False

rCount = 2
Total_Qty = 0
Insert_Qty = 0

strSQL = "select getdate()"
Set Rs = Conn.Execute(strSQL)
transdatetime = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")

'delete old data
'''strsql = "Truncate table QSMS_InSpect_Rule"
'Conn.Execute (strSql)
With xlsBook.Worksheets(Trim(Shift_Item))
    While Trim(.Cells(rCount, 1)) <> ""
        COMPPN = Replace(Trim(.Cells(rCount, 1) & vbNullString), "'", " ")
        Upper = Val(Trim(.Cells(rCount, 2) & vbNullString))
        Lower = Val(Trim(.Cells(rCount, 3) & vbNullString))
        Hz = Val(Trim(.Cells(rCount, 4) & vbNullString))     '''----- (0018)
        Volt = Val(Trim(.Cells(rCount, 5) & vbNullString))
        Ampere = Val(Trim(.Cells(rCount, 6) & vbNullString))
        Unit = Trim(.Cells(rCount, 7) & vbNullString)
        ChkNum = Val(Trim(.Cells(rCount, 8) & vbNullString))
        BaseQty = Val(Trim(.Cells(rCount, 9) & vbNullString))
        delflag = UCase(Trim(.Cells(rCount, 10) & vbNullString))
        
        If Unit = "" Then
            MsgBox "error:Uint is  Null !", vbCritical
            ErrRow = ErrRow & "," & rCount
            Exit Sub
        End If
        
        If Upper < Lower Then
            MsgBox Trim(COMPPN) & " is Upper smaller than Lower !", vbCritical
            ErrRow = ErrRow & "," & rCount
            Exit Sub
        End If
           
        If ChkPNCQ = "Y" Then   ''1211
            strSQL = "Select * from QSMS_InSpect_Rule Where Comppn='" & Trim(COMPPN) & "'"
            If Rs.State Then Rs.Close
            Rs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
            If Rs.EOF = False Then
                 If delflag = "Y" Then
                     strSQL = "Delete from QSMS_InSpect_Rule where Comppn='" & Trim(COMPPN) & "'  "
                     Conn.Execute strSQL
                 Else
                     MsgBox Trim(COMPPN) & " already exists, if you want to update, Step1: please delete the P/N (Set delflag=Y) ,Step2: Upload the P/N(Set delflag=N)", vbCritical
                     ErrRow = ErrRow & "," & rCount
                     xlsBook.Close
                     xlApp.Quit
                     Set xlApp = Nothing
                     Set xlsBook = Nothing
                     Exit Sub
                 End If
            Else
                 If delflag = "N" Then
                     strSQL = "Insert QSMS_InSpect_Rule(Comppn,Upper,Lower,Hz,Volt,Ampere,Unit,ChkNum,BaseQTy,UID,TransDateTime) Values('" & Trim(COMPPN) & "','" & Trim(Upper) & "','" & Trim(Lower) & "','" & Trim(Hz) & "','" & Trim(Volt) & "','" & Trim(Ampere) & "',N'" & Trim(Unit) & "', '" & Trim(ChkNum) & "','" & Trim(BaseQty) & "','" & g_userName & "','" & transdatetime & "'  )"
                     Conn.Execute strSQL
                 Else
                     MsgBox Trim(COMPPN) & " not  exists, if you want to upload, please Set delflag=N", vbCritical
                     ErrRow = ErrRow & "," & rCount
                     xlsBook.Close
                     xlApp.Quit
                     Set xlApp = Nothing
                     Set xlsBook = Nothing
                     Exit Sub
                 End If
            End If
        Else
             ''''''''''''''''''''''''''''''(1235) begin''''''''''''''''''''''''''''''''''''''
             strSQL = "Select * from QSMS_InSpect_Rule with(nolock) Where CompPN='" & Trim(COMPPN) & "'"
             If Rs.State Then Rs.Close
             Rs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
             If Rs.EOF = False Then
                If delflag = "Y" Then
                    strSQL = "Delete from QSMS_InSpect_Rule where Comppn='" & Trim(COMPPN) & "'  "
                    Conn.Execute strSQL
                Else
                   MsgBox Trim(COMPPN) & " already exists, if you want to update, Step1: please delete the P/N (Set delflag=Y) ,Step2: Upload the P/N(Set delflag=N)", vbCritical
                   ErrRow = ErrRow & "," & rCount
                   xlsBook.Close
                   xlApp.Quit
                   Set xlApp = Nothing
                   Set xlsBook = Nothing
                End If
                Exit Sub
             Else
                strSQL = "Select * from QSMS_UploadInSpectRule_Temp with(nolock) Where CompPN='" & Trim(COMPPN) & "'"
                If Rs.State Then Rs.Close
                Rs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
                If Rs.EOF = False Then
                    strSQL = "Delete from QSMS_UploadInSpectRule_Temp where CompPN='" & Trim(COMPPN) & "'  "
                    Conn.Execute strSQL
                End If
                strSQL = "Insert QSMS_UploadInSpectRule_Temp(CompPN,Upper,Lower,Hz,Volt,Ampere,Unit,ChkNum,BaseQTy,UID,TransDateTime,DeleteFlag) Values('" & Trim(COMPPN) & "','" & Trim(Upper) & "','" & Trim(Lower) & "','" & Trim(Hz) & "','" & Trim(Volt) & "','" & Trim(Ampere) & "',N'" & Trim(Unit) & "', '" & Trim(ChkNum) & "','" & Trim(BaseQty) & "','" & g_userName & "','" & transdatetime & "' ,'" & delflag & "')"
                Conn.Execute strSQL
             End If
            ''''''''''''''''''''''''''''''(1235) end'''''''''''''''''''''''''''''''''''''''
             
'            If Left(Trim(.Cells(rCount, 1)), 2) <> "CS" And Left(Trim(.Cells(rCount, 1)), 2) <> "CH" Then
'                strSql = "Select * from QSMS_InSpect_Rule Where Comppn='" & Trim(compPN) & "'"
'                If rs.State Then rs.Close
'                rs.Open strSql, Conn, adOpenForwardOnly, adLockReadOnly
'                If rs.EOF = False Then
'                     If delflag = "Y" Then
'                         strSql = "Delete from QSMS_InSpect_Rule where Comppn='" & Trim(compPN) & "'  "
'                         Conn.Execute strSql
'                     Else
'                         MsgBox Trim(compPN) & " already exists, if you want to update, Step1: please delete the P/N (Set delflag=Y) ,Step2: Upload the P/N(Set delflag=N)", vbCritical
'                         ErrRow = ErrRow & "," & rCount
'                         xlsBook.Close
'                         xlApp.Quit
'                         Set xlApp = Nothing
'                         Set xlsBook = Nothing
'                         Exit Sub
'                     End If
'                Else
'                     If delflag = "N" Then
'                         strSql = "Insert QSMS_InSpect_Rule(Comppn,Upper,Lower,Hz,Volt,Ampere,Unit,ChkNum,BaseQTy,UID,TransDateTime) Values('" & Trim(compPN) & "','" & Trim(Upper) & "','" & Trim(Lower) & "','" & Trim(Hz) & "','" & Trim(Volt) & "','" & Trim(Ampere) & "',N'" & Trim(Unit) & "', '" & Trim(ChkNum) & "','" & Trim(BaseQty) & "','" & g_userName & "','" & transdatetime & "'  )"
'                         Conn.Execute strSql
'                     Else
'                         MsgBox Trim(compPN) & " not  exists, if you want to upload, please Set delflag=N", vbCritical
'                         ErrRow = ErrRow & "," & rCount
'                         xlsBook.Close
'                         xlApp.Quit
'                         Set xlApp = Nothing
'                         Set xlsBook = Nothing
'                         Exit Sub
'                     End If
'                End If
'            ElseIf Left(Trim(.Cells(rCount, 1)), 2) = "CS" Then ''电阻的值
'                strSql = "Select * from QSMS_InSpect_Rule Where Comppn='" & Trim(compPN) & "'"
'                If rs.State Then rs.Close
'                    rs.Open strSql, Conn, adOpenForwardOnly, adLockReadOnly
'                    If rs.EOF = False Then
'                        If delflag = "Y" Then
'                            strSql = "Delete from QSMS_InSpect_Rule where Comppn='" & Trim(compPN) & "'  "
'                            Conn.Execute strSql
'                        Else
'                            MsgBox Trim(compPN) & " already exists, if you want to update, Step1: please delete the P/N (Set delflag=Y) ,Step2: Upload the P/N(Set delflag=N)", vbCritical
'                            ErrRow = ErrRow & "," & rCount
'                            xlsBook.Close
'                            xlApp.Quit
'                            Set xlApp = Nothing
'                            Set xlsBook = Nothing
'                            Exit Sub
'                        End If
'                    ElseIf delflag = "N" Then
'                        digit3 = Mid(Trim(compPN), 3, 1)
'                        digit4 = Mid(Trim(compPN), 4, 1)
'                        digit5 = Mid(Trim(compPN), 5, 1)
'                        digit6 = Mid(Trim(compPN), 6, 1)
'                        digit8 = Mid(Trim(compPN), 8, 1)
'
'                        Select Case digit3
'                        Case "-"
'                        resistanceValue = (Val(digit4) * 10 + Val(digit5) + Val(digit6) * 0.1) * 0.1
'                        Case "+"
'                        resistanceValue = (Val(digit4) * 10 + Val(digit5) + Val(digit6) * 0.1) * 0.01
'                        Case Else
'                        resistanceValue = (Val(digit4) * 10 + Val(digit5) + Val(digit6) * 0.1) * (10 ^ Val(digit3))
'                        End Select
'
'                        Select Case digit8
'                        Case "B"
'                            Upper = resistanceValue * (1.001)
'                            Lower = resistanceValue * (0.999)
'                        Case "C"
'                            Upper = resistanceValue * (1.0025)
'                            Lower = resistanceValue * (0.9975)
'                        Case "D"
'                            Upper = resistanceValue * (1.005)
'                            Lower = resistanceValue * (0.995)
'                        Case "F"
'                            Upper = resistanceValue * (1.01)
'                            Lower = resistanceValue * (0.99)
'                        Case "G"
'                            Upper = resistanceValue * (1.02)
'                            Lower = resistanceValue * (0.98)
'                        Case "J"
'                            Upper = resistanceValue * (1.05)
'                            Lower = resistanceValue * (0.95)
'                        Case "K"
'                            Upper = resistanceValue * (1.1)
'                            Lower = resistanceValue * (0.9)
'                        Case "M"
'                            Upper = resistanceValue * (1.2)
'                            Lower = resistanceValue * (0.8)
'                        Case "Z"
'                            Upper = resistanceValue * (1.8)
'                            Lower = resistanceValue * (0.8)
'                        Case Else
'                            MsgBox ("The tolerance of this comppn does not exist")
'                        End Select
'                     ''strSql = "Insert QSMS_InSpect_Rule(Comppn,Upper,Lower,Hz,Volt,Ampere,Unit,ChkNum,BaseQTy,UID,TransDateTime) Values('" & Trim(compPN) & "','" & Trim(Upper) & "','" & Trim(Lower) & "','" & 0 & "','" & 0 & "','" & 0 & "',N'" & "R" & "', '" & 1 & "','" & 50 & "','" & g_userName & "','" & transdatetime & "'  )"
'                        strSql = "Insert QSMS_InSpect_Rule(Comppn,Upper,Lower,Hz,Volt,Ampere,Unit,ChkNum,BaseQTy,UID,TransDateTime) Values('" & Trim(compPN) & "','" & Trim(Upper) & "','" & Trim(Lower) & "','" & Trim(Hz) & "','" & Trim(Volt) & "','" & Trim(Ampere) & "',N'" & Trim(Unit) & "', '" & Trim(ChkNum) & "','" & Trim(BaseQty) & "','" & g_userName & "','" & transdatetime & "'  )"
'                        Conn.Execute strSql
'                  Else
'                        MsgBox Trim(compPN) & " not  exists, if you want to upload, please Set delflag=N", vbCritical
'                        ErrRow = ErrRow & "," & rCount
'                        xlsBook.Close
'                        xlApp.Quit
'                        Set xlApp = Nothing
'                        Set xlsBook = Nothing
'                        Exit Sub
'                  End If
'
'            ElseIf Left(Trim(.Cells(rCount, 1)), 2) = "CH" Then ''电容的值----- (1204)
'                    strSql = "Select * from QSMS_InSpect_Rule Where Comppn='" & Trim(compPN) & "'"
'                    If rs.State Then rs.Close
'                    rs.Open strSql, Conn, adOpenForwardOnly, adLockReadOnly
'
'                    If rs.EOF = False Then
'                        If delflag = "Y" Then
'                            strSql = "Delete from QSMS_InSpect_Rule where Comppn='" & Trim(compPN) & "'  "
'                            Conn.Execute strSql
'                        Else
'                            MsgBox Trim(compPN) & " already exists, if you want to update, Step1: please delete the P/N (Set delflag=Y) ,Step2: Upload the P/N(Set delflag=N)", vbCritical
'                            ErrRow = ErrRow & "," & rCount
'                            xlsBook.Close
'                            xlApp.Quit
'                            Set xlApp = Nothing
'                            Set xlsBook = Nothing
'                            Exit Sub
'                        End If
'                    ElseIf delflag = "N" Then
'                        digit3 = Mid(Trim(compPN), 3, 1) ''表示多少次方
'                        digit4 = Mid(Trim(compPN), 4, 2)
'                        digit7 = Mid(Trim(compPN), 7, 1)
'                        digit8 = Mid(Trim(compPN), 8, 1)
'
'                        Select Case digit3
'                        Case "-"
'                            CapacitanceValue = Val(digit4) * 0.1 * (10 ^ -12)
'                        Case "+"
'                            CapacitanceValue = Val(digit4) * 1 * (10 ^ -12)
'                        Case Else
'                            CapacitanceValue = (Val(digit4) * (10 ^ Val(digit3))) * (10 ^ -12)
'                        End Select
'
'                    ''料号第7码是0~9就取第8码的值做为上下限
'                        If digit7 = "0" Or digit7 = "1" Or digit7 = "2" Or digit7 = "3" Or digit7 = "4" Or digit7 = "5" Or digit7 = "6" Or digit7 = "7" Or digit7 = "8" Or digit7 = "9" Then
'                            Select Case digit8
'                                Case "B"
'                                    Upper = CapacitanceValue * (1.001)
'                                    Lower = CapacitanceValue * (0.999)
'                                Case "C"
'                                    Upper = CapacitanceValue * (1.0025)
'                                    Lower = CapacitanceValue * (0.9975)
'                                Case "D"
'                                    Upper = CapacitanceValue * (1.005)
'                                    Lower = CapacitanceValue * (0.995)
'                                Case "F"
'                                    Upper = CapacitanceValue * (1.01)
'                                    Lower = CapacitanceValue * (0.99)
'                                Case "G"
'                                    Upper = CapacitanceValue * (1.02)
'                                    Lower = CapacitanceValue * (0.98)
'                                Case "J"
'                                    Upper = CapacitanceValue * (1.05)
'                                    Lower = CapacitanceValue * (0.95)
'                                Case "K"
'                                    Upper = CapacitanceValue * (1.1)
'                                    Lower = CapacitanceValue * (0.9)
'                                Case "M"
'                                    Upper = CapacitanceValue * (1.2)
'                                    Lower = CapacitanceValue * (0.8)
'                                Case "X"
'                                    Upper = CapacitanceValue * (1.3)
'                                    Lower = CapacitanceValue * (1)
'                                Case "Y"
'                                    Upper = CapacitanceValue * (1.1)
'                                    Lower = CapacitanceValue * (1.35)
'                                Case "Z"
'                                    Upper = CapacitanceValue * (1.8)
'                                    Lower = CapacitanceValue * (1.2)
'                                Case "T"
'                                    Upper = CapacitanceValue * (1)
'                                    Lower = CapacitanceValue * (1)
'                                Case Else
'                                    MsgBox Trim(compPN) & " Capacitance of this comppn does not exist", vbCritical
'                                    ErrRow = ErrRow & "," & rCount
'                                    xlsBook.Close
'                                    xlApp.Quit
'                                    Set xlApp = Nothing
'                                    Set xlsBook = Nothing
'                                    Exit Sub
'                            End Select
'                            strSql = "Insert QSMS_InSpect_Rule(Comppn,Upper,Lower,Hz,Volt,Ampere,Unit,ChkNum,BaseQTy,UID,TransDateTime) Values('" & Trim(compPN) & "','" & Trim(Upper) & "','" & Trim(Lower) & "','" & Trim(Hz) & "','" & Trim(Volt) & "','" & Trim(Ampere) & "',N'" & Trim(Unit) & "', '" & Trim(ChkNum) & "','" & Trim(BaseQty) & "','" & g_userName & "','" & transdatetime & "'  )"
'                            Conn.Execute strSql
'                        End If
'
'                        ''料号第7码不是0~9就取第7码的值做为上下限
'                        If digit7 <> "0" And digit7 <> "1" And digit7 <> "2" And digit7 <> "3" And digit7 <> "4" And digit7 <> "5" And digit7 <> "6" And digit7 <> "7" And digit7 <> "8" And digit7 <> "9" Then
'                            Select Case digit7
'                                Case "B"
'                                    Upper = CapacitanceValue * (1.001)
'                                    Lower = CapacitanceValue * (0.999)
'                                Case "C"
'                                    Upper = CapacitanceValue * (1.0025)
'                                    Lower = CapacitanceValue * (0.9975)
'                                Case "D"
'                                    Upper = CapacitanceValue * (1.005)
'                                    Lower = CapacitanceValue * (0.995)
'                                Case "F"
'                                    Upper = CapacitanceValue * (1.01)
'                                    Lower = CapacitanceValue * (0.99)
'                                Case "G"
'                                    Upper = CapacitanceValue * (1.02)
'                                    Lower = CapacitanceValue * (0.98)
'                                Case "J"
'                                    Upper = CapacitanceValue * (1.05)
'                                    Lower = CapacitanceValue * (0.95)
'                                Case "K"
'                                    Upper = CapacitanceValue * (1.1)
'                                    Lower = CapacitanceValue * (0.9)
'                                Case "M"
'                                    Upper = CapacitanceValue * (1.2)
'                                    Lower = CapacitanceValue * (0.8)
'                                Case "X"
'                                    Upper = CapacitanceValue * (1.3)
'                                    Lower = CapacitanceValue * (1)
'                                Case "Y"
'                                    Upper = CapacitanceValue * (1.1)
'                                    Lower = CapacitanceValue * (1.35)
'                                Case "Z"
'                                    Upper = CapacitanceValue * (1.8)
'                                    Lower = CapacitanceValue * (1.2)
'                                Case "T"
'                                    Upper = CapacitanceValue * (1)
'                                    Lower = CapacitanceValue * (1)
'                                Case Else
'                                    MsgBox Trim(compPN) & " Capacitance of this comppn does not exist", vbCritical
'                                    ErrRow = ErrRow & "," & rCount
'                                    xlsBook.Close
'                                    xlApp.Quit
'                                    Set xlApp = Nothing
'                                    Set xlsBook = Nothing
'                                    Exit Sub
'                            End Select
'                            strSql = "Insert QSMS_InSpect_Rule(Comppn,Upper,Lower,Hz,Volt,Ampere,Unit,ChkNum,BaseQTy,UID,TransDateTime) Values('" & Trim(compPN) & "','" & Trim(Upper) & "','" & Trim(Lower) & "','" & Trim(Hz) & "','" & Trim(Volt) & "','" & Trim(Ampere) & "',N'" & Trim(Unit) & "', '" & Trim(ChkNum) & "','" & Trim(BaseQty) & "','" & g_userName & "','" & transdatetime & "'  )"
'                            Conn.Execute strSql
'                        End If
'                  Else
'                        strSql = "Insert QSMS_InSpect_Rule(Comppn,Upper,Lower,Hz,Volt,Ampere,Unit,ChkNum,BaseQTy,UID,TransDateTime) Values('" & Trim(compPN) & "','" & Trim(Upper) & "','" & Trim(Lower) & "','" & Trim(Hz) & "','" & Trim(Volt) & "','" & Trim(Ampere) & "',N'" & Trim(Unit) & "', '" & Trim(ChkNum) & "','" & Trim(BaseQty) & "','" & g_userName & "','" & transdatetime & "'  )"
'                        Conn.Execute strSql
'                  End If
'            End If
        End If
        rCount = rCount + 1
    Wend
End With

strSQL = "EXEC QSMS_UploadInSpectRule" ''(1235)
Conn.Execute (strSQL) ''(1235)

If Len(ErrRow) > 0 Then
    ErrRow = Mid(ErrRow, 2)
End If
strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_InSpect_Rule','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSQL)


xlsBook.Close
xlApp.Quit
Set xlApp = Nothing
Set xlsBook = Nothing
MsgBox ("*** Load  Finish ! Row: " & rCount - 2 & " ***")
End Sub

'''''''''''''''''''''''''''''''''''''''''add by Jing 2007.10.30 (0004)'''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Load_PNAlarmQty(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim COMPPN As String, Unit As String, ErrRow As String

Dim tmppn As String, tmpqty As Integer, tmpFlag As String, tmpDate As String
Dim i As Integer, tmpdel As String
Dim tmpSQL As String, strSQL As String
Dim tmpRS As New ADODB.Recordset, Rs As New ADODB.Recordset

If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
   Exit Sub
End If

On Error GoTo errhandle:
Set xlApp = CreateObject("Excel.Application")
Let xlApp.Visible = False
Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.DisplayAlerts = False

i = 2

strSQL = "select getdate()"
Set Rs = Conn.Execute(strSQL)
tmpDate = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")

With xlsBook.Worksheets(Trim(Shift_Item))
    While Trim(.Cells(i, 1)) <> ""
        tmppn = Trim(.Cells(i, 1))
        tmpqty = Trim(.Cells(i, 2))
        tmpFlag = Trim(.Cells(i, 3))
        
        tmpSQL = "select * from PNAlarmQty where pnprefix='" & tmppn & "'"
        Set tmpRS = Conn.Execute(tmpSQL)
        
        If tmpRS.EOF = False Then
            If tmpFlag = "Y" Then
                tmpSQL = "delete from PNAlarmQty where pnprefix='" & tmppn & "'"
            Else
                tmpSQL = "update PNAlarmQty set Qty='" & tmpqty & "',Uid='" & Trim(g_userName) & "',transdatetime='" & tmpDate & "' where pnprefix='" & tmppn & "'"
            End If
            Conn.Execute (tmpSQL)
        Else
            If tmpFlag = "N" Then
                tmpSQL = "insert into PNAlarmQty(PNPrefix,Qty,UID,Transdatetime) values('" & tmppn & "','" & tmpqty & "','" & Trim(g_userName) & "','" & tmpDate & "')"
                Conn.Execute (tmpSQL)
            End If
        End If
        Set tmpRS = Nothing
        i = i + 1
    Wend
End With

Txt_RowCount = i - 2
xlsBook.Close
xlApp.Quit
Set xlApp = Nothing
Set xlsBook = Nothing
MsgBox "*** Load  finish ! ***"
Exit Sub
errhandle:
    MsgBox Err.Description
End Sub

'''''''''''''''''''''''''''''''''''''''''add by Jing 2007.11.19 (0006)'''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Load_WOScheduleList(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim tmpRow As Integer, tmpcol As Integer
Dim tmpLine As String, tmpWo As String, tmpSeqid As Integer, tmpDate As String
Dim strTmp As String
Dim rsTmp As New ADODB.Recordset
Dim intCount As Integer

On Error GoTo errhandle:
intCount = 0
If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
   Exit Sub
End If

Set xlApp = CreateObject("Excel.Application")
Let xlApp.Visible = False
Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.DisplayAlerts = False

tmpRow = 2
strTmp = "select getdate()"
Set rsTmp = Conn.Execute(strTmp)
tmpDate = Format(rsTmp.Fields(0), "YYYYMMDDHHMMSS")
Set rsTmp = Nothing

With xlsBook.Worksheets(Trim(Shift_Item))
    While Trim(.Cells(tmpRow, 1)) <> ""
        tmpcol = 2
        tmpSeqid = 1
        tmpLine = Trim(.Cells(tmpRow, 1))
        strTmp = "SELECT TOP 1* FROM WOScheduleList WHERE LINE='" & tmpLine & "'"
        Set rsTmp = Conn.Execute(strTmp)
        If rsTmp.EOF = False Then
            strTmp = "DELETE FROM WOScheduleList WHERE LINE='" & tmpLine & "'"
            Conn.Execute (strTmp)
        End If
        Set rsTmp = Nothing
        While Trim(.Cells(tmpRow, tmpcol)) <> ""
            tmpWo = Trim(.Cells(tmpRow, tmpcol))
            strTmp = "INSERT INTO WOScheduleList VALUES('" & tmpLine & "','" & tmpSeqid & "','" & tmpWo & "','" & tmpDate & "','" & Trim(g_userName) & "')"
            Conn.Execute (strTmp)
            tmpSeqid = tmpSeqid + 1
            tmpcol = tmpcol + 1
            intCount = intCount + 1
        Wend
        tmpRow = tmpRow + 1
    Wend
End With
Txt_RowCount = intCount
xlsBook.Close
xlApp.Quit
Set xlApp = Nothing
Set xlsBook = Nothing
MsgBox "*** Load  finish ! ***"
Exit Sub
errhandle:
    MsgBox Err.Description
End Sub


'''''''''''''''''''''''''''''''''''''''''add by Jing 2007.11.26 (0007)'''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub load_XL_WOPlanSeq(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets

Dim intCount As Integer, tmpRow As Integer, i As Integer
Dim tmpDate As String, tmpShift As String, tmpLine As String, tmpWo As String, tmpPlanQty As String, tmpTotalQty As String, tmpSeqid As Integer, tmpTrandt As String, tmpPCBVendorCode As String, tmpBufferQty As String, tmpCapacity As String, tmpFlag As String  ''(0071)
Dim rsTmp As New ADODB.Recordset
Dim strTmp As String, strlog As String
Dim blerr As Boolean, blNum As Boolean, isSide As Boolean, blAlarm As Boolean, blTip As Boolean, isPCBVendorCode As Boolean, isBufferQty6 As Boolean, isBufferQty7 As Boolean, isCapacity7 As Boolean, isCapacity8 As Boolean, isDualLaneMode As Boolean '20110823 Maggie Add DualLaneMode (1072)
Dim checkTime As String, tomorrowDate As String, tmpSide As String, tmpStr As String, tmpLog As String, tmpDualLaneMode As String '20110823 Maggie Add DualLaneMode (1072)
Dim strEDate As String, XLJobDatetime As String, XLplanTime As String
Dim LineArray As String
Dim strTempfac As String  'add by Kevin 20080704 (0033)
Dim strSQL As String
Dim tempFac As String
Dim Rs As New ADODB.Recordset
Dim strStep As String, strTempStep As String, strLogSQL As String, strStepInfo As String, strTempStepInfo As String
Dim WoNum As Integer, tmpRow2 As Integer
Dim a As Integer
Dim strDualLane As String
Dim strLostWO As String ''1278
Dim PCBCompPN As String, WOPN As String, WORev As String


On Error GoTo errhandle:

    isSide = False
    isPCBVendorCode = False
    isBufferQty6 = False
    isBufferQty7 = False
    isCapacity7 = False
    isCapacity8 = False
    isDualLaneMode = False '20110823 Maggie Add DualLaneMode (1072)
    blerr = False
    blNum = True
    blAlarm = False
    blTip = False
    strStep = ""
    PCBCompPN = ""
    WOPN = ""
    WORev = ""
    
    '''1173  如果有定义，则根据定义的时间点检查是否允许上传，没有定义则还是用之前的方式检查
    strSQL = "select dbo.formatdate(getdate(),'yyyy-mm-dd')+' '+Value as XLTime from QSMS_ProConfig where Line='ALL' AND Station='XL' AND [KEY]='XLTime'"
    Set Rs = Conn.Execute(strSQL)
        If Rs.EOF = False Then
            Do While Not Rs.EOF
                XLJobDatetime = Rs!XLTime
                strTmp = "select getdate() as time"
                Set rsTmp = Conn.Execute(strTmp)
                XLplanTime = rsTmp!Time
                XLplanTime = "2014-07-15 17:00:00"
                If Abs(DateDiff("n", XLJobDatetime, XLplanTime)) < 15 Then
                    MsgBox ("Please do not upload XL_WOPlanSeq before or after 15 minutes of XL_Job execution!")
                    Exit Sub
                End If
                Rs.MoveNext
            Loop
        Else
            strSQL = "select top 1 * from QSMS_CheckBom where workOrder='XLJob'order by datetime desc"            ''''1053
            Set Rs = Conn.Execute(strSQL)
            If Not Rs.EOF Then
                XLJobDatetime = Rs!DateTime
                strTmp = "select getdate() as time"
                Set rsTmp = Conn.Execute(strTmp)
                XLplanTime = rsTmp!Time
                If DateDiff("n", XLJobDatetime, XLplanTime) < 30 Then
                    MsgBox ("Please do not upload XL_WOPlanSeq before or after 30 minutes of XL_Job execution!")
                    Exit Sub
                End If
            End If
    
        End If
    ''''''''''add by Kevin 20080704  (0033)
    ''''1. 获得厂区,如果有多个厂区则要求User选择
    strSQL = "exec GetFactory"
    Set Rs = Conn.Execute(strSQL)
    strTempfac = Rs!result
    If InStr(strTempfac, "or") > 0 Then
        tempFac = InputBox("Please Input Factory:     " & strTempfac, "Input Factory")
    Else
        tempFac = strTempfac
    End If
    
    ''''''(0068)
    strStep = "Step1:"
    strStepInfo = strStep & "Select Factory[" & tempFac & "]->"
    
    ''''2. 获得日期
    strTmp = "select getdate()"
    Set rsTmp = Conn.Execute(strTmp)
    tmpTrandt = Format(rsTmp.Fields(0), "YYYYMMDDHHNNSS")
    tomorrowDate = Format(Now() + 1, "YYYYMMDD")
    
    Set rsTmp = Nothing
    tmpDate = Left(Right(txtFilePath, 12), 8)
    If Len(Trim(tmpDate)) <> 8 Then
        MsgBox "Please check the file name foramt:" & txtFilePath, vbCritical, "Information"
        Exit Sub
    End If
    
    strStep = "Step2:"
    strStepInfo = strStepInfo & strStep & "Get Date[" & tmpDate & "]->"
    
    strEDate = Format(DateAdd("d", 1, Format(tmpDate, "0000-00-00")), "yyyymmdd")
    If Len(Trim(strEDate)) <> 8 Then
        MsgBox ("EndDate format is error! Please call QMS !")
        Exit Sub
    End If
       
    For i = 1 To Len(tmpDate)
        If IsNumeric(Mid(tmpDate, i, 1)) = False Then
            blNum = False
        End If
    Next i
           
    If Not blNum Then
        MsgBox ("The filename format is error!")
        Exit Sub
    End If
        
    ''''4. 开始上传排程
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    intCount = 0
    tmpRow = 2
    tmpRow2 = 2
        
    With xlsBook.Worksheets(Trim(Shift_Item))

        If Trim(.Cells(2, 1)) = "" Or Trim(.Cells(2, 2)) = "" Then
            MsgBox "The format of this file is wrong,please check", vbCritical, "Format Error"
            GoTo ErrDeal
        Else
            ''''''''''''''''''''''删除表中该天已存在的数据''''''''''''''''''''''''''
            'add factory=tempfac by kevin 20080704
            ''''3. By 日期厂区删除已经上传的排程
'            strTmp = "DELETE FROM WO_AssignPN_Vendor WHERE WO IN (SELECT WO FROM XL_WOPlanSeq WHERE date='" & tmpDate & "' and factory = '" & tempFac & "') "
'            Conn.Execute (strTmp)
            
            strTmp = "DELETE FROM XL_WOPlanSeq WHERE date='" & tmpDate & "' and factory = '" & tempFac & "' "   'add a condition factory = tempfac by kevin 20080704 (0033)
            Conn.Execute (strTmp)
            
            strStep = "Step3:"
            strStepInfo = strStepInfo & strStep & "DelPlan[Fac:" & tempFac & ",Date:" & tmpDate & "]"
            
            strLogSQL = "insert into qms_log(system_name,event_no,sn,user_name,desc1,trans_date) values('Upload_XL_WOPlanSeq','1','','" & Trim(g_userName) & "',N'" & strStepInfo & "',dbo.formatdate(getdate(),'YYYYMMDDHHNNSS'))"
            Conn.Execute strLogSQL
            
            strTmp = "DELETE FROM XL_WOPlanSeqBySide WHERE WO+Side in (select WO+SIDE from XL_WOPlanSeq_TraceBySide WHERE date='" & tmpDate & "')"
            Conn.Execute (strTmp)
            strTmp = "DELETE FROM XL_WOPlanSeq_TraceBySide WHERE date='" & tmpDate & "'"
            Conn.Execute (strTmp)
            
        End If
        
        If UCase(Trim(.Cells(1, 7))) = "SIDE" Then isSide = True
        
         ''(1284)
        If UCase(Trim(.Cells(1, 7))) = "PCBVENDORCODE" Then
            isPCBVendorCode = True
        End If
            
        If UCase(Trim(.Cells(1, 7))) = "BUFFERQTY" Then isBufferQty6 = True ''(1015)
        If UCase(Trim(.Cells(1, 8))) = "BUFFERQTY" Then isBufferQty7 = True ''(1015)
        If UCase(Trim(.Cells(1, 8))) = "CAPACITY" Then isCapacity7 = True ''(1015)
        If UCase(Trim(.Cells(1, 9))) = "CAPACITY" Then isCapacity8 = True ''(1015)
        If UCase(Trim(.Cells(1, 10))) = "DUALLANEMODE" Then isDualLaneMode = True '20110823 Maggie Add DualLaneMode (1072)
        
        While Trim(.Cells(tmpRow, 1)) <> "" And Trim(.Cells(tmpRow, 2)) <> "" 'And Trim(.Cells(tmpRow, 3)) <> ""
            
            ''''4.1 将记录步骤的变量清空
            ''strTempStep = ""
            ''strTempStepInfo = ""
            
            ''''4.2 同一班别同一条线的多个工单按照产销上传的排程排序
            If tmpShift = Trim(.Cells(tmpRow, 2)) And tmpLine = Trim(.Cells(tmpRow, 1)) Then
                tmpSeqid = tmpSeqid + 1
            Else
                tmpSeqid = 1
            End If
            
            ''''4.3 获得排程中的数据 Shift/Line/WO/PlanQty/TotalQty
            tmpShift = Trim(.Cells(tmpRow, 2))
            tmpLine = Trim(.Cells(tmpRow, 1))
            tmpWo = Trim(.Cells(tmpRow, 3))
            tmpPlanQty = Trim(.Cells(tmpRow, 4))
            tmpTotalQty = Trim(.Cells(tmpRow, 5))
            tmpFlag = Trim(.Cells(tmpRow, 6))   ''(0071)(1135)
            
            If isSide = True Then
                tmpSide = Trim(.Cells(tmpRow, 7))
            End If
            
            If isPCBVendorCode = True Then
                tmpPCBVendorCode = Trim(.Cells(tmpRow, 7))
            End If
            
            strSQL = "select top 1 MBPN, CompPN, REV from SAP_BOM where Work_Order='" & tmpWo & "' and CompPN like '%DA%'"
            Set Rs = Conn.Execute(strSQL)
            If Not Rs.EOF Then
                WOPN = Rs!MBPN
                WORev = Rs!Rev
                PCBCompPN = Rs!COMPPN
            End If
             
            If isBufferQty6 = True Then
                tmpBufferQty = Trim(.Cells(tmpRow, 7))
            End If
            
            If isBufferQty7 = True Then
                tmpBufferQty = Trim(.Cells(tmpRow, 8))
            End If
            
            If isCapacity7 = True Then
                tmpCapacity = Trim(.Cells(tmpRow, 8))
            End If
            
            If isCapacity8 = True Then
                tmpCapacity = Trim(.Cells(tmpRow, 9))
            End If
            
            '20110823 Maggie Add DualLaneMode
            If isDualLaneMode = True Then
                tmpDualLaneMode = Trim(.Cells(tmpRow, 10))
            End If
            
            If UCase(tmpWo) <> "SKIP" Then
                ''''4.4 只有Release OK并且CheckBom Pass 的工单才能上传
                strTmp = "select wo,line,checkbompassdatetime from sap_wo_list where wo='" & tmpWo & "'"
                Set rsTmp = Conn.Execute(strTmp)
                If rsTmp.EOF Then
                    blerr = True
                    .Cells(tmpRow, 8).Interior.ColorIndex = 3
                    .Cells(tmpRow, 8).Interior.Pattern = xlSolid
                    .Cells(tmpRow, 8) = "This WO is error or not release to sap!"
                Else
                    If Trim(rsTmp("line")) <> tmpLine Then
                        blerr = True
                        .Cells(tmpRow, 8).Interior.ColorIndex = 3
                        .Cells(tmpRow, 8).Interior.Pattern = xlSolid
                        .Cells(tmpRow, 8) = "Can not find the wo in this line !"
                    Else
                        If Trim(rsTmp("checkbompassdatetime")) = "" Then
                            blerr = True
                            .Cells(tmpRow, 8).Interior.ColorIndex = 3
                            .Cells(tmpRow, 8).Interior.Pattern = xlSolid
                            .Cells(tmpRow, 8) = "This WO not checkbompass !"
                        End If
                    End If
                End If
                '=============================================0049
                ''''4.5 检查线别是否定义
                strTmp = "select 0 from XL_WOPlanLine"
                If rsTmp.State Then rsTmp.Close
                Set rsTmp = Conn.Execute(strTmp)
                If rsTmp.EOF = False Then
                    strTmp = "select 0 from XL_WOPlanLine where line='" & tmpLine & "'"
                    If rsTmp.State Then rsTmp.Close
                    Set rsTmp = Conn.Execute(strTmp)
                    If rsTmp.EOF Then
                        MsgBox ("Line： " & tmpLine & " is new Line，please comfirm！")
                    End If
                End If
                '=============================================0049
                
                If blerr = False Then
                    '===========================add by kane 2008.03.17 (0021)==================================
                    strTmp = "select 0 from sap_wo_list a where exists(select 0 from sap_wo_list b where a.[group]=b.[group] and b.wo='" & tmpWo & "') and checkbompassdatetime='' "
                    Set rsTmp = Conn.Execute(strTmp)
                    If rsTmp.EOF = False Then
                        blerr = True
                        .Cells(tmpRow, 8).Interior.ColorIndex = 3
                        .Cells(tmpRow, 8).Interior.Pattern = xlSolid
                        .Cells(tmpRow, 8) = "There some wo did not check bom pass in this PCB "
                    End If
    
                    '=========================ADD By Kevin 2008.07.07(0035)=========================================
                    If tmpLine <> "L2C" Then
                        strTmp = "select 0 from sap_wo_list a,work_center b,site c where a.work_center like replace(b.work_center,'*','_') and b.plant=c.plant and c.factory='" & tempFac & "'and a.wo='" & tmpWo & "'"
                        Set rsTmp = Conn.Execute(strTmp)
                        If rsTmp.EOF = True Then
                            blerr = True
                            .Cells(tmpRow, 9).Interior.ColorIndex = 3
                            .Cells(tmpRow, 9).Interior.Pattern = xlSolid
                            .Cells(tmpRow, 9) = "The wo did not suit to be the factory! "
                        End If
                    End If
                    
                    ''''''''''''''Added by Jing 2008.03.31  (0023)''''''''''''''''
                    If Trim(tmpSide) = "" Then
                        strTmp = "select 0 from XL_WOPlanSeq where Factory='" & Trim(tempFac) & "' and  Date='" & tmpDate & "' and Shift='" & tmpShift & "' and Line='" & tmpLine & "' and wo in (select wo from sap_wo_list a where exists(select 0 from sap_wo_list b where a.[group]=b.[group] and wo='" & tmpWo & "'))"
                        Set rsTmp = Conn.Execute(strTmp)
                        If rsTmp.EOF = False Then
                            blerr = True
                            .Cells(tmpRow, 10).Interior.ColorIndex = 3
                            .Cells(tmpRow, 10).Interior.Pattern = xlSolid '0060
                            .Cells(tmpRow, 10) = "Upload Fail: A work order has been uploaded to the same PCBGroup. For work orders of the same PCBGroup, if the Date/Shift/Line/Factory is the same, you only need to upload one of the work orders.！"

'                            strTmp = "DELETE FROM WO_AssignPN_Vendor WHERE WO IN (SELECT WO FROM XL_WOPlanSeq WHERE date='" & tmpDate & "' and factory = '" & tempFac & "') "
'                            Conn.Execute (strTmp)
                            
                            '''(0039)
                            strTmp = "DELETE FROM XL_WOPlanSeq WHERE date='" & tmpDate & "' and factory = '" & tempFac & "' "
                            Conn.Execute (strTmp)
                            
                            strTmp = "DELETE FROM XL_WOPlanSeq_trace WHERE date='" & tmpDate & "' and factory = '" & tempFac & "' "
                            Conn.Execute (strTmp)
                            
                            MsgBox ("Upload fail ! Please upload again !" & Chr(13) & Chr(13) & "Notice the red fail part in the EXCEL file !"), vbCritical
                            Let xlApp.Visible = True
                            Exit Sub
                        End If
                    Else
                        strTmp = "select 0 from XL_WOPlanSeq_TraceBySide where Side='" & Trim(tmpSide) & "' and Date='" & tmpDate & "' and Shift='" & tmpShift & "' and Line='" & tmpLine & "' and wo in (select wo from sap_wo_list a where exists(select 0 from sap_wo_list b where a.[group]=b.[group] and wo='" & tmpWo & "'))"
                        Set rsTmp = Conn.Execute(strTmp)
                        If rsTmp.EOF = False Then
                            blerr = True
                            .Cells(tmpRow, 10).Interior.ColorIndex = 3
                            .Cells(tmpRow, 10).Interior.Pattern = xlSolid
                            .Cells(tmpRow, 10) = "Upload Fail: In its Group have been upload WO , or Multiple Same WO in a Date/Shift/Line/Side."
   
'                            strTmp = "DELETE FROM WO_AssignPN_Vendor WHERE WO IN (SELECT WO FROM XL_WOPlanSeq WHERE date='" & tmpDate & "' and factory = '" & tempFac & "') "
'                            Conn.Execute (strTmp)

                            '''(0039)
                            strTmp = "DELETE FROM XL_WOPlanSeq WHERE date='" & tmpDate & "' and factory = '" & tempFac & "' "
                            Conn.Execute (strTmp)
                            
                            strTmp = "DELETE FROM XL_WOPlanSeq_trace WHERE date='" & tmpDate & "' and factory = '" & tempFac & "' "
                            Conn.Execute (strTmp)
                            
                            strTmp = "DELETE FROM XL_WOPlanSeqBySide WHERE WO+Side in (select WO+SIDE from XL_WOPlanSeq_TraceBySide WHERE date='" & tmpDate & "')"
                            Conn.Execute (strTmp)
                            
                            strTmp = "DELETE FROM XL_WOPlanSeq_TraceBySide WHERE date='" & tmpDate & "'"
                            Conn.Execute (strTmp)
                                                    
                            MsgBox ("Upload fail ! Please upload again !" & Chr(13) & Chr(13) & "Notice the red fail part in the EXCEL file !"), vbCritical
                            Let xlApp.Visible = True
                            Exit Sub
                        End If
                    End If
                    
                    ''''''''''''''Updated by Jing 2008.01.08    (0011)'''''''''''''''
                    If (blerr = False) And (tmpTotalQty <> "0") Then
                        strTmp = "select * from qsms_wogroup where work_order='" & tmpWo & "'"
                        Set rsTmp = Conn.Execute(strTmp)
                        If rsTmp.EOF Then
                            blerr = True
                            .Cells(tmpRow, 8).Interior.ColorIndex = 3
                            .Cells(tmpRow, 8).Interior.Pattern = xlSolid
                            .Cells(tmpRow, 8) = "This WO is not define the GroupID !"
                        Else
                            If CDbl(tmpPlanQty) > CDbl(tmpTotalQty) Then    '''Updated by Jing  (0016)'''
                                blerr = True
                                .Cells(tmpRow, 8).Interior.ColorIndex = 3
                                .Cells(tmpRow, 8).Interior.Pattern = xlSolid
                                .Cells(tmpRow, 8) = "PlanQty must be less than TotalQty !"
                            Else
                                
                                ''''''added by Jing (0037)''''''
                                If tmpSide = "" Then
                                    strTmp = "select top 1 inputQty from XL_WOPlanSeq_trace where date='" & tmpDate & "' and shift='" & tmpShift & "' and line='" & tmpLine & "' and wo='" & tmpWo & "' and factory='" & tempFac & "' order by TransDateTime desc "
                                    Set rsTmp = Conn.Execute(strTmp)
                                    If rsTmp.EOF = False Then
                                        If CDbl(rsTmp("inputQty")) > CDbl(tmpTotalQty) Then
                                            blTip = True
                                            .Cells(tmpRow, 11).Interior.ColorIndex = 37
                                            .Cells(tmpRow, 11).Interior.Pattern = xlSolid
                                            .Cells(tmpRow, 11) = "WOInput Qty is less than Previous WOInput that was uploaded scheduling !"
'                                            Return
                                        End If
                                    End If
                                End If
                            
                                strTmp = "DELETE FROM XL_WOPlanSeq WHERE date='" & tmpDate & "' and shift='" & tmpShift & "' and line='" & tmpLine & "' and wo='" & tmpWo & "' and factory='" & tempFac & "' "  'add a condition factory = tempfac by kevin 20080704 (0033)
                                Conn.Execute (strTmp)
                                Set rsTmp = Nothing
                                
                                If tmpShift = "D" Then
                                    If tmpPCBVendorCode <> "" And PCBCompPN <> "" Then
                                        strTmp = "DELETE QSMS_Assigned_CompPN WHERE WO = '" + tmpWo + "' AND CompPN = '" + PCBCompPN + "' "
                                        Conn.Execute (strTmp)
                                    
                                        strTmp = "INSERT INTO QSMS_Assigned_CompPN(WO, CompPN, VendorCode, DateCode, LotCode, Flag, Owner, UID, TransDateTime)" & _
                                                 "VALUES ('" & tmpWo & "','" & PCBCompPN & "','" & tmpPCBVendorCode & "','','','Y','" & Trim(g_userName) & "','" & Trim(g_userName) & "',dbo.formatdate(GETDATE(),'yyyymmddhhnnss'))"
                                        Conn.Execute (strTmp)
                                    End If
                                    
                                    strTmp = "INSERT INTO XL_WOPlanSeq(date,shift,line,wo,planqty,seqid,transdatetime,opid,inputqty,begindatetime,enddatetime,Factory,BufferQty,Capacity,Flag) " & _
                                            "VALUES('" & tmpDate & "','" & tmpShift & "','" & tmpLine & "','" & tmpWo & "','" & tmpPlanQty & "','" & _
                                            tmpSeqid & "',dbo.formatdate(GETDATE(),'yyyymmddhhnnss'),'" & Trim(g_userName) & "','" & tmpTotalQty & "','" & tmpDate & "0740','" & tmpDate & "1940','" & tempFac & "','" & tmpBufferQty & "','" & tmpCapacity & "','" & tmpFlag & "')"  'add a condition factory = tempfac by kevin 20080704 (0033)     ''(0071)
    
                                    strlog = "INSERT INTO XL_WOPlanSeq_Trace(date,shift,line,wo,planqty,seqid,transdatetime,opid,inputqty,begindatetime,enddatetime,Factory,BufferQty,Capacity) " & _
                                            "VALUES('" & tmpDate & "','" & tmpShift & "','" & tmpLine & "','" & tmpWo & "','" & tmpPlanQty & "','" & _
                                            tmpSeqid & "',dbo.formatdate(GETDATE(),'yyyymmddhhnnss'),'" & Trim(g_userName) & "','" & tmpTotalQty & "','" & tmpDate & "0740','" & tmpDate & "1940','" & tempFac & "','" & tmpBufferQty & "','" & tmpCapacity & "')"  'add a condition factory = tempfac by kevin 20080704 (0033)
    
                                    Conn.Execute (strTmp)
                                    Conn.Execute (strlog)
    
                                    '''updated by Jing (0028)'''
                                    If tmpSide <> "" Then
                                        tmpStr = "Exec XL_WOPlanSeq_BySide '" & tmpWo & "','" & tmpSide & "','" & tmpTotalQty & "','" & Trim(g_userName) & "','" & tmpTrandt & "','" & tmpDate & "','" & tmpShift & "','" & tmpLine & "','" & tmpSeqid & "','" & tmpPlanQty & "'"
                                        Conn.Execute (tmpStr)
                                    End If
                                Else
                                    If tmpPCBVendorCode <> "" And PCBCompPN <> "" Then
                                        strTmp = "DELETE QSMS_Assigned_CompPN WHERE WO = '" + tmpWo + "' AND CompPN = '" + PCBCompPN + "' "
                                        Conn.Execute (strTmp)
                                    
                                        strTmp = "INSERT INTO QSMS_Assigned_CompPN(WO, CompPN, VendorCode, DateCode, LotCode, Flag, Owner, UID, TransDateTime)" & _
                                                 "VALUES ('" & tmpWo & "','" & PCBCompPN & "','" & tmpPCBVendorCode & "','','','Y','" & Trim(g_userName) & "','" & Trim(g_userName) & "',dbo.formatdate(GETDATE(),'yyyymmddhhnnss'))"
                                        Conn.Execute (strTmp)
                                    End If
                                    
                                    strTmp = "INSERT INTO XL_WOPlanSeq(date,shift,line,wo,planqty,seqid,transdatetime,opid,inputqty,begindatetime,enddatetime,Factory,BufferQty,Capacity,Flag) " & _
                                            "VALUES('" & tmpDate & "','" & tmpShift & "','" & tmpLine & "','" & tmpWo & "','" & tmpPlanQty & "','" & _
                                            tmpSeqid & "',dbo.formatdate(GETDATE(),'yyyymmddhhnnss'),'" & Trim(g_userName) & "','" & tmpTotalQty & "','" & tmpDate & "1940','" & strEDate & "0740','" & tempFac & "','" & tmpBufferQty & "','" & tmpCapacity & "','" & tmpFlag & "')"      ''''''Update by Jing 2008.02.01 (0014)'''''''' 'add a condition factory = tempfac by kevin 20080704 (0033)    ''(0071)
    
                                    strlog = "INSERT INTO XL_WOPlanSeq_Trace(date,shift,line,wo,planqty,seqid,transdatetime,opid,inputqty,begindatetime,enddatetime,Factory,BufferQty,Capacity) " & _
                                            "VALUES('" & tmpDate & "','" & tmpShift & "','" & tmpLine & "','" & tmpWo & "','" & tmpPlanQty & "','" & _
                                            tmpSeqid & "',dbo.formatdate(GETDATE(),'yyyymmddhhnnss'),'" & Trim(g_userName) & "','" & tmpTotalQty & "','" & tmpDate & "1940','" & strEDate & "0740','" & tempFac & "','" & tmpBufferQty & "','" & tmpCapacity & "')"     ''''''Update by Jing 2008.02.01 (0014)'''''''''add a condition factory = tempfac by kevin 20080704 (0033)
    
                                    Conn.Execute (strTmp)
                                    Conn.Execute (strlog)
                                    
                                    ''updated by Jing (0028)'''
                                    If tmpSide <> "" Then
                                        tmpStr = "Exec XL_WOPlanSeq_BySide '" & tmpWo & "','" & tmpSide & "','" & tmpTotalQty & "','" & Trim(g_userName) & "','" & tmpTrandt & "','" & tmpDate & "','" & tmpShift & "','" & tmpLine & "','" & tmpSeqid & "','" & tmpPlanQty & "'"
                                        Conn.Execute (tmpStr)
                                    End If
                                End If
                                
                                '20110823 Maggie DualLaneMode (1072)
                                If isDualLaneMode = True Then
                                    strDualLane = "Exec XL_WOPlanSeq_DualLane " & sq(Trim(tmpWo)) & " ," & sq(Trim(tmpDualLaneMode)) & "," & sq(Trim(g_userName)) & "," & sq(Trim(tmpTrandt))
                                    Conn.Execute (strDualLane)
                                End If
                                
                            End If
                        End If
                    End If
                    intCount = intCount + 1
                End If
            Else '0047
                tmpStr = "UPDATE XL_WOPlanLine SET UploadFlag='Y' where Line='" & tmpLine & "'and Factory='" & tempFac & "'"
                Conn.Execute (tmpStr)
            End If
            
            If blerr = True And blAlarm = False Then blAlarm = True
            blerr = False
            tmpRow = tmpRow + 1
        Wend
        If Mid(tmpTrandt, 9, 2) >= "08" And Mid(tmpTrandt, 9, 2) < "10" Then ''1278
               If strChk_XL_WOPlanSeq = "Y" Then
                  strTmp = "SELECT DISTINCT A.WO FROM XL_WOPlanSeq_Trace A WHERE A.[DATE]='" & tmpDate & "'"
                  strTmp = strTmp + "AND A.TransDateTime<'" & tmpTrandt & "' AND NOT EXISTS "
                  strTmp = strTmp + "(SELECT 0 FROM XL_WOPlanSeq_Trace B WHERE A.[DATE]=B.[DATE] AND A.WO=B.WO "
                  strTmp = strTmp + "AND A.LINE=B.LINE AND A.Shift=B.Shift AND A.factory=B.factory AND B.TransDateTime>='" & tmpTrandt & "')"
                  Set rsTmp = Conn.Execute(strTmp)
                  If Not rsTmp.EOF Then
                     strLostWO = ""
                     Do While Not rsTmp.EOF
                        strLostWO = Trim(rsTmp!WO) + ";" + strLostWO
                        rsTmp.MoveNext
                     Loop
                  End If
                  If strLostWO <> "" Then
                      MsgBox (strLostWO + " Lost In XL_WOPlan,Pls Check!"), vbCritical
                  End If
                  
               End If
            End If
            
        tmpStr = "select Line from XL_WOPlanLine where line not in(select Line from XL_WOPlanSeq where date='" & tmpDate & "' AND FACTORY='" & tempFac & "' and shift='" & tmpShift & "') and UploadFlag='N' AND FACTORY='" & tempFac & "'"
        Set rsTmp = Conn.Execute(tmpStr)
        While Not rsTmp.EOF
            LineArray = LineArray + Trim(rsTmp!Line) + ","
            rsTmp.MoveNext
        Wend
        tmpStr = "UPDATE XL_WOPlanLine SET UploadFlag='N' where UploadFlag='Y'"
        Conn.Execute (tmpStr)
        If Trim(LineArray) <> "" Then
            MsgBox "Line=" & LineArray & "have not Upload,Please check it;"
        End If
        
    End With
    
    Txt_RowCount = intCount
    
    If blAlarm Then
        MsgBox ("Upload fail ! Please upload again !" & Chr(13) & " Notice the red fail part in the EXCEL file !"), vbCritical
        Let xlApp.Visible = True
        Exit Sub
    Else
        If blTip Then
            MsgBox ("Upload OK ! Notice the blue part in the EXCEL file !"), vbInformation
            Let xlApp.Visible = True
            Exit Sub
        End If
    End If
ErrDeal:
    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox ("*** Load  finish ! ***")
    Exit Sub

errhandle:
    MsgBox Err.Description
End Sub
'''''''''''''''''''''''''''''''''''''''''add by Jing 2007.11.26 (0007)'''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub load_XL_WOPlanSeqShiftID(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets

Dim intCount As Integer, tmpRow As Integer, i As Integer
Dim tmpDate As String, tmpShift As String, tmpLine As String, tmpWo As String, tmpPlanQty As String, tmpTotalQty As String, tmpSeqid As Integer, tmpTrandt As String, tmpBufferQty As String, tmpCapacity As String, tmpFlag As String  '''(1133)
Dim rsTmp As New ADODB.Recordset
Dim strTmp As String, strlog As String
Dim blerr As Boolean, blNum As Boolean, isSide As Boolean, blAlarm As Boolean, blTip As Boolean, isBufferQty7 As Boolean, isBufferQty8 As Boolean, isCapacity8 As Boolean, isCapacity9 As Boolean, isDualLaneMode As Boolean '20110823 Maggie Add DualLaneMode (1072)(1133)
Dim checkTime As String, tomorrowDate As String, tmpSide As String, tmpStr As String, tmpLog As String, tmpDualLaneMode As String '20110823 Maggie Add DualLaneMode (1072)
Dim strEDate As String, XLJobDatetime As String, XLplanTime As String
Dim LineArray As String
Dim strTempfac As String  'add by Kevin 20080704 (0033)
Dim strSQL As String
Dim tempFac As String
Dim Rs As New ADODB.Recordset
Dim strStep As String, strTempStep As String, strLogSQL As String, strStepInfo As String, strTempStepInfo As String
Dim WoNum As Integer, tmpRow2 As Integer
Dim a As Integer
Dim strDualLane As String
Dim ShiftIDLine As Integer
Dim tmpShiftID As String
Dim isShiftId As Boolean        ''(1121)

On Error GoTo errhandle:

    isShiftId = False
    isSide = False
    isBufferQty7 = False
    isBufferQty8 = False
    isCapacity8 = False
    isCapacity9 = False         ''(1133)
    isDualLaneMode = False '20110823 Maggie Add DualLaneMode (1072)
    blerr = False
    blNum = True
    blAlarm = False
    blTip = False
    strStep = ""

    strSQL = "select top 1 * from QSMS_CheckBom where workOrder='XLJob'order by datetime desc"            ''''1053
    Set Rs = Conn.Execute(strSQL)
    If Not Rs.EOF Then
        XLJobDatetime = Rs!DateTime
        strTmp = "select getdate() as time"
        Set rsTmp = Conn.Execute(strTmp)
        XLplanTime = rsTmp!Time
        If DateDiff("n", XLJobDatetime, XLplanTime) < 30 Then
            MsgBox ("Please do not upload XL_WOPlanSeq before or after 30 minutes of XL_Job execution！")
            Exit Sub
        End If
    End If
    
    ''''''''''add by Kevin 20080704  (0033)
    ''''1. 获得厂区,如果有多个厂区则要求User选择
    strSQL = "exec GetFactory"
    Set Rs = Conn.Execute(strSQL)
    strTempfac = Rs!result
    If InStr(strTempfac, "or") > 0 Then
        tempFac = InputBox("Please Input Factory:     " & strTempfac, "Input Factory")
    Else
        tempFac = strTempfac
    End If
    
    ''''''(0068)
    strStep = "Step1:"
    strStepInfo = strStep & "Select Factory[" & tempFac & "]->"
    
    ''''2. 获得日期
    strTmp = "select getdate()"
    Set rsTmp = Conn.Execute(strTmp)
    tmpTrandt = Format(rsTmp.Fields(0), "YYYYMMDDHHNNSS")
    tomorrowDate = Format(Now() + 1, "YYYYMMDD")
    
    Set rsTmp = Nothing
    tmpDate = Left(Right(txtFilePath, 12), 8)
    If Len(Trim(tmpDate)) <> 8 Then
        MsgBox "Please check the file name foramt:" & txtFilePath, vbCritical, "Information"
        Exit Sub
    End If
    
    strStep = "Step2:"
    strStepInfo = strStepInfo & strStep & "Get Date[" & tmpDate & "]->"
    
    strEDate = Format(DateAdd("d", 1, Format(tmpDate, "0000-00-00")), "yyyymmdd")
    If Len(Trim(strEDate)) <> 8 Then
        MsgBox ("EndDate format is error! Please call QMS !")
        Exit Sub
    End If
       
    For i = 1 To Len(tmpDate)
        If IsNumeric(Mid(tmpDate, i, 1)) = False Then
            blNum = False
        End If
    Next i
           
    If Not blNum Then
        MsgBox ("The filename format is error!")
        Exit Sub
    End If
        
    ''''4. 开始上传排程
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    intCount = 0
    tmpRow = 2
    tmpRow2 = 2
        
    With xlsBook.Worksheets(Trim(Shift_Item))

        If Trim(.Cells(2, 1)) = "" Or Trim(.Cells(2, 2)) = "" Then
            MsgBox "The format of this file is wrong,please check", vbCritical, "Format Error"
            GoTo ErrDeal
        Else
            ''''''''''''''''''''''删除表中该天已存在的数据''''''''''''''''''''''''''
            'add factory=tempfac by kevin 20080704
            ''''3. By 日期厂区删除已经上传的排程
            strTmp = "DELETE FROM XL_WOPlanSeq WHERE date='" & tmpDate & "' and factory = '" & tempFac & "' "   'add a condition factory = tempfac by kevin 20080704 (0033)
            Conn.Execute (strTmp)
            
            strStep = "Step3:"
            strStepInfo = strStepInfo & strStep & "DelPlan[Fac:" & tempFac & ",Date:" & tmpDate & "]"
            
            strLogSQL = "insert into qms_log(system_name,event_no,sn,user_name,desc1,trans_date) values('Upload_XL_WOPlanSeq','1','','" & Trim(g_userName) & "',N'" & strStepInfo & "',dbo.formatdate(getdate(),'YYYYMMDDHHNNSS'))"
            Conn.Execute strLogSQL
            
            strTmp = "DELETE FROM XL_WOPlanSeqBySide WHERE WO+Side in (select WO+SIDE from XL_WOPlanSeq_TraceBySide WHERE date='" & tmpDate & "')"
            Conn.Execute (strTmp)
            strTmp = "DELETE FROM XL_WOPlanSeq_TraceBySide WHERE date='" & tmpDate & "'"
            Conn.Execute (strTmp)
        End If
        
        If UCase(Trim(.Cells(1, 7))) = "SIDE" Then isSide = True
            
'        If UCase(Trim(.Cells(1, 6))) = "BUFFERQTY" Then isBufferQty6 = True ''(1015)
'        If UCase(Trim(.Cells(1, 7))) = "BUFFERQTY" Then isBufferQty7 = True ''(1015)
'        If UCase(Trim(.Cells(1, 7))) = "CAPACITY" Then isCapacity7 = True ''(1015)
'        If UCase(Trim(.Cells(1, 8))) = "CAPACITY" Then isCapacity8 = True ''(1015)
'        If UCase(Trim(.Cells(1, 9))) = "DUALLANEMODE" Then isDualLaneMode = True '20110823 Maggie Add DualLaneMode (1072)
'        If UCase(Trim(.Cells(1, 6))) = "SHIFTID" Then isShiftId = True: ShiftIDLine = 6
'        If UCase(Trim(.Cells(1, 7))) = "SHIFTID" Then isShiftId = True: ShiftIDLine = 7
'        If UCase(Trim(.Cells(1, 8))) = "SHIFTID" Then isShiftId = True: ShiftIDLine = 8
'        If UCase(Trim(.Cells(1, 9))) = "SHIFTID" Then isShiftId = True: ShiftIDLine = 9
'        If UCase(Trim(.Cells(1, 10))) = "SHIFTID" Then isShiftId = True: ShiftIDLine = 10       '''(1121)
         
        If UCase(Trim(.Cells(1, 7))) = "BUFFERQTY" Then isBufferQty7 = True ''(1015)
        If UCase(Trim(.Cells(1, 8))) = "BUFFERQTY" Then isBufferQty8 = True ''(1015)
        If UCase(Trim(.Cells(1, 8))) = "CAPACITY" Then isCapacity8 = True ''(1015)
        If UCase(Trim(.Cells(1, 9))) = "CAPACITY" Then isCapacity9 = True ''(1015)
        If UCase(Trim(.Cells(1, 10))) = "DUALLANEMODE" Then isDualLaneMode = True '20110823 Maggie Add DualLaneMode (1072)
        If UCase(Trim(.Cells(1, 7))) = "SHIFTID" Then isShiftId = True: ShiftIDLine = 7
        If UCase(Trim(.Cells(1, 8))) = "SHIFTID" Then isShiftId = True: ShiftIDLine = 8
        If UCase(Trim(.Cells(1, 9))) = "SHIFTID" Then isShiftId = True: ShiftIDLine = 9
        If UCase(Trim(.Cells(1, 10))) = "SHIFTID" Then isShiftId = True: ShiftIDLine = 10
        If UCase(Trim(.Cells(1, 11))) = "SHIFTID" Then isShiftId = True: ShiftIDLine = 11       '''(1121)(1133)
        
        While Trim(.Cells(tmpRow, 1)) <> "" And Trim(.Cells(tmpRow, 2)) <> "" 'And Trim(.Cells(tmpRow, 3)) <> ""
            
            ''''4.1 将记录步骤的变量清空
            ''strTempStep = ""
            ''strTempStepInfo = ""
            
            ''''4.2 同一班别同一条线的多个工单按照产销上传的排程排序
            If tmpShift = Trim(.Cells(tmpRow, 2)) And tmpLine = Trim(.Cells(tmpRow, 1)) Then
                tmpSeqid = tmpSeqid + 1
            Else
                tmpSeqid = 1
            End If
            
            ''''4.3 获得排程中的数据 Shift/Line/WO/PlanQty/TotalQty
            tmpShift = Trim(.Cells(tmpRow, 2))
            tmpLine = Trim(.Cells(tmpRow, 1))
            tmpWo = Trim(.Cells(tmpRow, 3))
            tmpPlanQty = Trim(.Cells(tmpRow, 4))
            tmpTotalQty = Trim(.Cells(tmpRow, 5))
            tmpFlag = Trim(.Cells(tmpRow, 6))
            
            If isSide = True Then
                tmpSide = Trim(.Cells(tmpRow, 7))
            End If
             
            If isBufferQty7 = True Then
                tmpBufferQty = Trim(.Cells(tmpRow, 7))
            End If
            
            If isBufferQty8 = True Then
                tmpBufferQty = Trim(.Cells(tmpRow, 8))
            End If
            
            If isCapacity8 = True Then
                tmpCapacity = Trim(.Cells(tmpRow, 8))
            End If
            
            If isCapacity9 = True Then
                tmpCapacity = Trim(.Cells(tmpRow, 9))
            End If          ''''''(1133)
            
            If isShiftId = True Then
                If tmpShift = "D" Then
                    tmpShiftID = Trim(.Cells(tmpRow, ShiftIDLine))
                Else
                    strTmp = "select * from XL_TypeDateTime "
                    Set rsTmp = Conn.Execute(strTmp)
                    If rsTmp.EOF Then
                        MsgBox ("Upload fail ! Please upload again !" & Chr(13) & Chr(13) & "Notice!! the ShiftID is empty !"), vbCritical
                        Exit Sub
                    Else
                        tmpShiftID = CStr(12 / CInt(rsTmp("XL_Type")) + CInt(Trim(.Cells(tmpRow, ShiftIDLine))))
                    End If
                End If
            End If              ''(1121)
            
            '20110823 Maggie Add DualLaneMode
            If isDualLaneMode = True Then
                tmpDualLaneMode = Trim(.Cells(tmpRow, 10))      ''''(1133)
            End If
            
            If UCase(tmpWo) <> "SKIP" Then
                ''''4.4 只有Release OK并且CheckBom Pass 的工单才能上传
                strTmp = "select wo,line,checkbompassdatetime from sap_wo_list where wo='" & tmpWo & "'"
                Set rsTmp = Conn.Execute(strTmp)
                If rsTmp.EOF Then
                    blerr = True
                    .Cells(tmpRow, 8).Interior.ColorIndex = 3
                    .Cells(tmpRow, 8).Interior.Pattern = xlSolid
                    .Cells(tmpRow, 8) = "This WO is error or not release to sap!"
                Else
                    If Trim(rsTmp("line")) <> tmpLine Then
                        blerr = True
                        .Cells(tmpRow, 8).Interior.ColorIndex = 3
                        .Cells(tmpRow, 8).Interior.Pattern = xlSolid
                        .Cells(tmpRow, 8) = "Can not find the wo in this line !"
                    Else
                        If Trim(rsTmp("checkbompassdatetime")) = "" Then
                            blerr = True
                            .Cells(tmpRow, 8).Interior.ColorIndex = 3
                            .Cells(tmpRow, 8).Interior.Pattern = xlSolid
                            .Cells(tmpRow, 8) = "This WO not checkbompass !"
                        End If
                    End If
                End If
                '=============================================0049
                ''''4.5 检查线别是否定义
                strTmp = "select 0 from XL_WOPlanLine"
                If rsTmp.State Then rsTmp.Close
                Set rsTmp = Conn.Execute(strTmp)
                If rsTmp.EOF = False Then
                    strTmp = "select 0 from XL_WOPlanLine where line='" & tmpLine & "'"
                    If rsTmp.State Then rsTmp.Close
                    Set rsTmp = Conn.Execute(strTmp)
                    If rsTmp.EOF Then
                        MsgBox ("Line " & tmpLine & " is new Line, please comfirm!")
                    End If
                End If
                '=============================================0049
                
                If blerr = False Then
                    '===========================add by kane 2008.03.17 (0021)==================================
                    strTmp = "select 0 from sap_wo_list a where exists(select 0 from sap_wo_list b where a.[group]=b.[group] and b.wo='" & tmpWo & "') and checkbompassdatetime='' "
                    Set rsTmp = Conn.Execute(strTmp)
                    If rsTmp.EOF = False Then
                        blerr = True
                        .Cells(tmpRow, 8).Interior.ColorIndex = 3
                        .Cells(tmpRow, 8).Interior.Pattern = xlSolid
                        .Cells(tmpRow, 8) = "There some wo did not check bom pass in this PCB "
                    End If
    
                    '=========================ADD By Kevin 2008.07.07(0035)=========================================
                    strTmp = "select 0 from sap_wo_list a,work_center b,site c where a.work_center like replace(b.work_center,'*','_') and b.plant=c.plant and c.factory='" & tempFac & "'and a.wo='" & tmpWo & "'"
                    Set rsTmp = Conn.Execute(strTmp)
                    If rsTmp.EOF = True Then
                        blerr = True
                        .Cells(tmpRow, 9).Interior.ColorIndex = 3
                        .Cells(tmpRow, 9).Interior.Pattern = xlSolid
                        .Cells(tmpRow, 9) = "The wo did not suit to be the factory! "
                    End If
                    
                    ''''''''''''''Added by Jing 2008.03.31  (0023)''''''''''''''''
                    If Trim(tmpSide) = "" Then
                        strTmp = "select 0 from XL_WOPlanSeq where Factory='" & Trim(tempFac) & "' and  Date='" & tmpDate & "' and ShiftID='" & tmpShiftID & "' and Line='" & tmpLine & "' and wo in (select wo from sap_wo_list a where exists(select 0 from sap_wo_list b where a.[group]=b.[group] and wo='" & tmpWo & "'))"  ''(1121)
                        Set rsTmp = Conn.Execute(strTmp)
                        If rsTmp.EOF = False Then
                            blerr = True
                            .Cells(tmpRow, 10).Interior.ColorIndex = 3
                            .Cells(tmpRow, 10).Interior.Pattern = xlSolid '0060
                            .Cells(tmpRow, 10) = "Upload Fail: One work order has been uploaded to the same work order PCBGroup. For work orders of the same PCBGroup, if the Date/Shift/Line/Factory is the same, only one work order needs to be uploaded."

                            '''(0039)
                            strTmp = "DELETE FROM XL_WOPlanSeq WHERE date='" & tmpDate & "' and factory = '" & tempFac & "' "
                            Conn.Execute (strTmp)
                            
                            strTmp = "DELETE FROM XL_WOPlanSeq_trace WHERE date='" & tmpDate & "' and factory = '" & tempFac & "' "
                            Conn.Execute (strTmp)
                            
                            MsgBox ("Upload fail ! Please upload again !" & Chr(13) & Chr(13) & "Notice the red fail part in the EXCEL file !"), vbCritical
                            Let xlApp.Visible = True
                            Exit Sub
                        End If
                    Else
                        strTmp = "select 0 from XL_WOPlanSeq_TraceBySide where Side='" & Trim(tmpSide) & "' and Date='" & tmpDate & "' and Shift='" & tmpShift & "' and Line='" & tmpLine & "' and wo in (select wo from sap_wo_list a where exists(select 0 from sap_wo_list b where a.[group]=b.[group] and wo='" & tmpWo & "'))"
                        Set rsTmp = Conn.Execute(strTmp)
                        If rsTmp.EOF = False Then
                            blerr = True
                            .Cells(tmpRow, 10).Interior.ColorIndex = 3
                            .Cells(tmpRow, 10).Interior.Pattern = xlSolid
                            .Cells(tmpRow, 10) = "Upload Fail: In its Group have been upload WO , or Multiple Same WO in a Date/Shift/Line/Side."

                            '''(0039)
                            strTmp = "DELETE FROM XL_WOPlanSeq WHERE date='" & tmpDate & "' and factory = '" & tempFac & "' "
                            Conn.Execute (strTmp)
                            
                            strTmp = "DELETE FROM XL_WOPlanSeq_trace WHERE date='" & tmpDate & "' and factory = '" & tempFac & "' "
                            Conn.Execute (strTmp)
                            
                            strTmp = "DELETE FROM XL_WOPlanSeqBySide WHERE WO+Side in (select WO+SIDE from XL_WOPlanSeq_TraceBySide WHERE date='" & tmpDate & "')"
                            Conn.Execute (strTmp)
                            
                            strTmp = "DELETE FROM XL_WOPlanSeq_TraceBySide WHERE date='" & tmpDate & "'"
                            Conn.Execute (strTmp)
                            
                            MsgBox ("Upload fail ! Please upload again !" & Chr(13) & Chr(13) & "Notice the red fail part in the EXCEL file !"), vbCritical
                            Let xlApp.Visible = True
                            Exit Sub
                        End If
                    End If
                    
                    ''''''''''''''Updated by Jing 2008.01.08    (0011)'''''''''''''''
                    If (blerr = False) And (tmpTotalQty <> "0") Then
                        strTmp = "select * from qsms_wogroup where work_order='" & tmpWo & "'"
                        Set rsTmp = Conn.Execute(strTmp)
                        If rsTmp.EOF Then
                            blerr = True
                            .Cells(tmpRow, 8).Interior.ColorIndex = 3
                            .Cells(tmpRow, 8).Interior.Pattern = xlSolid
                            .Cells(tmpRow, 8) = "This WO is not define the GroupID !"
                        Else
                            If CDbl(tmpPlanQty) > CDbl(tmpTotalQty) Then    '''Updated by Jing  (0016)'''
                                blerr = True
                                .Cells(tmpRow, 8).Interior.ColorIndex = 3
                                .Cells(tmpRow, 8).Interior.Pattern = xlSolid
                                .Cells(tmpRow, 8) = "PlanQty must be less than TotalQty !"
                            Else
                                
                                ''''''added by Jing (0037)''''''
                                If tmpSide = "" Then
                                    strTmp = "select top 1 inputQty from XL_WOPlanSeq_trace where date='" & tmpDate & "' and shiftID='" & tmpShiftID & "' and line='" & tmpLine & "' and wo='" & tmpWo & "' and factory='" & tempFac & "' order by TransDateTime desc "         ''(1121)
                                    Set rsTmp = Conn.Execute(strTmp)
                                    If rsTmp.EOF = False Then
                                        If CDbl(rsTmp("inputQty")) > CDbl(tmpTotalQty) Then
                                            blTip = True
                                            .Cells(tmpRow, 11).Interior.ColorIndex = 37
                                            .Cells(tmpRow, 11).Interior.Pattern = xlSolid
                                            .Cells(tmpRow, 11) = "WOInput Qty is less than Previous WOInput that was uploaded scheduling !"
                                            Return
                                        End If
                                    End If
                                End If
                                
                                strTmp = "DELETE FROM XL_WOPlanSeq WHERE date='" & tmpDate & "' and shiftID='" & tmpShiftID & "' and line='" & tmpLine & "' and wo='" & tmpWo & "' and factory='" & tempFac & "' "  'add a condition factory = tempfac by kevin 20080704 (0033)(1121)
                                Conn.Execute (strTmp)
                                Set rsTmp = Nothing
                                
                                If tmpShift = "D" Then
                                    strTmp = "INSERT INTO XL_WOPlanSeq(date,shift,line,wo,planqty,seqid,transdatetime,opid,inputqty,begindatetime,enddatetime,Factory,BufferQty,Capacity,ShiftID,Flag) " & _
                                            "VALUES('" & tmpDate & "','" & tmpShift & "','" & tmpLine & "','" & tmpWo & "','" & tmpPlanQty & "','" & _
                                            tmpSeqid & "','" & tmpTrandt & "','" & Trim(g_userName) & "','" & tmpTotalQty & "','" & tmpDate & "0740','" & tmpDate & "1940','" & tempFac & "','" & tmpBufferQty & "','" & tmpCapacity & "','" & tmpShiftID & "','" & tmpFlag & "')"  'add a condition factory = tempfac by kevin 20080704 (0033)(1121)
    
                                    strlog = "INSERT INTO XL_WOPlanSeq_Trace(date,shift,line,wo,planqty,seqid,transdatetime,opid,inputqty,begindatetime,enddatetime,Factory,BufferQty,Capacity,ShiftID) " & _
                                            "VALUES('" & tmpDate & "','" & tmpShift & "','" & tmpLine & "','" & tmpWo & "','" & tmpPlanQty & "','" & _
                                            tmpSeqid & "','" & tmpTrandt & "','" & Trim(g_userName) & "','" & tmpTotalQty & "','" & tmpDate & "0740','" & tmpDate & "1940','" & tempFac & "','" & tmpBufferQty & "','" & tmpCapacity & "','" & tmpShiftID & "')"  'add a condition factory = tempfac by kevin 20080704 (0033)(1121)
    
                                    Conn.Execute (strTmp)
                                    Conn.Execute (strlog)
    
                                    '''updated by Jing (0028)'''
                                    If tmpSide <> "" Then
                                        tmpStr = "Exec XL_WOPlanSeq_BySide '" & tmpWo & "','" & tmpSide & "','" & tmpTotalQty & "','" & Trim(g_userName) & "','" & tmpTrandt & "','" & tmpDate & "','" & tmpShift & "','" & tmpLine & "','" & tmpSeqid & "','" & tmpPlanQty & "'"
                                        Conn.Execute (tmpStr)
                                    End If
                                Else
                                    strTmp = "INSERT INTO XL_WOPlanSeq(date,shift,line,wo,planqty,seqid,transdatetime,opid,inputqty,begindatetime,enddatetime,Factory,BufferQty,Capacity,ShiftID,Flag) " & _
                                            "VALUES('" & tmpDate & "','" & tmpShift & "','" & tmpLine & "','" & tmpWo & "','" & tmpPlanQty & "','" & _
                                            tmpSeqid & "','" & tmpTrandt & "','" & Trim(g_userName) & "','" & tmpTotalQty & "','" & tmpDate & "1940','" & strEDate & "0740','" & tempFac & "','" & tmpBufferQty & "','" & tmpCapacity & "','" & tmpShiftID & "','" & tmpFlag & "')"      ''''''Update by Jing 2008.02.01 (0014)'''''''' 'add a condition factory = tempfac by kevin 20080704 (0033)(1121)
    
                                    strlog = "INSERT INTO XL_WOPlanSeq_Trace(date,shift,line,wo,planqty,seqid,transdatetime,opid,inputqty,begindatetime,enddatetime,Factory,BufferQty,Capacity,ShiftID) " & _
                                            "VALUES('" & tmpDate & "','" & tmpShift & "','" & tmpLine & "','" & tmpWo & "','" & tmpPlanQty & "','" & _
                                            tmpSeqid & "','" & tmpTrandt & "','" & Trim(g_userName) & "','" & tmpTotalQty & "','" & tmpDate & "1940','" & strEDate & "0740','" & tempFac & "','" & tmpBufferQty & "','" & tmpCapacity & "','" & tmpShiftID & "')"     ''''''Update by Jing 2008.02.01 (0014)'''''''''add a condition factory = tempfac by kevin 20080704 (0033)(1121)
    
                                    Conn.Execute (strTmp)
                                    Conn.Execute (strlog)
                                    
                                    ''updated by Jing (0028)'''
                                    If tmpSide <> "" Then
                                        tmpStr = "Exec XL_WOPlanSeq_BySide '" & tmpWo & "','" & tmpSide & "','" & tmpTotalQty & "','" & Trim(g_userName) & "','" & tmpTrandt & "','" & tmpDate & "','" & tmpShift & "','" & tmpLine & "','" & tmpSeqid & "','" & tmpPlanQty & "'"
                                        Conn.Execute (tmpStr)
                                    End If
                                End If
                                
                                '20110823 Maggie DualLaneMode (1072)
                                If isDualLaneMode = True Then
                                    strDualLane = "Exec XL_WOPlanSeq_DualLane " & sq(Trim(tmpWo)) & " ," & sq(Trim(tmpDualLaneMode)) & "," & sq(Trim(g_userName)) & "," & sq(Trim(tmpTrandt))
                                    Conn.Execute (strDualLane)
                                End If
                                
                            End If
                        End If
                    End If
                    intCount = intCount + 1
                End If
            Else '0047
                tmpStr = "UPDATE XL_WOPlanLine SET UploadFlag='Y' where Line='" & tmpLine & "'and Factory='" & tempFac & "'"
                Conn.Execute (tmpStr)
            End If
            
            If blerr = True And blAlarm = False Then blAlarm = True
            blerr = False
            tmpRow = tmpRow + 1
        Wend
        
        tmpStr = "select Line from XL_WOPlanLine where line not in(select Line from XL_WOPlanSeq where date='" & tmpDate & "' AND FACTORY='" & tempFac & "' and shift='" & tmpShift & "') and UploadFlag='N' AND FACTORY='" & tempFac & "'"
        Set rsTmp = Conn.Execute(tmpStr)
        While Not rsTmp.EOF
            LineArray = LineArray + Trim(rsTmp!Line) + ","
            rsTmp.MoveNext
        Wend
        tmpStr = "UPDATE XL_WOPlanLine SET UploadFlag='N' where UploadFlag='Y'"
        Conn.Execute (tmpStr)
        If Trim(LineArray) <> "" Then
            MsgBox "Line=" & LineArray & "have not Upload,Please check it;"
        End If
        
    End With
    
    Txt_RowCount = intCount
    
    If blAlarm Then
        MsgBox ("Upload fail ! Please upload again !" & Chr(13) & " Notice the red fail part in the EXCEL file !"), vbCritical
        Let xlApp.Visible = True
        Exit Sub
    Else
        If blTip Then
            MsgBox ("Upload OK ! Notice the blue part in the EXCEL file !"), vbInformation
            Let xlApp.Visible = True
            Exit Sub
        End If
    End If
ErrDeal:
    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox ("*** Load  finish ! ***")
    Exit Sub

errhandle:
    MsgBox Err.Description
End Sub
Private Sub XL_MATERIALTOWHID(Shift_Item As String) '(0013)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim FixPNstr As String
Dim str1 As String
Dim i As Integer
Dim flag As Integer
Dim rsTmp As ADODB.Recordset
Dim strSQL As String, PrefixPN As String, WareHouseID As String, transdatetime As String, strCompPN As String, strRemark As String, strUID As String
Dim tmpRow As Integer
Dim blerr As Boolean
Dim sFactory As String
Dim delflag As String

blerr = False

On Error GoTo errhandle:
    
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    
    tmpRow = 2
    strUID = Trim(g_userName)
    
    With xlsBook.Worksheets(Trim(Shift_Item))
        While Trim(.Cells(tmpRow, 1)) <> "" And Trim(.Cells(tmpRow, 2)) <> ""
            PrefixPN = Trim(.Cells(tmpRow, 1))
            WareHouseID = Trim(.Cells(tmpRow, 2))
            ''Denver       2008.08.04      Add Factory  (0038)
            sFactory = Trim(.Cells(tmpRow, 3))
            delflag = Trim(.Cells(tmpRow, 4))
            
            FixPNstr = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789+<-%"    '''''1066
            For i = 1 To Len(PrefixPN)
                str1 = Mid(PrefixPN, i, 1)
                flag = InStr(FixPNstr, str1)
                Debug.Print flag
                If (flag < 1) Then
                    MsgBox ("The PrefixPN include " & str1 & ", please check it. ")
                    .Cells(tmpRow, 5) = "Please check it !"
                    .Cells(tmpRow, 5).Interior.ColorIndex = 3
                    .Cells(tmpRow, 5).Interior.Pattern = xlSolid
                    
                    Exit Sub
                End If
            Next

            If delflag = "Y" Then   ''''(0069)
                strSQL = "EXEC XL_MaterialToWHID '" & PrefixPN & "','" & WareHouseID & "','" & strUID & "'" & "," & sq(Trim(sFactory)) & ",'" & delflag & "'"
            Else
                strSQL = "EXEC XL_MaterialToWHID '" & PrefixPN & "','" & WareHouseID & "','" & strUID & "'" & "," & sq(Trim(sFactory))
            End If
            
            Set rsTmp = Conn.Execute(strSQL)
            
            If rsTmp("result") = 1 Then
                blerr = True
                .Cells(tmpRow, 6).Interior.ColorIndex = 3
                .Cells(tmpRow, 6).Interior.Pattern = xlSolid
                .Cells(tmpRow, 6) = rsTmp("description")
            End If
            tmpRow = tmpRow + 1
        Wend
    End With
    
    If blerr Then
        MsgBox ("Upload fail ! Notice the red fail part in the EXCEL file !"), vbCritical
        Let xlApp.Visible = True
        Exit Sub
    End If

    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox ("*** Load  Finish ! ***")
    Exit Sub
    
errhandle:
    MsgBox Err.Description
End Sub



'''''''''''''''''''''''''''''''''''''''''add by Jing 2007.12.05 (0008)'''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub load_XL_ImplementPN(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets

Dim rsTmp As ADODB.Recordset
Dim strSQL As String, StrPN As String, strUID As String
Dim tmpRow As Integer
Dim blerr As Boolean
blerr = False

On Error GoTo errhandle:
    
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    
    tmpRow = 2
    strUID = Trim(g_userName)
    
    With xlsBook.Worksheets(Trim(Shift_Item))
        While Trim(.Cells(tmpRow, 1)) <> ""
            StrPN = Trim(.Cells(tmpRow, 1))
            
            strSQL = "EXEC XL_UploadToImplementPN '" & StrPN & "','" & strUID & "'"
            Set rsTmp = Conn.Execute(strSQL)
            
            If rsTmp("result") = 0 Then
                blerr = True
                .Cells(tmpRow, 2).Interior.ColorIndex = 3
                .Cells(tmpRow, 2).Interior.Pattern = xlSolid
                .Cells(tmpRow, 2) = "This PN upload fail !"
            End If
            tmpRow = tmpRow + 1
        Wend
    End With
    
    If blerr Then
        MsgBox ("Upload fail ! Notice the red fail part in the EXCEL file !"), vbCritical
        Let xlApp.Visible = True
        Exit Sub
    End If

    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox ("*** Load  Finish ! ***")
    Exit Sub
    
errhandle:
    MsgBox Err.Description
End Sub
Private Sub Load_NoMachineDropCompPN(Shift_Item As String)              ''(0042)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim COMPPN As String, Unit As String, ErrRow As String

Dim tmppn As String, tmpqty As Integer, tmpFlag As String, tmpDate As String
Dim i As Integer, tmpdel As String
Dim tmpSQL As String, strSQL As String
Dim tmpRS As New ADODB.Recordset, Rs As New ADODB.Recordset

If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
   Exit Sub
End If

On Error GoTo errhandle:
Set xlApp = CreateObject("Excel.Application")
Let xlApp.Visible = False
Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.DisplayAlerts = False

i = 2

strSQL = "select getdate()"
Set Rs = Conn.Execute(strSQL)
tmpDate = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")

With xlsBook.Worksheets(Trim(Shift_Item))
    While Trim(.Cells(i, 1)) <> ""
        tmppn = Trim(.Cells(i, 1))
        tmpFlag = Trim(.Cells(i, 2))
        
        tmpSQL = "select * from QSMS_UnCheckCompPN where CompPN='" & tmppn & "' and Type='NOMDrop'"
        Set tmpRS = Conn.Execute(tmpSQL)
        
        If tmpRS.EOF = False Then
            If tmpFlag = "Y" Then
                tmpSQL = "delete from QSMS_UnCheckCompPN where CompPN='" & tmppn & "' and Type='NOMDrop'"
                Conn.Execute (tmpSQL)
            End If
        Else
            If tmpFlag <> "Y" Then
                tmpSQL = "insert into QSMS_UnCheckCompPN(Type,CompPN,UserID,Transdatetime) values('NOMDrop','" & tmppn & "','" & Trim(g_userName) & "','" & tmpDate & "')"
                Conn.Execute (tmpSQL)
            End If
        End If
        Set tmpRS = Nothing
        i = i + 1
    Wend
End With

Txt_RowCount = i - 2
xlsBook.Close
xlApp.Quit
Set xlApp = Nothing
Set xlsBook = Nothing
MsgBox "*** Load  finish ! ***"
Exit Sub                                ''(0042)
errhandle:
    MsgBox Err.Description
End Sub

Private Sub Load_PCB_SingleCompPN(Shift_Item As String)              ''(0063)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim COMPPN As String, Unit As String, ErrRow As String

Dim tmppn As String, tmpqty As Integer, tmpFlag As String, tmpDate As String
Dim i As Integer, tmpdel As String
Dim tmpSQL As String, strSQL As String
Dim tmpRS As New ADODB.Recordset, Rs As New ADODB.Recordset

If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
   Exit Sub
End If

On Error GoTo errhandle:
Set xlApp = CreateObject("Excel.Application")
Let xlApp.Visible = False
Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.DisplayAlerts = False

i = 2

strSQL = "select getdate()"
Set Rs = Conn.Execute(strSQL)
tmpDate = Format(Rs.Fields(0), "YYYYMMDDHHMMSS")

With xlsBook.Worksheets(Trim(Shift_Item))
    While Trim(.Cells(i, 1)) <> ""
        tmppn = Trim(.Cells(i, 1))
        tmpFlag = Trim(.Cells(i, 2))
        
        tmpSQL = "select * from QSMS_UnCheckCompPN where CompPN='" & tmppn & "' and Type='PCB_SingleCompPN'"
        Set tmpRS = Conn.Execute(tmpSQL)
        
        If tmpRS.EOF = False Then
            If tmpFlag = "Y" Then
                tmpSQL = "delete from QSMS_UnCheckCompPN where CompPN='" & tmppn & "' and Type='PCB_SingleCompPN'"
                Conn.Execute (tmpSQL)
            End If
        Else
            If tmpFlag <> "Y" Then
                tmpSQL = "insert into QSMS_UnCheckCompPN(Type,CompPN,UserID,Transdatetime) values('PCB_SingleCompPN','" & tmppn & "','" & Trim(g_userName) & "','" & tmpDate & "')"
                Conn.Execute (tmpSQL)
            End If
        End If
        Set tmpRS = Nothing
        i = i + 1
    Wend
End With

Txt_RowCount = i - 2
xlsBook.Close
xlApp.Quit
Set xlApp = Nothing
Set xlsBook = Nothing
MsgBox "*** Load  finish ! ***"
Exit Sub                                ''(0063)
errhandle:
    MsgBox Err.Description
End Sub

'''''''''''''''''''''''''''''''''''''''''add by Jing 2007.12.17 (0009)'''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Load_XL_WOPN(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets

Dim rsTmp As ADODB.Recordset
Dim strSQL As String, strWO As String, StrMBPN As String, strUpCompPN As String, strCompPN As String, strRemark As String, strUID As String
Dim tmpRow As Integer
Dim blerr As Boolean
blerr = False

On Error GoTo errhandle:
    
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    
    tmpRow = 2
    strUID = Trim(g_userName)
    
    With xlsBook.Worksheets(Trim(Shift_Item))
        While Trim(.Cells(tmpRow, 1)) <> "" And Trim(.Cells(tmpRow, 2)) <> "" And Trim(.Cells(tmpRow, 3)) <> "" And Trim(.Cells(tmpRow, 4)) <> ""
            strWO = Trim(.Cells(tmpRow, 1))
            StrMBPN = Trim(.Cells(tmpRow, 2))
            strUpCompPN = Trim(.Cells(tmpRow, 3))
            strCompPN = Trim(.Cells(tmpRow, 4))
            strRemark = Trim(.Cells(tmpRow, 5))
            
            strSQL = "select * from sap_bom where work_order='" & strWO & "'"
            Set rsTmp = Conn.Execute(strSQL)
            
            If rsTmp.EOF Then
                blerr = True
                .Cells(tmpRow, 6).Interior.ColorIndex = 3
                .Cells(tmpRow, 6).Interior.Pattern = xlSolid
                .Cells(tmpRow, 6) = "This WO is error or not in SAP_BOM!"
            Else
                strSQL = "EXEC XL_UploadToWOPN '" & strWO & "','" & StrMBPN & "','" & strUpCompPN & "','" & strCompPN & "','" & strRemark & "','" & strUID & "'"
                Set rsTmp = Conn.Execute(strSQL)
                
                If rsTmp("result") = 0 Then
                    blerr = True
                    .Cells(tmpRow, 7).Interior.ColorIndex = 3
                    .Cells(tmpRow, 7).Interior.Pattern = xlSolid
                    .Cells(tmpRow, 7) = "This WO upload fail !"
                End If
            End If
            tmpRow = tmpRow + 1
        Wend
    End With
    
    If blerr Then
        MsgBox ("Upload fail ! Notice the red fail part in the EXCEL file !"), vbCritical
        Let xlApp.Visible = True
        Exit Sub
    End If

    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox ("*** Load  Finish ! ***")
    Exit Sub
    
errhandle:
    MsgBox Err.Description
End Sub

Private Sub load_DailySchedule(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim strFactory As String
Dim rsTmp As ADODB.Recordset
Dim strSQL As String, strLine As String, strUID As String
Dim PN As String, WO As String, WO_QTY As String, Rev As String
Dim tmpRow As Integer
Dim i As Integer
Dim WorkDate(6) As String
Dim Qty(6) As String
Dim TempLine As String
Dim blerr As Boolean
Dim strToday As String
Dim FileName As String
Dim StrBU As String, TempBU As String
Dim M As Integer, n As Integer, h As Integer
blerr = False

On Error GoTo errhandle:
    M = 1
    n = 0
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    FileName = Trim(Dir(txtFilePath))
    tmpRow = 3
    strUID = Trim(g_userName)
    strFactory = Left(FileName, InStr(FileName, " ") - 1) '0052

    strSQL = "select dbo.formatdate(getdate(),'YYYYMMDD') AS TODAY"
'    If rsTmp.State Then rsTmp.Close
    Set rsTmp = Conn.Execute(strSQL)
    strToday = Trim(rsTmp!Today)

    With xlsBook.Worksheets(Trim(Shift_Item))
    
        For h = 0 To 5000
            TempBU = Trim(.Cells(3 + h, 1))
            TempBU = Mid(TempBU, InStr(TempBU, "-") + 1, 10)
            If TempBU <> "" And StrBU <> TempBU Then '0053
                StrBU = TempBU
                strSQL = "EXEC XL_UpdateDailySchedule '" & strFactory & "','" & StrBU & "','1'"
                Set rsTmp = Conn.Execute(strSQL)
            End If
        Next h
        strSQL = "EXEC XL_UpdateDailySchedule "
        Set rsTmp = Conn.Execute(strSQL)
        
        While Replace(Trim(CStr(.Cells(1, M))), "/", "") <> strToday And M < 256 '0056
            M = M + 1
        Wend
        For i = 0 To 6
            WorkDate(i) = Replace(Trim(.Cells(1, M + i)), "/", "")
        Next i

        While PN <> "*" And tmpRow < 10000
            TempLine = Trim(.Cells(tmpRow, 1)) '0052
            If Trim(TempLine) <> "" Then
                 strLine = TempLine
            End If
            PN = Trim(.Cells(tmpRow, 2))
            WO = Trim(.Cells(tmpRow, 3))
            WO_QTY = Trim(.Cells(tmpRow, 4))
            Rev = Trim(.Cells(tmpRow, 5))
            For i = 0 To 6
                Qty(i) = Trim(.Cells(tmpRow, M + i))
            Next i
            If PN <> "" Then
                For i = 0 To 6
                    If Qty(i) <> "" Then
                        Rev = Replace(Rev, "'", " ")
                        strSQL = "EXEC XL_UploadDailySchedule '" & strFactory & "','" & strLine & "','" & PN & "','" & WO & "','" & WO_QTY & "','" & Rev & "','" & WorkDate(i) & "','" & Qty(i) & "','" & strUID & "'"
                        Debug.Print strSQL
                        Set rsTmp = Conn.Execute(strSQL)
                        If rsTmp("result") = 0 Then
                            blerr = True
                            .Cells(tmpRow, 13).Interior.ColorIndex = 3
                            .Cells(tmpRow, 13).Interior.Pattern = xlSolid
                            .Cells(tmpRow, 13) = "This Line upload fail !Please check the Line;"
                        End If
                        Qty(i) = ""
                    End If
                Next i
             End If
            tmpRow = tmpRow + 1
        Wend
    End With
    
    If blerr Then
        MsgBox ("Upload fail ! Notice the red fail part in the EXCEL file !"), vbCritical
        Let xlApp.Visible = True
        Exit Sub
    End If

    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox ("*** Load  Finish ! Row: " & tmpRow & " ***")
    Exit Sub
    
errhandle:
    MsgBox Err.Description
End Sub

Private Sub load_WOPlanLine(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim strFactory As String
Dim rsTmp As ADODB.Recordset
Dim strSQL As String, strLine As String, strUID As String
Dim tmpRow As Integer
Dim blerr As Boolean
blerr = False

On Error GoTo errhandle:
    
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    
    tmpRow = 2
    strUID = Trim(g_userName)
'    StrSQL = "exec GetFactory"
'    Set rsTmp = Conn.Execute(StrSQL)
'    StrFactory = rsTmp!result

    strSQL = "Truncate table XL_WOPlanLine"
    Conn.Execute (strSQL)
    
    With xlsBook.Worksheets(Trim(Shift_Item))
        While Trim(.Cells(tmpRow, 1)) <> ""
            strLine = Trim(.Cells(tmpRow, 1))
            strFactory = Trim(.Cells(tmpRow, 2))
            
            strSQL = "EXEC XL_UploadWOPlanLine '" & strLine & "','" & strFactory & "','" & strUID & "'"
            Set rsTmp = Conn.Execute(strSQL)
            
            If rsTmp("result") = 0 Then
                blerr = True
                .Cells(tmpRow, 2).Interior.ColorIndex = 3
                .Cells(tmpRow, 2).Interior.Pattern = xlSolid
                .Cells(tmpRow, 2) = "This Line upload fail !Please check the Line;"
            End If
            tmpRow = tmpRow + 1
        Wend
    End With
    
    If blerr Then
        MsgBox ("Upload fail ! Notice the red fail part in the EXCEL file !"), vbCritical
        Let xlApp.Visible = True
        Exit Sub
    End If

    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox ("*** Load  Finish ! ***")
    Exit Sub
    
errhandle:
    MsgBox Err.Description
End Sub
'''''''''''''''''''''''''''''''''''''''''add by Jing 2008.02.19 (0017)'''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub load_XL_PNOneByOne(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets

Dim rsTmp As ADODB.Recordset
Dim strSQL As String, StrPN As String, strUID As String
Dim tmpRow As Integer
Dim blerr As Boolean
blerr = False

On Error GoTo errhandle:
    
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    
    tmpRow = 2
    strUID = Trim(g_userName)
    
    strSQL = "Truncate table XL_PNOneByOne"
    Conn.Execute (strSQL)
    
    With xlsBook.Worksheets(Trim(Shift_Item))
        While Trim(.Cells(tmpRow, 1)) <> ""
            StrPN = Trim(.Cells(tmpRow, 1))
            
            strSQL = "EXEC XL_UploadToPNOneByOne '" & StrPN & "','" & strUID & "'"
            Set rsTmp = Conn.Execute(strSQL)
            
            If rsTmp("result") = 0 Then
                blerr = True
                .Cells(tmpRow, 2).Interior.ColorIndex = 3
                .Cells(tmpRow, 2).Interior.Pattern = xlSolid
                .Cells(tmpRow, 2) = "This PN upload fail !"
            End If
            tmpRow = tmpRow + 1
        Wend
    End With
    
    If blerr Then
        MsgBox ("Upload fail ! Notice the red fail part in the EXCEL file !"), vbCritical
        Let xlApp.Visible = True
        Exit Sub
    End If

    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox ("*** Load  Finish ! ***")
    Exit Sub
    
errhandle:
    MsgBox Err.Description
End Sub


'''''''''''''''''''''''''''''''''''''''''add by Jing 2008.03.02 (0019)'''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub load_XL_PNInterval(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets

Dim rsTmp As ADODB.Recordset
Dim strSQL As String, StrMBPN As String, StrPN As String, StrInterval As String, strUID As String, strFactory As String
Dim tmpRow As Integer
Dim blerr As Boolean
blerr = False

On Error GoTo errhandle:
    
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    
    tmpRow = 2
    strUID = Trim(g_userName)
    
    With xlsBook.Worksheets(Trim(Shift_Item))
        While Trim(.Cells(tmpRow, 1)) <> "" And Trim(.Cells(tmpRow, 2)) <> "" And Trim(.Cells(tmpRow, 3)) <> ""
            StrMBPN = ""
            StrPN = Trim(.Cells(tmpRow, 1))
            StrInterval = Trim(.Cells(tmpRow, 2))
            strFactory = Trim(.Cells(tmpRow, 3))
            
            strSQL = "EXEC XL_UploadToPNInterval '" & StrMBPN & "','" & StrPN & "','" & StrInterval & "','" & strUID & "','" & strFactory & "'"
            Set rsTmp = Conn.Execute(strSQL)
            
            If rsTmp("result") = 0 Then
                blerr = True
                .Cells(tmpRow, 4).Interior.ColorIndex = 3
                .Cells(tmpRow, 4).Interior.Pattern = xlSolid
                .Cells(tmpRow, 4) = "This PN upload fail !"
            End If
            Txt_RowCount.text = tmpRow - 1 'add by Kevin show the action row 2009.04.07 (0048)
            tmpRow = tmpRow + 1
        Wend
    End With
    
    If blerr Then
        MsgBox ("Upload fail ! Notice the red fail part in the EXCEL file !"), vbCritical
        Let xlApp.Visible = True
        Exit Sub
    End If

    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    
    MsgBox ("*** Load  Finish ! ***")
    Exit Sub
    
errhandle:
    MsgBox Err.Description
End Sub

'''''''''''''''''''''''''''''''''''''''''added by Jing 2008.03.23 (0022)'''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub load_XL_ECWOPlan(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rsTmp As ADODB.Recordset
Dim strSQL As String, strDate As String, strLine As String, strModel As String, intQty As Integer, strOPID As String, strTransDT As String
Dim tmpRow As Integer
Dim tmpNum1 As String, tmpNum2 As String, tmpNum3 As String
Dim intNum As Integer
Dim blerr As Boolean
blerr = False

On Error GoTo errhandle:

    strSQL = "select getdate()"
    Set rsTmp = Conn.Execute(strSQL)
    strTransDT = Format(rsTmp.Fields(0), "YYYYMMDDHHMMSS")
    
    strDate = Left(Right(Trim(txtFilePath), 12), 8)
    
    If CheckNum(strDate) = False Then
        MsgBox ("Please check the name of xls !" & Chr(13) & "XLS name must include Date !")
        Exit Sub
    End If
    
    strOPID = Trim(g_userName)
    
    tmpRow = 3
    
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = True
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    
    With xlsBook.Worksheets(Trim(Shift_Item))
        While Trim(.Cells(tmpRow, 1)) <> "" Or (Trim(.Cells(tmpRow, 2)) <> "" And Trim(.Cells(tmpRow, 5)) <> "" And strLine <> "") Or (intNum <> 50 And strLine <> "")
            If (Trim(.Cells(tmpRow, 1)) <> "" And Trim(.Cells(tmpRow, 1)) <> strLine) Or strLine = "" Then
                If CheckNum(Trim(.Cells(tmpRow, 1))) = False Then
                    strLine = Left(Trim(.Cells(tmpRow, 1)), 1)
                End If
            End If
            
            If Trim(.Cells(tmpRow, 2)) <> "" And Trim(.Cells(tmpRow, 5)) <> "" Then
                strModel = Trim(.Cells(tmpRow, 2)) + "-" + Left(Trim(.Cells(tmpRow, 5)), 3)
                
                strSQL = "select * from modelname where modelname='" & strModel & "'"
                Set rsTmp = Conn.Execute(strSQL)
                
                If rsTmp.EOF = True Then
                    blerr = True
                    .Cells(tmpRow, 16).Interior.ColorIndex = 3
                    .Cells(tmpRow, 16).Interior.Pattern = xlSolid
                    .Cells(tmpRow, 16) = "Part_NO or Remark define error !"
                End If
                
                intQty = 0
                tmpNum1 = ""
                tmpNum2 = ""
                tmpNum3 = ""
                
                If Trim(.Cells(tmpRow, 11)) <> "" And CheckNum(Trim(.Cells(tmpRow, 11))) = True Then
                    tmpNum1 = Trim(.Cells(tmpRow, 11))
                    intQty = CInt(tmpNum1)
                End If
                
                If Trim(.Cells(tmpRow, 12)) <> "" And CheckNum(Trim(.Cells(tmpRow, 12))) = True Then
                    tmpNum2 = Trim(.Cells(tmpRow, 12))
                    intQty = intQty + CInt(tmpNum2)
                End If
                
                If Trim(.Cells(tmpRow, 13)) <> "" And CheckNum(Trim(.Cells(tmpRow, 13))) = True Then
                    tmpNum3 = Trim(.Cells(tmpRow, 13))
                    intQty = intQty + CInt(tmpNum3)
                End If
                
                If intQty <> 0 Then
                    strSQL = "Delete from XL_EC_WOPlan where date='" & strDate & "' and line='" & strLine & "' and model='" & strModel & "'"
                    Conn.Execute (strSQL)
                    
                    strSQL = "Insert into XL_EC_WOPlan(Date,Line,Model,Qty,OPID,TransDateTime) values('" & strDate & "','" & strLine & "','" & strModel & "','" & intQty & "','" & strOPID & "','" & strTransDT & "')"
                    Conn.Execute (strSQL)
                End If
                
                If CheckNum(Trim(.Cells(tmpRow, 11))) = False Or CheckNum(Trim(.Cells(tmpRow, 12))) = False Then
                    blerr = True
                    .Cells(tmpRow, 17).Interior.ColorIndex = 3
                    .Cells(tmpRow, 17).Interior.Pattern = xlSolid
                    .Cells(tmpRow, 17) = "This Row upload fail !"
                End If
                
                intNum = 0
            Else
                intNum = intNum + 1             '''用于累计空行数(如果连续有50行空行，将不对后面数据进行读取)'''
            End If
            
            tmpRow = tmpRow + 1
        Wend
    End With
    
    If blerr Then
        MsgBox ("Notice the red fail part in the EXCEL file !"), vbCritical
        Let xlApp.Visible = True
        Exit Sub
    End If

    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox ("*** Load  Finish ! ***")
    Exit Sub
    
errhandle:
    MsgBox Err.Description
End Sub

'''''''''''''''''''''''''''''''''''''''''add by Archer 2008.04.01 (0024)'''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub load_XL_DoubleTables(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rsTmp As ADODB.Recordset
Dim strSQL As String, StrJobGroup As String, strMachine As String, strLine As String, strUID As String
Dim tmpRow As Integer
Dim blerr As Boolean
blerr = False

On Error GoTo errhandle:
    
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    
    strSQL = "Truncate table DoubleTables"
    Set rsTmp = Conn.Execute(strSQL)
    
    tmpRow = 2
    strUID = Trim(g_userName)
    
    With xlsBook.Worksheets(Trim(Shift_Item))
        While Trim(.Cells(tmpRow, 1)) <> "" And Trim(.Cells(tmpRow, 2)) <> "" And Trim(.Cells(tmpRow, 3)) <> ""
            StrJobGroup = Trim(.Cells(tmpRow, 1))
            strMachine = Trim(.Cells(tmpRow, 2))
            strLine = Trim(.Cells(tmpRow, 3))                '''''''''''''''''''''''''''''''1049
            strSQL = "EXEC XL_UploadToDoubleTables '" & StrJobGroup & "','" & strMachine & "','" & strUID & "','" & strLine & "'"
            Set rsTmp = Conn.Execute(strSQL)
            
            If rsTmp("result") = 0 Then
                blerr = True
                .Cells(tmpRow, 5).Interior.ColorIndex = 3
                .Cells(tmpRow, 5).Interior.Pattern = xlSolid
                .Cells(tmpRow, 5) = "This setting upload fail !"
            End If
            tmpRow = tmpRow + 1
        Wend
    End With
    
    If blerr Then
        MsgBox ("Upload fail ! Notice the red fail part in the EXCEL file !"), vbCritical
        Let xlApp.Visible = True
        Exit Sub
    End If

    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox ("*** Load  Finish ! ***")
    Exit Sub
    
errhandle:
    MsgBox Err.Description
End Sub
''''''''''''''''''''''''''''''''''''(0059)''''''''''''''''''''''''''''''
Private Sub Load_QSMS_CheckCompPN(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rsTmp As ADODB.Recordset
Dim Rs As ADODB.Recordset
Dim strSQL As String, strJobPN As String, strCompPN As String, strUID As String
Dim tmpRow As Integer
Dim blerr As Boolean
blerr = False

On Error GoTo errhandle:
    
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    
    strSQL = "Truncate table QSMS_CompPNcheck_Temp" ''(1110)
    Set rsTmp = Conn.Execute(strSQL)
    
    tmpRow = 2
    strUID = Trim(g_userName)
    
    With xlsBook.Worksheets(Trim(Shift_Item))
        While Trim(.Cells(tmpRow, 1)) <> "" And Trim(.Cells(tmpRow, 2)) <> ""
            strJobPN = Trim(.Cells(tmpRow, 1))
            strCompPN = Trim(.Cells(tmpRow, 2))
            
            strSQL = "EXEC QSMS_CheckCompPN '" & strJobPN & "','" & strCompPN & "','" & strUID & "','W'"
            Set rsTmp = Conn.Execute(strSQL)
            
            If Trim(rsTmp("result")) <> "0" Then
                MsgBox ("Err: " + rsTmp("desc1"))
                GoTo NormalHandle
            End If
            tmpRow = tmpRow + 1
        Wend
    End With
    
strSQL = "EXEC QSMS_CheckCompPN "
Set Rs = Conn.Execute(strSQL)
If Not Rs.EOF Then
       Call CopyToExcel(Rs)
    Else
       MsgBox ("No Data"), vbCritical
End If

NormalHandle:
    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox ("*** Load  Finish ! ***")
    Exit Sub
    
errhandle:
    MsgBox Err.Description
End Sub
'Private Sub load_WorkHS_Equipment(Shift_Item As String)
'Dim xlApp As Excel.Application
'Dim xlsBook As Excel.Workbook
'Dim xlWs As Excel.Worksheets
'Dim rsTmp As ADODB.Recordset
'Dim strSQL As String, strDeviceName As String, strSupplier As String, strUID As String
'Dim tmpRow As Integer
'Dim blErr As Boolean
'blErr = False
'
'On Error GoTo ErrHandle:
'
'    Set xlApp = CreateObject("Excel.Application")
'    Let xlApp.Visible = False
'    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
'    xlApp.DisplayAlerts = False
'
'    tmpRow = 2
'    strUID = Trim(g_userName)
'
'    With xlsBook.Worksheets(Trim(Shift_Item))
'        While Trim(.Cells(tmpRow, 1)) <> "" And Trim(.Cells(tmpRow, 2)) <> ""
'            strDeviceName = Trim(.Cells(tmpRow, 1))
'            strSupplier = Trim(.Cells(tmpRow, 2))
'            strSQL = "EXEC WorkHS_Upload_Equipment '" & strDeviceName & "','" & strSupplier & "','" & strUID & "'"
'            Set rsTmp = Conn.Execute(strSQL)
'
'            If rsTmp("result") = 0 Then
'                blErr = True
'                .Cells(tmpRow, 5).Interior.ColorIndex = 3
'                .Cells(tmpRow, 5).Interior.Pattern = xlSolid
'                .Cells(tmpRow, 5) = "This setting upload fail !"
'            End If
'            tmpRow = tmpRow + 1
'        Wend
'    End With
'
'    If blErr Then
'        MsgBox ("Upload fail ! Notice the red fail part in the EXCEL file !"), vbCritical
'        Let xlApp.Visible = True
'        Exit Sub
'    End If
'
'    xlsBook.Close
'    xlApp.Quit
'    Set xlApp = Nothing
'    Set xlsBook = Nothing
'    MsgBox ("*** Load  Finish ! ***")
'    Exit Sub
'
'ErrHandle:
'    MsgBox Err.Description
'End Sub
'
''''''''''''''''''''''''''''''''''''Add by Archer (0027)''''''''''''''''''''''''''''''
'Private Sub load_WorkHS_LineConfig(Shift_Item As String)
'Dim xlApp As Excel.Application
'Dim xlsBook As Excel.Workbook
'Dim xlWs As Excel.Worksheets
'Dim rsTmp As ADODB.Recordset
'Dim strSQL As String, strLine As String, strSide As String, strSeqID As Integer, strDeviceSupplier As String, strUID As String
'Dim tmpRow As Integer, j As Integer
'Dim blErr As Boolean
'blErr = False
'
'On Error GoTo ErrHandle:
'
'    Set xlApp = CreateObject("Excel.Application")
'    Let xlApp.Visible = False
'    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
'    xlApp.DisplayAlerts = False
'
'    tmpRow = 1
'    strUID = Trim(g_userName)
'
'    With xlsBook.Worksheets(Trim(Shift_Item))
'        While Trim(.Cells(tmpRow, 1)) <> "" And Trim(.Cells(tmpRow, 2)) <> ""
'            j = 3
'            While Trim(.Cells(tmpRow, j)) <> "" And Trim(.Cells(tmpRow, j + 1)) <> ""
'                strLine = Trim(.Cells(tmpRow, 1))
'                strSide = Trim(.Cells(tmpRow, 2))
'                strSeqID = Trim(.Cells(tmpRow, j))
'                strDeviceSupplier = Trim(.Cells(tmpRow, j + 1))
'                strSQL = "EXEC WorkHS_Upload_LineConfig '" & strLine & "','" & strSide & "','" & strSeqID & "','" & strDeviceSupplier & "','" & strUID & "'"
'                Set rsTmp = Conn.Execute(strSQL)
'
'                If rsTmp("result") = 0 Then
'                    blErr = True
'                    .Cells(tmpRow, 5).Interior.ColorIndex = 3
'                    .Cells(tmpRow, 5).Interior.Pattern = xlSolid
'                    .Cells(tmpRow, 5) = "This setting upload fail !"
'                End If
'                j = j + 2
'            Wend
'            tmpRow = tmpRow + 1
'        Wend
'    End With
'
'    If blErr Then
'        MsgBox ("Upload fail ! Notice the red fail part in the EXCEL file !"), vbCritical
'        Let xlApp.Visible = True
'        Exit Sub
'    End If
'
'    xlsBook.Close
'    xlApp.Quit
'    Set xlApp = Nothing
'    Set xlsBook = Nothing
'    MsgBox ("*** Load  Finish ! ***")
'    Exit Sub
'
'ErrHandle:
'    MsgBox Err.Description
'End Sub

Public Function CheckNum(str As String) As Boolean
Dim i As Integer

For i = 1 To Len(Trim(str))
    If IsNumeric(Mid(str, i, 1)) = False Then
        CheckNum = False
        Exit Function
    End If
Next i
CheckNum = True
End Function

Private Sub Upload_JobGroup(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim Rs As New ADODB.Recordset
Dim i As Integer
Dim strSQL As String
Dim L As Integer

On Error GoTo errhandle:

Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.Visible = False
xlApp.UserControl = True
xlApp.DisplayAlerts = False


'''''''''''''————————初始化数据表————————''''''''
If Trim(xlApp.Worksheets(1).Cells(1, 1)) <> "Factory" Or Trim(xlApp.Worksheets(1).Cells(1, 2)) <> "Line" Or Trim(xlApp.Worksheets(1).Cells(1, 3)) <> "JobGroup" Or Trim(xlApp.Worksheets(1).Cells(1, 4)) <> "JobPN" Or Trim(xlApp.Worksheets(1).Cells(1, 5)) <> "Version" Then
    MsgBox "Excel Field Format Error.Format:Factory-Line-JobGroup-JobPN-Version", vbOKOnly + vbInformation, "Excel Field Format Error"
    
    xlApp.Quit
    Set xlApp = Nothing
    Exit Sub
End If
'''''''''''''Step 1:清空临时表 JobGroup_Temp
strSQL = "truncate table JobGroup_Temp"
If Rs.State = 1 Then Rs.Close
Set Rs = Conn.Execute(strSQL)

'''''''''''''Step 2:把Excel中数据上传到临时表：JobGroup_Temp，如Excel中有重复，重复数据不上传
i = 2
L = 0
With xlBook.Worksheets(Trim(Shift_Item))
    Do While Not (.Cells(i, 1) = "")
       
    '''''''''''''————————导入数据，用插入方法导入数据————————''''''''
        strSQL = "SELECT * FROM JobGroup_Temp WHERE Factory='" & Trim(.Cells(i, 1)) & "' and Line='" & Trim(.Cells(i, 2)) & "' and JobGroup='" & Trim(.Cells(i, 3)) & "' and JobPN='" & Trim(.Cells(i, 4)) & "' and Version='" & Trim(.Cells(i, 5)) & "'"
        If Rs.State = 1 Then Rs.Close
        Set Rs = Conn.Execute(strSQL)
        If Rs.EOF Then
            strSQL = "Insert Into JobGroup_Temp Values ('" & Trim(.Cells(i, 1)) & "','" & Trim(.Cells(i, 2)) & "','" & Trim(.Cells(i, 3)) & "','" & Trim(.Cells(i, 4)) & "','" & Trim(.Cells(i, 5)) & "','" & Trim(g_userName) & "',DBO.FormatDate(Getdate(),'YYYYMMDDHHNNSS'))"
            '''DBO.FormatDate(Getdate(),'YYYYMMDDHHNNSS')
            If Rs.State = 1 Then Rs.Close
            Set Rs = Conn.Execute(strSQL)
            L = L + 1
        End If
        i = i + 1
    Loop
End With
Set Rs = Nothing

xlApp.Quit
Set xlApp = Nothing

'''''''''''''Step 3:处理临时表 JobGroup_Temp，并把数据从临时表转移到主表JobGroup
strSQL = "EXEC QSMS_JobGroup"
If Rs.State = 1 Then Rs.Close
Set Rs = Conn.Execute(strSQL)


MsgBox "Excel Data:" & i - 2 & vbCrLf & "Upload:" & L & vbCrLf & "duplication:" & i - L - 2
Exit Sub

errhandle:
    MsgBox Err.Description
End Sub

''(0040)
Private Sub Load_LineFUJIServer(Shift_Item As String)
Dim xlsApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim rCount, Row_Count As Long
Dim Line, FUJIServer As String
Dim Total_Qty, Update_Qty, Insert_Qty, Delete_Qty As Long
Dim aryServerIP As Variant
Dim transdatetime As String
Dim strSQL As String
Dim Rs As ADODB.Recordset
Dim strDelete As String

If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
   Exit Sub
End If
Set xlsApp = CreateObject("Excel.Application")
Let xlsApp.Visible = False
Set xlsBook = xlsApp.Workbooks.Open(txtFilePath)
xlsApp.DisplayAlerts = False

rCount = 2
Total_Qty = 0
Insert_Qty = 0
Update_Qty = 0
Delete_Qty = 0

strSQL = "select getdate() as TransDateTime"
Set Rs = Conn.Execute(strSQL)
transdatetime = Format(Rs(0), "yyyymmddhhnnss")
With xlsBook.Worksheets(Trim(Shift_Item))

    While Trim(.Cells(rCount, 1)) <> ""
        Line = Trim(.Cells(rCount, 1) & vbNullString)
        FUJIServer = Trim(.Cells(rCount, 2) & vbNullString)
        strDelete = UCase(Trim(.Cells(rCount, 3)))
            
        aryServerIP = Split(FUJIServer, ".")
        If UBound(aryServerIP) <> 3 Then
            MsgBox "Excel file format error,please check FUJI Server IP: ROW:" & rCount + 1
            Exit Sub
        End If
            
        If strDelete = "Y" Then
           strSQL = "Delete from Line_FUJI where Line='" & Line & "'"
           Conn.Execute (strSQL)
           Delete_Qty = Delete_Qty + 1
        Else
            strSQL = "select * from Line_FUJI where Line='" & Line & "'"
            Set Rs = Conn.Execute(strSQL)
            If Rs.EOF Then
                strSQL = "Insert into Line_FUJI(Line,FUJI_Server,UID,Transdatetime) " & _
                " values('" & Trim(Line) & "','" & FUJIServer & "','" & g_userName & "','" & transdatetime & "')"
                Conn.Execute strSQL
                Insert_Qty = Insert_Qty + 1
            Else
                strSQL = "Update Line_FUJI set FUJI_Server='" & Trim(FUJIServer) & "',UID='" & g_userName & "',Transdatetime='" & transdatetime & "' where Line='" & Line & "'"
                Conn.Execute strSQL
                Update_Qty = Update_Qty + 1
            End If

            DoEvents
            DoEvents
            DoEvents
        End If
            
        rCount = rCount + 1
        Total_Qty = Total_Qty + 1
        Txt_RowCount = Total_Qty
    Wend
End With
strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_LineFUJIServer','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSQL)

xlsBook.Close
xlsApp.Quit
Set xlsApp = Nothing
Set xlsBook = Nothing
MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
             "Total Counter : " & Total_Qty & vbCrLf & _
             "Insert succeed : " & Insert_Qty & vbCrLf & _
             "Update succeed : " & Update_Qty & vbCrLf & _
             "Delete succeed :" & Delete_Qty & vbCrLf
End Sub
Private Sub XL_MaxDIDMaintainQty(Shift_Item As String)  '(0041)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rCount, Row_Count As Long
Dim COMPPN, DeletedFlag As String
Dim Qty As Integer
Dim Total_Qty, Deleted_Qty, Insert_Qty, Update_Qty As Long
Dim str As String
Dim Rs As ADODB.Recordset
Dim transdatetime As String
If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
   Exit Sub
End If
Set xlApp = CreateObject("Excel.Application")
Let xlApp.Visible = False
Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.DisplayAlerts = False
rCount = 2
Total_Qty = 0

Insert_Qty = 0
Deleted_Qty = 0
str = "select getdate()"
Set Rs = Conn.Execute(str)
transdatetime = Rs.Fields(0) 'Format(Rs.Fields(0), "YYYYMMDDHHMMSS")

With xlsBook.Worksheets(Trim(Shift_Item))
     
     While Trim(.Cells(rCount, 2)) <> ""
           COMPPN = Trim(.Cells(rCount, 1) & vbNullString)
           Qty = Replace(Trim(.Cells(rCount, 2) & vbNullString), "'", " ")
           DeletedFlag = Trim(.Cells(rCount, 3) & vbNullString)
           
           If UCase(DeletedFlag) = "Y" Then
                 str = "delete from XL_MaxDIDMaintainQty where CompPN='" & Trim(COMPPN) & "' "
                 Conn.Execute str
                 Deleted_Qty = Deleted_Qty + 1
           Else
                 str = "select * from XL_MaxDIDMaintainQty where CompPN='" & Trim(COMPPN) & "'"
                 Set Rs = Conn.Execute(str)
                 If Rs.EOF Then
                     str = "Insert into XL_MaxDIDMaintainQty(CompPN,Qty,UID,TransDateTime) " & _
                       " values('" & COMPPN & "','" & Trim(Qty) & "','" & Trim(g_userName) & "','" & Format(transdatetime, "yyyymmddhhnnss") & "')"
                     Conn.Execute str
                    Insert_Qty = Insert_Qty + 1
                 Else
                     str = "update XL_MaxDIDMaintainQty set Qty='" & Trim(Qty) & "'where  CompPN='" & Trim(COMPPN) & "'  and  Qty<>'" & Trim(Qty) & "'"
                     Conn.Execute str
                     Update_Qty = Update_Qty + 1
                 End If
           End If
           DoEvents
           
          rCount = rCount + 1
          Total_Qty = Total_Qty + 1
          Txt_RowCount = Total_Qty
    Wend
End With
str = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','XL_MaxMaintainQty','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (str)
xlsBook.Close
xlApp.Quit
Set xlApp = Nothing
Set xlsBook = Nothing
MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
             "Total Counter : " & Total_Qty & vbCrLf & _
             "Insert succeed : " & Insert_Qty & vbCrLf & _
             "Deleted succeed : " & Deleted_Qty & vbCrLf & _
             "Update succeed :" & Update_Qty & vbCrLf
            
End Sub

Private Sub lblFileFormat_Click()
Dim FileName As String
    If lblFileFormat <> "" Then
        FileName = lblFileFormat
        Call ShellExecute(0, "Open", FileName, vbNullString, vbNullString, SW_SHOWNORMAL)
    End If
End Sub

Private Sub Upload_NOCheckReplacePNSplicing(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim Rs As New ADODB.Recordset
Dim i As Integer
Dim strSQL As String
Dim L As Integer

On Error GoTo errhandle:

Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.Visible = False
xlApp.UserControl = True
xlApp.DisplayAlerts = False

strSQL = "truncate table QSMS_NOCheckReplacePNSplicing"
If Rs.State = 1 Then Rs.Close
Set Rs = Conn.Execute(strSQL)

'''''''''''''Step 2:把Excel中数据上传到临时表：JobGroup_Temp，如Excel中有重复，重复数据不上传
i = 2
With xlBook.Worksheets(Trim(Shift_Item))
    Do While Not (.Cells(i, 1) = "")
        strSQL = "Insert Into QSMS_NOCheckReplacePNSplicing(PrefixPN,UID,TransDateTime,FuncType) " & _
                 "Values ('" & Trim(.Cells(i, 1)) & "','" & Trim(g_userName) & "',DBO.FormatDate(Getdate(),'YYYYMMDDHHNNSS'),'" & Trim(.Cells(i, 2)) & "')" ''(1055)
        Conn.Execute (strSQL)
        i = i + 1
    Loop
End With
Set Rs = Nothing

xlApp.Quit
Set xlApp = Nothing

MsgBox "Excel Data:" & i - 2 & vbCrLf & "Upload:" & L
Exit Sub

errhandle:
    MsgBox Err.Description
End Sub


''20090803  Denver     Add upload PNGroup
Private Function Upload_PNGroup(sFile As String, sSheetName As String) As Boolean
On Error GoTo Err_Handler
    Dim xlApp As Excel.Application
    Dim xlsBook As Excel.Workbook
    Dim xlWs As Excel.Worksheets
    Dim rCount As Long
    Dim PN As String, PNGroup As String, OPID As String, UpdFlag As String
    Dim Total_Qty, Update_Qty, Insert_Qty As Long, Delete_Qty As Long
    Dim transdatetime As String
    Dim i As Integer
    Dim sSql As String
    Dim Rst As New ADODB.Recordset
    Dim intCol As Integer
    Dim sColHead As String
    Dim varColHead As Variant
    
    
    Upload_PNGroup = False
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(sFile)
    xlApp.DisplayAlerts = False
    
    sSql = "select getdate()"
    Set Rst = Conn.Execute(sSql)
    transdatetime = Format(Rst.Fields(0), "YYYYMMDDHHMMSS")

    rCount = 2
    Total_Qty = 0
    Insert_Qty = 0
    Update_Qty = 0
    Delete_Qty = 0
    Txt_RowCount = 0
    
    '==================================================================================================================================================
    'can not change
    sColHead = "PN,PNGroup,UpdFlag"
    '==================================================================================================================================================
    varColHead = Split(sColHead, ",")
    With xlsBook.Worksheets(Trim(sSheetName))
        For intCol = 0 To UBound(varColHead)
            If UCase(Trim(.Cells(1, intCol + 1) & vbNullString)) <> UCase(varColHead(intCol)) Then
'                MsgBox "File Format Error: " & sColHead & " !!", vbExclamation, "Prompt"
                MsgBox "The format of the uploaded EXCEL file is wrong, the correct format should be " & sColHead & "", vbExclamation, "Prompt"
                GoTo Normal_Exit
            End If
        Next intCol
        
        While Trim(.Cells(rCount, 1)) <> ""
            
            PN = Trim(.Cells(rCount, 1) & vbNullString)
            PNGroup = Trim(.Cells(rCount, 2) & vbNullString)
            UpdFlag = Trim(.Cells(rCount, 3) & vbNullString)
            
            If PNGroup = "" Or PN = "" Then
               MsgBox "PNGroup Or PN can not be blank!!"
               GoTo Normal_Exit
            End If
            
            '1 Add, 2 Update,3 Delete
            If UpdFlag <> "1" And UpdFlag <> "2" And UpdFlag <> "3" Then
                MsgBox "UpdFlag must be 1 or 2 or 3,it means that Add,Update,Delete!!"
                GoTo Normal_Exit
            End If
            sSql = "exec Upload_PNGroup " & sq(PN) & "," & sq(PNGroup) & "," & sq(g_userName) & "," & sq(transdatetime) & "," & sq(UpdFlag)
                
            Set Rst = Conn.Execute(sSql)
            If Rst("Result") <> 0 Then
                MsgBox Rst("Description")
                GoTo Normal_Exit
            Else
                Select Case UpdFlag
                Case "1"
                    Insert_Qty = Insert_Qty + 1
                Case "2"
                    Update_Qty = Update_Qty + 1
                Case "3"
                    Delete_Qty = Delete_Qty + 1
                End Select
             
            End If
            
            rCount = rCount + 1
            Total_Qty = Total_Qty + 1
            Txt_RowCount = Total_Qty
        Wend
    
    End With
    
    Upload_PNGroup = True
    MsgBox "*** Load  finish ! ***PN Group   " & vbCrLf & _
             "Total Counter : " & Total_Qty & vbCrLf & _
             "Insert succeed : " & Insert_Qty & vbCrLf & _
             "Update succeed : " & Update_Qty & vbCrLf & _
             "Delete succeed : " & Delete_Qty & vbCrLf

Normal_Exit:
    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    Exit Function
    
Err_Handler:
    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox Err.Number & "," & Err.Description
End Function

'''''''''''''''''''''''''''''''''''''''''add by Richie 2009.10.17 (0062)'''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UploadIC_CompPN(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rsTmp As ADODB.Recordset
Dim strUID As String
Dim sql As String
Dim tmpRow As Integer

Dim blerr As Boolean
Dim Model As String
Dim PN As String
Dim COMPPN As String
Dim flag As String
Dim location As String
'20110920 Maggie 增加上传栏位
Dim Customer As String, Functions As String, ICVendor As String, MfrPN As String, Rev As String, FWPN As String, FWFile As String
Dim FWFileSize As String, FWFilePath As String, CheckSum As String, ICMark As String, CutInDate As String, Programmer As String, Store As String, Remark As String

blerr = False

On Error GoTo errhandle:
    
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    tmpRow = 2
    strUID = Trim(g_userName)
    
    With xlsBook.Worksheets(Trim(Shift_Item))
        Model = Trim(.Cells(tmpRow, 1))
        PN = Trim(.Cells(tmpRow, 2))
        Customer = Trim(.Cells(tmpRow, 3))
        Functions = Trim(.Cells(tmpRow, 4))
        COMPPN = Trim(.Cells(tmpRow, 5))
        ICVendor = Trim(.Cells(tmpRow, 6))
        MfrPN = Trim(.Cells(tmpRow, 7))
        location = Trim(.Cells(tmpRow, 8))
        Rev = Trim(.Cells(tmpRow, 9))
        FWPN = Trim(.Cells(tmpRow, 10))
        FWFile = Trim(.Cells(tmpRow, 11))
        FWFilePath = Trim(.Cells(tmpRow, 12))
        FWFileSize = Trim(.Cells(tmpRow, 13))
        CheckSum = Trim(.Cells(tmpRow, 14))
        ICMark = Trim(.Cells(tmpRow, 15))
        CutInDate = Trim(.Cells(tmpRow, 16))
        Programmer = Trim(.Cells(tmpRow, 17))
        Store = Trim(.Cells(tmpRow, 18))
        Remark = Trim(.Cells(tmpRow, 19))
        flag = Trim(.Cells(tmpRow, 20))
        If Model = "" Or PN = "" Or COMPPN = "" Or flag = "" Then   ''(1122)
               blerr = True
               .Cells(tmpRow, 20).Interior.ColorIndex = 3
               .Cells(tmpRow, 20).Interior.Pattern = xlSolid
               .Cells(tmpRow, 20) = "Model or PN or CompPN or Flag 不能为空!!"
        End If
        While Model <> "" And PN <> "" And COMPPN <> "" And flag <> ""
            'sql = "EXEC IC_UploadCompPN '" & Model & "','" & PN & "','" & compPN & "','" & Location & "','" & strUID & "','" & flag & "'"
            sql = "EXEC IC_UploadCompPN '" & Model & "','" & PN & "','" & COMPPN & "','" & location & "','" & strUID & "','" & flag & "'," & sq(Customer) & ",N'" & Functions & "'," & _
                sq(ICVendor) & ", " & sq(MfrPN) & "," & sq(Rev) & "," & sq(FWPN) & "," & sq(FWFile) & "," & sq(FWFileSize) & "," & sq(CheckSum) & ",N'" & ICMark & "'," & sq(CutInDate) & "," & _
                sq(Programmer) & "," & sq(Store) & ",N'" & Remark & "'," & sq(FWFilePath) & ""
            Set rsTmp = Conn.Execute(sql)
            If rsTmp("result") = "0" Then
                blerr = True
                '.Cells(tmpRow, 5).Interior.ColorIndex = 3
                '.Cells(tmpRow, 5).Interior.Pattern = xlSolid
                '.Cells(tmpRow, 5) = rsTmp("description")
                .Cells(tmpRow, 21).Interior.ColorIndex = 3
                .Cells(tmpRow, 21).Interior.Pattern = xlSolid
                .Cells(tmpRow, 21) = rsTmp("description")
            End If
            tmpRow = tmpRow + 1
            'Model = Trim(.Cells(tmpRow, 1))
            'PN = Trim(.Cells(tmpRow, 2))
            'compPN = Trim(.Cells(tmpRow, 3))
            'Location = Trim(.Cells(tmpRow, 4))
            'flag = Trim(.Cells(tmpRow, 5))
            Model = Trim(.Cells(tmpRow, 1))
            PN = Trim(.Cells(tmpRow, 2))
            Customer = Trim(.Cells(tmpRow, 3))
            Functions = Trim(.Cells(tmpRow, 4))
            COMPPN = Trim(.Cells(tmpRow, 5))
            ICVendor = Trim(.Cells(tmpRow, 6))
            MfrPN = Trim(.Cells(tmpRow, 7))
            location = Trim(.Cells(tmpRow, 8))
            Rev = Trim(.Cells(tmpRow, 9))
            FWPN = Trim(.Cells(tmpRow, 10))
            FWFile = Trim(.Cells(tmpRow, 11))
            FWFilePath = Trim(.Cells(tmpRow, 12))
            FWFileSize = Trim(.Cells(tmpRow, 13))
            CheckSum = Trim(.Cells(tmpRow, 14))
            ICMark = Trim(.Cells(tmpRow, 15))
            CutInDate = Trim(.Cells(tmpRow, 16))
            Programmer = Trim(.Cells(tmpRow, 17))
            Store = Trim(.Cells(tmpRow, 18))
            Remark = Trim(.Cells(tmpRow, 19))
            flag = Trim(.Cells(tmpRow, 20))
        Wend
    End With
    
    If blerr Then
        MsgBox ("Upload fail ! Notice the red fail part in the EXCEL file !"), vbCritical
        Let xlApp.Visible = True
        Exit Sub
    End If

    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox ("*** Load  Finish ! ***")
    Exit Sub
    
errhandle:
    MsgBox Err.Description
End Sub
Private Sub UploadIC_ShearPin(Shift_Item As String)   '''(1125)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rsTmp As ADODB.Recordset
Dim strUID As String
Dim sql As String
Dim tmpRow As Integer

Dim blerr As Boolean
Dim Model As String
Dim PN As String
Dim COMPPN As String
Dim flag As String
Dim Thickness As String, ReservedLength As String, Remark As String
blerr = False

On Error GoTo errhandle:
    
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    tmpRow = 3
    strUID = Trim(g_userName)
    
    With xlsBook.Worksheets(Trim(Shift_Item))
        Model = Trim(.Cells(tmpRow, 1))
        PN = Trim(.Cells(tmpRow, 2))
        COMPPN = Trim(.Cells(tmpRow, 3))           ''''(1134)
        Thickness = Trim(.Cells(tmpRow, 4))
        ReservedLength = Trim(.Cells(tmpRow, 5))
        Remark = Trim(.Cells(tmpRow, 6))
        flag = Trim(.Cells(tmpRow, 7))
        If Model = "" Or PN = "" Or COMPPN = "" Or flag = "" Then
               blerr = True
               .Cells(tmpRow, 7).Interior.ColorIndex = 3
               .Cells(tmpRow, 7).Interior.Pattern = xlSolid
               .Cells(tmpRow, 7) = "Model or PN or CompPN or Flag 不能为空!!"
        End If
        While Model <> "" And PN <> "" And flag <> ""
            sql = "EXEC IC_UploadShearPin '" & Model & "','" & PN & "'," & sq(Thickness) & "," & sq(ReservedLength) & "," & sq(strUID) & "," & sq(flag) & ",N'" & Remark & "'," & sq(COMPPN) & ""
            Set rsTmp = Conn.Execute(sql)
            If rsTmp("result") = "0" Then
                blerr = True
                .Cells(tmpRow, 6).Interior.ColorIndex = 3
                .Cells(tmpRow, 6).Interior.Pattern = xlSolid
                .Cells(tmpRow, 6) = rsTmp("description")
            End If
            tmpRow = tmpRow + 1
            Model = Trim(.Cells(tmpRow, 1))
            PN = Trim(.Cells(tmpRow, 2))
            COMPPN = Trim(.Cells(tmpRow, 3))   ''''(1134)
            Thickness = Trim(.Cells(tmpRow, 4))
            ReservedLength = Trim(.Cells(tmpRow, 5))
            Remark = Trim(.Cells(tmpRow, 6))
            flag = Trim(.Cells(tmpRow, 7))
        Wend
    End With
    
    If blerr Then
        MsgBox ("Upload fail ! Notice the red fail part in the EXCEL file !"), vbCritical
        Let xlApp.Visible = True
        Exit Sub
    End If

    xlsBook.Close
    xlApp.Quit
    Set xlApp = Nothing
    Set xlsBook = Nothing
    MsgBox ("*** Load  Finish ! ***")
    Exit Sub
    
errhandle:
    MsgBox Err.Description
End Sub

Private Sub upload_traycompPN(Shift_Item As String)  '2010.08.17 add by kaitlyn (1001)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rowcnt As Long
Dim strSQL As String
Dim COMPPN, UID As String, BaseQty As Integer, delflag As String
Dim Rs As ADODB.Recordset
Dim blerr As Boolean
blerr = False

If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
    Exit Sub
End If
On Error GoTo errhandle:
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    rowcnt = 2
    UID = Trim(g_userName)
    
    With xlsBook.Worksheets(Trim(Shift_Item))
        While Trim(.Cells(rowcnt, 1)) <> ""
            COMPPN = Trim(.Cells(rowcnt, 1))
            BaseQty = .Cells(rowcnt, 2)
            delflag = .Cells(rowcnt, 3)
            strSQL = "EXEC upload_traycompPN '" & COMPPN & "','" & BaseQty & "','" & UID & "','" & delflag & "' "
            Set Rs = Conn.Execute(strSQL)
            If Rs("result") = 0 Then
                blerr = True
                .Cells(rowcnt, 4).Interior.ColorIndex = 3
                .Cells(rowcnt, 4).Interior.Pattern = xlSolid
                .Cells(rowcnt, 4) = Rs("errlog")
            End If
            rowcnt = rowcnt + 1
        Wend
    End With
    If blerr = True Then
        MsgBox ("Upload fail ! Notice the red fail part in the EXCEL file !"), vbCritical
        Let xlApp.Visible = True
        Exit Sub
    End If
    xlsBook.Close
    xlApp.Quit
    Set xlsBook = Nothing
    Set xlApp = Nothing
    MsgBox ("***upload finished!***")
    Exit Sub
errhandle:
    MsgBox Err.Description
    
End Sub

Private Sub upload_ComponentData(Shift_Item As String)  '2010.12.08 add by kaitlyn (1024)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rowcnt As Long
Dim strSQL As String
Dim COMPPN, UID As String, Item As String, FuncType As String, Value As String, delflag As String
Dim Rs As ADODB.Recordset
Dim blerr As Boolean
blerr = False

If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
    Exit Sub
End If
On Error GoTo errhandle:
    Set xlApp = CreateObject("Excel.Application")
    Let xlApp.Visible = False
    Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
    xlApp.DisplayAlerts = False
    rowcnt = 2
    UID = Trim(g_userName)
    
    With xlsBook.Worksheets(Trim(Shift_Item))
        While Trim(.Cells(rowcnt, 1)) <> ""
            COMPPN = Replace(Trim(.Cells(rowcnt, 1)), "*", "")
            Item = Trim(.Cells(rowcnt, 2))
            FuncType = Trim(.Cells(rowcnt, 3))
            Value = Trim(.Cells(rowcnt, 4))
            delflag = Trim(.Cells(rowcnt, 5))
            If Trim(Item) = "" Then  ''(1028)
                MsgBox "The item can not be null,please add factory in it!!!"
                Set xlApp = Nothing
                Set xlsBook = Nothing
                Exit Sub
            End If
            'save data to DB
            If UCase(delflag) = "Y" Then
                strSQL = "delete from Component_Data where CompPN='" & COMPPN & "' and item='" & Item & "' and Functype='" & FuncType & "' "
                Conn.Execute (strSQL)
            Else
                strSQL = "select 0 from Component_Data where CompPN='" & COMPPN & "' and item='" & Item & "' and Functype='" & FuncType & "' "
                Set Rs = Conn.Execute(strSQL)
                If Rs.EOF Then
                    strSQL = "insert into Component_Data(compPN,item,functype,value,UID,transdatetime) " & _
                             "values('" & COMPPN & "','" & Item & "','" & FuncType & "','" & Value & "','" & UID & "',[dbo].[FormatDate](getdate(),'YYYYMMDDHHNNSS') ) "
                    Conn.Execute (strSQL)
                Else
                    .Cells(rowcnt, 6).Interior.ColorIndex = 3
                    .Cells(rowcnt, 6).Interior.Pattern = xlSolid
                    .Cells(rowcnt, 6) = "this CompPN have been exists!"
                End If

            End If
            rowcnt = rowcnt + 1
        Wend
    End With
    strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Upload_component_data','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
    Conn.Execute (strSQL)
    xlsBook.Close
    xlApp.Quit
    Set xlsBook = Nothing
    Set xlApp = Nothing
    MsgBox ("***upload finished!***")
    Exit Sub
errhandle:
    MsgBox Err.Description
    
End Sub

Private Sub Load_WO_AssignPN(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlsBook As Excel.Workbook
Dim xlWs As Excel.Worksheets
Dim rCount As Long
Dim AssignPN, WO, MBPN, Version, COMPPN, Vendor As String, delflag As String
Dim Insert_Qty As Long, Update_Qty As Long, Del_Qty As Long
Dim strSQL As String
Dim Rs As ADODB.Recordset
  
If Trim(cboSheetName) = "" Or Trim(txtFilePath) = "" Then
   Exit Sub
End If
  
Set xlApp = CreateObject("Excel.Application")
Let xlApp.Visible = False
Set xlsBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.DisplayAlerts = False

rCount = 2
Insert_Qty = 0
Update_Qty = 0
Del_Qty = 0

With xlsBook.Worksheets(Trim(Shift_Item))

    While Trim(.Cells(rCount, 1)) <> ""
    
        WO = Replace(Trim(.Cells(rCount, 1) & vbNullString), vbCrLf, "")
        MBPN = Trim(.Cells(rCount, 2) & vbNullString)
        Version = Replace(Trim(.Cells(rCount, 3) & vbNullString), "'", " ")
        COMPPN = Replace(Trim(.Cells(rCount, 4) & vbNullString), " ", "")
        AssignPN = Trim(.Cells(rCount, 5) & vbNullString)
        Vendor = Trim(.Cells(rCount, 6) & vbNullString)
        delflag = Trim(.Cells(rCount, 7) & vbNullString)
        ''Check Workorder
        If IsNumeric(WO) = False Then
           MsgBox "The Work order:" & WO & " must be numeric !! ", vbCritical
           xlsBook.Close
           xlApp.Quit
           Set xlApp = Nothing
           Set xlsBook = Nothing
           Exit Sub
        End If
        ''Check unuseful data  (0066)
        If COMPPN = AssignPN And Vendor = "" Then
            MsgBox "Please check the unuseful data (CompPN=AssignPN and Vendor='') !! ", vbCritical
            xlsBook.Close
            xlApp.Quit
            Set xlApp = Nothing
            Set xlsBook = Nothing
            Exit Sub
        End If
        ''Check Wo's modelname
        strSQL = "select rtrim(MBPN)+'-'+rtrim(Rev) as model from WO_AssignPN_Vendor where wo='" & Trim(WO) & "'"
        Set Rs = Conn.Execute(strSQL)
        If Rs.EOF = False Then
           If UCase(Rs!Model) <> UCase(Trim(MBPN)) + "-" + UCase(Trim(Version)) Then
              MsgBox "WO: " & WO & "'s model is not match with last time uploaded:" & Rs!Model & ", check it or del the record which last time uploaded first!"
              xlsBook.Close
              xlApp.Quit
              Set xlApp = Nothing
              Set xlsBook = Nothing
              Exit Sub
           End If
        End If
        ''Save data to DB
        If UCase(Trim(delflag)) = "Y" Then
           strSQL = "delete from WO_AssignPN_Vendor where WO='" & WO & "' and AssignedCompPN='" & AssignPN & "'"
           Conn.Execute (strSQL)
           Del_Qty = Del_Qty + 1
        Else
           ''Exists or not
           strSQL = "select 0 from WO_AssignPN_Vendor where WO='" & WO & "' and AssignedCompPN='" & AssignPN & "'"
           Set Rs = Conn.Execute(strSQL)
           If Rs.EOF Then
              ''Insert
              strSQL = "insert into WO_AssignPN_Vendor(WO,MBPN,Rev,CompPN,AssignedCompPN,VendorCode,[UID],Transdatetime) Values " & _
               " ('" & WO & "','" & MBPN & "','" & Version & "','" & COMPPN & "','" & AssignPN & "','" & Vendor & "','" & Trim(g_userName) & "',[DBO].[FormatDate](getdate(),'YYYYMMDDHHNNSS') )"
              Conn.Execute (strSQL)
              Insert_Qty = Insert_Qty + 1
           Else
              ''Update
              strSQL = "update WO_AssignPN_Vendor set MBPN='" & MBPN & "',Rev='" & Version & "',CompPN='" & COMPPN & "',VendorCode='" & Vendor & "',UID='" & Trim(g_userName) & "',Transdatetime=[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS') " & _
              "where WO='" & WO & "' AND AssignedCompPN='" & AssignPN & "'"
              Conn.Execute (strSQL)
              Update_Qty = Update_Qty + 1
           End If
        End If
        Set Rs = Nothing
        rCount = rCount + 1
       
    Wend
 
End With
     
strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_WO_AssignPN','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSQL)

xlsBook.Close
xlApp.Quit
Set xlApp = Nothing
Set xlsBook = Nothing
  
Total_Qty = Insert_Qty + Update_Qty

 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Update succeed : " & Update_Qty & vbCrLf & _
               "Delete succeed : " & Del_Qty & vbCrLf
              
End Sub

Private Sub Upload_MachineData(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim Rs As New ADODB.Recordset
Dim i As Integer
Dim strSQL As String
Dim L As Integer
Dim Line As String
On Error GoTo errhandle:

Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.Visible = False
xlApp.UserControl = True
xlApp.DisplayAlerts = False

Line = xlBook.Worksheets(Trim(Shift_Item)).Cells(2, 1)
If Line <> "" Then
    strSQL = "delete from machine_data where line='" & Line & "' and FuncType='DualLaneMode' and Item='Independent'"  '1155
    If Rs.State = 1 Then Rs.Close
    Set Rs = Conn.Execute(strSQL)
End If
i = 2

With xlBook.Worksheets(Trim(Shift_Item))
    Do While Not (.Cells(i, 1) = "")
'        strsql = "delete from machine_data where line='" & Trim(.Cells(i, 1)) & "'  and FuncType='DualLaneMode' and Item='Independent'" '(1149)
'        Conn.Execute (strsql)
'
        strSQL = "Insert Into machine_data(Line,Machine,TransDateTime,FuncType,value,Item,UID) " & _
                 "Values ('" & Trim(.Cells(i, 1)) & "','" & Trim(.Cells(i, 2)) & "',DBO.FormatDate(Getdate(),'YYYYMMDDHHNNSS'),'DualLaneMode','Y','Independent','" & Trim(g_userName) & "')" ''(1055)
        Conn.Execute (strSQL)
        i = i + 1
    Loop
End With
Set Rs = Nothing

xlApp.Quit
Set xlApp = Nothing

MsgBox "Excel Data:" & i - 2 & vbCrLf & "Upload:" & L
Exit Sub

errhandle:
    MsgBox Err.Description
End Sub

Private Sub Upload_CompPNSpacer(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim Rs As New ADODB.Recordset
Dim i As Integer
Dim strSQL As String
Dim L As Integer
On Error GoTo errhandle:

Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.Visible = False
xlApp.UserControl = True
xlApp.DisplayAlerts = False


i = 2

With xlBook.Worksheets(Trim(Shift_Item))
    Do While Not (.Cells(i, 1) = "")
        If (Trim(.Cells(i, 3)) = "Y") Then
            strSQL = "delete from  CompPN_BaseData where CompPN='" & Trim(.Cells(i, 1)) & "'  and Type='Spacer'" '
            Conn.Execute (strSQL)
        Else
            strSQL = "select Value from CompPN_BaseData where CompPN='" & Trim(.Cells(i, 1)) & "'  and Type='Spacer'"
            Set Rs = Conn.Execute(strSQL)
            If Rs.EOF = False Then
                If (Trim(.Cells(i, 2)) <> Trim(Rs!Value)) Then
                    MsgBox "CompPN:" & .Cells(i, 1) & "'Spacer is different from last upload,please check!"
                End If
            Else
                strSQL = "Insert Into CompPN_BaseData(CompPN,Type,TransDateTime,value,UID) " & _
                         "Values ('" & Trim(.Cells(i, 1)) & "','Spacer',DBO.FormatDate(Getdate(),'YYYYMMDDHHNNSS'),N'" & Trim(.Cells(i, 2)) & "','" & Trim(g_userName) & "')"
                Conn.Execute (strSQL)
            End If
        End If
        i = i + 1
    Loop
End With
Set Rs = Nothing

xlApp.Quit
Set xlApp = Nothing

 MsgBox ("***upload finished!***")
Exit Sub

errhandle:
    MsgBox Err.Description
End Sub

Private Sub Upload_AVLC(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim Rs As New ADODB.Recordset
Dim i As Integer
Dim strSQL As String
Dim L As Integer
On Error GoTo errhandle:

Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.Visible = False
xlApp.UserControl = True
xlApp.DisplayAlerts = False


i = 2
strSQL = "delete from  QSMS_CustomerData_Temp"
Conn.Execute (strSQL)
With xlBook.Worksheets(Trim(Shift_Item))
    Do While Not (.Cells(i, 2) = "")
            strSQL = "Insert Into QSMS_CustomerData_Temp(KC,CustomerPN,CompPN,TransDateTime,UID) " & _
                         "Values ('" & Trim(.Cells(i, 2)) & "','" & Trim(.Cells(i, 6)) & "','" & Trim(.Cells(i, 7)) & "',DBO.FormatDate(Getdate(),'YYYYMMDDHHNNSS'),'" & Trim(g_userName) & "')"
            Conn.Execute (strSQL)
            i = i + 1
    Loop
End With
Set Rs = Nothing

strSQL = "delete from QSMS_CustomerData"
Conn.Execute (strSQL)
strSQL = "delete from QSMS_CustomerData_Temp output deleted.* into QSMS_CustomerData"
Conn.Execute (strSQL)
xlApp.Quit
Set xlApp = Nothing

 MsgBox ("***upload finished!***")
Exit Sub

errhandle:
    MsgBox Err.Description
End Sub
Private Sub Upload_A8_Manual(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim Rs As New ADODB.Recordset
Dim i As Integer
Dim strSQL As String
Dim L As Integer
On Error GoTo errhandle:

Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.Visible = False
xlApp.UserControl = True
xlApp.DisplayAlerts = False


i = 2
strSQL = "delete from  Asce_ManualUpload"
Conn.Execute (strSQL)
With xlBook.Worksheets(Trim(Shift_Item))
    Do While Not (.Cells(i, 2) = "")
            strSQL = "Insert Into Asce_ManualUpload(CompPN,Qty,TransDateTime,UID) " & _
                         "Values ('" & Trim(.Cells(i, 1)) & "','" & Trim(.Cells(i, 2)) & "',DBO.FormatDate(Getdate(),'YYYYMMDDHHNNSS'),'" & Trim(g_userName) & "')"
            Conn.Execute (strSQL)
            i = i + 1
    Loop
End With
Set Rs = Nothing

strSQL = "exec Asce_Shelf_ManualDispatch"
Conn.Execute (strSQL)

xlApp.Quit
Set xlApp = Nothing

 MsgBox ("***upload finished!***")
Exit Sub

errhandle:
    MsgBox Err.Description
End Sub
Private Sub Upload_A8_DIDType(Shift_Item As String)
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim Rs As New ADODB.Recordset
Dim i As Integer
Dim strSQL As String
Dim L As Integer
On Error GoTo errhandle:

Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Open(txtFilePath)
xlApp.Visible = False
xlApp.UserControl = True
xlApp.DisplayAlerts = False

'Dim sheet2 As Excel.Worksheet
'Set sheet2 = xlBook.Worksheets(0)

i = 2
With xlBook.Worksheets(Trim(Shift_Item))
    
    Do While Not (.Cells(i, 1) = "")
       
        strSQL = "EXEC A8_DIDType_Insert " & _
                 "@Comp='" & Trim(.Cells(i, 1)) & "'" & _
                 ",@Vendor='" & Trim(.Cells(i, 2)) & "'" & _
                 ",@Size='" & Trim(.Cells(i, 3)) & "'" & _
                 ",@PackingType='" & Trim(.Cells(i, 4)) & "'" & _
                 ",@DeleteFlag='" & Trim(.Cells(i, 5)) & "'" & _
                 ",@UID='" & g_userName & "'"
        Conn.Execute (strSQL)
        i = i + 1
    Loop
End With
Set Rs = Nothing

xlApp.Quit
Set xlApp = Nothing

 MsgBox ("***upload finished!***")
Exit Sub

errhandle:
    MsgBox Err.Description
End Sub


