VERSION 5.00
Begin VB.Form FrmTransferPhilips 
   Caption         =   "Uplaod Philips Machine Bom[20070920]"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   11820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   9120
      TabIndex        =   15
      Top             =   960
      Width           =   855
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
      Left            =   1680
      TabIndex        =   13
      Top             =   840
      Width           =   5055
   End
   Begin VB.Frame Frame1 
      Caption         =   "File select"
      Height          =   4095
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   11775
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
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton cmdDELALL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete All"
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
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4080
         Width           =   1335
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
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1560
         Width           =   495
      End
      Begin VB.ListBox ListFile 
         Height          =   2595
         Left            =   6360
         TabIndex        =   8
         Top             =   1080
         Width           =   3375
      End
      Begin VB.FileListBox File1 
         Height          =   2625
         Left            =   2760
         Pattern         =   "*.XLS"
         TabIndex        =   7
         Top             =   1080
         Width           =   2775
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2535
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   375
      Left            =   12000
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtFilePath 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   10095
   End
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   11880
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   12
      Top             =   600
      Width           =   1200
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Upload"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "File Name"
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
      Left            =   0
      TabIndex        =   14
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "Sheet Name"
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
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "FrmTransferPhilips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Private Type Jax
    Seq As Integer
    Jobpn As String
    Rev As String
    CompPN As String
    LR As String
    Slot As String
    Qty As Integer
    Enabled As Boolean
    Machine As String
End Type
Public Function LoadDataFile(FilePath As String) As Boolean
    Dim fso As New Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim buf As String
    Dim FileName As String
    Dim temp() As String
    
    'On Error Resume Next
        
    temp = Split(FilePath, "\")
    FileName = temp(UBound(temp))
        
    temp = Split(FileName, "-")
    If UBound(temp) <> 4 Then
        MsgBox ("FileName format must be Machine-PN-REV-BuildType-Side.xls!")
        LoadDataFile = False
        Exit Function
    End If
    Call Load_Philips_Machine_BOM(Trim(cboSheetName))
    LoadDataFile = True
End Function

Private Sub Load_Philips_Machine_BOM(Shift_Item As String)
  Dim xlApp As Excel.Application
  Dim xlsBook As Excel.Workbook
  Dim xlWs As Excel.Worksheets
  Dim aryPN() As String
  Dim rCount, Row_Count As Long
  Dim Machine, Jobpn, Version, Slot, UpcompPN, CompPN, Qty, LR, JobGroup, Str1, StrSlot As String, BuildType As String, Side As String, errMsg As String
  Dim Trolley, Lane, PartNumber, Count, tempmachine, TempJobPn, BrdSeq, tempVersion As String, TempJObGroup As String
  Dim Total_Qty, Update_Qty, Insert_Qty As Long
  Dim temp() As String
  Dim strSQL As String
  Dim AryBrdSeq(200, 1) As Integer
  Dim Str As String
  Dim i, j, m, n As Integer
  Dim k As Integer
  Dim flag As Integer
  Dim Seq As Integer
  Dim rs As ADODB.Recordset
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
  tempmachine = ""
  TempJobPn = ""
  tempVersion = ""
  TempJObGroup = ""
  
  With xlsBook.Worksheets(Trim(Shift_Item))
  
    ' Macro3 Macro
    ' 宏由 Administrator 录制，时间: 2006/9/22
'    '*************** Sort MEBom start ***********************************
'    .Columns("A:A").Select
'    .Range("A1:J430").Sort Key1:=.Range("A1"), Order1:=xlAscending, Header:= _
'        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
'        SortMethod:=xlPinYin ', DataOption1:=xlSortNormal
'    .Columns("B:B").Select
'    .Range("A1:J430").Sort Key1:=.Range("B1"), Order1:=xlAscending, Header:= _
'        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
'        SortMethod:=xlPinYin ', DataOption1:=xlSortNormal
'    .Columns("C:C").Select
'    .Range("A1:J430").Sort Key1:=.Range("C1"), Order1:=xlAscending, Header:= _
'        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
'        SortMethod:=xlPinYin ', DataOption1:=xlSortNormal
'    .Columns("H:H").Select
'    .Range("A1:J430").Sort Key1:=.Range("E1"), Order1:=xlAscending, Header:= _
'        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
'        SortMethod:=xlPinYin ', DataOption1:=xlSortNormal
'  '  **************** Sort MEBom end *************************************


    While Trim(.Cells(rCount, 1)) <> ""
        i = 0
        m = 0
        k = 0
        Trolley = Trim(.Cells(rCount, 1) & vbNullString)
        Slot = Trim(.Cells(rCount, 2) & vbNullString)
        Lane = Trim(.Cells(rCount, 3) & vbNullString)
        PartNumber = Trim(.Cells(rCount, 5) & vbNullString)
        Count = Trim(.Cells(rCount, 6) & vbNullString)
'        BrdSeq = CStr(Left(Trim(.Cells(rCount, 7) & vbNullString), 1))
        Str = CStr(Trim(.Cells(rCount, 7) & vbNullString))
        aryPN = Split(Trim(Str), ",")
        If CStr(UBound(aryPN) + 1) <> Count Then
             MsgBox (" The Electrical reference number of " & sq(PartNumber) & "is not right,Please check the EXCEL file!")
        End If
        
        temp = Split(Trim(File1.FileName), "-")
        Machine = Trim(temp(0))
        Jobpn = Trim(temp(1))
        Version = Trim(temp(2))
        JobGroup = Jobpn & "-" & Version
        BuildType = Trim(temp(3))
        Side = Trim(Left(temp(4), Len(temp(4)) - 4))
        
        CompPN = PartNumber
        Slot = Trolley & "-" & Slot
        '---------------------------------------------------
        Slot = Replace(Slot, " ", "")
        Debug.Print Slot
        StrSlot = "0123456789-"
        For n = 1 To Len(Slot)
            Str1 = Mid(Slot, n, 1)
            flag = InStr(StrSlot, Str1)
'            Debug.Print Flag
            If (flag < 1) Then
                MsgBox (" The Trolley or Slot in row " & sq(rCount) & " is not Numeral,Please check the EXCEL file Contents!")
                Exit Sub
            End If
        Next
      '----------------------------------------------------
'        Qty = Count
        LR = Lane
        If LR <> "0" And LR <> "1" And LR <> "2" Then
            MsgBox (" Lane=" & sq(LR) & " is not right,Please check the EXCEL file!")
            Exit Sub
        End If
'        Debug.Print Qty
'        If TopGuide = "Left" Then LR = "1"
'        If TopGuide = "Right" Then LR = "2"
'        If TopGuide = "N/A" Then LR = "0"
        'Load SeqBoardNum mapping data brdseq
        
        If Qty <> "0" Then
        '***********************************************************
            For j = 0 To UBound(aryPN)
                Seq = Mid(Trim(aryPN(j) & vbNullString), 1, InStr(1, aryPN(j), "-") - 1)
'                Debug.Print Seq
                If j = 0 Then
                    AryBrdSeq(0, 0) = Seq
                    AryBrdSeq(0, 1) = 1
                Else
                    For m = 0 To k
                        If AryBrdSeq(m, 0) = Seq Then
                            AryBrdSeq(m, 1) = AryBrdSeq(m, 1) + 1
                            GoTo Handler
                        End If
                    Next
                     k = k + 1
                     AryBrdSeq(k, 0) = Seq
                     AryBrdSeq(k, 1) = 1
Handler:             Debug.Print AryBrdSeq(k, 0)
                End If
            Next
        '***************************************************************
            While i <= k
                BrdSeq = CStr(AryBrdSeq(i, 0))
'                Debug.Print BrdSeq
                Qty = CStr(AryBrdSeq(i, 1))
'                Debug.Print Qty
                strSQL = "select * from PhilipsBrdSeqMapping where JobPN=" & sq(Jobpn) & " and brdseq=" & sq(1) & " and Rev=" & sq(Version)
                Set rs = Conn.Execute(strSQL)
                If rs.EOF = False Then
                    Jobpn = rs("BrdPN")
                    Version = rs("BrdRev")
                    JobGroup = Jobpn & "-" & Version
                End If
                Jobpn = Trim(temp(1))
                Version = Trim(temp(2))
                strSQL = "select * from PhilipsBrdSeqMapping where JobPN=" & sq(Jobpn) & " and brdseq=" & sq(BrdSeq) & " and Rev=" & sq(Version)
                Set rs = Conn.Execute(strSQL)
                If rs.EOF = True Then
                    MsgBox "because JobPN=" & sq(Jobpn) & " And BRDSEQ = " & BrdSeq & " is not define in PhilipsBrdSeqMapping table ," & Chr(10) & "Please check the Excel file."
                    Exit Sub
                Else
                    Jobpn = rs("BrdPN")
                    Version = rs("BrdRev")
                End If
                If (TempJObGroup = "" Or TempJObGroup <> JobGroup) Or (tempmachine = "" Or tempmachine <> Machine) Then
                   strSQL = "delete from QSMS_MEBom where Jobgroup='" & JobGroup & "' and Jobpn='" & Jobpn & "' and machine='" & Machine & "' and Version='" & Version & "' and BuildType='" & BuildType & "' and Side='" & Trim(Side) & "'"
                   Conn.Execute strSQL
                End If
                    strSQL = "select * from QSMS_MEBom where Jobgroup='" & JobGroup & "' and Machine='" & Trim(Machine) & "' and JobPN='" & Trim(Jobpn) & "'  and Version='" & Version & "'  and Slot='" & Slot & "' and LR='" & LR & "' and BuildType='" & BuildType & "'"
                    Set rs = Conn.Execute(strSQL)
                    If rs.EOF Then
                        strSQL = "Insert into QSMS_MEBom(Machine,JobPN,Version,CompPN,LR,Slot,Qty,JobGroup,BuildType,Side,UID) " & _
                                 " values('" & Trim(Machine) & "','" & Trim(Jobpn) & "','" & Trim(Version) & "', " & _
                                 " '" & CompPN & "','" & LR & "','" & Slot & "'," & Qty & ",'" & JobGroup & "','" & BuildType & "','" & Side & "','" & g_userName & "')"
                        Conn.Execute strSQL
                        Insert_Qty = Insert_Qty + 1
                       
                    Else
                        strSQL = "Update QSMS_MEBom set CompPN='" & Trim(CompPN) & "',Qty=" & Qty & ",BuildType='" & Trim(BuildType) & "',Side='" & Trim(Side) & "',UID='" & g_userName & "',TransDateTime=convert(char(8),getdate(),112) + left(replace(convert(char(8),getdate(),108), ':', ''),6),Jobgroup='" & JobGroup & "' " & _
                                 " where Machine='" & Trim(Machine) & "' and JobGroup='" & Trim(JobGroup) & "' and JobPN='" & Trim(Jobpn) & "' and Version='" & Version & "'  and LR='" & LR & "' and Slot='" & Slot & "' and BuildType='" & BuildType & "'"
                        Conn.Execute strSQL
                        
                        Update_Qty = Update_Qty + 1
                        DoEvents
                   End If
                   i = i + 1
            Wend
        End If
        tempmachine = Machine
        TempJObGroup = JobGroup
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
         
        rCount = rCount + 1
        Total_Qty = Total_Qty + 1
            
    Wend
End With
strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_QSMS_BOM','" & Left(Trim(txtFilePath), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSQL)

 xlsBook.Close
  xlApp.Quit
  Set xlApp = Nothing
  Set xlsBook = Nothing
 MsgBox "*** Load  finish ! ***" & Shift_Item & "   " & vbCrLf & _
               "Total Counter : " & Total_Qty & vbCrLf & _
               "Insert succeed : " & Insert_Qty & vbCrLf & _
               "Update succeed : " & Update_Qty & vbCrLf
              
End Sub


Private Sub CmdAdd_Click()
    Dim Pointer As Integer
    Dim FileName As String
    Dim i As Integer
    Dim temp() As String
    Dim Str As String, Machine As String, MBPN As String, Version As String, JobGroup As String

    If File1.ListCount <= 0 Then Exit Sub
    If File1.ListIndex < 0 Then Exit Sub
    If ListFile.ListCount >= 1 Then
        MsgBox ("please select one Excel file and upload,then select the other one.")
        Exit Sub
    End If
    Pointer = File1.ListIndex
    ListFile.AddItem Trim(File1.FileName)
    If File1.ListCount <> Pointer Then
       File1.ListIndex = Pointer
    End If
    If ListFile.ListCount <= 0 Then Exit Sub
    FileName = File1.Path & "\" & File1.FileName
    txtFilePath = FileName
    cboSheetName.Clear
    Call ReadAllSheetName(txtFilePath)
    cboSheetName.Enabled = True
    If cboSheetName.ListCount > 0 Then
        cboSheetName.ListIndex = 0
    End If
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
Private Sub cmdADDALL_Click()
Dim i As Integer
    If File1.ListCount <= 0 Then Exit Sub
    
    For i = 0 To File1.ListCount - 1
  
      File1.ListIndex = i
      ListFile.AddItem Trim(File1.FileName)
   
   
  Next i
End Sub

Private Sub cmdDEL_Click()
    Dim Pointer As Integer
    If ListFile.ListCount <= 0 Then Exit Sub
    If ListFile.ListIndex < 0 Then Exit Sub
      Pointer = ListFile.ListIndex
      ListFile.RemoveItem Pointer
        If ListFile.ListCount <> Pointer Then
           ListFile.ListIndex = Pointer
        End If
        
 
End Sub

Private Sub cmdDELALL_Click()
ListFile.Clear
End Sub

Private Sub cmdLoad_Click()
Dim FileName As String
Dim i As Integer
Dim temp() As String
Dim Str As String, Machine As String, MBPN As String, Version As String, JobGroup As String
'    If I = 0 Then
'        temp = Split(Trim(File1.FileName), "-")
'        If UBound(temp) <> 4 Then
'            MsgBox ("FileName format must be Machine-PN-REV-BuildType-Side.xls!")
'            Exit Sub
'        End If
'
'    End If
    
    If LoadDataFile(Trim(txtFilePath)) = False Then MsgBox ("             Fail" & Chr(10) & "please check the Excel file!")
    ListFile.Clear
    DoEvents
End Sub
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
 File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

