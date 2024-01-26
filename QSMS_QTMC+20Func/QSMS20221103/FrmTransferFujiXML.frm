VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmTransferFujiXML 
   Caption         =   "frm transfer FUji XML file to ME BOM[20171102]"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "File select"
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   11775
      Begin VB.CommandButton cmdDEL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete"
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
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
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
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1800
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
         TabIndex        =   10
         Top             =   1680
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
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2160
         Width           =   495
      End
      Begin VB.ListBox ListFile 
         Height          =   2400
         Left            =   6360
         TabIndex        =   8
         Top             =   1080
         Width           =   3375
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Left            =   2760
         Pattern         =   "*.XML"
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
      Left            =   11760
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   10095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11880
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Upload"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   840
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
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "FrmTransferFujiXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**********************************************************************************
'**文 件 名: frmTransferFUJIXML.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: ArcherYang
'**日    期: 2009.07.08
'**描    述: Get QSMS_MEBOM
'
'**EQMS_ID      修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**QMS          Archer      20090708        save the machine bom's location data to QSMS_MEBOM  (0001)
'***********************************************************************************/
Option Explicit
Dim strSql As String
Private Type Jax
    Seq As Integer
    Jobpn As String
    Rev As String
    compPN As String
    LR As String
    Slot As String
    Qty As Integer
    Enabled As Boolean
    machine As String
    location As String          '''(0001)
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
    If UBound(temp) <> 6 Then
        MsgBox ("FileName format must be Factory-Line-Machine-PN-REV-BuildType-Side.xml!")   '(0007)
        LoadDataFile = False
        Exit Function
    End If
    
    'Read in file content
    Set ts = fso.OpenTextFile(FilePath)
    
    If Err Then
        LoadDataFile = False
        Set fso = Nothing
        Exit Function
    End If
    
    Err.Clear
    
    buf = ts.ReadAll
    
    ts.Close
    Set ts = Nothing

    Call LoadArrayWithData(FileName, buf)

    Set fso = Nothing
    LoadDataFile = True
End Function

Private Sub LoadArrayWithData(FileName As String, s As String)
    Dim Data() As Jax
    Dim num As Integer, TabQty As Integer, SlotQty As Integer
    Dim p1 As Long, p2 As Long
    Dim temp() As String, temp1() As String
    Dim Factory As String, Line As String, machine As String, Version As String, MBPN As String, BuildType As String, side As String, jobgroup As String, strLR As String, StrSlot As String
    Dim str As String
    Dim BrdPN(500) As String, BrdRev(500) As String, NXT As Boolean, AIMEX As Boolean
    Dim rs As ADODB.Recordset
    Dim strLocation As String
    
    temp = Split(FileName, "-")
    If UBound(temp) <> 6 Then
        MsgBox ("FileName format must be Factory-Line-Machine-PN-REV-BuildType-Side.xml!")   '(0007)'(1026)
        Exit Sub
    End If
    machine = Trim(temp(2))
    Line = Trim(temp(1))
    side = Trim(Left(temp(6), Len(temp(6)) - 4))
    If InStr(machine, "NXT") > 0 Then
       NXT = True
    ElseIf InStr(machine, "AIMEX") > 0 Then   ''''(1161)
       AIMEX = True
    Else
       NXT = False
    End If
    If NXT = False And AIMEX = False Then
        'QMS             Denver         2011/01/04     can not Upload Fuji XML data                                   (1042)
        strSql = "select * from Machine where Machine='" & machine & "' and Line='" & Line & "'and side='" & side & "'"  '(1043)
        Set rs = Conn.Execute(strSql)
        If rs.EOF Then
           MsgBox "Can't find The Machine:" & machine & " in line:" & Line & " and side: " & side & " ,please define it in machinetype or check its format!", vbCritical, "ErrMessage"
           Exit Sub
        Else
           TabQty = rs.Fields("Qty")
           SlotQty = rs.Fields("MaxSlotNum")
        End If
    End If
    
'    If Len(Machine) <> 6 Then
'        MsgBox ("Machine name: " & Machine & " length must be 6!")
'        Exit Sub
'    End If
    Factory = Trim(temp(0))    'add by giant 2008/06/27 (0007)
    MBPN = Trim(temp(3))
    Version = Trim(temp(4))
    jobgroup = MBPN & "-" & Version
    BuildType = Trim(temp(5))
    
    
    'Load SeqBoardNum mapping data
    strSql = "select * from FujiBrdSeqMapping where JobPN=" & sq(MBPN) & " and Rev=" & sq(Version)
    Set rs = Conn.Execute(strSql)
    Do While Not rs.EOF
        BrdPN(Val(rs("BrdSeq"))) = rs("BrdPN")
        BrdRev(Val(rs("BrdSeq"))) = rs("BrdRev")
        rs.MoveNext
    Loop
    
    p1 = InStr(s, "<Unit>")
    If num = 0 Then
        p1 = InStr(p1 + Len("<Unit"), s, "<Unit")
    End If
    
    Do While True
        num = num + 1
        ReDim Preserve Data(num)
        
        Data(num).Seq = num
        Data(num).compPN = StrBetween(s, "<seqPartNum>", "</seqPartNum>", p1)
        Data(num).compPN = Replace(Data(num).compPN, "&lt;", "<")
        
        '******************************
        '****add by jeanson 2007/09/03
        strErrMessage = ""
        strErrMessage = FunPartNumberCheck(Trim(Data(num).compPN))
        If strErrMessage <> "PASS" Then
            MsgBox strErrMessage
        Exit Sub
        End If
        '******************************
        
        str = StrBetween(s, "<seqBrdNum>", "</seqBrdNum>", p1)
        If Not IsNumeric(str) Then
            MsgBox ("The SeqBoardNum:" & str & " must be numeric!")
            Exit Sub
        End If
        'Get PN & Rev from PN-REV
        If BrdPN(Val(str)) <> "" Then
            Data(num).Jobpn = BrdPN(Val(str))
            Data(num).Rev = BrdRev(Val(str))
        Else
            MsgBox ("Can not find the SeqBoardNum mapping:" & str)
            Exit Sub
        End If

        If Len(Data(num).Rev) > 5 Then
            MsgBox ("Rev:" & Data(num).Rev & " is too long!")
            Exit Sub
        End If
        
        ''Get Location Data '''(0001)
        strLocation = StrBetween(s, "<seqRef>", "</seqRef>", p1)
        Data(num).location = strLocation
        
        If ChkMEBOM_Location = "Y" And Trim(Data(num).location) = "" Then  ''''(1250)
            MsgBox ("Location:" & Data(num).location & " can not be empty!")
            Exit Sub
        End If
        
        p1 = InStr(p1, s, "<fsSetPos>") + Len("<fsSetPos>")
        p2 = InStr(p1, s, "</fsSetPos>")
        
        StrSlot = Replace(Mid(s, p1, p2 - p1), " ", "")
        temp1 = Split(StrSlot, "-")
        
         ''必须为 数字 - 数字 格式  1215 begin
        StrBU = ReadIniFile("COMMON", "BU", App.Path & "\set.ini")
        If StrBU = "NB6" Then
            If NXT = True Then
                If UBound(temp1) <> 1 Or IsNumeric(temp1(0)) = False Or IsNumeric(temp1(1)) = False Then
                 MsgBox "The NXT Machine slot" & StrSlot & "format is wrong,please check!", vbCritical, "ErrMessage"
                       Exit Sub
                 End If
           End If
        End If
        ''1215 end
        
        Select Case UBound(temp1)
               Case 0
                    If NXT = True Then
                       MsgBox "The NXT Machine slot" & StrSlot & "format is wrong,please check!", vbCritical, "ErrMessage"
                       Exit Sub
                    End If
                    strLR = "0"
                    If Int(temp1(0)) > 240 Or Int(temp1(0) < 0) Then
                       MsgBox "The Slot " & StrSlot & " foramt is wrong,please check!", vbCritical, "ErrMessage"
                       Exit Sub
                    Else
                       Data(num).Slot = temp1(0)
                       Data(num).machine = machine
                    End If
               Case 1
                    If temp1(1) > 100 And NXT = False Then   '(1143)
                       temp1(1) = temp1(1) - 100
                    End If
                    If NXT = False Then
                       If Int(temp1(0)) > TabQty Or Int(temp1(1)) > SlotQty Then
                          MsgBox "The Slot " & StrSlot & " foramt is wrong,please check!", vbCritical, "ErrMessage"
                          Exit Sub
                       Else
                          Data(num).Slot = temp1(0) + "-" + temp1(1)
                          Data(num).machine = machine
                          strLR = "0"
                       End If
                    Else
                       If NXT = True Then
                           If Len(temp1(0)) = 1 Then
                              Data(num).machine = machine + "0" + temp1(0)
                           Else
                              Data(num).machine = machine + temp1(0)
                           End If
                              strSql = "select * from Machine where Machine='" & Data(num).machine & "' and Line='" & Line & "'and side='" & side & "'"
                              Set rs = Conn.Execute(strSql) '(1043)
                           If rs.EOF Then
                              MsgBox "Can't find The Machine:" & Data(num).machine & " in line:" & Line & " and side: " & side & " ,please define it in machinetype or check its format!", vbCritical, "ErrMessage"
                             Exit Sub
                           Else
                              TabQty = rs.Fields("Qty")
                              SlotQty = rs.Fields("MaxSlotNum")
                           End If
                           If TabQty <> 1 Or Int(temp1(1)) > SlotQty Then
                              MsgBox "The Slot " & StrSlot & " foramt is wrong,please check!", vbCritical, "ErrMessage"
                              Exit Sub
                           End If
                           Data(num).Slot = temp1(1)
                           strLR = "0"
                       End If
                    End If
               Case 2
                    If temp1(1) > 100 And NXT = False Then  '(1143)
                       temp1(1) = Trim(Int(temp1(1)) - 100)
                    End If
                    If NXT = False And AIMEX = False Then
                       If (temp1(2) <> "1" And temp1(2) <> "2") Or Int(temp1(0)) > TabQty Or Int(temp1(1)) > SlotQty Then
                          MsgBox "The Slot " & StrSlot & " foramt is wrong,please check!", vbCritical, "ErrMessage"
                          Exit Sub
                       Else
                          strLR = temp1(2)
                          Data(num).Slot = temp1(0) + "-" + temp1(1)
                          Data(num).machine = machine
                       End If
                    ElseIf NXT = True Then
                           If Len(temp1(0)) = 1 Then
                              Data(num).machine = machine + "0" + temp1(0)
                           Else
                              Data(num).machine = machine + temp1(0)
                           End If
                           strSql = "select * from Machine where Machine='" & Data(num).machine & "' and Line='" & Line & "'and side='" & side & "'"
                           Set rs = Conn.Execute(strSql) '(1043)
                           If rs.EOF Then
                              MsgBox "Can't find The Machine:" & Data(num).machine & " in line:" & Line & " and side: " & side & " ,please define it in machinetype or check its format!", vbCritical, "ErrMessage"
                             Exit Sub
                           Else
                              TabQty = rs.Fields("Qty")
                              SlotQty = rs.Fields("MaxSlotNum")
                           End If
                           If (temp1(2) <> "1" And temp1(2) <> "2") Or TabQty <> 1 Or Int(temp1(1)) > SlotQty Then
                              MsgBox "The Slot " & StrSlot & " foramt is wrong,please check!", vbCritical, "ErrMessage"
                              Exit Sub
                           End If
                           Data(num).Slot = temp1(1)
                           strLR = temp1(2)
                  
                    ElseIf AIMEX = True Then    ''''(1161)针对正常的料，文档中的Slot如5-1-43，其中5表示第五个Machine， 1表示这个Machine的两个Side中的前面，43表示具体的位置
                           If Len(temp1(0)) = 1 Then
                              Data(num).machine = machine + "0" + temp1(0)
                           Else
                              Data(num).machine = machine + temp1(0)
                           End If
                           strSql = "select * from Machine where Machine='" & Data(num).machine & "' and Line='" & Line & "'and side='" & side & "'"
                           Set rs = Conn.Execute(strSql) '(1043)
                           If rs.EOF Then
                              MsgBox "Can't find The Machine:" & Data(num).machine & " in line:" & Line & " and side: " & side & " ,please define it in machinetype or check its format!", vbCritical, "ErrMessage"
                             Exit Sub
                           Else
                              TabQty = rs.Fields("Qty")
                              SlotQty = rs.Fields("MaxSlotNum")
                           End If
                           If TabQty <> 2 Or Int(temp1(2)) > SlotQty Then
                              MsgBox "The Slot " & StrSlot & " foramt is wrong,please check!", vbCritical, "ErrMessage"
                              Exit Sub
                           End If
                           Data(num).Slot = temp1(1) & "-" & temp1(2) '' ''''(1161) 故系统所取的Slot 为1-43
                           strLR = 0
                    
                    End If
                Case 3              ''''Add by Archer (20080624)
                    If BU = "NB4" Or BU = "NB7" Or BU = "ESBU" Then   '(1007)'1243
                        If temp1(2) > 100 And NXT = False Then   '(1143)
                           temp1(2) = Trim(Int(temp1(2)) - 100)
                        End If
                        If NXT = False Then
                           If (temp1(3) <> "1" And temp1(3) <> "2") Or Int(temp1(0)) > TabQty Or Int(temp1(2)) > SlotQty Then
                              MsgBox "The Slot " & StrSlot & " foramt is wrong,please check!", vbCritical, "ErrMessage"
                              Exit Sub
                           Else
                              strLR = temp1(3)
                              Data(num).Slot = temp1(0) + "-" + temp1(2)
                              Data(num).machine = machine
                           End If
                        Else
                           If NXT = True Then
                               If Len(temp1(0)) = 1 Then
                                  Data(num).machine = machine + "0" + temp1(0)
                               Else
                                  Data(num).machine = machine + temp1(0)
                               End If
                               strSql = "select * from Machine where Machine='" & Data(num).machine & "' and Line='" & Line & "'and side='" & side & "'"
                               Set rs = Conn.Execute(strSql) '(1043)
                               If rs.EOF Then
                                    MsgBox "Can't find The Machine:" & Data(num).machine & " in line:" & Line & " and side: " & side & " ,please define it in machinetype or check its format!", vbCritical, "ErrMessage"
                                   Exit Sub
                               Else
                                  TabQty = rs.Fields("Qty")
                                  SlotQty = rs.Fields("MaxSlotNum")
                               End If
                               If (temp1(3) <> "1" And temp1(3) <> "2") Or TabQty <> 1 Or Int(temp1(2)) > SlotQty Then
                                  MsgBox "The Slot " & StrSlot & " foramt is wrong,please check!", vbCritical, "ErrMessage"
                                  Exit Sub
                               End If
                               Data(num).Slot = temp1(2)
                               strLR = temp1(3)
                           End If
                        End If
                    End If
                Case 4
                    If AIMEX = True Then  '''针对AIMEX的Tray 料 ，其格式：5-2-A-2-2，其中5为Machine，2 为Machine的两个面的后面，A为层（每面分为2层，每层包含12列），2为每层的一列，最后一个2为LR
                        If Len(temp1(0)) = 1 Then
                            Data(num).machine = machine + "0" + temp1(0)
                        Else
                            Data(num).machine = machine + temp1(0)
                        End If
                        strSql = "select * from Machine where Machine='" & Data(num).machine & "' and Line='" & Line & "'and side='" & side & "'"
                        Set rs = Conn.Execute(strSql) '(1043)
                        If rs.EOF Then
                            MsgBox "Can't find The Machine:" & Data(num).machine & " in line:" & Line & " and side: " & side & " ,please define it in machinetype or check its format!", vbCritical, "ErrMessage"
                            Exit Sub
                        Else
                            TabQty = rs.Fields("Qty")
                            SlotQty = rs.Fields("MaxSlotNum")
                        End If
                        If TabQty <> 2 Then
                            MsgBox "The Slot " & StrSlot & " foramt is wrong,please check!", vbCritical, "ErrMessage"
                            Exit Sub
                        End If
                        If temp1(2) = "B" Then  '''如果当前为B层，为了防止与A层重复，故需要在原有的列上增加12。
                            temp1(3) = CInt(temp1(3)) + 12
                        End If
                        
                        Data(num).Slot = temp1(1) & "-" & temp1(3) ''' 故系统所取的Slot 为2-2，LR 为2
                        strLR = temp1(4)
                    End If
         
               Case Else
OtherBU:
                    MsgBox "The Slot " & StrSlot & " foramt is wrong,please check!", vbCritical, "ErrMessage"
                    Exit Sub
         End Select
        
        Data(num).LR = strLR
        Data(num).Qty = 1
        Data(num).Enabled = True
                       
        p1 = InStr(p1, s, "<Unit>")  'find next <Unit>
               
        If p1 = 0 Then Exit Do
    Loop
    
    Dim I, j As Integer
    
    For I = 1 To num
        For j = I + 1 To num
         
            If Data(j).Enabled = True Then
                If Data(I).compPN = Data(j).compPN And Data(I).Jobpn = Data(j).Jobpn And Data(I).Slot = Data(j).Slot And Data(I).machine = Data(j).machine Then
                    Data(I).Qty = Data(I).Qty + 1
                    Data(I).location = Data(I).location & ";" & Data(j).location
                    Data(j).Enabled = False
                End If
            End If
        Next

    Next
    
    
    Dim preJobPN As String, strMachine As String, preRev As String
    
    'Delete old data of the machine, jobPN, Rev
    For I = 1 To num
        If Data(I).Enabled = True Then
            If Data(I).Jobpn <> preJobPN Or Data(I).machine <> strMachine Then
                strSql = "delete from QSMS_MEBom where JobGroup='" & jobgroup & "' and Machine=" & sq(Data(I).machine) & " and JobPN=" & sq(Data(I).Jobpn) & _
                            " and Version=" & sq(Data(I).Rev) & " And BuildType = " & sq(BuildType) & " and Line=" & sq(Line) & " and Factory=" & sq(Factory) & ""    '(0007)
                Conn.Execute strSql
                
                strMachine = Data(I).machine
                preJobPN = Data(I).Jobpn
            End If
        End If
    Next I
    
    For I = 1 To num
        If Data(I).Enabled = True Then
            
            strSql = "insert into QSMS_MEBom(Machine,JobPN,JobGroup,Version,CompPN,LR,Slot,Qty,BuildType,Side,UID,Factory,Line,Location) values (" & _
                                sq(Data(I).machine) & "," & sq(Data(I).Jobpn) & "," & sq(jobgroup) & "," & sq(Data(I).Rev) & "," & _
                                sq(Data(I).compPN) & "," & sq(Data(I).LR) & "," & sq(Data(I).Slot) & "," & Data(I).Qty & "," & sq(BuildType) & "," & _
                                sq(side) & "," & sq(g_userName) & "," & sq(Factory) & "," & sq(Line) & "," & sq(";" & Data(I).location & ";") & ")" '(0007) '(1026)
            Conn.Execute strSql
        End If
    Next
strSql = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_FujiXML','" & Left(Trim(FileName), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSql)

End Sub

Private Sub CmdAdd_Click()
    Dim Pointer As Integer
    If File1.ListCount <= 0 Then Exit Sub
    If File1.ListIndex < 0 Then Exit Sub
    Pointer = File1.ListIndex
    ListFile.AddItem Trim(File1.FileName)
   ' lstWO_LIST.RemoveItem Pointer
    If File1.ListCount <> Pointer Then
       File1.ListIndex = Pointer
    End If
    

End Sub

Private Sub cmdADDALL_Click()
Dim I As Integer
    If File1.ListCount <= 0 Then Exit Sub
    
    For I = 0 To File1.ListCount - 1
  
      File1.ListIndex = I
      ListFile.AddItem Trim(File1.FileName)
   
   
  Next I
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
Dim I As Integer
Dim temp() As String
Dim str As String, Factory As String, Line As String, machine As String, side As String, MBPN As String, Version As String, jobgroup As String


If ListFile.ListCount <= 0 Then Exit Sub
    
For I = 0 To ListFile.ListCount - 1
    ListFile.ListIndex = I
    FileName = File1.Path & "\" & ListFile.Text
    txtFile = FileName
    
    'For first file, delete all machine bom by Jobgroup, Side ( except OTHERS (DIP))
    If I = 0 Then
        temp = Split(Trim(ListFile.Text), "-")
        side = Left(temp(6), Len(temp(6)) - 4)   '(1032)
        If UBound(temp) <> 6 Then
            MsgBox ("FileName format must be Factory-Line-Machine-PN-REV-BuildType-Side.xml!")   '(0007)'(1026)
            Exit Sub
        End If
       ' If CheckMachine(temp(1), temp(2), side) = False Then '(1032)   '(1043)marked by jocelyn.
       '     Exit Sub
       ' End If
        If temp(5) <> "1" And temp(5) <> "2" And temp(5) <> "3" And temp(5) <> "4" Then
           MsgBox ("BuildType must be 1,2,3 or 4.")
           Exit Sub
        End If
        
        If Trim(side) <> "S" And Trim(side) <> "C" And Trim(side) <> "Q" Then
           MsgBox ("Side must be S,C or Q.")
           Exit Sub
        End If
        
       ' If Trim(temp(5)) = "1" And Trim(Side) <> Mid(Trim(temp(2)), 2, 1) Then
       '    MsgBox ("The machine is " & temp(2) & ",side is " & Side & ",they are not match when buidtyp is 1.")
       '    Exit Sub
       ' End If
        
        If (Trim(temp(5)) = "2" Or Trim(temp(5)) = "5") And side <> "S" Then
           MsgBox ("The Side is " & side & ",BuildType is " & Trim(temp(5)) & ",they are not match,the Side must be S side.")
           Exit Sub
        End If
        
        If BU <> "ESBU" Then
            If Trim(temp(5)) = "3" And side <> "C" Then
                MsgBox ("The Side is " & side & ",BuildType is " & Trim(temp(5)) & ",they are not match,the Side must be C side.")
                Exit Sub
            End If
        Else
            If Trim(temp(5)) = "3" And (side <> "C" And side <> "Q") Then  ''1261
                MsgBox ("The Side is " & side & ",BuildType is " & Trim(temp(5)) & ",they are not match,the Side must be C side.")
                Exit Sub
            End If
        End If
        '******************************
        '****add by jeanson 2007/09/03
        strErrMessage = ""
        strErrMessage = FunPartNumberCheck(Trim(temp(3)))
        If strErrMessage <> "PASS" Then
            MsgBox strErrMessage
        Exit Sub
        End If
        '******************************
'        If Len(Trim(temp(1))) <> 11 Then
'           MsgBox ("The MBPN:" & Trim(temp(1)) & ",the length must be 11.Please check the MBPN!")
'           Exit Sub
'        End If
        
        If Len(Trim(temp(4))) <> 3 And Len(Trim(temp(4))) <> 2 Then
           MsgBox ("The Version:" & Trim(temp(4)) & ",the length must be 2 or 3.Please check the Version!")
           Exit Sub
        End If
        
        If MsgBox("Do you want to delete all machine bom by line, side (except OTHERS)?", vbYesNo) = vbYes Then
            Factory = Trim(temp(0)) '(0007)
            Line = Trim(temp(1))
            machine = Trim(temp(2))
            MBPN = Trim(temp(3))
            Version = Trim(temp(4))
            jobgroup = MBPN & "-" & Version
            
            str = "delete from QSMS_MEBOM where Jobgroup=" & sq(jobgroup) & " and side like " & sq(side) & " and Machine not like '%Other%' and Factory=" & sq(Factory) & "and Line=" & sq(Line) & ""
            Conn.Execute (str)
        End If
    End If
    
    
    If LoadDataFile(Trim(txtFile)) = False Then
      MsgBox ("Fail")
    Else
      MsgBox ("Finish") & I + 1
    End If
    DoEvents
Next I
'    If Trim(txtFile) = "" Or UCase(Right(Trim(txtFile), 3)) <> "XML" Then
'        MsgBox "You must select a XML file!!", vbInformation
'        Exit Sub
'    End If
'

End Sub

Private Sub cmdSelect_Click()
  CommonDialog1.ShowOpen
  txtFile = CommonDialog1.FileName
End Sub

Private Sub Dir1_Change()
 File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
'    Dim connStr As String
'    connStr = ReadIniFile("Connection", "connSMT", App.Path & "\set.ini")
'
'    If conn.State Then conn.Close
'    conn.Open connStr
End Sub

