VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmTransferPanaMSF 
   Caption         =   "TransferPanaMSF[2010-12-21]"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   11460
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1035
      Width           =   11460
      _ExtentX        =   20214
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5821
            MinWidth        =   5821
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5821
            MinWidth        =   5821
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2640
      Top             =   840
   End
   Begin VB.CommandButton cmdGetMEBom 
      Caption         =   "Upload"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9960
      TabIndex        =   2
      Top             =   360
      Width           =   1395
   End
   Begin VB.TextBox txtFile 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   8535
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   375
      Left            =   8760
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmTransferPanaMSF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type BomData
    Slot As String
    CompPN As String
    Jobpn As String
    Revision As String
End Type

Private Type BomQty
    Slot As String
    CompPN As String
    Jobpn As String
    Revision As String
    Qty As Integer
End Type

Dim i As Integer

Private Sub cmdGetMEBom_Click()
    If Trim(txtFile) = "" Then
        MsgBox "You must select a file!!", vbInformation
        Exit Sub
    End If
        
    If LoadDataFile(Trim(txtFile)) = False Then
        MsgBox ("Fail")
        Exit Sub
    End If
    
    StatusBar1.Panels(1) = txtFile & " OK"
    StatusBar1.Panels(3) = "Finished DateTime:" & Now
    i = 0
End Sub

Private Sub cmdSelect_Click()
  CommonDialog1.ShowOpen
  txtFile = CommonDialog1.FileName
End Sub

Private Function LoadDataFile(strFile As String) As Boolean
Dim Factory As String, machine As String, Jobpn As String, Revision As String, StrSlot As String, BuildType As String, Side As String, strCompPN As String, FNFile As Integer, strJobPN As String, line As String
Dim Arry() As String, Arry2() As String, temp() As String, BomFile As String, strCurrent As String, MSFData() As BomData, MSFBomQty() As BomQty
Dim strSQL As String, j As Integer, k As Integer, t As Integer, CheckS As Boolean
Dim JobGroup As String, m As Integer, rs As ADODB.Recordset

On Err GoTo Herr:

LoadDataFile = False
Arry = Split(strFile, "\")
BomFile = Arry(UBound(Arry))

temp = Split(Trim(BomFile), "-")
If UBound(temp) <> 6 Then
    MsgBox ("Filename format must be Factory-Line-Machine-PN-Rev-BuildType-Side !")  '(0007)  (1025)
    Exit Function
End If
If Trim(temp(5)) <> "1" And Trim(temp(5)) <> "2" And Trim(temp(5)) <> "3" Then
   MsgBox ("BuildType must be 1,2 or 3.")
   Exit Function
End If
If Left(Trim(temp(6)), 1) <> "S" And Left(Trim(temp(6)), 1) <> "C" And Left(Trim(temp(6)), 1) <> "Q" Then
   MsgBox ("Side must be S,C or Q.")
   Exit Function
End If
'******************************
'****add by jeanson 2007/09/03
strErrMessage = ""
strErrMessage = FunPartNumberCheck(Trim(temp(3)))
If strErrMessage <> "PASS" Then
    MsgBox strErrMessage
Exit Function
End If
'******************************
'If Len(Trim(temp(1))) <> 11 Then
'   MsgBox ("The JobPN:" & Trim(temp(1)) & ",length must be 11,Please check the JobPN!")
'   Exit Function
'End If
If Len(Trim(temp(4))) <> 3 And Len(Trim(temp(4))) <> 2 Then
   MsgBox ("The Version:" & Trim(temp(3)) & ",length must be 2 or 3,Please check the Version!")
   Exit Function
End If

Factory = Trim(temp(0)) 'add by giant 2008/06/27 (0007)
line = Trim(temp(1))
machine = Trim(temp(2))
Jobpn = Trim(temp(3))
Revision = Trim(temp(4))
BuildType = Trim(temp(5))
Side = Left(Trim(temp(6)), 1)
If CheckMachine(line, machine, Side) = False Then '(1032)
     Exit Function
End If


'If BuildType = "1" And Mid(Trim(machine), 2, 1) <> Trim(Side) Then
 '  MsgBox ("The machine is " & Trim(machine) & ",side is " & Trim(Side) & ",they are not match when buidtyp is 1.")
 '  Exit Function
'End If

If BuildType = "2" And Side <> "S" Then
   MsgBox ("The side is " & Side & ",BuildType is 2,they are not match,side must be S side.")
   Exit Function
End If

If BuildType = "3" And Side <> "C" Then
   MsgBox ("The side is " & Side & ",BuildType is 3,they are not match,side must be C side.")
   Exit Function
End If

'If InStr(BomFile, "-") = 0 Then
'    MsgBox ("Filename format must be Machine-PN-Rev !")
'    Exit Function
'End If
'm = InStr(BomFile, ".")
'Machine = Left(BomFile, InStr(BomFile, "-") - 1)
'JobPN = Left(Right(BomFile, Len(BomFile) - InStr(BomFile, "-")), InStr(Right(BomFile, Len(BomFile) - InStr(BomFile, "-")), "-") - 1)
'If m = 0 Then
'   Revision = Trim(Mid(BomFile, Len(Machine) + Len(JobPN) + 3))
'Else
'   Revision = Trim(Mid(BomFile, Len(Machine) + Len(JobPN) + 3, m - Len(Machine) - Len(JobPN) - 3))
'End If
JobGroup = Trim(Jobpn) + "-" + Trim(Revision)
    
If Trim(Factory) = "" Or Trim(machine) = "" Or Trim(Jobpn) = "" Or Trim(Revision) = "" Or Trim(line) = "" Then '(0007) (1025)
    MsgBox ("Filename format must be Factory-line-Machine-PN-Rev !")
    Exit Function
End If

FNFile = FreeFile
Open strFile For Input As #FNFile
StatusBar1.Panels(1) = "GetMEBom_" & BomFile
StatusBar1.Panels(2) = "Start DateTime:" & Now
While Not EOF(FNFile)
      i = i + 1
      Line Input #FNFile, strCurrent
      strCurrent = Trim(Replace(Replace(strCurrent, vbCrLf, ""), Chr(9), " "))
      If i > 1 Then
         Arry = Split(strCurrent, ",")
         StrSlot = Trim(Arry(1))
         strCompPN = Trim(Arry(2))
         strJobPN = Trim(Arry(6))
         If StrSlot <> "" And strCompPN <> "" Then
            j = j + 1
            ReDim Preserve MSFData(j)
            MSFData(j).Slot = StrSlot
            MSFData(j).CompPN = strCompPN
            
            Arry2 = Split(strJobPN, "-")
            If UBound(Arry2) > 0 Then
               If UBound(Arry2) <> 2 Then
                  MsgBox "BoardType wrong : " & strJobPN & ", must be PN-REV!"
                  Exit Function
               Else
                    '******************************
                    '****add by jeanson 2007/09/03
                    strErrMessage = ""
                    strErrMessage = FunPartNumberCheck(Trim(Arry2(1)))
                    If strErrMessage <> "PASS" Then
                        MsgBox strErrMessage
                  
                    Exit Function
                    End If
                    '******************************
'                  If Len(Trim(Arry2(1))) <> 11 Then
'                     MsgBox ("The JobPN:" & Trim(Arry2(1)) & ",length must be 11,Please check the JobPN!")
'                     Exit Function
'                  End If
                  If Len(Trim(Arry2(2))) <> 3 And Len(Trim(Arry2(2))) <> 2 Then
                     MsgBox ("The Version:" & Trim(Arry2(2)) & ",length must be 2 or 3,Please check the Version!")
                     Exit Function
                  End If
                  MSFData(j).Jobpn = Trim(Arry2(1))
                  MSFData(j).Revision = Trim(Arry2(2))
               End If
            Else
              MSFData(j).Jobpn = Trim(Jobpn)
              MSFData(j).Revision = Trim(Revision)
            End If
         End If
      End If
Wend
Close #FNFile

For i = 1 To UBound(MSFData)
    If i = 1 Then
       k = k + 1
       ReDim Preserve MSFBomQty(k)
       MSFBomQty(k).Slot = MSFData(i).Slot
       MSFBomQty(k).CompPN = MSFData(i).CompPN
       MSFBomQty(k).Jobpn = MSFData(i).Jobpn
       MSFBomQty(k).Revision = MSFData(i).Revision
       MSFBomQty(k).Qty = 1
    Else
       For t = 1 To UBound(MSFBomQty)
          If MSFBomQty(t).CompPN = MSFData(i).CompPN And MSFBomQty(t).Slot = MSFData(i).Slot And MSFBomQty(t).Jobpn = MSFData(i).Jobpn And MSFBomQty(t).Revision = MSFData(i).Revision Then
             MSFBomQty(t).Qty = MSFBomQty(t).Qty + 1
             CheckS = True
          End If
          If t = UBound(MSFBomQty) And CheckS = False Then
             k = k + 1
             ReDim Preserve MSFBomQty(k)
             MSFBomQty(k).Slot = MSFData(i).Slot
             MSFBomQty(k).CompPN = MSFData(i).CompPN
             MSFBomQty(k).Jobpn = MSFData(i).Jobpn
             MSFBomQty(k).Revision = MSFData(i).Revision
             MSFBomQty(k).Qty = 1
          End If
       Next t
    End If
    CheckS = False
Next i

For i = 1 To UBound(MSFBomQty)
    'Auto Get TraySlot
    'Marked by Lynn 2008/07/22 We let user upload trayslot instead of auto
'    If Val(MSFBomQty(i).Slot) > 200 Then
'       strSQL = "select * from TraySlot where Machine=" & sq(Machine) & " and Slot=" & sq(MSFBomQty(i).Slot) & ""
'       Set rs = Conn.Execute(strSQL)
'       If rs.EOF Then
'          strSQL = "Insert into TraySlot(Machine,Slot,UID) values(" & sq(Machine) & "," & sq(MSFBomQty(i).Slot) & ",'Auto')"
'          Conn.Execute (strSQL)
'       End If
'    End If
    
    '20101221 Maggie ÐÞ¸ÄBug,Ìí¼Ó¡¯    (1030)
    'strSql = "delete QSMS_MEBom where JobGroup='" & jobgroup & "' and Machine='" & Trim(machine) & "' and JobPN='" & MSFBomQty(i).Jobpn & "' and Version='" & MSFBomQty(i).Revision & "' and BuildType='" & BuildType & "' and Factory='" & Trim(Factory) & "and Line='" & Trim(Line) & "'" '(0007)(1025)
    strSQL = "delete QSMS_MEBom where JobGroup='" & JobGroup & "' and Machine='" & Trim(machine) & "' and JobPN='" & MSFBomQty(i).Jobpn & "' and Version='" & MSFBomQty(i).Revision & "' and BuildType='" & BuildType & "' and Factory='" & Trim(Factory) & "' and Line='" & Trim(line) & "'" '(0007)(1025)
    Conn.Execute (strSQL)
Next i
'(0007)
For i = 1 To UBound(MSFBomQty)
    strSQL = "Insert Into QSMS_MEBom(Machine,JobPN,JobGroup,Version,CompPN,LR,Slot,Qty,BuildType,Side,UID,Factory,Line) values('" & Trim(machine) & "','" & MSFBomQty(i).Jobpn & "','" & JobGroup & "','" & MSFBomQty(i).Revision & "','" & MSFBomQty(i).CompPN & "',0,'" & Trim(MSFBomQty(i).Slot) & "','" & MSFBomQty(i).Qty & "','" & Trim(BuildType) & "','" & Trim(Side) & "','" & Trim(g_userName) & "','" & Trim(Factory) & "','" & Trim(line) & "')"     '(1025)
    Conn.Execute (strSQL)
Next i

strSQL = "Insert into QSMS_Log(System_Name,Event_No,DID,User_Name,ReturnQty,Trans_Date) values('SMT_QSMS','Load_PanaMSF','" & Left(Trim(BomFile), 50) & "','" & Trim(g_userName) & "',0,[DBO].[FormatDate](getdate(), 'YYYYMMDDHHNNSS'))"
Conn.Execute (strSQL)


LoadDataFile = True
 MsgBox "*** Load  finish ! ***" & "   " & vbCrLf & _
               "Total Counter : " & UBound(MSFBomQty) & vbCrLf
Exit Function

Herr:
  MsgBox Err.Description, vbCritical, "ErrMessage!"
  LoadDataFile = False
End Function


Private Sub Form_Initialize()
  If App.PrevInstance Then
    MsgBox "The program has been running in this machine, this instance will close !"
    End
  End If
End Sub

