Attribute VB_Name = "Pub"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function ConvCStringToVBString Lib "kernel32" Alias "lstrcpyA" (ByVal lpsz As String, ByVal pt As Long) As Long
'Public Declare Function MeasureE_GPIB Lib ".\MeasureE" (ByVal Equip As String, ByVal PAD As Long, ByVal QCategory As String) As Long
'Public Declare Function MeasureE_RS232 Lib ".\MeasureE" (ByVal Equip As String, ByVal PAD As Long, ByVal QCategory As String) As Long
Public Declare Function MeasureE_RS232 Lib ".\MeasureE" (ByVal Equip As String, ByVal PAD As Long, ByVal QCategory As String, ByVal Frequency As Single, ByVal Voltage As Single, ByVal Current As Single) As Long
Public Declare Function MeasureE_GPIB Lib ".\MeasureE" (ByVal Equip As String, ByVal PAD As Long, ByVal QCategory As String, ByVal Frequency As Single, ByVal Voltage As Single, ByVal Current As Single) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const CB_SETDROPPEDWIDTH = &H160

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public connStr As String
Public Conn As New ADODB.Connection
Public ConnSMT As New ADODB.Connection '1168
Public UID As String
Public RackID, Qty As String
Public strSQL As String
Public ScanCompPN As String
Public ScanMSD As String
Public NeedMSD As Boolean
Public CheckReturnForbiddenPN As String
Public strKeyInPNByManual As Boolean
Public MaintainFeederDID As String
Public DeleteMeBomByLine As Boolean  '1131
Public strChkDIDByLine As String  '1276
Public strChk_XL_WOPlanSeq As String  '1278
Public PreDIDPrinted As String
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public currProc As String
Public chkQty As String

Public WO(5)
Public Model(5)
Public PN(5)
Public Machine(5)
Public Slot(5)
Public MachineUnit(5)
Public Work_Order(5)
Public DIDType(5)
Public ISCYL(5)
Public SeqID(5)
Public VenderCode(5)
Public LR(5)

Public MachineCH(5)  ''1044
Public SideCH(5)   ''1044
Public LRCH(5)   ''1044
Public SlotCH(5)   ''1044
Public ReelWidth(5)

Public PrinterType As String ''1044
Public PrintDpm As String ''1044

Type Settings_DataType

     ConnectStr As String
     PRNa_Port As String
     PRNa_Settings As String
     LabelAFile As String
     LabelSATOFIle As String
     AutoDispatchLabel As String
     AutoDispatchSatoLabel As String
     ChkDIDDispatch As String
     UpdateJobSide As String
     
     '20071226 Denver for DID CallBack and Return print DID
     DIDLabelGood As String
     DIDLabelBad As String
     DIDLabelSATOGood As String
     DIDLabelSATOBad As String
     
     '20100618 Maggie CompPrint
     CompPrintLabel As String
     '20101014 Maggie CompPNLabelPrint   '(1013)
     CompPNLabelPrint As String
     KFLabel As String
     '''added by Jing 2008.04.05'''
     AutoDispatchNewLabel As String
     AutoDispatchSatoNewLabel As String
     DIDLabelPath As String  ''(1080)
     
End Type
Type Extra
    WO As String
    Machine As String
    Slot As String
    LR As String
    Line As String
    Group As String
    Qty As Long         ''Integer (1094)
    Side As String
End Type

Type PtData
    Machine As String
    Side As String
    DIDWOGROUP As String
    Line As String
    BU As String
    location As String  ''1242
    Mark As String '' 1255
    jobgroup As String ''1277
End Type

Type DIDBasic
     compPN As String
     DID As String
     VendorCode As String
     DateCode As String
     LotCode As String
     Qty As Long
     IsGood As String
     DIDType As String
     location As String
     Mark As String  '1255
     WareHouseID As String
     ''ReelWidth   As String
End Type

Public DIDInfo As DIDBasic
Public IP As String
Public oEncrypt As New Encrypt
Public Settings As Settings_DataType
Public WorkDir As String
Public Profile As String
Public g_userName As String, g_userRight() As String, g_delrightUser As String, g_factory As String ''(1016)
Public hSECTION As String
Public DeleteType As String
Public ProgramDescription As String
Public MEBomPath As String, MEBomBKPath As String, MEBomErrPath As String '''' for transfer pana AMI
Public CheckExpireFlag As String ' for transfer pana AMI
Public BomTest As String
Public BU As String
Public BUDIDShow As String
Public ChkFujiSPL As String

Public Check_NonAVL As String
Public Check_DID As String
Public Check_AVL As String
'Public pFlagTest As String
Public StrDeleteLog As String
Public DIDHead As String
Public strErrMessage As String
Public strFileContents As String
Public imagePath As String
Public TestFilepath As String
Public IPQCFlag As String
Public Finishflag As Boolean
Public TestDIDFlag As String
Public ValKeyCode As Long
Public F1 As String, F2 As String, F4 As String, F6 As String, F7 As String, QB As String, QC As String, Factory As String, CreateDIDFlag As String
Public PrtCallBKandReturn As String
Public CheckBomPilotWO As Boolean
Public CheckBomLogon As String
Public strAccessFlag As String
Public g_strUserRight As String
Public CheckBomID As String
Public CheckBomRight As Boolean

Public AutoDispatchForAnotherBU As String
Public CheckPNGroup As String
Public DIDnotToQWMS As String

Public IPQC_ChkVendorPN As String  ''Add flag(whether check Vendor_PN)
Public NewRs232 As New Rs232     ''Add LCR EquipType class
Public NewRs6420 As New Rs232
Public IC_CompChk As String  ''Add flag(whether check IC Component)
Public ChkOldDIDLabelQty As String  ''(0061)
Public ChkOneByOneMaterial As String  ''(0076)
Public NPMMachineType As String  ''(1079)
Public ChkWOGroupID As String     ''(1128)
Public ChkPrintDIDType As String
Public PrintedSeqID As String
Public BatchControl As String
Public chkDomain  As String '(1165)
Public UnChkCompPN As String '(1187)
Public CheckNeedMSD As String '(1188)
Public CheckWOIFReduceXboard As String '(1190)
Public CheckMSDCallBack As String '(1191)
Public CheckBurnDID As String
Public NoKeepPWD As String
Public BGAWarehouse As String
Public CheckBSMaterial As String ''(1213)
Public ChkEQProgram As String ''(1219)
Public ChkDateCode As String ''(1222)
Public ChkPNCQ As String
Public PrintedVenderCode As String   ''1223
Public NewGroupIDRule As String      ''1225
Public UnChkBaseReelQty As String      ''1225
Public ChkMEBOM_Location As String     ''1250
Public DIDAutoOpen As String    ''1268
Public LabelPrintCheck As String    ''1274






  Public Function GetStringFromPointer(ByVal lpString As Long) As String
  Dim NullCharPos     As Long
  Dim szBuffer     As String
    
          szBuffer = String(1024, 0)
          ConvCStringToVBString szBuffer, lpString
          '   Look   for   the   null   char   ending   the   C   string
          NullCharPos = InStr(szBuffer, vbNullChar)
          GetStringFromPointer = Left(szBuffer, NullCharPos - 1)
  End Function


Public Function ReadIniFile(ByVal strSection As String, ByVal strKey As String, strFname As String) As String
Dim strValue As String * 255
Dim intRet As Long

'On Error Resume Next
intRet = GetPrivateProfileString(strSection, strKey, "", strValue, Len(strValue), strFname)
ReadIniFile = Left(strValue, intRet)
End Function

Public Function BuildMainConnection()
Dim pwd As String
Dim intDBConnectTime As Integer
On Error GoTo Handler
intDBConnectTime = 0
oEncrypt.key = "Quanta"
connStr = ReadIniFile("database", "Connection", App.Path & "\set.ini")
IP = Mid(connStr, InStr(1, connStr, "Server=") + 7)
pwd = ReadIniFile("database", "PWD", App.Path & "\set.ini")
If pwd > "" Then
   connStr = connStr & ";pwd=" & oEncrypt.Decrypt(pwd)
End If
'ConnStr = "Provider=SQLOLEDB.1;Password=qms7sa;Persist Security Info=True;User ID=Sa;Initial Catalog=SMT;Data Source=172.26.60.4;Net Library=TCP/IP"
'Conn.CommandTimeout = 0
'Conn.CursorLocation = adUseClient
'If Conn.State = 1 Then Conn.Close

DBReConnect:
'Conn.Open connStr
Exit Function
'-------------update by Sandy---2007/10/04--
Handler:
    If (Err.Number = -2147467259 And Mid(Err.Description, 1, 49) = "[DBNETLIB][ConnectionOpen (Connect()).]SQL Server") Or InStr(1, Err.Description, "SSPI") > 0 Then
        intDBConnectTime = intDBConnectTime + 1
        If intDBConnectTime = 1 Then
            connStr = Replace(connStr, ";Network Library=DBMSSOCN", "")
            GoTo DBReConnect
        End If
    Else
        MsgBox Err.Description + vbCrLf + "Please call QMS staff for help"
        End
    End If
    Exit Function
End Function


Public Function GetDID(ByVal compPN As String, TransDate As String) As String
Dim str As String, TempDID As String
Dim RS As ADODB.Recordset
Dim I As Integer
Dim YMD As String
''modify new DID format:PN+DIDHead+YMD+Seqno(3)   Lynn 2009/02/24
'''''Get YMD,Y--Year(BASE34), M--month (BASE34), D--days(BASE34)'''''
YMD = Base_B_EquivOf_A(Mid(TransDate, 3, 2), 34, Apple34Chars) + Base_B_EquivOf_A(Mid(TransDate, 5, 2), 34, Apple34Chars) + Base_B_EquivOf_A(Mid(TransDate, 7, 2), 34, Apple34Chars)
'DIDHead = "NB3"
'Conn.Open ("Provider=sqloledb;UID=qmsuser;Server=172.26.60.5;database=QSMS;Network Library=DBMSSOCN;pwd=QuantacnQms")
'***** lynn: right(DID,4) BASE34 *****
If StrBU = "MBU" Then
str = "select Max(right(DID,3)) as maxSN from QSMS_DID where DID like '" & compPN & "-" & DIDHead & YMD & "%'"
Else
'str = "select Max(right(DID,3)) as maxSN from QSMS_DID where DID like '" & compPN & "-" & DIDHead & YMD & "%' and did not like '%-A%'"
str = "select Max(right(DID,3)) as maxSN from QSMS_DID where DID like '" & compPN & "-" & DIDHead & YMD & "%' AND SUBSTRING(DID,12,2)<>'-A'"   ''(1145)
End If
Set RS = Conn.Execute(str)

If IsNull(Trim(RS.Fields(0))) = True Then
    GetDID = compPN + "-" + DIDHead + YMD + "001"
    Exit Function
Else
    TempDID = ConvertBase2Dec(Trim(RS.Fields(0)), 34) + 1
    TempDID = Base_B_EquivOf_A(TempDID, 34, Apple34Chars)
'**Sandy      2007.11.26     update to drive out the blank space in DID.---------(0007)
   GetDID = Trim(compPN) + "-" + DIDHead + YMD + Right("000" & TempDID, 3) '' Lynn modify new DID format 2009/02/24
   GetDID = Replace(Trim(GetDID), " ", "")
End If
'***** END *****
End Function

Public Function ChkRackID(ByVal compPN As String) As Boolean
Dim str As String
Dim RS As ADODB.Recordset
str = "select RackID,Qty from QSMS_RackID where CompPN='" & compPN & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   RackID = Trim(RS!RackID)
   Qty = Trim(RS!Qty)
   ChkRackID = True
Else
   RackID = ""
   Qty = ""
   ChkRackID = False
End If

End Function

Public Function GetSettings(Profile As String, hSECTION As String) As Long
       Dim sSECTION As String
       Dim hVal As Long
       Dim hStr As String
       
       sSECTION = "COMMON"
       
       With Settings
            
            .PRNa_Port = UCase(Trim(GetProfileData(Profile, sSECTION, "PRNa_Port")))
            .PRNa_Settings = UCase(Trim(GetProfileData(Profile, sSECTION, "PRNa_Settings")))
            .LabelAFile = App.Path & "\" & UCase(Trim(GetProfileData(Profile, sSECTION, "LabelAFile")))
            .LabelSATOFIle = App.Path & "\" & UCase(Trim(GetProfileData(Profile, sSECTION, "LabelFIle_SATO")))
            .ChkDIDDispatch = UCase(Trim(GetProfileData(Profile, sSECTION, "CheckDIDDispatch")))
            .UpdateJobSide = "N"
            .UpdateJobSide = UCase(Trim(GetProfileData(Profile, sSECTION, "UpdateJobSide")))
            .AutoDispatchLabel = App.Path & "\" & UCase(Trim(GetProfileData(Profile, sSECTION, "AutoDispatchLabel")))
            .AutoDispatchSatoLabel = App.Path & "\" & UCase(Trim(GetProfileData(Profile, sSECTION, "AutoDispatchSatoLabel")))
            
            ''20071226 Denver for DID CallBack and Return print DID
            .DIDLabelGood = App.Path & "\" & UCase(Trim(GetProfileData(Profile, sSECTION, "DIDLabelGood")))
            .DIDLabelBad = App.Path & "\" & UCase(Trim(GetProfileData(Profile, sSECTION, "DIDLabelBad")))
            .DIDLabelSATOGood = App.Path & "\" & UCase(Trim(GetProfileData(Profile, sSECTION, "DIDLabelSATOGood")))
            .DIDLabelSATOBad = App.Path & "\" & UCase(Trim(GetProfileData(Profile, sSECTION, "DIDLabelSATOBad")))
            
            '20100618 Maggie CompPrint
            .CompPrintLabel = App.Path & "\" & UCase(Trim(GetProfileData(Profile, sSECTION, "CompPrintLabel")))
            '20101014 Maggie CompPNLabelPrint    '(1013)
            .CompPNLabelPrint = App.Path & "\" & UCase(Trim(GetProfileData(Profile, sSECTION, "CompPNLabelPrint")))
            .KFLabel = App.Path & "\" & UCase(Trim(GetProfileData(Profile, sSECTION, "KFLabel")))
            '''added by Jing 2008.04.05 (0032)'''
            .AutoDispatchNewLabel = App.Path & "\" & UCase(Trim(GetProfileData(Profile, sSECTION, "AutoDispatchNewLabel")))
            .AutoDispatchSatoNewLabel = App.Path & "\" & UCase(Trim(GetProfileData(Profile, sSECTION, "AutoDispatchSatoNewLabel")))
            .DIDLabelPath = App.Path & "\" & UCase(Trim(GetProfileData(Profile, sSECTION, "DIDLabelPath"))) ''(1080)
            
       End With
       'Settings.ConnectStr = GetScrConnect("Connect_MainDb")
       
End Function

Public Function GetProfileData(Profile As String, hSECTION As String, hKey As String) As String
       Dim hLen As Long
       Dim hString As String
       
       hString = String(255, 0)
       hLen = GetPrivateProfileString(hSECTION, hKey, vbNullString, hString, Len(hString), Profile)
       GetProfileData = StrConv(LeftB(StrConv(hString, vbFromUnicode), hLen), vbUnicode)
End Function



Public Function ReplaceStr(SearchString As String, key As String, Value As String) As String
       Dim StartPos As Long
       StartPos = InStr(SearchString, key)
       Select Case StartPos
          Case 0
            ReplaceStr = SearchString
          Case Else
            ReplaceStr = Mid(SearchString, 1, StartPos - 1) & Value & Mid(SearchString, StartPos + Len(key))
       End Select
End Function



Public Function SendSap1(ByVal WO As String, ByVal status As String, ByVal Factor As Double)
Dim str As String
Dim RS As ADODB.Recordset
Dim tempitem, TempUpCompPN As String
Dim TransDate As String
Dim I As Long
On Error GoTo errhandle:
str = "Exec QSMSSap1 '" & WO & "'"
Conn.Execute str
Exit Function
errhandle:
         MsgBox Err.Description
End Function
Public Function CloseWoByManual(ByVal WO As String, CloseType As String) As Boolean

Dim str As String
Dim RS As ADODB.Recordset
Dim TempGroupID As String
Dim Dispatch_Flag As String
Dim AOI_Flag As String
Dim SAP1_Flag As String
Dim SAP2_Flag As String
Dim XBoardQtyInput As String   ''(1190)
Dim XBoardQty As Integer   ''(1190)

On Error GoTo errhandle:
CloseWoByManual = True

'(1)Check if the WO can be closed
If FrmCloseWO.Check1.Value = 0 Then
    Dispatch_Flag = "N"
Else
    Dispatch_Flag = "Y"
End If

If FrmCloseWO.Check2.Value = 0 Then
    AOI_Flag = "N"
Else
    AOI_Flag = "Y"
End If

If FrmCloseWO.Check3.Value = 0 Then
    SAP1_Flag = "N"
Else
    SAP1_Flag = "Y"
End If

If FrmCloseWO.Check4.Value = 0 Then
    SAP2_Flag = "N"
Else
    SAP2_Flag = "Y"
End If
str = "exec QSMSChkCloseWOByManual '" & WO & "','" & Dispatch_Flag & "','" & AOI_Flag & "','" & SAP1_Flag & "','" & SAP2_Flag & "','" & g_userName & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   If UCase(RS.Fields(0)) = "PASS" Then
   Else
       CloseWoByManual = False
       Call CopyToExcel(RS)
       MsgBox UCase(RS.Fields(0))
       Exit Function
   End If
End If

'(2) Check If have any DID need to return
str = "QSMS_WONeedReturnDID '" & WO & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
    ''(1153)
    Dim strResultNB5 As VbMsgBoxResult
    If BU = "NB5" Or BU = "NB3" Then   ''20180605 PU3导入 Seven
        strResultNB5 = MsgBox("There are some DID need to return by the WO,please check!!" & vbCrLf & "1.[Yes]close WO and delete DID;" & vbCrLf & "2.[No]close WO and not delete DID;" & vbCrLf & "3.[Cancel]Do not close WO", vbYesNoCancel, "Message!")
        If strResultNB5 = vbYes Then
            str = "exec QSMSCloseWODelDID '" & WO & "','" & g_userName & "' "
            Conn.Execute (str)
        ElseIf strResultNB5 = vbCancel Then
            CloseWoByManual = False
            Call CopyToExcel(RS)
            Exit Function
        End If
    Else
        If MsgBox("There are some DID need to return by the WO,please check the result!!Do you still want to Close the WO(If you close the WO, the DID will be delete which need to return!!)?", vbYesNo, "Message!") = vbYes Then
            str = "exec QSMSCloseWODelDID '" & WO & "','" & g_userName & "' "
            Conn.Execute (str)
        Else
            CloseWoByManual = False
            Call CopyToExcel(RS)
            Exit Function
        End If
    End If
End If

''MBU Xborad自动收缩C和S面材料使用量，发料量，需求量''(1190)
If CheckWOIFReduceXboard = "Y" Then
    str = "exec QSMS_CloseWO_CheckWOIFReduceXboard '" & WO & "'"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
       If UCase(RS.Fields(0)) = "1" Then
            XBoardQtyInput = InputBox("请输入" & WO & "的X板数量：", "CloseWO", 0)
            XBoardQty = CInt(Trim(XBoardQtyInput))
            str = "exec QSMS_CloseWO_ReduceXboard '" & WO & "'," & XBoardQty & ",'" & g_userName & "'"
            Set RS = Conn.Execute(str)
            If Not RS.EOF Then
               If UCase(RS.Fields(0)) = "PASS" Then
               Else
                   CloseWoByManual = False
                   Call CopyToExcel(RS)
                   Exit Function
               End If
            End If

       End If
    End If
End If


'(3) send SAP1 data ,include lost data and sended more data
str = "exec QSMS_SapCostPacking '" & WO & "','" & g_userName & "','" & CloseType & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   If UCase(RS.Fields(0)) = "PASS" Then
   Else
       CloseWoByManual = False
       Call CopyToExcel(RS)
       Exit Function
   End If
End If
Exit Function
errhandle:
         CloseWoByManual = False
         MsgBox Err.Description & " Please Call QMS "

End Function


Public Function ChkWoFinished(ByVal WO As String) As Boolean
Dim str As String
Dim TransDate As String
Dim RS As ADODB.Recordset

str = "select WofinishedFlag from QSMS_WO where Work_Order= '" & WO & "' "

Set RS = Conn.Execute(str)

If RS.EOF Then
   ChkWoFinished = False
Else
   If UCase((RS!wofinishedflag)) = "Y" Then
       str = "select distinct WofinishedFlag from QSMS_WO where Work_Order in (select wo from sap_wo_list where [group] in (select [group] from sap_wo_list where wo='" & WO & "')) and WofinishedFlag='N'"
       Set RS = Conn.Execute(str)
       If RS.EOF Then
          ChkWoFinished = True
       Else
          ChkWoFinished = False
       End If
   Else
       ChkWoFinished = False
   End If
End If


End Function
Public Function GetGroupID(ByVal WO As String) As String
Dim str As String
Dim TransDate As String
Dim RS As ADODB.Recordset
GetGroupID = ""
str = "select  GroupID from QSMS_WoGroup where Work_Order='" & WO & "'"
Set RS = Conn.Execute(str)
If RS.EOF Then
   GetGroupID = ""
Else
   GetGroupID = Trim(RS!GroupID)
End If

End Function
Public Function ChkGroupFinished(ByVal WO As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim TempGroupID As String
TempGroupID = GetGroupID(WO)
If TempGroupID = "" Then
   Exit Function
End If
str = "select DispatchFlag from QSMS_WoGroup where GroupID='" & TempGroupID & "' and DispatchFlag<>'Y'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   Exit Function
End If



str = "Exec QSMSGroupCompQty '" & TempGroupID & "'"
Conn.Execute str
MsgBox "The group finished the dispatching"
End Function
Public Function ChkPrgVer(Program As String, Version As String) As Boolean
    Dim str As String
    Dim RS As ADODB.Recordset
    
    ChkPrgVer = False
    str = "Select Version,Description From Program_Version Where Program ='" & Trim(Program) & "'"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
        If Trim(Version) >= Trim(RS!Version) Then
            str = "update Program_Version set Version='" & Version & "' where Program='" & Trim(Program) & "'"
            Conn.Execute str
            ChkPrgVer = True
            ProgramDescription = Trim(Program) & " " & Trim(Version) & ":" & Trim(RS!Description)
        Else
            ChkPrgVer = False
        End If
        
    End If
End Function

Public Function ChkNonAVL(ByVal DID As String, ByVal Customer As String, ByVal Model As String, ByVal MBPN As String, ByVal Work_Order As String) As Boolean
Dim str As String
Dim RS As ADODB.Recordset
Dim compPN, VendorCode, DateCode, LotCode As String
ChkNonAVL = True
str = "select CompPN,VendorCode,DateCode,LotCode from QSMS_DID where DID='" & Trim(DID) & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   compPN = Trim(RS!compPN)
   VendorCode = Trim(RS!VendorCode)
   DateCode = Trim(RS!DateCode)
   LotCode = Trim(RS!LotCode)
   
Else
   MsgBox "Can not find the DID,Please check"
   ChkNonAVL = False
   Exit Function
End If

str = "select VendorCode,DateCode,LotCode from QSMS_NonAVL where Customer='" & Trim(Customer) & "' and CompPN='" & Trim(compPN) & "' " & _
      " and '" & MBPN & "' like  rtrim(Model)+'%' and (vendorcode='" & VendorCode & "' or vendorcode='')" & _
      " and (datecode='" & DateCode & "' or datecode='') and (LotCode='" & LotCode & "' or lotcode='') " & _
      " and (work_Order='" & Work_Order & "' or work_order='')"

Set RS = Conn.Execute(str)
If RS.EOF Then
   ChkNonAVL = True
   Exit Function
Else
   ChkNonAVL = False
End If

'While Not Rs.EOF
'      If UCase(Trim(Rs!VendorCode)) = UCase(VendorCode) Or UCase(Trim(Rs!DateCode)) = UCase(DateCode) Or UCase(LotCode) = UCase(Trim(Rs!LotCode)) Then
'
'         ChkNonAVL = False
'      End If
'      Rs.MoveNext
'Wend

If Check_NonAVL <> "Y" Then
    ChkNonAVL = True 'mark by leimo temporary   20061201   '--add flag [Check_NonAVL] for NB3 by Lynn 2007/06/17
End If

If ChkNonAVL = False Then
   MsgBox "Check NonAVL failed"
End If
End Function
Public Function ChkAVL(ByVal compPN As String, ByVal VendorCode As String, ByVal Customer As String, ByVal Model As String) As Boolean

Dim str As String
Dim RS As ADODB.Recordset
Dim AVLCustomer, TempModel, ModelFlag As String
Dim ControlPart As Boolean
ChkAVL = True
'(1) get if AVL by Quanta of by customer
str = "Select AVL_Customer,ModelFlag from AVL_Vendor where Customer='" & Customer & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   AVLCustomer = Trim(RS!avl_customer)
   ModelFlag = Trim(RS!ModelFlag)
Else
   AVLCustomer = "Quanta"   'default use quanta AVL
End If

'(2) Check AVL
If UCase(AVLCustomer) = "QUANTA" Then
Else
    If UCase(ModelFlag) = "Y" Then
        str = "Select * from QSMS_AVL where Customer='" & AVLCustomer & "' and model='" & Model & "' and CompPN='" & compPN & "' and Vendorcode='" & VendorCode & "'"
    Else
        str = "Select * from QSMS_AVL where Customer='" & AVLCustomer & "' and model='' and CompPN='" & compPN & "' and Vendorcode='" & VendorCode & "'"
    End If
    Set RS = Conn.Execute(str)
    If RS.EOF Then
       ChkAVL = False
       'Exit Function
       'MsgBox "Check AVL failed,please check "
    End If
End If

'(3) check control parts ,currently only  for ES and AS
str = "Select VendorCode from QSMS_ControlPart Where Model='" & Model & "' and CompPN='" & compPN & "'"
Set RS = Conn.Execute(str)
If RS.EOF Then
   ChkAVL = True
   Exit Function
End If
ControlPart = False
While Not RS.EOF
      If UCase(Trim(RS!VendorCode)) = UCase(Trim(VendorCode)) Then
         ControlPart = True
      End If
      RS.MoveNext
Wend
If ControlPart = True Then
   ChkAVL = True
Else
   ChkAVL = False
    'MsgBox "Check contorl parts failed,please check "
End If
ChkAVL = True

End Function
Public Function GetAVLCustomer(ByVal Customer As String)
Dim str As String
Dim RS As ADODB.Recordset
str = "Select AVL_Customer from AVL_Vendor where Customer='" & Customer & "'"
Set RS = Conn.Execute(str)
If Not RS.EOF Then
   
End If
End Function
Public Function ChkQSMS_WO(ByVal WO As String) As Boolean
Dim str As String
Dim RS As ADODB.Recordset
ChkQSMS_WO = True
str = "Select WO from sap_wo_list where WO='" & WO & "' And status >= 10 "
Set RS = Conn.Execute(str)

If RS.EOF Then
   ChkQSMS_WO = False
End If

End Function

Public Function ChkMBWo(ByVal WO As String) As Boolean
Dim str As String
Dim RS As ADODB.Recordset
ChkMBWo = False

'如果 InitAOIFlag='Y'则该工单为大板工单
str = "select WO from Sap_Wo_list where wo='" & WO & "' and InitAOIFlag='Y'"
Set RS = Conn.Execute(str)
If RS.EOF Then
    ChkMBWo = False
Else
   ChkMBWo = True
End If
End Function

Public Function GetNotInheritDIDByWO(ByVal WO As String) As ADODB.Recordset
Dim str As String
Dim RS As ADODB.Recordset
'Str = "Select distinct a.* from QSMS_DID a,QSMS_Dispatch B where a.InheritFlag='N' and a.RemainQty<>0  and a.DID=b.DID and "
str = "Select distinct a.* from QSMS_DID a,QSMS_Dispatch B where a.InheritFlag='N' and a.RemainQty<>0  and a.DID=b.DID and a.TransDateTime=b.DIDDateTime and " & _
    "b.work_order in (select wo from sap_wo_list where [Group] in (select [group] from sap_wo_list where wo='" & WO & "'))"
Set RS = Conn.Execute(str)
Set GetNotInheritDIDByWO = RS


End Function


Public Function Delay_Time(ByVal DelaySec As Long) As Long
    Dim ExitTime As String
    ExitTime = DateAdd("s", DelaySec, Now)
    Do
      Select Case DateDiff("s", ExitTime, Now)
         Case Is < 0
         Case 0
         Case Is > 0
           Exit Do
      End Select
      DoEvents
    Loop
End Function

Public Function ChkWOItemFinished(ByVal WO As String) As Boolean
Dim str As String
Dim RS As ADODB.Recordset
Dim WoArray(100) As String
Dim tempwo As String
Dim I As Integer
Dim TransDate As String

str = "select getdate()"
Set RS = Conn.Execute(str)
TransDate = Format(RS.Fields(0), "YYYYMMDDHHNNSS")

For I = 1 To 100
    WoArray(I) = ""
Next I
I = 1

tempwo = Replace(Mid(WO, 3, Len(WO) - 3), "'", "")
While Len(tempwo) >= 9
      WoArray(I) = Mid(tempwo, 1, 9)
      If InStr(tempwo, ",") > 0 Then
         tempwo = Mid(tempwo, 11)
      Else
         tempwo = ""
      End If
      I = I + 1
Wend

I = 1
ChkWOItemFinished = True
While WoArray(I) <> ""
    str = "Select distinct work_order from QSMS_Wo where Work_Order = '" & WoArray(I) & "' and DispatchQty=0"
    Set RS = Conn.Execute(str)
    If Not RS.EOF Then
        ChkWOItemFinished = False
        str = "Update QSMS_WOGroup set DispatchFlag='N' Where Work_Order='" & WoArray(I) & "'"
        Conn.Execute str
    Else
        ChkWOItemFinished = True
        str = "Update QSMS_WOGroup set DispatchFlag='1' Where Work_Order='" & WoArray(I) & "'"
        Conn.Execute str
        'add by Giant for WO dispatch ok time --20070624
        str = "Update Sap_WO_List set DispatchOKDateTime='" & TransDate & "' where WO='" & WoArray(I) & "'"
        Conn.Execute str
    End If
    I = I + 1
Wend
End Function

Public Function ChkDIDBelongToGroup(ByVal DID As String, ByVal GroupID As String) As Boolean
Dim str As String
Dim RS As ADODB.Recordset
ChkDIDBelongToGroup = True
str = "Select distinct GroupID from QSMS_Dispatch where DID='" & DID & "' and DeletedFlag<>'Y'"
Set RS = Conn.Execute(str)
While Not RS.EOF
      If Trim(RS!GroupID) <> "" And UCase(Trim(RS!GroupID)) <> UCase(GroupID) Then
          ChkDIDBelongToGroup = False
          MsgBox "The DID has been dispatched to the GroupID : " & RS!GroupID & ":   and didn't return ,Please check"
          Exit Function
      End If
      RS.MoveNext
Wend
End Function

Public Function InsertIntoQSMSLog(ByVal AppName As String, ByVal SubFunction As String, ByVal Desc1 As String)
Dim str As String
Dim RS As ADODB.Recordset
Dim transdatetime As String
On Error GoTo ErrHdl:
str = "Select getdate()"
Set RS = Conn.Execute(str)
transdatetime = Format(RS(0), "YYYYMMDDHHMMSS")

str = "insert into QSMS_Error_LOG(AppName,SubFunction,SubID,DetailDesc,Col2,transdateTime) values ('" & AppName & "','" & SubFunction & "',Newid(),'" & Trim(Desc1) & "','" & g_userName & "','" & transdatetime & "')"
Conn.Execute (str)
Exit Function
ErrHdl:
    Exit Function
End Function
Public Function IsInteger(StrINT As String) As Boolean
Dim Integerregexp As RegExp   ''建立变量
Dim IntegerMatches As MatchCollection
Dim IntegerMatch As Match
IsInteger = False
Set Integerregexp = New RegExp  ' 建立正则表达式。
Integerregexp.IgnoreCase = True ' 设置是否区分大小写
Integerregexp.Global = True     ' 搜索全部匹配
'Modelregexp.Pattern = "(\w)+[@](\w)+[.](\w)+"  '设置模式
Integerregexp.Pattern = "^(([1-9])|([1-9][0-9]+))$" '设置模式^(([1-9])|([0-9]+))$
Set IntegerMatches = Integerregexp.Execute(StrINT)  ' 执行搜索
For Each IntegerMatch In IntegerMatches
    If StrINT <> IntegerMatch.Value Then
        IsInteger = False
        Exit Function
    Else
        IsInteger = True
    End If
Next



End Function
