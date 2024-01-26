Attribute VB_Name = "SubMain"

Option Explicit
Public cn As New ADODB.Connection
Public CNHistory As New ADODB.Connection

Public strLine As String, strStation As String, StrConn As String, strRights As String
Public connStr As String
Public SMTServer As String
Public QSMSServer As String
Public SMTDB As String
Public QSMSDB As String
Public sql As String
Public Rs As New ADODB.Recordset

Private Sub Main()
If App.Title <> App.EXEName Then
      If Command = "" Then
            MsgBox "Please Use MainMenu "
            End
        Else
            If InStr(1, Command, "<LINE=", vbTextCompare) > 0 Then
              strLine = Mid(Mid(Command, InStr(1, Command, "<LINE=", vbTextCompare) + Len("<LINE="), Len(Command)), 1, InStr(1, Mid(Command, InStr(1, Command, "<LINE=", vbTextCompare) + Len("<LINE="), Len(Command)), ">") - 1)
            End If
            If InStr(1, Command, "<STATION=", vbTextCompare) > 0 Then
              strStation = Mid(Mid(Command, InStr(1, Command, "<Station=", vbTextCompare) + Len("<Station="), Len(Command)), 1, InStr(1, Mid(Command, InStr(1, Command, "<STATION=", vbTextCompare) + Len("<STATION="), Len(Command)), ">") - 1)
            End If
            If InStr(1, Command, "<CONN=", vbTextCompare) > 0 Then
              StrConn = Mid(Mid(Command, InStr(1, Command, "<CONN=", vbTextCompare) + Len("<CONN="), Len(Command)), 1, InStr(1, Mid(Command, InStr(1, Command, "<CONN=", vbTextCompare) + Len("<CONN="), Len(Command)), ">") - 1)
            End If
        End If
        '0014
        If InStr(1, Command, "<RIGHT=", vbTextCompare) > 0 Then
            strRights = Mid(Mid(Command, InStr(1, Command, "<RIGHT=", vbTextCompare) + Len("<RIGHT="), Len(Command)), 1, InStr(1, Mid(Command, InStr(1, Command, "<RIGHT=", vbTextCompare) + Len("<RIGHT="), Len(Command)), ">") - 1)
        End If
        If InStr(1, Command, "<USERID=", vbTextCompare) > 0 Then
            g_userName = Mid(Mid(Command, InStr(1, Command, "<USERID=", vbTextCompare) + Len("<USERID="), Len(Command)), 1, InStr(1, Mid(Command, InStr(1, Command, "<USERID=", vbTextCompare) + Len("<USERID="), Len(Command)), ">") - 1)
        End If
        If InStr(1, Command, "<FACTORY=", vbTextCompare) > 0 Then
            g_factory = Mid(Mid(Command, InStr(1, Command, "<FACTORY=", vbTextCompare) + Len("<FACTORY="), Len(Command)), 1, InStr(1, Mid(Command, InStr(1, Command, "<FACTORY=", vbTextCompare) + Len("<FACTORY="), Len(Command)), ">") - 1)
        End If
        If InStr(1, Command, "<CHKDOMAIN=", vbTextCompare) > 0 Then ''1165
            chkDomain = Mid(Mid(Command, InStr(1, Command, "<CHKDOMAIN=", vbTextCompare) + Len("<CHKDOMAIN="), Len(Command)), 1, InStr(1, Mid(Command, InStr(1, Command, "<CHKDOMAIN=", vbTextCompare) + Len("<CHKDOMAIN="), Len(Command)), ">") - 1)
        End If
Else
    'End
    strLine = "M17"
    strStation = "QSMS"
'    StrConn = "PROVIDER=SQLOLEDB;SERVER=172.26.16.4;UID=qms;PWD=qms2010@0203;;DATABASE=SMT;<USERID=QMS><RIGHT=mnuReturnDID,CheckBom,ClearMachine,CompCompare,CycleTime,DeleteDIDmmuCompPNCompare,mmuDIDintegration,mmuQSMS_SapHis,mmuUnlockCompPNCompare,mmuWOInputPlan,MnuAutoDispatch,mnuChangeExtraDIDslot,mnuCheckDispatchQty,mnuCloseWO,mnuCompPrint,mnuDefineBuildType,mnuWipReport,mnuInSpection,MnuUpLoadBom,mnumaintainWOSeq,mnumaintainFeeder,mnuTransferPanaAMI,mnuTransferPanaMSF,mnuTransferFujiNexim,mnuMaintainDIDAutoDispatch,><FACTORY=F2>"
'    StrConn = "PROVIDER=SQLOLEDB;SERVER=10.226.32.101;UID=sa;PWD=qms7sa;;DATABASE=SMT;<USERID=QMS><RIGHT=mnuReturnDID,CheckBom,ClearMachine,CompCompare,CycleTime,DeleteDIDmmuCompPNCompare,mmuDIDintegration,mmuQSMS_SapHis,mmuUnlockCompPNCompare,mmuWOInputPlan,MnuAutoDispatch,mnuChangeExtraDIDslot,mnuCheckDispatchQty,mnuCloseWO,mnuCompPrint,mnuDefineBuildType,mnuWipReport,mnuInSpection,MnuUpLoadBom,mnumaintainWOSeq,mnumaintainFeeder,mnuTransferPanaAMI,mnuTransferPanaMSF,mnuTransferFujiNexim,mnuMaintainDIDAutoDispatch,><FACTORY=T2>"
'    StrConn = "PROVIDER=SQLOLEDB;SERVER=192.168.20.39;UID=sa;PWD=qms7sa;;DATABASE=SMT;<USERID=10804005><RIGHT=mnuReturnDID,CheckBom,ClearMachine,CompCompare,CycleTime,DeleteDIDmmuCompPNCompare,mmuDIDintegration,mmuQSMS_SapHis,MnuUpLoadBom,mnuDIDChkStock,mmuUnlockCompPNCompare,mmuWOInputPlan,MnuAutoDispatch,mnuChangeExtraDIDslot,mnuCheckDispatchQty,mnuCloseWO,mnuCompPrint,mnuDefineBuildType,mnuWipReport,mnuInSpection,MnuUpLoadBom,mnumaintainWOSeq,mnumaintainFeeder,mnuTransferPanaAMI,mnuTransferPanaMSF,mnuTransferFujiNexim,mnuMaintainDIDAutoDispatch,mnuGenXLPrior,><FACTORY=T2>"
    StrConn = "PROVIDER=SQLOLEDB;SERVER=10.94.7.11;UID=sa;PWD=pqmb#7sa;;DATABASE=SMT;<USERID=10804005><RIGHT=mnuReturnDID,CheckBom,ClearMachine,CompCompare,CycleTime,DeleteDIDmmuCompPNCompare,mmuDIDintegration,mmuQSMS_SapHis,MnuUpLoadBom,mnuDIDChkStock,mmuUnlockCompPNCompare,mmuWOInputPlan,MnuAutoDispatch,mnuChangeExtraDIDslot,mnuCheckDispatchQty,mnuCloseWO,mnuCompPrint,mnuDefineBuildType,mnuWipReport,mnuInSpection,MnuUpLoadBom,mnumaintainWOSeq,mnumaintainFeeder,mnuTransferPanaAMI,mnuTransferPanaMSF,mnuTransferFujiNexim,mnuMaintainDIDAutoDispatch,mnuGenXLPrior,mnuSendXLRemainDemand><FACTORY=F7>"
    If InStr(1, StrConn, "<RIGHT=", vbTextCompare) > 0 Then
        strRights = Mid(Mid(StrConn, InStr(1, StrConn, "<RIGHT=", vbTextCompare) + Len("<RIGHT="), Len(StrConn)), 1, InStr(1, Mid(StrConn, InStr(1, StrConn, "<RIGHT=", vbTextCompare) + Len("<RIGHT="), Len(StrConn)), ">") - 1)
    End If
    If InStr(1, StrConn, "<USERID=", vbTextCompare) > 0 Then
        g_userName = Mid(Mid(StrConn, InStr(1, StrConn, "<USERID=", vbTextCompare) + Len("<USERID="), Len(StrConn)), 1, InStr(1, Mid(StrConn, InStr(1, StrConn, "<USERID=", vbTextCompare) + Len("<USERID="), Len(StrConn)), ">") - 1)
    End If
    If InStr(1, StrConn, "<FACTORY=", vbTextCompare) > 0 Then
        g_factory = Mid(Mid(StrConn, InStr(1, StrConn, "<FACTORY=", vbTextCompare) + Len("<FACTORY="), Len(StrConn)), 1, InStr(1, Mid(StrConn, InStr(1, StrConn, "<FACTORY=", vbTextCompare) + Len("<FACTORY="), Len(StrConn)), ">") - 1)
    End If
    If InStr(1, StrConn, "<CHKDOMAIN=", vbTextCompare) > 0 Then ''1165
        chkDomain = Mid(Mid(StrConn, InStr(1, StrConn, "<CHKDOMAIN=", vbTextCompare) + Len("<CHKDOMAIN="), Len(StrConn)), 1, InStr(1, Mid(StrConn, InStr(1, StrConn, "<CHKDOMAIN=", vbTextCompare) + Len("<CHKDOMAIN="), Len(StrConn)), ">") - 1)
    End If
End If
    ProgLine = strLine
    Conn.CommandTimeout = 0
    Conn.CursorLocation = adUseClient
    If Conn.State = 1 Then Conn.Close
    Conn.Open StrConn
    
    ''Stephen connect 39.SMT
    strConnSMT = StrConn
    
  
    '''Get SMT Server (0002)
    g_userRight = Split(strRights, ",")
    SMTServer = GetKeyValue(StrConn, "server")
    If SMTServer = "" Then
        MsgBox "Cant't get SMT Server information !! Call QMS please! "
        End
    Else
    ''1168
        ConnSMT.CursorLocation = adUseClient
        ConnSMT.Open StrConn
        sql = "select smt_db,qsms_db,QSMS_Server from QSMS_SMT_DB where smt_server='" & Trim(SMTServer) & "'" ''AND  BU='" & tSettings.BU & "'"
        Set Rs = ConnSMT.Execute(sql)
        SMTDB = Rs!SMT_DB
        QSMSDB = Rs!qsms_db
        QSMSServer = Rs!QSMS_Server
    End If
    ''Get QSMS Server
    If QSMSServer = "" Then
        MsgBox "Can't get QSMS Server information ! Call QMS please! "
        End
    Else
        IP = QSMSServer
        StrConn = Replace(StrConn, SMTServer, QSMSServer)
        StrConn = Replace(StrConn, SMTDB, QSMSDB)
        StrConn = Replace(StrConn, LCase(SMTDB), QSMSDB)
    End If
    ''Connect QSMS DB
    If Conn.State = 1 Then Conn.Close
    Conn.CursorLocation = adUseClient
    Conn.Open StrConn
    
    Call Setting
    
    ''''''added by Jing (0028)''''''
    chkQty = ReadIniFile("QSMS", "MaxDIDGroupQty", App.path & "\set.ini")
    StrBU = ReadIniFile("COMMON", "BU", App.path & "\set.ini")   'add a flag to NB5 for DeleteMe_Bom  (0010)
    ''''''''''(0008)
   If g_factory = "" Then
        CreateDIDFlag = "N"
        Factory = g_factory
    Else
        CreateDIDFlag = "Y"
        Factory = g_factory
    End If

'    If CheckFacIP = False Then          ''(0014)
'        End
'    End If
    mdiMain.Show
'    frmSelectGroup.Show
    plant = "TB30"
End Sub
Public Sub Setting()
    Dim strSQL As String
    Dim Rs As New ADODB.Recordset
    
'    strSQL = "select [Key], [Value] from QSMS_ProConfig with(nolock) where Line in ('All','" & ProgLine & "') and station='QSMS' Union all select [Key], [Value] from vMesPE_ProConfig with(nolock) where Line in ('All','" & ProgLine & "') and station='QSMS'"
    strSQL = "exec [QSMS_GetSetting] @Line = '" & ProgLine & "', @Station='QSMS'"
    Set Rs = Conn.Execute(strSQL)
    
    CompPrintTimeSpan = 10
    AutoDispatchTimeSpan = 3
    PrintTime = Now
    
    If Rs.EOF = False Then
        While Not Rs.EOF
            Select Case UCase(Rs!key)
                Case "SCANCOMPPN"
                    ScanCompPN = UCase(Rs!Value)
                Case "SCANMSD"
                    ScanMSD = UCase(Rs!Value)
                Case "CHECKBOMLOGON"
                    CheckBomLogon = UCase(Rs!Value)
                Case UCase("CheckReturnForbiddenPN")
                    CheckReturnForbiddenPN = UCase(Rs!Value)
                Case UCase("ChkOldDIDLabelQty")  ''(0061)
                    ChkOldDIDLabelQty = UCase(Rs!Value)
                Case UCase("ChkOneByOneMaterial")  ''(0076)
                    ChkOneByOneMaterial = UCase(Rs!Value)
                Case UCase("NPMMachineType")  ''(1079)
                    NPMMachineType = Trim(Rs!Value)
                Case UCase("MaintainFeederDID")  ''(1103)
                    MaintainFeederDID = Trim(Rs!Value)
                Case UCase("ChkFujiSPL")  ''(1103)
                    ChkFujiSPL = Trim(Rs!Value)
                Case UCase("ChkWOGroupID")  ''(1128)
                    ChkWOGroupID = Trim(Rs!Value)
                Case UCase("chkPrintDIDType")
                    ChkPrintDIDType = Trim(Rs!Value)
                Case UCase("PrintedSeqID")        'ги1147)
                    PrintedSeqID = Trim(Rs!Value)
                Case UCase("BatchControl")        'ги1147)
                    BatchControl = Trim(Rs!Value)
                Case UCase("UnChkCompPN")        'ги1187)
                    UnChkCompPN = Trim(Rs!Value)
                Case UCase("CheckNeedMSD")
                    CheckNeedMSD = Trim(Rs!Value)   ''1188
                Case UCase("CheckWOIFReduceXboard")
                    CheckWOIFReduceXboard = UCase(Trim(Rs!Value))   ''(1190)
                Case UCase("CheckMSDCallBack")      ''1191
                    CheckMSDCallBack = UCase(Trim(Rs!Value))
                Case UCase("CheckBurnDID")      ''1191
                    CheckBurnDID = UCase(Trim(Rs!Value))
                 Case UCase("NoKeepPWD")      ''1191
                     NoKeepPWD = UCase(Trim(Rs!Value))
                 Case UCase("BGAWarehouse")      ''1205
                     BGAWarehouse = UCase(Trim(Rs!Value))
                Case UCase("ChkPNCQ")
                     ChkPNCQ = UCase(Trim(Rs!Value))  ''1211
                Case UCase("CheckBSMaterial")
                     CheckBSMaterial = UCase(Trim(Rs!Value))  ''1213
                Case UCase("ChkEQProgram")
                     ChkEQProgram = UCase(Trim(Rs!Value))  ''1219
                Case UCase("ChkDateCode")
                    ChkDateCode = UCase(Trim(Rs!Value)) ''1222
                Case UCase("CheckDIDByLine")
                    strChkDIDByLine = UCase(Trim(Rs!Value)) ''1276
                Case UCase("PrintedVenderCode")         ''1223
                    PrintedVenderCode = UCase(Trim(Rs!Value))
                Case UCase("NewGroupIDRule")         ''1225
                    NewGroupIDRule = UCase(Trim(Rs!Value))
                Case UCase("UnChkBaseReelQty")         ''1241
                    UnChkBaseReelQty = UCase(Trim(Rs!Value))
                Case UCase("ChkMEBOM_Location")
                    ChkMEBOM_Location = UCase(Trim(Rs!Value))
                Case UCase("DIDAutoOpen")       ''1268
                    DIDAutoOpen = UCase(Trim(Rs!Value))
                Case UCase("Chk_XL_WOPlanSeq")       ''1278
                    strChk_XL_WOPlanSeq = UCase(Trim(Rs!Value))
                Case UCase("LabelPrintCheck")       ''1274
                    LabelPrintCheck = UCase(Trim(Rs!Value))
                Case UCase("Check09Code")
                     Chk09Code = UCase(Trim(Rs!Value)) '1254
                Case UCase("DispatchCompPrint")  '1287
                    DispatchCompPrint = UCase(Trim(Rs!Value)) '1287
                Case UCase("Chk2DCode")  ''1295
                    Chk2DCode = UCase(Trim(Rs!Value))
                Case UCase("CheckBCMS")  ''1296
                    CheckBCMS = UCase(Trim(Rs!Value))
                Case UCase("OneByOneControl")  ''1297
                    OneByOneControl = UCase(Trim(Rs!Value))
                Case UCase("CheckLocation")  ''1298
                    CheckLocation = UCase(Trim(Rs!Value))
                Case UCase("AutoDispatchPrintlable")  ''1300
                    AutoDispatchPrintlable = UCase(Trim(Rs!Value))
                Case UCase("CheckDIDRemainQty")  ''1302
                    CheckDIDRemainQty = UCase(Trim(Rs!Value))
                Case UCase("CheckOldNewPrintType")  ''1304
                    CheckOldNewPrintType = UCase(Trim(Rs!Value))
                Case UCase("CompPrintTimeSpan")  ''1289
                    CompPrintTimeSpan = UCase(Trim(Rs!Value))
                Case UCase("AutoDispatchTimeSpan")  ''1289
                    AutoDispatchTimeSpan = UCase(Trim(Rs!Value))
                Case UCase("CompPrintModbus")  ''1290
                    CompPrintModbus = UCase(Trim(Rs!Value))
                Case UCase("DispatchChkQty")  ''1290
                    DispatchChkQty = UCase(Trim(Rs!Value))
            End Select
            Rs.MoveNext
        Wend
    End If
    
End Sub



Public Function SaveLog(System_Name As String, strIP As String, strUserID As String, StrEventDesc As String)
Dim Rs As New ADODB.Recordset
Dim strSQL As String
strSQL = "Insert into QMS_Log(System_Name,Event_No,SN,User_Name,Desc1,Trans_Date)" & _
            "Select '" & Trim(System_Name) & "','1','" & Trim(strIP) & "','" & Trim(strUserID) & "','('+Host_Name()+')" & Trim(StrEventDesc) & "',dbo.formatdate(getdate(),'YYYYMMDDHHNNSS')"
If Rs.State = 1 Then Rs.Close
Rs.CursorLocation = adUseClient
Set Rs = Conn.Execute(strSQL)
End Function



