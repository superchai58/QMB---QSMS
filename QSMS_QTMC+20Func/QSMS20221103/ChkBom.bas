Attribute VB_Name = "ChkBom"
Option Explicit
'/**********************************************************************************
'**文 件 名: FrmMaintainDID.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人:
'**日    期:
'**描    述: QSMS Maintain DID
'
'**修 改 人     修改日期        描    述
'-----------------------------------------------------------------------------
'**Udall      2008.02.28     当PCN的PCNWO为空或者与当前工单一致时，才生效 (0001)  (PCNWO='' or PCNWO='" & Woinfo.WO & "')
'**Giant      2008.04.28     对于Pilot的工单，CheckBom fail时AutoRelease要发Mail (0002)
'**Sandy      2008.06.19     add log in checkBOM process in all Exit and entrance(0003)
Type WOBasic
    WO As String
    MBPN As String
    Version As String
    WOqty As Long
    SingleSide As Boolean
    CombineQty As Long
    Line As String
    WOLine As String
    jobgroup As String
    Group As String
    Model As String
    GroupWoQty As Long
    Negative As Boolean
    Pilot As String
    CrashFlag As Boolean
    SourceQty As Long
    DestQty As Long
    InitAOIFlag As String ' will use to check if it is MB WO
    
End Type
Public Woinfo As WOBasic
Type MeBomBasic
     machine As String
     Jobpn As String
     Slot As String
     Qty As Long
     Version As String
End Type
Public MeBomINfo(2) As MeBomBasic
Public ReplaceItem As Long
Public JobPCN() As String, FactoryID() As String, LocalIP As String
Public RefreshFlag As String
Public StrBU As String

Public Function GetWoinfo(ByVal WO As String, ByVal RefreshFlag As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
Dim JobPNNum As Long

GetWoinfo = False
'------(1)get Basic wo information
str = "select WO,PN,Mb_Rev,line,Qty,CombineQty,Pilot,InitAOIFlag from Sap_WO_List where WO='" & WO & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
    With Woinfo
         .CombineQty = CLng(Trim(rs!CombineQty))
         .WO = Trim(rs!WO)
         .MBPN = Trim(rs!PN)
         .Model = Mid(.MBPN, 3, 3)
         .Version = Trim(rs!MB_Rev)
         .WOqty = CLng(Trim(rs!Qty))
         .Line = Trim(rs!Line)
         .WOLine = Trim(rs!Line)
         .Pilot = UCase(Trim(rs!Pilot))
        .SourceQty = rs!CombineQty
         .DestQty = Woinfo.CombineQty
         .InitAOIFlag = Trim(rs!InitAOIFlag)
    End With
Else
    InsSAP_BOM_FAIL WO, "", "NO Sap Wo List: "
    Exit Function
End If
'-----(1.0) check if it is WO build on multi line
str = "select distinct line from WO_MultiLine  where WO='" & WO & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   While Not rs.EOF
        Woinfo.Line = Woinfo.Line & "," & rs!Line
        rs.MoveNext
        
   Wend
   Woinfo.Line = "[" & Woinfo.Line & "]"
End If

Call UpdateRev(RefreshFlag)
'-------(1.1)check if Negative MB
str = "select MBPN from QSMS_NegativeBrd where MBPN='" & Woinfo.MBPN & "'"
Set rs = Conn.Execute(str)
If rs.EOF Then
   Woinfo.Negative = False
Else
   Woinfo.Negative = True
End If
'-------(2)Get JObpN Qty and group

str = "Select count(JobPN) from QSMS_JOBBom where Work_Order='" & WO & "'"
Set rs = Conn.Execute(str)
JobPNNum = rs.Fields(0)

str = "select [Group] from Sap_Wo_List where WO='" & WO & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
  Woinfo.Group = Trim(rs![Group])
  str = "Select Count(*) from Sap_Wo_List where [Group]='" & Trim(rs![Group]) & " '"
  Set rs = Conn.Execute(str)
  Woinfo.GroupWoQty = rs.Fields(0)
Else
   'GetSmallBoardGroup = ""
End If
'----(3)check if it is single side brd
str = "Select MBPN from QSMS_SingleSideBrd where MBPN='" & Woinfo.MBPN & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
     Woinfo.SingleSide = True
Else
     Woinfo.SingleSide = False
End If

'---(4)get JobGroup=MB-JobPN+MB-Rev
Woinfo.jobgroup = GetJobGroup(Woinfo.Group)
If Woinfo.jobgroup = "" Then
   GetWoinfo = False
   InsSAP_BOM_FAIL WO, "", "Can not get jobgroup: Maybe the model is not defined in the Model BU shopfloor(PE define),or the Sap Bom upload fail(Can not find Job Bom),please check!"
   Exit Function
End If
'---check if the WO 是单报板
Woinfo.CrashFlag = False
str = "select CombineQty from QSMS_BrdCombineQty where MBPN='" & Woinfo.MBPN & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
  Woinfo.CrashFlag = True
  Woinfo.SourceQty = rs!CombineQty
  Woinfo.DestQty = Woinfo.CombineQty
End If


GetWoinfo = True

End Function

Public Function ChkWOGrouplegal(ByVal WO As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset

str = "EXEC ChkWOGrouplegal '" & WO & "'"
Set rs = Conn.Execute(str)
If UCase(Trim(rs!result)) = "PASS" Then
    ChkWOGrouplegal = True
Else
    ChkWOGrouplegal = False
End If
End Function

Public Function GetMEBOMQty(ByVal WO As String, ByVal COMPPN As String, ByVal UpcompPN As String, ByVal Line As String, ByVal RefreshFlag As String, ByVal BuildType As String) As Long
'  CompLevel   JobPN             Side
'   01         41*********       Component side
'   02         51*********       solder side
'   00         others,which by manual instead of  by machine
Dim str As String
Dim rs As ADODB.Recordset
'Dim CompRs As ADODB.Recordset
Dim tempqty As Long
Dim Jobpn As String
Dim machine As String
Dim TransDate As String

Dim i As Long
GetMEBOMQty = 0

 
str = "select getdate()"
Set rs = Conn.Execute(str)
TransDate = Format(rs.Fields(0), "YYYYMMDD")


'''(1) get me bom Qty from QSMS_MeBom_Chk and Qsms_ReplacePn
'str = "Select Item from QSMS_MEBOM_CHk where CompPN='" & CompPN & "' and jobpn ='" & UpcompPN & "' and machine like '" & Line & "%' " & _
'      "and version='" & Woinfo.Version & "' and JobGroup in " & Woinfo.JobGroup & ""
'Set Rs = Conn.Execute(str)
'If Not Rs.EOF Then
'   If Rs!Item <> 0 Then
'      str = "Select Machine,sum(Qty) as Qty from QSMS_MeBom_Chk  where version='" & Woinfo.Version & "'  and  Machine like '" & Woinfo.Line & "%'  and jobpn ='" & UpcompPN & "' " & _
'       "and jobGroup in " & Woinfo.JobGroup & " and (CompPN='" & CompPN & "' or CompPN in " & _
'       "(Select COmpPN from QSMS_MEBOM_CHk where Item='" & Trim(Rs!Item) & "' and jobpn ='" & UpcompPN & "' and machine like '" & Line & "%' " & _
'      "and version='" & Woinfo.Version & "' and JobGroup in " & Woinfo.JobGroup & ")) " & _
'       " Group By machine"
'   Else
'       str = "Select Machine,sum(Qty) as Qty from QSMS_MeBom_Chk  where version='" & Woinfo.Version & "'  and  Machine like '" & Woinfo.Line & "%'  and jobpn ='" & UpcompPN & "' " & _
'       "and jobGroup in " & Woinfo.JobGroup & " and (CompPN='" & CompPN & "' or CompPN in " & _
'       "(select CompPN from QSMS_ReplacePn where version='" & Woinfo.Version & "' and  JobPN   = '" & UpcompPN & "' " & _
'       "and ID in (select ID from QSMS_ReplacePN where  JObPN = '" & UpcompPN & "' and CompPN='" & CompPN & "' and Version='" & Woinfo.Version & "'))) " & _
'       " Group By machine"
'   End If
'End If
 str = "Select Machine,sum(Qty) as Qty from QSMS_MeBom_Chk  where version='" & Woinfo.Version & "'  and  Machine like '" & Woinfo.Line & "%'  and jobpn ='" & UpcompPN & "' " & _
       "and BuildType='" & BuildType & "' and jobGroup in " & Woinfo.jobgroup & " and (CompPN='" & COMPPN & "' or CompPN in " & _
       "(select CompPN from QSMS_ReplacePn where version='" & Woinfo.Version & "' and  JobPN   = '" & UpcompPN & "' " & _
       "and ID in (select ID from QSMS_ReplacePN where  JObPN = '" & UpcompPN & "' and CompPN='" & COMPPN & "' and Version='" & Woinfo.Version & "'))) " & _
       " Group By machine"

Set rs = Conn.Execute(str)


While Not rs.EOF
      If InStr(1, UCase(Trim(rs!machine)), "OTHERS") > 0 Then
         If Woinfo.SingleSide = True And BuildType = "1" Then
             GetMEBOMQty = GetMEBOMQty + rs!Qty * Woinfo.CombineQty * 2
         Else
             If Woinfo.Negative = True Then
                GetMEBOMQty = GetMEBOMQty + rs!Qty * 4
             Else
                GetMEBOMQty = GetMEBOMQty + rs!Qty * Woinfo.SourceQty
             
             End If
         End If
      Else
         If Woinfo.SingleSide = True Then
            ' GetMEBOMQty = GetMEBOMQty + Rs!Qty * Woinfo.CombineQty
             GetMEBOMQty = GetMEBOMQty + rs!Qty '* 2
         Else
             GetMEBOMQty = GetMEBOMQty + rs!Qty
         End If
      End If
      rs.MoveNext
Wend





End Function


Public Function CheckBom(ByVal WO As String, Optional ByVal RefreshFlag As String = "N", Optional ByVal BuildType As String = "1") As Boolean
Dim str As String
Dim rs As ADODB.Recordset
Dim errMsg As String
Dim BomTest As String
Dim TempCompLevel As String, step As Integer, i As Integer
'On Error GoTo errHandler  -- can not use because AutoRelease also call CheckBom

CheckBom = True
'check if need update

ReDim Preserve JobPCN(0) 'for PCN data

step = 0
If GetWoinfo(WO, RefreshFlag) = False Then
    InsSAP_BOM_FAIL WO, "", "Get WO Info fail : "
    CheckBom = False
    Exit Function
End If
'check WO's QTY if equal in same WOGroup --add by Giant 20070810
If ChkWOGrouplegal(WO) = False Then
    CheckBom = False
    Exit Function
End If

str = "Delete from Sap_Bom_Fail where Work_Order in (select wo from sap_wo_list where [Group]='" & Woinfo.Group & "')"
Conn.Execute str
'(0) insert the sap bom to sap_bom_chk
str = "Delete from Sap_BOM_Chk where work_Order='" & WO & "'"
Conn.Execute str


str = "insert into sap_bom_chk(Work_Order,MBPN,Item, CompPN, Qty ,CompLevel,UpCompPN ,TransDateTime) select Work_Order,MBPN,Item, CompPN, Qty ,CompLevel,UpCompPN ,TransDateTime" & _
    " from sap_bom where work_order='" & WO & "'"
Conn.Execute str

str = "select * from Sap_BOM_Chk where Work_Order='" & WO & "'"
Set rs = Conn.Execute(str)
If rs.EOF Then
   InsSAP_BOM_FAIL WO, "", "NO SAP BOM : "
   CheckBom = False
End If

Call AdjustSapBOM(WO, RefreshFlag)
Call AdjustMeBOm(WO, RefreshFlag, BuildType)
If ChkBuildType = False Then
    CheckBom = False
    Exit Function
End If

'Check Sap bom and ReplacePN
str = "select Work_Order,MBPN,Item,UpCompPN,count(*) as Qty from Sap_BOM_Chk where Work_Order='" & WO & "'" & _
      "group by Work_Order,MBPN,Item,UpCompPN  having count(*)>1 order by Item"
Set rs = Conn.Execute(str)
While Not rs.EOF
      If ChkSAPReplacePN(Trim(rs!Work_Order), Trim(rs!MBPN), Trim(rs!Item), Trim(rs!UpcompPN)) = False Then
         CheckBom = False
      End If
      rs.MoveNext
Wend

step = 1
'(1) check sap bom if lost in Me bom
str = "select Work_order,CompPN ,UpCompPN from Sap_BOM_Chk where Work_order ='" & WO & "'  and CompPn not in " & _
     "(select a.CompPn from QSMS_MeBom_Chk a,QSMS_JobBOM b,Sap_WO_List c where b.Work_Order='" & WO & "' and " & _
     "b.work_order=c.Wo and c.MB_Rev=a.version and b.JObPN=a.JObPN and a.Machine like '" & Woinfo.Line & "%' and jobgroup in " & Woinfo.jobgroup & " )"
    ' " and substring(comppn,1,2) not in (select CompHead from QSMS_UnChkComp)"

Set rs = Conn.Execute(str)
While Not rs.EOF
    If ChkForbiddenPN(rs!Work_Order, rs!COMPPN) = False Then  'add by Giant -(0013)
        If ChkUnChkComp(Trim(rs!COMPPN)) = False Then
            If ChkReplacePN(Trim(rs!COMPPN), Trim(rs!Work_Order), "SAP_BOM", Trim(rs!UpcompPN), RefreshFlag) = False Then
               InsSAP_BOM_FAIL WO, Woinfo.MBPN, "Lost in ME BOM: " & rs!COMPPN & " UpcompPN :" & Trim(rs!UpcompPN)
               CheckBom = False
            End If
        End If
    Else
        CheckBom = False
        Exit Function
    End If
        rs.MoveNext
Wend
step = 2
'(2) check ME bom if lost in Sap Bom
str = "select a.CompPn,Machine,a.JobPN from QSMS_MeBom_Chk a,QSMS_JobBOM b,Sap_WO_List c where b.Work_Order='" & WO & "' and b.work_order=c.Wo and " & _
      "c.MB_Rev=a.version and b.JObPN=a.JObPN and A.Machine like '" & Woinfo.Line & "%' and JobGroup in " & Woinfo.jobgroup & " and a.CompPN not in " & _
      " (select Comppn from Sap_BOM_Chk where Work_order ='" & WO & "')"

Set rs = Conn.Execute(str)
While Not rs.EOF
    If ChkForbiddenPN(WO, rs!COMPPN) = False Then  'add by Giant -(0013)
        If ChkReplacePN(rs!COMPPN, WO, "ME_BOM", Trim(rs!Jobpn), RefreshFlag) = False Then
           InsSAP_BOM_FAIL WO, Woinfo.MBPN, "Lost in SAPBOM: " & rs!COMPPN & "  UpcompPN :" & rs!Jobpn
           CheckBom = False
        End If
    Else
        CheckBom = False
        Exit Function
    End If
        rs.MoveNext
Wend
'(3)check Comp Qty from sap bom to mebom
step = 3
If Woinfo.Negative = False Then
    str = "select Work_Order,MBPN,Item,CompPN,Qty,UpCompPN from Sap_BOm_Chk where Work_Order='" & WO & "'"
    'and substring(comppn,1,2) not in (select CompHead from QSMS_UnChkComp)"
    
    Set rs = Conn.Execute(str)
    While Not rs.EOF
          If ChkUnChkComp(Trim(rs!COMPPN)) = False Then
            If ChkCompQty(Trim(rs!Work_Order), Trim(rs!COMPPN), Woinfo.WOqty, Trim(rs!UpcompPN), RefreshFlag, BuildType) = False Then
                
                CheckBom = False
            End If
          End If
          rs.MoveNext
       
    Wend
    step = 4
    ''''(4)check compqty from mebom to sap bom
    str = "Select distinct CompPN,JObpN from QSMS_MeBom_Chk where jobpn in (select jobpn from qsms_Jobbom where work_order='" & Woinfo.WO & "') and version='" & Woinfo.Version & "' " & _
          " and JobGroup in " & Woinfo.jobgroup & " and machine like '" & Woinfo.Line & "%' and BuildType='" & BuildType & "'"
    Set rs = Conn.Execute(str)
    
    While Not rs.EOF
            If ChkUnChkComp(Trim(rs!COMPPN)) = False Then
            
                If ChkCompQty(Woinfo.WO, Trim(rs!COMPPN), Woinfo.WOqty, Trim(rs!Jobpn), RefreshFlag, BuildType) = False Then
        
                    CheckBom = False
                End If
           End If
          rs.MoveNext
    Wend
Else
    str = "exec CheckBOM '" & WO & "','" & RefreshFlag & "'"
    Conn.Execute str
    str = "select * from sap_Bom_fail where work_order='" & WO & "'"
    Set rs = Conn.Execute(str)
    If Not rs.EOF Then
       CheckBom = False
    End If
End If
step = 5    '(0002)
'(5)if check bom pass but the WO is Pilot = "NEW" Or Pilot = "EOL",the send check-bom fail alarm email
CheckBomPilotWO = True
If Woinfo.Pilot = "NEW" Or Woinfo.Pilot = "EOL" Then
    str = "select *  from Sap_BOM_Fail  where Work_Order ='" & WO & "'"
    Set rs = Conn.Execute(str)
    If Not rs.EOF Then
       CheckBomPilotWO = False
    End If
End If
'(5)if check bom pass but the comp is document and didn't have ECN NO then check bom fail
'If ChkECNPass(Wo) = False Then
'    CheckBom = False
'End If


step = 6
If CheckBom = True Or BomTest = "456" Or Woinfo.Pilot = "NEW" Or Woinfo.Pilot = "EOL" Then
     CheckBom = True 'Leimo 20061114 for Gary Auto release flag
    '2006/9/17, Alex Wang; Insert WO Status 10, means CheckBom pass
    str = "Exec WO_Status_Update '" & WO & "', '10'"
    Conn.Execute (str)
    Call InsertToQSMS_WO(WO, BuildType)
    '05->fail, 10->pass, 20->release sn ok
    str = "update SAP_WO_LIST set status='10' where wo='" & WO & "' and status<'20'"
    Conn.Execute (str)
    '记录第一次CheckBom Pass的时间  --20070601
    str = "update SAP_WO_LIST set CheckBomPassDateTime=dbo.formatdate(getdate(),'YYYYMMDDHHNNSS') where wo='" & WO & "' and CheckBomPassDateTime=''"
    Conn.Execute (str)
Else
    str = "update SAP_WO_LIST set status='05' where wo='" & WO & "' and status<'20'"
    Conn.Execute (str)
    
    If UBound(JobPCN) > 0 Then
        For i = 1 To UBound(JobPCN) ''Udall add it for show the PCN of the JObPN when check bom fail  071004
            InsSAP_BOM_FAIL WO, Woinfo.MBPN, JobPCN(i)
        Next i
    End If
End If

'delete sap_bom_chk
str = "Delete from Sap_BOM_Chk where work_Order='" & WO & "'"
Conn.Execute str

'Exit Function
'errHandler:
'    MsgBox ("CheckBom, step:" & step & ", " & Err.Description)
End Function
Public Function ChkCompQty(ByVal WO As String, COMPPN As String, ByVal WOqty As Long, ByVal UpcompPN As String, ByVal RefreshFlag As String, ByVal BuildType As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
Dim SAPBomQty, MEBomQty As Long
Dim i As Long
Dim TempCompLevel As String
Dim Item As String
Dim NegativeSideQty As Integer


'  CompLevel   JobPN             Side
'   01         41*********       Component side
'   02         51*********       solder side
'   00         others,which by manual insteda of  by machine
'(1) sum Sap bom  Qty accoding to WO & Item
'
'Str = "select sum(qty) from sap_bom where work_order='" & WO & "' and item='" & Item & "' and UpCompPN='" & Upcomppn & "'"
'Set Rs = Conn.Execute(Str)
'If Not Rs.EOF Then
'   SAPBomQty = Rs.Fields(0)
'End If

str = "Select Item from Sap_BOM_Chk where work_order='" & WO & "' and UpCompPN='" & UpcompPN & "' and ( CompPN='" & COMPPN & "' or CompPN in " & _
       "(select CompPN from QSMS_ReplacePn where version='" & Woinfo.Version & "' and  JobPN   = '" & UpcompPN & "' " & _
       "and ID in (select ID from QSMS_ReplacePN where  JObPN = '" & UpcompPN & "' and CompPN='" & COMPPN & "' and Version='" & Woinfo.Version & "'))) "
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   Item = Trim(rs!Item)
Else
   If ChkAddNewCompInMEBOM(WO, UpcompPN, COMPPN) = True Then
      ChkCompQty = True
      Exit Function
   End If
End If



'(1) Get ME BOm Qty

If BuildType = "4" Then
    MEBomQty = GetMEBOMQtyforWOMultiLine(COMPPN, UpcompPN)
Else
    MEBomQty = CDbl(GetMEBOMQty(Trim(WO), Trim(COMPPN), Trim(UpcompPN), Woinfo.Line, RefreshFlag, BuildType)) * WOqty
End If



MEBomQty = MEBomQty / Woinfo.CombineQty
'-----the board is unique side and build on both side
If Woinfo.SingleSide = True And BuildType = "1" Then
     MEBomQty = MEBomQty / 2
End If
'如果是单报板，因为MEＢＯＭ用的是Ｌａｙｏｕｔ的数量，而实际工单的combine Qty可能<Layout　Qty
If Woinfo.CrashFlag = True Then
   MEBomQty = MEBomQty * Woinfo.DestQty / Woinfo.SourceQty
End If
'If Trim(CompPN) = "BC521G30Z06" Then
'   MsgBox "K"
'End If
If Woinfo.Negative = True Then
   MEBomQty = MEBomQty * GetNegativeMESide(WO, COMPPN)

   'mark by leimo ,阴阳板单报板
   If Woinfo.CombineQty = 3 Then
      MEBomQty = MEBomQty * Woinfo.CombineQty / 4
   End If
   MEBomQty = GetNegativeNeedQty(COMPPN, WO, MEBomQty, UpcompPN)
   SAPBomQty = GetNegativeSapBOMQty(WO, COMPPN, Item, UpcompPN)

Else
    
    str = "select sum(qty) from sap_bom_Chk where work_order='" & WO & "' and item='" & Item & "' and UpCompPN='" & UpcompPN & "'"
    Set rs = Conn.Execute(str)
    If Not rs.EOF Then
       SAPBomQty = rs.Fields(0)
    End If

End If

If (SAPBomQty - MEBomQty <= 2) And (SAPBomQty - MEBomQty >= -2) Then
   ChkCompQty = True
        
Else
   If Woinfo.Negative = True And MEBomQty <> 0 Then
      If SAPBomQty Mod MEBomQty = 0 Then
          ChkCompQty = True
      End If
   Else
      ChkCompQty = False
   End If
End If

If ChkCompQty = False Then
   InsSAP_BOM_FAIL WO, Woinfo.MBPN, "Comp Qty does not match: " & COMPPN & " (SAP_BOM Qty:" & SAPBomQty & ")" & "(ME Bom Qty:" & MEBomQty & ")" & " UpcomppN:" & UpcompPN
    
End If


End Function

Public Sub InsSAP_BOM_FAIL(ByVal Work_Order As String, ByVal MBPN As String, ERR_DESC As String)
    
    Dim Tran_Date As String, Tran_Time As String
    Tran_Date = Format(Now, "YYYYMMDD")
    Tran_Time = Format(Now, "HHNNSS")
    strsql = "Insert SAP_BOM_FAIL(Work_Order,MBPN,ERR_DESC,Tran_Date,Tran_Time) values('" & Trim(Work_Order) & "','" & Trim(MBPN) & "',N'" & (ERR_DESC) & "'," & _
        " '" & Tran_Date & "','" & Tran_Time & "')"
    Conn.Execute strsql
End Sub
Public Function ChkSAPReplacePN(ByVal WO As String, ByVal MBPN As String, ByVal Item As String, ByVal UpcompPN As String) As Boolean
Dim str As String, str2 As String
Dim rs As ADODB.Recordset, Rs2 As ADODB.Recordset
Dim COMPPN As String, strCompPN As String, CompPN2 As String, strID() As String
Dim t As Integer, i As Integer, j As Integer

t = 0
ChkSAPReplacePN = True
str = "select * from Sap_BOM_Chk where Work_Order='" & WO & "' and MBPN='" & MBPN & "' and Item='" & Item & "' and UpCompPN='" & UpcompPN & "'"
Set rs = Conn.Execute(str)
ReDim strID(rs.RecordCount - 1)
While rs.EOF = False
      COMPPN = COMPPN & "'" & Trim(rs!COMPPN) & "',"
      strCompPN = strCompPN & Trim(rs!COMPPN) & ","
      str2 = "select * from QSMS_ReplacePN where JobPN='" & UpcompPN & "' and CompPN='" & Trim(rs!COMPPN) & "' and Version='" & Woinfo.Version & "'"
      Set Rs2 = Conn.Execute(str2)
      If Rs2.EOF Then
         ChkSAPReplacePN = False
      Else
        strID(t) = Trim(Rs2!ID)
        t = t + 1
      End If
      rs.MoveNext
Wend
If ChkSAPReplacePN = False Then
   InsSAP_BOM_FAIL WO, Woinfo.MBPN, "替代料设置错误: " & strCompPN & " 在SAP Bom中是作为替代料,但在替代料表中没有设置完整,请核对!" & "UpcompPN :" & Trim(UpcompPN)
   Exit Function
End If
For i = 0 To t - 1
   For j = i + 1 To t - 1
       If strID(i) <> strID(j) Then
          ChkSAPReplacePN = False
       End If
   Next j
Next i
If ChkSAPReplacePN = False Then
   InsSAP_BOM_FAIL WO, Woinfo.MBPN, "替代料设置错误: " & strCompPN & " 在SAP Bom是作为替代料,但在替代料表中设置有差异,请核对!" & "UpcompPN :" & Trim(UpcompPN)
   Exit Function
End If

COMPPN = Mid(COMPPN, 1, Len(COMPPN) - 1)
str = "select * from Sap_BOM_Chk where Work_Order='" & WO & "' and MBPN='" & MBPN & "' and UpCompPN='" & UpcompPN & "' and CompPN in" & _
      "(select CompPN from QSMS_ReplacePN where JobPN='" & UpcompPN & "' and ID='" & strID(0) & "' and Version='" & Woinfo.Version & "' and CompPN not in(" & COMPPN & "))"
Set rs = Conn.Execute(str)
While rs.EOF = False
      CompPN2 = CompPN2 & Trim(rs!COMPPN) & ","
      If rs!Item <> Item Then
         ChkSAPReplacePN = False
      End If
    rs.MoveNext
Wend
If ChkSAPReplacePN = False Then
   InsSAP_BOM_FAIL WO, Woinfo.MBPN, "替代料设置错误: " & strCompPN & " 在SAP Bom是作为替代料；" & strCompPN & CompPN2 & " 在替代料表中设置为替代料；" & CompPN2 & " 在 SAPBom 中有出现但却不是做为替代料,请核对!" & "UpcompPN :" & Trim(UpcompPN)
   Exit Function
End If

End Function


Public Function ChkReplacePN(ByVal COMPPN As String, ByVal WO As String, ByVal Ctype As String, ByVal UpcompPN As String, ByVal RefreshFlag As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
Dim ID As String
Dim Jobpn As String
ID = ""
ChkReplacePN = True
'(0)check if the comppn is second source
If SecondSource(COMPPN) = True Then
   Exit Function
End If


'(1) Get Replace ID accroding to WO.Version & ComppN
str = "select ID from QSMS_ReplacePN where  CompPN='" & COMPPN & "' and JobPN='" & UpcompPN & "' and Version='" & Woinfo.Version & "'"
Set rs = Conn.Execute(str)

If rs.EOF Then
   If Woinfo.Negative = True Then
      str = "Select count(*) from SAP_BOM where Work_Order='" & WO & "' and CompPN='" & COMPPN & "'"
      Set rs = Conn.Execute(str)
      If rs.Fields(0) >= 2 Then
         ChkReplacePN = False
         Exit Function
      End If
      If rs.Fields(0) = 1 Or rs.Fields(0) = 0 Then
         str = "select ID from QSMS_ReplacePN where  CompPN='" & COMPPN & "' and JobPN in (select JobPN from QSMS_JobBOM where work_order='" & WO & "') and Version='" & Woinfo.Version & "'"
         Set rs = Conn.Execute(str)
         If rs.EOF Then
            Exit Function
         Else
            ID = Trim(rs!ID)
         End If
      End If
   Else
       ChkReplacePN = False
       Exit Function
   End If
  
Else
   ID = Trim(rs!ID)
End If


Select Case UCase(Ctype)
       Case "SAP_BOM" ' check if lost in MeBom
             str = "select a.CompPn from QSMS_MeBom_Chk a,QSMS_ReplacePN b where a.JObPN=b.JobPN and  a.version=b.version  " & _
                    "and a.version='" & Woinfo.Version & "' and a.compPn=b.compPN and b.ID='" & ID & "' and a.JObPN= '" & UpcompPN & "' and A.machine like '" & Woinfo.Line & "%'"
             Set rs = Conn.Execute(str)
             If rs.EOF Then
                    ChkReplacePN = False
                    Exit Function
             End If
       Case "ME_BOM" ' check if lost in SAP BOM
             If Woinfo.Negative = True Then
                str = "select ID from QSMS_ReplacePN where  JObPN in(select jobpn from qsms_jobbom where work_order='" & WO & "') and Version='" & Woinfo.Version & "' and ID='" & ID & "'" & _
                      " and CompPN in (select CompPN from Sap_BOm_Chk where Work_Order='" & WO & "' )"
             Else
                str = "select ID from QSMS_ReplacePN where  JObPN ='" & UpcompPN & "' and Version='" & Woinfo.Version & "' and ID='" & ID & "'" & _
                " and CompPN in (select CompPN from Sap_BOm_Chk where Work_Order='" & WO & "' )"
             End If
             Set rs = Conn.Execute(str)
             If rs.EOF Then
                   If ChkAddNewCompInMEBOM(Woinfo.WO, UpcompPN, COMPPN) = False Then 'check if it is PCN---document comppn
                      ChkReplacePN = False
                      Exit Function
                   End If
             End If
End Select
End Function




Public Function UpdateRev(ByVal RefreshFlag As String)
Dim str As String
Dim rs As ADODB.Recordset
Dim TransDate As String
str = "select getdate()"
Set rs = Conn.Execute(str)
TransDate = Format(rs.Fields(0), "YYYYMMDD")

If UCase(RefreshFlag) = "Y" Then
    str = "select NewVersion from QSMS_DocuComp where  version='" & Woinfo.Version & "' and  " & _
          " JobPN in (select JobPn from QSMS_JobBOM where Work_Order='" & Woinfo.WO & "') " & _
          " and BeginDate<='" & TransDate & "' and EndDate>='" & TransDate & "' and (PCNWO='' or PCNWO='" & Woinfo.WO & "') and FuncType='UpdRev' and EffectiveFlag='Y'"    ''(0001)  (PCNWO='' or PCNWO='" & Woinfo.WO & "')
    Set rs = Conn.Execute(str)
    If rs.EOF Then
    Else
       
        str = "exec PCN_UpdateRev '" & Woinfo.WO & "','" & Woinfo.MBPN & "','" & rs!NewVersion & "'"
       
        Conn.Execute str
        Woinfo.Version = Trim(rs!NewVersion)
    End If

End If
End Function

Public Function InsertToQSMS_WO(ByVal WO As String, ByVal BuildType As String)
Dim str As String
Dim rs As ADODB.Recordset
Dim RsBom As ADODB.Recordset
Dim Item As Long
Dim SideValue As String
Dim TransDateTime As String
Dim ChkBOMFlag As Boolean
Dim mailbody As String
Dim FileNum As Integer

str = "select work_order from QSMS_Dispatch where work_order='" & WO & "'"
Set rs = Conn.Execute(str)
If rs.EOF Then
   str = "Delete QSMS_Wo where Work_Order='" & WO & "'"
   Conn.Execute str
End If
'mark by leimo 20070301
'If Woinfo.SingleSide = True Then
'   SideValue = GetSingleSideDispatch(WO, Woinfo.Line)
'End If
str = "Select getdate()"
Set rs = Conn.Execute(str)
TransDateTime = Format(rs(0), "YYYYMMDDHHMMSS")

'(1) update qsms_wo set freshflag='0'
'Refreshflag 0--- the first time has but new without
'Refreshflag 1--- the first time has and new has too
'Refreshflag 2--- the first time without and new has
str = "Select distinct Work_Order from QSMS_WO where Work_Order='" & WO & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   ChkBOMFlag = True
   str = "Update QSMS_WO set RefreshFlag='0' where work_order='" & WO & "'"
   Conn.Execute (str)
Else
   ChkBOMFlag = False
End If

Item = 0
ReplaceItem = 0

str = "select Machine,CompPN,Slot,LR,Qty,JobPN,JobGroup,Item,BuildType,Side from QSMS_MeBom_Chk where JobPN in (select JobPn from QSMS_JobBom where Work_Order='" & WO & "') and version='" & Woinfo.Version & "' " & _
      "and (Machine like '" & Woinfo.Line & "%' ) and newflag<>'Y' and JobGroup in " & Woinfo.jobgroup & " and buildtype='" & BuildType & "' order by comppn,side,jobpn,machine,slot,lr "
Set RsBom = Conn.Execute(str)
While Not RsBom.EOF
      
      Call InsertQSMSWOByComp(WO, Trim(RsBom!COMPPN), Trim(RsBom!machine), RsBom!Qty, Trim(RsBom!Slot), Trim(RsBom!LR), Trim(RsBom!Jobpn), Trim(RsBom!jobgroup), Trim(RsBom!Item), Trim(RsBom!BuildType), Trim(RsBom!Side))
      RsBom.MoveNext
Wend
'(1.1) update PlanQty,PlanNeedQty,PlanBalanceQty by leimo 20070725
       str = "Exec QSMS_UpdatePlanDispatchQty '" & WO & "','QSMS'"
       Conn.Execute str

'(2):  (1) delete record from qsms_wo(2)Delete from QSMS_Dispatch (3)Restore DID qty where refreshflag='Y'
'save the record to QSMS_WODiff where RefreshFlag=0 or RefreshFlag=2
If ChkBOMFlag = True Then

        str = "select Work_Order,Line,WoQty,JobPN,JobGroup,Machine,CompPN,Slot,LR,Item,BaseQty,NeedQty,DispatchQty,BalanceQty,MachineFinishedFlag,WoFinishedFlag,RefreshFlag,'" & TransDateTime & "',BuildType,Side " & _
              "from QSMS_WO where Work_Order='" & WO & "' and (RefreshFlag='2' or (RefreshFlag='0' and DispatchQty>0))"
            
         Set rs = Conn.Execute(str)
        Do While Not rs.EOF
            str = "delete qsms_wo_diff where work_order='" & rs!Work_Order & "' and machine='" & rs!machine & "' and comppn='" & rs!COMPPN & "' and slot='" & rs!Slot & "' and lr='" & rs!LR & "'"
            Conn.Execute str
            str = "insert into QSMS_WO_Diff(Work_Order,Line,WoQty,JobPN,JobGroup,Machine,CompPN,Slot,LR,Item,BaseQty,NeedQty,DispatchQty,BalanceQty,MachineFinishedFlag,WoFinishedFlag,RefreshFlag,OriginalFlag,ChkDateTime,BuildType,Side) " & _
                  "values " & _
                  "('" & rs!Work_Order & "','" & rs!Line & "','" & rs!WOqty & "','" & rs!Jobpn & "','" & rs!jobgroup & "','" & rs!machine & "','" & rs!COMPPN & "','" & rs!Slot & "','" & rs!LR & "','" & rs!Item & "'," & _
                  "'" & rs!BaseQty & "','" & rs!NeedQty & "','" & rs!DispatchQty & "','" & rs!BalanceQty & "','" & rs!MachinefinishedFlag & "','" & rs!wofinishedflag & "','" & rs!RefreshFlag & "','" & rs!RefreshFlag & "','" & TransDateTime & "','" & rs!BuildType & "','" & rs!Side & "')"
            Conn.Execute str
            rs.MoveNext
        Loop
        
        str = "Delete from QSMS_WO where work_order='" & WO & "' and RefreshFlag='0' "
        Conn.Execute str

     'check if need send mail
     'str = "Select work_order,line,woqty,jobgroup,machine,comppn,replace(slot,'-','--'),lr,baseqty,needqty,dispatchqty from QSMS_WO_Diff where work_Order='" & Wo & "' and (RefreshFlag='2' or (RefreshFlag='0' and DispatchQty>0)) and ChkDateTime='" & TransDateTime & "' order by refreshflag"
     str = "Select top 1 *  from QSMS_WO_Diff where work_Order='" & WO & "' and RefreshFlag='0' and DispatchQty>0 and ChkDateTime='" & TransDateTime & "'"
     Set rs = Conn.Execute(str)
     If Not rs.EOF Then
     '20061211 mark by leiom temp.
        str = "exec Auto_ChangeMachineSlot '" & WO & "'"
        Set rs = Conn.Execute(str)

        '**** sendmail ****
'        If Dir(App.Path + "\CheckBomDiff.xls", vbDirectory) <> "" Then
'            Kill App.Path + "\CheckBomDiff.xls"
'        End If
'        fileNum = FreeFile()
'        Open App.Path + "\CheckBomDiff.xls" For Output As #fileNum
'        Print #fileNum, "Work_order" & Chr(9) & "line" & Chr(9) & "WoQty" & Chr(9) & "jobgroup" & Chr(9) & "Machine" & Chr(9) & "Comppn" & Chr(9) & "slot" & Chr(9) & "LR" & Chr(9) & "BaseQty" & Chr(9) & "NeedQty" & Chr(9) & "DispatchQty"
'        Print #fileNum, Rs.GetString
'        Close #fileNum
'
'        mailbody = "Dear all," + vbCrLf + "Please check the difference between the new MEBom & old MEBom, and re-maintain it to the right Machine!" + vbCrLf + "Thanks ! "
'        BU = ReadIniFile("Common", "BU", App.Path & "\set.ini")
'        'Call CopyToExcel(Rs)
'        If BU <> "" Then
'            SendMailDiff BU & "_QSMS", "CheckBomDiff", mailbody, App.Path + "\CheckBomDiff.xls", ""
'        End If
        '**** end ****
     End If
     
     

End If

'(3)Update machine finished flag

Call UpdateMachineFlagByWO(WO)

End Function
Public Function InsertQSMSWOByComp(ByVal WO As String, ByVal COMPPN As String, ByVal machine As String, ByVal BaseQty As Long, ByVal Slot As String, ByVal LR As String, ByVal Jobpn As String, ByVal jobgroup As String, ByVal Item As String, ByVal BuildType As String, ByVal Side As String)
Dim str As String
Dim rs As ADODB.Recordset
Dim TempRs As ADODB.Recordset
Dim ID As String
Dim NeedQty As Long
Dim DispatchQty As Long
'Dim SecondSource As Boolean
'Dim JobGroup As String
'##################for JuJi system,one slot has two subslot(L,R),so maybe need insert the record by LR.----mark by leimo 20060516##################

NeedQty = 0
'If CompPN = "BC521G30Z06" Then
'   MsgBox "OK"
'End If


NeedQty = CDbl(BaseQty) / Woinfo.CombineQty * Woinfo.WOqty
If Woinfo.SingleSide = True And BuildType = "1" Then
    If InStr(1, UCase(machine), "OTHERS") > 0 Then
    Else
       NeedQty = NeedQty / 2
    End If
End If
'If UCase(Trim(CompPN)) = "DA0C0ATB4B1" Then
'   MsgBox "OK"
'End If


If InStr(1, UCase(machine), "OTHERS") > 0 Then
     NeedQty = BaseQty * Woinfo.WOqty
     BaseQty = BaseQty * Woinfo.SourceQty
Else
'如果是单报板，因为MEＢＯＭ用的是Ｌａｙｏｕｔ的数量，而实际工单的combine Qty可能<Layout　Qty
     If Woinfo.CrashFlag = True Then
        NeedQty = NeedQty * Woinfo.DestQty / Woinfo.SourceQty
     End If
End If
If BuildType = "4" Then
   NeedQty = GetNeedQtyforWOMultiLine(BaseQty, machine, Side)
   If NeedQty = 0 Then
      Exit Function
   End If
End If

If Woinfo.Negative = True Then
   NeedQty = GetNegativeNeedQty(COMPPN, WO, NeedQty, Jobpn)
End If
'SecondSource = ChkIFSecondSourceComp(Wo, CompPN, Machine, BaseQty, Slot, LR, JobPN, JobGroup)

str = "select ID from QSMS_ReplacePN where  CompPN='" & COMPPN & "' and JObPN ='" & Jobpn & "' and Version='" & Woinfo.Version & "'"
Set rs = Conn.Execute(str)
If rs.EOF Then
     'check if it is second source add by leimo 20061123
   
     If Item = "0" Then 'if it is not second source
             str = "select Work_Order from QSMS_WO where Work_Order='" & WO & "' and CompPN='" & COMPPN & "' and Slot='" & Slot & "' and LR='" & Trim(LR) & "' and Machine='" & machine & "' and Side='" & Side & "' and JobGroup='" & jobgroup & "'"
             Set TempRs = Conn.Execute(str)
             If TempRs.EOF Then
                       
                str = "insert into QSMS_Wo(Work_Order,Line,WoQty,JobPN,JobGroup,Machine,CompPN,Slot,LR,Item,BaseQty,NeedQty,DispatchQty,BalanceQty,MachineFinishedFlag,WoFinishedFlag,RefreshFlag,BuildType,Side) values" & _
                     "('" & Woinfo.WO & "','" & Woinfo.WOLine & "'," & Woinfo.WOqty & ",'" & Jobpn & "','" & jobgroup & "','" & machine & "','" & COMPPN & "','" & Slot & "','" & LR & "','0', " & BaseQty & "," & NeedQty & ",0,-" & NeedQty & ",'N','N','2','" & BuildType & "','" & Side & "' )"
                      
                Conn.Execute str
            Else
                str = "Update QSMS_Wo set BaseQty='" & BaseQty & "',NeedQty=" & NeedQty & ", Item='0',Balanceqty=dispatchQty-" & NeedQty & ",JobGroup='" & jobgroup & "',RefreshFlag='1',Side='" & Side & "',buildtype='" & BuildType & "' where Work_order='" & WO & "' and CompPN='" & COMPPN & "' and Slot='" & Slot & "' and LR='" & LR & "' and Machine='" & machine & "' and side='" & Side & "' and JobGroup='" & jobgroup & "'"
                Conn.Execute str
            End If
    Else 'it is second source
           ReplaceItem = ReplaceItem + 1
           Call InsertSeconSourceComp(ReplaceItem, NeedQty, WO, COMPPN, machine, BaseQty, Slot, LR, Jobpn, jobgroup, BuildType, Side)
           
                  
    End If
Else
   ReplaceItem = ReplaceItem + 1
   ID = Trim(rs!ID)
   'MAX FUNCTION might return null
   str = "select ISNULL(max(DispatchQty),0) as DispatchQty from QSMS_WO where Work_Order='" & WO & "' and CompPN in (select CompPN from QSMS_ReplacePN where  ID='" & ID & "' and JObPN='" & Jobpn & "' and version='" & Woinfo.Version & "') " & _
         "and Slot='" & Slot & "' and LR='" & LR & "' and Machine='" & machine & "' and Side='" & Side & "' and JobGroup='" & jobgroup & "'"
   Set rs = Conn.Execute(str)
   If Not rs.EOF Then
      DispatchQty = rs!DispatchQty
   End If

   str = "select CompPN from QSMS_ReplacePN where  ID='" & ID & "' and JObPN='" & Jobpn & "' and version='" & Woinfo.Version & "'"
   Set rs = Conn.Execute(str)
   While Not rs.EOF
    
         str = "select Work_Order from QSMS_WO where Work_Order='" & WO & "' and CompPN='" & rs!COMPPN & "' and Slot='" & Slot & "' and LR='" & LR & "' and Machine='" & machine & "' and Side='" & Side & "' and JobGroup='" & jobgroup & "'"
         Set TempRs = Conn.Execute(str)
        
         If TempRs.EOF Then
        
             str = "insert into QSMS_Wo(Work_Order,Line,WoQty,JobPN,JobGroup,Machine,CompPN,Slot,LR,Item,BaseQty,NeedQty,DispatchQty,BalanceQty,MachineFinishedFlag,WoFinishedFlag,RefreshFlag,BuildType,Side) values" & _
             "('" & Woinfo.WO & "','" & Woinfo.WOLine & "'," & Woinfo.WOqty & ",'" & Jobpn & "','" & jobgroup & "','" & machine & "','" & rs!COMPPN & "','" & Slot & "','" & LR & "','" & ReplaceItem & "', " & BaseQty & "," & NeedQty & "," & DispatchQty & "," & DispatchQty & "-" & NeedQty & ",'N','N','2','" & BuildType & "','" & Side & "' )"
            Conn.Execute str
         Else
            str = "Update QSMS_Wo set BaseQty='" & BaseQty & "',NeedQty=" & NeedQty & " , Item='" & ReplaceItem & "',Balanceqty=DispatchQty-" & NeedQty & ",JobGroup='" & jobgroup & "' ,RefreshFlag='1',Side='" & Side & "',buildtype='" & BuildType & "' where Work_order='" & WO & "' and CompPN='" & Trim(rs!COMPPN) & "' and Slot='" & Slot & "' and LR='" & LR & "' and Machine='" & machine & "' and side='" & Side & "' and JobGroup='" & jobgroup & "'"
            Conn.Execute str
         End If
         rs.MoveNext
   Wend
   If Item <> "0" Then
        Call InsertSeconSourceComp(ReplaceItem, NeedQty, WO, COMPPN, machine, BaseQty, Slot, LR, Jobpn, jobgroup, BuildType, Side)
   End If
End If
End Function
Public Function ChkIFSecondSourceComp(ByVal WO As String, ByVal COMPPN As String, ByVal machine As String, ByVal BaseQty As Long, ByVal Slot As String, ByVal LR As String, ByVal Jobpn As String, ByVal jobgroup As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
ChkIFSecondSourceComp = False
str = "select * from QSMS_MEBOM_CHk where Jobpn='" & Jobpn & "' and JobGroup='" & jobgroup & "' and machine='" & machine & "'  " & _
      "and version='" & Woinfo.Version & "' and CompPN='" & COMPPN & "' and item<>0 and slot='" & Slot & "' and LR='" & LR & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
    ChkIFSecondSourceComp = True
End If
End Function
Public Function InsertSeconSourceComp(ByVal ReplaceItem As Long, NeedQty As Long, ByVal WO As String, ByVal COMPPN As String, ByVal machine As String, ByVal BaseQty As Long, ByVal Slot As String, ByVal LR As String, ByVal Jobpn As String, ByVal jobgroup As String, ByVal BuildType As String, ByVal Side As String)
Dim rs As ADODB.Recordset, TempRs As ADODB.Recordset
Dim str As String
Dim DispatchQty As Long
DispatchQty = 0

str = "select DispatchQty from QSMS_WO where Work_Order='" & WO & "' and CompPN in " & _
     "(select CompPN from QSMS_MEBOM_CHk where JobGroup='" & jobgroup & "' and Machine='" & machine & "' and JObPN='" & Jobpn & "' and version='" & Woinfo.Version & "' and  slot='" & Slot & "' and LR='" & LR & "') " & _
      "and Slot='" & Slot & "' and LR='" & LR & "' and Machine='" & machine & "' and Side='" & Side & "' and JobGroup='" & jobgroup & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   DispatchQty = rs!DispatchQty
End If


str = "select CompPN from QSMS_MEBOM_CHk   where JobGroup='" & jobgroup & "' and Machine='" & machine & "' and JObPN='" & Jobpn & "' and version='" & Woinfo.Version & "' and  slot='" & Slot & "' and LR='" & LR & "' and Side='" & Side & "' "
Set rs = Conn.Execute(str)
While Not rs.EOF
 
      str = "select Work_Order from QSMS_WO where Work_Order='" & WO & "' and CompPN='" & rs!COMPPN & "' and Slot='" & Slot & "' and LR='" & LR & "' and Machine='" & machine & "' and Side='" & Side & "' and JobGroup='" & jobgroup & "'"
      Set TempRs = Conn.Execute(str)
     
      If TempRs.EOF Then
     
          str = "insert into QSMS_Wo(Work_Order,Line,WoQty,JobPN,JobGroup,Machine,CompPN,Slot,LR,Item,BaseQty,NeedQty,DispatchQty,BalanceQty,MachineFinishedFlag,WoFinishedFlag,RefreshFlag,BuildType,Side) values" & _
          "('" & Woinfo.WO & "','" & Woinfo.WOLine & "'," & Woinfo.WOqty & ",'" & Jobpn & "','" & jobgroup & "','" & machine & "','" & rs!COMPPN & "','" & Slot & "','" & LR & "','" & ReplaceItem & "', " & BaseQty & "," & NeedQty & "," & DispatchQty & "," & DispatchQty & "-" & NeedQty & ",'N','N','2','" & BuildType & "','" & Side & "' )"
         Conn.Execute str
      Else
         str = "Update QSMS_Wo set BaseQty='" & BaseQty & "',NeedQty=" & NeedQty & " , Item='" & ReplaceItem & "',Balanceqty=DispatchQty-NeedQty,JobGroup='" & jobgroup & "' ,RefreshFlag='1',Side='" & Side & "' ,buildtype='" & BuildType & "' where Work_order='" & WO & "' and CompPN='" & Trim(rs!COMPPN) & "' and Slot='" & Slot & "' and LR='" & LR & "' and Machine='" & machine & "' and side='" & Side & "' and JobGroup='" & jobgroup & "'"
         Conn.Execute str
      End If
      rs.MoveNext
Wend
End Function
Public Function ChkUnChkComp(ByVal COMPPN As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
Dim TempHead As String
ChkUnChkComp = False
str = "Select CompHead from QSMS_UnchkComp order by CompHead"
Set rs = Conn.Execute(str)
While Not rs.EOF
      TempHead = Trim(rs!CompHead)
      If UCase(TempHead) = UCase(Mid(Trim(COMPPN), 1, Len(TempHead))) Then
         ChkUnChkComp = True
      End If
      rs.MoveNext
Wend


End Function

Public Function GetJobGroup(ByVal Group As String) As String
Dim str As String
Dim rs As ADODB.Recordset
Dim Jobpn As String
Dim jobgroup As String

Jobpn = ""
''判断是否为EMS机种
    str = "exec QSMS_GetEMSFlag '" & Group & "'"
    Set rs = Conn.Execute(str)
    If Trim(rs("EMSFlag")) = "NONE" Then
        GetJobGroup = ""
        Exit Function
    End If
    
    If Trim(rs("EMSFlag")) <> "Y" Then   ''该机种不是EMS机种时，先判断PCB中是否有大板（PN中包含MB字符）
        str = "Select a.JobPN,b.MB_Rev from QSMS_JobBom a,Sap_Wo_List b where b.[Group]='" & Group & "' and a.Work_Order=b.WO and b.PN like '%MB%'"
        Set rs = Conn.Execute(str)
        If rs.EOF Then   ''当该PCB为纯小板时，选择所有的JobPN 和 MB_Rev 为JobGroup
            str = "Select a.JobPN,b.MB_Rev from QSMS_JobBom a,Sap_Wo_List b where b.[Group]='" & Group & "' and a.Work_Order=b.WO"
            Set rs = Conn.Execute(str)
            If rs.EOF Then
                GetJobGroup = ""
                Exit Function
            Else
                While Not rs.EOF
                    jobgroup = Trim(rs!Jobpn) + "-" + Trim(rs!MB_Rev)
                    Jobpn = Jobpn + "'" + jobgroup + "'" + ","
                    rs.MoveNext
                Wend
            End If
        Else
            While Not rs.EOF    ''当该PCB为大板时，选择大板的JobPN 和 MB_Rev 为JobGroup
                jobgroup = Trim(rs!Jobpn) + "-" + Trim(rs!MB_Rev)
                Jobpn = Jobpn + "'" + jobgroup + "'" + ","
                rs.MoveNext
            Wend
        End If
    Else        ''当该PCB为EMS机种时，选择工单InitAOIFlag标志为“Y”的JobPN 和 MB_Rev 为JobGroup
        str = "Select a.JobPN,b.MB_Rev from QSMS_JobBOM a,Sap_Wo_List B where b.[Group]='" & Group & "' and a.work_order=b.wo and b.InitAOIFlag='Y'"
        Set rs = Conn.Execute(str)
        If rs.EOF Then
                GetJobGroup = ""
                Exit Function
        End If
        
        While Not rs.EOF
              jobgroup = Trim(rs!Jobpn) + "-" + Trim(rs!MB_Rev)
              Jobpn = Jobpn + "'" + jobgroup + "'" + ","
              rs.MoveNext
        Wend
    End If
    
'''功能：取工单组的JobGroup
'''如果InitAOIFlag='Y' 则该工单为大板工单，同一工单Group 里的JobGroup是由大板的JobPN +大板的Rev 构成
''    Str = "Select a.JobPN,b.MB_Rev from QSMS_JobBOM a,Sap_Wo_List B where b.[Group]='" & Group & "' and a.work_order=b.wo and b.InitAOIFlag='Y'"
''    Set Rs = Conn.Execute(Str)
''    If Rs.EOF Then
''            GetJobGroup = ""
''            Exit Function
''    End If
''
''    While Not Rs.EOF
''          JobGroup = Trim(Rs!Jobpn) + "-" + Trim(Rs!Mb_Rev)
''          Jobpn = Jobpn + "'" + JobGroup + "'" + ","
''          Rs.MoveNext
''    Wend
    Jobpn = Mid(Jobpn, 1, Len(Jobpn) - 1)
    Jobpn = "(" + Jobpn + ")"
    GetJobGroup = Jobpn
End Function


Public Function ChkNegativeSide(ByVal Work_Order As String, ByVal COMPPN As String, ByVal UpcompPN As String) As Long
Dim str As String
Dim rs As ADODB.Recordset
'(1) sap bom side---maybe sap bom has two side and each of them separate to mebom (two sides)
'Str = "Select count(*) from Sap_BOM_Chk where work_order='" & Work_Order & "' and ComppN='" & CompPN & "'"
str = "Select count(distinct upcomppn) from Sap_BOM_Chk where work_order='" & Work_Order & "'  and (CompPN='" & COMPPN & "' or CompPN in " & _
       "(select CompPN from QSMS_ReplacePn where version='" & Woinfo.Version & "' and  JobPN  in (select jobpn from qsms_jobbom where work_order='" & Work_Order & "') " & _
       "and ID in (select ID from QSMS_ReplacePN where  JObPN in (select jobpn from qsms_jobbom where work_order='" & Work_Order & "') and CompPN='" & COMPPN & "' and Version='" & Woinfo.Version & "'))) "
Set rs = Conn.Execute(str)
ChkNegativeSide = rs.Fields(0)


End Function
Public Function GetNegativeMESide(ByVal Work_Order As String, ByVal COMPPN As String) As Long
Dim str As String
Dim rs As ADODB.Recordset
''(1) sap bom has one side and mebom has one side
str = "Select distinct JobPn from QSMS_MeBom_Chk where JobPn in (select jobpn from qsms_jobbom where work_order='" & Work_Order & "')" & _
      " and version='" & Woinfo.Version & "' and (CompPN='" & COMPPN & "' or CompPN in " & _
       "(select CompPN from QSMS_ReplacePn where version='" & Woinfo.Version & "' and  JobPN  in (select jobpn from qsms_jobbom where work_order='" & Work_Order & "') " & _
       "and ID in (select ID from QSMS_ReplacePN where  JObPN in (select jobpn from qsms_jobbom where work_order='" & Work_Order & "') and CompPN='" & COMPPN & "' and Version='" & Woinfo.Version & "'))) "
Set rs = Conn.Execute(str)
GetNegativeMESide = rs.RecordCount

End Function

Public Function GetNegativeSapBOMQty(ByVal WO As String, ByVal COMPPN As String, ByVal Item As String, ByVal UpcompPN As String) As Long
Dim str As String
Dim rs As ADODB.Recordset
Dim RsQty As ADODB.Recordset
Dim JobNum As Long
Dim SAPBomQty As Long
Dim tempqty As Long
Dim tempitem As String
tempitem = ""
tempqty = 0
'Str = "Select distinct UpComppN,Item from sap_bom_CHk where Work_Order='" & Wo & "' and CompPN='" & CompPN & "'"
str = "Select distinct UpComppN,Item from sap_bom_CHk where Work_Order='" & WO & "' and  upcomppn='" & UpcompPN & "' and (CompPN='" & COMPPN & "' or CompPN in " & _
       "(select CompPN from QSMS_ReplacePn where version='" & Woinfo.Version & "' and  JobPN  ='" & UpcompPN & "' " & _
       "and ID in (select ID from QSMS_ReplacePN where  JObPN='" & UpcompPN & "' and CompPN='" & COMPPN & "' and Version='" & Woinfo.Version & "'))) "

Set rs = Conn.Execute(str)
JobNum = rs.RecordCount
While Not rs.EOF
          str = "select sum(qty) from sap_bom_CHk where work_order='" & WO & "' and item ='" & rs!Item & "' and upcomppn='" & rs!UpcompPN & "' "
'          Str = "select sum(qty) from sap_bom_CHk where work_order='" & Wo & "' and item ='" & Rs!Item & "' and " & _
                "upcomppn in (select jobpn from qsms_jobbom where work_order='" & Wo & "') "
          Set RsQty = Conn.Execute(str)
          tempqty = tempqty + RsQty.Fields(0)
        
          rs.MoveNext
Wend
  'mark by leimo 20070125 19:26:00
'If JobNum = 0 Then
'    GetNegativeSapBOMQty = 0
'Else
'    GetNegativeSapBOMQty = tempqty / JobNum
'End If
 'add by leimo 20070125 19:26:00
GetNegativeSapBOMQty = tempqty
End Function




Public Function UpdateMachineFlagByWO(ByVal WO As String)
Dim str As String, TransDate As String
Dim rs As ADODB.Recordset
Dim rsMachine As ADODB.Recordset
str = "select getdate()"
Set rs = Conn.Execute(str)
TransDate = Format(rs.Fields(0), "YYYYMMDDHHNNSS")

str = "Select Distinct Machine from QSMS_Wo where Work_Order='" & WO & "'  order by machine"
Set rs = Conn.Execute(str)
While Not rs.EOF
      str = "Select distinct Machine  from QSMS_WO where work_order='" & WO & "' and Machine='" & Trim(rs!machine) & "' and BalanceQty<0"
      Set rsMachine = Conn.Execute(str)
      If rsMachine.EOF Then
         str = "Update QSMS_WO set MachineFinishedFlag='Y' where work_order='" & WO & "' and Machine='" & Trim(rs!machine) & "'"
         Conn.Execute str
      Else
         str = "Update QSMS_WO set MachineFinishedFlag='N' where work_order='" & WO & "' and Machine='" & Trim(rs!machine) & "'"
         Conn.Execute str
      End If
      rs.MoveNext
Wend
str = "Select distinct Machine from QSMS_Wo where work_Order='" & WO & "' and MachineFinishedFlag='N'"
Set rs = Conn.Execute(str)
If rs.EOF Then
    str = "Update QSMS_Wo set WoFinishedFlag='Y' where work_Order='" & WO & "'"
    Conn.Execute str
Else
   str = "Update QSMS_Wo set WoFinishedFlag='N' where work_Order='" & WO & "'"
   Conn.Execute str
End If
 
End Function


Public Function AdjustSapBOM(ByVal WO As String, ByVal RefreshFlag As String)
Dim str As String
Dim rs As ADODB.Recordset, Rs2 As ADODB.Recordset
Dim RsECN As ADODB.Recordset
Dim SapJObPN, MEJObPN, PCN, COMPPN, FuncType, NewCompPN As String
Dim LocNum, SapQty, AdjustQty, AddSeq As Long
Dim s As Integer

s = 0
LocNum = 0
'(0)  adjust sap bom if exit in talbe qsms_updatejobpn
'currently only for OT1,which CPU belong to 21,but in MEBOM it belong to 41
str = "select a.ComPPN,SourceJobPN,DestJobpn from QSMS_UpdateJobPN a,Sap_BOm_Chk B where a.model='" & Woinfo.Model & "' and b.UpCompPN=a.SourceJobPN " & _
      " and b.work_order='" & Woinfo.WO & "'  and a.CompPN=b.CompPN"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   str = "Update Sap_Bom_Chk set UpcomppN='" & Trim(rs!DestJobpn) & "' where UpcompPN='" & Trim(rs!SourceJobpn) & "' " & _
         "and CompPN='" & Trim(rs!COMPPN) & "' and work_order='" & Woinfo.WO & "'"
   Conn.Execute str
End If

''Query all PCN of the JobPN add by Udall 071004
str = "Select a.Qty, b.FuncType, b.PCN,b.JobPN,b.OldCOmpPN,b.NewCompPN,B.LocNum from Sap_Bom_Chk a ,QSMS_DocuComp B where a.work_Order='" & WO & "' and a.upcompPN=b.JobPN and " & _
      "b.NewVersion='" & Woinfo.Version & "' and B.EffectiveFlag='Y' and (B.PCNWO='' or B.PCNWO='" & WO & "') and ECN=''"                                   ''(0001)  (PCNWO='' or PCNWO='" & Woinfo.WO & "')
Set rs = Conn.Execute(str)
While Not rs.EOF
    s = s + 1
    ReDim Preserve JobPCN(s)
    JobPCN(s) = "PCN:" & Trim(rs!PCN) & ",JobPN:" & Trim(rs!Jobpn) & ",OldCompPN:" & Trim(rs!oldCompPN) & ",NewCompPN:" & Trim(rs!NewCompPN) & ",FuncType:" & Trim(rs!FuncType)
    rs.MoveNext
Wend

'(1)check if any  record according to jobpn & CompPN and version and purpose and effectiveflag
'Str = "Select a.Qty, b.FuncType, b.PCN,b.JobPN,b.OldCOmpPN,b.NewCompPN,B.LocNum from Sap_Bom_Chk a ,QSMS_DocuComp B where a.work_Order='" & WO & "' and a.upcompPN=b.JobPN and " & _
'      "b.NewVersion='" & Woinfo.Version & "' and a.CompPN=b.OldCompPN  and B.EffectiveFlag='Y' and ECN='' and b.Purpose='B'"
'Set Rs = Conn.Execute(Str)
'If Rs.EOF Then
'   Exit Function
'End If


'maybe there is more than one PCN for one work order(adjust upcomppn)
'While Not Rs.EOF
'      LocNum = Rs!LocNum
'      SapQty = Rs!Qty
'      SapJObPN = Trim(Rs!Jobpn)
'      PCN = Trim(Rs!PCN)
'      CompPN = Trim(Rs!oldCompPN)
'      AdjustQty = LocNum * Woinfo.WOqty
'      If (SapQty - AdjustQty <= 2) And (SapQty - AdjustQty >= -2) Then
'         Str = "Select JobPN from QSMS_DocuComp where PCN='" & PCN & "' and oldCompPN='" & CompPN & "' and Purpose='B'"
'         Set RsECN = Conn.Execute(Str)
'         MEJObPN = Trim(RsECN!Jobpn)
'         Str = "Insert into Sap_Bom_Chk(Work_Order,MBPN,Item,CompPN,Qty, CompLevel, UpCompPN , TransDateTime )" & _
'               "select work_order ,MBPN,item,CompPN,Qty,CompLevel,'" & MEJObPN & "',TransDateTime from " & _
'               "sap_bom_Chk where work_order='" & Wo & "' and comppn='" & CompPN & "' and UpcompPN='" & SapJObPN & "'"
'         Conn.Execute Str
'
'         Str = "delete from sap_bom_chk where work_order='" & Wo & "' and UpCompPN='" & SapJObPN & "' and CompPN='" & CompPN & "'"
'         Conn.Execute Str
'      Else
'         Str = "Select JobPN from QSMS_DocuComp where PCN='" & PCN & "' and CompPN='" & CompPN & "' and Purpose='B'"
'         Set RsECN = Conn.Execute(Str)
'         MEJObPN = Trim(RsECN!Jobpn)
'         'Str = "Select * from SAP_BOM_Chk where work_order='" & Wo & "' and upcomppn='" & MEJObPN & "' and "
'         Str = "Insert into Sap_Bom_Chk(Work_Order,MBPN,Item,CompPN,Qty, CompLevel, UpCompPN , TransDateTime )" & _
'               "select work_order ,MBPN,item,CompPN," & AdjustQty & ",CompLevel,'" & MEJObPN & "',TransDateTime from " & _
'               "sap_bom_Chk where work_order='" & Wo & "' and comppn='" & CompPN & "' and UpcompPN='" & SapJObPN & "'"
'         Conn.Execute Str
'         Str = "Update  Sap_Bom_Chk set Qty=" & SapQty & "- " & AdjustQty & " where work_order='" & Wo & "' and comppn='" & CompPN & "' and UpcompPN='" & SapJObPN & "'"
'         Conn.Execute Str
'      End If
'      Rs.MoveNext
'
'Wend
'(2)if refresh than check if any delete or update
'If UCase(RefreshFlag) = "Y" Then
'by leimo 20070413 get comp from 内部行文 for Type="ADD"
AddSeq = 0
str = "select PCN,JobPN,NewVersion,NewCompPN,LocNum from QSMS_DocuComp where FuncType='Add' and EffectiveFlag='Y' and ECN='' and NewVersion='" & Woinfo.Version & "' and (PCNWO='' or PCNWO='" & WO & "') and JobPN in (select JobPN from QSMS_JobBOM where Work_Order='" & Woinfo.WO & "')"                ''(0001)  (PCNWO='' or PCNWO='" & Woinfo.WO & "')
Set rs = Conn.Execute(str)
While Not rs.EOF
     AddSeq = AddSeq + 1
     AdjustQty = rs!LocNum * Woinfo.WOqty
      'Update by leimo 20070720---BookMark_1
    'From: The LocNum means add Qty, for example SAPBOM_Qty=10,and need update to 13, then LocNum=3
    'TO:   The LocNum means Total Qty,for example SAPBOM_Qty=10,and need update to 13, then LocNum=13
    str = "Select * from Sap_BOM_Chk where work_order='" & WO & "' and comppn='" & rs!NewCompPN & "' and UpcompPN='" & rs!Jobpn & "'"
    Set Rs2 = Conn.Execute(str)
    If Rs2.EOF Then
          str = "Insert into Sap_Bom_Chk(Work_Order,MBPN,Item,CompPN,Qty, CompLevel, UpCompPN , TransDateTime )" & _
                " select '" & Woinfo.WO & "','" & Woinfo.MBPN & "','Item'+'" & AddSeq & "','" & rs!NewCompPN & "'," & AdjustQty & ",'00','" & rs!Jobpn & "',''"
          Conn.Execute str
    Else
          InsSAP_BOM_FAIL WO, "", "The CompPN:" & Trim(rs!NewCompPN) & " has already existed in the Sap Bom,the PCN:" & Trim(rs!PCN) & " is useless!"     ''Add by Udall 20071010
    End If
'    Else
'         'Update by leimo 20070720---BookMark_1
'         'From: The LocNum means add Qty, for example SAPBOM_Qty=10,and need update to 13, then LocNum=3
'         'TO:   The LocNum means Total Qty,for example SAPBOM_Qty=10,and need update to 13, then LocNum=13
'        ' Str = "Update  Sap_Bom_Chk set Qty=Qty+" & AdjustQty & " where work_order='" & WO & "' and comppn='" & Rs!NewCompPN & "' and UpcompPN='" & Rs!Jobpn & "'"
'         Str = "Update  Sap_Bom_Chk set Qty=" & AdjustQty & " where work_order='" & WO & "' and comppn='" & Rs!NewCompPN & "' and UpcompPN='" & Rs!Jobpn & "'"
'         Conn.Execute Str
'
'    End If
    rs.MoveNext
Wend
   str = "Select a.Qty, b.FuncType, b.PCN,b.JobPN,b.OldCOmpPN,b.NewCompPN,B.LocNum from Sap_Bom_Chk a ,QSMS_DocuComp B where a.work_Order='" & WO & "' and a.upcompPN=b.JobpN and " & _
      "b.NewVersion='" & Woinfo.Version & "' and (B.PCNWO='' or B.PCNWO='" & WO & "' ) and a.CompPN=b.OldCompPN  and B.EffectiveFlag='Y' and ECN='' "                                       ''(0001)  (PCNWO='' or PCNWO='" & Woinfo.WO & "')
   Set rs = Conn.Execute(str)
   If rs.EOF Then
      Exit Function
   End If
   While Not rs.EOF
          LocNum = rs!LocNum
          SapQty = rs!Qty
          SapJObPN = Trim(rs!Jobpn)
          PCN = Trim(rs!PCN)
          COMPPN = Trim(rs!oldCompPN)
          NewCompPN = Trim(rs!NewCompPN)
          FuncType = Trim(rs!FuncType)
          AdjustQty = LocNum * Woinfo.WOqty
          
'          If (SapQty - AdjustQty <= 2) And (SapQty - AdjustQty >= -2) Then
             Select Case UCase(FuncType)
                    Case "DELETE"
                            str = "Delete from Sap_bom_Chk where work_order='" & WO & "' and comppn='" & COMPPN & "' and upcomppn='" & SapJObPN & "'"
                            Conn.Execute str
                    Case "UPDATE"
                            If UCase(Trim(NewCompPN)) = UCase(Trim(COMPPN)) Then  'update Qty
                                 str = "Update  Sap_Bom_Chk set Qty=" & AdjustQty & " where work_order='" & WO & "' and comppn='" & NewCompPN & "' and UpcompPN='" & SapJObPN & "'"
                                 Conn.Execute str
                            Else ' Update Componnet
                                str = "Update  Sap_Bom_Chk set CompPN='" & NewCompPN & "' where work_order='" & WO & "' and comppn='" & COMPPN & "' and UpcompPN='" & SapJObPN & "'"
                                Conn.Execute str
                            End If
'
'                            Str = "Select * from Sap_BOM_Chk where work_order='" & WO & "' and comppn='" & NewCompPN & "' and UpcompPN='" & SapJObPN & "'"
'                            Set Rs2 = Conn.Execute(Str)
'                            If Rs2.EOF Then
'                                Str = "Update  Sap_Bom_Chk set CompPN='" & NewCompPN & "' where work_order='" & WO & "' and comppn='" & CompPN & "' and UpcompPN='" & SapJObPN & "'"
'                                Conn.Execute Str
'                            Else
'                                 'Update by Leimo 20070720, detail see BookMark_1
'                                 'Str = "Update  Sap_Bom_Chk set Qty=Qty+" & AdjustQty & " where work_order='" & WO & "' and comppn='" & NewCompPN & "' and UpcompPN='" & SapJObPN & "'"
'                                 Str = "Update  Sap_Bom_Chk set Qty=" & AdjustQty & " where work_order='" & WO & "' and comppn='" & NewCompPN & "' and UpcompPN='" & SapJObPN & "'"
'                                 Conn.Execute Str
'                                 Str = "delete from Sap_Bom_Chk where work_order='" & WO & "' and comppn='" & CompPN & "' and UpcompPN='" & SapJObPN & "'"
'                                 Conn.Execute Str
'                            End If
                    Case Else
             End Select
'          Else
'             Select Case UCase(FuncType)
'                    Case "DELETE"
'                            Str = "Update  Sap_Bom_Chk set Qty=" & SapQty & "- " & AdjustQty & " where work_order='" & WO & "' and comppn='" & CompPN & "' and UpcompPN='" & SapJObPN & "'"
'                            Conn.Execute Str
'                    Case "UPDATE"
'                             Str = "Select * from Sap_BOM_Chk where work_order='" & WO & "' and comppn='" & NewCompPN & "' and UpcompPN='" & SapJObPN & "'"
'                             Set Rs2 = Conn.Execute(Str)
'                             If Rs2.EOF Then
'                                Str = "Insert into Sap_Bom_Chk(Work_Order,MBPN,Item,CompPN,Qty, CompLevel, UpCompPN , TransDateTime )" & _
'                                      "select work_order ,MBPN,'A'+item,'" & NewCompPN & "'," & AdjustQty & ",CompLevel,UpCompPN,TransDateTime from " & _
'                                       "sap_bom_Chk where work_order='" & WO & "' and comppn='" & CompPN & "' and UpcompPN='" & SapJObPN & "'"
'                                Conn.Execute Str
'                             Else
'                                 'Update by Leimo 20070720, detail see BookMark_1
'                                 'Str = "Update  Sap_Bom_Chk set Qty=Qty+" & AdjustQty & " where work_order='" & WO & "' and comppn='" & NewCompPN & "' and UpcompPN='" & SapJObPN & "'"
'                                 Str = "Update  Sap_Bom_Chk set Qty=" & AdjustQty & " where work_order='" & WO & "' and comppn='" & NewCompPN & "' and UpcompPN='" & SapJObPN & "'"
'                                 Conn.Execute Str
'                             End If
'
'                            Str = "Update  Sap_Bom_Chk set Qty=" & SapQty & "- " & AdjustQty & " where work_order='" & WO & "' and comppn='" & CompPN & "' and UpcompPN='" & SapJObPN & "'"
'                            Conn.Execute Str
'                    Case Else
'            End Select
'          End If
          rs.MoveNext
   Wend
  
'End If



End Function
Public Function AdjustMeBOm(ByVal WO As String, ByVal RefreshFlag As String, ByVal BuildType As String)
Dim str As String
Dim rs As ADODB.Recordset
Dim RsBom As ADODB.Recordset
Dim Item As Long
Dim oldCompPN, NewCompPN, Jobpn, PCN As String
Item = 0
str = "Delete from QSMS_MEBOM_CHk where  " & _
      " version='" & Woinfo.Version & "' and jobpn in (select jobpn from QSMS_JobBOM where work_Order='" & Woinfo.WO & "') " & _
      " and jobgroup in " & Woinfo.jobgroup & " and machine like '" & Woinfo.Line & "%'"
Conn.Execute str

str = "Insert into QSMS_MEBOM_CHk(Machine,JobPN,JobGroup,Item,version,CompPN,LR,Slot,Qty,Status,UID,TransDateTime,BuildType,Side)" & _
      "select Machine,JobPN,JobGroup,'0',version,CompPN,LR,Slot,Qty,Status,UID,TransDateTime,BuildType,Side from QSMS_MEBOm where  " & _
      " version='" & Woinfo.Version & "' and jobpn in (select jobpn from QSMS_JobBOM where work_Order='" & Woinfo.WO & "') " & _
      " and jobgroup in " & Woinfo.jobgroup & " and machine like '" & Woinfo.Line & "%' and BuildType='" & BuildType & "'"
Conn.Execute str



'(1)check if any second source record according to jobpn & CompPN and version and effectiveflag
str = "select distinct a.CompPN, b.NewCompPN,b.jobPN ,b.PCN from QSMS_MEBOM_CHk a ,QSMS_DocuComp B where  a.JobPN=b.JobPN and " & _
      "b.NewVersion=a.Version  and a.CompPN=b.OldCompPN  and B.EffectiveFlag='Y' and (B.PCNWO='' or B.PCNWO='" & WO & "') and ECN='' and FuncType='2ndSource' and " & _
      "a.version='" & Woinfo.Version & "' and a.jobpn in (select jobpn from QSMS_JobBOM where work_Order='" & Woinfo.WO & "') " & _
      " and a.jobgroup in " & Woinfo.jobgroup & " and a.machine like '" & Woinfo.Line & "%'"                        ''(0001)  (PCNWO='' or PCNWO='" & Woinfo.WO & "')
Set rs = Conn.Execute(str)
If rs.EOF Then
   Exit Function
End If

While Not rs.EOF
      NewCompPN = Trim(rs!NewCompPN)
      oldCompPN = Trim(rs!COMPPN)
      Jobpn = Trim(rs!Jobpn)
      PCN = Trim(rs!PCN)
      Item = Item + 1
      str = "Insert into QSMS_MEBOM_CHk(Machine,JobPN,JobGroup,Item,version,CompPN,LR,Slot,NewFlag,Qty,Status,UID,TransDateTime,BuildType,Side)" & _
            "select Machine,JobPN,JobGroup," & Item & ",Version,'" & NewCompPN & "',LR,Slot,'Y',0,Status,UID,TransDateTime,BuildType,Side from QSMS_MEBOM_CHk " & _
            " where JobPN='" & Jobpn & "' and CompPN='" & oldCompPN & "' and " & _
            "version='" & Woinfo.Version & "' and jobpn ='" & Jobpn & "' " & _
            " and jobgroup in " & Woinfo.jobgroup & " and machine like '" & Woinfo.Line & "%'"
      Conn.Execute (str)
      
      'insert into QSMS_MEBOM_PCN
      Call InsertIntoQSMS_MEBOMPCN(oldCompPN, NewCompPN, PCN, Jobpn)

      str = "Update QSMS_MEBOM_CHk set item=" & Item & " where jobpn='" & Jobpn & "' and CompPN='" & oldCompPN & "' and " & _
            "version='" & Woinfo.Version & "' and jobpn ='" & Jobpn & "' " & _
            " and jobgroup in " & Woinfo.jobgroup & " and machine like '" & Woinfo.Line & "%' and BuildType='" & BuildType & "'"
      Conn.Execute str
      rs.MoveNext
Wend


      

End Function
Public Function UpdOT1CPUCompMeBOM()
Dim str As String
Dim rs As ADODB.Recordset
Dim CpuCompPNFromSap As String
''(0)  adjust sap bom if exit in talbe qsms_updatejobpn
''currently only for OT1,which CPU belong to 21,but in MEBOM it belong to 41
'
'Str = "select a.ComPPN,SourceJobPN,DestJobpn from QSMS_UpdateJobPN a,Sap_BOm_Chk B where a.model='" & Woinfo.Model & "' and b.UpCompPN=a.SourceJobPN " & _
'      " and b.work_order='" & Woinfo.Wo & "'  and a.CompPN=b.CompPN"
'Set Rs = Conn.Execute(Str)
'If Not Rs.EOF Then
'   CpuCompPNFromSap = Trim(Rs!CompPN)
'Else
'   Exit Function
'End If
'
'
'Str = "Update QSMS_MEBOM_Chk set CompPN='" & CpuCompPNFromSap & "' from QSMS_UpdateJobPN a,QSMS_MEBOm_Chk B where a.model='" & Woinfo.Model & "' and b.Jobpn=a.DestJobPN " & _
'        "b.version='" & Woinfo.Version & "' and b.jobpn in (select jobpn from QSMS_JobBOM where work_Order='" & Woinfo.Wo & "') " & _
'      " and b.jobgroup in " & Woinfo.JobGroup & " and b.machine like '" & Woinfo.Line & "%' and b.CompPN =a.CompPN"
'Set Rs = Conn.Execute(Str)
'If Not Rs.EOF Then
'
'End If
End Function
Public Function InsertIntoQSMS_MEBOMPCN(ByVal oldCompPN As String, ByVal NewCompPN As String, ByVal PCN As String, ByVal Jobpn As String)
Dim str As String
Dim rs As ADODB.Recordset
Dim RsBom As ADODB.Recordset
'(1)Me bom use old comppn
 str = "select '" & PCN & "','Y',Machine,JobPN,JobGroup,Version,'" & NewCompPN & "',LR,Slot,0,Status,UID,TransDateTime from QSMS_MEBOM_CHk " & _
          " where JobPN='" & Jobpn & "' and CompPN='" & oldCompPN & "' and " & _
          "version='" & Woinfo.Version & "' and jobgroup in " & Woinfo.jobgroup & " and machine like '" & Woinfo.Line & "%'"
Set RsBom = Conn.Execute(str)
While Not RsBom.EOF
    str = "select * from QSMS_MEBOM_PCN where JobPN='" & Jobpn & "' and CompPN='" & NewCompPN & "' and " & _
      "version='" & Woinfo.Version & "' and jobgroup in " & Woinfo.jobgroup & " and machine = '" & Trim(RsBom!machine) & "' " & _
      "and Slot='" & Trim(RsBom!Slot) & "' and LR='" & Trim(RsBom!LR) & "' and PCN='" & PCN & "'"
    Set rs = Conn.Execute(str)
    If rs.EOF Then
        str = "Insert into QSMS_MEBOM_PCN(PCN,EffectiveFlag,Machine,JobPN,JobGroup,version,CompPN,LR,Slot,Qty,Status,UID,TransDateTime)" & _
              "select '" & PCN & "','Y',Machine,JobPN,JobGroup,Version,'" & NewCompPN & "',LR,Slot,0,Status,UID,TransDateTime from QSMS_MEBOM_CHk " & _
              " where JobPN='" & Jobpn & "' and CompPN='" & oldCompPN & "' and machine = '" & Trim(RsBom!machine) & "' and Slot='" & Trim(RsBom!Slot) & "' and LR='" & Trim(RsBom!LR) & "' " & _
              "and version='" & Woinfo.Version & "' and jobgroup in " & Woinfo.jobgroup & " "
        Conn.Execute (str)
    End If
    RsBom.MoveNext
Wend
'(2) MEBOM use new comppn

 str = "select '" & PCN & "','Y',Machine,JobPN,JobGroup,Version,'" & oldCompPN & "',LR,Slot,0,Status,UID,TransDateTime from QSMS_MEBOM_CHk " & _
          " where JobPN='" & Jobpn & "' and CompPN='" & NewCompPN & "' and " & _
          "version='" & Woinfo.Version & "' and jobgroup in " & Woinfo.jobgroup & " and machine like '" & Woinfo.Line & "%'"
Set RsBom = Conn.Execute(str)
While Not RsBom.EOF
    str = "select * from QSMS_MEBOM_PCN where JobPN='" & Jobpn & "' and CompPN='" & NewCompPN & "' and " & _
      "version='" & Woinfo.Version & "' and jobgroup in " & Woinfo.jobgroup & " and machine = '" & Trim(RsBom!machine) & "' " & _
      "and Slot='" & Trim(RsBom!Slot) & "' and LR='" & Trim(RsBom!LR) & "' and PCN='" & PCN & "'"
    Set rs = Conn.Execute(str)
    If rs.EOF Then
        str = "Insert into QSMS_MEBOM_PCN(PCN,EffectiveFlag,Machine,JobPN,JobGroup,version,CompPN,LR,Slot,Qty,Status,UID,TransDateTime)" & _
              "select '" & PCN & "','Y',Machine,JobPN,JobGroup,Version,'" & oldCompPN & "',LR,Slot,0,Status,UID,TransDateTime from QSMS_MEBOM_CHk " & _
              " where JobPN='" & Jobpn & "' and CompPN='" & oldCompPN & "' and machine = '" & Trim(RsBom!machine) & "' and Slot='" & Trim(RsBom!Slot) & "' and LR='" & Trim(RsBom!LR) & "' " & _
              "and version='" & Woinfo.Version & "' and jobgroup in " & Woinfo.jobgroup & " "
        Conn.Execute (str)
    End If
    RsBom.MoveNext
Wend

'Str = "select * from QSMS_MEBOM_PCN where JobPN='" & JobPN & "' and CompPN='" & OldCompPN & "' and " & _
'      "version='" & Woinfo.Version & "' and jobgroup in " & Woinfo.JobGroup & " and machine like '" & Woinfo.Line & "%' and PCN='" & PCN & "'"
'Set Rs = Conn.Execute(Str)
'If Rs.EOF Then
'    Str = "Insert into QSMS_MEBOM_PCN(PCN,EffectiveFlag,Machine,JobPN,JobGroup,version,CompPN,LR,Slot,Qty,Status,UID,TransDateTime)" & _
'          "select '" & PCN & "','Y',Machine,JobPN,JobGroup,Version,'" & OldCompPN & "',LR,Slot,0,Status,UID,TransDateTime from QSMS_MEBOM_CHk " & _
'          " where JobPN='" & JobPN & "' and CompPN='" & NewCompPN & "' and " & _
'          "version='" & Woinfo.Version & "' and jobgroup in " & Woinfo.JobGroup & " and machine like '" & Woinfo.Line & "%'"
'    Conn.Execute (Str)
'End If
End Function


Public Function SecondSource(ByVal COMPPN As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
SecondSource = False
str = "select * from QSMS_MeBOM_CHk where CompPN='" & COMPPN & "' and newflag='Y'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   SecondSource = True
  
End If

End Function
Public Function ChkAddNewCompInMEBOM(ByVal WO As String, ByVal Jobpn As String, ByVal COMPPN) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
ChkAddNewCompInMEBOM = False
str = "Select * from QSMS_DocuComp where JobPN='" & Jobpn & "' and NewCompPN='" & COMPPN & "' and (PCNWO='' or PCNWO='" & WO & "') and EffectiveFlag='Y' and ECN='' and NewVersion='" & Woinfo.Version & "'"        ''(0001)  (PCNWO='' or PCNWO='" & Woinfo.WO & "')
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   ChkAddNewCompInMEBOM = True
End If


End Function

Public Function ChkECNPass(ByVal WO As String) As Boolean
Dim str As String
Dim rs As ADODB.Recordset
ChkECNPass = True
str = "Select a.Qty, b.FuncType, b.PCN,b.JobPN,b.OldCOmpPN,b.NewCompPN,B.LocNum from Sap_Bom_Chk a ,QSMS_DocuComp B where a.work_Order='" & WO & "' and a.upcompPN=b.JobPN and " & _
      "b.NewVersion='" & Woinfo.Version & "' and a.CompPN=b.OldCompPN and (B.PCNWO='' or B.PCNWO='" & WO & "') and B.EffectiveFlag='N' and ECN=''"                  ''(0001)  (PCNWO='' or PCNWO='" & Woinfo.WO & "')
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   ChkECNPass = False
   InsSAP_BOM_FAIL WO, Woinfo.MBPN, "Did not have ECN NO: " & rs!oldCompPN
End If


End Function

'Public Function SendMailDiff(SystemName As String, Subject As String, Body As String, Optional AttachPath1 As String = "D:\CheckBomFail.xls", Optional AttachPath2 As String = "D:\ReleaseInfo.txt")
'    Dim Rs As ADODB.Recordset
'
'    Dim strTo As String
'    Dim strCc As String
'    Dim strBcc As String
'
'    strTo = ""
'    strsql = "select distinct emailaddress from userdetail A,userright B where A.UserName=B.UserName and B.appname='SMT_QSMS'"
'    Set Rs = Conn.Execute(strsql)
'    While Not Rs.EOF
'        strTo = strTo & ";" & Rs!emailAddress
'        Rs.MoveNext
'    Wend
'
'    Body = Replace(Body, "'", " ")
'    strsql = "exec sp_send_cdosysmail_internet '', '" & SystemName & "','" & strTo & "' , '" & strCc & "','" & strBcc & "' , '" & Subject & "' , '" & Body & "','" & AttachPath1 & "' ,'" & AttachPath2 & "'"
'    SMT_Conn.Execute strsql
'    DoEvents
'End Function


Public Function GetSingleSideDispatch(ByVal WO As String, ByVal Line As String) As String
Dim str As String
Dim rs As ADODB.Recordset
Dim SSide As Boolean
Dim CSide As Boolean
SSide = False
CSide = False
GetSingleSideDispatch = "BOTH"
str = "select distinct Machine from QSMS_Wo where work_order ='" & WO & "' and machine like '" & Line & "%' and Machine not like '" & Line & "%others%'"
Set rs = Conn.Execute(str)
While Not rs.EOF
   If UCase(Mid(rs!machine, 2, 1)) = "S" Then
      SSide = True
   End If
   If UCase(Mid(rs!machine, 2, 1)) = "C" Then
      CSide = True
   End If
   rs.MoveNext
Wend
If SSide = True Then
   GetSingleSideDispatch = "S"
End If
If CSide = True Then
   GetSingleSideDispatch = "C"
End If
If SSide = True And CSide = True Then
   GetSingleSideDispatch = "BOTH"
End If

End Function

Public Function SingleSideConfirm(ByVal Line As String, ByVal Side As String, ByVal WO As String)

Dim str As String, machine As String
Dim rs As ADODB.Recordset

Select Case UCase(Side)
       Case "C"
                machine = Trim(Line) & "S"
               ' Str = "Delete QSMS_WO where work_order in " & _
                      "(select wo from sap_wo_list where [group] in (select [group] from sap_wo_list where wo='" & Trim(Wo) & "'))  " & _
                      "and machine  like '" & Machine & "%' and Machine not like '" & Line & "%others%'"
                str = "Delete QSMS_WO where work_order ='" & WO & "' " & _
                       "and machine  like '" & machine & "%' and Machine not like '" & Line & "%others%'"
                Conn.Execute str
                machine = Trim(Line) & "C"
                
                str = "Update QSMS_WO set NeedQty=NeedQty * 2 where  Machine not like '" & Line & "%others%' and work_Order ='" & WO & "'"
                Conn.Execute str
                str = "Update QSMS_Wo set BalanceQty=DispatchQty-NeedQty where work_Order ='" & WO & "'"
                Conn.Execute str
       Case "S"
                machine = Trim(Line) & "C"
                str = "Delete QSMS_WO where work_order ='" & WO & "'" & _
                     "and machine  like '" & machine & "%' and Machine not like '" & Line & "%others%'"
                Conn.Execute str
                machine = Trim(Line) & "S"
                
                str = "Update QSMS_WO set NeedQty=NeedQty * 2 where  Machine not like '" & Line & "%others%' and work_Order ='" & WO & "'"
                Conn.Execute str
                str = "Update QSMS_Wo set BalanceQty=DispatchQty-NeedQty where work_Order ='" & WO & "'"
                Conn.Execute str
      Case Else

 
End Select

End Function
Public Sub Copy2Excel(ByVal rst As ADODB.Recordset)
 Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object
 Dim recArray As Variant
 Dim strDB As String
 Dim fldCount As Long
 Dim recCount As Long
 Dim iCol As Long
 Dim iRow As Long
 Dim strFileName, Trans_Date As String
 On Error GoTo errHandler


    Set xlApp = CreateObject("Excel.application")
    xlApp.Visible = True
    Set xlsBook = xlApp.Workbooks.Open(App.Path & "\QSMS_ReturnDID_Summary.XLS")
    Set xlWs = xlsBook.Worksheets(1)
    xlApp.DisplayAlerts = False
    xlApp.UserControl = True
    Set xlWs = Nothing
    Set xlWs = xlsBook.Worksheets("QSMS_ReturnDID_Report")
    xlWs.Cells(1, 1) = "Today:" & Format(Now, "MM/DD")
'    Set xlApp = CreateObject("Excel.Application")
'    Set xlsBook = xlApp.Workbooks.Add
'    'important for disabled alerts
'    xlApp.DisplayAlerts = False
'    Set xlWS = xlApp.Worksheets(1)
  
    
    ' Copy field names to the fiRst row of the worksheet
    fldCount = rst.Fields.Count

    For iCol = 1 To fldCount
        xlWs.Cells(2, iCol).Select
        xlApp.Selection.Interior.ColorIndex = 6
        xlWs.Cells(2, iCol).Value = rst.Fields(iCol - 1).Name
        xlApp.Selection.HorizontalAlignment = xlCenter
        xlApp.Selection.VerticalAlignment = xlCenter
    Next
    
    xlWs.Cells(3, 1).CopyFromRecordset rst
    xlApp.Rows("3:2").Select
    xlApp.ActiveWindow.FreezePanes = True
    xlApp.ActiveWindow.SmallScroll Down:=0

    ' Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlApp.Visible = True
    
    
    ' Close ADO objects
    rst.Close
    Set rst = Nothing
    Trans_Date = Format(Now, "YYYYMMDD")

    Set xlApp = Nothing
    Set xlsBook = Nothing

    Exit Sub

errHandler:
    MsgBox ("CopyToExcel, " & Err.Description & "; please contact QMS!")
End Sub

Public Sub CopyToExcel(ByVal rst As ADODB.Recordset)
 Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object
 Dim recArray As Variant
 Dim strDB As String
 Dim fldCount As Long
 Dim recCount As Long
 Dim iCol As Long
 Dim iRow As Long
 Dim strFileName, Trans_Date As String
 On Error GoTo errHandler
    
 If chkDomain = "N" Then ''1165
    Call ExportToHtml(rst)

 Else
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add
    'important for disabled alerts
    xlApp.DisplayAlerts = False
    Set xlWs = xlApp.Worksheets(1)
  
    xlApp.UserControl = True
    
    ' Copy field names to the fiRst row of the worksheet
    fldCount = rst.Fields.Count

    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Select
        xlApp.Selection.Interior.ColorIndex = 6
        xlWs.Cells(1, iCol).Value = rst.Fields(iCol - 1).Name
        xlApp.Selection.HorizontalAlignment = xlCenter
        xlApp.Selection.VerticalAlignment = xlCenter
    Next
    
    xlWs.Cells(2, 1).CopyFromRecordset rst
    xlApp.Rows("2:2").Select
    xlApp.ActiveWindow.FreezePanes = True
    xlApp.ActiveWindow.SmallScroll Down:=0

    ' Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlApp.Visible = True
    
    
    ' Close ADO objects
    rst.Close
    Set rst = Nothing
    Trans_Date = Format(Now, "YYYYMMDD")

    Set xlApp = Nothing
    Set xlsBook = Nothing
 
 End If
 
    Exit Sub

errHandler:
    MsgBox ("CopyToExcel, " & Err.Description & "; please contact QMS!")
End Sub

Public Function ChkForbiddenPN(WO As String, COMPPN As String) As Boolean

Dim str As String
Dim rs As ADODB.Recordset
Dim TempJObGroup As String

On Error GoTo ErrHandle:

    str = "exec QSMS_ChkForbiddenPN '" & WO & "','" & COMPPN & "'"
    Set rs = Conn.Execute(str)
    If Not rs.EOF Then
        If UCase(rs.Fields(0)) = "PASS" Then
            ChkForbiddenPN = False
        Else
            ChkForbiddenPN = True
            Exit Function
        End If
    End If
    
    Exit Function
ErrHandle:
    ChkForbiddenPN = False
    MsgBox Err.Description & " Please Call QMS "

End Function

Public Function ChkBuildType() As Boolean

Dim str As String
Dim rs As ADODB.Recordset
Dim TempJObGroup As String

On Error GoTo ErrHandle:
TempJObGroup = Replace(Woinfo.jobgroup, "'", "")
TempJObGroup = Replace(TempJObGroup, "(", "")
TempJObGroup = Replace(TempJObGroup, ")", "")
ChkBuildType = True

'(1)Check Build Type and side
str = "exec QSMS_ChkBuildType '" & Woinfo.WO & "','" & TempJObGroup & "'"
'Str = "exec QSMSChkCloseWOByManual '" & Wo & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   If UCase(rs.Fields(0)) = "PASS" Then
   Else
       ChkBuildType = False
       Exit Function
   End If
End If

Exit Function
ErrHandle:
         ChkBuildType = False
         MsgBox Err.Description & " Please Call QMS "

End Function
Public Function GetNegativeNeedQty(ByVal COMPPN As String, ByVal WO As String, ByVal NeedQty As Long, ByVal Jobpn As String) As Long
Dim str As String
Dim rs As ADODB.Recordset
Dim NeedTotalQty, SAPTotalQty As Long
Dim Item As String
Item = 0
'当除不断的时候， SAP 的Total Qty可能会大于QSMS_WO 的Need Qty
GetNegativeNeedQty = NeedQty
If Woinfo.CrashFlag = True Then
    str = "select ISNULL(sum(NeedQty),0) as NeedQty,Item from qsms_wo where work_order='" & WO & "' and Comppn='" & COMPPN & "' group by item"
    Set rs = Conn.Execute(str)
    If Not rs.EOF Then
        NeedTotalQty = NeedQty + rs!NeedQty
        Item = Trim(rs!Item)
        'NeedTotalQty = Rs!NeedQty
    End If
    
    'Str = "select isnull(sum(Qty),0) as SAPQty from SAP_BOM where work_order='" & WO & "' and Comppn='" & CompPN & "'"
    If Item <> 0 Then
         str = "select isnull(sum(Qty),0) as SAPQty from SAP_BOM  where work_order='" & WO & "' and comppn in (select comppn from qsms_wo where work_order='" & WO & "' and Item='" & Item & "')"
    Else
         str = "select isnull(sum(Qty),0) as SAPQty from SAP_BOM  where work_order='" & WO & "' and Comppn='" & COMPPN & "'"
    End If
    Set rs = Conn.Execute(str)
    If Not rs.EOF Then
        SAPTotalQty = rs!SapQty
    End If
    If NeedTotalQty >= SAPTotalQty Then
       str = "select ISNULL(sum(NeedQty),0) as NeedQty  from qsms_wo where work_order='" & WO & "' and Comppn='" & COMPPN & "' and Jobpn='" & Jobpn & "'"
       Set rs = Conn.Execute(str)
       If Not rs.EOF Then
          If (NeedQty - rs!NeedQty <= 2) And (NeedQty - rs!NeedQty >= -2) Then
             GetNegativeNeedQty = rs!NeedQty
             Exit Function
          End If
       End If
      
    End If
    If (NeedTotalQty - SAPTotalQty <= 2) And (NeedTotalQty - SAPTotalQty >= -2) Then
       GetNegativeNeedQty = SAPTotalQty - NeedTotalQty + NeedQty
    End If
End If



End Function
Public Function GetNeedQtyforWOMultiLine(ByVal BaseQty As Long, ByVal machine As String, ByVal Side As String) As Long
Dim str As String
Dim rs As ADODB.Recordset

GetNeedQtyforWOMultiLine = 0
str = "Select Qty from WO_MultiLine where WO='" & Woinfo.WO & "' and Line='" & Mid(machine, 1, 1) & "' and side='" & Side & "'"
Set rs = Conn.Execute(str)
If Not rs.EOF Then
  
    GetNeedQtyforWOMultiLine = CDbl(BaseQty) / Woinfo.CombineQty * rs!Qty
    
Else
    Exit Function
End If


If InStr(1, UCase(machine), "OTHERS") > 0 Then
     GetNeedQtyforWOMultiLine = BaseQty * rs!Qty / Woinfo.SourceQty
Else
'如果是单报板，因为MEＢＯＭ用的是Ｌａｙｏｕｔ的数量，而实际工单的combine Qty可能<Layout　Qty
     If Woinfo.CrashFlag = True Then
        GetNeedQtyforWOMultiLine = GetNeedQtyforWOMultiLine * Woinfo.DestQty / Woinfo.SourceQty
     End If
End If

End Function

'GetMEBOMQtyforWOMultiLine @Wo varchar(50),@JobGroup varchar(50),@CompPN varchar(50),@UpCompPN varchar(50) ,@Line varchar(50) as
Public Function GetMEBOMQtyforWOMultiLine(ByVal COMPPN As String, ByVal UpcompPN As String) As Long
Dim str As String
Dim rs As ADODB.Recordset
Dim TempJObGroup As String

TempJObGroup = Replace(Woinfo.jobgroup, "'", "")
TempJObGroup = Replace(TempJObGroup, "(", "")
TempJObGroup = Replace(TempJObGroup, ")", "")


GetMEBOMQtyforWOMultiLine = 0
str = "Exec GetMEBOMQtyforWOMultiLine '" & Woinfo.WO & "','" & TempJObGroup & "','" & COMPPN & "','" & UpcompPN & "','" & Woinfo.Line & "'"
Set rs = Conn.Execute(str)

If Not rs.EOF Then
    GetMEBOMQtyforWOMultiLine = rs.Fields(0)
Else
    Exit Function
End If

End Function


Public Sub ToExcel(strsql As String)
 Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object
 Dim recArray As Variant
 Dim strDB As String
 Dim fldCount As Long
 Dim recCount As Long
 Dim iCol As Long
 Dim iRow As Long
 Dim strFileName, Trans_Date As String
 Dim i As Integer
 Dim rst As New ADODB.Recordset
 
 On Error GoTo errHandler
 
    If chkDomain = "N" Then ''1167
        Set rst = Conn.Execute(strsql)
        Call ExportToHtml(rst)
        Exit Sub
    End If

    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add
    Set xlWs = xlApp.Worksheets.Add     ' win10默认只有一个sheet1如果有sheet2的操作会报错所以这里添加一个sheet
    
    xlApp.Worksheets(1).Activate
    
    For i = 1 To 2
        Set xlWs = xlApp.Worksheets(i)
        Set rst = Conn.Execute(strsql)
         
        
            If i = 2 Then
                Set rst = rst.NextRecordset
                xlApp.Worksheets(2).Activate
            End If
            
            If rst.EOF = False Then
                xlApp.UserControl = True
            
                fldCount = rst.Fields.Count
            
                For iCol = 1 To fldCount
                    xlWs.Cells(1, iCol).Select
                    xlApp.Selection.Interior.ColorIndex = 6
                    xlWs.Cells(1, iCol).Value = rst.Fields(iCol - 1).Name
                    xlApp.Selection.HorizontalAlignment = xlCenter
                    xlApp.Selection.VerticalAlignment = xlCenter
                Next
            
                xlWs.Cells(2, 1).CopyFromRecordset rst
            
                xlApp.Rows("2:2").Select
                xlApp.ActiveWindow.FreezePanes = True
                xlApp.ActiveWindow.SmallScroll Down:=0
        
                ' Auto-fit the column widths and row heights
                xlApp.Selection.CurrentRegion.Columns.AutoFit
                xlApp.Selection.CurrentRegion.Rows.AutoFit
            Else
                MsgBox ("NO DATA !")
            End If
            Set rst = Nothing
    Next

    xlApp.Visible = True
    Trans_Date = Format(Now, "YYYYMMDD")

    Set xlApp = Nothing
    Set xlsBook = Nothing

    Exit Sub

errHandler:
    MsgBox ("ToExcel, " & Err.Description & "; please contact QMS!")
End Sub
Public Sub CopyToTemplateExcel(ByVal rst As ADODB.Recordset, strRptFileName As String) '''1104
 Dim xlApp As Excel.Application
 Dim xlsBook As Excel.Workbook
 Dim xlWs As Object
 Dim recArray As Variant
 Dim strDB As String
 Dim fldCount As Long
 Dim recCount As Long
 Dim iCol As Long
 Dim iRow As Long
 Dim xlSheet As New Excel.Worksheet
 Dim strFileName, Trans_Date As String
 On Error GoTo errHandler
    
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlsBook = xlApp.Workbooks.Add(App.Path + "\Template\" & strRptFileName & ".xls")
    'Set xlsBook = xlApp.Workbooks.Add
    'important for disabled alerts
    xlApp.DisplayAlerts = False
    Set xlWs = xlApp.Worksheets(1)
   Set xlSheet = xlsBook.Sheets("IC_BURN")
       xlSheet.Activate
    xlApp.UserControl = True
 
    ' Copy field names to the fiRst row of the worksheet
    fldCount = rst.Fields.Count

    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Select
        xlApp.Selection.Interior.ColorIndex = 6
      '  xlWs.Cells(1, iCol).Value = rst.Fields(iCol - 1).Name
        xlApp.Selection.HorizontalAlignment = xlCenter
        xlApp.Selection.VerticalAlignment = xlCenter
    Next
    
    xlWs.Cells(3, 1).CopyFromRecordset rst
    xlApp.Rows("3:2").Select
    xlApp.ActiveWindow.FreezePanes = True
    xlApp.ActiveWindow.SmallScroll Down:=0

    ' Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlApp.Visible = True
    
    
    ' Close ADO objects
    rst.Close
    Set rst = Nothing
    Trans_Date = Format(Now, "YYYYMMDD")

    Set xlApp = Nothing
    Set xlsBook = Nothing

    Exit Sub

errHandler:
    MsgBox ("CopyToExcel, " & Err.Description & "; please contact QMS!")
End Sub

Public Function GetCheckBomData(ByVal Work_Order As String, ByVal g_userName As String, ByVal DualModel As String)            ''(0019)
Dim str As String, step As Integer
Dim rs As ADODB.Recordset
Dim BuildType As String
Dim strErrDesc As String
On Error GoTo errHandler
step = 0
If Trim(Work_Order) = "" Then
   MsgBox "Please check the WO"
   Exit Function
Else
    str = "Exec QSMS_RegisterCheckBOM '" & Trim(Work_Order) & "','0',''" '(0008)
    Set rs = Conn.Execute(str)
    If Not rs.EOF Then
        If rs("rtnCode") = 0 Then
            MsgBox ("Now,SomeBody is doing CheckBOM in computer " & Trim(rs("hostname")) & ",The system don't allow more than one person to do CheckBOM")
            Exit Function
        End If
    End If
End If
step = 1
'delete old failure log
str = "delete from Sap_BOM_Fail  where Work_Order ='" & Trim(Work_Order) & "' "
Conn.Execute (str)
str = "Insert into QSMS_LOG(system_name,event_no,DID,user_name,ReturnQty,trans_date) values ('SMT_QSMS','CheckBOM','" & Work_Order & "','" & g_userName & "','1',dbo.FormatDate(getdate(),'yyyymmddhhnnss'))"
Conn.Execute (str) '(0003)
'再次检查CheckBOM的工单是否已经注册
str = "select 0 from QSMS_CheckBOM where WorkOrder='" & Trim(Work_Order) & "' "
Set rs = Conn.Execute(str)
If rs.EOF Then
    str = "Exec QSMS_RegisterCheckBOM '" & Trim(Work_Order) & "','0',''" '(0008)
    Set rs = Conn.Execute(str)
    If Not rs.EOF Then
        If rs("rtnCode") = 0 Then
            MsgBox ("Now,SomeBody is doing CheckBOM in computer " & Trim(rs("hostname")) & ",The system don't allow more than one person to do CheckBOM")
            Exit Function
        End If
    End If
End If
step = 2
'Check BOM
'get Built Type
str = "select BuildType from sap_wo_list where Wo='" & Trim(Work_Order) & "'"
Set rs = Conn.Execute(str)
If rs.EOF Then
    MsgBox "PMC didn't release the WO,Please Check"
    str = "Insert into QSMS_LOG(system_name,event_no,DID,user_name,ReturnQty,trans_date) values ('SMT_QSMS','CheckBOM','" & Work_Order & "','" & g_userName & "','4',dbo.FormatDate(getdate(),'yyyymmddhhnnss'))"
    Conn.Execute (str) '(0003)
    Exit Function
Else
    BuildType = Trim(rs!BuildType)
    If BuildType <> "1" And BuildType <> "2" And BuildType <> "3" And BuildType <> "4" Then
       MsgBox "BuildType Error,Please call QMS"
        str = "Insert into QSMS_LOG(system_name,event_no,DID,user_name,ReturnQty,trans_date) values ('SMT_QSMS','CheckBOM','" & Work_Order & "','" & g_userName & "','4',dbo.FormatDate(getdate(),'yyyymmddhhnnss'))"
        Conn.Execute (str) '(0003)
       Exit Function
    End If
End If

'20070724 Denver  Update SAP_WO_List firstPass info
Dim chkBOMPass As String

If DualModel = "N" Then    ''(1179)
    str = "Exec QSMS_CheckBomSP '" & Trim(Work_Order) & "','N','" & Trim(BuildType) & "'"               ''(0019)
Else
    str = "Exec QSMS_CheckBomSP_Dual '" & Trim(Work_Order) & "','N','" & Trim(BuildType) & "'"          ''(1179)
End If
Set rs = Conn.Execute(str)
If rs.EOF = False Then
    If Woinfo.Negative = True And Woinfo.Pilot = "NEW" Then
       MsgBox "Check BOM OK"
       str = "Exec QSMS_RegisterCheckBOM '" & Trim(Work_Order) & "','1',''" '(0008)
        Set rs = Conn.Execute(str)
        If Not rs.EOF Then
            If rs("rtnCode") = 0 Then
                MsgBox ("Clear the WorkOder from  XL_CheckBOM fail,Please check XL_CheckBOM")
            End If
        End If
            '--(0018)
        str = "Insert into QSMS_LOG(system_name,event_no,DID,user_name,ReturnQty,trans_date) values ('SMT_QSMS','CheckBOM','" & Work_Order & "','" & g_userName & "','3',dbo.FormatDate(getdate(),'yyyymmddhhnnss'))"
        Conn.Execute (str) '(0003)
        str = "update qsms_error_log set col1=dbo.formatdate(getdate(),'yyyymmddhhnnss') where subid='" & Trim(Work_Order) & "' and appname='SMT_QSMS' and subfunction='ReplacePN' and col1<>''"
        Conn.Execute (str)
        str = "insert into qsms_log(system_name,event_no,DID,user_name,ReturnQty,trans_date) values('SMT_QSMS','CheckBOMResult','" & Work_Order & "','" & g_userName & "','0',dbo.FormatDate(getdate(),'yyyymmddhhnnss'))"
        Conn.Execute (str) '0073
       Exit Function
    End If
    MsgBox "Check bom fail"
    chkBOMPass = "N"
Else
    chkBOMPass = "Y"
End If

str = "Select Site from Site"
Set rs = Conn.Execute(str)
If rs.EOF = False Then
    If rs.Fields(0) = "ES" Or rs.Fields(0) = "ESBU" Or rs.Fields(0) = "CC" Then
        str = "Exec UpdSAP_firstPass " & sq(Work_Order) & "," & sq(chkBOMPass)
        Conn.Execute str
    End If
End If
'20070724 End

'--(0018)
'--0073
If chkBOMPass = "Y" Then
    str = "update qsms_error_log set col1=dbo.formatdate(getdate(),'yyyymmddhhnnss') where subid='" & Trim(Work_Order) & "' and appname='SMT_QSMS' and subfunction='ReplacePN' and col1=''"
    Conn.Execute (str)
    str = "insert into qsms_log(system_name,event_no,DID,user_name,ReturnQty,trans_date) values('SMT_QSMS','CheckBOMResult','" & Work_Order & "','" & g_userName & "','0',dbo.FormatDate(getdate(),'yyyymmddhhnnss'))"
    Conn.Execute (str)
Else
    str = "insert into qsms_log(system_name,event_no,DID,user_name,ReturnQty,trans_date) values('SMT_QSMS','CheckBOMResult','" & Work_Order & "','" & g_userName & "','-1',dbo.FormatDate(getdate(),'yyyymmddhhnnss'))"
    Conn.Execute (str)
End If

step = 3
str = "select *  from Sap_BOM_Fail  where Work_Order ='" & Trim(Work_Order) & "' "
Set rs = Conn.Execute(str)
If Not rs.EOF Then
   Call CopyToExcel(rs)
Else
    str = "EXEC QSMS_ReCountDispatchQty '" & Trim(Work_Order) & "','1' "
    Conn.Execute (str)
    MsgBox "Check BOM OK"
End If
str = "Exec QSMS_RegisterCheckBOM '" & Trim(Work_Order) & "','1',''" '(0008)
Set rs = Conn.Execute(str)
If Not rs.EOF Then
    If rs("rtnCode") = 0 Then
        MsgBox ("Clear the WorkOder from  XL_CheckBOM fail,Please check XL_CheckBOM")
    End If
End If
'delete old failure log
str = "delete from Sap_BOM_Fail  where Work_Order ='" & Trim(Work_Order) & "' "
Conn.Execute (str)

str = "Insert into QSMS_LOG(system_name,event_no,DID,user_name,ReturnQty,trans_date) values ('SMT_QSMS','CheckBOM','" & Work_Order & "','" & g_userName & "','2',dbo.FormatDate(getdate(),'yyyymmddhhnnss'))"
Conn.Execute (str) '(0003)


Exit Function

errHandler:
    strErrDesc = Err.Description  'save error description into table-- add by kane 090804
    MsgBox ("GetCheckBomData, Step:" & CStr(step) & ", " & strErrDesc)
'    str = "Insert into QSMS_LOG(system_name,event_no,DID,user_name,ReturnQty,trans_date) values ('SMT_QSMS','CheckBOM','" & Work_Order & ";" & Left(strErrDesc, 150) & "','" & g_userName & "','5',dbo.FormatDate(getdate(),'yyyymmddhhnnss'))"
'    Conn.Execute (str) '(0003)
End Function

