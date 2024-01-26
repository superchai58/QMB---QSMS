Attribute VB_Name = "ChkBom"
Type WOBasic
    WO As String
    MBPN As String
    Version As String
    WOqty As Integer
    CombineQty As Integer
    Line As String
    JobPn41 As String
    JobPn51 As String
    Group As String
    GroupWoQty As Integer
End Type
Public Woinfo As WOBasic
Type MeBomBasic
     Machine As String
     JobPN As String
     Slot As String
     Qty As Integer
     Version As String
End Type
Public MeBomINfo(2) As MeBomBasic
Public ReplaceItem As Integer

Public Function GetWoinfo(ByVal WO As String)
Dim Str As String
Dim Rs As ADODB.Recordset
Str = "select WO,PN,Mb_Rev,line,Qty,CombineQty from Sap_WO_List where WO='" & WO & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
    With Woinfo
         .CombineQty = CLng(Trim(Rs!CombineQty))
         .WO = Trim(Rs!WO)
         .MBPN = Trim(Rs!PN)
         .Version = Trim(Rs!Mb_Rev)
         .WOqty = CLng(Trim(Rs!Qty))
         .Line = Trim(Rs!Line)
    End With
Else
    InsSAP_BOM_FAIL WO, "", "NO Sap Wo List: "
End If
Str = "Select JobPN from QSMS_JOBBom where Work_Order='" & WO & "'"
Set Rs = Conn.Execute(Str)
While Not Rs.EOF
      If Mid(Rs!JobPN, 1, 2) = "41" Then
         Woinfo.JobPn41 = Trim(Rs!JobPN)
      Else
         Woinfo.JobPn51 = Trim(Rs!JobPN)
      End If
      Rs.MoveNext
Wend

Str = "select [Group] from Sap_Wo_List where WO='" & WO & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
  Woinfo.Group = Trim(Rs![Group])
  Str = "Select Count(*) from Sap_Wo_List where [Group]='" & Trim(Rs![Group]) & " '"
  Set Rs = Conn.Execute(Str)
  Woinfo.GroupWoQty = Rs.Fields(0)
Else
   GetSmallBoardGroup = ""
End If

End Function
Public Function GetMEBOMQty(ByVal WO As String, ByVal CompPN As String, ByVal CompLevel As String, ByVal Line As String) As Integer
'  CompLevel   JobPN             Side
'   01         41*********       Component side
'   02         51*********       solder side
'   00         others,which by manual insteda of  by machine
Dim Str As String
Dim Rs As ADODB.Recordset
'Dim CompRs As ADODB.Recordset
Dim TempQty As Integer
Dim JobPN As String
Dim Machine As String

Dim I As Integer
Select Case CompLevel
       Case "01"
            JobPN = "41"
            Machine = Line + "C"
            
       Case "02"
            JobPN = "51"
            Machine = Line + "S"
       Case "00"
            Str = "select JobPN from QSMS_ReplacePN where version='" & Woinfo.Version & "' and JObPN in (select JObPN from QSMS_JobBom where Work_Order='" & WO & "' and jobpn like '" & JobPN & "%') and CompPN='" & CompPN & "' "
            Set Rs = Conn.Execute(Str)
            If Not Rs.EOF Then
                JobPN = Mid(Trim(Rs!JobPN), 1, 2)
            End If
           
            Machine = "Others"
            
End Select



TempQty = 0

'Str = "Select sum(A.Qty) as Qty from QSMS_MeBOM a,QSMS_JobBOM b where b.Work_Order='" & Wo & "' and a.JobPN=b.JobPN  and a.jobpn like '" & JobPn & "%' " & _
       "and (CompPN='" & CompPN & "' or CompPN in (select CompPn from CompRs))"
Str = "Select sum(Qty) as Qty from QSMS_MeBOM  where version='" & Woinfo.Version & "' and   jobpn in (select JobPN from QSMS_JObBOm where Work_Order='" & WO & "') and jobpn like '" & JobPN & "%' " & _
       "and (CompPN='" & CompPN & "' or CompPN in " & _
       "(select CompPN from QSMS_ReplacePn where version='" & Woinfo.Version & "' and JobPN in (select JobPN from QSMS_JobBom where work_order='" & WO & "' and jobpn like '" & JobPN & "%') " & _
       "and ID in (select ID from QSMS_ReplacePN where version='" & Woinfo.Version & "' and JObPN in (select JObPN from QSMS_JobBom where Work_Order='" & WO & "' and jobpn like '" & JobPN & "%') and CompPN='" & CompPN & "'))) "
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   If Trim(Rs!Qty & vbNullString) <> "" Then
      GetMEBOMQty = Rs!Qty
   End If
End If



End Function


Public Function CheckBom(ByVal WO As String) As Boolean
Dim Str As String
Dim Rs As ADODB.Recordset
Dim Errmsg As String
CheckBom = True
Call GetWoinfo(WO)

Str = "Delete from Sap_Bom_Fail where Work_Order in (select wo from sap_wo_list where [Group]='" & Woinfo.Group & "')"
Conn.Execute Str

Str = "select * from sap_Bom where Work_Order='" & WO & "'"
Set Rs = Conn.Execute(Str)
If Rs.EOF Then
   InsSAP_BOM_FAIL WO, "", "NO SAP BOM : "
   CheckBom = False
End If


'(1) check sap bom if lost in Me bom
'Str = "select Work_order,CompPN from sap_bom where Work_order ='" & WO & "'  and CompPn not in " & _
     "(select a.CompPn from QSMS_MeBom a,QSMS_JobBOM b,Sap_WO_List c where b.Work_Order='" & WO & "' and b.work_order=c.Wo and c.MB_Rev=a.version and b.JObPN=a.JObPN )"
Str = "select Work_order,CompPN from sap_bom where Work_order in (Select Wo from sap_wo_list where [Group]='" & Trim(Woinfo.Group) & "')  and CompPn not in " & _
     "(select a.CompPn from QSMS_MeBom a,QSMS_JobBOM b,Sap_WO_List c where b.Work_Order='" & WO & "' and b.work_order=c.Wo and c.MB_Rev=a.version and b.JObPN=a.JObPN )"
Set Rs = Conn.Execute(Str)
While Not Rs.EOF
        If ChkReplacePN(Rs!CompPN, WO, "SAP_BOM") = False Then
           InsSAP_BOM_FAIL WO, MBPN, "Lost in ME BOM: " & Rs!CompPN
           CheckBom = False
        End If
        Rs.MoveNext
Wend
'(2) check ME bom if lost in Sap Bom
'Str = "select a.CompPn from QSMS_MeBom a,QSMS_JobBOM b,Sap_WO_List c where b.Work_Order='" & WO & "' and b.work_order=c.Wo and c.MB_Rev=a.version and b.JObPN=a.JObPN  and a.CompPN not in " & _
      " (select Comppn from sap_bom where Work_order ='" & WO & "')"
Str = "select a.CompPn from QSMS_MeBom a,QSMS_JobBOM b,Sap_WO_List c where b.Work_Order='" & WO & "' and b.work_order=c.Wo and c.MB_Rev=a.version and b.JObPN=a.JObPN  and a.CompPN not in " & _
      " (select Comppn from sap_bom where Work_order  in (Select Wo from sap_wo_list where [Group]='" & Trim(Woinfo.Group) & "'))"
Set Rs = Conn.Execute(Str)
While Not Rs.EOF
     If ChkReplacePN(Rs!CompPN, WO, "ME_BOM") = False Then
        InsSAP_BOM_FAIL WO, MBPN, "Lost in SAPBOM: " & Rs!CompPN
        CheckBom = False
     End If
     Rs.MoveNext
Wend
'(3)check Comp Qty

'Str = "select Work_Order,MBPN,Item,CompPN,Qty,CompLevel from Sap_BOm where Work_Order='" & WO & "' "
Str = "select Work_Order,MBPN,Item,CompPN,Qty,CompLevel from Sap_BOm where Work_Order in (Select Wo from sap_wo_list where [Group]='" & Trim(Woinfo.Group) & "')"
Set Rs = Conn.Execute(Str)
While Not Rs.EOF
      If ChkCompQty(Rs!Work_Order, Trim(Rs!MBPN), Rs!CompPN, Woinfo.WOqty, Trim(Rs!Item), Trim(Rs!CompLevel)) = False Then
          'InsSAP_BOM_FAIL Wo, Woinfo.MBPN, "Comp Qty does not match: " & Rs!CompPN
          CheckBom = False
      End If
      Rs.MoveNext
      
Wend
If CheckBom = True Then
   Call InsertToQSMS_WO(WO)
End If
End Function
Public Function ChkCompQty(ByVal WO As String, MBPN As String, CompPN As String, ByVal WOqty As Integer, ByVal Item As String, ByVal CompLevel As String) As Boolean
Dim Str As String
Dim Rs As ADODB.Recordset
Dim SAPBomQty, MEBomQty As Long
Dim I As Integer
Dim TempCompLevel As String
TempCompLevel = CompLevel

'  CompLevel   JobPN             Side
'   01         41*********       Component side
'   02         51*********       solder side
'   00         others,which by manual insteda of  by machine
'(1) sum Sap bom  Qty accoding to WO & Item
Str = "select sum(qty) from sap_bom where work_order='" & WO & "' and item='" & Item & "' and CompLevel='" & CompLevel & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   SAPBomQty = Rs.Fields(0)
End If

If Trim(WO) = Trim(Woinfo.WO) Then
   SAPBomQty = SAPBomQty + GetSmallBoardSapQty(CompPN, CompLevel)
Else
   If ChkCompBelongMBBom(CompPN, Woinfo.MBPN, MBPN) = True Then
      ChkCompQty = True
      Exit Function
   Else
       TempCompLevel = GetSBCompLevel(Trim(CompPN), Trim(MBPN), Trim(Woinfo.MBPN))
   End If
End If

'(2) Get ME BOm Qty


MEBomQty = CDbl(GetMEBOMQty(Trim(WO), Trim(CompPN), Trim(TempCompLevel), Woinfo.Line)) * WOqty
Select Case UCase(Mid(MBPN, 3, 3))
       Case "VC1", "VC2", "K2M"
             MEBomQty = MEBomQty / 2
      Case Else
         
End Select
MEBomQty = MEBomQty / Woinfo.CombineQty
If SAPBomQty = MEBomQty Then
   ChkCompQty = True
        
Else
   ChkCompQty = False
End If

If ChkCompQty = False Then
   InsSAP_BOM_FAIL WO, Woinfo.MBPN, "Comp Qty does not match: " & CompPN & " (SAP_BOM Qty:" & SAPBomQty & ")" & "(ME Bom Qty:" & MEBomQty & ")" & " CompLevel:" & CompLevel
    
End If


End Function

Public Sub InsSAP_BOM_FAIL(ByVal Work_Order As String, ByVal MBPN As String, ERR_DESC As String)
    Dim strSQL As String
    Dim Tran_Date As String, Tran_Time As String
    Tran_Date = Format(Now, "YYYYMMDD")
    Tran_Time = Format(Now, "HHNNSS")
    strSQL = "Insert SAP_BOM_FAIL(Work_Order,MBPN,ERR_DESC,Tran_Date,Tran_Time) values('" & Trim(Work_Order) & "','" & Trim(MBPN) & "','" & (ERR_DESC) & "'," & _
        " '" & Tran_Date & "','" & Tran_Time & "')"
    Conn.Execute strSQL
End Sub

Public Function ChkReplacePN(ByVal CompPN As String, ByVal WO As String, ByVal Ctype As String) As Boolean
Dim Str As String
Dim Rs As ADODB.Recordset
Dim ID As String
ID = ""
ChkReplacePN = True
'(1) Get Replace ID accroding to WO.Version & ComppN
Str = "select ID from QSMS_ReplacePN where version='" & Woinfo.Version & "' and CompPN='" & CompPN & "' and JObPN in (select JOBPN from QSMS_JOBBOm  where Work_Order='" & WO & "')"
Set Rs = Conn.Execute(Str)
If Rs.EOF Then
   ChkReplacePN = False
   Exit Function
Else
   ID = Trim(Rs!ID)
End If
Select Case UCase(Ctype)
       Case "SAP_BOM" ' check if lost in MeBom
             Str = "select a.CompPn from QSMS_MEBom a,QSMS_ReplacePN b where a.JObPN=b.JobPN and a.version=b.version and a.compPn=b.compPN and b.ID='" & ID & "'"
             Set Rs = Conn.Execute(Str)
             If Rs.EOF Then
                ChkReplacePN = False
                Exit Function
             End If
       Case "ME_BOM" ' check if lost in SAP BOM
             Str = "select ID from QSMS_ReplacePN where version='" & Woinfo.Version & "' and JObPN in (select JOBPN from QSMS_JOBBOm  where Work_Order='" & WO & "') and ID='" & ID & "'" & _
                   " and CompPN in (select CompPN from Sap_BOm where Work_Order='" & WO & "')"
             Set Rs = Conn.Execute(Str)
             If Rs.EOF Then
                ChkReplacePN = False
                Exit Function
             End If
End Select
End Function
Public Function InsertToQSMS_WO(ByVal WO As String)
Dim Str As String
Dim Rs As ADODB.Recordset
Dim RsBom As ADODB.Recordset
Dim Item As Integer
Item = 0
ReplaceItem = 0
Str = "select Machine,CompPN,Slot,Qty,JobPN from QSMS_MEBom where JobPN in (select JobPn from QSMS_JobBom where Work_Order='" & WO & "') and version='" & Woinfo.Version & "' "
Set RsBom = Conn.Execute(Str)
While Not RsBom.EOF
      
      Call InsertQSMSWOByComp(WO, Trim(RsBom!CompPN), Trim(RsBom!Machine), RsBom!Qty, Trim(RsBom!Slot), Trim(RsBom!JobPN))
      RsBom.MoveNext
Wend
End Function
Public Function InsertQSMSWOByComp(ByVal WO As String, ByVal CompPN As String, ByVal Machine As String, ByVal BaseQty As Integer, ByVal Slot As String, ByVal JobPN As String)
Dim Str As String
Dim Rs As ADODB.Recordset
Dim TempRs As ADODB.Recordset
Dim ID As String
Dim NeedQty As Long
'##################for JuJi system,one slot has two subslot(L,R),so maybe need insert the record by LR.----mark by leimo 20060516##################

NeedQty = 0
BaseQty = BaseQty / Woinfo.CombineQty
NeedQty = CDbl(BaseQty) * Woinfo.WOqty
Str = "select ID from QSMS_ReplacePN where version='" & Woinfo.Version & "' and CompPN='" & CompPN & "' and JObPN ='" & JobPN & "'" 'in (select JOBPN from QSMS_JOBBOm  where Work_Order='" & WO & "')"
Set Rs = Conn.Execute(Str)
If Rs.EOF Then
     
     Str = "select Work_Order from QSMS_WO where Work_Order='" & WO & "' and CompPN='" & CompPN & "' and Slot='" & Slot & "' and Machine='" & Machine & "'"
     Set TempRs = Conn.Execute(Str)
     If TempRs.EOF Then
               
        Str = "insert into QSMS_Wo(Work_Order,Line,WoQty,MBPN,Machine,CompPN,Slot,Item,BaseQty,NeedQty,DispatchQty,BalanceQty,MachineFinishedFlag,WoFinishedFlag) values" & _
             "('" & Woinfo.WO & "','" & Woinfo.Line & "'," & Woinfo.WOqty & ",'" & Woinfo.MBPN & "','" & Machine & "','" & CompPN & "','" & Slot & "','0', " & BaseQty & "," & NeedQty & ",0,-" & NeedQty & ",'N','N' )"
              
        Conn.Execute Str
    Else
        Str = "Update QSMS_Wo set BaseQty='" & BaseQty & "',NeedQty=" & NeedQty & ", Item='0',Balanceqty=dispatchQty-NeedQty where Work_order='" & WO & "' and CompPN='" & CompPN & "' and Slot='" & Slot & "' and Machine='" & Machine & "' "
        Conn.Execute Str
    End If
Else
   ReplaceItem = ReplaceItem + 1
   ID = Trim(Rs!ID)
   Str = "select CompPN from QSMS_ReplacePN where version='" & Woinfo.Version & "' and ID='" & ID & "' and JObPN='" & JobPN & "'" 'in (select JOBPN from QSMS_JOBBOm  where Work_Order='" & WO & "')"
   Set Rs = Conn.Execute(Str)
   While Not Rs.EOF
    
         Str = "select Work_Order from QSMS_WO where Work_Order='" & WO & "' and CompPN='" & Rs!CompPN & "' and Slot='" & Slot & "' and Machine='" & Machine & "'"
         Set TempRs = Conn.Execute(Str)
        
         If TempRs.EOF Then
        
             Str = "insert into QSMS_Wo(Work_Order,Line,WoQty,MBPN,Machine,CompPN,Slot,Item,BaseQty,NeedQty,DispatchQty,BalanceQty,MachineFinishedFlag,WoFinishedFlag) values" & _
             "('" & Woinfo.WO & "','" & Woinfo.Line & "'," & Woinfo.WOqty & ",'" & Woinfo.MBPN & "','" & Machine & "','" & Rs!CompPN & "','" & Slot & "','" & ReplaceItem & "', " & BaseQty & "," & NeedQty & ",0,-" & NeedQty & ",'N','N' )"
            Conn.Execute Str
         Else
            Str = "Update QSMS_Wo set BaseQty='" & BaseQty & "',NeedQty=" & NeedQty & " , Item='" & ReplaceItem & "',Balanceqty=DispatchQty-NeedQty where Work_order='" & WO & "' and CompPN='" & Trim(Rs!CompPN) & "' and Slot='" & Slot & "' and Machine='" & Machine & "'"
            Conn.Execute Str
         End If
         Rs.MoveNext
   Wend
End If
End Function

Public Function ClearBomFail(ByVal WO As String)

End Function

Public Function GetReplaceCompPN(ByVal Work_Order As String, ByVal CompPN As String, Woinfo As WOBasic) As ADODB.Recordset
Dim Str As String
Dim Rs As ADODB.Recordset
Dim TempID As String
Str = "select ID from QSMS_ReplacePN where version='" & Woinfo.Version & "' and JObPN in (select JObPN from QSMS_JobBom where Work_Order='" & WO & "') and CompPN='" & CompPN & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   TempID = Trim(Rs!ID)
End If

Str = "select CompPN from QSMS_ReplacePn where version='" & Woinfo.Version & "' and JobPN in (select JobPN from QSMS_JobBom where work_order='" & WO & "') and ID ='" & TempID & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
   Set GetReplaceCompPN = Rs
Else
   Str = "Select 'None' as CompPN"
   Set Rs = Conn.Execute(Str)
   Set GetReplaceCompPN = Rs
End If

End Function

'Public Function GetSmallBoardGroup(ByVal Work_Order As String) As String
'Dim Str As String
'Dim Rs As ADODB.Recordset
'Str = "select [Group] from Sap_Wo_List where WO='" & Work_Order & "'"
'Set Rs = Conn.Execute(Str)
'If Not Rs.EOF Then
'  GetSmallBoardGroup = Trim(Rs!Group)
'  Str = "Select Count(*) from Sap_Wo_List where [Group]='" & " Trim(Rs!Group) & " '"
'  Set Rs = Conn.Execute(Str)
'  Woinfo.GroupWoQty = Rs.Fields(0)
'Else
'   GetSmallBoardGroup = ""
'End If
'
'End Function

Public Function GetSmallBoardSapQty(ByVal CompPN As String, ByVal CompLevel As String) As Integer
Dim Str As String
Dim Rs As ADODB.Recordset
Dim TempRs As ADODB.Recordset
Dim JobPN As String

GetSmallBoardSapQty = 0
If Woinfo.GroupWoQty < 2 Then
   GetSmallBoardSapQty = 0
   Exit Function
End If
Select Case CompLevel
       Case "01"
            JobPN = "41"
       Case "02"
            JobPN = "51"
       Case "00"
            Str = "select JobPN from QSMS_ReplacePN where version='" & Woinfo.Version & "' and JObPN in (select JObPN from QSMS_JobBom where Work_Order='" & WO & "' and jobpn like '" & JobPN & "%') and CompPN='" & CompPN & "' "
            Set Rs = Conn.Execute(Str)
            If Not Rs.EOF Then
                JobPN = Mid(Trim(Rs!JobPN), 1, 2)
            End If
End Select

Str = "Select Wo ,PN from Sap_Wo_List where [Group]='" & Woinfo.Group & "'"
Set Rs = Conn.Execute(Str)
While Not Rs.EOF
      If Woinfo.WO = Rs!WO Then
      Else
          Str = "Select Qty from QSMS_SBBOm where MBPN='" & Woinfo.MBPN & "' and SbPN='" & Trim(Rs!PN) & "' and JobPN like '" & JobPN & "%'  and CompPN='" & CompPN & "'"
          Set TempRs = Conn.Execute(Str)
          If TempRs.EOF Then
             GetSmallBoardSapQty = 0
          Else
             Str = "select Item,CompLevel from Sap_Bom where Work_Order='" & Trim(Rs!WO) & "' and CompPN='" & Trim(CompPN) & "'"
             Set TempRs = Conn.Execute(Str)
             If Not TempRs.EOF Then
                    Str = "select sum(qty) from sap_bom where work_order='" & Rs!WO & "' and item='" & TempRs!Item & "' and CompLevel='" & TempRs!CompLevel & "'"
                    Set Rs = Conn.Execute(Str)
                    If Not Rs.EOF Then
                       GetSmallBoardSapQty = Rs.Fields(0)
                    End If
             End If
          End If
      End If
     Rs.MoveNext
Wend

End Function
Public Function ChkCompBelongMBBom(ByVal CompPN As String, ByVal SBWo As String, ByVal SBPN) As Boolean
Dim Str As String
Dim Rs As ADODB.Recordset
Dim TempRs As ADODB.Recordset
Dim CompLevel As String
ChkCompBelongMBBom = True
Str = "Select JobPN from QSMS_SBBOm where MBPN='" & Woinfo.MBPN & "' and SbPN='" & Trim(SBPN) & "'  and CompPN='" & CompPN & "'"
Set TempRs = Conn.Execute(Str)
While Not TempRs.EOF
       If Mid(Trim(TempRs!JobPN), 1, 2) = "41" Then
          CompLevel = "01"
       Else
          CompLevel = "02"
       End If
      Str = "Select CompPN from SAP_BOM where Work_Order='" & Trim(Woinfo.WO) & "' and CompPN='" & CompPN & "' and CompLevel='" & CompLevel & "'"
      Set Rs = Conn.Execute(Str)
      If Rs.EOF Then
         ChkCompBelongMBBom = False
      Else
         ChkCompBelongMBBom = True
      End If
      TempRs.MoveNext
Wend


End Function

Public Function GetSBCompLevel(ByVal CompPN As String, ByVal SBPN As String, ByVal MBPN As String)
Dim Str As String
Dim Rs As ADODB.Recordset
Str = "Select JobPN from QSMS_SBBOm where MBPN='" & Woinfo.MBPN & "' and SbPN='" & Trim(SBPN) & "'  and CompPN='" & CompPN & "'"
Set Rs = Conn.Execute(Str)
If Not Rs.EOF Then
       If Mid(Trim(Rs!JobPN), 1, 2) = "41" Then
          GetSBCompLevel = "01"
       Else
          GetSBCompLevel = "02"
       End If
End If
End Function
