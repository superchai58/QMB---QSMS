Attribute VB_Name = "mdlProgramVersion"
'/**********************************************************************************
'**�� �� ��: SMT_QSMS.frm
'**Copyright (C) 2007-2010 QMS
'**�ļ����:
'**�� �� ��: Giant
'**��    ��: 2010.08.12
'**��    ��: ---to be added
'
'EQMS_ID        '**�� �� ��     �޸�����        ��    ��
                '-----------------------------------------------------------------------------
'QMS             Jeanson        2010/08/12     save the modification log into one unified file '(1000)
'                kaitlyn        2010/08/17     add a template of upload_traycompPN            '(1001)
'RQ10061811      Lynn           2010/08/24     Check Forbidden PN when maintain feeder            '(1002)
'QMS             Giant          2010/08/25     Print Bug           '(1003)
'QMS             Jocelyn        2010/08/26     Add appname limit when user login '(1004)
'QMS             Denver         2010/08/27     can not Print DID Label(Return DID/ReturnComp/DIDCallBack)  '(1005)
'QMS             Udall          2010/09/06     ���ӳ���,���ɢ�Ϸ��� '(1006)
'QMS             Jocelyn        2010/09/08     Add NB7 in '(1007)
'QMS             Bay            2010/09/22     Add display RealQty (1008)
'QMS             Jocelyn        2010/09/23     Add XL_MaxDIDMaintainQty load��Execl���ܣ�1009��
'QMS             Walton         2010/09/27     Modify QSMS Trace report file format from Excel to CSV file '(1010)
'AS              Udall          2010/09/27     ���AP�ڷ�"����"����Ҳ��Ҫ�˶�MSD    (1011)
'QMS             Jocelyn        2010/10/12     �ж�upLoad_MachineTypeʱ��execl�ļ��е�SeqIDByLine�Ƿ�Ϊ���� '(1012)
'                Maggie         2010/10/14     �����¾��Ϻ�ת��ʱ��Label��ӡ  '(1013)
'QMS             Udall          2010/10/18     Add form function frmDispatchDataToQWMS      (1014)
'QMS             Giant          2010/10/25     Add new column BufferQty for upload XL_WOPlanSeq      (1015)
'RQ10110212      Giant          2010/11/02     Make up some bug for re-print DID      (1016)
'QMS             Jocelyn        2010/11/03     Print buffer not enough,can't print    (1017)
'MFG             Kaitlyn        2010/11/12     Get PN/VendorCode/DateCode/LotCode/Qty from 2D barcode  (1018)
'RQ10110907      Maggie         2010/11/15     Save Printer setting in local Registry (1019)
'QMS             kaitlyn        2010/11/22     get returnDID 's infomation from QSMS_DID_TOWH  (1020)
'QMS             kailtyn        2010/11/24     upLoad machineType:the first letter can't be accepted [0-9]  '(1021)
'QMs             Austin         2010/11/24     ����MSD���������txtInSpection��λ.  (1022)
'QMS             Austin         2010/11/30     �޸�ǿ�����ϴ�ӡ���ַ�50�Ĵ�         (1023)
'QMS             kaitlyn        2010/12/08     UploadData Add item:Component_Data  (1024)
'QMS             Walton         2010/12/09     UploadData/TransferPanaMSF: the table MEbom add column :line(1025)
'QMS             Jocelyn        2010/12/09     FrmTransferFujiXML/frmTransferPanaAMI: the table MEbom add column :line(1026)
'QMS             Jocelyn        2010/12/14     Print buffer not enough,can't print    (1027)
'QMS             kaitlyn        2010/12/16     uploadData:1.��functype=sheetname(ȡ��xls�����ƣ���2.component_data��itemΪ���� ��1028��
'QMS             Udall          2010/12/17     Add Side for Machine        (1029)
'QMS             Maggie         2010/12/21     �޸�Bug              (1030)
'QMS             Giant          2010/12/22     ����MEBOM for 3���߱�              (1031)
'QMS             Jocelyn        2010/12/27     check machine whether define or not while uploading machine bom  (1032)
'QMS             Walton         2010/12/27     Modify Report: query Machine Bom by WO            (1033)
'QMS             Bay            2010/12/29     Modify MainTainFeeder :  Add   @machine  (1034)
'QMS             kaitlyn        2010/12/29     update delete me_bom frm   (1035)
'QMS             Jocelyn        2010/12/30     ȡ���߱��machine����ȡmachine   (1036)
'Qms             Walton         2010/12/30     add line and side           (1037)
'                Maggie         2010/12/31     FrmReturnComp:CCBUֻ�����txtCompPN,txtQty   (1038)
'QMS             Udall          2011/01/01     ��Machine����ΪPanaʱ�˶�Machine��MappingID���ȱ���Ϊ7��         (1039)
'QMS             Austin         2011/01/03     ��ʱȡ��MappingID�����飬��θĴ���                           (1040)
'QMS             Austin         2011/01/05     �޸�ǰ��ȡgroupID(2,8),�޸ĺ��ȡGroupID(3,8)                   (1041)
'QMS             Denver         2011/01/04     can not Upload Fuji XML data                                   (1042)
'QMS             Jocelyn        2011/01/05     update method to upload fuji XML data                          (1043)
'QMS             Jocelyn        2011/01/07     Get ComPPN from 2D barcode                                    '(1045)
'QMS             kaitlyn        2011/01/10     upload machinetype:others��̨����mappingID   (1046)
'QMS             Giant          2011/01/10     NB5 upload MEBOM:*others��̨If defined in machine   (1047)
'QMS             Denver         2011/01/12     Unify SP:XL_ChkAnotherBUDID to Add Para:Factory in CallBack/ReturnDID   (1048)
'QMS             Alawn          2011/01/14     Add add a column line for XL_DoubleTable      (1049)
'QMS             Jocelyn        2011/01/28     modify limit qty from 20000 to 60000 '(1050)
'QMS             Bay            2011/01/28     Modify   Print Bug (1051)
'QMS             Giant          2011/02/10     Modify Integer to String,user input >50000 (1052)
'QMS             Walton         2011/02/21     check XL is running when upload XL plan (1053)
'QMS             Alawn          2011/03/01     Modify Bug for get machine         (1054)
'QMS             Giant          2011/03/14     add column "FuncType" for table  QSMS_NoCheckReplacePNSplicing   (1055)
'QMS             Giant          2011/03/24     modify get date   (1056)
'RQ11041312      Alawn          2011/04/19     Add DelDateTime for ForbiddenPN Report                 (1056)
'QMS             Austin         2011/05/11     ��ֹ���ϵ�ʱ��,�������OK��ť,���²������DID,�ڴ�����֮ǰ��OK��ťDisable   1058
'QMS             Jocelyn        2011/05/11     ��ʲ�ѯ '1059
'QMS             Jocelyn        2011/05/17     ȥ��byʱ������� '1060
'QMS             Jeanson        2011/05/27     query unlink feeder and DID by WO'1061
'CC              Denver         2011/06/21     Add Dummy ECN function
'PO-RQ11062703   Felix          2011/07/04     Adjust DID dspatch label format (1063)
'RQ11060735      Bay            2011/07/14     Add DID and CompPN check Function    (1064)
'RQ11060107      kaitlyn        2011/07/20     upload machinetype:�޸ļ����� (1065)
'RQ11071314      Alawn          2011/07/21     MATERIALTOWHID�ܹ��ϴ�%���Ϻ�     ��1066��
'RQ11071314      Alawn          2011/07/21     Add PrintAutoDispatchLabelNetWorkPort     ��1067��
'QMS             Kaitlyn        2011/08/10     upload qsms_mebom:�����machinename����(1068)
'QMS             Giant          2011/08/16     get reelwidth from [stockdata] for NPM machine(1069)
'QMS             Austin         2011/08/24     Modify upload AMI�ļ� for NPM DualLane=MIX
'QMS             Bay            2011/08/24     SET.INI add FLAG: DIDScan , ���DIDScan="Y" QSMS Maintain DID auto Dispatch��Return DID �е�CompPN��DID����ͨ�� scaner ���� (1071)
'QMS             Maggie         2011/08/24     �޸�upload XL_WOPlanSeq����¼DualLane=MIX (1072)
'QMS             Alawn          2011/09/14     ���Buildtype=4��wo����ɾ��ME_Bom��ʱ��ɾ�������ߵ�����,����ME_BOM�����ߵ�����   ��1073��
'RQ11090907      Bay            2011/09/25     ���Cϵ��ɢ�ϣ����Ͻ��ѷ��ϵ�ɢ�ϻ�δ�������ߵĲ������Ϻ�����ע���µ�DID���������ߣ����߽�ɢ�ϵĶ���ת�Ƶ�������ɣ�
                                               'Ŀ�ģ�1���ٲ��߽��ϵĶ�����2���Ͳ��ߴ��ϵĻ���(1074)
'QMS             Maggie         2011/10/17     �ϴ�AMI�ļ�ʱ����ļ�����,�����Ƿ�ΪNPM DualLaneMode(1075)
'RQ11101304      Alawn          2011/10/25     QSMS Maintain Feeder �е�DID����ͨ�� scaner ���� ��1076��
'QMS             Jerry          2011/10/25     ��Report��ӿ��Բ�ѯ�������Ϻ�����ϵ������Ϣ          (1077)
'QMS             Bay            2011/11/20     Modify Bug :  SQL����﷨����(1078)
'QMS             Giant          2011/11/22     Read MCType from QSMS_ProConfig(1079)
'RQ11120813      kaitlyn        2011/12/19     ����200���ӡ�����޸�DIDģ��ѡ�񲿷�(1080)
'qms             Jerry          2011/12/22     MaintainFeederʱ��QSMS_LOG���¼(1081)
'qms             Jerry          2011/12/22     ��Report��ӿ��Բ�ѯMaintainFeeder�������Ϣ(1082)
'QMS             Walton         2011/12/28     1>�����¼����ĳ�ͨ��MainMenu��¼��
                                        '''    2>��Factory����ͨ��MainMenu������û�д������վ��еķ�ʽSetting.ini����     ''''(1084)
'QMS             Kaitlyn        2011/12/29     �ϴ�PANA me_bomʱ���޸�machinename�ļ�����1083��
'QMS             Jerry          2012/01/12     ����NB4 ��zebra_300_old.txt�����ӱ���<MACHINE1>��1084��
'QMS             Jerry          2012/02/16     ��FrmInSpection�������Ӹ���COMPPN��ѯ��Ӧ���ݵĹ��ܣ�1085��
'QMS             Jerry          2012/03/16     ����NB5 ��zebra_300_new.txt�����ӱ���<SLOT1>��1086��
'QMS             Walton         2012/03/20     ��ˢ���DIDΪ�н���ʱ��Ҫ���֮ǰ���нӵ��ϵ�RealQty�Ƿ����0.������0,�򽫴���֮ǰ��DID .  (1087)
'QMS             Jerry          2012/03/27     1.frmMaintainDIDAutoDispatch��������������ʾ.  (1088)
'                                              2.frmMaintainDIDAutoDispatch���ִ�����ʾʱ��������Ϣ
'                                              3.frmMaintainDIDAutoDispatch������ˢ��COMPNʱ��ֱ������TxtGroupQty
'QMS             Jerry          2012/04/17     ��Report��ӿ��Բ�ѯPDA_DistributeDIDLog�������Ϣ.  (1089)
'QMS             kaitlyn        2012/04/17     fixed bug:�ֶ��н�ʱby side�н��޸�line+side��ȡֵ��������machineǰ�����ʾ(1090)
'QMS             Jerry          2012/04/25     1.����FrmDIDInteGration�����һ��BUG(1091)
'                                              2.����FrmDIDIGration�����strDay������ȡֵ
'                                              3.FrmDIDIGration�������ӱ���<MACHINE1>
'QMS             Bay            2012/04/26     ʹ��API������Ƶ�ļ� (1092)
'QMS             Walton         2012/05/14     ��ӡDIDʱ����DIDΪָ�����������ӹ�������,��Ҫ�޸�SP��[QSMS_GetDIDPrintInfo]��XL_GetDidPrintInfo_Return ���Բ���NB7  ��1093��
'QMS             kaitlyn        2012/05/16     ��13���Ͼ�����Ϊ50K����Ҫ�޸ı�������Ϊlong (1094)
'QMS             Alawn          2012/05/16     �޸���ӡDID��¼��������                      ��1095��
'QMS             Bay            2012/05/31     ���FrmDIDDistribution���壬�������� �����ߣ��棬��̨���з����Ӧ��DID���з��� ��1096��
'RQ12052311      Dreamu         2012/06/07     Get PN/VendorCode/DateCode/LotCode from 2D barcode,��ӡ����DateCode��LotCode�ֿ����� (1097)
'RQ12052204      Dreamu         2012/06/08     �޸�QSMS\PD\UpdateDIDRealQTY��ӡ�Reason����λ   (1098)
'QMS             scofield       2012/06/21     modify limit qty from 60000 to 120000 '(1099)
'QMS             Allsa          2012/06/27     ���ϴ�����CompPN��VendorCode��DateCode��LotCode��λ��������Key���Ҳ����ø��ƺ���� '(1100)
                                               '�������޸�����RQ12052311�н�CompPrint�����ԭ��������KeyVendorCode��DateCode��LotCode��λ��ȡ�����ˣ��������϶���������滹����Ҫ��Key�ģ���Ϊ�������ֻ��ָ��Ա�ſ��Ե�¼����
'RQ12062909      Dreamu         2012/07/10     DID Label�ϵ�Dataץȡ��ʱ�侫ȷ���롣Format=120629080000      (1101)
'RQ12051615      Alawn          2012/07/12     ��ֹPD\CompPNCompare�����ܹ����ƺ�ճ�� ��1102��
'QMS             Lynn           2012/07/25     ����flag��MaintainFeederDID�����������Ϊfuji��ʱ��ֻ��Ҫˢ��DID,CompPN,Feeder���� ��1103��
'QMS             Walton         2012/07/01     �޸�IC_CompPN��ѯͨ��PN���������IC_Burn�ĵ���ģ��    (1104)
'QMS             Walton         2012/08/13     ��ӡDIDʱ����DIDΪָ������������DIDType����,��Ҫ�޸�SP��[QSMS_GetDIDPrintInfo]��XL_GetDidPrintInfo_Return ���Բ���NB7  ��1105��
'QMS             Walton         2012/08/17     ����ָ���������ĵ�        ��1106��
'QMS             Lynn           2012/08/29     Return DID��ʱ�������һ��DID������������0������ʾ        ��1107��
'QMS             Austin         2012/09/24     if MCType = NPMMachineType,ReelWidthȡNPMReelWidth        ��1108��
'EQMS           kaitlyn         2012/10/16     ��ӡDIDʱ�ǲ�������DID����ʾ����(CYL:��ͬgroup�����������ų̣��ȶԵ�ǰ�����������һ�����������ϵĲ��죩(1109)
'QMS            kaitlyn         2012/10/30     �ϴ��ĵط���table QSMS_CompPNcheck ��ΪQSMS_CompPNcheck_temp(��Ϊ��table����PN_Receiving���õ�����������1110��
'QMS            kaitlyn         2012/11/13     fixed bug:ϵͳ�жϲ��ǵ����DID��returnʱ��ֱ��ɾ��DID������ǵ���Ĳ�ɾ����Ŀ����Ϊ�˷�ֹDID�ظ�.�ȶԵ�ʱ��ȵ���YYYYMMDD,����returnʱ���¼��ʱ����YYMMDDHHNNSS(1111)
'QMS            Dreamu          2012/11/14     �޸ĳ�ʽ��ʹ��ӡLabel��ʱ��ӷ�������ץȡ   (1112)
'QMS            kaitlyn         2012/11/20     update sql(ͬһ��DIDֻ�ܽ��ϵ�һ��predid�ϣ�DIDSLOTLINK�п�)�����Ƕ���fujitraxû�п��������޸�sql��ץȡ���һ�����ϵ���Ϣ)��1113��
'QMS            Link            2012/12/03     QSMS\MCC\CompPint\���Զ���2Dbarcode�л�����ݣ����QTY��1114��
'RQ12111406     Cynthia         2012/12/05     ��ӡLabel����Ӳ��������ں�ʱ�䣬����Ϊ��<DateTime>��(1115)
'RQ12111408     kaitlyn         2012/12/06     returnDIDʱ�Զ�����realqty��ѡ�� ��1116��
'RQ12120309     Dreamu          2012/12/12     NB6��frmCompPrint������Ӵ�ӡ����Mark   (1117)
'RQ12120301     Alawn           2012/12/27     ���SAP2��ѯ��GetGroupIDDataByCompPN ��1118��
'QMS            Ava             2012/12/27     �޸�frmCompPrint��ȡ�ļ��ķ�ʽ  (1119)
'RQ12122103     Allsa           2013/01/15     NB1��һ��������ˢ�����EMMC������Ҫ�ȶ�ImageVersion��MBPN�Ķ�Ӧ��ϵ  (1120)
'QMS            Newton          2013/01/24     �����ϴ��ų�������һ��ѡ��XL_WOPLANSEQSHIFTID��������һ����λ��ShiftID(1121)
'QMS            Walton          2013/01/29     �ϴ�IC_CompPN ʱ��Ҫ��������Ƿ����   (1122)
'QMS            Jerry           2013/01/29     �ڵ�����Excel������һ��ѡ��XL_WOPLANSEQSHIFTID(1123)
'QMS            Ava             2013/01/31     NB5ֻ��ȡPP10����    (1124)
'QMS            Walton          2013/01/31     �����ϴ�IC_ShearPin��ѡ� (1125)
'RQ13031414     Link            2013/03/20     mnuMaintainDIDȨ��ֻ����frmMaintainDID���壬����mnuMaintainDIDAutoDispatchȨ�޸���FrmMaintainDIDAutoDispatch����  ����1126��
'QMS            Jerry           2013/03/29     �ڵ�����Excel������һ��ѡ��GetMEBom_ByGroupID(1127)
'QMS            Ava             2013/04/02     ��������ӵ�GroupID֮ǰ������Check������ ��1128��
'QMS            Jerry           2013/04/11     UploadData Add item:Machine_Data�� ��1129��
'QMS            Jerry           2013/04/19     NB4���Ӹ���Typeѡ���ֶ��㼸СʱXL�� ��1130��
'QMS            Ava             2013/05/13     ��Ӹ����߱�ɾ��MEBOM��userRight��DeleteMeBomByLine����1131��
'QMS            Scofield        2013/05/13     ���ڣ�1085�����MBU IPQCɢ�ϣ��Ϻų��ȶ���11�룩�޷���ȡ��ȷ��CHENUMֵ��MARK��Len(lblcomppn) > 12����ж���������1132��
'QMS            Newton          2013/05/23     �ϴ�XL_WOPlanSeqShiftIDʱ����Flag��ʶ��Щ����ΪCTO���� (1133)
'QMS            Walton          2013/05/27     UploadData �������е�UploadIC_ShearPin����CompPN���� (1134)
'QMS            Walfan          2013/04/30     �ϴ�XL_WOPanSeqʱ����Flag��ʶ��Щ����ΪCTO���� (1135)
'QMS            Jerry           2013/06/17     Delete MEBOM����ɾ��ME_BOMʱ���Ը���JjobPN (1136)
'QMS            Alawn           2013/07/08     �޸�CompPNCompare������line��ƥ���bug ��1137��
'QMS            Walfan          2013/07/16     Report����DispatchDID ��ѯ����CompPN���� (1138)
'QMS            Dreamu          2013/08/26     FrmMainWoSeq�����޶�BUG,��Ӳ�ѯ��ʷDB   (1139)
'QMS            Jerry           2013/09/02     ��FrmUrgentInsertWO��Query��ť�µ��õ�SP��ΪXL_SpecialCaseByWO_New  (1140)
'QMS            Walton          2013/09/03     ���Machine_Data��ģ��        (1141��
'QMS            Walton          2013/10/17     �ڴ�ӡDID������DIDType ����   (1142)
'RQ13102407     Allsa           2013/11/01     ��ǰ��Fuji�豸û�е���FuJiTraxʱ���������ϴ�����100��slot��Ϊ��ʹDIDSlotLink��ʽ����ȥ��������Ҫ����100.
                                               '��ǰFuji��NXT�豸������FujiTrax����Tray�̵�Slot��900���ϵģ����Ҵ�Slot��Ҫ��FujiTraxϵͳ��ƥ�䡣�豸��ϣ�����ϴ�ʱ����Ҫ��һ����100�Ķ�����������NXT�豸ȡ�����ǳ�ʽȡ����ȥ100�Ķ�����(1143)
'QMS            Jerry           2013/11/25     �޸�FrmGenXLMD�����cboType���ɴ�Tble:XL_TypeDateTime�ж�ȡ    (1144)
'QMS            kaitlyn         2013/12/06     ��ȡDID�Ĺ����и��ģ����ϵ�DID���ϺŶ�����-A����Ҫ���ֿ������������DID���ظ�(1145)
'QMS            kaitlyn         2013/12/06     fixed bug:maintain feederʱ����Ҫ��Ҫ���LR�ĺϷ��ԣ�ֻĿǰֻ��0��1��2���� (1046)
'QMS            Jerry           2013/12/10     ����Flag��PrintedSeqID,�ܿ��Ƿ��ӡDIDʱ���ӱ���COUNT������Mark��DID�����GoupID�����Slot��LR�ĵڼ����ϣ������������(1147)
'QMS            Jerry           2013/12/10     �������棨1101���޸ģ�NB4���滻�˱���Data�󣬻���ס��������ݣ�����NB4�Ļ�YYYYMMDD��ʽ(1148)
'QMS            Walfan          2014/01/13     ����Bug:�ϴ�machine_data����ʱ��by lineɾ�����ݣ������������ɾ (1149)
'QMS            Scofield        2014/01/20     �����ϴ�MB_REV=�յ�MEBOM   (1150)
'QMS            Walton          2014/02/10     MaintainFeeder ������������Check ����WO�İ�Feeder��������Լ�������������ʾ����    (1151)
'QMS            Van             2014/02/11     ��ԣ�1150���޸ģ�ȡ��tempversion=""������   (1152)
'QMS            Ava             2014/02/17     NB5 closeWOʱ��ѡ��ɾ��DID    (1153)
'QMS            Jerry           2014/02/21     UploadData Add item:CompPN_Spacer,�ϴ���ӦCompPN��Ӧ�ĵ�Ƭ��Ϣ�� ��1154��
'QMS            Jerry           2014/03/11     ���1149���޸�By lineɾ������ʱ�ټ���FuncType='DualLaneMode' and Item='Independent'���������� ��1155��
'QMS            Walton          2014/03/13     ȡ��IC_CompPN�ϴ��ĳ�ͨ����ʽUploadBasicData�ϴ� ��
'QMS            Walton          2014/03/24     ȡ��IC_CompPN�Ĳ�ѯ                           (1157)
'QMS            Newton          2014/03/28     NB3ֻ��ȡPP10����    (1158)
'RQ14040405     Anker           2014/04/14     MaintainFeeder֮����Check����Ѵ˹�������δ��MaintainFeeder�Ļ�̨Slot ��L/R�Լ������Ϻ���ʾ��excel��  (1159)
'RQ14031414     kaitlyn         2014/04/16     Ϊ�˼������ϵĴ�����ʵ���Ų��,return����ʱ��������ԭ�߱����������K17�����ϣ�Return ʱ��K17 ������������K17�������ٿ������߱���û������û�����˲֡�������ԭ�߱�����ʱ�������趨����SLOT��֮ǰ���ĸ����ĸ�SLOT �������������������SLOT ����н�ʱ����һ��  (1160)
'RQ14042101     Walton          2014/04/21     ����Fuji �µĻ���AIMEX�Ĵ���               (1161)
'RQ14042802     Dreamu          2014/04/29     �޸�ReturnDID���棬��Return DIDʱ����Ӷ�������ʾ��Ϣ����  (1162)
'QMS            Newton          2014/05/16     �޸�CompPrint��ReturnDID���ܣ���ʽֱ�Ӵӵ��ϻ���Com���ж�ȡQty  (1163)
'RQ14052112     Cynthia         2014/06/08     �޸�ͨ��BuildType=4(����������)��WO��ȡLine�ķ�ʽ��ʹ�ÿ���ֱ�Ӱڵ�WO��ѯ��ɾ��2���ߵ�ME_Bom���� ��1164��
'QMS            Cynthia         2014/06/10     ���Ӵ���chkDomain��������Quantan��ĵ�����ͳһexportΪhtml��ʽ ��1165��
'QMS            Newton          2014/06/11     ���߿��ǵ���ҪΪ���ϻ�����Shopfloor���ԣ�ȡ��(1163) ��1166��
'QMS            Cynthia         2014/06/12     ֧�ֶ�����������ΪHtml��ʽ�ļ� ��1167��
'QMS            Cynthia         2014/06/13     �޸�Report�����CboReportType��Program_DefineItem�л�ȡ ��1168��
'QMS            Cynthia         2014/06/14     ʹ��PrepareMaterialByWONew����PrepareMaterialByWO,ȡ��ģ�� (1169)
'QMS            Cynthia         2014/06/16     ת��Htmlʱ������Nullֵͳһ����Ϊ"",�����һλΪ"=",����ͳһ����[].��1170��
'QMS            Cynthia         2014/06/18     ͳһʹ��utf-8�ַ���exprot to html ��1171��
'QMS            Ava             2014/07/11     NPM�����ϴ�MEBOM������Skip���顢վλ��Ԫ��  ��1172��
'QMS            Lynn            2014/07/15     �ϴ������ų̵ĵط����Ӽ�飬��������ִ������JOB��ǰ��15�����ϴ��ų�  ��1173��
'QMS            Newton          2014/08/05     PT200�豸���������PositionData��������BG��PROGDIST�� ��1174��
'QMS            Giant           2014/08/15     NB5��һ������������û�ϣ�����С�����С����Բ�������������Ϊ׼��ȡ������InitAOIFlag='Y'�� ��1175��
'RQ14082505     Alawn           2014/09/03     ������xl����ʹ��MCCPreMaterial����  1176
'RQ14091504     Alawn           2014/09/22     ��Feeder��DIDʱ�����밴DID sequence�����������ϵ����϶�ʵ��һ������DID sequence��������OP�������򵥻������õر�����Ϸ��� 1177
'RQ14090106     Anker           2014/10/09     �ϴ�PanaAMI���ĵ�ʱ���������и����ֶ����ڵ��±���ץȡֵ  ��1178��
'QMS            Jerry           2014/11/05     ���NB4��Report��������checkBox:Dual������CheckBomʱ�ǰ��ն������ǹ���ģʽCheck  ��1179��
'RQ14110502     Cynthia         2014/11/19     �޸�CompPrint�����ӡ�Ϻ�Label��QTY����ԭ��Integer��ΪLong��������ߴ�ӡ������Ϻţ�1180��
'QMS            Sarah           2014/11/19     NB6�����ֲ��������� FrmStartSplitLineMC ��1181��
'QMS            Ava             2014/11/21     �ϴ�MEBOM�޸ģ�[Machines]��MCNAME�У�����NPM����MCNAMEֻ��ȡNPM��֮ǰ�Ĳ���   (1182)
'QMS            Ava             2014/11/21     �ϴ�MEBOM�޸ģ�Tray�̲��ϵ��Ϻź����"-T"��"-N"��ץȡ�Ϻ�ʱֻץȡ11λ���Ϻ� (1183)
'QMS            Ava             2014/11/27     1178�� SD(idxSD).TA = Trim(GetPosition(SD_Header, "TA"))δ��Arry   ��1184��
'QMS            Alawn           2014/12/09     RQ14111916  ���幤��BuildType=4 ʱ������Station����  1185
'QMS            Jerry           2014/12/19     ��MaintainFeeder�����TxtJobGroup�Ļ�ԭ����cboJobGroup  1186
'QMS            Ava             2014/12/23     MaintainFeeder����TxtCompPN��ˢ���DID�л�ȡ������flag:UnChkCompPN    1187
'QMS            Ava             2015/01/06     ����flag��checkNeedMSD�����ΪY������ʾ�ò�����濾 1188
'QMS            Walton          2015/01/22     ����ѯJOBPN�ĳ�SP��ʽ    1189
'QMS            Scofield        2015/02/03     ���Xboard������CLOSEWOʱ����Xboard���������ն�Ӧ��C���S���������    1190
'RQ14122224     Cynthia         2015/03/04     ����MSD���ϣ�����Flag=CheckMSDCallBack������flag��CallBack��NewDID����ʱ��н�OLD DID�Ŀ���ʱ��  1191
'QMS            Walton          2015/03/09     �����ڴ�ӡDID ʱ���DIDΪ��¼���ϣ���DID Text ������BurnRev ��Ϣ      1192
'QMS            Udall           2015/03/13     ��������Tray���ϴ�-T or -N����,���ȡ������������,ֻ���ڳ���<>11 and <>14ʱ�Ž���      1193
'QMS            Alawn           2015/03/30     �ϴ�MEBOM�����ӵȴ�1s 1194
'QMS            Jerry           2015/03/31     UpLoadData�����ϴ�AVLC�Ĺ��� 1195
'QMS            Walton          2015/04/01     ����DID_2D��ӡ�ı���    1196
'RQ15032012     Walton          2015/04/03     ReturnDID ʱ��DIDδ��ʹ�ã�����ʾ����  1197
'QMS            Jerry           2015/04/03     NB5���Ӱ�JOBGroupɾ��MEBOM�Ĺ���  1198
'QMS            Jerry           2015/04/03     ����Flag:NoKeepPWD��Check�Ƿ���Ҫ��¼����  1199
'RQ15041004     Cynthia         2015/05/06     MSD����Return/CallBack���ٷ��ϵģ�����Ӧ��DID��״̬�����������ܷ��� 1200
'QMS            Frame           2015/06/02     ����3523�ͺ�LCR������  1201
'QMS            Frame           2015/06/03      ����DID֮���ȥ����callPicture���������û��ͼƬ���˳�Function��������ִ��  1202
'QMS            Jerry           2015/06/16      �ϴ�ME_BOMʱץȡ[BoardData]�е�PCBSize����X="###",����¼��ϵͳ 1203
'QMS            Newton          2015/07/29      ����8110G�ͺ�LCR������ͳ 1204
'QMS            Newton          2015/08/05      ����NB3����һ�����ϲ֣�PE��������ѡ�����ϵ��ĸ��ֱ� 1205
'QMS            Frame           2015/08/24      ����BY WO��CompPN ������ѯ�Ĺ��� 1206
'QMS            Eason           2015/09/09      ����NB5���ӽ�������ְԱ���󶨵�DID��UID����Ϊ��ְԱ����UID  1207
'QMS            Newton          2015/10/13      ����NB3 B/S����Barcode�ϵ�Datecode�ǹ��Datecode���ڷ���ʱҪˢ�������ϵ�����Datecode����¼  1208
'RQ15101219     Yan             20151023        User��ɾ�����������Detele��ʱ��ʽȥ��鵱ǰ�����Ƿ��ϴ���XL,����е�ǰ�����ǲ�����ɾ���� 1209
'RQ15101203     Frame           2015/10/27      �����ϱ���ʱ��ǰ���Feeder��״̬������Ҫ������δ�ڵ���FujiTraxServerע���Feeder������Ӧ����ʾ(ֻ���Fuji��Feeder) 1210
'QMS            Eason           2015/10/30      ����Flag ChkPNCQ �ж������������ϺŲ������� 1211
'QMS            Eason           2015/11/12      PU5��report����CheckWO_WastagePNѡ�����ѯ�����Ϻ͸�Ƶ�ϡ� 1212
'QMS            Newton          2015/11/15      ��PE�ϴ�B/S �����Ϻţ�ϵͳ�Զ��ж�  1213
'QMS            Udall           2015/11/30      ��table:NozzleLocation������Line,Side,BuildType������   1214
'QMS            Frame           2015/12/04      �����ж�,<fsSetPos>��</fsSetPos>����Ϊ ���� - ���� ģʽ 1215
'QMS            Jerry           2015/12/17      ����(1085),����ˢ���DID�Զ���ȡǰ11�� 1216
'QMS            Giant           2015/12/18      [PartsData]���ִ��ڲ�������IDNO����Ҫ�ų�Ϊ�յĲ��� 1217
'QMS            Season          2015/11/12      �н��ϵ�ʱ�����DID֮ǰû�з������λ�ã������DID���ܳнӵ����λ�� ----(1218)
'QMS            Giant           2015/12/29      PU5��Ҫ�޸��ϴ�MEBOM���ļ�����ʽ�������Դ�Report�е�����������ݣ����FLAG:ChkEQProgram ----(1219)
'RQ16010604     Kver            2016/01/06      �޸Ľ���FrmInheritDIDByWO��ֻ��ʾδ�ؽ�Ĺ���Group  ----��1220��
'QMS            Austin          2016/01/10      ���ϵ�ʱ�򣬸��ݶ���Ĳֱ��ж��Ƿ�����Ϻ���ͬ������D\C��ͬ�����ݣ�����У�����ʾ (1221)
'RQ16011311     Yan             2016/01/19      ϵͳ����DA0��ʹ��ʱ�俨�أ����������DA0���޷�ʹ��(1222)
'QMS            Jerry           2016/01/21      DIDLabel������ӡ����<VenderCode1>��<LR1>,��Flag:PrintedVenderCode�ܿ�(1223)
'RQ16012215     Kver            2016/01/25      ����Maintain DID�д�ӡPU5ʱ���ʽΪ"YYYYMMDDHHNNSS" (1224)
'QMS            Jerry           2016/02/25      PU4 GroupID�����޸�Ϊline+YY+XXXX,Flag��NewGroupIDRule (1225)
'RQ16021812     Giant           2016/02/29      PU5��DID lable������ӡ�д����ϵ�������Ϻű�����MainPN������ӡʱֱ�Ӷ�ȡSF�ϴ��ĵ�2nd PN��Ӧ�����ϵMain PN��ӡ��û�ж����ֱ�Ӵ�ӡDID�����Ϻ� (1226)
'QMS            Jerry           2016/03/07      ReturnDID��ӡLable������ӡ����<VenderCode1>,��Flag:PrintedVenderCode�ܿ�(1227)
'QMS            Jerry           2016/03/24      ��MSD��Check(1188)����TxtLotCode_KeyPress(1228)
'QMS            Jerry           2016/04/21      ReturnCompPN��ӡLable������ӡ����<VenderCode1>,��Flag:PrintedVenderCode�ܿ�(1229)
'QMS            Newton          2016/05/12      B/S����Barcode�ϵ�Lotcode��ͳһ�޸�Ϊ"@@@@"����ʽ��鵽LotcodeΪ"@@@@"����ΪB/S����(1230)
'QMS            Giant           2016/06/14      �޸�substring(host_name(),1,9)Ϊֱ��ȡhost_name (1231)
'QMS            YanYe           20160627        ���ܼ�1 �������1 ÿ���ϴ����һ��LOCATIONG�����޷��ռ���1232��
'QMS            YanYe           2016/06/27      �Զ��������Ӳ���OldReturnDID ��1233��
'QMS            Kelly           2016/08/11      �޸�QueryInspect��Ĭ��ʱ���Ϊǰһ�쵽���죨1234��
'QMS            Kelly           2016/08/12      �޸�UploadData-> COMPPNINSPECTRULE,ԭ��ͨ��VB�������������ֵ���������ݣ��ָ�Ϊͨ���洢ʵ�֣�1235��
'QMS            Kelly           2016/08/30      InSpection�������Ӳ�����У�1236��
'QMS            Udall           2016/09/20      Report��ѯ������ Type:ReturnDID_ByDate,����Byʱ���ѯ������Ϣ��1237��
'QMS            Yan             20160920        ����FujiNexim�����ϴ�MachineBOM��1238��
'QMS            Yan             20160927        �ڽ���Groupʱ��Ҫ����Ӧ�Ĺ����Ƿ���CHECK BOM PASS (1239)
'QMS            Udall           20161025        ��PU9 ��������������ϵͳΪ׼������DID���ݺ�FujiTrax��RealQty��ƥ������������SPץȡFujiTrax��DIDʣ������ (1240)
'QMS            Jerry           20161104        ����Flag:UnChkBaseReelQty,����ֵΪYʱȡ��ɢ�Ϻϲ�ʱDID��������С��BaseReelQty������ (1241)
'RQ16110703     Season          20161117        ��ӡDIDʱ��Ҫ���ӡ����Item Location��Qty (1242)
'QMS            Light           20161202        PU9����1-A-3-1�������͵Ĵ��� (1243)
'QMS            Yan             20161215        ��DID Label����Ҫ��ʾ����  ��1244��
'QMS            Seven           20161226        ���ڻ�Ϊ���ϲ��ǵ���������,������Ϊ�����߼�(1245)
'QMS            Frame           20161226        ��ֹ1����ť˫������(1246)
'QMS            Frame           20170116        QTY��λ��CompPN���������ӷ���(1247)
'QMS            Kim             20170217        PU5��ά����A�׹��ز�������ʾ:�ò�������հ�װ���˻ؿ�(1248)
'RQ17011010     Friday          20170307        �Ż�CompPrint,���һ����λStandard��ΪSAP���.��װ������ʾ(1249)
'QMS            Sarah           20170308        PU9����Flag��ChkMEBOM_Location�ϴ�MachineBOMʱ��Location��Ϊ�գ�1250��
'QMS            Yan             20170313        PU9����IsTrayComp ���Tray���Ƿ��ж���XY����������1251��
'QMS            Kim             20170320        PU5���ӱ�����WHID�� ��ӡ��Ӧ�ֱ𡣣�1252��
'QMS            Seven           20170321        PU3���Ӱ�JOBGroupɾ��MEBOM�Ĺ���  (1253)
'QMS            Seven           20170417        ��Ϊ����(�͹���)ʱ��Ҫˢ��09Code����(1254)
'RQ17041211     Season          20170420        ��ӡDIDʱ��Ҫ���ӡ����Mark (1255)
'QMS            Sarah           20170512        �޸�bug,QSMS_MEBOM�ְڼ��ϴ����excelȡ����������ֹ����location���ң�1256��
'QMS            Kim             20170522        PU5����CHeckDID���棺DID��ʵ��BarCode���Աȣ�1257��
'QMS            Feix            20170614        PU9���LCK�������ƣ�1258��
'QMS            Kim             20170719        PU5�������ά��BIOS��EC��Ϣ���壨1259��
'QMS            Feix            20170725        ��ӵ�������Slot�����ڿգ�1260��
'QMS            Luck            20170803        CompPrintǰ���Check��1261��
'QMS            Kim             20170808        PU5����DID�Զ�����CompPN,VendorCode,DateCode,LotCode��Ϣ��1262��
'QMS            Ju              20170808        �������BU="NB3"��1263��
'QMS            Ju              20170808        ��ɾ��Machine�д�"Others"��QSMS_MEBOM�����ݣ�1264��
'QMS            Feix            20171026        ��ӵ�������Slot�����ڿգ�1260��
'QMS            Light           20171102        DIP���ϵ����Զ������,�ϴ�Build TypeΪ3�ĵ�����ϴ�Q����ϣ�1261��
'QMS            Kim             20171228        PU5���ӽ����ϱ��������ѣ�1265��
'QMS            Kim             20180118        PU5����HP�Ϻ�תΪ����Ϻţ�1266��
'QMS            Feix            20180403        ESBU�����ʾ��Ϣ��1267��
'QMS            Luck            20180626        ��ӡDIDʱ�Զ���ά������δOpen��DID��¼Open(1268)
'QMS            Seven           20180709        ��ѯ����PU3��Camera��DID��1269��
'QMS            Seven           20181024        ��ѯME_BOM_WO����Location��Ϣ��1270��
'QMS            Feix            20181024        ESBU����µĽ���TransferFujiNexim_MI(1271��
'QMS            YanYe           20181127        ���Location �����Ƿ���QTY��ƥ��    1272
'QMS            Friday          20181129        NB6����ҪCompPrintǰ���Check  (1273)
'QMS            Lucy            20181214        ESBU��Ӵ�ӡʱתCompPN  (1274)
'QMS            Seven           20181230        ���DID����¼���ܣ�1275��
'QMS            Seven           20190307        CheckDID����ѡ���߱��ܣ�1276��
'QMS            Ellen           20190326        ����DID��ӡ����JobGroup��1277��
'QMS            Seven           20190430        �嵥����ų��Ƿ�©����1278��
'QMS            Light           20190522        Add machine hour C/T check (1279)
'QMS            Seven           20190528        HW��DID���ֿ͹��ϣ�1280��
'QMS            Seven           20190615        ���ϵ�����ʾ:Feeder��ʹ�ô��Ե�Ƭ��1281��
'QMS            Rain            20190711        ���Ӵ�ӡ��ϢVendorCode ��1282��
'QMS            Henry           20210104        ����WOGroupʱ�����ָ�����ϣ�1283��
'QMS            Henry           20210219        PMC�ϴ�����ʱ��ָ��PCBVendorCode��1284��
'QMS            Light           20210304        Add ����UniqueID 001401
'QMS            Stephen         20211118        'Add Quanta Delivery Label (1286)
'QMS            Stephen         20220414        MCC->Report->Report->ME_BOM_WO ����COMPPN2 ���� (1287)
'QMS            Henry           20220517        SpecialCase -> GenXLPrior (1288)
'QMS            Henry           20220829        ?����ӡ�ȴ�10�� (1289)
'QMS            Henry           20220829        �Ԅӻ�Modbus�z�y�˻`�Ƿ�˺�� (1290)
'**********************************************************************************/


