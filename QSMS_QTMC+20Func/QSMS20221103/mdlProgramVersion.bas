Attribute VB_Name = "mdlProgramVersion"
'/**********************************************************************************
'**文 件 名: SMT_QSMS.frm
'**Copyright (C) 2007-2010 QMS
'**文件编号:
'**创 建 人: Giant
'**日    期: 2010.08.12
'**描    述: ---to be added
'
'EQMS_ID        '**修 改 人     修改日期        描    述
                '-----------------------------------------------------------------------------
'QMS             Jeanson        2010/08/12     save the modification log into one unified file '(1000)
'                kaitlyn        2010/08/17     add a template of upload_traycompPN            '(1001)
'RQ10061811      Lynn           2010/08/24     Check Forbidden PN when maintain feeder            '(1002)
'QMS             Giant          2010/08/25     Print Bug           '(1003)
'QMS             Jocelyn        2010/08/26     Add appname limit when user login '(1004)
'QMS             Denver         2010/08/27     can not Print DID Label(Return DID/ReturnComp/DIDCallBack)  '(1005)
'QMS             Udall          2010/09/06     增加厂别,针对散料发料 '(1006)
'QMS             Jocelyn        2010/09/08     Add NB7 in '(1007)
'QMS             Bay            2010/09/22     Add display RealQty (1008)
'QMS             Jocelyn        2010/09/23     Add XL_MaxDIDMaintainQty load及Execl功能（1009）
'QMS             Walton         2010/09/27     Modify QSMS Trace report file format from Excel to CSV file '(1010)
'AS              Udall          2010/09/27     针对AP在非"祥龙"材料也需要核对MSD    (1011)
'QMS             Jocelyn        2010/10/12     判断upLoad_MachineType时的execl文件中的SeqIDByLine是否为数字 '(1012)
'                Maggie         2010/10/14     增加新旧料号转换时的Label打印  '(1013)
'QMS             Udall          2010/10/18     Add form function frmDispatchDataToQWMS      (1014)
'QMS             Giant          2010/10/25     Add new column BufferQty for upload XL_WOPlanSeq      (1015)
'RQ10110212      Giant          2010/11/02     Make up some bug for re-print DID      (1016)
'QMS             Jocelyn        2010/11/03     Print buffer not enough,can't print    (1017)
'MFG             Kaitlyn        2010/11/12     Get PN/VendorCode/DateCode/LotCode/Qty from 2D barcode  (1018)
'RQ10110907      Maggie         2010/11/15     Save Printer setting in local Registry (1019)
'QMS             kaitlyn        2010/11/22     get returnDID 's infomation from QSMS_DID_TOWH  (1020)
'QMS             kailtyn        2010/11/24     upLoad machineType:the first letter can't be accepted [0-9]  '(1021)
'QMs             Austin         2010/11/24     输入MSD后，鼠标跳到txtInSpection栏位.  (1022)
'QMS             Austin         2010/11/30     修改强行推料打印，字符50的传         (1023)
'QMS             kaitlyn        2010/12/08     UploadData Add item:Component_Data  (1024)
'QMS             Walton         2010/12/09     UploadData/TransferPanaMSF: the table MEbom add column :line(1025)
'QMS             Jocelyn        2010/12/09     FrmTransferFujiXML/frmTransferPanaAMI: the table MEbom add column :line(1026)
'QMS             Jocelyn        2010/12/14     Print buffer not enough,can't print    (1027)
'QMS             kaitlyn        2010/12/16     uploadData:1.卡functype=sheetname(取消xls的限制）；2.component_data卡item为厂别 （1028）
'QMS             Udall          2010/12/17     Add Side for Machine        (1029)
'QMS             Maggie         2010/12/21     修改Bug              (1030)
'QMS             Giant          2010/12/22     导出MEBOM for 3码线别              (1031)
'QMS             Jocelyn        2010/12/27     check machine whether define or not while uploading machine bom  (1032)
'QMS             Walton         2010/12/27     Modify Report: query Machine Bom by WO            (1033)
'QMS             Bay            2010/12/29     Modify MainTainFeeder :  Add   @machine  (1034)
'QMS             kaitlyn        2010/12/29     update delete me_bom frm   (1035)
'QMS             Jocelyn        2010/12/30     取消线别加machine名来取machine   (1036)
'Qms             Walton         2010/12/30     add line and side           (1037)
'                Maggie         2010/12/31     FrmReturnComp:CCBU只需清空txtCompPN,txtQty   (1038)
'QMS             Udall          2011/01/01     当Machine类型为Pana时核对Machine的MappingID长度必须为7码         (1039)
'QMS             Austin         2011/01/03     暂时取消MappingID规则检查，如何改待定                           (1040)
'QMS             Austin         2011/01/05     修改前截取groupID(2,8),修改后截取GroupID(3,8)                   (1041)
'QMS             Denver         2011/01/04     can not Upload Fuji XML data                                   (1042)
'QMS             Jocelyn        2011/01/05     update method to upload fuji XML data                          (1043)
'QMS             Jocelyn        2011/01/07     Get ComPPN from 2D barcode                                    '(1045)
'QMS             kaitlyn        2011/01/10     upload machinetype:others机台不卡mappingID   (1046)
'QMS             Giant          2011/01/10     NB5 upload MEBOM:*others机台If defined in machine   (1047)
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
'QMS             Austin         2011/05/11     防止退料的时候,连续点击OK按钮,导致产生多个DID,在处理完之前是OK按钮Disable   1058
'QMS             Jocelyn        2011/05/11     多笔查询 '1059
'QMS             Jocelyn        2011/05/17     去除by时间的条件 '1060
'QMS             Jeanson        2011/05/27     query unlink feeder and DID by WO'1061
'CC              Denver         2011/06/21     Add Dummy ECN function
'PO-RQ11062703   Felix          2011/07/04     Adjust DID dspatch label format (1063)
'RQ11060735      Bay            2011/07/14     Add DID and CompPN check Function    (1064)
'RQ11060107      kaitlyn        2011/07/20     upload machinetype:修改检查规则 (1065)
'RQ11071314      Alawn          2011/07/21     MATERIALTOWHID能够上传%的料号     （1066）
'RQ11071314      Alawn          2011/07/21     Add PrintAutoDispatchLabelNetWorkPort     （1067）
'QMS             Kaitlyn        2011/08/10     upload qsms_mebom:不检查machinename规则(1068)
'QMS             Giant          2011/08/16     get reelwidth from [stockdata] for NPM machine(1069)
'QMS             Austin         2011/08/24     Modify upload AMI文件 for NPM DualLane=MIX
'QMS             Bay            2011/08/24     SET.INI add FLAG: DIDScan , 如果DIDScan="Y" QSMS Maintain DID auto Dispatch与Return DID 中的CompPN和DID必须通过 scaner 输入 (1071)
'QMS             Maggie         2011/08/24     修改upload XL_WOPlanSeq，记录DualLane=MIX (1072)
'QMS             Alawn          2011/09/14     针对Buildtype=4的wo，在删除ME_Bom的时候，删除两条线的数据,导出ME_BOM两条线的数据   （1073）
'RQ11090907      Bay            2011/09/25     针对C系列散料，物料将已发料的散料还未发到产线的材料整合后重新注册新的DID并发到产线（产线接散料的动作转移到物料完成）
                                               '目的：1减少产线接料的动作，2降低产线错料的机率(1074)
'QMS             Maggie         2011/10/17     上传AMI文件时检查文件名称,区分是否为NPM DualLaneMode(1075)
'RQ11101304      Alawn          2011/10/25     QSMS Maintain Feeder 中的DID必须通过 scaner 输入 （1076）
'QMS             Jerry          2011/10/25     在Report添加可以查询正常接料和替代料的相关信息          (1077)
'QMS             Bay            2011/11/20     Modify Bug :  SQL语句语法错误(1078)
'QMS             Giant          2011/11/22     Read MCType from QSMS_ProConfig(1079)
'RQ11120813      kaitlyn        2011/12/19     增加200点打印机，修改DID模板选择部分(1080)
'qms             Jerry          2011/12/22     MaintainFeeder时在QSMS_LOG里记录(1081)
'qms             Jerry          2011/12/22     在Report添加可以查询MaintainFeeder的相关信息(1082)
'QMS             Walton         2011/12/28     1>将其登录界面改成通过MainMenu登录。
                                        '''    2>将Factory参数通过MainMenu传，若没有传，则按照旧有的方式Setting.ini传送     ''''(1084)
'QMS             Kaitlyn        2011/12/29     上传PANA me_bom时，修改machinename的检查规则（1083）
'QMS             Jerry          2012/01/12     重庆NB4 在zebra_300_old.txt中增加变量<MACHINE1>（1084）
'QMS             Jerry          2012/02/16     在FrmInSpection界面增加根据COMPPN查询对应电容的功能（1085）
'QMS             Jerry          2012/03/16     重庆NB5 在zebra_300_new.txt中增加变量<SLOT1>（1086）
'QMS             Walton         2012/03/20     当刷入的DID为承接料时需要检查之前所承接的料的RealQty是否大于0.若大于0,则将带出之前的DID .  (1087)
'QMS             Jerry          2012/03/27     1.frmMaintainDIDAutoDispatch界面增加声音提示.  (1088)
'                                              2.frmMaintainDIDAutoDispatch出现错误提示时清空相关信息
'                                              3.frmMaintainDIDAutoDispatch界面在刷入COMPN时，直接跳到TxtGroupQty
'QMS             Jerry          2012/04/17     在Report添加可以查询PDA_DistributeDIDLog的相关信息.  (1089)
'QMS             kaitlyn        2012/04/17     fixed bug:手动承接时by side承接修改line+side的取值，不在用machine前两码表示(1090)
'QMS             Jerry          2012/04/25     1.修正FrmDIDInteGration界面的一个BUG(1091)
'                                              2.更改FrmDIDIGration界面的strDay变量的取值
'                                              3.FrmDIDIGration界面增加变量<MACHINE1>
'QMS             Bay            2012/04/26     使用API播放音频文件 (1092)
'QMS             Walton         2012/05/14     打印DID时若该DID为指定用料则增加工单变量,需要修改SP：[QSMS_GetDIDPrintInfo]，XL_GetDidPrintInfo_Return 可以参照NB7  （1093）
'QMS             kaitlyn        2012/05/16     增13寸料卷，基数为50K，需要修改变量类型为long (1094)
'QMS             Alawn          2012/05/16     修复补印DID记录工号问题                      （1095）
'QMS             Bay            2012/05/31     添加FrmDIDDistribution窗体，用于物料 按照线，面，机台进行放入对应的DID进行分类 （1096）
'RQ12052311      Dreamu         2012/06/07     Get PN/VendorCode/DateCode/LotCode from 2D barcode,打印出的DateCode和LotCode分开两栏 (1097)
'RQ12052204      Dreamu         2012/06/08     修改QSMS\PD\UpdateDIDRealQTY添加“Reason”栏位   (1098)
'QMS             scofield       2012/06/21     modify limit qty from 60000 to 120000 '(1099)
'QMS             Allsa          2012/06/27     发料窗体中CompPN，VendorCode，DateCode，LotCode栏位均不可手Key，且不能用复制和黏贴 '(1100)
                                               '另外在修改需求RQ12052311中将CompPrint界面的原来可以手KeyVendorCode，DateCode，LotCode栏位的取消掉了，但产线上对于这个界面还是需要手Key的，因为这个界面只有指导员才可以登录进来
'RQ12062909      Dreamu         2012/07/10     DID Label上的Data抓取的时间精确到秒。Format=120629080000      (1101)
'RQ12051615      Alawn          2012/07/12     禁止PD\CompPNCompare界面能够复制和粘贴 （1102）
'QMS             Lynn           2012/07/25     增加flag：MaintainFeederDID，如果开启且为fuji的时候，只需要刷入DID,CompPN,Feeder即可 （1103）
'QMS             Walton         2012/07/01     修改IC_CompPN查询通过PN来查且添加IC_Burn的导出模板    (1104)
'QMS             Walton         2012/08/13     打印DID时若该DID为指定用料则增加DIDType变量,需要修改SP：[QSMS_GetDIDPrintInfo]，XL_GetDidPrintInfo_Return 可以参照NB7  （1105）
'QMS             Walton         2012/08/17     增加指定料声音文档        （1106）
'QMS             Lynn           2012/08/29     Return DID的时候，如果上一个DID的数量还大于0，则提示        （1107）
'QMS             Austin         2012/09/24     if MCType = NPMMachineType,ReelWidth取NPMReelWidth        （1108）
'EQMS           kaitlyn         2012/10/16     打印DID时是差异料在DID上显示出来(CYL:相同group，根据祥龙排程，比对当前这个工单和上一个工单新增料的差异）(1109)
'QMS            kaitlyn         2012/10/30     上传的地方将table QSMS_CompPNcheck 改为QSMS_CompPNcheck_temp(因为此table名和PN_Receiving中用到的重名）（1110）
'QMS            kaitlyn         2012/11/13     fixed bug:系统判断不是当天的DID，return时候直接删除DID，如果是当天的不删除，目的是为了防止DID重复.比对的时候比的是YYYYMMDD,但是return时候记录的时间是YYMMDDHHNNSS(1111)
'QMS            Dreamu          2012/11/14     修改程式，使打印Label的时间从服务器上抓取   (1112)
'QMS            kaitlyn         2012/11/20     update sql(同一个DID只能接料到一个predid上，DIDSLOTLINK有卡)，但是对于fujitrax没有卡，所以修改sql，抓取最后一个接料的信息)（1113）
'QMS            Link            2012/12/03     QSMS\MCC\CompPint\中自动从2Dbarcode中获得数据，添加QTY（1114）
'RQ12111406     Cynthia         2012/12/05     列印Label上添加产生的日期和时间，变量为“<DateTime>”(1115)
'RQ12111408     kaitlyn         2012/12/06     returnDID时自动带出realqty并选中 （1116）
'RQ12120309     Dreamu          2012/12/12     NB6在frmCompPrint界面添加打印变量Mark   (1117)
'RQ12120301     Alawn           2012/12/27     添加SAP2查询到GetGroupIDDataByCompPN （1118）
'QMS            Ava             2012/12/27     修改frmCompPrint读取文件的方式  (1119)
'RQ12122103     Allsa           2013/01/15     NB1的一个需求：若刷入的是EMMC的料则要比对ImageVersion和MBPN的对应关系  (1120)
'QMS            Newton          2013/01/24     产销上传排程中新增一个选项XL_WOPLANSEQSHIFTID，新增了一个栏位：ShiftID(1121)
'QMS            Walton          2013/01/29     上传IC_CompPN 时需要检查数据是否存在   (1122)
'QMS            Jerry           2013/01/29     在导出的Excel中新增一个选项XL_WOPLANSEQSHIFTID(1123)
'QMS            Ava             2013/01/31     NB5只获取PP10工单    (1124)
'QMS            Walton          2013/01/31     增加上传IC_ShearPin的选项。 (1125)
'RQ13031414     Link            2013/03/20     mnuMaintainDID权限只负责frmMaintainDID窗体，新增mnuMaintainDIDAutoDispatch权限负责FrmMaintainDIDAutoDispatch窗体  ’（1126）
'QMS            Jerry           2013/03/29     在导出的Excel中新增一个选项GetMEBom_ByGroupID(1127)
'QMS            Ava             2013/04/02     工单在添加到GroupID之前，先做Check动作。 （1128）
'QMS            Jerry           2013/04/11     UploadData Add item:Machine_Data。 （1129）
'QMS            Jerry           2013/04/19     NB4增加根据Type选择手动点几小时XL。 （1130）
'QMS            Ava             2013/05/13     添加根据线别删除MEBOM，userRight：DeleteMeBomByLine。（1131）
'QMS            Scofield        2013/05/13     由于（1085）造成MBU IPQC散料（料号长度都是11码）无法获取正确的CHENUM值，MARK掉Len(lblcomppn) > 12这个判断条件。（1132）
'QMS            Newton          2013/05/23     上传XL_WOPlanSeqShiftID时增加Flag标识那些工单为CTO工单 (1133)
'QMS            Walton          2013/05/27     UploadData 界面上中的UploadIC_ShearPin增加CompPN变量 (1134)
'QMS            Walfan          2013/04/30     上传XL_WOPanSeq时增加Flag标识那些工单为CTO工单 (1135)
'QMS            Jerry           2013/06/17     Delete MEBOM界面删除ME_BOM时可以根据JjobPN (1136)
'QMS            Alawn           2013/07/08     修复CompPNCompare界面中line不匹配的bug （1137）
'QMS            Walfan          2013/07/16     Report界面DispatchDID 查询增加CompPN条件 (1138)
'QMS            Dreamu          2013/08/26     FrmMainWoSeq界面修订BUG,添加查询历史DB   (1139)
'QMS            Jerry           2013/09/02     将FrmUrgentInsertWO的Query按钮下调用的SP改为XL_SpecialCaseByWO_New  (1140)
'QMS            Walton          2013/09/03     添加Machine_Data的模板        (1141）
'QMS            Walton          2013/10/17     在打印DID中增加DIDType 变量   (1142)
'RQ13102407     Allsa           2013/11/01     以前的Fuji设备没有导入FuJiTrax时，产线有上传大于100的slot，为了使DIDSlotLink程式看上去清晰，需要减掉100.
                                               '当前Fuji的NXT设备导入了FujiTrax，其Tray盘的Slot有900以上的，而且此Slot需要和FujiTrax系统做匹配。设备不希望在上传时还需要多一个加100的动作，因此针对NXT设备取消我们程式取消减去100的动作。(1143)
'QMS            Jerry           2013/11/25     修改FrmGenXLMD界面的cboType改由从Tble:XL_TypeDateTime中读取    (1144)
'QMS            kaitlyn         2013/12/06     获取DID的规则中更改：退料的DID和料号都含有-A，需要区分开，否则产生的DID会重复(1145)
'QMS            kaitlyn         2013/12/06     fixed bug:maintain feeder时候需要需要检查LR的合法性，只目前只有0，1，2三种 (1046)
'QMS            Jerry           2013/12/10     增加Flag：PrintedSeqID,管控是否打印DID时增加变量COUNT，用以Mark该DID是这个GoupID中这个Slot，LR的第几盘料，方便产线找料(1147)
'QMS            Jerry           2013/12/10     由于上面（1101）修改，NB4在替换了变量Data后，会遮住后面的内容，所以NB4改回YYYYMMDD格式(1148)
'QMS            Walfan          2014/01/13     修正Bug:上传machine_data定义时，by line删除数据，会造成数据误删 (1149)
'QMS            Scofield        2014/01/20     允许上传MB_REV=空的MEBOM   (1150)
'QMS            Walton          2014/02/10     MaintainFeeder 界面上增加了Check 整个WO的绑定Feeder的情况。以及更新了数据显示部分    (1151)
'QMS            Van             2014/02/11     针对（1150）修改，取消tempversion=""的条件   (1152)
'QMS            Ava             2014/02/17     NB5 closeWO时可选择不删除DID    (1153)
'QMS            Jerry           2014/02/21     UploadData Add item:CompPN_Spacer,上传对应CompPN对应的垫片信息。 （1154）
'QMS            Jerry           2014/03/11     针对1149的修改By line删除数据时再加上FuncType='DualLaneMode' and Item='Independent'两个条件。 （1155）
'QMS            Walton          2014/03/13     取消IC_CompPN上传改成通过程式UploadBasicData上传 。
'QMS            Walton          2014/03/24     取消IC_CompPN的查询                           (1157)
'QMS            Newton          2014/03/28     NB3只获取PP10工单    (1158)
'RQ14040405     Anker           2014/04/14     MaintainFeeder之后点击Check，会把此工单中尚未做MaintainFeeder的机台Slot 、L/R以及材料料号显示在excel中  (1159)
'RQ14031414     kaitlyn         2014/04/16     为了减少退料的次数及实物的挪动,return发料时，先满足原线别需求，如果是K17的退料，Return 时如K17 有需求，先满足K17的需求，再看其它线别有没有需求，没有再退仓。在满足原线别需求时，请先设定面与SLOT，之前在哪个面哪个SLOT 就让他先满足这个面与SLOT ，与承接时逻理一样  (1160)
'RQ14042101     Walton          2014/04/21     增加Fuji 新的机型AIMEX的处理               (1161)
'RQ14042802     Dreamu          2014/04/29     修改ReturnDID界面，在Return DID时，添加独用料提示信息功能  (1162)
'QMS            Newton          2014/05/16     修改CompPrint与ReturnDID功能，程式直接从点料机的Com口中读取Qty  (1163)
'RQ14052112     Cynthia         2014/06/08     修改通过BuildType=4(跨线面生产)的WO获取Line的方式，使得可以直接摆到WO查询和删除2条线的ME_Bom数据 （1164）
'QMS            Cynthia         2014/06/10     增加传送chkDomain参数，非Quantan域的电脑则统一export为html格式 （1165）
'QMS            Newton          2014/06/11     产线考虑到需要为点料机增加Shopfloor电脑，取消(1163) （1166）
'QMS            Cynthia         2014/06/12     支持多个结果集导出为Html格式文件 （1167）
'QMS            Cynthia         2014/06/13     修改Report界面的CboReportType从Program_DefineItem中获取 （1168）
'QMS            Cynthia         2014/06/14     使用PrepareMaterialByWONew代替PrepareMaterialByWO,取消模板 (1169)
'QMS            Cynthia         2014/06/16     转换Html时，对于Null值统一更新为"",如果第一位为"=",则在统一加上[].（1170）
'QMS            Cynthia         2014/06/18     统一使用utf-8字符集exprot to html （1171）
'QMS            Ava             2014/07/11     NPM机器上传MEBOM，增加Skip区块、站位及元件  （1172）
'QMS            Lynn            2014/07/15     上传祥龙排程的地方增加检查，不允许在执行祥龙JOB的前后15分钟上传排程  （1173）
'QMS            Newton          2014/08/05     PT200设备软件升级后PositionData块新增了BG和PROGDIST列 （1174）
'QMS            Giant           2014/08/15     NB5有一种情况：大板上没料，但是小板上有。所以不能再以主工单为准，取消条件InitAOIFlag='Y'。 （1175）
'RQ14082505     Alawn           2014/09/03     不允许xl材料使用MCCPreMaterial发料  1176
'RQ14091504     Alawn           2014/09/22     即Feeder绑定DID时检查必须按DID sequence操作。从上料到接料都实现一条龙按DID sequence操作，让OP操作更简单化，更好地避免错料发生 1177
'RQ14090106     Anker           2014/10/09     上传PanaAMI的文档时，在数组中根据字段所在的下标来抓取值  （1178）
'QMS            Jerry           2014/11/05     针对NB4，Report界面增加checkBox:Dual来区分CheckBom时是按照独立还是共享模式Check  （1179）
'RQ14110502     Cynthia         2014/11/19     修改CompPrint界面打印料号Label的QTY，由原先Integer改为Long，方便产线打印大基数料号（1180）
'QMS            Sarah           2014/11/19     NB6新增分仓启动界面 FrmStartSplitLineMC （1181）
'QMS            Ava             2014/11/21     上传MEBOM修改，[Machines]的MCNAME列，对于NPM机的MCNAME只截取NPM及之前的部分   (1182)
'QMS            Ava             2014/11/21     上传MEBOM修改，Tray盘材料的料号后会有"-T"和"-N"，抓取料号时只抓取11位的料号 (1183)
'QMS            Ava             2014/11/27     1178） SD(idxSD).TA = Trim(GetPosition(SD_Header, "TA"))未加Arry   （1184）
'QMS            Alawn           2014/12/09     RQ14111916  定义工单BuildType=4 时，增加Station设置  1185
'QMS            Jerry           2014/12/19     将MaintainFeeder界面的TxtJobGroup改回原来的cboJobGroup  1186
'QMS            Ava             2014/12/23     MaintainFeeder界面TxtCompPN从刷入的DID中获取，增加flag:UnChkCompPN    1187
'QMS            Ava             2015/01/06     增加flag：checkNeedMSD，如果为Y，则提示该材料需烘烤 1188
'QMS            Walton          2015/01/22     将查询JOBPN改成SP方式    1189
'QMS            Scofield        2015/02/03     针对Xboard允许在CLOSEWO时输入Xboard数量，回收对应的C面和S面材料消耗    1190
'RQ14122224     Cynthia         2015/03/04     对于MSD材料，新增Flag=CheckMSDCallBack，开启flag则CallBack的NewDID开封时间承接OLD DID的开封时间  1191
'QMS            Walton          2015/03/09     增加在打印DID 时如果DID为烧录材料，则DID Text 中增加BurnRev 信息      1192
'QMS            Udall           2015/03/13     对于重庆Tray盘料带-T or -N的料,其截取条件进行修正,只有在长度<>11 and <>14时才进行      1193
'QMS            Alawn           2015/03/30     上传MEBOM，增加等待1s 1194
'QMS            Jerry           2015/03/31     UpLoadData增加上传AVLC的功能 1195
'QMS            Walton          2015/04/01     增加DID_2D打印的变量    1196
'RQ15032012     Walton          2015/04/03     ReturnDID 时若DID未被使用，则提示物料  1197
'QMS            Jerry           2015/04/03     NB5增加按JOBGroup删除MEBOM的功能  1198
'QMS            Jerry           2015/04/03     增加Flag:NoKeepPWD，Check是否需要记录密码  1199
'RQ15041004     Cynthia         2015/05/06     MSD对于Return/CallBack后再发料的，检查对应的DID的状态，不符合则不能发料 1200
'QMS            Frame           2015/06/02     增加3523型号LCR量测仪  1201
'QMS            Frame           2015/06/03      输入DID之后会去调用callPicture函数，如果没有图片则退出Function继续往下执行  1202
'QMS            Jerry           2015/06/16      上传ME_BOM时抓取[BoardData]中的PCBSize，即X="###",并记录进系统 1203
'QMS            Newton          2015/07/29      增加8110G型号LCR量测仪统 1204
'QMS            Newton          2015/08/05      重庆NB3增加一个退料仓，PE可以自行选择退料到哪个仓别 1205
'QMS            Frame           2015/08/24      增加BY WO和CompPN 批量查询的功能 1206
'QMS            Eason           2015/09/09      重庆NB5增加将物料离职员工绑定的DID的UID更换为在职员工的UID  1207
'QMS            Newton          2015/10/13      重庆NB3 B/S材料Barcode上的Datecode是广达Datecode，在发料时要刷入料盘上的来料Datecode并记录  1208
'RQ15101219     Yan             20151023        User在删除工单点击‘Detele’时程式去检查当前工单是否上传过XL,如果有当前工单是不可以删除的 1209
'RQ15101203     Frame           2015/10/27      在物料备料时提前检查Feeder的状态，对需要保养或未在当线FujiTraxServer注册的Feeder进行相应的提示(只针对Fuji的Feeder) 1210
'QMS            Eason           2015/10/30      增加Flag ChkPNCQ 判断如果是重庆的料号不做处理 1211
'QMS            Eason           2015/11/12      PU5在report增加CheckWO_WastagePN选项，供查询高抛料和高频料。 1212
'QMS            Newton          2015/11/15      由PE上传B/S 材料料号，系统自动判断  1213
'QMS            Udall           2015/11/30      在table:NozzleLocation中增加Line,Side,BuildType等数据   1214
'QMS            Frame           2015/12/04      增加判断,<fsSetPos>和</fsSetPos>必须为 数字 - 数字 模式 1215
'QMS            Jerry           2015/12/17      改善(1085),对于刷入的DID自动截取前11码 1216
'QMS            Giant           2015/12/18      [PartsData]部分存在不连续的IDNO，需要排除为空的部分 1217
'QMS            Season          2015/11/12      承接料的时候，如果DID之前没有发过这个位置，则这个DID不能承接到这个位置 ----(1218)
'QMS            Giant           2015/12/29      PU5需要修改上传MEBOM的文件名格式，并可以从Report中导出保存的数据，添加FLAG:ChkEQProgram ----(1219)
'RQ16010604     Kver            2016/01/06      修改界面FrmInheritDIDByWO，只显示未关结的工单Group  ----（1220）
'QMS            Austin          2016/01/10      退料的时候，根据定义的仓别，判断是否存在料号相同，但是D\C不同的数据，如果有，则提示 (1221)
'RQ16011311     Yan             2016/01/19      系统增加DA0料使用时间卡关，超过半年的DA0料无法使用(1222)
'QMS            Jerry           2016/01/21      DIDLabel新增打印变量<VenderCode1>和<LR1>,由Flag:PrintedVenderCode管控(1223)
'RQ16012215     Kver            2016/01/25      更改Maintain DID中打印PU5时间格式为"YYYYMMDDHHNNSS" (1224)
'QMS            Jerry           2016/02/25      PU4 GroupID规则修改为line+YY+XXXX,Flag：NewGroupIDRule (1225)
'RQ16021812     Giant           2016/02/29      PU5在DID lable新增打印有搭配关系的主料料号变量＜MainPN＞，打印时直接读取SF上传的的2nd PN对应搭配关系Main PN打印，没有定义的直接打印DID材料料号 (1226)
'QMS            Jerry           2016/03/07      ReturnDID打印Lable新增打印变量<VenderCode1>,由Flag:PrintedVenderCode管控(1227)
'QMS            Jerry           2016/03/24      将MSD的Check(1188)移至TxtLotCode_KeyPress(1228)
'QMS            Jerry           2016/04/21      ReturnCompPN打印Lable新增打印变量<VenderCode1>,由Flag:PrintedVenderCode管控(1229)
'QMS            Newton          2016/05/12      B/S材料Barcode上的Lotcode会统一修改为"@@@@"，程式检查到Lotcode为"@@@@"就判为B/S材料(1230)
'QMS            Giant           2016/06/14      修改substring(host_name(),1,9)为直接取host_name (1231)
'QMS            YanYe           20160627        不能减1 ，如果减1 每次上传最后一个LOCATIONG都是无法收集（1232）
'QMS            YanYe           2016/06/27      自动发料增加参数OldReturnDID （1233）
'QMS            Kelly           2016/08/11      修改QueryInspect，默认时间改为前一天到今天（1234）
'QMS            Kelly           2016/08/12      修改UploadData-> COMPPNINSPECTRULE,原先通过VB代码计算上下限值并保存数据，现改为通过存储实现（1235）
'QMS            Kelly           2016/08/30      InSpection界面增加测量电感（1236）
'QMS            Udall           2016/09/20      Report查询中新增 Type:ReturnDID_ByDate,用于By时间查询退料信息（1237）
'QMS            Yan             20160920        增加FujiNexim机型上传MachineBOM（1238）
'QMS            Yan             20160927        在建立Group时需要检查对应的工单是否在CHECK BOM PASS (1239)
'QMS            Udall           20161025        因PU9 导入退料数量以系统为准，出现DID数据和FujiTrax中RealQty不匹配的情况，增加SP抓取FujiTrax中DID剩余数量 (1240)
'QMS            Jerry           20161104        增加Flag:UnChkBaseReelQty,当其值为Y时取消散料合并时DID数量必须小于BaseReelQty的限制 (1241)
'RQ16110703     Season          20161117        打印DID时需要多打印两个Item Location和Qty (1242)
'QMS            Light           20161202        PU9增加1-A-3-1这种类型的处理 (1243)
'QMS            Yan             20161215        在DID Label上需要显示中文  （1244）
'QMS            Seven           20161226        由于华为退料仓是单独出来的,新增华为退料逻辑(1245)
'QMS            Frame           20161226        防止1个按钮双击两次(1246)
'QMS            Frame           20170116        QTY栏位和CompPN的输入增加防呆(1247)
'QMS            Kim             20170217        PU5有维护的A阶贵重材料需提示:该材料需真空包装再退回库(1248)
'RQ17011010     Friday          20170307        优化CompPrint,添加一个栏位Standard作为SAP规格.包装规格的显示(1249)
'QMS            Sarah           20170308        PU9增加Flag：ChkMEBOM_Location上传MachineBOM时卡Location不为空（1250）
'QMS            Yan             20170313        PU9增加IsTrayComp 检查Tray料是否有定义XY的数量。（1251）
'QMS            Kim             20170320        PU5增加变量《WHID》 打印对应仓别。（1252）
'QMS            Seven           20170321        PU3增加按JOBGroup删除MEBOM的功能  (1253)
'QMS            Seven           20170417        华为发料(客供料)时需要刷入09Code功能(1254)
'RQ17041211     Season          20170420        打印DID时需要多打印变量Mark (1255)
'QMS            Sarah           20170512        修复bug,QSMS_MEBOM手摆件上传针对excel取消排序处理，防止料与location错乱（1256）
'QMS            Kim             20170522        PU5增加CHeckDID界面：DID与实物BarCode作对比（1257）
'QMS            Feix            20170614        PU9添加LCK次数限制（1258）
'QMS            Kim             20170719        PU5添加增加维护BIOS和EC信息窗体（1259）
'QMS            Feix            20170725        添加导出条件Slot不等于空（1260）
'QMS            Luck            20170803        CompPrint前添加Check（1261）
'QMS            Kim             20170808        PU5增加DID自动带出CompPN,VendorCode,DateCode,LotCode信息（1262）
'QMS            Ju              20170808        添加条件BU="NB3"（1263）
'QMS            Ju              20170808        不删除Machine中带"Others"的QSMS_MEBOM表数据（1264）
'QMS            Feix            20171026        添加导出条件Slot不等于空（1260）
'QMS            Light           20171102        DIP材料导入自动插件机,上传Build Type为3的单面板上传Q面材料（1261）
'QMS            Kim             20171228        PU5增加禁用料报错弹框提醒（1265）
'QMS            Kim             20180118        PU5增加HP料号转为广达料号（1266）
'QMS            Feix            20180403        ESBU添加提示信息（1267）
'QMS            Luck            20180626        打印DID时自动对维护过的未Open的DID记录Open(1268)
'QMS            Seven           20180709        查询带出PU3和Camera的DID（1269）
'QMS            Seven           20181024        查询ME_BOM_WO带出Location信息（1270）
'QMS            Feix            20181024        ESBU添加新的界面TransferFujiNexim_MI(1271）
'QMS            YanYe           20181127        检查Location 数据是否与QTY相匹配    1272
'QMS            Friday          20181129        NB6不需要CompPrint前添加Check  (1273)
'QMS            Lucy            20181214        ESBU添加打印时转CompPN  (1274)
'QMS            Seven           20181230        添加DID烤记录功能（1275）
'QMS            Seven           20190307        CheckDID新增选择线别功能（1276）
'QMS            Ellen           20190326        增加DID打印变量JobGroup（1277）
'QMS            Seven           20190430        插单检查排除是否漏传（1278）
'QMS            Light           20190522        Add machine hour C/T check (1279)
'QMS            Seven           20190528        HW退DID区分客供料（1280）
'QMS            Seven           20190615        备料弹窗提示:Feeder需使用磁性垫片（1281）
'QMS            Rain            20190711        增加打印信息VendorCode （1282）
'QMS            Henry           20210104        加入WOGroup时，检查指定发料（1283）
'QMS            Henry           20210219        PMC上传祥龙时，指定PCBVendorCode（1284）
'QMS            Light           20210304        Add 解析UniqueID 001401
'QMS            Stephen         20211118        'Add Quanta Delivery Label (1286)
'QMS            Stephen         20220414        MCC->Report->Report->ME_BOM_WO 新增COMPPN2 對應 (1287)
'QMS            Henry           20220517        SpecialCase -> GenXLPrior (1288)
'QMS            Henry           20220829        ?死列印等待10秒 (1289)
'QMS            Henry           20220829        自動化Modbus檢測標籤是否撕下 (1290)
'**********************************************************************************/


