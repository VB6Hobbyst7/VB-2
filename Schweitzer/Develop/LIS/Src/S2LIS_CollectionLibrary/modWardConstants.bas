Attribute VB_Name = "modWardConstants"
Option Explicit

'Global objS2Code As New clsHosComCode
'Global objSysInfo As New clsS2DSO
'Global objInitLPFactory As clsInitLPFactory

Global objAPSbarcode As New clsBarcode
Global objBBSbarcode As New clsBarcode
Global objLISbarcode As New clsBarcode
Global blnAPSBarFg As Boolean
Global blnBBSBarFg As Boolean
Global blnLISBarFg As Boolean

'Global MyUser As New clsEmployee
Global BuildingCd As String     '-- �ǹ��ڵ�
Global BuildingNm As String     '-- �ǹ���
Global BuildingNo As Integer    '-- �ǹ���ȣ
Global FileServer As String     '-- File Server IP

Global SB_ServerNm As String    '-- ������
Global SB_DatabaseNm As String  '-- ����Ÿ���̽���
Global SB_LoginId As String     '-- �α�Ƶ�
Global SB_Password As String    '-- �н�����

Global SB_ConnStatus As Integer

Global Const HospitalNm = "��õ�ǰ����� �μ� �溴��"
Global Const CentralLab = "10"   '-- �߾Ӱ˻��
Global Const CentralLabNm = "�߾�"  '-- �߾Ӱ˻��
Global Const CentralNo = 1      '-- �߾Ӱ˻��
Global Const EmergencyNo = 5        '-- ���޼���
Global Const EmergencyLab = "50"    '-- ���޼���
Global Const EmergencyLabNm = "���޼���"   '-- ���޼�Ÿ
Global Const AneLab = "40"    '-- ���̼���
Global Const AneLabNm = "���̼���"    '-- ���̼���
Global Const WomLab = "20"          '-- ����Ŭ����
Global Const WomLabNm = "����Ŭ����"  '-- ����Ŭ����
Global Const HrtLab = "30"          '-- ���弾��
Global Const HrtLabNm = "���弾��"  '-- ���弾��


'Global declare the data class
Global Const DatabaseName$ = "Lab"
Global Const Connect$ = "Lab/Lab"
Global Const ConnectString = "dsn=sybaseODBC;uid=hisbase;pwd=hispass;"
'
'Tab Item Constants
Global Const Cs_Tab_Collect = 1
Global Const Cs_Tab_Result = 2
Global Const Cs_Tab_Micro = 3
Global Const Cs_Tab_Review = 4
Global Const Cs_Tab_QC = 5
Global Const Cs_Tab_Manager = 6
Global Const Cs_Tab_Statistic = 7


'��������===============>>>> ���İ��� ��.
Global Const HosptGb = "10"

Global Const CCD_ChgBldInfo = "change building info"
Global Const CCD_ChgDBInfo = "change database"
                                          
Global Const CS_BarFormat = "0#########"   '10�ڸ�

Global Const CS_DateMask = "0###-##-##"
Global Const CS_DateLMask = "0###/##/##"
Global Const CS_DateSMask = "0#-##-##"
Global Const CS_TimeSMask = "0#:##"
Global Const CS_TimeLMask = "0#:##:##"
Global Const CS_BlankMask = "____/__/__"
Global Const CS_DateSFormat = "YY-MM-DD"
Global Const CS_DateFormat = "YYYY-MM-DD"
Global Const CS_TimeSFormat = "HH:MM"
Global Const CS_TimeLFormat = "HH:MM:SS"
Global Const CS_DateDbFormat = "YYYYMMDD"
Global Const CS_TimeDbFormat = "HHMMSS"
Global Const CS_LabDay = "YYYYMMDD"     '�ϴ���
Global Const CS_LabMonth = "YYYYMM"     '������
Global Const CS_LabYear = "YYYY"        '�����

' Interface ���� ����Ÿ ���
Global Const CS_EqpError = "ERR"

'****************  Sybase DB ���� System Data & Time ���ϴ� �Լ�  ****************'
Global Const CS_SybaseDate = "convert(char(8),getdate(),112)"
Global Const CS_SybaseTime = "substring(convert(char(8),getdate(),108),1,2)+" & _
                             "substring(convert(char(8),getdate(),108),4,2)+" & _
                             "substring(convert(char(8),getdate(),108),7,2)"
'***********************************************************************************'

' File Server Path
Global Const RegHdSet As String = "SchLis"
Global Const RegSsSet As String = "Setup"
Global Const RegK1Set As String = "SvrIP"

' App Path
Global Const RegHdApp As String = "SchLis"
Global Const RegSsApp As String = "App"
Global Const RegK1App As String = "Path"

' Registry ��� (�������� �ɼ�)
Global Const RegHdOpt As String = "SchLis"
Global Const RegSsOpt As String = "Options"
Global Const RegK1Opt As String = "ShowAtStart"

' Registry ��� (�ǹ�����)
Global Const RegHdBld As String = "SchLis"
Global Const RegSsBld As String = "Building"
Global Const RegK1Bld As String = "Key"
Global Const RegK2Bld As String = "Name"
Global Const RegK3Bld As String = "No"

' Registry ��� (����Ÿ���̽�����)
Global Const RegHdSvr As String = "SchLis"
Global Const RegSsSvr As String = "Server"
Global Const RegK1Svr As String = "Key"
Global Const RegK2Svr As String = "DB"
Global Const RegK3Svr As String = "UID"
Global Const RegK4Svr As String = "PWD"

' ���̺��� ��� ( LIS Tables )
Global Const TB_HIS001 = "h1ptntinfo"   'ȯ�ڱ⺻������
Global Const TB_HIS002 = "h1admin"      'ȯ�ڱ⺻������
Global Const TB_HIS003 = "hzdept"       '�μ�������
Global Const TB_HIS004 = "hzdept"       '����������
Global Const TB_HIS005 = "HIS005"       '���󸶽���
Global Const TB_HIS007 = "hzempl"       '�ǻ縶����
Global Const TB_HIS008 = "HIS008"       '�󺴸�����
Global Const TB_HIS009 = "h1actmat"     '����������

Global Const TB_LAB001 = "h7lab001"     '�˻��׸񸶽���
Global Const TB_LAB002 = "h7lab002"     '�˻絿�Ǿ����
Global Const TB_LAB003 = "h7lab003"     '�˻纰��񸶽���
Global Const TB_LAB004 = "h7lab004"     '������ü������
Global Const TB_LAB005 = "h7lab005"     '����ġ������
Global Const TB_LAB006 = "h7lab006"     '��񸶽���
Global Const TB_LAB007 = "h7lab007"     '�ܺΰ˻縶����
Global Const TB_LAB008 = "h7lab008"     'Worksheet������
Global Const TB_LAB009 = "h7lab009"     '�������׳���
Global Const TB_LAB011 = "h7lab011"     'QC������
Global Const TB_LAB012 = "h7lab012"     'QC�˻�����������
Global Const TB_LAB013 = "h7lab013"     '�̻���QC������
Global Const TB_LAB014 = "h7lab014"     'QC��Ʈ�Ѹ�����
Global Const TB_LAB015 = "h7lab015"     '����������

Global Const TB_LAB021 = "h7lab021"     'QC Control Master
Global Const TB_LAB022 = "h7lab022"     'QC Item Master
Global Const TB_LAB023 = "h7lab023"     'QC Master
Global Const TB_LAB024 = "h7lab024"     'QC Item Information
Global Const TB_LAB025 = "h7lab025"     'QC Schedule
Global Const TB_LAB026 = "h7lab026"     'QC �������
Global Const TB_LAB027 = "h7lab027"     'QC ��������
Global Const TB_LAB028 = "h7lab028"     'QC Text ����

Global Const TB_LAB031 = "h7lab031"     '�����ڵ帶����1
Global Const TB_LAB032 = "h7lab032"     '�����ڵ帶����2
Global Const TB_LAB033 = "h7lab033"     '�����ڵ帶����3
Global Const TB_LAB034 = "h7lab034"     '�����ڵ帶����3
Global Const TB_LAB035 = "h7lab035"     '���ø�������
Global Const TB_LAB036 = "h7lab036"     '��Ÿ�˻����ø�������
Global Const TB_LAB099 = "h7lab099"     '��ȣ�ο�������

Global Const TB_LAB101 = "h7lab101"     'ó��Header
Global Const TB_LAB102 = "h7lab102"     'ó��Body
Global Const TB_LAB103 = "h7lab103"     'QCó��Body

Global Const TB_LAB201 = "h7lab201"     'ä����������
Global Const TB_LAB202 = "h7lab202"
Global Const TB_LAB203 = "h7lab203"     '���Ӱ˻系��
Global Const TB_LAB204 = "h7lab204"     '�ϰ�ä������
Global Const TB_LAB205 = "h7lab205"     '�ܺ��Ƿڳ���

Global Const TB_LAB301 = "h7lab301"     'Worksheet����
Global Const TB_LAB302 = "h7lab302"     '�Ϲݰ������
Global Const TB_LAB303 = "h7lab303"     '�Ϲ��ؽ�Ʈ�������
Global Const TB_LAB304 = "h7lab304"     'FootNote����
Global Const TB_LAB305 = "h7lab305"     'Supplemental����
Global Const TB_LAB306 = "h7lab306"     '�ڵ�ȭ��� ���۳���
Global Const TB_LAB307 = "h7lab307"     'QC�������
Global Const TB_LAB308 = "h7lab308"     '�Ϲݰ����������

Global Const TB_LAB350 = "h7lab350"     '��Ÿ�˻缳������
Global Const TB_LAB351 = "h7lab351"     '��Ÿ�˻�������
Global Const TB_LAB352 = "h7lab352"     '��Ÿ�˻�Numeric���
Global Const TB_LAB353 = "h7lab353"     '��Ÿ�˻�Text���
Global Const TB_LAB354 = "h7lab354"     '��Ÿ�˻��������

Global Const TB_LAB401 = "h7lab401"     '�̻���Worksheet����
Global Const TB_LAB401A = "h7lab401A"   '�̻���Worksheet����(Body)
Global Const TB_LAB401B = "h7lab401B"   '�̻���Worksheet���ܳ���
Global Const TB_LAB402 = "h7lab402"     '�̻���Worksheet����
Global Const TB_LAB403 = "h7lab403"     '�̻���Growth Reading����
Global Const TB_LAB404 = "h7lab404"     '�̻����������
Global Const TB_LAB405 = "h7lab405"     '�̻����������������
Global Const TB_LAB406 = "h7lab406"     '�̻���QC�������
Global Const TB_LAB407 = "h7lab407"     '�̻�����������

Global Const TB_LAB601 = "h7lab601"     '���������
Global Const TB_LAB602 = "h7lab602"     '�����������������
Global Const TB_LAB603 = "h7lab603"     '������µ���������

Global Const TB_LAB999 = "LIS99_DB..h7lab999"     '�� system �������


' �����ڵ�1 (LAB031) Index ���
Global Const CD1_Index = "C100"
Global Const CD1_Panel = "C101"         ' Paneló�� Item
Global Const CD1_MultiSpc = "C102"      ' ������ü
Global Const CD1_Detail = "C103"        ' Detail Items
Global Const CD1_KeyMap = "C104"        ' Keyboard mapping
Global Const CD1_AttrItem = "C105"      ' �Ӽ� ���� Item
Global Const CD1_SpcMedia = "C106"      ' ��ü�� - ����
Global Const CD1_MediaBio = "C107"      ' ���� - Bio Chemical Item
Global Const CD1_MicroAnti = "C108"     ' ���� - �׻���
Global Const CD1_Machine = "C109"       ' ��� - Item
Global Const CD1_ItemResult = "C110"    ' Item - ����ڵ�
Global Const CD1_WAResult = "C111"      ' WorkArea - ����ڵ�
Global Const CD1_QcControl = "C112"     ' QC Control
Global Const CD1_MBatchRst = "C113"     ' �̻��� �p� - ��ġ ��� �ڵ�
Global Const CD1_RelTest = "C114"       ' ���ð˻��ڵ�
Global Const CD1_ColListTm = "C115"     ' �ǹ��� ä������Ʈ ��½ð�
Global Const CD1_CumItem = "C116"       ' ���������ȸ Item

' �����ڵ�2 (LAB032) Index ���
Global Const CD2_Index = "C200"
Global Const CD2_DrGrade = "C201"       ' �ǻ�Grade
Global Const CD2_BedGrade = "C202"      ' ������
Global Const CD2_BedStatus = "C203"     ' �������
Global Const CD2_DeptDiv = "C204"       ' ���з�
Global Const CD2_HighItem = "C205"      ' �ٺ�ó��
Global Const CD2_PocItem = "C206"       ' Point of Care
Global Const CD2_Bypass = "C207"        ' Bypass
Global Const CD2_RoundTime = "C208"     ' Roundä�� �ð���
Global Const CD2_ColTeam = "C209"       ' ä����
Global Const CD2_OutLab = "C210"        ' �ܺ��Ƿ�ó
Global Const CD2_RefLab = "C211"        ' Referral Lab
Global Const CD2_Vander = "C212"        ' Vander �ڵ�
Global Const CD2_WorkArea = "C213"      ' Work Area
Global Const CD2_Section = "C214"       ' Section
Global Const CD2_Specimen = "C215"      ' ��ü
Global Const CD2_VerifyFg = "C216"      ' Auto Verify On/Off
Global Const CD2_SGroup = "C217"        ' ��ü��
Global Const CD2_Media = "C218"         ' ����
Global Const CD2_Microbe = "C219"       ' ��
Global Const CD2_Species = "C220"       ' ����
Global Const CD2_AntiBiotic = "C221"    ' �׻���
Global Const CD2_BioChemical = "C222"   ' ��ȭ���� �����˻�
Global Const CD2_Volume = "C223"        ' �����ڵ�
Global Const CD2_Infect = "C224"        ' ����������
Global Const CD2_QCOrderTime = "C225"   ' QC�ڵ�ó�� �ð���
Global Const CD2_BedDiv = "C226"        ' �����з�
Global Const CD2_NoGrowth = "C227"      ' �̻��� Nogrowth Code
Global Const CD2_WorkSheetName = "C228" ' ��ũ��Ʈ �̸�
Global Const CD2_StoreCd = "C229"       ' ��������
Global Const CD2_Buildings = "C230"     ' �ǹ��ڵ�
Global Const CD2_MWSKinds = "C231"      ' �̻��� �p�� ����
Global Const CD2_FileServer = "C232"    ' File Server Location
Global Const CD2_StaticItem = "C233"    ' ���� ��� �׸�
Global Const CD2_StaticGroup = "C234"   ' ���� ��� Workarea
Global Const CD2_PrinterId = "C235"     ' Printer ID
Global Const CD2_StartDate = "C236"     ' ���� �˻��Ⱓ ����


' �����ڵ�3 (LAB033) Index ���
Global Const CD3_Index = "C300"
Global Const CD3_ScrLock = "C301"       ' Screen Lock Interval
Global Const CD3_PrgOnOff = "C302"      ' Program On/Off
Global Const CD3_FnctOnOff = "C303"     ' Fuction On/Off
Global Const CD3_InfectCond = "C304"    ' �������� ����
Global Const CD3_BarFormat = "C305"     ' Barcode Label Format
Global Const CD3_BarTime = "C306"       ' ���Ӱ˻� Barcode Label ��½���
Global Const CD3_WSPrtTime = "C307"     ' ��Ÿ�˻� Worksheet ��½���
Global Const CD3_Hospital = "C308"      ' �����̸�, �ּ�, �˻���̸�
Global Const CD3_CumulTime = "C309"     ' ������� ��½���
Global Const CD3_LabelTime = "C310"     ' ���� Label ��½���
Global Const CD3_TempUnit = "C311"      ' ����� �µ� ����
Global Const CD3_DateFormat = "C312"    ' ��¥ Format
Global Const CD3_TimeFormat = "C313"    ' �ð� Format

' Template (LAB034) Index ���
Global Const CD4_Index = "C400"
Global Const CD4_Morphology = "C401"    ' �� ����
Global Const CD4_UncolReason = "C402"   ' ��ä�� ����
Global Const CD4_Remark = "C403"        ' Remark
Global Const CD4_FootNote = "C404"      ' Foot Note
Global Const CD4_WarnInfect = "C405"    ' Warning/Infection
Global Const CD4_TextResult = "C406"    ' Text ���
Global Const CD4_SPTextResult = "C407"  ' ��Ÿ�˻� Text ���
Global Const CD4_DCReason = "C408"      ' ó����� ����
Global Const CD4_CancelReason = "C409"  ' ������� ����
Global Const CD4_ModifyReason = "C410"  ' ������� ����
Global Const CD4_QCRejReason = "C411"   ' QC Reject ����
Global Const CD4_TempReason = "C412"    ' �µ��� Reject ����
Global Const CD4_ClinicalNotice = "C413" ' Clinical Notice

'Help Context ID
Global Const HLP_LogOn = 1003
Global Const HLP_EmpMaster = 1004
Global Const HLP_Order = 1011
Global Const HLP_Round = 1012
Global Const HLP_NulCol = 1013
Global Const HLP_Access = 1014
Global Const HLP_SendPt = 1015
Global Const HLP_Referral = 1016
Global Const HLP_BarReprint = 1017
Global Const HLP_WSBuild = 1021
Global Const HLP_AccEntry = 1022
Global Const HLP_InstEntry = 1023
Global Const HLP_WSEntry = 1024
Global Const HLP_ItemEntry = 1025
Global Const HLP_Modify = 1026
Global Const HLP_RstView = 1041
Global Const HLP_AllRstView = 1042
Global Const HLP_ItemMaster = 1061
Global Const HLP_SpcMaster = 1062
Global Const HLP_RefMaster = 1063
Global Const HLP_WSMaster = 1064

Global Const LIS_ORDDIV = "L"
Global Const BBS_ORDDIV = "B"

Global Const Abnormal_High = "��"
Global Const Abnormal_Low = "��"
Global Const Abnormal_Flag = "*"
Global Const Abnormal_Delta = "D"
Global Const Abnormal_Panic = "P"


'�˻�����
Global Const TST_RouTest = "0"      ' ��κ� �˻�
Global Const TST_SpeTest = "1"      ' Ư�� �˻�
Global Const TST_MicTest = "2"      ' �̻��� �˻�
Global Const TST_BldTest = "3"      ' ������ �˻�

' ��ȣ�ο�����
Global Const NO_Specimen = "01"     '��ü��ȣ
Global Const NO_LabNo = "02"        '������ȣ
Global Const NO_WorkNo = "03"       '�Ϲ�Worksheet Unit
Global Const NO_WSUnit = "09"       '�̻���Worksheet Unit

' BussDiv
Global Const CS_BussOut = "1"       '�ܷ�ȯ��
Global Const CS_BussIn = "2"        '�Կ�ȯ��
Global Const CS_BussEr = "3"        '����ȯ��

' Panel Flag
Global Const PN_Group = "G"         'Group Item
Global Const PN_Detail = "D"        'Detail Item
Global Const PN_Normal = ""         '�Ϲ� Item

' FootNote ���� (FootNoteFg in lab201)
Global Const RST_FootNote = "Y"

' Status Code
Global Const STS_Order = "0"        'ó��
Global Const STS_HaveSpc = "1"      'ä��
Global Const STS_Access = "2"       '����
Global Const STS_Worksheet = "3"    'In-Process
Global Const STS_MidRst = "4"       'Partial Verify / �߰����
Global Const STS_FinRst = "5"       'Ȯ�� / ����Ȯ��
Global Const STS_Modify = "6"       '����

Global Const STNM_Order = "ó��"
Global Const STNM_Collect = "ä��"
Global Const STNM_Access = "����"
Global Const STNM_Worksheet = "�˻���"
Global Const STNM_Partial = "�κ�"
Global Const STNM_MidRst = "�߰�"
Global Const STNM_FinRst = "����"
Global Const STNM_Verify = "���"
Global Const STNM_Modify = "����"
Global Const STNM_Reading = "�ǵ�"

' �̻��� Worksheet �ۼ� ��� Flag
Global Const MWS_Ready = "1"        'Worksheet �ۼ�
Global Const MWS_Holding = "2"      'Worksheet build ����
Global Const MWS_Growth = "3"       'Growth ���� - ����
Global Const MWS_Final = "4"        '������� �Է� �Ϸ� - ����

' �̻��� Worksheet ��� ���� Flag
Global Const MRT_GSen = "S"
Global Const MRT_MSen = "C"
Global Const MRT_Stain = "G"
Global Const MRT_AFC = "M"
Global Const MRT_AFS = "N"
Global Const MRT_Both = "B"

Global Const MNM_GSen = "�Ϲ� ������"        ' ������ �Է�ȭ��� ǥ��
Global Const MNM_MSen = "MIC ������"
Global Const MNM_AFC = "AFB/Fungus Culture"
Global Const MNM_AFS = "AFB/Fungus Stain"

Global Const MCD_GSen = "GS"                 ' �˻翡 ���� ������ �׻��� �з� ������
Global Const MCD_MSen = "MS"

' �̻��� ������ ��� ���� ���� (SenFg in lab404)
Global Const MRT_SenRst = "Y"
Global Const MRT_SenRstCd = "RISPN-"

' ��Ÿ�˻� ���ΰ�� ���� ����
Global Const ERT_ValRst = "Y"  '(ValFg in lab351)
Global Const ERT_TxtRst = "Y"  '(TxtFg in lab351)

' ��Ÿ�˻� Worksheet Flag
Global Const EWS_OK = "1"
Global Const EWS_NO = "0"

'Other Error Codes
Global Const CONNECT_SUCCESS = 0
Global Const CONNECT_ERROR = 1

'Open Recordset �� parameter
Global Const adOpenForwardOnly = 0
Global Const adOpenKeyset = 1
Global Const adOpenDynamic = 2
Global Const adOpenStatic = 3


