Attribute VB_Name = "modCmtConstants"
Option Explicit

Global Const HospitalNm = "��õ�ǰ����� �μ� �溴��"

'Global declare the data class
Global Const DatabaseName$ = "Lab"
Global Const Connect$ = "Lab/Lab"
Global Const ConnectString = "dsn=sybaseODBC;uid=hisbase;pwd=hispass;"
'
Global SB_ServerNm As String    '-- ������
Global SB_DatabaseNm As String  '-- ����Ÿ���̽���
Global SB_LoginId As String     '-- �α�Ƶ�
Global SB_Password As String    '-- �н�����

Global SB_ConnStatus As Integer

'��������===============>>>> ���İ��� ��.
Global Const HosptGb = "10"

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

'****************  Sybase DB ���� System Data & Time ���ϴ� �Լ�  ****************'
Global Const CS_SybaseDate = "convert(char(8),getdate(),112)"
Global Const CS_SybaseTime = "substring(convert(char(8),getdate(),108),1,2)+" & _
                             "substring(convert(char(8),getdate(),108),4,2)+" & _
                             "substring(convert(char(8),getdate(),108),7,2)"
'***********************************************************************************'

'Color
Global Const CR_LIGHT_BLUE = &HB9602F
Global Const CR_LIGHT_BLUE1 = &HDF6A3E
Global Const CR_LIGHT_RED = &H7477EF
Global Const CR_LIGHT_YELLOW = &HC0FFFF
Global Const CR_GREY = &HC0C0C0
Global Const CR_BROWN = &H404080


' ���̺�� ��� ( LIS Tables )
Global Const T_HIS001 = "h1ptntinfo"   'ȯ�ڱ⺻������
Global Const T_HIS002 = "h1admin"      'ȯ�ڱ⺻������
Global Const T_HIS003 = "hzdept"       '�μ�������
Global Const T_HIS004 = "hzdept"       '����������
Global Const T_HIS005 = "HIS005"       '���󸶽���
Global Const T_HIS007 = "hzempl"       '�ǻ縶����
Global Const T_HIS008 = "h2diag"       '�󺴸�����
Global Const T_HIS009 = "h1actmat"     '����������

Global Const T_LAB001 = "h7lab001"     '�˻��׸񸶽���
Global Const T_LAB002 = "h7lab002"     '�˻絿�Ǿ����
Global Const T_LAB003 = "h7lab003"     '�˻纰��񸶽���
Global Const T_LAB004 = "h7lab004"     '������ü������
Global Const T_LAB005 = "h7lab005"     '����ġ������
Global Const T_LAB006 = "h7lab006"     '��񸶽���
Global Const T_LAB007 = "h7lab007"     '�ܺΰ˻縶����
Global Const T_LAB008 = "h7lab008"     'Worksheet������
Global Const T_LAB009 = "h7lab009"     '�������׳���
Global Const T_LAB011 = "h7lab011"     'QC������
Global Const T_LAB012 = "h7lab012"     'QC�˻�����������
Global Const T_LAB013 = "h7lab013"     '�̻���QC������
Global Const T_LAB014 = "h7lab014"     'QC��Ʈ�Ѹ�����
Global Const T_LAB015 = "h7lab015"     '����������

Global Const T_LAB021 = "h7lab021"     'QC Control Master
Global Const T_LAB022 = "h7lab022"     'QC Item Master
Global Const T_LAB023 = "h7lab023"     'QC Master
Global Const T_LAB024 = "h7lab024"     'QC Item Information
Global Const T_LAB025 = "h7lab025"     'QC Schedule
Global Const T_LAB026 = "h7lab026"     'QC �������
Global Const T_LAB027 = "h7lab027"     'QC ��������
Global Const T_LAB028 = "h7lab028"     'QC Text ����

Global Const T_LAB031 = "h7lab031"     '�����ڵ帶����1
Global Const T_LAB032 = "h7lab032"     '�����ڵ帶����2
Global Const T_LAB033 = "h7lab033"     '�����ڵ帶����3
Global Const T_LAB034 = "h7lab034"     '�����ڵ帶����3
Global Const T_LAB035 = "h7lab035"     '���ø�������
Global Const T_LAB036 = "h7lab036"     '��Ÿ�˻����ø�������
Global Const T_LAB099 = "h7lab099"     '��ȣ�ο�������

Global Const T_LAB101 = "h7lab101"     'ó��Header
Global Const T_LAB102 = "h7lab102"     'ó��Body
Global Const T_LAB103 = "h7lab103"     'QCó��Body

Global Const T_LAB201 = "h7lab201"     'ä����������
Global Const T_LAB202 = "h7lab202"
Global Const T_LAB203 = "h7lab203"     '���Ӱ˻系��
Global Const T_LAB204 = "h7lab204"     '�ϰ�ä������
Global Const T_LAB205 = "h7lab205"     '�ܺ��Ƿڳ���

Global Const T_LAB301 = "h7lab301"     'Worksheet����
Global Const T_LAB302 = "h7lab302"     '�Ϲݰ������
Global Const T_LAB303 = "h7lab303"     '�Ϲ��ؽ�Ʈ�������
Global Const T_LAB304 = "h7lab304"     'FootNote����
Global Const T_LAB305 = "h7lab305"     'Supplemental����
Global Const T_LAB306 = "h7lab306"     '�ڵ�ȭ��� ���۳���
Global Const T_LAB307 = "h7lab307"     'QC�������
Global Const T_LAB308 = "h7lab308"     '�Ϲݰ����������

Global Const T_LAB350 = "h7lab350"     '��Ÿ�˻缳������
Global Const T_LAB351 = "h7lab351"     '��Ÿ�˻�������
Global Const T_LAB352 = "h7lab352"     '��Ÿ�˻�Numeric���
Global Const T_LAB353 = "h7lab353"     '��Ÿ�˻�Text���
Global Const T_LAB354 = "h7lab354"     '��Ÿ�˻��������

Global Const T_LAB401 = "h7lab401"     '�̻���Worksheet����
Global Const T_LAB402 = "h7lab402"     '�̻���Worksheet����
Global Const T_LAB403 = "h7lab403"     '�̻���Growth Reading����
Global Const T_LAB404 = "h7lab404"     '�̻����������
Global Const T_LAB405 = "h7lab405"     '�̻����������������
Global Const T_LAB406 = "h7lab406"     '�̻���QC�������
Global Const T_LAB407 = "h7lab407"     '�̻�����������

'** ���հ���/�ǵ� ���� ���� ���̺� **'
Global Const T_LAB501 = "h7lab501"     '�Կ�ȯ�ڳ���
Global Const T_LAB502 = "h7lab502"     '��������
Global Const T_LAB503 = "h7lab503"     '�������
Global Const T_LAB504 = "h7lab504"     '��������
Global Const T_LAB505 = "h7lab505"     '��������
Global Const T_LAB506 = "h7lab506"     'Template

Global Const T_LAB999 = "h7lab999"     '�� system �������

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
Global Const CD2_PtDiv = "C237"         ' ȯ�ڱ���


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

'�˻�����
Global Const TST_RouTest = "0"      ' ��κ� �˻�
Global Const TST_SpeTest = "1"      ' Ư�� �˻�
Global Const TST_MicTest = "2"      ' �̻��� �˻�

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

