VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISConst"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-----------------------------------------------------------------------------'
'   ���ϸ� : clsLISConst
'   �ۼ��� :
'   ��  �� : LIS ���뺯��
'   �ۼ��� :
'   ��  �� :
'       1. 5.0.5: �̻��(2005-01-03)
'          - �����к� Workarea(CBC_WorkArea) �����߰�
'          - WBC �˻��ڵ弳��(LC3_WBCCode) �����߰�
'          - NRBC �˻��ڵ弳��(LC3_NRBCCode) �����߰�
'-----------------------------------------------------------------------------'

Option Explicit

'������
'--------
Public LIS_ORDDIV           As String                         '�ӻ󺴸� ó�汸��
Public APS_ORDDIV           As String                         '���ܺ��� ó�汸��
Public BBS_ORDDIV           As String                         '�������� ó�汸��

Public POC_ORDDIV           As String                         'POC ����
Public CMT_ORDDIV           As String                         'CMT ����

'Public HospitalNm           As String
Public CentralLab           As String
Public CentralLabNm         As String
Public EmergencyNo          As Long
Public EmergencyLab         As String
Public EmergencyLabNm       As String
Public AneLab               As String
Public AneLabNm             As String
Public WomLab               As String
Public WomLabNm             As String
Public HrtLab               As String
Public HrtLabNm             As String
Public CS_BarFormat         As String
Public CS_AllCaption        As String
Public CS_DefaultBuilding   As String
Public APS_BUILDCD          As String
Public APS_BUILDNM          As String

Public LABName              As String
Public APSName              As String
Public BBSName              As String

Public APSBarNm             As String
Public BBSBarNm             As String

'�����ڵ� ������1(T_LAB031)
'---------------------------
Public LC2_Index        As String
Public LC2_Panel        As String           ' Paneló�� Item
Public LC2_MultiSpc     As String           ' ������ü
Public LC2_Detail       As String           ' Detail Items
Public LC2_AttrItem     As String           ' �Ӽ� ���� Item
Public LC2_SpcMedia     As String           ' ��ü�� - ����
Public LC2_MediaBio     As String           ' ���� - Bio Chemical Item
Public LC2_MicroAnti    As String           ' ���� - �׻���
Public LC2_Machine      As String           ' ��� - Item
Public LC2_ItemResult   As String           ' Item - ����ڵ�
Public LC2_WAResult     As String           ' WorkArea - ����ڵ�
Public LC2_QcControl    As String           ' QC Control
Public LC2_MBatchRst    As String           ' �̻��� �p� - ��ġ ��� �ڵ�
Public LC2_RelTest      As String           ' ���ð˻��ڵ�
Public LC2_ColListTm    As String           ' �ǹ��� ä������Ʈ ��½ð�
Public LC2_CumItem      As String           ' ���������ȸ Item
Public LC2_MicQcRst     As String           ' �̻���QC ����ڵ�
Public LC2_MicTemp      As String           ' �̻���QC Template
Public LC2_SpeTemp      As String           ' ��Ÿ�˻� Template
Public LC2_SpeAddTemp   As String           ' ��Ÿ�˻� ���� Template
Public LC2_Calculation  As String           ' ��Ÿ�˻� ���� Template
Public LC2_KeyMap       As String

Public LC2_OGYITEM      As String           ' ����ΰ� �˻��׸��ڵ�
Public LC2_TempletTest  As String           ' Ư���˻� ���ø� �˻��׸� ����.
Public LC2_TempletText1 As String           ' Ư���˻� ���ø� text1
Public LC2_TempletText2 As String           ' Ư���˻� ���ø� text1
Public LC2_TempletText3 As String           ' Ư���˻� ���ø� text1
Public LC2_SelReview    As String           ' �׸� Review
Public LC2_TelePhone    As String           ' TelePhone����

'** �߰� ������ Ȯ�� �˻��׸� ����� ����... By M.G.Choi 2006.09.01
Public LC2_DoctTest     As String           ' ������ Ȯ�� �˻��׸� ���
' �������� �߰� By M.G.Choi 2007.07.02 =================================
Public LC2_Warning        As String      'Warning Message
'=======================================================================



'�����ڵ� ������2(T_LAB032)
'---------------------------
Public LC3_Index            As String
Public LC3_DrGrade          As String           ' �ǻ�Grade
Public LC3_BedGrade         As String           ' ������
Public LC3_BedStatus        As String           ' �������
Public LC3_DeptDiv          As String           ' ���з�
Public LC3_HighItem         As String           ' �ٺ�ó��
Public LC3_ByPass           As String           ' ByPass�׸��ڵ�
Public LC3_PocItem          As String           ' Point of Care

Public LC3_RoundTime        As String           ' Roundä�� �ð���
Public LC3_ColTeam          As String           ' ä����
Public LC3_OutLab           As String           ' �ܺ��Ƿ�ó
Public LC3_RefLab           As String           ' Referral Lab
Public LC3_Vander           As String           ' Vander �ڵ�
Public LC3_WorkArea         As String           ' Work Area
Public LC3_Section          As String           ' Section
Public LC3_Specimen         As String           ' ��ü
Public LC3_VerifyFg         As String           ' Auto Verify On/Off
Public LC3_SGroup           As String           ' ��ü��
Public LC3_Media            As String           ' ����
Public LC3_Microbe          As String           ' ��
Public LC3_Species          As String           ' ����
Public LC3_AntiBiotic       As String           ' �׻���
Public LC3_BioChemical      As String           ' ��ȭ���� �����˻�
Public LC3_Volume           As String           ' �����ڵ�
Public LC3_Infect           As String           ' ����������
Public LC3_QCOrderTime      As String           ' QC�ڵ�ó�� �ð���
Public LC3_BedDiv           As String           ' �����з�
Public LC3_NoGrowth         As String           ' �̻��� Nogrowth Code
Public LC3_WorkSheetName    As String           ' ��ũ��Ʈ �̸�
Public LC3_StoreCd          As String           ' ��������
Public LC3_Buildings        As String           ' �ǹ��ڵ�
Public LC3_MWSKinds         As String           ' �̻��� �p�� ����
Public LC3_FileServer       As String           ' File Server Location
Public LC3_StaticItem       As String           ' ���� ��� �׸�
Public LC3_StaticGroup      As String           ' ���� ��� Workarea
Public LC3_PrinterId        As String           ' Printer ID
Public LC3_StartDate        As String           ' ���� �˻��Ⱓ ����
Public LC3_PtDiv            As String           ' ȯ�ڱ���
Public LC3_PartCd           As String           ' �а��ڵ�
Public LC3_HospCd           As String           ' �Ƿ�������
Public LC3_Method           As String           ' �ܺ��������� ����ڵ�
Public LC3_BldQcRst         As String           ' ��������QC ����ڵ�
Public LC3_ICUTestCd        As String           ' ICU �˻��׸�
Public LC3_POCTestCd        As String           ' POC �˻��׸�
Public LC3_POCTimes         As String           ' POC �Է½ð���
Public LC3_ElectronicSign   As String           ' �ӻ󺴸� ���ڼ�������
Public LC3_ReportTesctCd    As String           ' ȸ���뷹��Ʈ �˻��ڵ�
Public LC3_BatchColDept     As String           ' �ϰ�ä�� ���� �μ��ڵ�
Public LC3_MachHistory      As String           ' �����°�������
Public LC3_DiffCount        As String           ' bone marrow diff count
Public LC3_WBCDiffCode      As String           ' WBC Diff �˻��ڵ�
Public LC3_DiffKeyMap       As String           ' Diff Count Key Mapping
Public LC3_TUBERCLE         As String           ' ���ٱհ˻� �����ڵ�
Public LC3_ItemWorkLoad     As String           ' Item�� WorkLoad
Public LC3_EtcWorkLoad      As String           ' Etc Workload
Public LC3_HosFrmUsing      As String           ' ������ ȭ�� �ε�
Public LC3_WorkTime         As String
Public LC3_WBCCode          As String           ' WBC �˻��ڵ� ����
Public LC3_NRBCCode         As String           ' NRBC �˻��ڵ� ����

' �������� �߰� By M.G.Choi 2004.11.01 =================================
Public LC3_INFECTION        As String           ' �������� �Ƿڰ�ü
Public LC3_INFECTIONTEST    As String           ' �������� ��������

Public LC3_RESULTREADTEST   As String           ' �������� ��������
'=======================================================================

'���ø� ������(T_COM004)
'---------------------------
Public LC4_Index            As String
Public LC4_Morphology       As String           ' �� ����
Public LC4_UncolReason      As String           ' ��ä�� ����
Public LC4_Remark           As String           ' Remark
Public LC4_FootNote         As String           ' Foot Note
Public LC4_WarnInfect       As String           ' Warning/Infection
Public LC4_TextResult       As String           ' Text ���
Public LC4_SPTextResult     As String           ' ��Ÿ�˻� Text ���
Public LC4_DCReason         As String           ' ó����� ����
Public LC4_CancelReason     As String           ' ������� ����
Public LC4_ModifyReason     As String           ' ������� ����
Public LC4_QCRejReason      As String           ' QC Reject ����
Public LC4_TempReason       As String           ' �µ��� Reject ����
Public LC4_ClinicalNotice   As String           ' Clinical Notice
Public LC4_Peripheral       As String           ' Peripheral
Public LC4_BMComment        As String           ' bone marrow comment
Public LC4_Hematologic      As String           ' hematologic

Public LC4_TempletCommet    As String           ' ���ø� commet
Public LC4_TempletDignosis  As String           ' ���ø� ����
Public LC4_Calibration      As String           ' QC Calibration �Ұ�
Public LC4_TestItemComment  As String           ' �˻��׸� Comment���
Public LC4_AccessComment    As String           ' �ܷ������� Comment ���ø�

' �������� �߰� By M.G.Choi 2004.11.01 =================================
Public LC4_Infection        As String    '�������� �˻���
'=======================================================================

'������Ʈ �Ķ���� ������(T_COM005)
'---------------------------
Public LC5_Index As String
Public LC5_ScrLock As String        ' Screen Lock Interval
Public LC5_PrgOnOff As String       ' Program On/Off
Public LC5_FnctOnOff As String      ' Fuction On/Off
Public LC5_InfectCond As String     ' �������� ����
Public LC5_BarFormat As String      ' Barcode Label Format
Public LC5_BarTime As String        ' ���Ӱ˻� Barcode Label ��½���
Public LC5_WSPrtTime As String      ' ��Ÿ�˻� Worksheet ��½���
Public LC5_Hospital As String       ' �����̸�, �ּ�, �˻���̸�
Public LC5_CumulTime As String      ' ������� ��½���
Public LC5_LabelTime As String      ' ���� Label ��½���
Public LC5_TempUnit As String       ' ����� �µ� ����
Public LC5_DateFormat As String     ' ��¥ Format
Public LC5_TimeFormat As String     ' �ð� Format

'��ȣ�ο� ������(T_COM099)
'-------------------------
Public COM99_LIS_Specimen As String     '��ü��ȣ
Public COM99_LIS_LabNo As String        '������ȣ
Public COM99_LIS_SeqNo As String       '��������(Workarea ��)
Public COM99_LIS_WorkNo As String       '�Ϲ�Worksheet Unit
Public COM99_LIS_WSUnit As String       '�̻���Worksheet Unit
Public COM99_LIS_Report As String       '�ӻ󺴸� ����� ������� ����

Public LIS_BarDiv  As String            '���ڵ屸��(1:�ӻ�,2:�غ�,3:����)
Public LIS_BarFormat As String          '���ڵ�����

Public LIS_LabDayFormat As String             '������ȣ�ο����� - �ϴ���
Public LIS_LabMonthFormat As String           '������ȣ�ο����� - ������
Public LIS_LabYearFormat As String            '������ȣ�ο����� - �����


'������ Ŭ������ ���
'-------------------------
Public NoIndex_WSBuild  As Long          'WorkSheet Sequence
'
Public Splt_Delimeter   As String
'
'Constants for Result- Data Entry Type
Public RESULT_BY_DEFAULT   As Long
Public RESULT_BY_ACCESSION As Long
Public RESULT_BY_WORKSHEET As Long
Public RESULT_BY_EQUIPMENT As Long
Public RESULT_BY_ITEM      As Long
Public RESULT_BY_MODIFY    As Long
Public RESULT_BY_DIFFCOUNT As Long

'ó��Ӽ��ڵ�
'---------------------
Public LIS_ATTR_SPECIAL_INTEREST    As String           'Special Interest for Ex.
Public LIS_ATTR_CLINICAL_SUMMARY    As String           'ȯ�ں���
Public LIS_ATTR_OPERATION_FINDINGS  As String           '�����Ұ�
Public LIS_ATTR_CLINICAL_DX_CD      As String           '�ӻ������ڵ�(�ӻ��)
Public LIS_ATTR_OPERATION_CD        As String           '�����ڵ�
Public LIS_ATTR_OPERATION_DX_CD     As String           '�������ڵ�
Public LIS_ATTR_LMP                 As String           '����������
Public LIS_ATTR_MESTRUATION_CYCLE   As String           '�����ֱ�
Public LIS_ATTR_MENAPAUSE           As String           '���⳪��

'��Ÿ ���뼱��
'-------------
Public HLDIV_HIGH_CD        As String
Public HLDIV_HIGH_FG        As String
Public HLDIV_LOW_CD         As String
Public HLDIV_LOW_FG         As String

Public LIS_RST_REQUIRED     As String
Public LIS_RST_ALTERNATIVE  As String

Public LIS_Nothing          As String

'ó�� STATUS
'-----------
Public STS_LIS_Order        As String   'ó��
Public STS_LIS_HaveSpc      As String   'ä��
Public STS_LIS_Access       As String   '����
Public STS_LIS_Worksheet    As String   'In-Process
Public STS_LIS_Reading      As String   '�ǵ�
Public STS_LIS_MidRst       As String   'Partial Verify / �߰����
Public STS_LIS_FinRst       As String   '����Ȯ��
Public STS_LIS_Partial      As String   '�κ�
Public STS_LIS_Verify       As String   '���
Public STS_LIS_Modify       As String   '����
Public STS_LIS_Cancel       As String   'ä��/���

'Panel Flag
'-------------------------
Public PN_Group     As String           'Group Item
Public PN_Detail    As String           'Detail Item
Public PN_Normal    As String           '�Ϲ� Item

'�ܺ��Ƿ�(SEND OUT) STATUS
'-------------------------
Public STS_OUTLAB_ACCESSION As String   '����
Public STS_OUTLAB_TRANSFER  As String   '����
Public STS_OUTLAB_WITHDRAW  As String   'ȸ��
Public STS_OUTLAB_RESULT    As String   '���


'��Ź�˻�(REFERRAL) STATUS
'-------------------------
Public STS_REFERRAL_ACCESSION               '����
Public STS_REFERRAL_INPROCESS               '�˻�
Public STS_REFERRAL_RESULT                  'ȸ��
Public STS_REFERRAL_RESULTTRANS             '���


Public MIC_WorkArea As String               '�̻��� Work Area
Public CBC_WorkArea As String               '�����к� Work Area

' �̻��� Worksheet �ۼ� ��� Flag
Public MWS_Ready        As String           'Worksheet �ۼ�
Public MWS_Holding      As String           'Worksheet build ����
Public MWS_Growth       As String           'Growth ���� - ����
Public MWS_Final        As String           '������� �Է� �Ϸ� - ����
Public MWS_Selected     As String
Public MWS_Deselected   As String
Public MWS_ForCulture   As Long             '�̻��� ���˻翡 �ش��ϴ� Worksheet�ڵ�
Public MWS_ForSensi     As Long             '�̻��� ���˻翡 �ش��ϴ� Worksheet�ڵ�
Public MWS_ForStain     As Long             '�̻��� Stain�˻翡 �ش��ϴ� Worksheet�ڵ�
Public MWS_ForAll       As Long             '�̻��� ���˻翡 �ش��ϴ� Worksheet�ڵ�

' �̻��� Worksheet ��� ���� Flag
Public MRT_GenSen   As String               'General Sensitivity
Public MRT_MicSen   As String               'MIC Sensitivity
Public MRT_Stain    As String               'Gram Stain
Public MRT_AFC      As String               'AFB Culture
Public MRT_AFS      As String               'AFB Stain
Public MRT_AFN      As String               'AFB Sensitivity
Public MRT_FUC      As String               'Fungus Culture
Public MRT_FUS      As String               'Fungus Stain
Public MRT_Both     As String               '?
Public MRT_OTHER    As String               '?

Public MNM_GSen As String                   ' ������ �Է�ȭ��� ǥ��
Public MNM_MSen As String
Public MNM_AFC  As String
Public MNM_AFS  As String

Public MCD_GSen As String                   ' �˻翡 ���� ������ �׻��� �з� ������
Public MCD_MSen As String

' �̻��� ������ ��� ���� ���� (SenFg in lab404)
Public MRT_NoSenRst As String
Public MRT_SenRst   As String
Public MRT_SenRstCd As String

' ��Ÿ�˻� ���ΰ�� ���� ����
Public ERT_ValRst   As String               '(ValFg in lab351)
Public ERT_TxtRst   As String               '(TxtFg in lab351)

' ��Ÿ�˻� Worksheet Flag
Public EWS_OK       As String
Public EWS_NO       As String

' ��Ÿ�˻� ����� ��� Work Area
Public ERT_REPORT   As String

' QC Level �ڵ�
Public QC_LEVEL_HIGH    As String
Public QC_LEVEL_NORMAL  As String
Public QC_LEVEL_LOW     As String

'������ WorkArea
Public RI_WORKAREA      As String

' Interface ���� ����Ÿ ���
Public CS_EqpError      As String

'DECLARE ENUM
'-------------
Public Enum enOrdFG                           'ó������
    Order_OCS = 0
    Order_BackUp = 1
    Order_ItemAdd = 2
    Order_Reverse = 3
    Order_Change = 4
    Order_Repeat = 5
    Order_DC = 6
    Order_QC = 7
    Order_SpecialStudy = 8
End Enum

Public Enum enBussDiv                       '��������
    BussDiv_OutPatient = "1"
    BussDiv_InPatient = "2"
    BussDiv_Emergency = "3"
    BussDiv_HealthCaer = "4"
    BussDiv_Referral = "5"
    BussDiv_ICU = "9"
End Enum

Public Enum enSpOrdDiv                      'Ư��ó�汸��
    SpOrdDiv_None = 0
    SpOrdDiv_VIP = 1
    SpOrdDiv_Research = 2
    SpOrdDiv_Other = 3
End Enum

Public Enum enTrustFg                       '��Ź����
    TrustFg_None = 0
    TrustFg_Referral = 1
End Enum

Public Enum enStsCd                       'ó�� STATUS
    StsCd_LIS_Order = 0                   'ó��
    StsCd_LIS_Collection = 1              'ä��
    StsCd_LIS_Accession = 2               '����
    StsCd_LIS_InProcess = 3               '�˻���
    StsCd_LIS_MidRst = 4                  '�߰�����
    StsCd_LIS_FinRst = 5                  '���Ȯ��(�������)
    StsCd_LIS_Modify = 6                  '�������
    StsCd_LIS_Cancel = 7                  '������
End Enum

'*
Public Enum enTestDiv                       '�˻籸��
    TST_RouTest = "0"      ' �Ϲ� �˻�
    TST_SpeTest = "1"      ' Ư�� �˻�
    TST_MicTest = "2"      ' �̻��� �˻�
    TST_AboTest = "3"      ' ������ �˻�
End Enum

Public Enum enLabDiv                       '������ȣ�ο�����
    LabDiv_ByDay = 1        ' �ϴ���
    LabDiv_ByMonth = 2      ' ������
    LabDiv_ByYear = 3       ' �����
    LabDiv_BySpc = 4        ' ��ü������
End Enum

Public Enum enFlag
    Flag_Off = "0"                        '��
    Flag_On = "1"                         '��
End Enum


Public Enum enDoneFg                         '��������
    DoneFg_Order = 0
    DoneFg_Collection = 1
    DoneFg_Accession = 2
End Enum

'
'Public Enum FootnoteFg
'    FootNote_Exist = "Y"
'    FootNote_None = ""
'End Enum

Public Enum enLabNoY2kFlag
    LabNoY2kFlag_20C = "19"
    LabNoY2kFlag_21C = "20"
End Enum

Public Type Field
    Field As String
End Type

Public Type ORDB
    PtId As String
    OrdDt As String
    OrdNo As Long
    OrdSeq As Long
    OrdCd As String
    SpcCd As String
    RcvDt As String
    RCVTM As String
    StsCd As enStsCd
    StatFg As String
    DoneFg As enDoneFg
    WorkArea As String
    AccDt As String
    AccSeq As Long
    ReceptNo As Long
    Mesg As String
End Type

Public Type ACCESSION
    WorkArea As String
    AccDt As String
    AccSeq As Integer
    PtId As String
    OrdDt As String
    Sex As String
    Age As Integer
    BedInDt As String
    ReqDt As String
    ReqTm As String
    DeptCd As String
    OrdDoct As String
    MajDoct As String
    EntId As String
    OrdDiv As String
    OrdFG As enOrdFG
    SpOrdDiv As enSpOrdDiv
    DoneFg As enDoneFg
    WardID As String
    HosilId As String
    BedID As String
    RefHosCd As String
    TrustFg As enTrustFg
    TrustDt As String
    OrdNo As Long
    OrdTm As String
    EntDt As String
    EntTm As String
    OrgAccNo As String
    HosCd As String
    Seq As Integer
    RcvDt As String
    RCVTM As String
    SPCYY As String
    SPCNO As String
    StsCd As enStsCd
    StatFg As String
    FZFG As String
End Type

Public Enum enSaveCollection
    Save_Ward_Collection = 0                        '����ä��
    Save_Nurse_Collection = 1                       '������ȣ��ä��
    save_OutPatient_Collection = 2                  '�ܷ���ȣ��ä��
End Enum

'= ó�� Body ='
Public Type OrderBody
   OrdSeq As Integer    ' ó�� Seq
   OrdCd As String      ' ó���ڵ�
   SpcCd As String      ' ��ü�ڵ�
   StoreCd As String    ' ��������
   DcFg As String       ' DC ����
   DcDt As String       ' DC ����
   DcNo As Integer      ' DC ó���ȣ
   AttrCd As String     ' �Ӽ�����
   ExamDt As String     ' ���� �ǽ�����
   ExamTm As String     ' ���� �ǽýð�
   ExamDoct As String   ' ���� �ǽ��ǻ�
   StsCd As String      ' Status
   StatFg As String     ' ���޿���
   InsDiv As String     ' �޿�����
   DoneFg As String
End Type

'= ó�� Header ='
Public Type OrderHeader
   OrdNo As Integer     ' ó���ȣ
   ReqDt As String      ' �ǽÿ�û����
   ReqTm  As String     ' �ǽÿ�û�ð�
   OrdBody() As OrderBody
End Type


Public Type ResultTable
    TestCd As String     '/* �˻��׸� �ڵ�   */
    RstVal As String     '/* ���(Numeric)   */
    RstCd As String      '/* ����ڵ�(Alpha) */
    RstUnit As String    '/* Unit            */
    HLDiv As String      '/* High/Low(H:Hig,L:Low) */
    DPDiv As String      '/* Delta/Panic(D:Delta,P:Panic) */
    SpcCd As String      '/* ��ü�ڵ�        */
    StatFg As String     '/* ���޿���('0':��,'1':��) */
    LastRst As String    '/* �ֱٰ��        */
    LastVfyDt As String  '/* �ֱٰ��Ȯ����  */
    LastVfyTm As String  '/* �ֱٰ��Ȯ�νð� */
    LastVfyId As Integer '/* �ֱٰ��Ȯ����  */
    VfyDt As String      '/* ���Ȯ������    */
    VfyTm As String      '/* ���Ȯ�νð�    */
    VfyId As Integer     '/* ���Ȯ����      */
    AttrCd As String     '/* �Ӽ��ڵ�   */
    MfyFg As String      '/* ��������('0':��,'1':��) */
    GrpFg As String      '/* �׷��� �������('0':��,'1':��) */
    ValFg As String      '/* ��ġ �������('0':��,'1':��)  - ��Ÿ�˻翡���� ���*/
    TxtFg As String      '/* TEXT �������('0':��,'1':��) */
    RstType As String    '/* �������(N,Alpha,����,Free) */
    RstDiv As String     '/Alternative,Required ���� */
    DetailFg As String   '/* ���׸�Group Seq */
    PtId As String       '/* ȯ�� ID */
    OrdDt As String      '/* ó���� */
    OrdNo As Integer     '/* ó���ȣ */
    OrdSeq As Integer    '/* ó��Seq */
    SenFg As String      '/* ��������� ���� */
    AutoFg As String
    EqpCd As String
    
    TestDiv As String
End Type

Public Type OutLabTable
    WorkArea As String  '/* Work Area(������ȣ��) */
    AccDt As String     '/* ������(������ȣ��)    */
    AccSeq As Integer   '/* ��������(������ȣ��)  */
    TestCd As String    '/* �˻��׸�        */
    PtId As String      '/* ȯ��ID  */
    SpcCd As String     '/* ��ü�ڵ�  */
    OutLabCd As String  '/* ����ڵ�        */
    StsCd As String     '/* STATUS          */
    RcvDt As String     '/* ��������        */
    SendDt As String    '/* SEND Date       */
    ChargeDt As String  '/* û������        */
    SendId As String    '/* ������ ID   */
End Type

Public Type tpMicWorkSheet
    SpcGrp As String
    SpcGrpNm As String
    WsCode As String
    WsName As String
    WsType As String
    WsRstType As String
End Type

Public Type tTmpValue
    sTpcd As String
    sTpnm As String
    swsfg As String           ' 1 : check,   0: uncheck
    sTpdata As String
    sValueCnt As Long
    sRstFields As String
    sSaveDecision As Integer  ' 1 : save ,   0 : Unsave
End Type

Private Sub Class_Initialize()

    '�����ڵ� ������1(T_COM002)
    LC2_Index = "C100"
    LC2_Panel = "C101"         ' Paneló�� Item
    LC2_MultiSpc = "C102"      ' ������ü
    LC2_Detail = "C103"        ' Detail Items
    LC2_AttrItem = "C105"      ' �Ӽ� ���� Item
    LC2_SpcMedia = "C106"      ' ��ü�� - ����
    LC2_MediaBio = "C107"      ' ���� - Bio Chemical Item
    LC2_MicroAnti = "C108"     ' ���� - �׻���
    LC2_Machine = "C109"       ' ��� - Item
    LC2_ItemResult = "C110"    ' Item - ����ڵ�
    LC2_WAResult = "C111"      ' WorkArea - ����ڵ�
    LC2_QcControl = "C112"     ' QC Control
    LC2_MBatchRst = "C113"     ' �̻��� �p� - ��ġ ��� �ڵ�
    LC2_RelTest = "C114"       ' ���ð˻��ڵ�
    LC2_ColListTm = "C115"     ' �ǹ��� ä������Ʈ ��½ð�
    LC2_CumItem = "C116"       ' ���������ȸ Item
    LC2_MicQcRst = "C117"      ' �̻���QC ����ڵ�
    LC2_MicTemp = "C118"       ' �̻���QC Template
    LC2_SpeTemp = "C119"       ' ��Ÿ�˻� ��� Template
    LC2_SpeAddTemp = "C120"    ' ��Ÿ�˻� ��� ���� Template
    LC2_Calculation = "C121"   ' ������ ����
    LC2_OGYITEM = "C122"       ' ����ΰ� �˻��׸� �ڵ�
    LC2_TempletTest = "C124"   ' ���ø� �˻��׸� ����.
    LC2_TempletText1 = "C125"
    LC2_TempletText2 = "C126"
    LC2_TempletText3 = "C127"
    LC2_SelReview = "C128"      ' SelReview �˻��ڵ�/��ü��
    LC2_TelePhone = "C129"      ' ��ȭ��ȣ����
    
    '** �߰� ������ Ȯ�� �˻��׸� By M.G.Choi 2006.09.01
    LC2_DoctTest = "C150"
    ' �������� �߰� By M.G.Choi 2007.07.02 =================================
    LC2_Warning = "C151"        ' Warning Message
    '=======================================================================
    '---------------------------------------------------
    
    '�����ڵ� ������2(T_COM003)
    LC3_Index = "C200"
    LC3_DrGrade = "C201"       ' �ǻ�Grade
    LC3_BedGrade = "C202"      ' ������
    LC3_BedStatus = "C203"     ' �������
    LC3_DeptDiv = "C204"       ' ���з�
    LC3_HighItem = "C205"      ' �ٺ�ó��
    LC3_PocItem = "C206"       ' Point of Care
    LC3_ByPass = "C207"        ' Bypass
    LC3_RoundTime = "C208"     ' Roundä�� �ð���
    LC3_ColTeam = "C209"       ' ä����
    LC3_OutLab = "C210"        ' �ܺ��Ƿ�ó
    LC3_RefLab = "C211"        ' Referral Lab
    LC3_Vander = "C212"        ' Vander �ڵ�
    LC3_WorkArea = "C213"      ' Work Area
    LC3_Section = "C214"       ' Section
    LC3_Specimen = "C215"      ' ��ü
    LC3_VerifyFg = "C216"      ' Auto Verify On/Off
    LC3_SGroup = "C217"        ' ��ü��
    LC3_Media = "C218"         ' ����
    LC3_Microbe = "C219"       ' ��
    LC3_Species = "C220"       ' ����
    LC3_AntiBiotic = "C221"    ' �׻���
    LC3_BioChemical = "C222"   ' ��ȭ���� �����˻�
    LC3_Volume = "C223"        ' �����ڵ�
    LC3_Infect = "C224"        ' ����������
    LC3_QCOrderTime = "C225"   ' QC�ڵ�ó�� �ð���
    LC3_BedDiv = "C226"        ' �����з�
    LC3_NoGrowth = "C227"      ' �̻��� Nogrowth Code
    LC3_WorkSheetName = "C228" ' ��ũ��Ʈ �̸�
    LC3_StoreCd = "C229"       ' ��������
    LC3_Buildings = "C230"     ' �ǹ��ڵ�
    LC3_MWSKinds = "C231"      ' �̻��� �p�� ����
    LC3_FileServer = "C232"    ' File Server Location
    LC3_StaticItem = "C233"    ' ���� ��� �׸�
    LC3_StaticGroup = "C234"   ' ���� ��� Workarea
    LC3_PrinterId = "C235"     ' Printer ID
    LC3_StartDate = "C236"     ' ���� �˻��Ⱓ ����
    LC3_PtDiv = "C237"         ' ȯ�ڱ���
    LC3_PartCd = "C238"        ' �а��ڵ�
    LC3_HospCd = "C239"        ' �Ƿ�������
    LC3_Method = "C240"        ' �ܺ��������� ����ڵ�
    LC3_BldQcRst = "C241"      ' ��������QC ����ڵ�
    LC3_ICUTestCd = "C242"     ' ICU �׸��ڵ�
    LC3_POCTestCd = "C243"     ' POC �׸��ڵ�
    LC3_POCTimes = "C244"      ' POC �Է½ð���
    LC3_ElectronicSign = "C245" '���ڼ���
    LC3_ReportTesctCd = "C246" ' ȸ���뷹��Ʈ �˻��ڵ�
    LC3_WBCDiffCode = "C247"    'WBC Diff �˻��ڵ�
    LC3_BatchColDept = "C248"
    LC3_DiffCount = "C251"      ' bone marrow diff count
    LC3_MachHistory = "C252"    ' �����°�������
    LC3_TUBERCLE = "C253"       ' ���ٱհ˻� �����ڵ�
    LC3_ItemWorkLoad = "C254"   ' Item WorkLoad
    LC3_EtcWorkLoad = "C255"    ' Etc WorkLoad
    LC3_DiffKeyMap = "C256"     ' Diff Key Mapping
    LC3_HosFrmUsing = "C261"    ' �ӻ󺴸� ������ �ε�ȭ�� ���� ���������� "C262" 'C257�� ���� ��
    LC3_WorkTime = "C258"       ' �����ð��� �ۼ�
    LC3_WBCCode = "C270"        ' WBC �˻��ڵ� ����
    LC3_NRBCCode = "C271"       ' NRBC �˻��ڵ� ����
    
    ' �������� �߰� By M.G.Choi 2004.11.01 =================================
    LC3_INFECTION = "C259"      ' �������� �Ƿڰ�ü
    LC3_INFECTIONTEST = "C260"  ' �������� ��������
    
    LC3_RESULTREADTEST = "C265" ' �����ǵ����� �˻��׸�
    '=======================================================================
    
    

    '���ø� ������(T_COM004)
    LC4_Index = "C400"
    LC4_Morphology = "C401"         ' �� ����
    LC4_UncolReason = "C402"        ' ��ä�� ����
    LC4_Remark = "C403"             ' Remark
    LC4_FootNote = "C404"           ' Foot Note
    LC4_WarnInfect = "C405"         ' Warning/Infection
    LC4_TextResult = "C406"         ' Text ���
    LC4_SPTextResult = "C407"       ' ��Ÿ�˻� Text ���
    LC4_DCReason = "C408"           ' ó����� ����
    LC4_CancelReason = "C409"       ' ������� ����
    LC4_ModifyReason = "C410"       ' ������� ����
    LC4_QCRejReason = "C411"        ' QC Reject ����
    LC4_TempReason = "C412"         ' �µ��� Reject ����
    LC4_ClinicalNotice = "C413"     ' Clinical Notice
    LC4_Peripheral = "C420"         ' Peripheral
    LC4_BMComment = "C421"          ' bone marrow comment
    LC4_Hematologic = "C422"        ' hematologic
    LC4_TempletCommet = "C423"      ' ���ø� commet
    LC4_TempletDignosis = "C424"    ' ���ø� ����
    LC4_Calibration = "C425"        ' ��� Calibration
    LC4_TestItemComment = "C426"    ' �˻��׸� Comment
    LC4_AccessComment = "C427"      ' �ܷ������� Comment ���ø�
    
    ' �������� �߰� By M.G.Choi 2004.11.01 =================================
    LC4_Infection = "C428"          '�������� �˻���
    '=======================================================================
    
    '������Ʈ �Ķ���� ������(T_COM005)
    LC5_Index = "C500"
    LC5_ScrLock = "C501"       ' Screen Lock Interval
    LC5_PrgOnOff = "C502"      ' Program On/Off
    LC5_FnctOnOff = "C503"     ' Fuction On/Off
    LC5_InfectCond = "C504"    ' �������� ����
    LC5_BarFormat = "C505"     ' Barcode Label Format
    LC5_BarTime = "C506"       ' ���Ӱ˻� Barcode Label ��½���
    LC5_WSPrtTime = "C507"     ' ��Ÿ�˻� Worksheet ��½���
    LC5_Hospital = "C508"      ' �����̸�, �ּ�, �˻���̸�
    LC5_CumulTime = "C509"     ' ������� ��½���
    LC5_LabelTime = "C510"     ' ���� Label ��½���
    LC5_TempUnit = "C511"      ' ����� �µ� ����
    LC5_DateFormat = "C512"    ' ��¥ Format
    LC5_TimeFormat = "C513"    ' �ð� Format

    '������
    LIS_ORDDIV = "L"
    APS_ORDDIV = "A"
    BBS_ORDDIV = "B"
    POC_ORDDIV = "P"
    CMT_ORDDIV = "C"
        'Constants for Result- Data Entry Type
    RESULT_BY_DEFAULT = 1
    RESULT_BY_ACCESSION = 2
    RESULT_BY_WORKSHEET = 3
    RESULT_BY_EQUIPMENT = 4
    RESULT_BY_ITEM = 5
    RESULT_BY_MODIFY = 6
    RESULT_BY_DIFFCOUNT = 7

    '���ܺ��� ó��Ӽ��ڵ�
    LIS_ATTR_SPECIAL_INTEREST = "A1"        'Special Interest for Ex.
    LIS_ATTR_CLINICAL_SUMMARY = "A2"        'ȯ�ں���
    LIS_ATTR_OPERATION_FINDINGS = "A4"      '�����Ұ�
    LIS_ATTR_CLINICAL_DX_CD = "A3"          '�ӻ������ڵ�(�ӻ��)
    LIS_ATTR_OPERATION_CD = "A5"            '�����ڵ�
    LIS_ATTR_OPERATION_DX_CD = "A6"         '�������ڵ�
    LIS_ATTR_LMP = "A7"                     '����������
    LIS_ATTR_MESTRUATION_CYCLE = "A8"       '�����ֱ�
    LIS_ATTR_MENAPAUSE = "A9"               '���⳪��


    '��Ÿ���뼱��
    STS_LIS_Order = "ó��"
    STS_LIS_HaveSpc = "ä��"
    STS_LIS_Access = "����"
    STS_LIS_Worksheet = "�˻���"
    STS_LIS_Reading = "�ǵ�"
    STS_LIS_MidRst = "�߰�"
    STS_LIS_FinRst = "����"
    STS_LIS_Verify = "���"
    STS_LIS_Partial = "�κ�"
    STS_LIS_Modify = "����"
    STS_LIS_Cancel = "���"
    '
    PN_Group = "G"         'Group Item
    PN_Detail = "D"        'Detail Item
    PN_Normal = ""         '�Ϲ� Item

    '�ܺ��Ƿ�(SEND OUT) STATUS
    STS_OUTLAB_ACCESSION = "1"              '����
    STS_OUTLAB_TRANSFER = "2"               '����
    STS_OUTLAB_WITHDRAW = "3"               'ȸ��
    STS_OUTLAB_RESULT = "4"                 '���

    '��Ź�˻�(REFERRAL) STATUS
    STS_REFERRAL_ACCESSION = "1"            '����
    STS_REFERRAL_INPROCESS = "2"            '�˻�
    STS_REFERRAL_RESULT = "3"               'ȸ��
    STS_REFERRAL_RESULTTRANS = "4"          '���

'    HospitalNm = "��õ�ǰ����� �μ� �溴��"
    CentralLab = "10"           '-- �߾Ӱ˻��
    CentralLabNm = "�߾�"       '-- �߾Ӱ˻��
    WomLab = "20"               '-- ����Ŭ����
    WomLabNm = "����Ŭ����"     '-- ����Ŭ����
    HrtLab = "30"               '-- ���弾��
    HrtLabNm = "���弾��"       '-- ���弾��
    AneLab = "40"               '-- ���̼���
    AneLabNm = "���̼���"       '-- ���̼���
    EmergencyNo = 5             '-- ���޼���
    EmergencyLab = "50"         '-- ���޼���
    EmergencyLabNm = "���޼���"   '-- ���޼�Ÿ

    CS_AllCaption = "(��ü)"
    CS_DefaultBuilding = "10"
    APS_BUILDCD = "10"
    APS_BUILDNM = "����"
    LABName = "�ӻ�"
    APSName = "�غ�"
    APSBarNm = "AP"
    BBSName = "����"
    BBSBarNm = "XM"
    LIS_BarDiv = "1"

    LIS_LabDayFormat = "YYYYMMDD"     '�ϴ���
    LIS_LabMonthFormat = "YYYYMM"     '������
    LIS_LabYearFormat = "YYYY"        '�����

    NoIndex_WSBuild = 10            'WorkSheet Sequence
    Splt_Delimeter = "$"

    HLDIV_HIGH_CD = "H"
    HLDIV_HIGH_FG = "��"
    HLDIV_LOW_CD = "L"
    HLDIV_LOW_FG = "��"

    LIS_RST_REQUIRED = "R"
    LIS_RST_ALTERNATIVE = "A"

    LIS_Nothing = "(����)"

' �̻��� Worksheet �ۼ� ��� Flag
    MWS_Ready = "1"                     'Worksheet �ۼ�
    MWS_Holding = "2"                   'Worksheet build ����
    MWS_Growth = "3"                    'Growth ���� - ����
    MWS_Final = "4"                     '������� �Է� �Ϸ� - ����

    MWS_Selected = "1"
    MWS_Deselected = "0"

    MWS_ForAll = 0
    MWS_ForCulture = 1
    MWS_ForStain = 2
    MWS_ForSensi = 3

    MRT_GenSen = "S"        'General Sensi
    MRT_MicSen = "C"        'MIC Sensi
    MRT_Stain = "G"         'Gram Stain
    MRT_AFC = "M"           'AFB Culture
    MRT_AFS = "B"           'AFB Stain
    MRT_AFN = "A"           'AFB Sensi
    MRT_FUC = "F"           'Fungus Culture
    MRT_FUS = "N"           'Fungus Stain
    MRT_Both = "B"
    MRT_OTHER = "O"

    MNM_GSen = "�Ϲ� ������"        ' ������ �Է�ȭ��� ǥ��
    MNM_MSen = "MIC ������"
    MNM_AFC = "AFB/Fungus Culture"
    MNM_AFS = "AFB/Fungus Stain"

    MCD_GSen = "GS"                 ' �˻翡 ���� ������ �׻��� �з� ������
    MCD_MSen = "MS"

' �̻��� ������ ��� ���� ���� (SenFg in lab404)
    MRT_NoSenRst = ""
    MRT_SenRst = "Y"
    MRT_SenRstCd = "RISPN-"

' ��Ÿ�˻� ���ΰ�� ���� ����
    ERT_ValRst = "Y"  '(ValFg in lab351)
    ERT_TxtRst = "Y"  '(TxtFg in lab351)

' ��Ÿ�˻� Worksheet Flag
    EWS_OK = "1"
    EWS_NO = "0"

    ERT_REPORT = "BM"

' QC Level �ڵ�
    QC_LEVEL_HIGH = "H"
    QC_LEVEL_NORMAL = "N"
    QC_LEVEL_LOW = "L"

' Interface ���� ����Ÿ ���
    CS_EqpError = "Error"
    RI_WORKAREA = "RI"
    
' ��ȣ�ο�����
    COM99_LIS_Specimen = "L001"           '��ü��ȣ
    COM99_LIS_LabNo = "L002"              '������ȣ
    COM99_LIS_WorkNo = "L003"             '�Ϲ�Worksheet Unit
    COM99_LIS_SeqNo = "L004"              '��������(WorkArea��)
    COM99_LIS_WSUnit = "L009"             '�̻���Worksheet Unit
    COM99_LIS_Report = "L010"             '����� ������� ����
    LIS_BarFormat = "0########"           '9�ڸ�(�����ھֺ���)
    CS_BarFormat = "000000000"
    MIC_WorkArea = "04"
    CBC_WorkArea = "01"
End Sub