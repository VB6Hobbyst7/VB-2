VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBBSConst"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'ó�� ����(Status)
Public Enum BBSOrderStatus
    stsORDER = 0        'ó��
    stsCOLLECT          'ä��
    stsACCESS           '����
    stsINPROCESS        '������(Cross-Matching�����Ͻ� )
    stsEnd              '�Ϸ�(pheresisó��)
End Enum

'2001-11-22�߰� : ȣȯ�� ������ ���Ͽ� �ϳ� �� ����.
Public Enum BBSOrdStatus
    stsORDER = 0        'ó��
    stsCOLLECT          'ä��
    stsACCESS           '����
    stsREQUEST          '������û
    stsINPROCESS        '������(Cross-Matching�����Ͻ� )
    stsEnd              '�Ϸ�(pheresisó��)
End Enum

'���� ����(Status)
Public Enum BBSBloodStatus
    stsENTER = 0        '�԰�
    stsRETURN           '��ȯ
    stsASSIGN           'Assign
    stsDELIVERY         '���
    stsEXPIRE           '���
    stsBAG              'ȸ��
End Enum

Public Enum BBSOrderStatusColor
    clrOrder = 0           'black
    clrCOLLECT = 65280     'green
    clrACCESS = 16711680   'blue
    clrINPROCESS = 255     'red
    cIrEND = 12583104
End Enum

Public Enum BBSOrdStatusColor
    clrOrder = 0           'black
    clrCOLLECT = 65280     'green
    clrACCESS = 16711680   'blue
    clrREQUEST = &HFF      'red
    clrINPROCESS = &H4A4189
    cIrEND = &H8000&
End Enum

'���Ÿ��
'��ü�߰���û ó������
Public Enum BBSSpcAddFg
    stsORDER = 0
    stsCOLLECT
End Enum
'������� ����(BBS206)
Public Enum BBSSaveStatue
    stsUsed = 0         '��밡��
    stsNotUsed          '�����
    stsNotUsedPoss      '����Ļ��
End Enum

'��ҿ���
Public Enum BBSCancelStatus
    stsNotCancel = 0    '��Ҿƴ�
    stsCancel           '��һ���
End Enum
'��⿩��
Public Enum BBSExpFg
    stsNotExpFg = 0     '��� �ƴ�
    stsExpfg            '���
End Enum
'�Կ�����
Public Enum BBSBUSSDIV
    stsNotBed = 1
    stsBed
    stsEr
End Enum

Public Enum DonorStatus
    '�����ڻ���(����,�ڰ�,����)
    stsAccessSave = 1
    stsAccessVerify = 2
    stsAskSave = 3
    stsAskVerify = 4
    stsDonation = 5
    stsFinish = 6
    stsPrint = 7
    
    '�����ڻ���(pheresis)
End Enum

Public Enum XMethod
    xmMAJOR = 0
    xmMINOR
    xmBOTH
    xmNONE
End Enum


Public CD2_Buildings        As String

'================================
'        �����ͼ� �⺻����
'================================
Public PRESENTDATE_FORMAT   As String               '��¥����="YYYYMMDD"
Public PRESENTTIME_FORMAT   As String               '��¥����="HHMMSS"

Public C_WORKAREA           As String               ' Workarea='B'
Public BC2_KEEP_HOUR        As String               ' xm��ü�����ϼ�(�ð�)
Public BC2_XM_STEP          As String               ' xm�˻� step
Public BN_SPC_NO            As String               ' ��ü��ȣ
Public BN_ACC_NO            As String               ' ������ȣ
Public BC2_COMPO            As String               ' ��������
Public BC2_BRANCH           As String               ' �ں����ڵ�
Public BC2_LOCAL            As String               ' local ����
Public BC2_BLD_SRC          As String               ' ���׿�
Public BC2_EXP_RESON        As String               ' ������
Public BC2_ACTIVE_BLD_SRC   As String               ' �ŷ� ���׿�
Public BC2_ACCNO_CRITERION  As String               ' ������ȣ ��������
Public BC2_CENTER           As String               ' �����ڵ�
Public BC2_Trans_Reason     As String               ' ��������
Public BC2_REACTION_REASON  As String               ' ���� ���ۿ���
Public BC2_REACTION_TEST    As String               ' ���� ���ۿ�� �˻��׸�
Public BC2_CANCELRSN        As String               ' ������һ���
Public BC2_REACTION_TEST2   As String               ' ������ �˻��� �ϼ�
Public BC2_RETURN_HOUR      As String               ' ����� ��ȯ ���ɽð�
Public BC2_COMPO_GROUP      As String               ' �������� ���׷�
Public BC2_File_Server      As String           ' ���ϼ�����ġ
Public BC2_SPC_DEFAULT      As String               ' ��ü��ȣ�տ� �ٴ� ����
Public BC2_CompoValue       As String               ' ���������� ���ݼ���

'================================
'        ABO���� ����
'================================
Public BC2_ABO_RH           As String               ' ABO�� RH�ڵ� (��������� �ѹ� �߻���Ű�� ����) cdval1 �ڵ�
Public BC2_RST_ABO          As String               ' ABO����ڵ�
Public BC2_RST_RH           As String               ' RH ����ڵ�
Public BC2_RST_ABOSUB       As String               ' ABO SUBTYPE  ����ڵ�
Public BC2_RST_RHSUB        As String               ' RH  SUBGROUP ����ڵ�
Public BC2_RST_DU           As String               ' DU Test ����ڵ�
Public BC2_ABO_TEST         As String               ' ABO��� �׸�
Public BC2_BARCODE          As String               ' ���ڵ� ��½� �ٴ� ����
Public BC2_ADDSPC           As String               ' �߰�ä�� ����
Public BC2_ORDVAL           As String               ' ����
Public BC2_REASONVAL        As String               ' ��������
'================================
'   ������ �˻��Ƿڿ��� ���
'================================
Public BN_DONOR_ID          As String               ' ������ ���� ID
Public CD1_MULTISPC         As String               ' ������ü
Public CD2_SPECIMEN         As String               ' ��ü
Public CD2_WORKAREA         As String               ' Work Area
Public CD2_SGROUP           As String               ' ��ü��
Public NO_SPECIMEN          As String               ' ��ü��ȣ
Public NO_LABNO             As String               ' ������ȣ
Public NO_WORKNO            As String               ' �Ϲ�Worksheet Unit
Public NO_WSUNIT            As String               ' �̻���Worksheet Unit
Public BN_TMP_ID            As String               ' ������ �ӽ� ȯ�� ID
Public BC2_KIT              As String               ' KIT (PHERESIS�� ���)
Public BC2_DONOR_REASON     As String               ' �����ڰ˻��Ƿ� ��������
Public BC2_PHERESIS_COMPO   As String               ' �������� ��������(01;����,02;������,03;������)cdval1
Public BC2_UNACCEPT_REASON  As String               ' ������ ����(01;ALT,02:B������.....��)CDVAL1�ڵ�
Public BC2_JOB              As String               ' ����
Public BC2_TMP_ID           As String               ' �˻��Ƿڿ� �ӽ�ȯ�� ID
Public BC2_MATERIAL         As String               ' �߰���᳻��
Public BC2_DS_TEST          As String               ' donor screening �˻��׸�
Public BC2_ASK              As String               ' ���������ڵ�
Public BC2_RELATION         As String               ' ȯ�ڿ��� ����


Public BC2_RC_COMPO As String   '���׿����� ����ϴ� ����
Public BC2_RC_ABO As String    '���׿����� ����ϴ� ABO TYPE
Public BC2_RC_VOL As String      '���׿����� ����ϴ� ������ �뷮�� ���Ѵ�.
Public BC2_XM_COMPO As String

'================================
'      ���ڵ� ��������
'================================
Public BC2_BLOOD_BAR        As String               '���ڵ� ��� 'ȣȯ�� �������� �����Ұ�.... FRM301_FILE �� ������ ��

'���º� ǥ�� �̸�
Public STS_NM_ORDER As String 'ó��
Public STS_NM_COLLECT As String 'ä��
Public STS_NM_ACCESS As String '����
Public STS_NM_INPROGRESS As String '������
Public STS_NM_REQUEST As String '��û
Public STS_NM_DONE As String '�Ϸ�
Public STS_NM_END As String '����

'-----------------------
Private Sub Class_Initialize()

    C_WORKAREA = "B"
    
    BC2_KEEP_HOUR = "B001"              ' xm��ü�����ϼ�(�ð�)
    BC2_XM_STEP = "B002"                ' xm�˻� step
    BC2_DS_TEST = "B003"                ' donor screening �˻��׸�
    BC2_COMPO = "B004"                  ' ��������
    BC2_BRANCH = "B005"                 ' �ں����ڵ�
    BC2_LOCAL = "B006"                  ' local ����
    BC2_BLD_SRC = "B007"                ' ���׿�
    BC2_ASK = "B008"                    ' ���������ڵ�
    BC2_RELATION = "B009"               ' ȯ�ڿ��� ����
    BC2_CompoValue = "B032"             ' ���������� ���ݼ���
    BC2_EXP_RESON = "B010"              ' ������
    BC2_ACTIVE_BLD_SRC = "B012"         ' �ŷ� ���׿�
    BC2_ACCNO_CRITERION = "B013"        ' ������ȣ ��������
    
    
    BC2_Trans_Reason = "B015"           ' ��������
    BC2_JOB = "B016"                    ' ����
    BC2_TMP_ID = "B017"                 ' �˻��Ƿڿ� �ӽ�ȯ�� ID
    BC2_MATERIAL = "B018"               ' �߰���᳻��
    BC2_REACTION_TEST = "B019"          ' ���� ���ۿ�� �˻��׸�
    BC2_CANCELRSN = "B020"              ' ������һ���
    BC2_ABO_TEST = "B021"               ' ABO����׸�
    BC2_REACTION_TEST2 = "B022"         ' ������ �˻��� �ϼ�
    BC2_RETURN_HOUR = "B023"            ' ��ȯ ���� �ð�(��)
    BC2_COMPO_GROUP = "B024"            ' �������� ��� �׷�
    BC2_KIT = "B025"                    ' KIT (PHERESIS�� ���)
    BC2_DONOR_REASON = "B026"           ' �����˻��Ƿ� ��������
    BC2_PHERESIS_COMPO = "B027"         ' �������� ��������(��¿�)
    BC2_UNACCEPT_REASON = "B028"        ' �����ݻ���(��¿�)
    BC2_ABO_RH = "B029"                 ' ABO & RH �ڵ�(��������� ���)
    BC2_BLOOD_BAR = "B030"              ' ���� ���ڵ� ���
    BC2_REACTION_REASON = "B031"        ' ���� ���ۿ� ����Ʈ
    
'�˻����ڵ�
    BC2_RST_ABO = "B111"
    BC2_RST_RH = "B112"
    BC2_RST_ABOSUB = "B113"
    BC2_RST_RHSUB = "B114"
    BC2_RST_DU = "B115"

'�������ڰ���
    BC2_File_Server = "B202"            ' ���ϼ�����ġ
    BC2_SPC_DEFAULT = "3"               ' ��ü��ȣ �տ� �ٴ� ����
    BC2_BARCODE = "000000000"           ' ���ڵ���½� �ٴ� ����
    BC2_ADDSPC = "�߰�ä��"             ' �߰�ä��
    BC2_ORDVAL = "A5"                   ' ��������
    BC2_REASONVAL = "B1"                ' ����
'������ �˻��Ƿڿ��� ���Ǵ� Cdindex
    CD1_MULTISPC = "C102"               ' ������ü
    CD2_SPECIMEN = "C215"               ' ��ü
    CD2_WORKAREA = "C213"               ' Work Area
    CD2_SGROUP = "C217"                 ' ��ü��

    BN_TMP_ID = "C001"                  ' �˻��Ƿڿ� �ӽ� ȯ�� ID
    BN_SPC_NO = "B002"                  ' ��ü ��ȣ
    BN_ACC_NO = "B004"                  ' ������ȣ
    BN_DONOR_ID = "B006"                ' ������ ID

'������ �˻��Ƿڿ��� ���Ǵ� Noindex
    NO_SPECIMEN = "L001"                ' ��ü��ȣ
    NO_LABNO = "L002"                   ' ������ȣ
    NO_WORKNO = "L003"                  ' �Ϲ�Worksheet Unit
    NO_WSUNIT = "L009"                  ' �̻���Worksheet Unit

    BC2_CENTER = "C230"
    
    BC2_RC_COMPO = "B301"    '���׿����� ����ϴ� ����
    BC2_RC_ABO = "B302"  '���׿����� ����ϴ� ABO TYPE
    BC2_RC_VOL = "B303"  '���׿����� ����ϴ� ���� �� �뷮�� ����
    BC2_XM_COMPO = "B032" 'XM�ʼ� �Է� ����
    
'FORMAT �Լ� ����
    PRESENTDATE_FORMAT = "YYYYMMDD"
    PRESENTTIME_FORMAT = "HHMMSS"
    
    STS_NM_ORDER = "ó��" 'ó��
    STS_NM_COLLECT = "ä��" 'ä��
    STS_NM_ACCESS = "����" '����
    STS_NM_INPROGRESS = "������" '������
    STS_NM_REQUEST = "��û" '��û
    STS_NM_DONE = "�Ϸ�" '�ϰ�
    STS_NM_END = "����" '����
End Sub
