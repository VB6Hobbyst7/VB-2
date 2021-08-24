Attribute VB_Name = "modIISCommon"
'-----------------------------------------------------------------------------'
'   ���ϸ�  : modIISCommon.bas
'   ��  ��  : �����ڵ� Ŭ����
'   ��  ��  :
'          - �̻��� �Ϲݰ�����(IISGENSENSI), MIC(IISMIC) ������� �����߰�
'          - PT(%)���� ����ڵ�(IISPT) �����߰�
'          - CONCAT(FCONCAT) �����߰�
'          - �˻��׸� ����ڵ�(mCRESULTCD) �����߰�
'          - CBC Workarea(mIISCBCWA) �����߰�
'          - ���Ѱ�(mIISLMTLOW), ���Ѱ�(mIISLMTHIGH) ����ڵ� �����߰�
'-----------------------------------------------------------------------------'

Option Explicit

'## Database ��������
Public mDbCon           As ADODB.Connection     '����Db Connection
Public mCliCon          As ADODB.Connection     'ClientDb Connection
Public mDbType          As String               'DB Type (0:Oracle, 1:Sybase, 2:MS-SQL, 3:ACCESS)
Public mSource          As String               'Data Source
Public mCatalog         As String               'Initial Catalog
Public mUid             As String               'User ID
Public mPwd             As String               'Password
Public mUserCancel      As Boolean              '����ڰ� DB������ �����ϸ� True

'## Error �÷���
Public mError           As clsIISError          '���� Ŭ����

'## ������Ʈ���� DB���� ����
Public Const cDBSERVER  As String = "DbServer"
Public Const cDBTYPE    As String = "DbType"
Public Const cSOURCE    As String = "Source"
Public Const cCATALOG   As String = "Catalog"
Public Const cUID       As String = "Uid"
Public Const cPWD       As String = "Pwd"

'## ������Ʈ�� ����
Public mAppName         As String       'App Name
Public mExePath         As String       'EXE ���ϰ��
Public mLogPath         As String       'Log ���ϰ��
Public mClientDbPath    As String       'ClientDb ���+���ϸ�
Public mIniPath         As String       'INI ���ϰ��

'## ����� ����
Public mEmpId           As String       '����� ���̵�
Public mEmpNm           As String       '����� �̸�

'## mdiIISMain Form
Public mMainFrm         As Object       'mdiIISMain ��

'## ��ű�ȣ
Public Const mENQ As Long = &H5         'Chr(5),  ""
Public Const mACK As Long = &H6         'Chr(6),  ""
Public Const mSTX As Long = &H2         'Chr(2),  ""
Public Const mETB As Long = &H17        'Chr(23), ""
Public Const mETX As Long = &H3         'Chr(3),  ""
Public Const mEOT As Long = &H4         'Chr(4),  ""
Public Const mNAK As Long = &H15        'Chr(21), ""
Public Const mSOH As Long = &H1         'Chr(1),  ""
Public Const mDLE As Long = &H10        'Chr(16), ""
Public Const mSYN As Long = &H16        'Chr(22), ""

'## ������ ����
Public mPROJECTCODE     As String       'Project Code
Public mPROJECTTYPE     As String       'Project Type(A:�ڻ�, B:Ÿ��, C:����)
Public mHOSPITALNM      As String       '�����̸�
Public mSPCLEN          As Long         'SPC Length
Public mSPCYYLEN        As Long         'SPCYY Length
Public mSPCNOLEN        As Long         'SPCNO Length
Public mIISNEGATIVE     As String       'Negative
Public mIISPOSITIVE     As String       'Positive
Public mIISGRAYZONE     As String       'Grayzone
Public mIISERROR        As String       'Error
Public mIISSPCLEN       As Long         '��ü���б���
Public mIISSPCSERUM     As String       'Sereum
Public mIISSPCURINE     As String       'Urine
Public mIISSPCPLASMA    As String       'Plasma
Public mIISSPCCSF       As String       'CSF
Public mIISSPCBLOOD     As String       'Blood
Public mIISSPCFLUID     As String       'Boyd Fluid
Public mIISSPCCAPD      As String       'CAPD
Public mIISQCLOW        As String       'QC Low Level
Public mIISQCNORMAL     As String       'QC Normal Level
Public mIISQCHIGH       As String       'QC High Level
Public mIISPANICCHECK   As String       'Panic üũ ���̺�
Public mIISMICWA        As String       '�̻��� Workarea
Public mIISMQTCD        As String       '�̻��� ��������� �����ڵ�
Public mIISGENSENSI     As String       '�̻��� �Ϲݰ����� ����ڵ�
Public mIISMIC          As String       '�̻��� MIC ����ڵ�
Public mIISSERUMINDEX   As String       'Hitachi7600����� Serum Index �������(0:��,1:��)
Public mIISPT           As String       'PT(%)���� 100�̻��϶� ����ڵ�
Public mIISCBCWA        As String       'CBC Workarea

Public mIISREACTIVE     As String       'Reactive
Public mIISNREACTIVE    As String       'NonReactive
Public mIISWPOSITIVE    As String       'WaekPositive

'   - ���Ѱ�, ���Ѱ� ����
Public mIISLMTLOW       As String       '���Ѱ�
Public mIISLMTHIGH      As String       '���Ѱ�

'## �����ڵ� �ε���
Public mCODE            As String       'CDINDEX
Public mCSPCCD          As String       '��ü
Public mCWACD           As String       'WorkArea
Public mCDETAILCD       As String       '���׸�
Public mCPANELCD        As String       '�׷��׸�
Public mCREPEATCD       As String       '�ٺ�ó��
Public mCLOCATIONCD     As String       '�ǹ��ڵ�
Public mCVANDCD         As String       '��ü�ڵ�
Public mCFOOTNOTECD     As String       'FootNote
Public mCSPCRMKCD       As String       '��ü Remark
Public mCACCRSNCD       As String       '������� ����
Public mCMDYRSNCD       As String       '������� ����
Public mCQCREJECTCD     As String       'QC Reject ����
Public mCMnmCd          As String       '���ڵ�
Public mCRESULTCD       As String       '�˻��׸� ����ڵ�

'## ���̺�
Public mTHIS001         As String       'ȯ�� ������
Public mTHIS002         As String       '���� ������
Public mTHIS003         As String       '����� ������
Public mTHIS004         As String       'ó���� ������
Public mTHIS005         As String       '���� ������1 (�α�������)
Public mTHIS006         As String       '���� ������2 (��������)

Public mTCOM001         As String       '�����ڵ� ������
Public mTIIS001         As String       '�����ڵ� ������1
Public mTIIS002         As String       '�����ڵ� ������2
Public mTIIS003         As String       '���ø� ������

Public mTIIS101         As String       '��ü ������
Public mTIIS102         As String       '�˻��׸� ������
Public mTIIS103         As String       '������ü ������
Public mTIIS104         As String       '����ġ ������
Public mTIIS105         As String       '���ø� ������
Public mTIIS106         As String       'ǲ��Ʈ ������
Public mTIIS107         As String       'QC �Ұ߸�����

Public mTIIS201         As String       'ó�泻��(H)
Public mTIIS202         As String       'ó�泻��(B)
Public mTIIS203         As String       '��������
Public mTIIS204         As String       '�������
Public mTIIS205         As String       '������ ���
Public mTIIS206         As String       '����� ���

Public mTIIS301         As String       'QC ��Ʈ��(H)
Public mTIIS302         As String       'QC ��Ʈ��(B)
Public mTIIS303         As String       'QC ������(H)
Public mTIIS304         As String       'QC ������(B)
Public mTIIS305         As String       'QC ������
Public mTIIS306         As String       'QC ����
Public mTIIS307         As String       'QC ���
Public mTIIS308         As String       'QC �Ұ߳���

Public mTIIS401         As String       '��� ������
Public mTIIS402         As String       '������ ������
Public mTIIS403         As String       '��� �˻��׸� ������(H)
Public mTIIS404         As String       '��� �˻��׸� ������(B)
Public mTIIS405         As String       '������۳���
Public mTIIS406         As String       '����������

Public mTIIS501         As String       '�̻��� WorkSheet(H)
Public mTIIS502         As String       '�̻��� WorkSheet(B)
Public mTIIS503         As String       '�̻��� WorkSheet �߰�����
Public mTIIS504         As String       '�̻��� �������
Public mTIIS505         As String       '�̻��� ������ �������

'   - Ư���˻� ������� ���̺��� �߰�
Public mTIIS601         As String       'Ư���˻� �������

Public mTACC203         As String       'ClientDb ��������
Public mTACC204         As String       'ClientDb �������

'## �ʵ�
'HIS001 (ȯ�� ������)
Public mFPTID           As String       'ȯ��ID
Public mFPTNM           As String       '�̸�
Public mFJUMIN          As String       '�ֹι�ȣ
Public mFSEX            As String       '����
Public mFAGE            As String       '����

'HIS002 (�μ� ������)
Public mFDEPTCD         As String       '�μ��ڵ�
Public mFDEPTNM         As String       '�μ���

'HIS003 (���� ������)
Public mFWARDCD         As String       '�����ڵ�
Public mFWARDNM         As String       '������

'HIS004 (ó���� ������)
Public mFDOCTCD         As String       'ó�����ڵ�
Public mFDOCTNM         As String       'ó���Ǹ�

'HIS006 (���� ������, ��������)
Public mFEMPID          As String       '����ID
Public mFEMPNM          As String       '�����̸�

'�����ڵ� ������2
Public mFSPCCD          As String       '��ü�ڵ�
Public mFSPCNM          As String       '��ü��

'## Database�� ������
Public mFCONCAT         As String       'Concatenate ������(Oracle:||, MS-SQL:+)
