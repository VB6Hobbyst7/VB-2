Attribute VB_Name = "modConst"
Option Explicit


'    acpt.instcd , --�����ȣ
'    acpt.acptdd , --��������
'    acpt.acptno , --������ȣ
'    acpt.acptitemno , --�����׸��ȣ
'    acpt.PTNO , --������ȣ
'    acpt.PID , --��Ϲ�ȣ
'    acpt.TESTCD , --�˻��ڵ�
'    test.testengnm , --�����˻��
'    spcm.SPCNM , --��ü��
'    acpt.prcpgenrflag , --�Կ� / �ܷ�����
'    dept.deptengabbr , --�������
'    acpt.prcpdd,        --ó������,
'    acpt.execprcpuniqno , --�ǽ�ó�����Ϲ�ȣ
'    acpt.prcpno , --ó���ȣ
'    ptbs.hngnm , --ȯ�ڸ�
'    ptbs.sex , --����
'    ptbs.brthdd , --����
'    com.fn_zz_getage(ptbs.rrgstno1, ptbs.rrgstno2, acpt.acptdd, 'A', ptbs.brthdd) as age  -- �������ڱ��� ����
'

'-- �������̽� ȯ������
Public Const colSPECNO = 0      '�̻��
Public Const colCHECKBOX = 1
Public Const colEXAMDATE = 2    '�������̽�����
Public Const colSAVESEQ = 3     '�������(��¥��)
Public Const colHOSPDATE = 4    '������������                   ==> ��������
'Public Const colBARCODE = 5     '���ڵ�                         ==> ������ȣ
Public Const colBARCODE = 9     '���ڵ�                         ==> ������ȣ
Public Const colSEQNO = 6       '�Ϸù�ȣ                       ==> �����׸��ȣ
Public Const colRACKNO = 7      '����ȣ                         ==> ������ڵ�(X)
Public Const colPOSNO = 8       '������                         ==> �������
'Public Const colCHARTNO = 9     'íƮ��ȣ                       ==> ������ȣ
Public Const colCHARTNO = 5     'íƮ��ȣ                       ==> ������ȣ
Public Const colPID = 10        'ȯ�ڹ�ȣ,���Ϲ�ȣ,������ȣ     ==> ��Ϲ�ȣ
Public Const colINOUT = 11      '�Կ�/�ܷ�
Public Const colPNAME = 12      '�̸�
Public Const colPSEX = 13       '����
Public Const colPAGE = 14       '����
Public Const colPJUMIN = 15     '�ֹ�                           ==> ó���ȣ
Public Const colKEY1 = 16       '����1                          ==> ó������
Public Const colKEY2 = 17       '����2                          ==> �ǽ�ó�����Ϲ�ȣ
Public Const colOCNT = 18       '��������                       ==> �˻��
Public Const colRCNT = 19       '�������                       ==> ��ü��
Public Const colSTATE = 20      '�˻����
'-- ��ũ����Ʈ ��
Public Const colITEMS = 21

'-- �������̽� ���
Public Const colRSPECNO = 0
Public Const colRCHECKBOX = 1
Public Const colRSEQNO = 2
Public Const colRORDERCD = 3
Public Const colRTESTCD = 4
Public Const colRSUBCD = 5
Public Const colRTESTNM = 6
Public Const colRCHANNEL = 7
Public Const colRMACHRESULT = 8
Public Const colRLISRESULT = 9
Public Const colRFLAG = 10
Public Const colRJUDGE = 11
Public Const colRREF = 12

'-- �˻縶����
Public Const colLSPECNO = 0
Public Const colLMACHCODE = 1
Public Const colLSEQNO = 2
Public Const colLOCHANNEL = 3
Public Const colLRCHANNEL = 4
Public Const colLTESTCD = 5
Public Const colLTESTNM = 6
Public Const colLABBRNM = 7
Public Const colLRESSPEC = 8
Public Const colLLOW = 9
Public Const colLHIGH = 10
Public Const colLRSTTYPE = 11
Public Const colLCUTUSE = 12
Public Const colLCOLIN = 13
Public Const colLCOLCOMP = 14
Public Const colLCOLOUT = 15
Public Const colLCOMOUT = 16
Public Const colLCOHIN = 17
Public Const colLCOHCOMP = 18
Public Const colLCOHOUT = 19


'===============================
Public Const SPCLEN As Integer = 10

Public Const STX As String = ""
Public Const ETX As String = ""
Public Const ENQ As String = ""
Public Const ACK As String = ""
Public Const NAK As String = ""
Public Const EOT As String = ""
Public Const ETB As String = ""
Public Const FS  As String = ""
Public Const RS  As String = ""
Public Const GS  As String = ""
Public Const SB As String = ""  'Chr(11)
Public Const EB As String = ""   'Chr(28)


Public strRecvData()    As String
Public intPhase         As Integer
Public strState         As String
Public intBufCnt        As Integer
Public blnIsETB         As Boolean
Public intSndPhase      As Integer
Public intFrameNo       As Integer
'===============================

Public strErrMsg   As String
