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
Public Const colSAVESEQ = 2     '�������(��¥��)
Public Const colEXAMDATE = 3    '�������̽�����
Public Const colHOSPDATE = 4    '������������
Public Const colRCPDATE = 5     '�Ƿ�����
Public Const colJUBNO = 6       '������ȣ
Public Const colCHARTNO = 7     '��Ϲ�ȣ
Public Const colPNAME = 8       '�̸�
Public Const colPSEX = 9        '����
Public Const colPAGE = 10       '����
Public Const colPART = 11       '��
Public Const colROOM = 12       '����
Public Const colTESTCD = 13     '�˻��ڵ�
Public Const colTESTNM = 14     '�˻��׸�
Public Const colTESTDATE = 15   '�˻������
Public Const colSPCPART = 16    '��ü����
Public Const colBARCODE = 17    '��ü��ȣ

Public Const colRELTEST = 18    '��������

Public Const colSPCCD = 19      '��ü�ڵ�
Public Const colSPCNM = 20      '��ü��
Public Const colRESULT = 21     '�˻���

Public Const colHPVIC = 22      'IC
Public Const colPRERESULT = 23  '�������
Public Const colMETHOD = 24     'Method
Public Const colREMARK = 25     'Remakr


Public Const colRSTDATE = 26    '�˻纸����
Public Const colDOCTOR = 27     '�ǵ��ǻ�
Public Const colPRINT = 28      '�������
Public Const colSTATE = 29      '�˻����
'-- ��ũ����Ʈ ��
Public Const colITEMS = 30

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
