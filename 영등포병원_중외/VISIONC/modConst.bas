Attribute VB_Name = "modConst"
Option Explicit

'-- �������̽� ȯ������
Public Const colSPECNO = 0      '�̻��
Public Const colCHECKBOX = 1
Public Const colEXAMDATE = 2    '�������̽�����
Public Const colSAVESEQ = 3     '�������(��¥��)
Public Const colHOSPDATE = 4    '������������
Public Const colBARCODE = 5     '���ڵ�
Public Const colSEQNO = 6       '�Ϸù�ȣ
Public Const colRACKNO = 7      '����ȣ
Public Const colPOSNO = 8       '������
Public Const colCHARTNO = 9     'íƮ��ȣ
Public Const colPID = 10        'ȯ�ڹ�ȣ,���Ϲ�ȣ,������ȣ
Public Const colINOUT = 11      '�Կ�/�ܷ�
Public Const colPNAME = 12      '�̸�
Public Const colPSEX = 13       '����
Public Const colPAGE = 14       '����
Public Const colPJUMIN = 15     '�ֹ�
Public Const colKEY1 = 16       '����1
Public Const colKEY2 = 17       '����2
Public Const colOCNT = 18       '��������
Public Const colRCNT = 19       '�������
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
