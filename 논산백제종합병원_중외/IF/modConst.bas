Attribute VB_Name = "modConst"
Option Explicit

'-- �������̽� ȯ������
Public Const colSPECNO = 0      '�̻��
Public Const colCHECKBOX = 1
Public Const colEXAMDATE = 2    '�������̽�����
Public Const colSAVESEQ = 3     '�������(��¥��)
Public Const colHOSPDATE = 4    '������������
Public Const colGUBUN = 3
Public Const colBARCODE = 5     '���ڵ�
Public Const colSEQNO = 6       '�Ϸù�ȣ
Public Const colRACKNO = 7      '����ȣ
Public Const colPOSNO = 8       '������
Public Const colINOUT = 9       '�Կ�/�ܷ�
Public Const colCHARTNO = 10    'íƮ��ȣ
Public Const colPID = 11        'ȯ�ڹ�ȣ,���Ϲ�ȣ,������ȣ
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
Public Const colLGUBUN = 3
Public Const colLOCHANNEL = 4
Public Const colLRCHANNEL = 5
Public Const colLTESTCD = 6
Public Const colLTESTNM = 7
Public Const colLABBRNM = 8
Public Const colLRESSPEC = 9
Public Const colLLOW = 10
Public Const colLHIGH = 11
Public Const colLLOWF = 12
Public Const colLHIGHF = 13
Public Const colLRSTTYPE = 14
Public Const colLCUTUSE = 15
Public Const colLCOLIN = 16
Public Const colLCOLCOMP = 17
Public Const colLCOLOUT = 18
Public Const colLCOMOUT = 19
Public Const colLCOHIN = 20
Public Const colLCOHCOMP = 21
Public Const colLCOHOUT = 22
'-- QC
Public Const colLQCLab = 23
Public Const colLQCLot = 24
Public Const colLQCAnalyte = 25
Public Const colLQCMethod = 26
Public Const colLQCInstrument = 27
Public Const colLQCReagent = 28
Public Const colLQCUnit = 29
Public Const colLQCTemp = 30

Public Const colLUseResSpec = 31


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
Public Const R_S As String = ""
Public Const SB  As String = ""    'Chr(11)
Public Const EB  As String = ""     'Chr(28)
Public Const SYN As String = ""    'Chr(22)
Public Const EF As String = ""    'EOF Chr(26)


Public RcvBuffer        As String
Public strRecvData()    As String
Public intPhase         As Integer
Public strState         As String
Public intBufCnt        As Integer
Public blnIsETB         As Boolean
Public intSndPhase      As Integer
Public intFrameNo       As Integer

'===============================

Public strErrMsg   As String
