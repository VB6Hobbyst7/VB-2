Attribute VB_Name = "modConst"
Option Explicit

'-- �������̽� ȯ������
Public Const colSPECNO = 0      '�̻��
Public Const colCHECKBOX = 1
Public Const colEXAMDATE = 2    '�������̽�����
Public Const colEXAMTIME = 3    '�������̽�����
Public Const colSAVESEQ = 4     '�������(��¥��)
Public Const colER = 5          '���޿���
Public Const colRT = 6          '���
Public Const colHOSPDATE = 7    '������������
Public Const colBARCODE = 8     '��ü��ȣ(���ڵ�)
Public Const colSPECIMEN = 9    '��ü
Public Const colRACKNO = 10     '����ȣ
Public Const colPOSNO = 11      '������
Public Const colSEQNO = 12      '�Ϸù�ȣ
Public Const colPNAME = 13      '�̸�
Public Const colPSEX = 14       '����
Public Const colPAGE = 15       '����
Public Const colPID = 16        '���Ϲ�ȣ,ȯ�ڹ�ȣ,������ȣ
Public Const colCHARTNO = 17    'íƮ��ȣ
Public Const colDEPT = 18       '�Ƿڰ�
Public Const colINOUT = 19      '�Կ�/�ܷ�
Public Const colOCNT = 20       '��������
Public Const colRCNT = 21       '�������
Public Const colSTATE = 22      '�˻����
Public Const colITEMS = 23      '�˻��'s (��ũ����Ʈ ��)

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
Public Const colRPREVRESULT = 13

'-- �˻縶����
Public Const colLSPECNO = 0
Public Const colLMACHCODE = 1
Public Const colLSEQNO = 2
Public Const colLOCHANNEL = 3
Public Const colLRCHANNEL = 4
Public Const colLTESTCD = 5
Public Const colLTESTNM = 6
Public Const colLABBRNM = 7
Public Const colLRESSPECUSE = 8
Public Const colLRESSPEC = 9
Public Const colLMLOW = 10
Public Const colLMHIGH = 11
Public Const colLFLOW = 12
Public Const colLFHIGH = 13

'-- QC
Public Const colLQCLab = 22
Public Const colLQCLot = 23
Public Const colLQCAnalyte = 24
Public Const colLQCMethod = 25
Public Const colLQCInstrument = 26
Public Const colLQCReagent = 27
Public Const colLQCUnit = 28
Public Const colLQCTemp = 29

Public Const colLUseResSpec = 30

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
Public Const EF  As String = ""    'EOF Chr(26)


Public pBuffer          As Variant
Public RcvBuffer        As String
Public strRecvData()    As String
Public intPhase         As Integer
Public strState         As String
Public intBufCnt        As Integer
Public blnIsETB         As Boolean
Public intSndPhase      As Integer
Public intFrameNo       As Integer

Public strErrMsg        As String
