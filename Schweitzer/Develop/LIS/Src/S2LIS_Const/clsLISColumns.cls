VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISColumns"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum enORDSHEET  'ó�������̺�
    tcORDNO = 1
    tcTESTNM = 2
    tcTESTCD = 3
    tcSPCCD = 4
    tcINSURFG = 5
    tcSTATCHK = 6
    tcREQDTTM = 7
    tcSTATFG = 8
    tcWORKAREA = 9
    tcSTORECD = 10
    tcRNDFG = 11
    tcTESTDIV = 12
    tcMULTIFG = 13
    tcSPCGRP = 14
    tcORDSEQ = 15
    tcABBRNM = 16
    tcBARCNT = 17
    tcTESTFLAG = 18
    tcBUILDCD = 19
    tcBUILDNM = 20
    tcSPCABBR = 21
    tcLABDIV = 22
    tcLABRANGE = 23
End Enum

Public Enum enWARDCOLLIST   '�����ϰ�ä�����̺�
    tcCHECK = 1
    tcWARD = 2
    tcPTID = 3
    tcPTNM = 4
    tcSTATSPC = 5
    tcGENSPC = 6
    tcMONCOL = 7
    tcDOB = 8
    tcBedInDT = 9
    tcDEPTCD = 10
    tcORDDOCT = 11
    tcMAJDOCT = 12
    tcWARDID = 13
    tcROOMID = 14
    tcPTSEX = 15
End Enum

Public Enum enCOLLIST   '�ܷ�/���� ä�����̺�
    tcCHECK = 1
    tcORDDT = 2
    tcORDNO = 3
    tcTESTNM = 4
    tcSPCNM = 5
    tcDOCTNM = 6
    tcSTATFG = 7
    tcREQDTTM = 8
    tcORDDATE = 9
    tcORDNUM = 10
    tcORDSEQ = 11
    tcTESTCD = 12
    tcSPCCD = 13
    tcWORKAREA = 14
    tcSTORECD = 15
    tcTESTDIV = 16
    tcMULTIFG = 17
    tcSPCGRP = 18
    tcORDDOCT = 19
    tcMAJDODT = 20
    tcSTATFLAG = 21
    tcDEPTCD = 22
    tcABBRNM = 23
    tcBARCNT = 24
    tcBUILDCD = 25
    tcBUILDNM = 26
    tcPAYDT = 27
    tcSPCABBR = 28
    tcLABDIV = 29
    tcLABRANGE = 30
    tcWARDID = 31
    tcROOMID = 32
    tcBEDID = 33
    tcORDDIV = 34
    tcFRZFG = 35
    tcHold = 36
    tcBedInDT = 37
End Enum

Public Enum enOUTLAB    '�ܺ��Ƿ����̺�
    tcCHECK = 1
    tcLABNO
    tcPTID
    tcPTNM
    tcDEPTNM
    tcSEX
    tcAGE
    tcTESTNM
    tcINSUR
    tcSPCNM
    tcSTSCD
    tcWORKAREA
    tcACCDT
    tcACCSEQ
    tcTESTCD
End Enum

Public Enum enREVIEW1   'ó��װ����ȸ ���̺�
    tcORDDT = 1
    tcORDNO
    tcTESTNM
    tcSPCNM
    tcSTATFG
    tcDOCTNM
    tcRCVDT
    tcSTSNM
    tcORDDATE
    tcCOLDATE
    tcORDDOCT
    tcSPCNAME
    tcORDNUM
    tcWORKAREA
    tcACCDT
    tcACCSEQ
    tcVFYNM
    tcVFYDATE
    tcTAT
    tcTESTCD
    tcSPCCD
    tcSPCYY
    tcSPCNO
    tcORDDIV
    tcUNITQTY
    tcREQDATE
    tcREQTIME
    tcWARDID
    tcHOSILID
    tcSTSCD
End Enum

Public Enum enSPCGRP    '��ü�����̺�
    tcGRPNM = 0
    tcGRPCD
    tcWSUNIT
    tcFNSHDT
    tcFNSHTM
    tcWORKAREA
    tcFROMSEQ
    tcTOSEQ
    tcWSGRP
    tcSELFG
End Enum

'�̻��� �����ΰ�� ���̺�
Public Enum enSTAIN
    tcTESTNM = 1    '�˻��׸�
    tcTESTCD        '�ڵ�
    tcRSTDIV        '���Ÿ��
    tcDETAIL        '������
    tcRSTCD         '����ڵ�
    tcRSTNM         '�˻���
    tcLASTRST       '�ֱٰ��
    tcEXCPT         '����
    tcOLDRST        '���������
    tcOLDDT         'ó����
    tcOLDNO         'ó���ȣ
    tcOLDSEQ        'ó��Seq
End Enum


'QC Table �� Column
Public Enum enQC
    SCOL_TESTNM = 1
    SCOL_RESULT = 2
    SCOL_RSTUNIT = 3
    SCOL_RADIV = 4
    SCOL_EQPNM = 9
    SCOL_TXT = 10
End Enum