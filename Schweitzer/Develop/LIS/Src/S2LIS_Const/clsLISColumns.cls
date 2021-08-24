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

Public Enum enORDSHEET  '처방등록테이블
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

Public Enum enWARDCOLLIST   '병동일괄채혈테이블
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

Public Enum enCOLLIST   '외래/병동 채혈테이블
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

Public Enum enOUTLAB    '외부의뢰테이블
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

Public Enum enREVIEW1   '처방및결과조회 테이블
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

Public Enum enSPCGRP    '검체군테이블
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

'미생물 스테인결과 테이블
Public Enum enSTAIN
    tcTESTNM = 1    '검사항목
    tcTESTCD        '코드
    tcRSTDIV        '결과타입
    tcDETAIL        '디테일
    tcRSTCD         '결과코드
    tcRSTNM         '검사결과
    tcLASTRST       '최근결과
    tcEXCPT         '보류
    tcOLDRST        '수정전결과
    tcOLDDT         '처방일
    tcOLDNO         '처방번호
    tcOLDSEQ        '처방Seq
End Enum


'QC Table 의 Column
Public Enum enQC
    SCOL_TESTNM = 1
    SCOL_RESULT = 2
    SCOL_RSTUNIT = 3
    SCOL_RADIV = 4
    SCOL_EQPNM = 9
    SCOL_TXT = 10
End Enum
