Attribute VB_Name = "modConst"
Option Explicit

'-- 인터페이스 환자정보
Public Const colSPECNO = 0      '미사용
Public Const colCHECKBOX = 1
Public Const colEXAMDATE = 2    '인터페이스일자
Public Const colEXAMTIME = 3    '인터페이스일자
Public Const colSAVESEQ = 4     '저장순번(날짜별)
Public Const colER = 5          '응급여부
Public Const colRT = 6          '재검
Public Const colHOSPDATE = 7    '병원접수일자
Public Const colBARCODE = 8     '검체번호(바코드)
Public Const colSPECIMEN = 9    '검체
Public Const colRACKNO = 10     '렉번호
Public Const colPOSNO = 11      '포지션
Public Const colSEQNO = 12      '일련번호
Public Const colPNAME = 13      '이름
Public Const colPSEX = 14       '성별
Public Const colPAGE = 15       '나이
Public Const colPID = 16        '병록번호,환자번호,내원번호
Public Const colCHARTNO = 17    '챠트번호
Public Const colDEPT = 18       '의뢰과
Public Const colINOUT = 19      '입원/외래
Public Const colOCNT = 20       '오더갯수
Public Const colRCNT = 21       '결과갯수
Public Const colSTATE = 22      '검사상태
Public Const colITEMS = 23      '검사명's (워크리스트 용)

'-- 인터페이스 결과
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

'-- 검사마스터
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
