Attribute VB_Name = "modConst"
Option Explicit

'-- 인터페이스 환자정보
Public Const colSPECNO = 0      '미사용
Public Const colCHECKBOX = 1
Public Const colEXAMDATE = 2    '인터페이스일자
Public Const colSAVESEQ = 3     '저장순번(날짜별)
Public Const colHOSPDATE = 4    '병원접수일자
Public Const colBARCODE = 5     '바코드
Public Const colSEQNO = 6       '일련번호
Public Const colRACKNO = 7      '렉번호
Public Const colPOSNO = 8       '포지션
Public Const colINOUT = 9       '입원/외래
Public Const colCHARTNO = 10    '챠트번호
Public Const colPID = 11        '환자번호,병록번호,내원번호
Public Const colPNAME = 12      '이름
Public Const colPSEX = 13       '성별
Public Const colPAGE = 14       '나이
Public Const colPJUMIN = 15     '주민
Public Const colKEY1 = 16       '여분1
Public Const colKEY2 = 17       '여분2
Public Const colOCNT = 18       '오더갯수
Public Const colRCNT = 19       '결과갯수
Public Const colSTATE = 20      '검사상태
'-- 워크리스트 용
Public Const colITEMS = 21

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

'-- 검사마스터
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
Public Const colLLOWF = 11
Public Const colLHIGHF = 12
Public Const colLRSTTYPE = 13
Public Const colLCUTUSE = 14
Public Const colLCOLIN = 15
Public Const colLCOLCOMP = 16
Public Const colLCOLOUT = 17
Public Const colLCOMOUT = 18
Public Const colLCOHIN = 19
Public Const colLCOHCOMP = 20
Public Const colLCOHOUT = 21
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
Public strQState        As String
Public strLogData       As String

Public blnIsSB         As Boolean
'===============================

Public strErrMsg   As String
