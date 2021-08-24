Attribute VB_Name = "modConst"
Option Explicit


'    acpt.instcd , --기관기호
'    acpt.acptdd , --접수일자
'    acpt.acptno , --접수번호
'    acpt.acptitemno , --접수항목번호
'    acpt.PTNO , --병리번호
'    acpt.PID , --등록번호
'    acpt.TESTCD , --검사코드
'    test.testengnm , --영문검사명
'    spcm.SPCNM , --검체명
'    acpt.prcpgenrflag , --입원 / 외래구분
'    dept.deptengabbr , --진료과명
'    acpt.prcpdd,        --처방일자,
'    acpt.execprcpuniqno , --실시처방유일번호
'    acpt.prcpno , --처방번호
'    ptbs.hngnm , --환자명
'    ptbs.sex , --성별
'    ptbs.brthdd , --생일
'    com.fn_zz_getage(ptbs.rrgstno1, ptbs.rrgstno2, acpt.acptdd, 'A', ptbs.brthdd) as age  -- 접수일자기준 나이
'

'-- 인터페이스 환자정보
Public Const colSPECNO = 0      '미사용
Public Const colCHECKBOX = 1
Public Const colSAVESEQ = 2     '저장순번(날짜별)
Public Const colEXAMDATE = 3    '인터페이스일자
Public Const colHOSPDATE = 4    '병원접수일자
Public Const colRCPDATE = 5     '의뢰일자
Public Const colJUBNO = 6       '접수번호
Public Const colCHARTNO = 7     '등록번호
Public Const colPNAME = 8       '이름
Public Const colPSEX = 9        '성별
Public Const colPAGE = 10       '나이
Public Const colPART = 11       '과
Public Const colROOM = 12       '병실
Public Const colTESTCD = 13     '검사코드
Public Const colTESTNM = 14     '검사항목
Public Const colTESTDATE = 15   '검사시행일
Public Const colSPCPART = 16    '검체종류
Public Const colBARCODE = 17    '검체번호

Public Const colRELTEST = 18    '상관성결과

Public Const colSPCCD = 19      '검체코드
Public Const colSPCNM = 20      '검체명
Public Const colRESULT = 21     '검사결과

Public Const colHPVIC = 22      'IC
Public Const colPRERESULT = 23  '이전결과
Public Const colMETHOD = 24     'Method
Public Const colREMARK = 25     'Remakr


Public Const colRSTDATE = 26    '검사보고일
Public Const colDOCTOR = 27     '판독의사
Public Const colPRINT = 28      '보고서출력
Public Const colSTATE = 29      '검사상태
'-- 워크리스트 용
Public Const colITEMS = 30

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
