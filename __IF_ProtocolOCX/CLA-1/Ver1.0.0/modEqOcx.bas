Attribute VB_Name = "modEqOcx"
Option Explicit

'--- MSComm event constants
Public Const MSCOMM_EV_SEND = 1
Public Const MSCOMM_EV_RECEIVE = 2
Public Const MSCOMM_EV_CTS = 3
Public Const MSCOMM_EV_DSR = 4
Public Const MSCOMM_EV_CD = 5
Public Const MSCOMM_EV_RING = 6
Public Const MSCOMM_EV_EOF = 7

'--- MSComm error code constants
Public Const MSCOMM_ER_BREAK = 1001
Public Const MSCOMM_ER_CTSTO = 1002
Public Const MSCOMM_ER_DSRTO = 1003
Public Const MSCOMM_ER_FRAME = 1004
Public Const MSCOMM_ER_OVERRUN = 1006
Public Const MSCOMM_ER_CDTO = 1007
Public Const MSCOMM_ER_RXOVER = 1008
Public Const MSCOMM_ER_RXPARITY = 1009
Public Const MSCOMM_ER_TXFULL = 1010

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Sample 에 대한 정보
Type SAMPLE_INFO
    ID          As String
    SEQNO       As String
    RACK        As String
    POS         As String
    QCGBN       As String
    KIND        As String       '1st/Rerun 구분
    ORDCNT      As Integer
    IFCD()      As String
    SVOL()      As String
    SINDEX      As Boolean      'Serum Index 검사여부
    PATINFO     As String       'Patient Info.
    SAMPINFO    As String       'Sample Info.
    SAMPTYPE    As String       'Sample Type(검체종류)
End Type

'수신된 결과에 대한 정보
Type RESULT_INFO
    ID          As String
    PATINFO     As String
    SAMPINFO    As String
    SEQNO       As String
    RACK        As String
    POS         As String
    QCGBN       As String
    KIND        As String
    RSTCNT      As Integer
    IFCD        As String
    RST1        As String
    RST2        As String
    RSTDT       As String
    UNIT        As String
    FLAG        As String
    ALARMCD     As String
    INSTID      As String
End Type

'임시로 비밀번호 설정
Public Const gcOpenPW = "ACK"
Public Const gcEditPW = "MEDI@CK"

Public gsSiteNm As String       '거래처정보

'
'   ASTM Protocol CheckSum 계산
'
Public Function ChkSum_ASTM(ByVal Para As String) As String

    Dim i   As Integer
    Dim Tmp As Integer
    Dim ChkS1   As Integer
    Dim ChkS2   As String
    
    For i = 1 To Len(Para)
        Tmp = Asc(Mid$(Para, i, 1))
        ChkS1 = ChkS1 + Tmp
    Next i
    ChkS1 = ChkS1 Mod 256
    ChkS2 = Right$("0" & Hex$(ChkS1), 2)
    
    ChkSum_ASTM = ChkS2
    
End Function

