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

'Sample �� ���� ����
Type SAMPLE_INFO
    ID      As String
    SEQNO   As String
    RACK    As String
    POS     As String
    QCGBN   As String
    KIND    As String       '1st/Rerun ����
    ORDCNT  As Integer
    IFCD()  As String
    SVOL()  As String
    SINDEX  As Boolean      'Serum Index �˻翩��
    SPCCD   As String       '��ü����
    RSTDT   As String       '2005/6/10 yk
    OTHER   As String
    CMT1    As String       '2005/8/1 yk
    INSTID  As String       '2006/2/9 yk
    INSTNM  As String       '2006/10/11 yk
    CONTROLNAME As String   '2007/01/17    by YeounJu
End Type

'���ŵ� ����� ���� ����
Type RESULT_INFO
    ID      As String
    SEQNO   As String
    RACK    As String
    POS     As String
    QCGBN   As String
    KIND    As String
    RSTCNT  As Integer
    IFCD    As String
    RST1    As String
    RST2    As String
    UNIT    As String
    FLAG    As String
    ALARMCD As String
    INSTID  As String
    INSTNM  As String       '2006/10/11 yk
    RSTDT   As String       'Date/Time Results Reported or Last Modified (2005/6/10 �߰� yk)
    SPCCD   As String       '��ü���� (2006/10/11 �߰� yk)
    OPERID  As String       'Operation ID (       "      )
    OTHER   As String
End Type

''�ӽ÷� ��й�ȣ ����
'Public Const gcOpenPW = "ACK"
'Public Const gcEditPW = "MEDI@CK"

'�ӽ÷� ��й�ȣ ����
Public Const pOpenPW = "ACK"
Public Const pEditPW = "MEDI@CK"

Public gsSiteNm As String       '�ŷ�ó����

'
'   ASTM Protocol CheckSum ���
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

