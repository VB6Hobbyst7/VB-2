Attribute VB_Name = "mscomm"
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
'*                                                              *
'*  SLBI_30F  = BIOMIC 결과받기                          *
'*                                                              *
'*  System    : 신촌세브란스병원 시스템                         *
'*  Subsystem : 임상병리 관리 시스템                            *
'*                                                              *
'*  Designed  : 1997-08-30                                      *
'*  Coded     : 1997-08-30 유은자                               *
'*  Modified  :                                                 *
'*                                                              *
'*  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *  *
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

Public Function p_funGet_Bn2BCC_2(ByVal strPara As String) As Integer

    Dim intIdx  As Integer
    Dim intBcc  As Integer
    
    intBcc = 0
    For intIdx = 1 To Len(strPara)
        intBcc = intBcc + Asc(Mid$(strPara, intIdx, 1))
    Next
    
    p_funGet_Bn2BCC_2 = 0
    If intBcc Mod 64 = 0 Then p_funGet_Bn2BCC_2 = 1

End Function

Public Function p_funGet_Bn2BCC_1(ByVal strPara As String) As String

    Dim intIdx  As Integer
    Dim intBcc  As Integer
    
    intBcc = 0
    For intIdx = 1 To Len(strPara)
        intBcc = intBcc + Asc(Mid$(strPara, intIdx, 1))
    Next
    
    intBcc = 64 - (intBcc Mod 64)
    
    If intBcc < 32 Then intBcc = intBcc + 64
    
    p_funGet_Bn2BCC_1 = Chr(intBcc)
        
End Function


