Attribute VB_Name = "modCommunication"
Option Explicit


Declare Function GetDefaultCommConfig Lib "kernel32" Alias "GetDefaultCommConfigA" (ByVal lpszName As String, lpCC As COMMCONFIG, lpdwSize As Long) As Long

'-- 환자정보
Type PatData
    BARCODE     As String
    ChartNo     As String
    PID         As String
    NAME        As String
    SEX         As String
    AGE         As String
End Type

Public mPatient As PatData

'-- 수신한 오더정보
Type RecvData
    OrgBarNo    As String
    BarNo       As String
    Seq         As String
    RackNo      As String
    TubePos     As String
    NoOrder     As Boolean
    Order       As String
    IsSending   As Boolean
    SendCnt     As Integer
    isresult    As Boolean
    PID         As String
    SPCCD       As String
    SampleData  As String
    'for PLIS
    WA          As String
    AccSeq      As Long
    'for ACLTOP
    MsgID       As String
    Sender      As String
    Receiver    As String
    Version     As String
    PNAME       As String
    Count       As Integer
    Items()     As String
    'for H7180
    Func        As String
    Function    As String
    'for LH780
    BlkCnt      As Integer
    'for AU480
    SmpType     As String
    
    'for BS240
    BSMType     As String
    BSMaker     As String
    BSMchNm     As String
    BSDtTm      As String
    BSModel     As String
    BSSTime     As String
    BSETime     As String
    BSQryId     As String
    BSQRF       As String
    'for ACCESS2 배치오더
    DestRow     As Integer
    'for 의사랑
    OCNT        As Integer
    'for E411 batch (응급)
    STAT        As String
    NextData    As String
    LastORDER   As Boolean
End Type

Public mOrder As RecvData

'-- 수신한 결과정보
Type IntfData
    SpcmNo   As String
    Seq      As String
    PatNo    As String
    BarNo    As String
    RackNo   As String
    TubePos  As String
    MnmCd    As String
    MnmNm    As String
    MCnt     As String
    RST      As String
    SpcPos   As String
    RsltDate As String
    RsltTime As String
    RsltSeq  As String
    TestCd   As String
    Kind     As String
    Rerun    As String
    IntBase  As String
    Result   As String
    EqpCd    As String
    '-- 비트워크리스트용
    RESODRSEQ   As String
    RESSEQ      As String
    RESSUBSEQ   As String
    CARBAR_CMT  As String
    MTBRIF_CMT  As String
    CARBAR_CMTCD  As String
    MTBRIF_CMTCD  As String
    CMNTCD      As String
    'for 9180
    strNa       As String
    strK        As String
    strCl       As String
End Type

Public mResult As IntfData

'for ADVIA1650
Public iPendingFlag    As Integer
Public iTotQueryFlag   As Integer
Public iTmpPendingFlag As Integer
Public iIdleFlag   As Integer
Public iOrderFlag  As Integer
Public iResultFlag As Integer
Public sRcvState   As String
Public sSndState   As String
Public sSndPacket()    As String
Public sQueryBarcd()   As String

'for ADVIA2120
Public Const mc_sSampleType    As String = "1"
Public Const mc_sPatInfo       As String = ""
Public Const mc_sSampInfo      As String = ""
Public Const mc_sSiteNm        As String = ""
Public Const mc_sRerunGbn      As String = ""
Public Const mc_bSerumIndex    As Integer = False
Public Const mc_sEqName        As String = ""
Public Const mc_bUseBarcode    As Boolean = False
Public Const mc_iPhase         As Integer = 1
Public Const mc_iSendPhase     As Integer = 1
Public Const mc_sTestMode      As String = "0"
Public Const mc_iFrameN        As Integer = 1
Public Const mc_sID            As String = ""
Public Const mc_sSeq           As String = ""
Public Const mc_sRack          As String = ""
Public Const mc_sPos           As String = ""
Public Const mc_iOrdCnt        As Integer = 0
Public Const mc_sTIFCd         As String = ""
Public Const mc_bPortOpen      As Boolean = False
Public Const mc_sOpenPW        As String = ""
Public Const mc_sEditPW        As String = ""
Public Const mc_bReserveEnd    As Boolean = False

'속성 변수:
Public mp_sSampleType          As String
Public mp_sPatInfo             As String
Public mp_sSampInfo            As String
Public mp_sSiteNm              As String
Public mp_sRerunGbn            As String
Public mp_bSerumIndex          As Boolean
Public mp_sEqName              As String
Public mp_bUseBarcode          As Boolean
Public mp_iPhase               As Integer
Public mp_iSendPhase           As Integer
Public mp_sTestMode            As String
Public mp_iFrameN              As Integer
Public mp_sID                  As String
Public mp_sSeq                 As String
Public mp_sRack                As String
Public mp_sPos                 As String
Public mp_iOrdCnt              As Integer
Public mp_sTIFCd               As String
Public mp_bPortOpen            As Boolean
Public mp_sOpenPW              As String
Public mp_sEditPW              As String
Public mp_bReserveEnd          As Boolean

Public Const mc_iMaxCnt     As Integer = 100
Public msMT                 As String
Public msTimerFlag          As String
Public msSndPacket          As String

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'for PFA-200, CT500
Public miLineNo             As Integer


'for RAPIDPOINT500
Public aMod                 As String
Public iIID                 As String
Public AckOn                As Boolean
Public Sample_Seq           As String

'for ACLTOP
Public mPNo                 As Integer
Public mOCnt                As Integer

'===== User Define
'인터페이스에서 사용
Public strFRcvBuffer   As String
Public strFWkBuf       As String
Public strFState       As String
Public blnFSend        As Boolean
Public blnFEndChk      As Boolean
Public blnFSTXChk      As Boolean
Public strFRstEnd      As String

Public strFRcvState    As String
Public strFSndState    As String
Public msAllBarCd   As String
Public maAllBarCd() As String
Public TimerFlag    As Integer
Public SavBuffer    As String
Public ii_SendCnt   As Integer
Public m_aTemp()    As String
Public miSendCnt    As Integer
Public msSendBuff   As String

'속성 변수:
Public m_p_sPatInfo As Variant
Public m_EqName As String
Public m_bUseBarcode As Boolean
Public m_iPhase As Integer
Public m_iSendPhase As Integer
Public m_sTestMode As String
Public m_iFrameN As Integer
Public m_p_sID As String
Public m_p_sSeq As String
Public m_p_sRack As String
Public m_p_sPos As String
Public m_p_iOrdCnt As Integer
Public m_p_sTIFCd As String
Public m_PortOpen As Boolean
Public m_OpenPW As String
Public m_EditPW As String
Public m_IFMode As String

'for IF
Public strOldBarno          As String   '수신한 바코드번호

'for CT500
Public mColor               As Boolean

'For E-170/Hitachi7600
Public bSTXChk      As Boolean
Public bEndChk      As Boolean
Public RstEnd       As String



Public Sub Serial_Protocol()

    SetRawData "[Rx]" & pBuffer
        
    Select Case UCase(gHOSP.MACHNM)
        
'        Case "AFIAS6"
'                Call Phase_Serial_AFIAS6
                
'        Case "VERSACELL"
'                Call Phase_Serial_VERSACELL
                
'        Case "ADVIA1800-1", "ADVIA1800-2"
'                Call Phase_Serial_ADVIA1800
                
'        Case "ADVIA2120-1", "ADVIA2120-2"
'                Call Phase_Serial_ADVIA2120
                
'        Case "RAPIDLAB348"
'                Call Phase_Serial_RAPIDLAB348
                
'        Case "RAPIDPOINT500"
'                Call Phase_Serial_RAPIDPOINT500
        
'        Case "PFA200"
'                Call Phase_Serial_PFA200
                
'        Case "ACLTOP"
'                Call Phase_Serial_ACLTOP
                
'
'        Case "VESCUBE"
'                Call Phase_Serial_VESCUBE
                
                
'        Case "CT500"
'                Call Phase_Serial_CT500
                
        Case Else
            
    End Select
    

End Sub


Public Function EnumSerPorts(port As Integer) As Long
    Dim cc As COMMCONFIG, ccsize As Long
    ccsize = LenB(cc)
    EnumSerPorts = GetDefaultCommConfig("COM" + Trim(Str(port)) + Chr(0), cc, ccsize)
End Function



'-----------------------------------------------------------------------------'
'   기능 : 해당 문자열의 CheckSum을 구함
'   인수 :
'       - pMsg : 문자열
'   반환 : CheckSum
'-----------------------------------------------------------------------------'
Public Function GetChkSum(ByVal pMsg As String) As String
    Dim lngChkSum   As Long
    Dim i           As Long

    For i = 1 To Len(pMsg)
        lngChkSum = (lngChkSum + Asc(Mid(pMsg, i, 1))) Mod 256
    Next

    If lngChkSum = 0 Then
        GetChkSum = "00"
    Else
        GetChkSum = Mid("0" & Hex(lngChkSum), Len(Hex(lngChkSum)), 2)
    End If
End Function

Public Function CheckSum_ADVIA2120(ByVal pMsg As String) As String
    Dim i%
    Dim sXOR$
    
    sXOR = ""
    sXOR = Mid(pMsg, 1, 1)
    
    For i = 2 To Len(pMsg)
        sXOR = Chr(Int(Asc(sXOR)) Xor Int(Asc(Mid(pMsg, i, 1))))
    Next
    
    If sXOR = Chr(3) Then
        sXOR = Chr(127)
    End If
    
    CheckSum_ADVIA2120 = Chr(2) & pMsg & sXOR & Chr(3)
    
End Function

Function GetPrevResult(pBarNo As String, pEqipCode As String, pExamCode As String) As String
    Dim RS_L        As ADODB.Recordset
    Dim strPrevRslt As String
    
    GetPrevResult = ""
    
    If pBarNo = "" Then
        Exit Function
    End If

    SQL = ""
    SQL = SQL & "SELECT TOP 1 EXAMDATE, RESULT                              " & vbCrLf
    SQL = SQL & "  FROM PATRESULT                                           " & vbCrLf
    SQL = SQL & " WHERE BARCODE     = '" & pBarNo & "'                      " & vbCrLf
    SQL = SQL & "   AND EQUIPCODE   = '" & pEqipCode & "'                   " & vbCrLf
    SQL = SQL & "   AND EXAMCODE    = '" & pExamCode & "'                   " & vbCrLf
    SQL = SQL & "   AND EXAMDATE    < '" & Format(Now, "yyyy-mm-dd") & "'   " & vbCrLf
    SQL = SQL & " ORDER By EXAMDATE DESC "
    
    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
        strPrevRslt = Trim(RS_L.Fields("RESULT")) & ""
    End If
    
    RS_L.Close
    
    GetPrevResult = strPrevRslt
    
End Function

Public Function SndMore() As String
    Dim strSndMsg As String
    
    SndMore = ""
    
    strSndMsg = ">"
    'strSndMsg = STX & strSndMsg & ETX ' & GetChkSum(strSndMsg) & vbCr
    strSndMsg = STX & strSndMsg & ETX & GetChkSum(strSndMsg) & vbCr
    'strSndMsg = strSndMsg & vbCrLf
        
    SndMore = strSndMsg
    
End Function




