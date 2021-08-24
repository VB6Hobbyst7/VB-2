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
    JUMIN       As String
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
    IsResult    As Boolean
    PID         As String
    ChartNo     As String
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
    'for BC6200
    MsgCtrlID   As String
    TstType     As String
    
    
    Company     As String
    Product     As String
    DateTime    As String
    MsgType     As String
    MsgSeq      As String
    
    QryTime     As String
    QryFmtCd    As String
    QryPrio     As String
    QryID       As String
    QLR         As String   'Quality Limited Request
    SmplBarcode As String
    WSF         As String   'What Subject Field
    QRL         As String   'Query Result Level
    
    BDtTm       As String
    EDtTm       As String
    WDQ         As String   'Which Data/Time Qualifier
    WTQ         As String   'Which Data/Time Status Qualifier
    DSQ         As String   'Data/Time Selection Qualifier
    
    'for MSH
    MSHCorpName     As String
    MSHDeviceModel  As String
    MSHSysDateTime  As String
    MSHMessageType  As String
    MSHMessageID    As String
    MSHProduct      As String
    MSHHL7Version   As String
    MSHResultType   As String
    MSHChrEncoding  As String
    
                    'Corp.name          : MINDRAY
                    'Device Model       : BS-380
                    'System date/time   : 20130504083053
                    'Message Type       : QRY^Q02
                    'Message ID         : 1
                    'Product            : P
                    'HL7 Version        : 2.3.1
                    'Resut Type         : 0 (Sample) , 1 (Calib. Result)
                    'Character Encoding : ASCII

    QRDQryTime          As String
    QRDQryFormatCode    As String
    QRDQryPriority      As String
    QRDNum              As String
    QRDQLRequest        As String
    QRDSampleBarcode    As String
    QRDWSFilter         As String
    QRDQryResultLevel   As String
    
                    'Qry Time(2)                    : 20180611153634
                    'Qry Format Code(3)             : R
                    'Qry Priority(4)                : D
                    'Quantity Limited Request(8)    : RD
                    'Sample Barcode(9)              : 0019
                    'What Subject Filter(10)        : OTH
                    'Query Results Level(13)        : T

    QRFProduct              As String
    QRFWherStartDtTm        As String
    QRFWherEndDtTm          As String
    QRFWhichDtTmQualifier   As String
    QRFWhichStatusQualifier As String
    QRFDtTmSelecQualifier   As String
    
                    'Which Date/Time Qualifier          : RCT
                    'Which date/Time Status Qualifier   : COR
                    'Date/Time Selection Qualifier      : ALL

    SPCID    As String
    'OCNT    As Integer
    
End Type

Public mOrder As RecvData


'-- 전송할 오더정보
Type SendData
    Count           As Integer
    SndCnt          As Integer
    Data(1000)      As String
End Type

Public mSend As SendData



Type MSH_Field
    Field(20) As String
End Type

Public MSH As MSH_Field

Type QRD_Field
    Field(12) As String
End Type

Public QRD As QRD_Field

Type QRF_Field
    Field(9) As String
End Type

Public QRF As QRF_Field


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
    TESTCD   As String
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
    'for QC
    LabNab      As String
    TestQCCd    As String
    MType       As String
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



'''Public Sub Serial_Protocol()
'''
'''    SetRawData "[Rx]" & pBuffer
'''
'''    Select Case UCase(gHOSP.MACHNM)
'''
''''        Case "AFIAS6"
''''                Call Phase_Serial_AFIAS6
'''
''''        Case "VERSACELL"
''''                Call Phase_Serial_VERSACELL
'''
''''        Case "ADVIA1800-1", "ADVIA1800-2"
''''                Call Phase_Serial_ADVIA1800
'''
''''        Case "ADVIA2120-1", "ADVIA2120-2"
''''                Call Phase_Serial_ADVIA2120
'''
''''        Case "RAPIDLAB348"
''''                Call Phase_Serial_RAPIDLAB348
'''
''''        Case "RAPIDPOINT500"
''''                Call Phase_Serial_RAPIDPOINT500
'''
''''        Case "PFA200"
''''                Call Phase_Serial_PFA200
'''
''''        Case "ACLTOP"
''''                Call Phase_Serial_ACLTOP
'''
''''
''''        Case "VESCUBE"
''''                Call Phase_Serial_VESCUBE
'''
'''
''''        Case "CT500"
''''                Call Phase_Serial_CT500
'''
'''        Case Else
'''
'''    End Select
'''
'''
'''End Sub


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

Function GetPrevResult(pBarno As String, pEqipCode As String, pExamCode As String) As String
    Dim RS_L        As ADODB.Recordset
    Dim strPrevRslt As String
    
    GetPrevResult = ""
    
    If pBarno = "" Then
        Exit Function
    End If

    SQL = ""
    SQL = SQL & "SELECT TOP 1 EXAMDATE, RESULT                              " & vbCrLf
    SQL = SQL & "  FROM PATRESULT                                           " & vbCrLf
    SQL = SQL & " WHERE BARCODE     = '" & pBarno & "'                      " & vbCrLf
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
    strSndMsg = STX & strSndMsg & ETX '& GetChkSum(strSndMsg) & vbCr
    'strSndMsg = strSndMsg & vbCrLf
        
    SndMore = strSndMsg
    
End Function

Public Sub SetResultBufFree()
    
    With mResult
        .SpcmNo = ""
        .Seq = ""
        .PatNo = ""
        .BarNo = ""
        .RackNo = ""
        .TubePos = ""
        .MnmCd = ""
        .MnmNm = ""
        .MCnt = ""
        .RST = ""
        .SpcPos = ""
        .RsltDate = ""
        .RsltTime = ""
        .RsltSeq = ""
        .TESTCD = ""
        .Kind = ""
        .Rerun = ""
        .IntBase = ""
        .Result = ""
        .EqpCd = ""
        .RESODRSEQ = ""
        .RESSEQ = ""
        .RESSUBSEQ = ""
        .CARBAR_CMT = ""
        .MTBRIF_CMT = ""
        .CARBAR_CMTCD = ""
        .MTBRIF_CMTCD = ""
        .CMNTCD = ""
        .strNa = ""
        .strK = ""
        .strCl = ""
        .LabNab = ""
        .TestQCCd = ""
    End With
    
End Sub

Public Sub SetOrderBufFree()
    
    With mOrder
        .OrgBarNo = ""
        .BarNo = ""
        .Seq = ""
        .RackNo = ""
        .TubePos = ""
        .NoOrder = True
        .Order = ""
        .IsSending = False
        .SendCnt = 0
        .IsResult = False
        .PID = ""
        .ChartNo = ""
        .SPCCD = ""
        .SampleData = ""
        'for PLIS
        .WA = ""
        .AccSeq = 0
        'for ACLTOP
        .MsgID = ""
        .Sender = ""
        .Receiver = ""
        .Version = ""
        .PNAME = ""
        .Count = 0
        Erase mOrder.Items
        'for H7180
        .Func = ""
        .Function = ""
        'for LH780
        .BlkCnt = 0
        'for AU480
        .SmpType = ""
        'for BS240
        .BSMType = ""
        .BSMaker = ""
        .BSMchNm = ""
        .BSDtTm = ""
        .BSModel = ""
        .BSSTime = ""
        .BSETime = ""
        .BSQryId = ""
        .BSQRF = ""
        'for ACCESS2 배치오더
        .DestRow = 0
        'for 의사랑
        .OCNT = 0
        'for BC6200
        .MsgType = ""
        .MsgCtrlID = ""
        .TstType = ""
        'for MSH
        .MSHCorpName = ""
        .MSHDeviceModel = ""
        .MSHSysDateTime = ""
        .MSHMessageType = ""
        .MSHMessageID = ""
        .MSHProduct = ""
        .MSHHL7Version = ""
        .MSHResultType = ""
        .MSHChrEncoding = ""
        .QRDQryTime = ""
        .QRDQryFormatCode = ""
        .QRDQryPriority = ""
        .QRDNum = ""
        .QRDQLRequest = ""
        .QRDSampleBarcode = ""
        .QRDWSFilter = ""
        .QRDQryResultLevel = ""
        .QRFProduct = ""
        .QRFWherStartDtTm = ""
        .QRFWherEndDtTm = ""
        .QRFWhichDtTmQualifier = ""
        .QRFWhichStatusQualifier = ""
        .QRFDtTmSelecQualifier = ""
        .SPCID = ""
    End With
    
End Sub





