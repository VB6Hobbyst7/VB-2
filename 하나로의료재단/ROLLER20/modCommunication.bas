Attribute VB_Name = "modCommunication"
Option Explicit

Public pBuffer As Variant

'-- 수신한 오더정보
Type RecvData
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
    'for ACLTOP
    MsgID       As String
    Sender      As String
    Receiver    As String
    Version     As String
    PName       As String
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
    RsltSeq  As String
    TestCd   As String
    Kind     As String
    Rerun    As String
    IntBase  As String
    RESULT   As String
    EqpCd    As String
    SEX      As String
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

Public Sub TCP_Protocol()

    Select Case UCase(gHOSP.MACHNM)
        Case "BA400"
                Call Phase_TCP_BA400
        Case ""
        
    End Select
    
End Sub

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


'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
'                                 2. 장비수신정보 화면표시,
'                                 3. 처방코드 가져오기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Public Sub SetPatInfo(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If IsNumeric(pBarno) And IsNumeric(Trim(GetText(frmMain.spdOrder, i, colBARCODE))) Then
                    If Val(Trim(GetText(frmMain.spdOrder, i, colBARCODE))) = Val(pBarno) Then
                        If Trim(GetText(frmMain.spdOrder, i, colSTATE)) = "" Or InStr(GetText(frmMain.spdOrder, i, colSTATE), "오더") > 0 Then
                            intRow = i
                            Exit For
                        End If
                    End If
                End If
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" And GetText(frmMain.spdOrder, i, colSTATE) = "" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mResult.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
            End Select
        End If
        
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        Call SetText(.spdOrder, "1", intRow, colCHECKBOX)
        
        '-- 장비결과인덱스 화면표시
        Call SetText(.spdOrder, mResult.RsltSeq, intRow, colSAVESEQ)
        Call SetText(.spdOrder, mResult.RsltDate, intRow, colEXAMDATE)
        
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mResult.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mResult.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mResult.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mResult.TubePos, intRow, colPOSNO)
    
        '-- 환자정보 표시
        'Call vasActiveCell(.spdOrder, intRow, colBARCODE)
        
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
    
        '-- 검사자 정보 가져오기
        If Len(pBarno) = 9 Then
            Call GetSampleInfo(intRow, .spdOrder)
        ElseIf Len(pBarno) = 12 Then
            Call GetSampleInfo_SCL(intRow, .spdOrder)
        End If
        
        .spdOrder.RowHeight(-1) = 15
    
    End With
    
    '-- 현재 Row
    gRow = intRow
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Public Sub SendOrder_E411()
    
    
End Sub


Public Sub Phase_TCP_BA400()
 
End Sub
    

'Public Sub Phase_Serial_E411()
'
'
'End Sub
'
'Public Sub Phase_Serial_AU400()
'    Dim Buffer      As Variant
'    Dim BufChar     As String
'    Dim lngBufLen   As Long
'    Dim i           As Long
'
'    lngBufLen = Len(pBuffer)
'
'    For i = 1 To lngBufLen
'        BufChar = Mid$(pBuffer, i, 1)
'        Select Case BufChar
'            Case STX
'                intBufCnt = 1
'                Erase strRecvData
'                ReDim Preserve strRecvData(intBufCnt)
'            Case ETB
'            Case ETX
'                Call SerialRcvData_AU400
'            Case Else
'                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'        End Select
'    Next i
'
'End Sub

'Public Sub Phase_Serial_AU680()
'    Dim Buffer      As Variant
'    Dim BufChar     As String
'    Dim lngBufLen   As Long
'    Dim i           As Long
'
'    lngBufLen = Len(pBuffer)
'
'    For i = 1 To lngBufLen
'        BufChar = Mid$(pBuffer, i, 1)
'        Select Case BufChar
'            Case STX
'                intBufCnt = 1
'                Erase strRecvData
'                ReDim Preserve strRecvData(intBufCnt)
'            Case ETB
'            Case ETX
'                Call SerialRcvData_AU680
'            Case Else
'                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'        End Select
'    Next i
'
'End Sub

'Public Sub Phase_Serial_AFIAS6()
'    Dim Buffer      As Variant
'    Dim BufChar     As String
'    Dim lngBufLen   As Long
'    Dim i           As Long
'
'    lngBufLen = Len(pBuffer)
'
'    For i = 1 To lngBufLen
'        BufChar = Mid$(pBuffer, i, 1)
'        Select Case BufChar
'            Case "$" 'SOH
'                intBufCnt = 1
'                Erase strRecvData
'                ReDim Preserve strRecvData(intBufCnt)
'            Case vbCr
'                Call SerialRcvData_AFIAS6
'            Case Else
'                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'        End Select
'    Next i
'
'End Sub


'Public Sub Phase_Serial_AU480()
'    Dim Buffer      As Variant
'    Dim BufChar     As String
'    Dim lngBufLen   As Long
'    Dim i           As Long
'
'    lngBufLen = Len(pBuffer)
'
'    For i = 1 To lngBufLen
'        BufChar = Mid$(Buffer, i, 1)
'        Select Case BufChar
'            Case STX
'                intBufCnt = 1
'                Erase strRecvData
'                ReDim Preserve strRecvData(intBufCnt)
'            Case ETB
'            Case ETX
'                Call SerialRcvData_AU480
'            Case Else
'                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'        End Select
'    Next i
'
'End Sub

'Public Sub Phase_Serial_XN1000()
'    Dim Buffer      As Variant
'    Dim BufChar     As String
'    Dim lngBufLen   As Long
'    Dim i           As Long
'
'    lngBufLen = Len(pBuffer)
'    With frmMain
'        For i = 1 To lngBufLen
'            BufChar = Mid$(pBuffer, i, 1)
'            Select Case intPhase
'                Case 1      '## Estabilshment Phase
'                    Select Case BufChar
'                        Case ENQ
'                            intBufCnt = 1
'                            Erase strRecvData
'                            ReDim Preserve strRecvData(intBufCnt)
'                            intPhase = 2
'                            .comEqp.Output = ACK
'                            DoEvents
'                            SetRawData "[Tx]" & ACK
'                        Case ACK
'                            If strState = "Q" Then Call SendOrder
'
'                    End Select
'                Case 2      '## Transfer Phase
'                    Select Case BufChar
'                        Case ENQ
'                            Erase strRecvData
'                            .comEqp.Output = ACK
'                            DoEvents
'                            SetRawData "[Tx]" & ACK
'                        Case STX
'                            intBufCnt = 1
'                            Erase strRecvData
'                            ReDim Preserve strRecvData(intBufCnt)
'                        Case ETB
'                            blnIsETB = True
'                            intPhase = 3
'                        Case ETX
'                            intBufCnt = intBufCnt + 1
'                            ReDim Preserve strRecvData(intBufCnt)
'                            intPhase = 3
'                        Case vbCr, vbLf
'                        Case EOT
'                            intPhase = 1
'                        Case Else
'                            If blnIsETB = False Then
'                                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'                            Else
'                                blnIsETB = False
'                            End If
'                    End Select
'                Case 3      '## Transfer Phase
'                    Select Case BufChar
'                        Case vbCr
'                        Case vbLf
'                            intPhase = 4
'                            .comEqp.Output = ACK
'                            DoEvents
'                            SetRawData "[Tx]" & ACK
'                    End Select
'                Case 4      '## Termination Phase
'                    Select Case BufChar
'                        Case STX
'                            intPhase = 2
'                        Case EOT
'                            Call SerialRcvData_XN1000
'                            If strState = "Q" Then
'                                intSndPhase = 1
'                                intFrameNo = 1
'                                .comEqp.Output = ENQ
'                                DoEvents
'                                SetRawData "[Tx]" & ENQ
'                            End If
'
'                            intPhase = 1
'                    End Select
'            End Select
'        Next i
'    End With
'
'End Sub


'Public Sub SerialRcvData_XN1000()
'    Dim RS_L            As ADODB.Recordset
'    Dim strRcvBuf       As String   '수신한 Data
'    Dim strRcvBufOrd    As String
'    Dim strRcvBufRst    As String
'    Dim strType         As String   '수신한 Record Type
'    Dim strBarno        As String   '수신한 바코드번호
'    Dim strSeq          As String   '수신한 Sequence
'    Dim strRackNo       As String   '수신한 Rack Or Disk No
'    Dim strTubePos      As String   '수신한 Tube Position
'    Dim strIntBase      As String   '수신한 장비기준 검사명
'    Dim strMachResult   As String   '수신한 장비결과
'    Dim strResult       As String   '수신한 결과(정성)
'    Dim strIntResult    As String   '수신한 결과(정량)
'    Dim strQCResult     As String   '수신한 결과(QC)
'    Dim strFlag         As String   '수신한 Abnormal Flag
'    Dim strComm         As String   '수신한 Comment
'
'    Dim lsOrderCode     As String   '처방코드
'    Dim lsTestCode      As String   '검사코드
'    Dim lsTestName      As String   '검사명
'    Dim lsSeqNo         As String   '로컬DB 검사Seq
'
'    Dim lsRstRow        As String   '결과스프레드 현재 Row
'    Dim intCnt          As Integer  '통신 Frame 갯수
'    Dim intCol          As Integer  '결과컬럼 갯수
'    Dim strJudge        As String   '결과판정
'    Dim Res             As Integer
'
'
'    With frmMain
'        For intCnt = 1 To UBound(strRecvData)
'            strRcvBuf = strRecvData(intCnt)
'
'            '-- 테스트용 -----------------
'            If .fraCommTest.Visible = False Then
'                Call SetSQLData("RCV", strRcvBuf, "A")
'            End If
'            '-- 테스트용 -----------------
'
'            strType = Mid$(strRcvBuf, 1, 1)
'            If IsNumeric(strType) Then
'                strType = Mid$(strRcvBuf, 2, 1)
'            End If
'
'            Select Case strType
'                Case "H"
'                Case "Q"    '## Inquiry Order
'                        strBarno = Trim(Mid(strRcvBuf, 14, 26))
'                        strSeq = Mid(strRcvBuf, 9, 5)
'                        strRackNo = Mid(strRcvBuf, 3, 4)
'                        strTubePos = Mid(strRcvBuf, 7, 2)
'
'                        With mOrder
'                            .BarNo = strBarno
'                            .Seq = strSeq
'                            .RackNo = strRackNo
'                            .TubePos = strTubePos
'                        End With
'
'                        If strBarno = "" Then
'                            strBarno = "NoOrder_" & Trim(strSeq)
'                        End If
'
'                        Call GetOrder(strBarno, gHOSP.RSTTYPE)
'
'                        strState = "Q"
'
'                Case "P"
'
'                Case "O"
'                    '4O|1||1^6^          201404240002^B|^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NEUT#\^^^^LYMPH#\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH%\^^^^MONO%\^^^^EC|1||
'
'                    strRcvBufOrd = Trim$(mGetP(strRcvBuf, 4, "|"))
'                    strBarno = Trim$(mGetP(strRcvBufOrd, 3, "^"))
'                    strSeq = ""
'                    strRackNo = Trim$(mGetP(strRcvBufOrd, 1, "^"))
'                    strTubePos = Trim$(mGetP(strRcvBufOrd, 2, "^"))
'
'                    With mResult
'                        .BarNo = strBarno
'                        .SpcPos = strSeq
'                        .RackNo = strRackNo
'                        .TubePos = strTubePos
'                        .RsltDate = Format(Now, "yyyymmddhhmmss")
'                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'                    End With
'
'                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'                    If gRow <= 0 Then
'                        Exit Sub
'                    End If
'
'                    strState = "O"
'                    .spdResult.MaxRows = 0
'
'                Case "R"
'                    strRcvBufRst = Trim(mGetP(strRcvBuf, 3, "|"))
'                    strIntBase = Trim$(mGetP(strRcvBufRst, 5, "^"))
'                    strResult = Trim(mGetP(strRcvBuf, 4, "|"))
'
'                    If strIntBase <> "" And strResult <> "" Then
'                        If gPatOrdCd <> "" Then
'                            SQL = ""
'                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                            SQL = SQL & "  FROM EQPMASTER" & vbCr
'                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.HOSPCD & "' " & vbCr
'                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
'                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
'
'                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                lsTestCode = Trim(RS_L.Fields("TESTCODE"))
'                                lsTestName = Trim(RS_L.Fields("TESTNAME"))
'                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                '-- 결과Row 추가
'                                lsRstRow = .spdResult.DataRowCnt + 1
'                                If .spdResult.MaxRows < lsRstRow Then
'                                    .spdResult.MaxRows = lsRstRow
'                                End If
'
'                                '소수점 처리, 결과 형태 처리
'                                strMachResult = strResult
'                                strResult = SetResult(strResult, strIntBase)
'                                strJudge = SetJudge(strResult, strIntBase)
'
'                                '진행상태 표시("결과")
'                                SetText .spdOrder, "결과", gRow, colSTATE
'
'                                '결과값 표시
'                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                        SetText .spdOrder, strResult, gRow, intCol
'                                        Exit For
'                                    End If
'                                Next
'
'                                '-- 결과 List
'                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
'                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
'                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
'                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
'                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
'                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
'                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
'                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
'                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
'                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
'
'                                '-- 로컬 저장
'                                SetLocalDB gRow, lsRstRow, "1", ""
'
'                                strState = "R"
'
'                                '-- 결과Count
'                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                                    SetText .spdOrder, "1", gRow, colRCNT
'                                Else
'                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                                End If
'
'                            End If
'                        Else
'                            SQL = ""
'                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                            SQL = SQL & "  FROM EQPMASTER" & vbCr
'                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.HOSPCD & "' " & vbCr
'                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
'
'                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                                lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                '-- 결과Row 추가
'                                lsRstRow = .spdResult.DataRowCnt + 1
'                                If .spdResult.MaxRows < lsRstRow Then
'                                    .spdResult.MaxRows = lsRstRow
'                                End If
'
'                                '소수점 처리, 결과 형태 처리
'                                strMachResult = strResult
'                                strResult = SetResult(strResult, strIntBase)
'                                strJudge = SetJudge(strResult, strIntBase)
'
'                                '진행상태 표시("결과")
'                                SetText .spdOrder, "결과", gRow, colSTATE
'
'                                '결과값 표시
'                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                        SetText .spdOrder, strResult, gRow, intCol
'                                        Exit For
'                                    End If
'                                Next
'
'                                '-- 결과 List
'                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
'                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
'                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
'                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
'                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
'                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
'                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
'                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
'                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
'
'                                '-- 로컬 저장
'                                SetLocalDB gRow, lsRstRow, "1", ""
'
'                                If strState <> "R" Then
'                                    strState = ""
'                                End If
'
'                                '-- 결과Count
'                                If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                                    SetText .spdOrder, "1", gRow, colRCNT
'                                Else
'                                    SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                                End If
'                            End If
'
'                        End If
'
'                    End If
'
'                    .spdResult.RowHeight(-1) = 14
'
'                Case "L"
'                    '## DB에 결과저장
'                    If .optTrans(0).Value = True And strState = "R" Then
'                        Res = SaveTransData(gRow)
'
'                        If Res = -1 Then
'                            '-- 저장 실패
''                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                            SetText .spdOrder, "Failed", gRow, colSTATE
'                        Else
'                            '-- 저장 성공
''                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                            SetText .spdOrder, "저장완료", gRow, colSTATE
'                            SetText .spdOrder, "0", gRow, colCHECKBOX
'
'                                  SQL = "Update PATRESULT Set " & vbCrLf
'                            SQL = SQL & " sendflag = '2' " & vbCrLf
'                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
'                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
'
''                            Res = SendQuery(gLocal, SQL)
'                            If Res = -1 Then
''                                SaveQuery SQL
'                                Exit Sub
'                            End If
'                        End If
'                        strState = ""
'                    End If
'
'            End Select
'        Next
'    End With
'
'End Sub


'Public Sub SerialRcvData_AU480()
'
'
'End Sub

'Public Sub SerialRcvData_AU400()
'    Dim RS_L            As ADODB.Recordset
'    Dim strRcvBuf       As String   '수신한 Data
'    Dim strType         As String   '수신한 Record Type
'    Dim strBarno        As String   '수신한 바코드번호
'    Dim strSeq          As String   '수신한 Sequence
'    Dim strRackNo       As String   '수신한 Rack Or Disk No
'    Dim strTubePos      As String   '수신한 Tube Position
'    Dim strIntBase      As String   '수신한 장비기준 검사명
'    Dim strMachResult   As String   '수신한 장비결과
'    Dim strResult       As String   '수신한 결과(정성)
'    Dim strIntResult    As String   '수신한 결과(정량)
'    Dim strQCResult     As String   '수신한 결과(QC)
'    Dim strFlag         As String   '수신한 Abnormal Flag
'    Dim strComm         As String   '수신한 Comment
'
'    Dim lsOrderCode     As String   '처방코드
'    Dim lsTestCode      As String   '검사코드
'    Dim lsTestName      As String   '검사명
'    Dim lsSeqNo         As String   '로컬DB 검사Seq
'
'    Dim lsRstRow        As String   '결과스프레드 현재 Row
'    Dim intCnt          As Integer  '통신 Frame 갯수
'    Dim intCol          As Integer  '결과컬럼 갯수
'    Dim strJudge        As String   '결과판정
'    Dim Res             As Integer
'
'    Dim strTmp          As String
'
'    With frmMain
'        For intCnt = 1 To UBound(strRecvData)
'            strRcvBuf = strRecvData(intCnt)
'
'            '-- 테스트용 -----------------
'            If .fraCommTest.Visible = False Then
'                Call SetSQLData("RCV", strRcvBuf, "A")
'            End If
'            '-- 테스트용 -----------------
'
'            strRcvBuf = strRecvData(intCnt)
'            strType = Mid$(strRcvBuf, 1, 2)
'
'            Select Case strType
'                Case "R "    '## Inquiry Order
'                        'R 000101 0001                1608270009
'                        strBarno = Trim(Mid(strRcvBuf, 14, 26))
'                        strSeq = Mid(strRcvBuf, 9, 5)
'                        strRackNo = Mid(strRcvBuf, 3, 4)
'                        strTubePos = Mid(strRcvBuf, 7, 2)
'
'                        With mOrder
'                            .BarNo = strBarno
'                            .Seq = strSeq
'                            .RackNo = strRackNo
'                            .TubePos = strTubePos
'                        End With
'
'                        If strBarno = "" Then
'                            strBarno = "NoOrder_" & Trim(strSeq)
'                        End If
'
'                        Call GetOrder(strBarno, gHOSP.RSTTYPE)
'
'                        strState = "Q"
'
'                Case "D "    '## Result
'                        'D 000101 0001                1608270009    E001   9.3  002   5.8  
'                        strBarno = Trim$(Mid$(strRcvBuf, 14, 10))
'                        strRackNo = Mid(strRcvBuf, 3, 4)
'                        strTubePos = Mid(strRcvBuf, 7, 2)
'                        strSeq = Trim(Mid(strRcvBuf, 9, 5))
'
'                        With mResult
'                            .BarNo = strBarno
'                            .SpcPos = strSeq
'                            .Seq = strSeq
'                            .RackNo = strRackNo
'                            .TubePos = strTubePos
'                            .RsltDate = Format(Now, "yyyymmddhhmmss")
'                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'                        End With
'
'                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'                        If gRow <= 0 Then
'                            Exit Sub
'                        End If
'
'                        strTmp = Mid$(strRcvBuf, 29)
'
'                        Do While Len(strTmp) >= 11
'                            strIntBase = Mid$(strTmp, 2, 2)
'                            strResult = Mid$(strTmp, 4, 6)
'                            strComm = Mid$(strTmp, 10, 1)
'
'                            If strIntBase <> "" And strResult <> "" Then
'                                If gPatOrdCd <> "" Then
'                                    SQL = ""
'                                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                                    SQL = SQL & "  FROM EQPMASTER" & vbCr
'                                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
'                                    SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
'
'                                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                        lsTestCode = Trim(RS_L.Fields("TESTCODE"))
'                                        lsTestName = Trim(RS_L.Fields("TESTNAME"))
'                                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                        '-- 결과Row 추가
'                                        lsRstRow = .spdResult.DataRowCnt + 1
'                                        If .spdResult.MaxRows < lsRstRow Then
'                                            .spdResult.MaxRows = lsRstRow
'                                        End If
'
'                                        '소수점 처리, 결과 형태 처리
'                                        strMachResult = strResult
'                                        strResult = SetResult(strResult, strIntBase)
'                                        strJudge = SetJudge(strResult, strIntBase)
'
'                                        '진행상태 표시("결과")
'                                        SetText .spdOrder, "결과", gRow, colSTATE
'
'                                        '결과값 표시
'                                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                                SetText .spdOrder, strResult, gRow, intCol
'                                                Exit For
'                                            End If
'                                        Next
'
'                                        '-- 결과 List
'                                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
'                                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
'                                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
'                                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
'                                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
'                                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
'                                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
'                                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
'                                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
'                                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
'
'                                        '-- 로컬 저장
'                                        SetLocalDB gRow, lsRstRow, "1", ""
'
'                                        strState = "R"
'
'                                        '-- 결과Count
'                                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                                            SetText .spdOrder, "1", gRow, colRCNT
'                                        Else
'                                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                                        End If
'
'                                    End If
'                                Else
'                                    SQL = ""
'                                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                                    SQL = SQL & "  FROM EQPMASTER" & vbCr
'                                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
'
'                                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                        lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                                        lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                        '-- 결과Row 추가
'                                        lsRstRow = .spdResult.DataRowCnt + 1
'                                        If .spdResult.MaxRows < lsRstRow Then
'                                            .spdResult.MaxRows = lsRstRow
'                                        End If
'
'                                        '소수점 처리, 결과 형태 처리
'                                        strMachResult = strResult
'                                        strResult = SetResult(strResult, strIntBase)
'                                        strJudge = SetJudge(strResult, strIntBase)
'
'                                        '진행상태 표시("결과")
'                                        SetText .spdOrder, "결과", gRow, colSTATE
'
'                                        '결과값 표시
'                                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                                SetText .spdOrder, strResult, gRow, intCol
'                                                Exit For
'                                            End If
'                                        Next
'
'                                        '-- 결과 List
'                                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
'                                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
'                                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
'                                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
'                                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
'                                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
'                                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
'                                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
'                                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
'
'                                        '-- 로컬 저장
'                                        SetLocalDB gRow, lsRstRow, "1", ""
'
'                                        If strState <> "R" Then
'                                            strState = ""
'                                        End If
'
'                                        '-- 결과Count
'                                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                                            SetText .spdOrder, "1", gRow, colRCNT
'                                        Else
'                                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                                        End If
'                                    End If
'
'                                End If
'
'                            End If
'                            strTmp = Mid$(strTmp, 12)
'                        Loop
'
'                        .spdResult.RowHeight(-1) = 14
'
'                        '## DB에 결과저장
'                        If .optTrans(0).Value = True And strState = "R" Then
'                            Res = SaveTransData(gRow)
'
'                            If Res = -1 Then
'                                '-- 저장 실패
'    '                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                                SetText .spdOrder, "Failed", gRow, colSTATE
'                            Else
'                                '-- 저장 성공
'    '                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                                SetText .spdOrder, "저장완료", gRow, colSTATE
'                                SetText .spdOrder, "0", gRow, colCHECKBOX
'
'                                      SQL = "Update PATRESULT Set " & vbCrLf
'                                SQL = SQL & " sendflag = '2' " & vbCrLf
'                                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
'                                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'                                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'                                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
'
'    '                            Res = SendQuery(gLocal, SQL)
'                                If Res = -1 Then
'    '                                SaveQuery SQL
'                                    Exit Sub
'                                End If
'                            End If
'                            strState = ""
'                        End If
'
'            End Select
'        Next
'    End With
'
'End Sub
'
'Public Sub SerialRcvData_AU680()
'    Dim RS_L            As ADODB.Recordset
'    Dim strRcvBuf       As String   '수신한 Data
'    Dim strType         As String   '수신한 Record Type
'    Dim strBarno        As String   '수신한 바코드번호
'    Dim strSeq          As String   '수신한 Sequence
'    Dim strRackNo       As String   '수신한 Rack Or Disk No
'    Dim strTubePos      As String   '수신한 Tube Position
'    Dim strIntBase      As String   '수신한 장비기준 검사명
'    Dim strMachResult   As String   '수신한 장비결과
'    Dim strResult       As String   '수신한 결과(정성)
'    Dim strIntResult    As String   '수신한 결과(정량)
'    Dim strQCResult     As String   '수신한 결과(QC)
'    Dim strFlag         As String   '수신한 Abnormal Flag
'    Dim strComm         As String   '수신한 Comment
'
'    Dim lsOrderCode     As String   '처방코드
'    Dim lsTestCode      As String   '검사코드
'    Dim lsTestName      As String   '검사명
'    Dim lsSeqNo         As String   '로컬DB 검사Seq
'
'    Dim lsRstRow        As String   '결과스프레드 현재 Row
'    Dim intCnt          As Integer  '통신 Frame 갯수
'    Dim intCol          As Integer  '결과컬럼 갯수
'    Dim strJudge        As String   '결과판정
'    Dim Res             As Integer
'
'    Dim strTmp          As String
'
'    With frmMain
'        For intCnt = 1 To UBound(strRecvData)
'            strRcvBuf = strRecvData(intCnt)
'
'            '-- 테스트용 -----------------
'            If .fraCommTest.Visible = False Then
'                Call SetSQLData("RCV", strRcvBuf, "A")
'            End If
'            '-- 테스트용 -----------------
'
'            strRcvBuf = strRecvData(intCnt)
'            strType = Mid$(strRcvBuf, 1, 2)
'
'            Select Case strType
'                Case "R "    '## Inquiry Order
'                        'R 000101 0001                1608270009
'                        strBarno = Trim(Mid(strRcvBuf, 14, 20))
'                        strSeq = Mid(strRcvBuf, 9, 5)
'                        strRackNo = Mid(strRcvBuf, 3, 4)
'                        strTubePos = Mid(strRcvBuf, 7, 2)
'
'                        With mOrder
'                            .BarNo = strBarno
'                            .Seq = strSeq
'                            .RackNo = strRackNo
'                            .TubePos = strTubePos
'                        End With
'
'                        If strBarno = "" Then
'                            strBarno = "NoOrder_" & Trim(strSeq)
'                        End If
'
'                        Call GetOrder(strBarno, gHOSP.RSTTYPE)
'
'                        strState = "Q"
'
'                Case "D "    '## Result
'                        '1234567890123456789012345678901234567890
'                        'D 000501 0001            02001035    E001  2.60Pr002  75.1Pr003    98r 004   7.1r 005   3.4Nr007    80Pr008    75Pr009    61r 011  0.68r 097   132Nr098   4.3r 099   101r
'                        'D 000101 0001                1608270009    E001   9.3  002   5.8  
'                        strBarno = Trim$(Mid$(strRcvBuf, 14, 20))
'                        strRackNo = Mid(strRcvBuf, 3, 4)
'                        strTubePos = Mid(strRcvBuf, 7, 2)
'                        strSeq = Trim(Mid(strRcvBuf, 10, 4))
'
'                        With mResult
'                            .BarNo = strBarno
'                            .SpcPos = strSeq
'                            .Seq = strSeq
'                            .RackNo = strRackNo
'                            .TubePos = strTubePos
'                            .RsltDate = Format(Now, "yyyymmddhhmmss")
'                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'                        End With
'
'                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'                        If gRow <= 0 Then
'                            Exit Sub
'                        End If
'
'                        'strTmp = Mid$(strRcvBuf, 29)
'                        strTmp = Mid$(strRcvBuf, 39)
'
'                        'Do While Len(strTmp) >= 11
'                        Do While Len(strTmp) >= 10
'                            strIntBase = Mid$(strTmp, 2, 2)
'                            strResult = Mid$(strTmp, 4, 6)
'                            strComm = Mid$(strTmp, 10, 1)
'
'                            If strIntBase <> "" And strResult <> "" Then
'                                If gPatOrdCd <> "" Then
'                                    SQL = ""
'                                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                                    SQL = SQL & "  FROM EQPMASTER" & vbCr
'                                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
'                                    SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
'
'                                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                        lsTestCode = Trim(RS_L.Fields("TESTCODE"))
'                                        lsTestName = Trim(RS_L.Fields("TESTNAME"))
'                                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                        '-- 결과Row 추가
'                                        lsRstRow = .spdResult.DataRowCnt + 1
'                                        If .spdResult.MaxRows < lsRstRow Then
'                                            .spdResult.MaxRows = lsRstRow
'                                        End If
'
'                                        '소수점 처리, 결과 형태 처리
'                                        strMachResult = strResult
'                                        strResult = SetResult(strResult, strIntBase)
'                                        strJudge = SetJudge(strResult, strIntBase)
'
'                                        '진행상태 표시("결과")
'                                        SetText .spdOrder, "결과", gRow, colSTATE
'
'                                        '결과값 표시
'                                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                                SetText .spdOrder, strResult, gRow, intCol
'                                                Exit For
'                                            End If
'                                        Next
'
'                                        '-- 결과 List
'                                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
'                                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
'                                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
'                                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
'                                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
'                                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
'                                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
'                                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
'                                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
'
'                                        '-- 로컬 저장
'                                        SetLocalDB gRow, lsRstRow, "1", ""
'
'                                        strState = "R"
'
'                                        '-- 결과Count
'                                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                                            SetText .spdOrder, "1", gRow, colRCNT
'                                        Else
'                                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                                        End If
'
'                                    End If
'                                Else
'                                    SQL = ""
'                                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                                    SQL = SQL & "  FROM EQPMASTER" & vbCr
'                                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
'
'                                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                        lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                                        lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                        '-- 결과Row 추가
'                                        lsRstRow = .spdResult.DataRowCnt + 1
'                                        If .spdResult.MaxRows < lsRstRow Then
'                                            .spdResult.MaxRows = lsRstRow
'                                        End If
'
'                                        '소수점 처리, 결과 형태 처리
'                                        strMachResult = strResult
'                                        strResult = SetResult(strResult, strIntBase)
'                                        strJudge = SetJudge(strResult, strIntBase)
'
'                                        '진행상태 표시("결과")
'                                        SetText .spdOrder, "결과", gRow, colSTATE
'
'                                        '결과값 표시
'                                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                                SetText .spdOrder, strResult, gRow, intCol
'                                                Exit For
'                                            End If
'                                        Next
'
'                                        '-- 결과 List
'                                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
'                                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
'                                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
'                                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
'                                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
'                                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
'                                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
'                                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
'                                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
'
'                                        '-- 로컬 저장
'                                        SetLocalDB gRow, lsRstRow, "1", ""
'
'                                        If strState <> "R" Then
'                                            strState = ""
'                                        End If
'
'                                        '-- 결과Count
'                                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                                            SetText .spdOrder, "1", gRow, colRCNT
'                                        Else
'                                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                                        End If
'                                    End If
'
'                                End If
'
'                            End If
'                            strTmp = Mid$(strTmp, 12)
'                        Loop
'
'                        .spdResult.RowHeight(-1) = 14
'
'                        '## DB에 결과저장
'                        If .optTrans(0).Value = True And strState = "R" Then
'                            Res = SaveTransData(gRow)
'
'                            If Res = -1 Then
'                                '-- 저장 실패
'    '                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                                SetText .spdOrder, "Failed", gRow, colSTATE
'                            Else
'                                '-- 저장 성공
'    '                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                                SetText .spdOrder, "저장완료", gRow, colSTATE
'                                SetText .spdOrder, "0", gRow, colCHECKBOX
'
'                                      SQL = "Update PATRESULT Set " & vbCrLf
'                                SQL = SQL & " sendflag = '2' " & vbCrLf
'                                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
'                                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'                                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'                                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
'
'    '                            Res = SendQuery(gLocal, SQL)
'                                If Res = -1 Then
'    '                                SaveQuery SQL
'                                    Exit Sub
'                                End If
'                            End If
'                            strState = ""
'                        End If
'
'            End Select
'        Next
'    End With
'
'End Sub

'Public Sub SerialRcvData_AFIAS6()
'    Dim RS_L            As ADODB.Recordset
'    Dim strRcvBuf       As String   '수신한 Data
'    Dim strType         As String   '수신한 Record Type
'    Dim strBarno        As String   '수신한 바코드번호
'    Dim strSeq          As String   '수신한 Sequence
'    Dim strRackNo       As String   '수신한 Rack Or Disk No
'    Dim strTubePos      As String   '수신한 Tube Position
'    Dim strIntBase      As String   '수신한 장비기준 검사명
'    Dim strMachResult   As String   '수신한 장비결과
'    Dim strResult       As String   '수신한 결과(정성)
'    Dim strIntResult    As String   '수신한 결과(정량)
'    Dim strQCResult     As String   '수신한 결과(QC)
'    Dim strFlag         As String   '수신한 Abnormal Flag
'    Dim strComm         As String   '수신한 Comment
'
'    Dim lsOrderCode     As String   '처방코드
'    Dim lsTestCode      As String   '검사코드
'    Dim lsTestName      As String   '검사명
'    Dim lsSeqNo         As String   '로컬DB 검사Seq
'
'    Dim lsRstRow        As String   '결과스프레드 현재 Row
'    Dim intCnt          As Integer  '통신 Frame 갯수
'    Dim intCol          As Integer  '결과컬럼 갯수
'    Dim strJudge        As String   '결과판정
'    Dim Res             As Integer
'
'    Dim strTmp          As String
'
'    With frmMain
'        For intCnt = 1 To UBound(strRecvData)
'            strRcvBuf = strRecvData(intCnt)
'
'            '-- 테스트용 -----------------
'            If .fraCommTest.Visible = False Then
'                Call SetSQLData("RCV", strRcvBuf, "A")
'            End If
'            '-- 테스트용 -----------------
'
'            strRcvBuf = strRecvData(intCnt)
'            strBarno = Trim(mGetP(strRcvBuf, 5, "|"))
'            strRackNo = ""
'            strTubePos = ""
'            strSeq = ""
'
'            With mResult
'                .BarNo = strBarno
'                .SpcPos = strSeq
'                .Seq = strSeq
'                .RackNo = strRackNo
'                .TubePos = strTubePos
'                .RsltDate = Format(Now, "yyyymmddhhmmss")
'                .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'            End With
'
'            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'            If gRow <= 0 Then
'                Exit Sub
'            End If
'
'            strIntBase = mGetP(strRcvBuf, 8, "|")
'            strResult = mGetP(strRcvBuf, 11, "|")
'
'            If strIntBase <> "" And strResult <> "" Then
'                If gPatOrdCd <> "" Then
'                    SQL = ""
'                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                    SQL = SQL & "  FROM EQPMASTER" & vbCr
'                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
'                    SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
'
'                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                        lsTestCode = Trim(RS_L.Fields("TESTCODE"))
'                        lsTestName = Trim(RS_L.Fields("TESTNAME"))
'                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                        '-- 결과Row 추가
'                        lsRstRow = .spdResult.DataRowCnt + 1
'                        If .spdResult.MaxRows < lsRstRow Then
'                            .spdResult.MaxRows = lsRstRow
'                        End If
'
'                        '소수점 처리, 결과 형태 처리
'                        strMachResult = strResult
'                        strResult = SetResult(strResult, strIntBase)
'                        strJudge = SetJudge(strResult, strIntBase)
'
'                        '진행상태 표시("결과")
'                        SetText .spdOrder, "결과", gRow, colSTATE
'
'                        '결과값 표시
'                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                SetText .spdOrder, strResult, gRow, intCol
'                                Exit For
'                            End If
'                        Next
'
'                        '-- 결과 List
'                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
'                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
'                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
'                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
'                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
'                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
'                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
'                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
'                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
'
'                        '-- 로컬 저장
'                        SetLocalDB gRow, lsRstRow, "1", ""
'
'                        strState = "R"
'
'                        '-- 결과Count
'                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                            SetText .spdOrder, "1", gRow, colRCNT
'                        Else
'                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                        End If
'
'                    End If
'                Else
'                    SQL = ""
'                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                    SQL = SQL & "  FROM EQPMASTER" & vbCr
'                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
'
'                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                        lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                        lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                        '-- 결과Row 추가
'                        lsRstRow = .spdResult.DataRowCnt + 1
'                        If .spdResult.MaxRows < lsRstRow Then
'                            .spdResult.MaxRows = lsRstRow
'                        End If
'
'                        '소수점 처리, 결과 형태 처리
'                        strMachResult = strResult
'                        strResult = SetResult(strResult, strIntBase)
'                        strJudge = SetJudge(strResult, strIntBase)
'
'                        '진행상태 표시("결과")
'                        SetText .spdOrder, "결과", gRow, colSTATE
'
'                        '결과값 표시
'                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                SetText .spdOrder, strResult, gRow, intCol
'                                Exit For
'                            End If
'                        Next
'
'                        '-- 결과 List
'                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
'                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
'                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
'                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
'                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
'                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
'                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
'                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
'                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
'
'                        '-- 로컬 저장
'                        SetLocalDB gRow, lsRstRow, "1", ""
'
'                        If strState <> "R" Then
'                            strState = ""
'                        End If
'
'                        '-- 결과Count
'                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                            SetText .spdOrder, "1", gRow, colRCNT
'                        Else
'                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                        End If
'                    End If
'
'                End If
'
'            End If
'
'            .spdResult.RowHeight(-1) = 14
'
'            '## DB에 결과저장
'            If .optTrans(0).Value = True And strState = "R" Then
'                Res = SaveTransData_MCC(gRow)
'
'                If Res = -1 Then
'                    '-- 저장 실패
'                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                    SetText .spdOrder, "Failed", gRow, colSTATE
'                Else
'                    '-- 저장 성공
'                    SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                    SetText .spdOrder, "저장완료", gRow, colSTATE
'                    SetText .spdOrder, "0", gRow, colCHECKBOX
'
'                          SQL = "Update PATRESULT Set " & vbCrLf
'                    SQL = SQL & " sendflag = '2' " & vbCrLf
'                    SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
'                    SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'                    SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'                    SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
'
'                    If DBExec(AdoCn_Local, SQL) Then
'                        '-- 성공
'                    End If
'                End If
'                strState = ""
'            End If
'        Next
'    End With
'
'End Sub


Public Sub SerialRcvData_E411()
   

End Sub


Function SaveTransData(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strSex          As String
    Dim strAge          As String

    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strSubCode      As String
    Dim strEqpcd        As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strRefVal       As String
        
On Error GoTo ErrHandle

    With frmMain
        SaveTransData = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        'AdoCn.BeginTrans
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                      SQL = " Update SLA_LabResult  " & vbCr
                SQL = SQL & "   Set Result     = '" & sResult & "' " & vbCr
                SQL = SQL & "      ,NormalFlag = '0' " & vbCr
                SQL = SQL & "      ,PanicFlag  = '0' " & vbCr
                SQL = SQL & "      ,DeltaFlag  = '0' " & vbCr
                SQL = SQL & "      ,TransFlag  = '1' " & vbCr
                SQL = SQL & "      ,ResultID   = ''  " & vbCr
                SQL = SQL & "      ,ResultDate = '" & Trim(Format(Now, "yyyy-mm-dd")) & "'" & vbCr
                SQL = SQL & "      ,ResultTime = '" & Trim(Format(Time, "HH:MM:SS")) & "'" & vbCr
                SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "'" & vbCr
                SQL = SQL & "   And OrderCode IN (" & gAllOrdCd & ") " & vbCr
                SQL = SQL & "   And LabCode = '" & strTestCd & "'" & vbCr
                SQL = SQL & "   And TransFlag < '2' "

                Call SetSQLData("결과저장", SQL)
                Call DBExec(AdoCn, SQL)
                
                SaveTransData = 1
                
            End If
        Next intRow
        
        If SaveTransData = 1 Then
                  SQL = " Update SLA_LabMaster " & vbCr
            SQL = SQL & "   Set JStatus = '2' " & vbCr
            SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "' " & vbCr
            'SQL = SQL & "   And OrderCode = '" & strTestCd & "'" & vbCr
            SQL = SQL & "   And OrderCode IN (" & gAllOrdCd & ") " & vbCr
            SQL = SQL & "   And RECEIPTDATE = '" & Format(strHospDate, "yyyy-mm-dd") & "'" & vbCr
            SQL = SQL & "   And JStatus < '3' "
            
            Call SetSQLData("상태저장", SQL)
            Call DBExec(AdoCn, SQL)
            
        End If
        
        'AdoCn.CommitTrans
        
    
    End With

Exit Function

ErrHandle:
    SaveTransData = -1
    'AdoCn.RollbackTrans
    
End Function


Function SaveTransData_MCC_R(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strSex          As String
    Dim strAge          As String

    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strSubCode      As String
    Dim strEqpcd        As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strRefVal       As String
    
    'Dim strReturn       As String
    Dim intReturn       As Long
    Dim strMSG          As String
    
    Dim prm0 As New ADODB.Parameter
    Dim Prm1 As New ADODB.Parameter
    Dim Prm2 As New ADODB.Parameter
    Dim Prm3 As New ADODB.Parameter
    Dim Prm4 As New ADODB.Parameter
    Dim Prm5 As New ADODB.Parameter
    
    
    Dim intBarno  As Double
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_MCC_R = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdROrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdROrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdROrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdROrder, argSpcRow, colBARCODE))
        strChartNo = Trim(GetText(.spdROrder, argSpcRow, colCHARTNO))
        strPatID = Trim(GetText(.spdROrder, argSpcRow, colPID))
        
        If IsNumeric(strBarcode) Then
            intBarno = CDbl(strBarcode)
        Else
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                      SQL = "Exec UP_LIS_INTERFACE_U$001 " & intBarno
                SQL = SQL & "," & strTestCd
                SQL = SQL & "," & sResult
                SQL = SQL & "," & gHOSP.MACHCD

                'AdoCn.Execute SQL
                
                Set AdoCmd = New ADODB.Command
                Set AdoCmd.ActiveConnection = AdoCn
                With AdoCmd
                    .CommandTimeout = 15
                    .CommandText = "UP_LIS_INTERFACE_U$001"
                    .CommandType = adCmdStoredProc
                    
                    
                    Set Prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, intBarno)      '바코드번호
                    .Parameters.Append Prm1

                    Set Prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       '처방코드
                    .Parameters.Append Prm2

                    Set Prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '결과값
                    .Parameters.Append Prm3

                    Set Prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, gHOSP.MACHCD)    '장비코드
                    .Parameters.Append Prm4

                    .Execute
                    
                End With
                
                'Call SetSQLData("결과저장", SQL)
                
                SaveTransData_MCC_R = 1
                
            End If
        Next intRow
        
    End With

Exit Function

ErrHandle:
    SaveTransData_MCC_R = -1
    
End Function

Function SaveTransData_MCC(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strSex          As String
    Dim strAge          As String

    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strSubCode      As String
    Dim strEqpcd        As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strRefVal       As String
    
    'Dim strReturn       As String
    Dim intReturn       As Long
    Dim strMSG          As String
    
    Dim prm0 As New ADODB.Parameter
    Dim Prm1 As New ADODB.Parameter
    Dim Prm2 As New ADODB.Parameter
    Dim Prm3 As New ADODB.Parameter
    Dim Prm4 As New ADODB.Parameter
    Dim Prm5 As New ADODB.Parameter
    
    
    Dim intBarno  As Double
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_MCC = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        
        If IsNumeric(strBarcode) Then
            intBarno = CDbl(strBarcode)
        Else
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                      SQL = "Exec UP_LIS_INTERFACE_U$001 " & intBarno
                SQL = SQL & "," & strTestCd
                SQL = SQL & "," & sResult
                SQL = SQL & "," & gHOSP.MACHCD

                'AdoCn.Execute SQL
                
                Set AdoCmd = New ADODB.Command
                Set AdoCmd.ActiveConnection = AdoCn
                With AdoCmd
                    .CommandTimeout = 15
                    .CommandText = "UP_LIS_INTERFACE_U$001"
                    .CommandType = adCmdStoredProc
                    
                    
                    Set Prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, intBarno)      '바코드번호
                    .Parameters.Append Prm1

                    Set Prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       '처방코드
                    .Parameters.Append Prm2

                    Set Prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '결과값
                    .Parameters.Append Prm3

                    Set Prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, gHOSP.MACHCD)    '장비코드
                    .Parameters.Append Prm4

                    .Execute
                    
                End With
                
                'Call SetSQLData("결과저장", SQL)
                
                SaveTransData_MCC = 1
                
            End If
        Next intRow
        
    End With

Exit Function

ErrHandle:
    SaveTransData_MCC = -1
    
End Function

Function SaveTransData_MCC_VERSACELL(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strSex          As String
    Dim strAge          As String

    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strSubCode      As String
    Dim strEqpcd        As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strRefVal       As String
    
    Dim intReturn       As Long
    Dim strMSG          As String
    
    Dim prm0 As New ADODB.Parameter
    Dim Prm1 As New ADODB.Parameter
    Dim Prm2 As New ADODB.Parameter
    Dim Prm3 As New ADODB.Parameter
    Dim Prm4 As New ADODB.Parameter
    Dim Prm5 As New ADODB.Parameter
    
    Dim intBarno        As Double
    
    Dim strMachCD       As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_MCC_VERSACELL = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        
        If IsNumeric(strBarcode) Then
            intBarno = CDbl(strBarcode)
        Else
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "      ,DISKNO "                                                        'VERSACELL 에서 결과저장용 장비코드를 저장해놓고 있다.
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                Call SetText(.vasTemp, RS_L.Fields("DISKNO").Value & "", intRow, 7)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
            strMachCD = Trim(GetText(.vasTemp, intRow, 7))      '장비코드 : ADVIA1800, CENTAURXP
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                      SQL = "Exec UP_LIS_INTERFACE_U$001 " & intBarno
                SQL = SQL & "," & strTestCd
                SQL = SQL & "," & sResult
                SQL = SQL & "," & strMachCD

                'AdoCn.Execute SQL
                
                Set AdoCmd = New ADODB.Command
                Set AdoCmd.ActiveConnection = AdoCn
                With AdoCmd
                    .CommandTimeout = 15
                    .CommandText = "UP_LIS_INTERFACE_U$001"
                    .CommandType = adCmdStoredProc
                                        
                    Set Prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, intBarno)      '바코드번호
                    .Parameters.Append Prm1

                    Set Prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       '처방코드
                    .Parameters.Append Prm2

                    Set Prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '결과값
                    .Parameters.Append Prm3

                    Set Prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, strMachCD)    '장비코드
                    .Parameters.Append Prm4

                    .Execute
                    
                End With
                
                Call SetSQLData("결과저장", SQL)
                
                SaveTransData_MCC_VERSACELL = 1
                
            End If
        Next intRow
                
    End With

Exit Function

ErrHandle:
    SaveTransData_MCC_VERSACELL = -1
    
End Function

Function SaveTransData_HANARO_ROLLER20(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strSex          As String
    Dim strAge          As String

    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strSubCode      As String
    Dim strEqpcd        As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strRefVal       As String
    
    Dim intReturn       As Long
    Dim strMSG          As String
    
    Dim intBarno        As Double
    
    Dim strMachCD       As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_HANARO_ROLLER20 = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        
        If IsNumeric(strBarcode) Then
            intBarno = CDbl(strBarcode)
        Else
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "      ,DISKNO "                                                        'VERSACELL 에서 결과저장용 장비코드를 저장해놓고 있다.
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                Call SetText(.vasTemp, RS_L.Fields("DISKNO").Value & "", intRow, 7)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))      '검사코SUB
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
            strMachCD = Trim(GetText(.vasTemp, intRow, 7))      '장비코드 : ADVIA1800, CENTAURXP
                        
            '-- 장비결과적용
            If gHOSP.SAVELIS = "Y" Then
                sResult = sResult2
            Else
                sResult = sResult1
            End If
                        
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- 서버저장
                SQL = ""
                SQL = SQL & "UPDATE GOLIATH.SZDAT01T                                    " & vbCrLf
                SQL = SQL & "   Set STATUS              = '3'                   " & vbCrLf  '검사상태 (1.준비  3.완료)
                SQL = SQL & "     , NRSLT               = '" & sResult & "'     " & vbCrLf  '숫자결과
                SQL = SQL & "     , INSERT_DATE_DEVICE  = SysDate               " & vbCrLf
                SQL = SQL & "     , INSERT_DEVICE_ID    = '" & gHOSP.USERID & "'" & vbCrLf
                SQL = SQL & "     , UPDATE_DATE_DEVICE  = SysDate               " & vbCrLf
                SQL = SQL & "     , UPDATE_DEVICE_ID    = '" & gHOSP.USERID & "'" & vbCrLf
                SQL = SQL & "     , UPLOAD_UPDATE_DATE  = SysDate               " & vbCrLf
                SQL = SQL & " Where COMP                = '" & gHOSP.HOSPCD & "'" & vbCrLf  '회사코드
                SQL = SQL & "   And WORKCD              = '" & gHOSP.MACHCD & "'" & vbCrLf   '작업장치
                SQL = SQL & "   And ACPTNO              = '" & strBarcode & "'  " & vbCrLf  '접수번호(검체번호)
                SQL = SQL & "   And CODE1               = '" & strTestCd & "'   " & vbCrLf  '검사코드
                If Trim(strSubCode) <> "" Then
                    SQL = SQL & "   And CODE2               = '" & strTestCd & "' " & vbCrLf  '검사코드순번
                End If
                SQL = SQL & "   And STATUS  = '1'                   " & vbCrLf  '검사상태 (1.준비  3.완료)
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
            
                SaveTransData_HANARO_ROLLER20 = 1
                
            End If
        Next intRow
                
    End With

Exit Function

ErrHandle:
    SaveTransData_HANARO_ROLLER20 = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_HANARO_ROLLER20" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function

Function SaveTransData_HANARO_ROLLER20_R(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strSex          As String
    Dim strAge          As String

    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strSubCode      As String
    Dim strEqpcd        As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strRefVal       As String
    
    Dim intReturn       As Long
    Dim strMSG          As String
    
    Dim intBarno        As Double
    
    Dim strMachCD       As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_HANARO_ROLLER20_R = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdROrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdROrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdROrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdROrder, argSpcRow, colBARCODE))
        strChartNo = Trim(GetText(.spdROrder, argSpcRow, colCHARTNO))
        strPatID = Trim(GetText(.spdROrder, argSpcRow, colPID))
        
        If IsNumeric(strBarcode) Then
            intBarno = CDbl(strBarcode)
        Else
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "      ,DISKNO "                                                        'VERSACELL 에서 결과저장용 장비코드를 저장해놓고 있다.
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                Call SetText(.vasTemp, RS_L.Fields("DISKNO").Value & "", intRow, 7)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))      '검사코SUB
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
            strMachCD = Trim(GetText(.vasTemp, intRow, 7))      '장비코드 : ADVIA1800, CENTAURXP
                        
            '-- 장비결과적용
            If gHOSP.SAVELIS = "Y" Then
                sResult = sResult2
            Else
                sResult = sResult1
            End If
                        
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- 서버저장
                SQL = ""
                SQL = SQL & "UPDATE GOLIATH.SZDAT01T                                    " & vbCrLf
                SQL = SQL & "   Set STATUS              = '3'                   " & vbCrLf  '검사상태 (1.준비  3.완료)
                SQL = SQL & "     , NRSLT               = '" & sResult & "'     " & vbCrLf  '숫자결과
                SQL = SQL & "     , INSERT_DATE_DEVICE  = SysDate               " & vbCrLf
                SQL = SQL & "     , INSERT_DEVICE_ID    = '" & gHOSP.USERID & "'" & vbCrLf
                SQL = SQL & "     , UPDATE_DATE_DEVICE  = SysDate               " & vbCrLf
                SQL = SQL & "     , UPDATE_DEVICE_ID    = '" & gHOSP.USERID & "'" & vbCrLf
                SQL = SQL & "     , UPLOAD_UPDATE_DATE  = SysDate               " & vbCrLf
                SQL = SQL & " Where COMP                = '" & gHOSP.HOSPCD & "'" & vbCrLf  '회사코드
                SQL = SQL & "   And WORKCD              = '" & gHOSP.MACHCD & "'" & vbCrLf   '작업장치
                SQL = SQL & "   And ACPTNO              = '" & strBarcode & "'  " & vbCrLf  '접수번호(검체번호)
                SQL = SQL & "   And CODE1               = '" & strTestCd & "'   " & vbCrLf  '검사코드
                If Trim(strSubCode) <> "" Then
                    SQL = SQL & "   And CODE2               = '" & strTestCd & "' " & vbCrLf  '검사코드순번
                End If
                SQL = SQL & "   And STATUS  = '1'                   " & vbCrLf  '검사상태 (1.준비  3.완료)
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
            
                SaveTransData_HANARO_ROLLER20_R = 1
                
            End If
        Next intRow
                
    End With

Exit Function

ErrHandle:
    SaveTransData_HANARO_ROLLER20_R = -1
    Screen.MousePointer = 0
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "_SaveTransData_HANARO_ROLLER20_R" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show
    
End Function


Function SaveTransData_SCL(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    Dim RsLocal         As ADODB.Recordset
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    
    Dim strEqpcd        As String
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strTestCdSub    As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
    Dim Prm1 As New ADODB.Parameter
    Dim Prm2 As New ADODB.Parameter
    Dim Prm3 As New ADODB.Parameter
    
'1) 결과를 Minterface 테이블에 저장 한다(기존자료에 update).
'* 저장 시 Minterface 테이블의 flag필드를 ‘2’로 update한다.
'* 재검인 경우 Minterface 테이블의 reflag 필드를 ‘1’로 update한다. (Procedure는 결과 처리시 기존의 데이터는 이력으로 남기고 재검카운트를 증가하고 update된 결과를 저장한다.)
'* qc data는 MintQC 테이블에 flag 필드를 ‘1’로하여 신규로 insert하고 Procedure는 해당 일자 및 장비의 flag ‘1’인 결과를 찾아 처리하고 flag 필드를 ‘2’로 update.

'2) 결과 저장이 완료되면 해당 procedure를 call 한다.
'==========  batch slrtrm55p(
'pmach : char(3) => 장비코드,
'perr : char(1) => 인증확인 및 에러코드),

'==========  real slrtrm56p(
'pbarc : char(12) => 바코드번호,
'pmach : char(3) => 장비코드,
'perr : char(1) => 인증확인 및 에러코드)
    
On Error GoTo ErrHandle
    
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmMain
        SaveTransData_SCL = -1
        
        strSaveSeq = Trim(GetText(SPD, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(SPD, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(SPD, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(SPD, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(SPD, argSpcRow, colPID))
        strPatNm = Trim(GetText(SPD, argSpcRow, colPNAME))
        strChartNo = Trim(GetText(SPD, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT,REFJUDGE    " & vbCrLf
        SQL = SQL & "  FROM PATRESULT                                                               " & vbCrLf
        SQL = SQL & " WHERE MID(EXAMDATE,1,8)    = '" & Mid(strExamDate, 1, 8) & "'                                     " & vbCrLf
        SQL = SQL & "   AND SAVESEQ     = " & strSaveSeq & vbCrLf
        SQL = SQL & "   AND BARCODE     = '" & strBarcode & "'                                      " & vbCrLf
        SQL = SQL & "   AND EXAMCODE    <> ''                                                       " & vbCrLf
        
        Call SetSQLData("로컬조회", SQL, "A")
        
        Set RsLocal = New ADODB.Recordset
        Set RsLocal = AdoCn_Local.Execute(SQL, , 1)
        If Not RsLocal.EOF = True And Not RsLocal.BOF = True Then
            Do Until RsLocal.EOF
                strEqpcd = RsLocal.Fields("EQUIPCODE").Value & ""
                strOrdCd = RsLocal.Fields("ORDERCODE").Value & ""
                strTestCd = RsLocal.Fields("EXAMCODE").Value & ""
                strTestCdSub = RsLocal.Fields("EXAMSUBCODE").Value & ""
                sResult1 = RsLocal.Fields("EQUIPRESULT").Value & ""
                sResult2 = RsLocal.Fields("RESULT").Value & ""
                strJudge = RsLocal.Fields("REFJUDGE").Value & ""
                
                '-- 장비결과적용
                If gHOSP.SAVELIS = "Y" Then
                    sResult = sResult2
                Else
                    sResult = sResult1
                End If
                
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- 서버저장
                    SQL = ""    'DB명.테이블명 대소문자 구분해야 함
                    SQL = SQL & "Update HLISILIB.MINTERFACE  " & vbCrLf
                    SQL = SQL & "   Set Result      = '" & Trim(sResult) & "' " & vbCrLf
                    SQL = SQL & "     , Rltflag     = 'N'                       " & vbCrLf
                    SQL = SQL & "     , Updtdate    = (select substring(char(curdate()),1,4) || substring(char(curdate()),6,2) || substring(char(curdate()),9,2) || substring(char(curtime()),4,2) || substring(char(curtime()),7,2) || substring(char(curtime()),10,2) from sysibm.sysdummy1) " & vbCrLf
                    'SQL = SQL & "     , Testercode  = '" & gHOSP.USERID & "'    " & vbCrLf
                    SQL = SQL & "     , Testercode  = '" & frmMain.txtSCLID.Text & "'    " & vbCrLf
                    
                    SQL = SQL & "     , Flag        = '2'                       " & vbCrLf
                    'SQL = SQL & "     , Updtempl    = '" & gHOSP.USERID & "'    " & vbCrLf
                    SQL = SQL & "     , Updtempl    = '" & frmMain.txtSCLID.Text & "'    " & vbCrLf
                    SQL = SQL & " Where barcodeno   = '" & strBarcode & "'      " & vbCrLf
                    SQL = SQL & "   And mcode       = '" & gHOSP.SCLMACHCD & "'    " & vbCrLf
                    SQL = SQL & "   And itemcode    = '" & Mid(strTestCd, 1, 5) & "'" & vbCrLf
                    If Len(strTestCd) > 5 Then
                       SQL = SQL & "   And dcode = '" & Mid(strTestCd, 6) & "'"
                    End If

                    Call SetSQLData("결과저장", SQL, "A")
                    AdoCn_SCL.Execute SQL
                
                
                    '-- 검사결과저장 = SLRTRM55P
                    Set AdoCmd_SCL = New ADODB.Command
                    Set AdoCmd_SCL.ActiveConnection = AdoCn_SCL
                    With AdoCmd_SCL
                        .CommandTimeout = 15
                        .CommandText = "SLRTRM56P"
                        .CommandType = adCmdStoredProc
                        
                        Set Prm1 = .CreateParameter("pbarc", adChar, adParamInput, 12, strBarcode)
                        .Parameters.Append Prm1
                        Set Prm2 = .CreateParameter("pmach", adChar, adParamInput, 3, gHOSP.SCLMACHCD)
                        .Parameters.Append Prm2
                        Set Prm3 = .CreateParameter("perr", adChar, adParamOutput, 1, "")
                        .Parameters.Append Prm3
                        
                        .Execute
                        
                        Call SetSQLData("결과저장", strBarcode & "," & gHOSP.SCLMACHCD, "A")
                        
                        Set AdoCmd_SCL = Nothing
                        
                    End With
                End If
                RsLocal.MoveNext
            Loop
        End If
        
        RsLocal.Close
        
        SaveTransData_SCL = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_SCL = -1
    
    Screen.MousePointer = 1
    
    strErrMsg = ""
    strErrMsg = strErrMsg & "위    치 : " & gHOSP.MACHNM & "SaveTransData_SCL" & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류번호 : " & Err.Number & vbNewLine & vbNewLine
    strErrMsg = strErrMsg & "오류내용 : " & Err.Description & vbNewLine
    frmErrMsg.txtErr = vbNewLine & strErrMsg
    frmErrMsg.Show vbModal
    
End Function



Function SaveTransData_MCC_VERSACELL_R(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strSex          As String
    Dim strAge          As String

    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strSubCode      As String
    Dim strEqpcd        As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strRefVal       As String
    
    Dim intReturn       As Long
    Dim strMSG          As String
    
    Dim prm0 As New ADODB.Parameter
    Dim Prm1 As New ADODB.Parameter
    Dim Prm2 As New ADODB.Parameter
    Dim Prm3 As New ADODB.Parameter
    Dim Prm4 As New ADODB.Parameter
    Dim Prm5 As New ADODB.Parameter
    
    Dim intBarno        As Double
    
    Dim strMachCD       As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_MCC_VERSACELL_R = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdROrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdROrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdROrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdROrder, argSpcRow, colBARCODE))
        strChartNo = Trim(GetText(.spdROrder, argSpcRow, colCHARTNO))
        strPatID = Trim(GetText(.spdROrder, argSpcRow, colPID))
        
        If IsNumeric(strBarcode) Then
            intBarno = CDbl(strBarcode)
        Else
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "      ,DISKNO "                                                        'VERSACELL 에서 결과저장용 장비코드를 저장해놓고 있다.
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        SQL = SQL & "   AND SENDFLAG <> '2' "
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                Call SetText(.vasTemp, RS_L.Fields("DISKNO").Value & "", intRow, 7)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
            strMachCD = Trim(GetText(.vasTemp, intRow, 7))      '장비코드 : ADVIA1800, CENTAURXP
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                      SQL = "Exec UP_LIS_INTERFACE_U$001 " & intBarno
                SQL = SQL & "," & strTestCd
                SQL = SQL & "," & sResult
                SQL = SQL & "," & strMachCD

                'AdoCn.Execute SQL
                
                Set AdoCmd = New ADODB.Command
                Set AdoCmd.ActiveConnection = AdoCn
                With AdoCmd
                    .CommandTimeout = 15
                    .CommandText = "UP_LIS_INTERFACE_U$001"
                    .CommandType = adCmdStoredProc
                    
                    
                    Set Prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, intBarno)      '바코드번호
                    .Parameters.Append Prm1

                    Set Prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       '처방코드
                    .Parameters.Append Prm2

                    Set Prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '결과값
                    .Parameters.Append Prm3

                    Set Prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, strMachCD)    '장비코드
                    .Parameters.Append Prm4

                    .Execute
                    
                End With
                
                Call SetSQLData("결과저장", SQL)
                
                SaveTransData_MCC_VERSACELL_R = 1
                
            End If
        Next intRow
                
    End With

Exit Function

ErrHandle:
    SaveTransData_MCC_VERSACELL_R = -1
    
End Function

Function SetJudge(asResult As String, asEquipCode As String)
    Dim RS_L        As ADODB.Recordset
    Dim i As Integer
    Dim sLVal As String
    Dim sHVal As String
    Dim sEquipCode As String
    Dim sEquipRes As String
    Dim sResFlag As String
    
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    sResFlag = ""
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT REFLOW, REFHIGH " & vbCr
    SQL = SQL & "  FROM EQPMASTER " & vbCr
    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & "   AND RSLTCHANNEL = '" & sEquipCode & "'"

    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
        If IsNumeric(Trim(RS_L.Fields("REFLOW")) & "") = True And IsNumeric(Trim(RS_L.Fields("REFHIGH")) & "") = True Then
            sLVal = Trim(RS_L.Fields("REFLOW")) & ""
            sHVal = Trim(RS_L.Fields("REFHIGH")) & ""
            If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
                sResFlag = ""
            ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
                sResFlag = "H"
            ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
                sResFlag = "L"
            End If
        End If
    End If
 
    SetJudge = sResFlag
    
End Function

Function SetJudge_Hanaro_Roller20(asResult As String, asEquipCode As String, asSex As String)
    Dim RS_L        As ADODB.Recordset
    Dim i As Integer
    Dim sLVal As String
    Dim sHVal As String
    Dim sEquipCode As String
    Dim sEquipRes As String
    Dim sResFlag As String
    
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    sResFlag = ""
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    If UCase(mResult.SEX) = "M" Or mResult.SEX = "남" Or mResult.SEX = "남자" Then
        sLVal = 0
        sHVal = 15
        
        If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
            sResFlag = ""
        ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
            sResFlag = "H"
        ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
            sResFlag = "L"
        End If
            
    Else
        sLVal = 0
        sHVal = 20
        
        If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
            sResFlag = ""
        ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
            sResFlag = "H"
        ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
            sResFlag = "L"
        End If
    End If
 
    SetJudge_Hanaro_Roller20 = sResFlag
    
End Function

Function SetResult(asResult As String, asEquipCode As String)
    Dim RS_L        As ADODB.Recordset
    Dim i As Integer
    Dim sEquipCode As String
    Dim sEquipRes As String
    Dim sResult As String
    Dim sPoint As Integer
    Dim sResType As String
    
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT RESPREC, REFLOW, REFHIGH " & vbCr
    SQL = SQL & "  FROM EQPMASTER " & vbCr
    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & "   AND RSLTCHANNEL = '" & sEquipCode & "'"

    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
        If IsNumeric(Trim(RS_L.Fields("RESPREC")) & "") = True Then
            sPoint = CInt(Trim(RS_L.Fields("RESPREC")))
            sResType = ""
            For i = 0 To sPoint
                If i = 0 Then
                    sResType = "#0"
                ElseIf i = 1 Then
                    sResType = sResType & ".0"
                Else
                    sResType = sResType & "0"
                End If
            Next
            sResult = Format(sEquipRes, sResType)
        Else
            sResult = sEquipRes
        End If
    End If
 
    SetResult = sResult
    
End Function
