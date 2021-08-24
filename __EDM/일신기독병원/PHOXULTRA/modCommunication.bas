Attribute VB_Name = "modCommunication"
Option Explicit

Public pBuffer As Variant

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
    Result   As String
    EqpCd    As String
    '-- 비트워크리스트용
    RESODRSEQ   As String
    RESSEQ      As String
    RESSUBSEQ   As String
End Type

Public mResult As IntfData

'-- OCS접수정보
Type OcsData
    FPID    As String
    FOrdCd  As String
    FPNM    As String
    FJNO    As String
    FJNO1   As String
    FJNO2   As String
    FWard   As String
    FRoom   As String
    FDept   As String
    FMdDt   As String
    FAcDt   As String
    FAcTm   As String
    FHpDt   As String
    FHpTm   As String
    FDoct   As String
    FDocID  As String
    FWorkNo As String
    FBarCode As String
End Type

Public mOCS As OcsData

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
    Dim blnRcp      As Boolean
    
    intRow = -1
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
'                If IsNumeric(pBarno) And IsNumeric(Trim(GetText(frmMain.spdOrder, i, colBARCODE))) Then
'                    If Val(Trim(GetText(frmMain.spdOrder, i, colBARCODE))) = Val(pBarno) Then
'                        If Trim(GetText(frmMain.spdOrder, i, colSTATE)) = "" Or InStr(GetText(frmMain.spdOrder, i, colSTATE), "오더") > 0 Then
'                            intRow = i
'                            Exit For
'                        End If
'                    End If
'                End If
            
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = Trim(pBarno) Then
                    mResult.BarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                    mResult.PatNo = Trim(GetText(frmMain.spdOrder, i, colPID))
                    intRow = i
                    Exit For
                End If
            
            Next i
        Else
            Select Case pType
                '-- Seq
                Case "1"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
'                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
'                            mOrder.BarNo = pBarno
                            
                            mResult.BarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mResult.PatNo = Trim(GetText(frmMain.spdOrder, i, colPID))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            mResult.BarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mResult.PatNo = Trim(GetText(frmMain.spdOrder, i, colPID))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        'If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" And GetText(frmMain.spdOrder, i, colSTATE) = "" Then
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            mResult.BarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mResult.PatNo = Trim(GetText(frmMain.spdOrder, i, colPID))
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
    
        '-- 환자정보 표시
        'Call vasActiveCell(.spdOrder, intRow, colBARCODE)
        
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        If GetSampleInfo(intRow, .spdOrder) = 1 Then
            '-- 접수 마스터 저장
            Call Reg_acawnifh_ILSIN
                    
            SetText .spdOrder, mOCS.FBarCode, intRow, colCHARTNO             '
        
        End If
        
        .spdOrder.RowHeight(-1) = 12
    
    End With
    
    '-- 현재 Row
    gRow = intRow
    
End Sub


Public Sub SetPatInfo_BS240_HL7(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    With frmMain
        For i = 1 To .spdOrder.DataRowCnt
            If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = Trim(pBarno) Then
                intRow = i
                Exit For
            End If
        Next i
        
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
    
        '-- 환자정보 표시
        'Call vasActiveCell(.spdOrder, intRow, colBARCODE)
        
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
    
    End With
    
    '-- 현재 Row
    gRow = intRow
    
End Sub

Public Sub SetPatInfo_BS220_HL7(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    With frmMain
        For i = 1 To .spdOrder.DataRowCnt
            If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = Trim(pBarno) Then
                intRow = i
                Exit For
            End If
        Next i
        
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
    
        '-- 환자정보 표시
        'Call vasActiveCell(.spdOrder, intRow, colBARCODE)
        
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
    
    End With
    
    '-- 현재 Row
    gRow = intRow
    
End Sub

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
Public Sub SetPatInfo_H7080(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    With frmMain
        For i = 1 To .spdOrder.DataRowCnt
            If IsNumeric(pBarno) And IsNumeric(Trim(GetText(frmMain.spdOrder, i, colCHARTNO))) Then
                If Val(Trim(GetText(frmMain.spdOrder, i, colCHARTNO))) = Val(pBarno) Then
                    intRow = i
                    Exit For
                End If
            End If
        Next i
        
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
        'Call SetText(.spdOrder, mResult.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mResult.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mResult.TubePos, intRow, colPOSNO)
    
        '-- 환자정보 표시
        'Call vasActiveCell(.spdOrder, intRow, colBARCODE)
        
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
    
    End With
    
    '-- 현재 Row
    gRow = intRow
    
End Sub


'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Public Sub SendOrder_E411()
    
    
End Sub


'Public Sub Phase_TCP_BA400()
'
'End Sub
'
'Public Sub Phase_TCP_OSMOPRO()
'
'    Dim Buffer      As Variant
'    Dim BufChar     As String
'    Dim lngBufLen   As Long
'    Dim i           As Long
'
'    lngBufLen = Len(pBuffer)
'
'    For i = 1 To lngBufLen
'        BufChar = Mid$(pBuffer, i, 1)
'
'        Select Case intPhase
'            Case 1      '## Estabilshment Phase
'                Select Case BufChar
'                    Case ENQ
'                        Erase strRecvData
'                        intPhase = 2
'                        'frmMain.comEqp.Output = ACK
'                        frmMain.wSck.SendData ACK
'                        SetRawData "[Tx]" & ACK
'
'                    Case ACK
'                        If strState = "Q" Then
'                            'Call SendOrder_VERSACELL
'                        End If
'                End Select
'            Case 2      '## Transfer Phase
'                Select Case BufChar
'                    Case ENQ
'                        Erase strRecvData
'                        frmMain.comEqp.Output = ACK
'                        SetRawData "[Tx]" & ACK
'                    Case STX
'                        If intBufCnt = 0 Then
'                            intBufCnt = 1
'                            Erase strRecvData
'                            ReDim Preserve strRecvData(intBufCnt)
'                        Else
'                            intBufCnt = intBufCnt + 1
'                            ReDim Preserve strRecvData(intBufCnt)
'                        End If
'                    Case ETB
'                        blnIsETB = True
'                        intPhase = 3
'                    Case ETX
'                        intBufCnt = intBufCnt + 1
'                        ReDim Preserve strRecvData(intBufCnt)
'                        intPhase = 3
'                    Case vbCr
'                        intBufCnt = intBufCnt + 1
'                        ReDim Preserve strRecvData(intBufCnt)
'                    Case EOT
'                        intPhase = 1
'                    Case Else
'                        If blnIsETB = False Then
'                            strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'                        Else
'                            blnIsETB = False
'                        End If
'                End Select
'            Case 3      '## Transfer Phase
'                Select Case BufChar
'                    Case vbCr
'                        intPhase = 4
'                        'frmMain.comEqp.Output = ACK
'                        frmMain.wSck.SendData ACK
'                        SetRawData "[Tx]" & ACK
'                End Select
'            Case 4      '## Termination Phase
'                Select Case BufChar
'                    Case STX
'                        intPhase = 2
'                    Case EOT
'                        Call frmMain.SerialRcvData_OSMOPRO
''                        If strState = "Q" Then
''                            intSndPhase = 1
''                            intFrameNo = 1
''                            frmMain.comEqp.Output = ENQ
''                            SetRawData "[Tx]" & ENQ
''                        End If
'                        intPhase = 1
'                End Select
'        End Select
'    Next i
'
'End Sub
    

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
'    Dim strOrderCode     As String   '처방코드
'    Dim strTestCode      As String   '검사코드
'    Dim strTestName      As String   '검사명
'    Dim strSeqNo         As String   '로컬DB 검사Seq
'
'    Dim strRstRow        As String   '결과스프레드 현재 Row
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
'                                strTestCode = Trim(RS_L.Fields("TESTCODE"))
'                                strTestName = Trim(RS_L.Fields("TESTNAME"))
'                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                '-- 결과Row 추가
'                                strRstRow = .spdResult.DataRowCnt + 1
'                                If .spdResult.MaxRows < strRstRow Then
'                                    .spdResult.MaxRows = strRstRow
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
'                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                        SetText .spdOrder, strResult, gRow, intCol
'                                        Exit For
'                                    End If
'                                Next
'
'                                '-- 결과 List
'                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
'                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
'                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
'                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
'                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
'                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
'                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
'                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
'
'                                '-- 로컬 저장
'                                SetLocalDB gRow, strRstRow, "1", ""
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
'                                strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                                strTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                                strSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                '-- 결과Row 추가
'                                strRstRow = .spdResult.DataRowCnt + 1
'                                If .spdResult.MaxRows < strRstRow Then
'                                    .spdResult.MaxRows = strRstRow
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
'                                    If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                        SetText .spdOrder, strResult, gRow, intCol
'                                        Exit For
'                                    End If
'                                Next
'
'                                '-- 결과 List
'                                SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
'                                SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
'                                SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'                                SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
'                                SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
'                                SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
'                                SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
'                                SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
'                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
'
'                                '-- 로컬 저장
'                                SetLocalDB gRow, strRstRow, "1", ""
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
'    Dim strOrderCode     As String   '처방코드
'    Dim strTestCode      As String   '검사코드
'    Dim strTestName      As String   '검사명
'    Dim strSeqNo         As String   '로컬DB 검사Seq
'
'    Dim strRstRow        As String   '결과스프레드 현재 Row
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
'                                        strTestCode = Trim(RS_L.Fields("TESTCODE"))
'                                        strTestName = Trim(RS_L.Fields("TESTNAME"))
'                                        strSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                        '-- 결과Row 추가
'                                        strRstRow = .spdResult.DataRowCnt + 1
'                                        If .spdResult.MaxRows < strRstRow Then
'                                            .spdResult.MaxRows = strRstRow
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
'                                            If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                                SetText .spdOrder, strResult, gRow, intCol
'                                                Exit For
'                                            End If
'                                        Next
'
'                                        '-- 결과 List
'                                        SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
'                                        SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
'                                        SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'                                        SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'                                        SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
'                                        SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
'                                        SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
'                                        SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
'                                        SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
'                                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
'
'                                        '-- 로컬 저장
'                                        SetLocalDB gRow, strRstRow, "1", ""
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
'                                        strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                                        strTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                                        strSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                        '-- 결과Row 추가
'                                        strRstRow = .spdResult.DataRowCnt + 1
'                                        If .spdResult.MaxRows < strRstRow Then
'                                            .spdResult.MaxRows = strRstRow
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
'                                            If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                                SetText .spdOrder, strResult, gRow, intCol
'                                                Exit For
'                                            End If
'                                        Next
'
'                                        '-- 결과 List
'                                        SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
'                                        SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
'                                        SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'                                        SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
'                                        SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
'                                        SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
'                                        SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
'                                        SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
'                                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
'
'                                        '-- 로컬 저장
'                                        SetLocalDB gRow, strRstRow, "1", ""
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
'    Dim strOrderCode     As String   '처방코드
'    Dim strTestCode      As String   '검사코드
'    Dim strTestName      As String   '검사명
'    Dim strSeqNo         As String   '로컬DB 검사Seq
'
'    Dim strRstRow        As String   '결과스프레드 현재 Row
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
'                                        strTestCode = Trim(RS_L.Fields("TESTCODE"))
'                                        strTestName = Trim(RS_L.Fields("TESTNAME"))
'                                        strSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                        '-- 결과Row 추가
'                                        strRstRow = .spdResult.DataRowCnt + 1
'                                        If .spdResult.MaxRows < strRstRow Then
'                                            .spdResult.MaxRows = strRstRow
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
'                                            If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                                SetText .spdOrder, strResult, gRow, intCol
'                                                Exit For
'                                            End If
'                                        Next
'
'                                        '-- 결과 List
'                                        SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
'                                        SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
'                                        SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'                                        SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
'                                        SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
'                                        SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
'                                        SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
'                                        SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
'                                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
'
'                                        '-- 로컬 저장
'                                        SetLocalDB gRow, strRstRow, "1", ""
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
'                                        strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                                        strTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                                        strSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                        '-- 결과Row 추가
'                                        strRstRow = .spdResult.DataRowCnt + 1
'                                        If .spdResult.MaxRows < strRstRow Then
'                                            .spdResult.MaxRows = strRstRow
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
'                                            If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                                SetText .spdOrder, strResult, gRow, intCol
'                                                Exit For
'                                            End If
'                                        Next
'
'                                        '-- 결과 List
'                                        SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
'                                        SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
'                                        SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'                                        SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
'                                        SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
'                                        SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
'                                        SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
'                                        SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
'                                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
'
'                                        '-- 로컬 저장
'                                        SetLocalDB gRow, strRstRow, "1", ""
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
'    Dim strOrderCode     As String   '처방코드
'    Dim strTestCode      As String   '검사코드
'    Dim strTestName      As String   '검사명
'    Dim strSeqNo         As String   '로컬DB 검사Seq
'
'    Dim strRstRow        As String   '결과스프레드 현재 Row
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
'                        strTestCode = Trim(RS_L.Fields("TESTCODE"))
'                        strTestName = Trim(RS_L.Fields("TESTNAME"))
'                        strSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                        '-- 결과Row 추가
'                        strRstRow = .spdResult.DataRowCnt + 1
'                        If .spdResult.MaxRows < strRstRow Then
'                            .spdResult.MaxRows = strRstRow
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
'                            If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                SetText .spdOrder, strResult, gRow, intCol
'                                Exit For
'                            End If
'                        Next
'
'                        '-- 결과 List
'                        SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
'                        SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
'                        SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'                        SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
'                        SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
'                        SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
'                        SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
'                        SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
'                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
'
'                        '-- 로컬 저장
'                        SetLocalDB gRow, strRstRow, "1", ""
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
'                        strTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                        strTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                        strSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                        '-- 결과Row 추가
'                        strRstRow = .spdResult.DataRowCnt + 1
'                        If .spdResult.MaxRows < strRstRow Then
'                            .spdResult.MaxRows = strRstRow
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
'                            If strTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                SetText .spdOrder, strResult, gRow, intCol
'                                Exit For
'                            End If
'                        Next
'
'                        '-- 결과 List
'                        SetText .spdResult, strSeqNo, strRstRow, colRSEQNO                '순번
'                        SetText .spdResult, strOrderCode, strRstRow, colRORDERCD          '처방코드
'                        SetText .spdResult, strTestCode, strRstRow, colRTESTCD            '검사코드
'                        SetText .spdResult, strTestName, strRstRow, colRTESTNM            '검사명
'                        SetText .spdResult, strIntBase, strRstRow, colRCHANNEL           '장비채널
'                        SetText .spdResult, strMachResult, strRstRow, colRMACHRESULT     '장비결과
'                        SetText .spdResult, strResult, strRstRow, colRLISRESULT          'LIS결과
'                        SetText .spdResult, strJudge, strRstRow, colRJUDGE                     '판정
'                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), strRstRow, colRREF          '참고치
'
'                        '-- 로컬 저장
'                        SetLocalDB gRow, strRstRow, "1", ""
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


'Function SaveTransData(ByVal argSpcRow As Integer) As Integer
'    Dim RS_L            As ADODB.Recordset
'    Dim intRow          As Integer
'    Dim strDate         As String
'
'    Dim strSaveSeq      As String
'    Dim strExamDate     As String
'    Dim strHospDate     As String
'    Dim strBarcode      As String
'    Dim strChartNo      As String
'    Dim strPatID        As String
'    Dim strSex          As String
'    Dim strAge          As String
'
'    Dim strOrdCd        As String
'    Dim strTestCd       As String
'    Dim strSubCode      As String
'    Dim strEqpcd        As String
'    Dim sResult         As String
'    Dim sResult1        As String
'    Dim sResult2        As String
'    Dim strRefVal       As String
'
'On Error GoTo ErrHandle
'
'    With frmMain
'        SaveTransData = -1
'        intRow = 0
'
'        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
'        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
'        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
'        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
'        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
'        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
'
'
'        '-- Local에서 환자별로 결과값 가져오기
'        .vasTemp.MaxRows = 0
'
'              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
'        SQL = SQL & "  FROM PATRESULT " & vbCr
'        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
'        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
'        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
'        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
'
'        Call SetSQLData("로컬결과조회", SQL)
'
'        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'            .vasTemp.MaxRows = RS_L.RecordCount
'            Do Until RS_L.EOF
'                intRow = intRow + 1
'                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
'                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
'                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
'                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
'                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
'                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
'                RS_L.MoveNext
'            Loop
'        End If
'
'        RS_L.Close
'
'        sResult = ""
'        sResult1 = ""
'        sResult2 = ""
'
'        'AdoCn.BeginTrans
'
'        '-- 서버로 결과값 저장하기
'        For intRow = 1 To .vasTemp.DataRowCnt
'            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
'            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
'            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
'
'            '-- 장비결과적용
'            If .optSaveResult(0).Value = True Then
'                sResult = sResult1
'            Else
'                sResult = sResult2
'            End If
'
'            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
'                      SQL = " Update SLA_LabResult  " & vbCr
'                SQL = SQL & "   Set Result     = '" & sResult & "' " & vbCr
'                SQL = SQL & "      ,NormalFlag = '0' " & vbCr
'                SQL = SQL & "      ,PanicFlag  = '0' " & vbCr
'                SQL = SQL & "      ,DeltaFlag  = '0' " & vbCr
'                SQL = SQL & "      ,TransFlag  = '1' " & vbCr
'                SQL = SQL & "      ,ResultID   = ''  " & vbCr
'                SQL = SQL & "      ,ResultDate = '" & Trim(Format(Now, "yyyy-mm-dd")) & "'" & vbCr
'                SQL = SQL & "      ,ResultTime = '" & Trim(Format(Time, "HH:MM:SS")) & "'" & vbCr
'                SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "'" & vbCr
'                SQL = SQL & "   And OrderCode IN (" & gAllOrdCd & ") " & vbCr
'                SQL = SQL & "   And LabCode = '" & strTestCd & "'" & vbCr
'                SQL = SQL & "   And TransFlag < '2' "
'
'                Call SetSQLData("결과저장", SQL)
'                Call DBExec(AdoCn, SQL)
'
'                SaveTransData = 1
'
'            End If
'        Next intRow
'
'        If SaveTransData = 1 Then
'                  SQL = " Update SLA_LabMaster " & vbCr
'            SQL = SQL & "   Set JStatus = '2' " & vbCr
'            SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "' " & vbCr
'            'SQL = SQL & "   And OrderCode = '" & strTestCd & "'" & vbCr
'            SQL = SQL & "   And OrderCode IN (" & gAllOrdCd & ") " & vbCr
'            SQL = SQL & "   And RECEIPTDATE = '" & Format(strHospDate, "yyyy-mm-dd") & "'" & vbCr
'            SQL = SQL & "   And JStatus < '3' "
'
'            Call SetSQLData("상태저장", SQL)
'            Call DBExec(AdoCn, SQL)
'
'        End If
'
'        'AdoCn.CommitTrans
'
'
'    End With
'
'Exit Function
'
'ErrHandle:
'    SaveTransData = -1
'    'AdoCn.RollbackTrans
'
'End Function

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
    Dim prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim prm3 As New ADODB.Parameter
    Dim prm4 As New ADODB.Parameter
    Dim prm5 As New ADODB.Parameter
    
    
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
                    
                    
                    Set prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, intBarno)      '바코드번호
                    .Parameters.Append prm1

                    Set prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       '처방코드
                    .Parameters.Append prm2

                    Set prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '결과값
                    .Parameters.Append prm3

                    Set prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, gHOSP.MACHCD)    '장비코드
                    .Parameters.Append prm4

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

Function SaveTransData_GINUS(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim blnSave         As Boolean
    Dim intRow          As Integer
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    
    Dim strTestCd       As String
    Dim strSubCode      As String
    Dim strEqpcd        As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
On Error GoTo ErrHandle
    
    blnSave = False
    intRow = 0
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmMain
        SaveTransData_GINUS = -1
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT                                                   " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'                            " & vbCr                      '장비코드
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'        " & vbCr  '검사일
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        
        Set RS_L = New ADODB.Recordset
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
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사SUB코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            strJudge = SetJudge(strTestCd, sResult)
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- 검사결과 저장하기
                SQL = ""
                SQL = SQL & "Update scrrslth                                            " & vbCr
                SQL = SQL & "   SET exam_stus  = decode(exam_stus, '0','1',exam_stus)   " & vbCr
                SQL = SQL & "     , exam_rslt  = '" & sResult & "'                      " & vbCr
                SQL = SQL & "     , mach_rslt  = '" & sResult & "'                      " & vbCr
                SQL = SQL & "     , exam_dt    = '" & Format(Now, "yyyymmddhhmm") & "'  " & vbCr
                SQL = SQL & "     , exam_empno = '" & gHOSP.USERID & "'                 " & vbCr
                SQL = SQL & " WHERE hos_org_no = '" & gHOSP.HOSPCD & "'                 " & vbCr
                SQL = SQL & "   AND smp_no     = '" & strBarcode & "'                   " & vbCr
                SQL = SQL & "   AND cd         = '" & strEqpcd & "'                     " & vbCr
                SQL = SQL & "   AND prcp_seq   = '" & mGetP(strSubCode, 1, "|") & "'    " & vbCr
                SQL = SQL & "   AND exam_seq   = '" & mGetP(strSubCode, 2, "|") & "'    " & vbCr
                SQL = SQL & "   AND rept_seq   = '" & mGetP(strSubCode, 3, "|") & "'    " & vbCr
                                    
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                        
                '-- 상태변경1
                SQL = ""
                SQL = SQL & "UPDATE scrprexh                                        " & vbCr
                SQL = SQL & "   SET smp_stus   = decode(smp_stus, '0','1',smp_stus) " & vbCr
                SQL = SQL & " WHERE hos_org_no = '" & gHOSP.HOSPCD & "'             " & vbCr
                SQL = SQL & "   AND smp_no     = '" & strBarcode & "'               " & vbCr
                SQL = SQL & "   AND cd         = '" & strTestCd & "'                " & vbCr
                SQL = SQL & "   AND smp_stus   <> '3'"
                        
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
                '-- 상태변경2
                SQL = ""
                SQL = SQL & "UPDATE mosxpslh                                " & vbCr
                SQL = SQL & "   SET prcp_stus_cd = '4'                      " & vbCr
                SQL = SQL & " WHERE hos_org_no   = '" & gHOSP.HOSPCD & "'   " & vbCr
                SQL = SQL & "   AND smp_no       = '" & strBarcode & "'           " & vbCr
                SQL = SQL & "   AND prcp_stus_cd <> '6'                     " & vbCr
                SQL = SQL & "   AND prcp_typ_cd  IN ('O','C')               " & vbCr
                        
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                        
            End If
        Next intRow
        
        SaveTransData_GINUS = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_GINUS = -1
    
End Function

Function SaveTransData_JWINFO(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_JWINFO = -1
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        intRow = 0
        strJudge = ""
        blnSave = False
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strOrdCd = Trim(GetText(.vasTemp, intRow, 2))       '처방코드
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
                '-- 결과저장
                      SQL = "Update SLA_LabResult                       " & vbCr
                SQL = SQL & "   Set Result      = '" & sResult & "'     " & vbCr
                SQL = SQL & "      ,NormalFlag  = '0'                   " & vbCr
                SQL = SQL & "      ,PanicFlag   = '0'                   " & vbCr
                SQL = SQL & "      ,DeltaFlag   = '0'                   " & vbCr
                SQL = SQL & "      ,TransFlag   = '1'                   " & vbCr
                SQL = SQL & "      ,ResultID    = '" & gHOSP.USERID & "'" & vbCr
                SQL = SQL & "      ,ResultDate  = '" & Trim(Format(Now, "yyyy-mm-dd")) & "'" & vbCr
                SQL = SQL & "      ,ResultTime  = '" & Trim(Format(Time, "HH:MM:SS")) & "'" & vbCr
                SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "'  " & vbCr
                SQL = SQL & "   AND OrderCode = '" & strOrdCd & "'      " & vbCr
                SQL = SQL & "   And LabCode = '" & strTestCd & "'       " & vbCr
                SQL = SQL & "   And TRANSFLAG < '2'                     " & vbCr
                'SQL = SQL & "   AND ReceiptNo = '" & strChartNo & "' "& vbCr
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
                '-- 상태변경
                      SQL = "Update SLA_LabMaster                       " & vbCr
                SQL = SQL & "   Set JStatus = '2'                       " & vbCr
                SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "'  " & vbCr
                SQL = SQL & "   AND OrderCode   = '" & strOrdCd & "'    " & vbCr
                SQL = SQL & "   And JStatus < '3'                       " & vbCr
                    
                Call SetSQLData("상태변경", SQL, "A")
                AdoCn.Execute SQL
                
                SaveTransData_JWINFO = 1
            
            End If
        Next intRow
                       
        SaveTransData_JWINFO = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_JWINFO = -1
    
End Function

Function SaveTransData_ILSIN(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strPatNm        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
    Dim strWorkNum      As String
    Dim strDocID        As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_ILSIN = -1
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatNm = Trim(GetText(.spdOrder, argSpcRow, colPNAME))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strWorkNum = strPatID
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strDocID = strIO
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        intRow = 0
        strJudge = ""
        blnSave = False
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatNm) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
        SQL = ""
        SQL = SQL & "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        'SQL = SQL & "SELECT EQUIPCODE,EXAMCODE,RESULT,EQUIPRESULT,REFFLAG,PANICVALUE,DELTAVALUE,PSEX,SEQNO,PAGE,PID,DISKNO,POSNO,EXAMSUBCODE,INOUT " & vbCrLf
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        SQL = SQL & "   AND SENDFLAG < '2'     "
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strOrdCd = Trim(GetText(.vasTemp, intRow, 2))       '처방코드
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
                '-- 결과저장
                SQL = ""
                SQL = SQL & "Insert Into acrcpnwh                                      " & vbCr
                SQL = SQL & "( CPNWDPCD, CPNWHPNO, CPNWDATE, CPNWSLIP, CPNWITEM        " & vbCr
                SQL = SQL & " , CPNWOITP, CPNWWKNO, CPNWCODE, CPNWSMPL, CPNWMACH        " & vbCr
                SQL = SQL & " , CPNWDISP, CPNWIDNO, CPNWDLTF, CPNWPNCF, CPNWSTAT        " & vbCr
                SQL = SQL & " , CPNWSPCL, CPNWRCLF, CPNWRTCD, CPNWNVAL, CPNWSIGN        " & vbCr
                SQL = SQL & " , CPNWRSLT, CPNWPRNT, CPNWUPDT, CPNWUSER, CPNWUSR2        " & vbCr
                SQL = SQL & " , CPNWSMYR, CPNWSMSN, CPNWSMS1, CPNWSMS2                  " & vbCr
                SQL = SQL & " , CPNWDESC, CPNWTATD, CPNWTATF, CPNWPANL)                 " & vbCr
                SQL = SQL & " Values                                                    " & vbCr
                SQL = SQL & "('LA'" & vbCr
                SQL = SQL & ",'01'" & vbCr
                SQL = SQL & ",'" & Format(Now, "yyyymmdd") & "'" & vbCr
                SQL = SQL & ",'LAE'" & vbCr
                SQL = SQL & ",'00'" & vbCr
                SQL = SQL & ",'A'" & vbCr
                SQL = SQL & ",'" & mOCS.FWorkNo & "'" & vbCr
                SQL = SQL & ",'" & strTestCd & "'" & vbCr
                SQL = SQL & ",'004'" & vbCr
                SQL = SQL & ",'998'" & vbCr
                SQL = SQL & ",'99'" & vbCr
                SQL = SQL & ",'" & mOCS.FPID & "'" & vbCr
                SQL = SQL & ",'N'" & vbCr
                SQL = SQL & ",'N'" & vbCr
                SQL = SQL & ",'1'" & vbCr   '결과보고
                SQL = SQL & ",''" & vbCr
                SQL = SQL & ",'N'" & vbCr
                SQL = SQL & ",''" & vbCr
                SQL = SQL & ",'0'" & vbCr
                SQL = SQL & ",''" & vbCr
                SQL = SQL & ",'" & sResult & "'" & vbCr
                SQL = SQL & ",''" & vbCr
                SQL = SQL & ",'" & Format(Now, "yyyymmddhhmm") & "'" & vbCr
                SQL = SQL & ",'" & mOCS.FDocID & "'" & vbCr
                SQL = SQL & ",''" & vbCr
                SQL = SQL & ",'" & Mid(strChartNo, 1, 2) & "'" & vbCr
                SQL = SQL & ",'" & Mid(strChartNo, 3, 6) & "'" & vbCr
                SQL = SQL & ",1" & vbCr
                SQL = SQL & ",1" & vbCr
                SQL = SQL & ",''" & vbCr
                SQL = SQL & ",''" & vbCr
                SQL = SQL & ",''" & vbCr
                SQL = SQL & ",'')" & vbCr
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
                '-- 상태변경
                SQL = ""
                SQL = SQL & "Update ocsipslh Set " & vbCr
                SQL = SQL & " IPSLSTAT = '7' " & vbCr
                SQL = SQL & ",IPSLMTCD = '004' " & vbCr
                SQL = SQL & ",IPSLACDP = 'LA'  " & vbCr
                SQL = SQL & ",IPSLSMDT = '" & Format(Now, "yyyymmddhhmm") & "'" & vbCr
                SQL = SQL & ",IPSLSMYR = '" & Mid(strChartNo, 1, 2) & "'" & vbCr
                SQL = SQL & ",IPSLSMSN = '" & Mid(strChartNo, 3, 6) & "'" & vbCr
                SQL = SQL & ",IPSLSMS1 = 1                        " & vbCr
                SQL = SQL & ",IPSLSMS2 = 1                        " & vbCr
                SQL = SQL & ",IPSLWORK = '" & mOCS.FWorkNo & "'" & vbCr
                If mOCS.FAcDt <> "" Then
                    SQL = SQL & ",IPSLACDT = '" & mOCS.FAcDt & "'" & vbCr
                    SQL = SQL & ",ipslhpdt = '" & mOCS.FAcDt & "'" & vbCr
                Else
                    SQL = SQL & ",IPSLACDT = '" & Format(Now, "yyyymmdd") & "'" & vbCr
                    SQL = SQL & ",ipslhpdt = '" & Format(Now, "yyyymmdd") & "'" & vbCr
                End If
                If mOCS.FAcTm <> "" Then
                    SQL = SQL & ",IPSLACTM = '" & mOCS.FAcTm & "'" & vbCr
                    SQL = SQL & ",ipslhptm = '" & mOCS.FAcTm & "'" & vbCr
                Else
                    SQL = SQL & ",IPSLACTM = '" & Format(Now, "hhmm") & "'" & vbCr
                    SQL = SQL & ",ipslhptm = '" & Format(Now, "hhmm") & "'" & vbCr
                End If
                SQL = SQL & " Where ipslidno = '" & mOCS.FPID & "'" & vbCr
                SQL = SQL & "   And ipslhpno = '01'                 " & vbCr
                SQL = SQL & "   And ipslmddt = '" & mOCS.FMdDt & "'" & vbCr
                SQL = SQL & "   And ipslcode = '" & mOCS.FOrdCd & "'" & vbCr
                SQL = SQL & "   And ipslflag = 'O' " & vbCr
                SQL = SQL & "   And ipslstat = '0' "
            
                    
                Call SetSQLData("상태변경", SQL, "A")
                AdoCn.Execute SQL
                
                SaveTransData_ILSIN = 1
            
            End If
        Next intRow
                       
        SaveTransData_ILSIN = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_ILSIN = -1
    
End Function


Function SaveTransData_JAINCOM(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_JAINCOM = -1
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        intRow = 0
        strJudge = ""
        blnSave = False
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            'strOrdCd = Trim(GetText(.vasTemp, intRow, 2))       '처방코드
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
                '-- 결과저장
                      SQL = "Update SLA_LabResult                       " & vbCr
                SQL = SQL & "   Set Result      = '" & sResult & "'     " & vbCr
                SQL = SQL & "      ,NormalFlag  = '0'                   " & vbCr
                SQL = SQL & "      ,PanicFlag   = '0'                   " & vbCr
                SQL = SQL & "      ,DeltaFlag   = '0'                   " & vbCr
                SQL = SQL & "      ,TransFlag   = '1'                   " & vbCr
                SQL = SQL & "      ,ResultID    = '" & gHOSP.USERID & "'" & vbCr
                SQL = SQL & "      ,ResultDate  = '" & Trim(Format(Now, "yyyy-mm-dd")) & "'" & vbCr
                SQL = SQL & "      ,ResultTime  = '" & Trim(Format(Time, "HH:MM:SS")) & "'" & vbCr
                SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "'  " & vbCr
                SQL = SQL & "   AND OrderCode = '" & strOrdCd & "'      " & vbCr
                SQL = SQL & "   And LabCode = '" & strTestCd & "'       " & vbCr
                SQL = SQL & "   And TRANSFLAG < '2'                     " & vbCr
                
                '검사정보 MASTER Update
                      SQL = "UPDATE JAIN_SCP.SCPRST41                               " & vbCr
                SQL = SQL & "   SET SCP41TSTDAT = '" & Format(Now, "YYYYMMDD") & "' " & vbCr '결과일자 => YYYYMMDD"
                SQL = SQL & "     , SCP41SNDYN  = 'N'                               " & vbCr '고정값 : 'N'
                SQL = SQL & "     , SCP41RSTYN  = 'Y'                               " & vbCr '고정값 : 'Y'
                SQL = SQL & "     , SCP41TSTUID = '" & gHOSP.USERID & "'            " & vbCr '검사자사번
                SQL = SQL & " WHERE SCP41SPMNO2 = '" & strBarcode & "'              " & vbCr '바코드번호
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
                '검사정보 DETAIL UPDATE
                      SQL = "UPDATE JAIN_SCP.SCPRST42                               " & vbCr
                SQL = SQL & "   SET SCP42TSTDAT = '" & Format(Now, "YYYYMMDD") & "' " & vbCr '결과일자 => YYYYMMDD"
                SQL = SQL & "     , SCP42RSTCD  = 'N'                               " & vbCr '결과형식 => 숫자 : 'N', 문자 : 'X', 장문 : 'R'
                SQL = SQL & "     , SCP42RESULT = '" & sResult & "'                 " & vbCr '결과값
                SQL = SQL & " WHERE SCP42SPMNO2 = '" & strBarcode & "'              " & vbCr '바코드번호
                SQL = SQL & "   AND SCP42SUGACD = '" & strTestCd & "'               " & vbCr '수가코드
                    
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
                SaveTransData_JAINCOM = 1
            
            End If
        Next intRow
                       
        SaveTransData_JAINCOM = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_JAINCOM = -1
    
End Function

Function SaveTransData_KCHART(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_KCHART = -1
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        intRow = 0
        strJudge = ""
        blnSave = False
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT,REFJUDGE " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
                Call SetText(.vasTemp, RS_L.Fields("REFJUDGE").Value & "", intRow, 7)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strOrdCd = Trim(GetText(.vasTemp, intRow, 2))       '진료검사ID
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '진료지원ID
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
            strRefVal = Trim(GetText(.vasTemp, intRow, 6))      '판정
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- 결과저장
                'SQL = SQL & "    ,  수정자 = 'IIS', " & vbCr
                      SQL = "Update TB_진료검사                                   " & vbCr
                SQL = SQL & "   Set 검사결과              = '" & sResult & "'     " & vbCr
                SQL = SQL & "     , 하이로우              = '" & strRefVal & "'   " & vbCr
                SQL = SQL & "     , 검사상태              = '2'                   " & vbCr
                SQL = SQL & "     , 연동구분              = '1'                   " & vbCr
                SQL = SQL & "     , 수정일자              = GetDate()             " & vbCr
                SQL = SQL & " Where 진료검사ID            = '" & strOrdCd & "'    " & vbCr
                SQL = SQL & "   And 진료지원ID            = '" & strSubCode & "'  " & vbCr
                SQL = SQL & "   And 검체번호              = '" & strBarcode & "'  " & vbCr
                SQL = SQL & "   And (처방코드 + 서브코드) = '" & strTestCd & "'   " & vbCr
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
                SaveTransData_KCHART = 1
            
            End If
        Next intRow
                       
        SaveTransData_KCHART = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_KCHART = -1
    
End Function

Function SaveTransData_HWASAN(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_HWASAN = -1
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        intRow = 0
        strJudge = ""
        blnSave = False
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
                '-- 결과저장
                'strDate = Format(Now, "yyyymmdd")
                'strTime = Format(Now, "hhmmss")

                SQL = ""
                SQL = SQL & "INSERT INTO TC206 " & vbCr
                SQL = SQL & " (SPCNO, EQUIPCD, TESTCD, TRANSDT, TRANSTM, RSTVAL, SPCDIV) " & vbCr
                SQL = SQL & "  Values " & vbCr
                SQL = SQL & " ('" & strBarcode & "'                    " & vbCr
                SQL = SQL & " ,'" & gHOSP.MACHCD & "'                  " & vbCr
                SQL = SQL & " ,'" & strTestCd & "'                     " & vbCr
                SQL = SQL & " ,'" & Trim(Format(Now, "yyyymmdd")) & "' " & vbCr
                SQL = SQL & " ,'" & Trim(Format(Time, "hhmmss")) & "'  " & vbCr
                SQL = SQL & " ,'" & sResult & "'                       " & vbCr
                SQL = SQL & " ,'')                                     " & vbCr
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
                       
        SaveTransData_HWASAN = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_HWASAN = -1
    
End Function


Function SaveTransData_KOMAIN(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_KOMAIN = -1
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        intRow = 0
        strJudge = ""
        blnSave = False
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- 결과저장
                
                If gHOSP.BARUSE = "Y" Then
                    '바코드 사용
                    SQL = "EXEC AP_INF_BAR_RESULT '" & strBarcode & "', '" & gHOSP.MACHCD & "', '" & mGetP(strTestCd, 1, "/") & "', '" & mGetP(strTestCd, 2, "/") & "', '" & sResult & "'"
                Else
                    '바코드 미사용
'AP_INF_S_UPDATE
'- 검사번호 LID (바코드처럼..)
'- serial
'- rorder
'-   장비코드(XP:101,BS240 : 201)
'- 검사코드
'- sub 코드
'- 결과
  
  
                    SQL = "EXEC AP_INF_S_UPDATE '" & strBarcode & "','" & strPatID & "','" & strSubCode & "','" & gHOSP.MACHCD & "', '" & mGetP(strTestCd, 1, "/") & "', '" & mGetP(strTestCd, 2, "/") & "', '" & sResult & "'"
                    
                    'EXEC AP_INF_S_UPDATE '1','1','11','201', 'B2570', '1', '18.419311'
                End If
                
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                                
                SaveTransData_KOMAIN = 1
            
            End If
        Next intRow
                       
        SaveTransData_KOMAIN = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_KOMAIN = -1
    
End Function

Function SaveTransData_SY(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    Dim strReturn       As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_SY = -1
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        intRow = 0
        strJudge = ""
        blnSave = False
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- 결과저장
                
                '-- 결과업데이트
                SQL = ""
                SQL = SQL & "Exec Interface_SetPatientResult02 "
                SQL = SQL & "  '" & strHospDate & "'"
                SQL = SQL & " , " & strPatID
                SQL = SQL & " ,'" & mGetP(strSubCode, 1, "|") & "'"
                SQL = SQL & " ,'" & mGetP(strSubCode, 2, "|") & "'"
                SQL = SQL & " ,'" & mGetP(strSubCode, 3, "|") & "'"
                SQL = SQL & " ,'" & sResult & "'"
                SQL = SQL & " ,''"
                SQL = SQL & " ,''"
                SQL = SQL & " , 0"
                SQL = SQL & " , 0"
                SQL = SQL & " , 0"
                SQL = SQL & " ,'" & gHOSP.MACHCD & "'"
                SQL = SQL & " ,'" & strReturn & "'"
                
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                                
                SaveTransData_SY = 1
            
            End If
        Next intRow
                       
        SaveTransData_SY = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_SY = -1
    
End Function

Function SaveTransData_KYU(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim RS_S            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    Dim strSubCodes     As String
    Dim strSlip1        As String
    Dim strSlip2        As String
        
    Dim prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim prm3 As New ADODB.Parameter
    Dim prm4 As New ADODB.Parameter
    Dim prm5 As New ADODB.Parameter
    Dim prm6 As New ADODB.Parameter
    Dim prm7 As New ADODB.Parameter
    Dim prm8 As New ADODB.Parameter
        
    Dim intBcNow    As Integer
    Dim intBcFive   As Integer
    Dim intBcAdd    As Integer
    Dim strADT      As String
        
        
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_KYU = -1
        
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strSlip1 = Trim(GetText(.spdOrder, argSpcRow, colRACKNO))
        strSlip2 = Trim(GetText(.spdOrder, argSpcRow, colPOSNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        
        strDate = Format(Now, "yyyy-mm-dd")
        intBcNow = DateDiff("d", "1999-01-01", strDate)
        intBcFive = Mid(strBarcode, 1, 5) '06351
        intBcAdd = intBcFive - intBcNow
        strADT = Format(Now + intBcAdd, "yyyy-mm-dd")
        
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT,EXAMSUBCODE " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            If intRow = 1 Then
                '검체접수 (검체도착)
                Set AdoCmd = New ADODB.Command
                Set AdoCmd.ActiveConnection = AdoCn
                With AdoCmd
                    .CommandTimeout = 15
                    .CommandText = "EXAM_INTERFACE_ARR_U"
                    .CommandType = adCmdStoredProc
                    
                    Set prm1 = .CreateParameter("I_PTNO", adVarChar, adParamInput, 20, strPatID)
                    .Parameters.Append prm1
                    Set prm2 = .CreateParameter("I_JEOBSUDT", adDate, adParamInput, 10, strADT)
                    .Parameters.Append prm2
                    Set prm3 = .CreateParameter("I_SLIPNO1", adInteger, adParamInput, 2, strSlip1)
                    .Parameters.Append prm3
                    Set prm4 = .CreateParameter("I_SLIPNO2", adInteger, adParamInput, 5, strSlip2)
                    .Parameters.Append prm4
                    
                    .Execute
                    
                    Call SetSQLData("검체도착", strPatID & "," & strADT & "," & strSlip1 & "," & strSlip2, "A")
                    Set AdoCmd = Nothing
                End With
            End If
            
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
            'strSubCodes = Trim(GetText(.vasTemp, intRow, 7))      '저장용 코드 : ex) 999|888|777
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- 검사결과업데이트
                      
'EXAM_INTERFACE_U
'      I_PTNO VARCHAR2
'    , I_JEOBSUDT      DATE
'    , I_SLIPNO1       NUMBER
'    , I_SLIPNO2       NUMBER
'    , I_ITEMCD        VARCHAR2
'    , I_RESULT        VARCHAR2
'    , I_JANGBI        NUMBER
'    , I_SABUN         VARCHAR2
'
    
                '-- 검사결과저장 = PG_SLA_INTERFACEMGT.SP_SLA_INTERFACEMGT_U02
                Set AdoCmd = New ADODB.Command
                Set AdoCmd.ActiveConnection = AdoCn
                With AdoCmd
                    .CommandTimeout = 15
                    .CommandText = "EXAM_INTERFACE_U"
                    .CommandType = adCmdStoredProc
                    
                    Set prm1 = .CreateParameter("I_PTNO", adVarChar, adParamInput, 20, strPatID)
                    .Parameters.Append prm1
                    Set prm2 = .CreateParameter("I_JEOBSUDT", adDate, adParamInput, 10, strADT)
                    .Parameters.Append prm2
                    Set prm3 = .CreateParameter("I_SLIPNO1", adInteger, adParamInput, 2, strSlip1)
                    .Parameters.Append prm3
                    Set prm4 = .CreateParameter("I_SLIPNO2", adInteger, adParamInput, 5, strSlip2)
                    .Parameters.Append prm4
                    Set prm5 = .CreateParameter("I_ITEMCD", adVarChar, adParamInput, 20, strTestCd)
                    .Parameters.Append prm5
                    Set prm6 = .CreateParameter("I_RESULT", adVarChar, adParamInput, 50, sResult)
                    .Parameters.Append prm6
                    Set prm7 = .CreateParameter("I_JANGBI", adInteger, adParamInput, 10, gHOSP.MACHCD)
                    .Parameters.Append prm7
                    Set prm8 = .CreateParameter("I_SABUN", adVarChar, adParamInputOutput, 10, gHOSP.USERID)
                    .Parameters.Append prm8
                    
                    .Execute
                    
                    Call SetSQLData("결과저장", strPatID & "," & strADT & "," & strSlip1 & "," & strSlip2 & "," & strTestCd & "," & sResult & "," & gHOSP.MACHCD & "," & gHOSP.USERID, "A")
                    
                    Set AdoCmd = Nothing
                    
                End With
    
    
            End If
        Next intRow
               
        SaveTransData_KYU = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_KYU = -1
    
End Function

Function SaveTransData_BIT(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim blnSave         As Boolean
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strSubCode      As String
    Dim strEqpcd        As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strRefVal       As String
    Dim strJudge        As String
    
On Error GoTo ErrHandle
        
    SaveTransData_BIT = -1
    intRow = 0
    strJudge = ""
    blnSave = False

    With frmMain
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT,EXAMSUBCODE " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
        strSubCode = ""
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사SUB코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" And strSubCode <> "" Then
                '-- 결과저장
                      SQL = "Update ResInf " & vbCr
                SQL = SQL & "   Set ResRltVal = '" & sResult & "'" & vbCr   '검사결과
                SQL = SQL & "     , ResRepTyp = 'I'              " & vbCr
                SQL = SQL & " Where ResOcmNum = '" & Space(10 - Len(strPatID)) & strPatID & "'  " & vbCr            '10
                SQL = SQL & "   And ResLabCod = '" & strTestCd & "'                             " & vbCr   '검사코드
                SQL = SQL & "   And ResOdrSeq = '" & mGetP(strSubCode, 1, "|") & "'             " & vbCr
                SQL = SQL & "   And ResSeq    = '" & mGetP(strSubCode, 2, "|") & "'             " & vbCr
                SQL = SQL & "   And ResSubSeq = '" & mGetP(strSubCode, 3, "|") & "'             " & vbCr
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
               
        SaveTransData_BIT = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_BIT = -1
    
End Function

Function SaveTransData_BIT70(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim blnSave         As Boolean
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    
    Dim strOrdCd        As String
    Dim strTestCd       As String
    Dim strSubCode      As String
    Dim strEqpcd        As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strRefVal       As String
    Dim strJudge        As String
    
    Dim strDate         As String
    Dim strTime         As String
    
On Error GoTo ErrHandle
        
    SaveTransData_BIT70 = -1
    intRow = 0
    strJudge = ""
    blnSave = False

    With frmMain
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        
        If Trim(strChartNo) = "" Then
            Exit Function
        End If
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT,EXAMSUBCODE " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
        strSubCode = ""
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사SUB코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" And strSubCode <> "" Then
                '-- 결과저장
                SQL = ""
                SQL = SQL & "UPDATE ME_LABDAT " & vbCr
                SQL = SQL & "   Set LABRESULT = '" & sResult & "'       " & vbCr    '검사결과
                SQL = SQL & "     , LABENDDEP = '2'                     " & vbCr    '처리상태      2:접수, 3:결과입력
                SQL = SQL & "     , LABRSTDTE = '" & strDate & "'       " & vbCr    '결과입력일자  YYYY-MM-DD
                SQL = SQL & "     , LABRSTTIM = '" & strTime & "'       " & vbCr    '결과입력시간  hh:mm:ss
                SQL = SQL & "     , LABRSTUID = '" & gHOSP.USERID & "'  " & vbCr    '결과입력ID
                SQL = SQL & "     , LABRSTCOM = '" & gHOSP.MACHNM & "'  " & vbCr    '결과입력컴퓨터명
                SQL = SQL & " WHERE LABATTEND = '" & strPatID & "'      " & vbCr    '내원번호
                SQL = SQL & "   And LABODRDTE = '" & strHospDate & "'   " & vbCr    '처방일자
                SQL = SQL & "   And LABODRCOD = '" & strTestCd & "'     " & vbCr    '검사코드
                SQL = SQL & "   And LABODRSTP = '" & strSubCode & "'    " & vbCr    '검사일련번호
                'SQL = SQL & "   And LABBARCOD = '" & lsID & "'" & vbCr  '바코드
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
               
        SaveTransData_BIT70 = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_BIT70 = -1
    
End Function

Function SaveTransData_BIGUBCARE(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_BIGUBCARE = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        'strDate = Format(Now, "yyyy-mm-dd")
        'strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strChartNo) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "'" & vbCr                        '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사코드SUB
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                
                '-- 서버저장
                SQL = ""
                SQL = SQL & " UPDATE IntRst SET"
                SQL = SQL & "      IntRstVal = '" & sResult & "'" & vbCr
                SQL = SQL & "    , IntRstDte = '" & Format(Now, "yyyymmddhhmm") & "'" & vbCr
                SQL = SQL & " WHERE IntLabNum = '" & strBarcode & "'" & vbCr
                SQL = SQL & "   AND IntLabCod + cast(IntLabseq as varchar(3)) = '" & strTestCd & "'" & vbCr
                'SQL = SQL & "   And IntLabSeq = " & strSubCode & vbCr
                SQL = SQL & "   And IntChtNum = '" & strChartNo & "'" & vbCr
                SQL = SQL & "   And IntOdrDte = '" & strHospDate & "'" & vbCr
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
        
        SaveTransData_BIGUBCARE = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_BIGUBCARE = -1
    
End Function

Function SaveTransData_AMIS(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_AMIS = -1
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        intRow = 0
        strJudge = ""
        blnSave = False
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strOrdCd = Trim(GetText(.vasTemp, intRow, 2))       '처방코드
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
                '-- 결과저장
'                SQL = SQL & "      , printflag     = '5' "
                'SQL = SQL & "      , ANALYZERCODE  = '" & gHOSP.MACHCD & "'"  'GEMINI = 43
                SQL = "                      "
                SQL = SQL & "Update RESULTOFNUM                                    " & vbCr
                SQL = SQL & "   Set RESULTINDATE   = to_char(sysdate,'yyyymmdd')   " & vbCr
                SQL = SQL & "     , RESULTINTIME   = to_char(sysdate,'HH24MI')     " & vbCr
                SQL = SQL & "     , RESULTINID     = '" & gHOSP.USERID & "'        " & vbCr
                SQL = SQL & "     , RESULTFLAG     = '1'                           " & vbCr
                SQL = SQL & "     , TEXTRESULTVAL  = '" & sResult & "'             " & vbCr
                '-- 결과가 수치형이면
                If IsNumeric(sResult) Then
                    SQL = SQL & "     , NUMRESULTVAL = '" & sResult & "'           " & vbCr
                End If
                SQL = SQL & " Where SPCMNO         = '" & strBarcode & "'       " & vbCr
                SQL = SQL & "   And ORDERCODE      = '" & strOrdCd & "'         " & vbCr
                SQL = SQL & "   And RESULTITEMCODE = '" & strTestCd & "'        " & vbCr
                SQL = SQL & "   And RESULTFLAG < '3'                            " & vbCr
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
                '-- 상태변경
                SQL = "                "
                SQL = SQL & "Update REGISTINFOS                         " & vbCr
                SQL = SQL & "   Set RESULTSTATE  = '1'                  " & vbCr
                SQL = SQL & "      ,RsvAcptState = '4'                  " & vbCr
                SQL = SQL & " Where SPCMNO       = '" & strBarcode & "' " & vbCr
                SQL = SQL & "   AND ORDERCODE    = '" & strOrdCd & "'   " & vbCr
                SQL = SQL & "   AND CLAS         = 4                    " & vbCr
                SQL = SQL & "   AND RESULTSTATE < '4'                   " & vbCr
                    
                Call SetSQLData("상태변경", SQL, "A")
                AdoCn.Execute SQL
                
            
            End If
        Next intRow
                       
        SaveTransData_AMIS = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_AMIS = -1
    
    
End Function


Function SaveTransData_MSINFOTEC(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_MSINFOTEC = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        mPatient.AGE = Trim(GetText(.spdOrder, argSpcRow, colPAGE))
        mPatient.SEX = Trim(GetText(.spdOrder, argSpcRow, colPSEX))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사SUB코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- H/L 판정
                'strJudge = getMSINFOTECJudge(strTestCd, sResult)
                strJudge = SetJudge(sResult, strTestCd)
                
                '-- 서버저장
                SQL = ""
                SQL = SQL & " Update emr.LRESULT                    " & vbCr
                SQL = SQL & "   Set RSFL = 'Y'                      " & vbCr
                SQL = SQL & "     , RSLT = '" & sResult & "'        " & vbCr
                SQL = SQL & "     , HLFL = '" & strJudge & "'       " & vbCr
                SQL = SQL & "     , RSDT = SYSDATE                  " & vbCr
                SQL = SQL & "     , RSID = '" & gHOSP.USERID & "'   " & vbCr
                SQL = SQL & " Where SPNO = '" & strBarcode & "'     " & vbCr
                SQL = SQL & "   And PAID = '" & strPatID & "'       " & vbCr
                SQL = SQL & "   And ORCD = '" & strTestCd & "'      " & vbCr
                SQL = SQL & "   And ORQN = " & strSubCode & vbCr
                SQL = SQL & "   And OKFL <> 'Y'                     " & vbCr   '-- 결과확정유무
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
        
        SaveTransData_MSINFOTEC = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_MSINFOTEC = -1
    
End Function

Function SaveTransData_MEDICHART(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_MEDICHART = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        mPatient.AGE = Trim(GetText(.spdOrder, argSpcRow, colPAGE))
        mPatient.SEX = Trim(GetText(.spdOrder, argSpcRow, colPSEX))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사SUB코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- H/L 판정
                strJudge = SetJudge(sResult, strTestCd)
                
                If strJudge = "L" Then
                    strJudge = "2"
                ElseIf strJudge = "H" Then
                    strJudge = "1"
                Else
                    strJudge = "0"
                End If
                
                '-- 서버저장
                SQL = ""
                SQL = SQL & "Update TB_검사항목 "
                SQL = SQL & "   Set 검사결과        = '" & sResult & "'"
                SQL = SQL & "     , 진료지원상태    = 5 " & vbCr      '1 : 처치중, 5 : 완료
                SQL = SQL & "     , 하이로우        = '" & strJudge & "'" & vbCr
                'SQL = SQL & " Where 진료년   = '" & strYear & "'" & vbCr
                'SQL = SQL & "   and 진료월   = '" & strMonth & "'" & vbCr
                'SQL = SQL & "   and 진료일   = '" & strDay & "'" & vbCr
                SQL = SQL & " Where (진료년+진료월+진료일) = '" & strHospDate & "'" & vbCr
                SQL = SQL & "   and 챠트번호 = '" & strChartNo & "'" & vbCr
                SQL = SQL & "   And (처방코드+서브코드) = '" & strTestCd & "'" & vbCr
                'SQL = SQL & "   And 서브코드 = '" & strSubCD & "'" & vbCr
                
                
                'strHospDate
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
        
        SaveTransData_MEDICHART = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_MEDICHART = -1
    
End Function

Function SaveTransData_MEDIIT(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_MEDIIT = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        mPatient.AGE = Trim(GetText(.spdOrder, argSpcRow, colPAGE))
        mPatient.SEX = Trim(GetText(.spdOrder, argSpcRow, colPSEX))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사SUB코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- H/L 판정
                strJudge = SetJudge(sResult, strTestCd)
                
                '-- 서버저장
                SQL = ""
                SQL = SQL & "Update trures                                      " & vbCr
                SQL = SQL & "   Set RESULT_VALUE    = '" & sResult & "'         " & vbCr
                SQL = SQL & "     , RESULT_DECISION = '" & strJudge & "'        " & vbCr
                SQL = SQL & " WHERE request_date    = '" & strHospDate & "'     " & vbCr
                SQL = SQL & "   And exam_no         = '" & strPatID & "'        " & vbCr
                SQL = SQL & "   And exam_code       = '" & strTestCd & "'       " & vbCr
                SQL = SQL & "   And (RESULT_VALUE = '' or RESULT_VALUE IS NULL) " & vbCr
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
        
        SaveTransData_MEDIIT = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_MEDIIT = -1
    
End Function

Function SaveTransData_MEDITOLISS(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_MEDITOLISS = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
                
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사SUB코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- H/L 판정
                mResult.RsltDate = strHospDate
                mResult.BarNo = strBarcode
                
                strJudge = SetJudge(sResult, strTestCd)
                
                '-- 서버저장
                SQL = ""
                SQL = SQL & "Update MEDITOLISS..TOTRES                      " & vbCr
                SQL = SQL & "   Set RESULT_VALUE    = '" & sResult & "'     " & vbCr
                SQL = SQL & "     , RESULT_DECISION = '" & strJudge & "'    " & vbCr
                SQL = SQL & " WHERE REQUEST_DATE    = '" & strHospDate & "' " & vbCr
                SQL = SQL & "   AND EXAM_NO         = '" & strBarcode & "'  " & vbCr
                SQL = SQL & "   AND EXAM_CODE       = '" & strTestCd & "'   " & vbCr
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
        
        SaveTransData_MEDITOLISS = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_MEDITOLISS = -1
    
End Function

Function SaveTransData_MOD(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_MOD = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        mPatient.AGE = Trim(GetText(.spdOrder, argSpcRow, colPAGE))
        mPatient.SEX = Trim(GetText(.spdOrder, argSpcRow, colPSEX))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strOrdCd = Trim(GetText(.vasTemp, intRow, 2))       '처방번호
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사SUB코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- H/L 판정
                'strJudge = SetJudge(sResult, strTestCd)
                
                '-- 서버저장
                SQL = ""
                SQL = SQL & "Update EXAMRES "
                SQL = SQL & "   Set RESULT      = '" & sResult & "'     " & vbCr
                SQL = SQL & " Where PID         = '" & strPatID & "'    " & vbCr
                SQL = SQL & "   and SPECIMENID  = '" & strBarcode & "'  " & vbCr
                SQL = SQL & "   and RECENO      = '" & strOrdCd & "'    " & vbCr
                SQL = SQL & "   and SEQNO       = '" & strSubCode & "'  " & vbCr
                SQL = SQL & "   and EXAMCODE    = '" & strTestCd & "'   " & vbCr
                SQL = SQL & "   And (EXAMEND    = '' Or EXAMEND IS NULL) "
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
        
        SaveTransData_MOD = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_MOD = -1
    
End Function



'>> RESA
'RESA_MEDM_ID    NVARCHAR    10          병원 인덱스     NOT NULL
'RESA_KIND       SMALLINT                0외래 1입원     NOT NULL
'RESA_KEY        NVARCHAR    20          날짜(6)+타이머(7)+슬립구분(2)+행(3)+seq같은거갯수(2) 예) 16021962420800200201       NOT NULL
'RESA_SEQ        INT                     resa 시퀀스     NOT NULL
'RESA_CNT        SMALLINT                resa count      NOT NULL
'RESA_CHAM_INDEX NVARCHAR    10          환자 인덱스     NULL
'RESA_GWAM_ID    NVARCHAR    3           진료과      NULL
'RESA_DATE       NVARCHAR    8           날짜        NULL
'RESA_DEPT_ID    NVARCHAR    20          전달부서        NULL
'RESA_SLIP_ID    NVARCHAR    30          슬립아이디      NULL
'RESA_CODE       NVARCHAR    20          검사코드        NULL
'RESA_TIME       NVARCHAR    4           ST (채취시간)       NULL
'RESA_FRESULT    NVARCHAR    50          임상참고치 (최소)       NULL
'RESA_TRESULT    NVARCHAR    50          임상참고치 (최대)       NULL
'RESA_RESULT     NVARCHAR    50          결과        NULL
Function SaveTransData_NEOSOFT(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_NEOSOFT = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사SUB코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- 서버저장
                SQL = ""
                SQL = SQL & "Update E_ORDER..RESA" & Format(Now, "yyyy") & vbCr
                SQL = SQL & "   Set RESA_RESULT     = '" & sResult & "'         " & vbCr '검사결과
                'SQL = SQL & "      ,RESA_BIGO5      = '1'                       " & vbCr '검사결과가 인터페이스에서 저장되면 챠트번호란 우측 상단에 빨간 점이 나타나게 하는것..
                SQL = SQL & " Where RESA_CHAM_INDEX = '" & Val(strBarcode) & "' " & vbCr '검체번호
                SQL = SQL & "   and RESA_DATE       = '" & strHospDate & "'     " & vbCr
                SQL = SQL & "   and RESA_CODE       = '" & strTestCd & "'       " & vbCr
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
        
        SaveTransData_NEOSOFT = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_NEOSOFT = -1
    
End Function

Function SaveTransData_ONITGUM(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_ONITGUM = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        strDate = Format(Now, "yyyymmdd")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
                
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사SUB코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- 서버저장
                SQL = ""
                SQL = SQL & "Update ONIT..GUMJIN_INTERFACE                  " & vbCr
                SQL = SQL & "   Set RESULT          = '" & sResult & "'     " & vbCr
                SQL = SQL & "     , ACT_RETURN_DATE = '" & strDate & "'     " & vbCr
                SQL = SQL & " Where PER_GUMJIN_DATE = '" & strHospDate & "' " & vbCr
                SQL = SQL & "   And PER_GUM_NUM = " & Val(strBarcode) & vbCr
                SQL = SQL & "   And EDPSCODE = '" & strTestCd & "'          " & vbCr
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
        
        SaveTransData_ONITGUM = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_ONITGUM = -1
    
End Function


Function SaveTransData_ONITEMR(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_ONITEMR = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        strDate = Format(Now, "yyyymmdd")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
                
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사SUB코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- 서버저장
                SQL = ""
                SQL = SQL & "UPDATE " & gSQLDB.DB & "..JUN370_RESULTTB" & vbCr
                SQL = SQL & "   SET RESULT      = '" & sResult & "'   " & vbCr
                SQL = SQL & " WHERE WAITSEQNO   = '" & strBarcode & "'" & vbCr
                SQL = SQL & "   AND MAP2SEQNO   = '" & strTestCd & "' " & vbCr
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
        
        SaveTransData_ONITEMR = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_ONITEMR = -1
    
End Function

Function SaveTransData_PLIS(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_PLIS = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        strDate = Format(Now, "yyyymmdd")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
                
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strOrdCd = Trim(GetText(.vasTemp, intRow, 2))       'WA
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     'ACCSEQ
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                If mResult.Kind = "QC" Then
                    '-- QC 결과저장
                    SQL = ""
                    SQL = SQL & "UPDATE plis..s2lab026                  " & vbCr
                    SQL = SQL & "   SET eqpcd   = '" & gHOSP.MACHCD & "'" & vbCr
                    If IsNumeric(sResult) And InStr(sResult, "+") <= 0 And InStr(sResult, "-") <= 0 Then
                        SQL = SQL & "  , rstval  = '" & sResult & "'    " & vbCr
                    End If
                    SQL = SQL & "  , rstcd       = '" & sResult & "'    " & vbCr
                    SQL = SQL & "  , rsttype     = 'N'                  " & vbCr
                    SQL = SQL & " WHERE workarea = '" & strOrdCd & "'   " & vbCr
                    SQL = SQL & "   AND accdt    = '" & strHospDate & "'" & vbCr
                    SQL = SQL & "   AND accseq   = " & strSubCode & vbCr
                    SQL = SQL & "   AND testcd   = '" & strTestCd & "'  " & vbCr
                    SQL = SQL & "   And (vfydt IS NULL OR vfydt= '')    " & vbCr
                Else
                    '-- 결과저장
                    SQL = ""
                    SQL = SQL & " UPDATE plis..s2lab302 " & vbCr
                    SQL = SQL & "    SET eqpcd   = '" & gHOSP.MACHCD & "'" & vbCr
                    If IsNumeric(sResult) And InStr(sResult, "+") <= 0 And InStr(sResult, "-") <= 0 Then
                        SQL = SQL & "  , rstval  = '" & sResult & "'" & vbCr
                    End If
                    SQL = SQL & "  , rstcd       = '" & sResult & "'    " & vbCr
                    SQL = SQL & "  , rsttype     = 'N'                  " & vbCr
                    SQL = SQL & " WHERE workarea = '" & strOrdCd & "'   " & vbCr
                    SQL = SQL & "   AND accdt    = '" & strHospDate & "'" & vbCr
                    SQL = SQL & "   AND accseq   = " & strSubCode & vbCr
                    SQL = SQL & "   AND testcd   = '" & strTestCd & "'  " & vbCr
                    SQL = SQL & "   And (vfydt IS NULL OR vfydt= '')    " & vbCr
                End If
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
        
        SaveTransData_PLIS = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_PLIS = -1
    
End Function


Function SaveTransData_TWIN(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_TWIN = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
            strOrdCd = Trim(GetText(.vasTemp, intRow, 2))       'MASTER코드
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사SUB코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '########  투윈 버전에 따라 두가지 결과저장 형태가 있슴
                '-- 서버저장
'                SQL = ""
'                SQL = SQL & "Update twexam_general_sub " & vbCr
'                SQL = SQL & "   Set RESULT1  = '" & sResult & "'" & vbCrLf '검사결과
'                SQL = SQL & " Where PTNO     = '" & strBarcode & "'" & vbCrLf        '검체번호
'                SQL = SQL & "   and jeobsudt = to_date('" & strHospDate & "', 'yyyy/mm/dd hh24/mi/ss') " & vbCrLf
'                SQL = SQL & "   and itemcd = '" & strTestCd & "'"
'                SQL = SQL & "   and verify <> 'Y'"
'
'                Call SetSQLData("결과저장", SQL, "A")
'                AdoCn.Execute SQL
                
                '-- 서버저장
                SQL = ""
                SQL = SQL & "Update TW_HSP_OCS.TWEXAM_RESULTC           " & vbCr
                SQL = SQL & "   Set STATUS      = '4'                   " & vbCr  '검사상태
                SQL = SQL & "     , RESULT      = '" & sResult & "'     " & vbCr  '검사결과
                SQL = SQL & "     , RESULTDATE  = TRUNC(SYSDATE)        " & vbCr  '검사전송시간
                SQL = SQL & " Where SPECNO      = '" & strBarcode & "'  " & vbCr  '검체번호
                SQL = SQL & "   And MASTERCODE  = '" & strOrdCd & "'    " & vbCr  '마스터코드
                SQL = SQL & "   And SUBCODE     = '" & strTestCd & "'   " & vbCr  '검사코드
                SQL = SQL & "   And STATUS      <= '3'                  " & vbCr  '검사상태(=검체접수)
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
                '-- 상태업데이트
                SQL = ""
                SQL = SQL & "Update TW_HSP_OCS.TWEXAM_SPECMST           " & vbCr
                SQL = SQL & "   Set STATUS     = '4'                    " & vbCr '검사상태 [결과등록(4:부분전송)]
                SQL = SQL & "     , RESULTDATE = TRUNC(SYSDATE)         " & vbCr
                SQL = SQL & " Where SPECNO     = '" & strBarcode & "'   " & vbCr '검체번호
                SQL = SQL & "   And STATUS     <= '3'                   " & vbCr '검사상태 [3:검체접수]
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
        
        SaveTransData_TWIN = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_TWIN = -1
    
End Function

Function SaveTransData_UBCARE(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    Dim strXMLBody      As String
    
    Dim strPName    As String
    Dim strPJumin    As String
    Dim strExamNo    As String
    Dim strSpcType    As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_UBCARE = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strXMLBody = ""
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT,EXAMNO,SAMPLETYPE,REFVALUE " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
                Call SetText(.vasTemp, RS_L.Fields("EXAMNO").Value & "", intRow, 7)
                Call SetText(.vasTemp, RS_L.Fields("SAMPLETYPE").Value & "", intRow, 8)
                Call SetText(.vasTemp, RS_L.Fields("REFVALUE").Value & "", intRow, 9)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strEqpcd = Trim(GetText(.vasTemp, intRow, 1))
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
            strExamNo = Trim(GetText(.vasTemp, intRow, 7))
            strSpcType = Trim(GetText(.vasTemp, intRow, 8))
            strRefVal = Trim(GetText(.vasTemp, intRow, 9))
                  
            If strIO = "입원" Then
                strIO = "1"
            ElseIf strIO = "외래" Then
                strIO = "0"
            End If
            
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                strXMLBody = strXMLBody & "<검사>"
                strXMLBody = strXMLBody & "<업체>" & gHOSP.LABCD & "</업체>"
                strXMLBody = strXMLBody & "<요양기관번호>" & gHOSP.HOSPCD & "</요양기관번호>"
                strXMLBody = strXMLBody & "<차트번호>" & strChartNo & "</차트번호>"
                strXMLBody = strXMLBody & "<수진자명>" & strPName & "</수진자명>"
                strXMLBody = strXMLBody & "<주민등록번호>" & strPJumin & "</주민등록번호>"
                strXMLBody = strXMLBody & "<내원번호>" & strPatID & "</내원번호>"
                strXMLBody = strXMLBody & "<의뢰일>" & strHospDate & "</의뢰일>"
                strXMLBody = strXMLBody & "<검사번호>" & strExamNo & "</검사번호>"
                strXMLBody = strXMLBody & "<검사ID>" & strTestCd & "</검사ID>"
                strXMLBody = strXMLBody & "<업체검사ID>" & strEqpcd & "</업체검사ID>"
                strXMLBody = strXMLBody & "<검체>" & strSpcType & "</검체>"
                strXMLBody = strXMLBody & "<결과치>" & sResult & "</결과치>"
                strXMLBody = strXMLBody & "<참조치>" & strRefVal & "</참조치>"
                strXMLBody = strXMLBody & "<소견></소견>"
                strXMLBody = strXMLBody & "<결과일>" & strExamDate & "</결과일>"
                strXMLBody = strXMLBody & "<입원외래구분>" & strIO & "</입원외래구분>"
                strXMLBody = strXMLBody & "</검사>"
            End If
        Next intRow
        
        If strXMLBody <> "" Then
            Call SaveXMLFile_UBCARE(strXMLBody)
            strXMLBody = ""
            SaveTransData_UBCARE = 1
        End If
                
    End With

Exit Function

ErrHandle:
    SaveTransData_UBCARE = -1
    
End Function

Public Sub SaveXMLFile_UBCARE(strXMLBody As String, Optional argFlag As Integer = 0)
    Dim FilNum, FilNum1
    Dim FindFile As String
    Dim TxtString1 As String
    Dim AllString1 As String
    Dim i As Long
    
    Dim strXmlLine      As String
    Dim strXmlAll       As String
    Dim strXmlAllBody   As String
    
    Dim strXml          As String
    Dim strXmlHeader    As String
    Dim strXmlTail      As String
    
    strXmlAll = ""
    strXmlAllBody = ""
    
    strXml = ""
    strXmlHeader = ""
    strXmlTail = ""
    
    strXmlHeader = "<?xml version=""1.0"" encoding=""euc-kr""?>" & vbCrLf & _
                   "<?xml-stylesheet type=""text/xsl"" href=C:\UBCare\SINAI\IF\Form\ExamIF_Form_05.xsl""?>" & vbCrLf & _
                   "<UBCare검사정보>"
    
    strXmlTail = "</UBCare검사정보>"
    
    If gHOSP.PARTCD = "C" Then
        FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_Out1.xml")
        
        If FindFile <> "" Then
            FilNum1 = FreeFile
            Open "C:\UBCare\SINAI\IF\ExamIF_Out1.xml" For Input As FilNum1
            
            Do While Not EOF(FilNum1)
                Input #FilNum1, strXmlLine
                strXmlAll = strXmlAll & strXmlLine
            Loop
    
            Close #FilNum1
            i = InStr(1, strXmlAll, "</UBCare검사정보>")
            strXmlAllBody = Mid(strXmlAll, 1, i - 1)
            strXml = strXmlAllBody & strXMLBody & strXmlTail
            Kill "C:\UBCare\SINAI\IF\ExamIF_Out1.xml"
        Else
            strXml = strXmlHeader & strXMLBody & strXmlTail
        End If
        
        FilNum = FreeFile
        
        If argFlag = 0 Then
            Open "C:\UBCare\SINAI\IF\ExamIF_Out1.xml" For Output As FilNum
        Else
            Open "C:\UBCare\SINAI\IF\ExamIF_Out1.xml" For Append As FilNum
        End If
    Else
        FindFile = Dir("C:\UBCare\SINAI\IF\ExamIF_Out2.xml")
        
        If FindFile <> "" Then
            FilNum1 = FreeFile
            Open "C:\UBCare\SINAI\IF\ExamIF_Out2.xml" For Input As FilNum1
            
            Do While Not EOF(FilNum1)
                Input #FilNum1, strXmlLine
                strXmlAll = strXmlAll & strXmlLine
            Loop
    
            Close #FilNum1
            i = InStr(1, strXmlAll, "</UBCare검사정보>")
            strXmlAllBody = Mid(strXmlAll, 1, i - 1)
            strXml = strXmlAllBody & strXMLBody & strXmlTail
            Kill "C:\UBCare\SINAI\IF\ExamIF_Out2.xml"
        Else
            strXml = strXmlHeader & strXMLBody & strXmlTail
        End If
        
        FilNum = FreeFile
        
        If argFlag = 0 Then
            Open "C:\UBCare\SINAI\IF\ExamIF_Out2.xml" For Output As FilNum
        Else
            Open "C:\UBCare\SINAI\IF\ExamIF_Out2.xml" For Append As FilNum
        End If

    End If
    
    Print #FilNum, strXml
    Close FilNum
    
    Call SetSQLData("결과저장", strXml, "A")
    
    
End Sub


Function SaveTransData_NAVY(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_NAVY = -1
        'intRow = 0
        'strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
'        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
'        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        
'        strDate = Format(Now, "yyyy-mm-dd")
'        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Call SetSQLData("로컬결과조회", SQL)
        
        Set RS_L = New ADODB.Recordset
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
                '-- 서버저장
                SQL = ""
                SQL = SQL & " Update SLXWORKT"
                SQL = SQL & "   Set PROCSTAT = 'E' " & vbCr
                SQL = SQL & "     , RSLTTEXT = '" & sResult & "'" & vbCr
                SQL = SQL & " Where HOSPID = '" & gHOSP.HOSPCD & "'" & vbCr
                SQL = SQL & "   And PATNO = '" & strPatID & "'" & vbCr
                SQL = SQL & "   And WORKCODE = '" & strChartNo & "'" & vbCr
                SQL = SQL & "   And EXAMCODE = '" & strTestCd & "'" & vbCr
                'SQL = SQL & "   And SPCID = '" & strBarcode & "'" & vbCr
                'SQL = SQL & "   And ORDSEQNO = '" & strORSQ & "'"
                                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
        
        SaveTransData_NAVY = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_NAVY = -1
    
End Function

'-- 검사결과 서버저장
Function SaveTransData(ByVal argSpcRow As Integer) As Integer
    
    SaveTransData = -1
    
    Select Case gEMR
        Case "AMIS"
            SaveTransData = SaveTransData_AMIS(argSpcRow)
        
        Case "BIGUBCARE"
            SaveTransData = SaveTransData_BIGUBCARE(argSpcRow)
        
        Case "BIT"
            SaveTransData = SaveTransData_BIT(argSpcRow)

        Case "BIT70"
            SaveTransData = SaveTransData_BIT70(argSpcRow)
        
        Case "EMEDI"
            SaveTransData = SaveTransData_AMIS(argSpcRow)
        
        Case "EONM"
            SaveTransData = SaveTransData_EONM(argSpcRow)
            
        Case "EASYS"
            SaveTransData = SaveTransData_EASYS(argSpcRow)
            
        Case "GINUS"
            SaveTransData = SaveTransData_GINUS(argSpcRow)
        
        Case "GSEN"
            SaveTransData = SaveTransData_MSINFOTEC(argSpcRow)
        
        Case "HWASAN"
            SaveTransData = SaveTransData_HWASAN(argSpcRow)
        
        Case "ILSIN"
            SaveTransData = SaveTransData_ILSIN(argSpcRow)
        
        Case "JAINCOM"
            SaveTransData = SaveTransData_JAINCOM(argSpcRow)
        
        Case "JWINFO"
            SaveTransData = SaveTransData_JWINFO(argSpcRow)
        
        Case "KCHART"
            SaveTransData = SaveTransData_KCHART(argSpcRow)
        
        Case "KOMAIN"
            SaveTransData = SaveTransData_KOMAIN(argSpcRow)
        
        Case "KYU"
            SaveTransData = SaveTransData_KYU(argSpcRow)
        
        Case "MEDICHART"
            SaveTransData = SaveTransData_MEDICHART(argSpcRow)
        
        Case "MEDIIT"
            SaveTransData = SaveTransData_MEDIIT(argSpcRow)
        
        Case "MEDITOLISS"
            SaveTransData = SaveTransData_MEDITOLISS(argSpcRow)
        
        Case "MCC"
            SaveTransData = SaveTransData_MCC(argSpcRow)
        
        Case "MOD"
            SaveTransData = SaveTransData_MOD(argSpcRow)
        
        Case "MSINFOTEC"
            SaveTransData = SaveTransData_MSINFOTEC(argSpcRow)

        Case "NEOSOFT"
            SaveTransData = SaveTransData_NEOSOFT(argSpcRow)

        Case "ONITGUM"
            SaveTransData = SaveTransData_ONITGUM(argSpcRow)

        Case "ONITEMR"
            SaveTransData = SaveTransData_ONITEMR(argSpcRow)

        Case "PLIS"
            SaveTransData = SaveTransData_PLIS(argSpcRow)

        Case "SY"
            SaveTransData = SaveTransData_SY(argSpcRow)
        
        Case "TWIN"
            SaveTransData = SaveTransData_TWIN(argSpcRow)

        Case "UBCARE"
            SaveTransData = SaveTransData_UBCARE(argSpcRow)

        
        Case Else
            SaveTransData = -1
    End Select


End Function
                    


Function SaveTransData_EONM(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim blnSave         As Boolean
    Dim intRow          As Integer
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    
    Dim strTestCd       As String
    Dim strSubCode      As String
    Dim strEqpcd        As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
On Error GoTo ErrHandle
    
    blnSave = False
    intRow = 0
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmMain
        SaveTransData_EONM = -1
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT                                                   " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'                            " & vbCr                      '장비코드
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'        " & vbCr  '검사일
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        
        Set RS_L = New ADODB.Recordset
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
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- 서버저장
                SQL = ""
                SQL = SQL & "Update TB_H141_LISTAKEBODY                     " & vbCr
                SQL = SQL & "   SET H141_RSLTYN    ='Y'                     " & vbCr
                SQL = SQL & " WHERE H141_TSAMPLENO = '" & strBarcode & "'   " & vbCr
                SQL = SQL & "   AND H141_SUGACD    = '" & strTestCd & "'    " & vbCr
                
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                
                SQL = ""
                SQL = SQL & "UPDATE TB_H131_SPPRESULT                       " & vbCr
                SQL = SQL & "   SET H131_RESULT  = '" & sResult & "'        " & vbCr
                SQL = SQL & " WHERE H131_SPPTYPE = '" & gHOSP.PARTCD & "'   " & vbCr    'L010
                SQL = SQL & "   AND H131_SEQNO   = '" & strSubCode & "'     " & vbCr
                    
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
            
                SQL = ""
                SQL = SQL & "UPDATE TB_H130_SPPRECEIVE                              " & vbCr
                SQL = SQL & "   SET H130_RSLTDAT = TO_CHAR(SYSDATE, 'YYYYMMDD')     " & vbCr
                SQL = SQL & "      ,H130_RSLTTM  = TO_CHAR(SYSDATE, 'HH24:MI:SS')   " & vbCr
                SQL = SQL & " WHERE H130_SPPTYPE = '" & gHOSP.PARTCD & "'           " & vbCr    'L010
                SQL = SQL & "   AND H130_SEQNO   = '" & strSubCode & "'             " & vbCr
                    
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
            
                SQL = ""
                SQL = SQL & "UPDATE TB_H140_LISTAKEHEAD                     " & vbCr
                SQL = SQL & "   SET H140_RSLTYN    = 'Y'                    " & vbCr
                SQL = SQL & " WHERE H140_TSAMPLENO = '" & strBarcode & "'   " & vbCr
                                    
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                        
            End If
        Next intRow
        
        SaveTransData_EONM = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_EONM = -1
    
End Function

Function SaveTransData_EASYS(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim blnSave         As Boolean
    Dim intRow          As Integer
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    
    Dim strTestCd       As String
    Dim strSubCode      As String
    Dim strEqpcd        As String
    Dim sResult         As String
    Dim sResult1        As String
    Dim sResult2        As String
    Dim strJudge        As String
    
On Error GoTo ErrHandle
    
    blnSave = False
    intRow = 0
    strJudge = ""
    sResult = ""
    sResult1 = ""
    sResult2 = ""

    With frmMain
        SaveTransData_EASYS = -1
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT                                                   " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'                            " & vbCr                      '장비코드
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'        " & vbCr  '검사일
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        
        Set RS_L = New ADODB.Recordset
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
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            strSubCode = Trim(GetText(.vasTemp, intRow, 4))     '검사코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            'strJudge = getEASYSJudge(strTestCd, sResult)
            'strJudge = SetJudge_EASYS(strTestCd, sResult)
            strJudge = SetJudge(strTestCd, sResult)
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- 서버저장
                SQL = ""
                SQL = SQL & "UPDATE H3LAB_RESULT    " & vbCr
                SQL = SQL & "   SET STS_CD     = 'R'" & vbCr
                SQL = SQL & "      ,RESULT_VAL = '" & sResult & "'      " & vbCr '수치형결과
                SQL = SQL & "      ,RESULT_NM  = '" & sResult & "'      " & vbCr '(수치형 + 서술형 등 모든결과)
                SQL = SQL & "      ,HL_GB      = '" & strJudge & "'     " & vbCr
                SQL = SQL & " WHERE RECEPT_NO  = '" & strBarcode & "'   " & vbCr
                SQL = SQL & "   And ORD_CD     = '" & strTestCd & "'    " & vbCr
                SQL = SQL & "   And STS_CD     = 'A'                    " & vbCr

                                    
                Call SetSQLData("결과저장", SQL, "A")
                AdoCn.Execute SQL
                        
            End If
        Next intRow
        
        SaveTransData_EASYS = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_EASYS = -1
    
End Function



Public Function getEASYSJudge(ByVal pOrdCD As String, ByVal pResult As String) As String
    Dim RSJ         As ADODB.Recordset
    Dim strLow      As String
    Dim strHigh     As String
    
    getEASYSJudge = ""
    
          SQL = "Select REFLOW, REFHIGH  "
    SQL = SQL & "  From EQPMASTER"
    SQL = SQL & " Where EQUIPCD = '" & gHOSP.MACHCD & "' "
    SQL = SQL & "   And TESTCODE =  '" & pOrdCD & "'"
    
    Set RSJ = New ADODB.Recordset
    Set RSJ = AdoCn_Local.Execute(SQL, , 1)
    If Not RSJ.EOF = True And Not RSJ.BOF = True Then
        strLow = Trim(RSJ.Fields("REFLOW") & "")
        strHigh = Trim(RSJ.Fields("REFHIGH") & "")
        
        If strLow <> "" And strHigh <> "" And pResult <> "" And IsNumeric(strLow) And IsNumeric(strHigh) And IsNumeric(pResult) Then
            If Val(pResult) > Val(strHigh) Then
                getEASYSJudge = "H"
            ElseIf Val(pResult) < Val(strLow) Then
                getEASYSJudge = "L"
            Else
                getEASYSJudge = " "
            End If
        Else
            getEASYSJudge = " "
        End If
    Else
        getEASYSJudge = ""
    End If
        
    RSJ.Close
    
End Function



Function GetOrderSeqCode(argExamDt As String, argPID As String, argPCD As String) As String
    Dim RS As ADODB.Recordset
    
    '-- SEQ 가져오기
    
          SQL = "SELECT /*+ INDEX(rslt scrrslth_ux1) INDEX (coif scccoifm_ix1) */" & vbCr
    SQL = SQL & "       rslt.smp_no, rslt.prcp_seq, rslt.exam_seq, rslt.rept_seq, rslt.cd, rslt.pt_no, rslt.exam_stus, rslt.mach_rslt, rslt.exam_rslt ," & vbCr
    SQL = SQL & "       coif.exam_nm, prex.acp_dt, ptbs.pt_nm, ptbs.ssn_1, ptbs.ssn_2, xpsl.pt_no, " & vbCr
    SQL = SQL & "       DECODE(xpsl.gnl_add_typ_cd,'3','I',xpsl.prcp_knd_cd), xpsl.adms_ymd, xpsl.mn_sub_typ_cd, xpsl.med_dpt_cd, xpsl.med_ymd, coif.spc_cd, codm.cd_desc" & vbCr
    SQL = SQL & "  FROM scrrslth rslt, scccoifm coif, scccodem codm, scrprexh prex, mosxpslh xpsl, pmcptbsm ptbs" & vbCr
    SQL = SQL & " WHERE rslt.hos_org_no   = '" & gHOSP.HOSPCD & "'" & vbCr & vbCr
    SQL = SQL & "  AND SUBSTR(prex.acp_dt,1,8) BETWEEN '" & argExamDt & "' AND '" & argExamDt & "'" & vbCr
    SQL = SQL & "  AND rslt.smp_no       = '" & argPID & "'" & vbCr
    SQL = SQL & "  AND rslt.cd           = '" & argPCD & "'" & vbCr
    SQL = SQL & "  AND rslt.exam_stus  IN ('0','1','2')" & vbCr
    SQL = SQL & "  AND coif.hos_org_no   = rslt.hos_org_no" & vbCr
    SQL = SQL & "  AND coif.exam_cd      = rslt.cd" & vbCr
    SQL = SQL & "  AND SUBSTR(prex.acp_dt,1,8) BETWEEN coif.fr_dt AND coif.to_dt" & vbCr
    SQL = SQL & "  AND coif.exam_mach_cd = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & "  AND codm.hos_org_no   = coif.hos_org_no" & vbCr
    SQL = SQL & "  AND codm.typ_cd       = '02'" & vbCr
    SQL = SQL & "  AND codm.cd           = coif.spc_cd" & vbCr
    SQL = SQL & "  AND SUBSTR(prex.acp_dt,1,8) BETWEEN codm.fr_dt AND codm.to_dt" & vbCr
    SQL = SQL & "  AND prex.hos_org_no   = rslt.hos_org_no" & vbCr
    SQL = SQL & "  AND prex.smp_no       = rslt.smp_no" & vbCr
    SQL = SQL & "  AND prex.prcp_seq     = rslt.prcp_seq" & vbCr
    SQL = SQL & "  AND prex.exam_seq     = rslt.exam_seq" & vbCr
    SQL = SQL & "  AND xpsl.hos_org_no   = prex.hos_org_no" & vbCr
    SQL = SQL & "  AND xpsl.smp_no       = prex.smp_no" & vbCr
    SQL = SQL & "  AND xpsl.acp_no       = prex.prcp_seq" & vbCr
    SQL = SQL & "  AND xpsl.prcp_typ_cd IN ('O','C')" & vbCr
    SQL = SQL & "  AND ptbs.hos_org_no   = prex.hos_org_no" & vbCr
    SQL = SQL & "  AND ptbs.pt_no        = prex.pt_no" & vbCr

    Call SetSQLData("SEQ찾기", SQL)

    '-- Record Count 가져옴
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            GetOrderSeqCode = GetOrderSeqCode & Trim(RS.Fields("prcp_seq")) & "|" & Trim(RS.Fields("exam_seq")) & "|" & Trim(RS.Fields("rept_seq")) & "|"
            RS.MoveNext
        Loop
    End If
    
    If GetOrderSeqCode <> "" Then
        GetOrderSeqCode = Mid(GetOrderSeqCode, 1, Len(GetOrderSeqCode) - 1)
    End If
    
    Set RS = Nothing
    
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
    Dim prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim prm3 As New ADODB.Parameter
    Dim prm4 As New ADODB.Parameter
    Dim prm5 As New ADODB.Parameter
    
    
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
                    
                    
                    Set prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, intBarno)      '바코드번호
                    .Parameters.Append prm1

                    Set prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       '처방코드
                    .Parameters.Append prm2

                    Set prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '결과값
                    .Parameters.Append prm3

                    Set prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, gHOSP.MACHCD)    '장비코드
                    .Parameters.Append prm4

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
    Dim prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim prm3 As New ADODB.Parameter
    Dim prm4 As New ADODB.Parameter
    Dim prm5 As New ADODB.Parameter
    
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
                                        
                    Set prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, intBarno)      '바코드번호
                    .Parameters.Append prm1

                    Set prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       '처방코드
                    .Parameters.Append prm2

                    Set prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '결과값
                    .Parameters.Append prm3

                    Set prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, strMachCD)    '장비코드
                    .Parameters.Append prm4

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
    Dim prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim prm3 As New ADODB.Parameter
    Dim prm4 As New ADODB.Parameter
    Dim prm5 As New ADODB.Parameter
    
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
                    
                    
                    Set prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, intBarno)      '바코드번호
                    .Parameters.Append prm1

                    Set prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       '처방코드
                    .Parameters.Append prm2

                    Set prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '결과값
                    .Parameters.Append prm3

                    Set prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, strMachCD)    '장비코드
                    .Parameters.Append prm4

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

    Select Case gEMR
        Case "AMIS"                         '아미스
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EMEDI"                        '이메디
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "BIT"                          '비트
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)

        Case "BIT70"                        '비트 HIB70
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EASYS"                        '이지스
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "EONM"                         '이온엠
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "GSEN"                         '지센커뮤니케이션즈(이챠트)
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "HWASAN"                       '화산
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "JAINCOM"                       '자인컴
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
        
        Case "JWINFO"                       '중외정보
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "KCHART"                       '다대소프트
                SetJudge = SetJudge_KCHART(asResult, asEquipCode)
        
        Case "KOMAIN"                       '중외정보
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "KYU"                          '건양대학교병원
                '워크리스트 기능없음
                'SetJudge =  SetJudge_KYU(asResult,asEquipCode)
        Case "MEDICHART"                    '메디챠트
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "MEDIIT"
                SetJudge = SetJudge_LOCAL(asResult, asEquipCode)
            
        Case "MEDITOLISS"                    '
                SetJudge = SetJudge_MEDITOLISS(asResult, asEquipCode)
            
        Case "MSINFOTEC"                    'MS인포텍
                SetJudge = SetJudge_MSINFOTEC(asResult, asEquipCode)
                
    End Select
    
End Function

Function SetJudge_LOCAL(asResult As String, asEquipCode As String)
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
    
    If Not IsNumeric(sEquipRes) Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT REFLOW, REFHIGH                     " & vbCr
    SQL = SQL & "  FROM EQPMASTER                           " & vbCr
    SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & "   AND RSLTCHANNEL = '" & sEquipCode & "'  " & vbCr

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
 
    SetJudge_LOCAL = sResFlag
    
End Function

Function SetJudge_EASYS(asResult As String, asEquipCode As String) As String
    Dim RSJ         As ADODB.Recordset
    Dim strLow      As String
    Dim strHigh     As String
    
    SetJudge_EASYS = ""
    
          SQL = "Select REFLOW, REFHIGH                     " & vbCr
    SQL = SQL & "  From EQPMASTER                           " & vbCr
    SQL = SQL & " Where EQUIPCD  = '" & gHOSP.MACHCD & "'   " & vbCr
    SQL = SQL & "   And TESTCODE = '" & asEquipCode & "'    " & vbCr
    
    Set RSJ = New ADODB.Recordset
    Set RSJ = AdoCn_Local.Execute(SQL, , 1)
    If Not RSJ.EOF = True And Not RSJ.BOF = True Then
        strLow = Trim(RSJ.Fields("REFLOW") & "")
        strHigh = Trim(RSJ.Fields("REFHIGH") & "")
        
        If strLow <> "" And strHigh <> "" And asResult <> "" And IsNumeric(strLow) And IsNumeric(strHigh) And IsNumeric(asResult) Then
            If Val(asResult) > Val(strHigh) Then
                SetJudge_EASYS = "H"
            ElseIf Val(asResult) < Val(strLow) Then
                SetJudge_EASYS = "L"
            Else
                SetJudge_EASYS = " "
            End If
        Else
            SetJudge_EASYS = " "
        End If
    Else
        SetJudge_EASYS = ""
    End If
        
    RSJ.Close

End Function

Function SetJudge_MSINFOTEC(asResult As String, asEquipCode As String) As String
    Dim RSJ         As ADODB.Recordset
    Dim sqlRet      As Integer
    Dim sqlDoc      As String
    Dim strAge      As String
    Dim strSex      As String
    Dim stryy, strmm, strdd, strDate  As String
    
On Error GoTo ErrorTrap
    
    SetJudge_MSINFOTEC = ""
    
    asResult = Replace(asResult, "<", "")
    asResult = Replace(asResult, ">", "")
    
    strAge = mPatient.AGE
    strSex = mPatient.SEX
    
    If strAge <> "" Then
        If strAge <= 7 Then
            SQL = "Select YMAX as MAX, YMIN as MIN "
        Else
            If strSex = "M" Then
                     SQL = "Select MMAX as MAX, MMIN as MIN "
            Else
                     SQL = "Select WMAX as MAX, WMIN as MIN "
            End If
        End If
    Else
        SQL = "Select MMAX as MAX, MMIN as MIN "
    End If
    
    SQL = SQL & "  From emr.LABMAST"
    SQL = SQL & " Where ORCD =  '" & asEquipCode & "'"
    
    Set RSJ = AdoCn.Execute(SQL)
    Do Until RSJ.EOF
        If IsNumeric(asResult) And IsNumeric(RSJ.Fields("MAX")) And IsNumeric(RSJ.Fields("MIN")) Then
            If Val(asResult) > Val(RSJ.Fields("MAX")) Then
                SetJudge_MSINFOTEC = "H"
            ElseIf Val(asResult) < Val(RSJ.Fields("MIN")) Then
                SetJudge_MSINFOTEC = "L"
            Else
                SetJudge_MSINFOTEC = " "
            End If
        Else
            SetJudge_MSINFOTEC = " "
        End If
        RSJ.MoveNext
    
    Loop
    
    RSJ.Close

Exit Function

ErrorTrap:
    SetJudge_MSINFOTEC = ""
    
End Function

Function SetJudge_MEDITOLISS(asResult As String, asEquipCode As String) As String
    Dim RSJ         As ADODB.Recordset
    Dim strRefVal   As String
    
On Error GoTo ErrorTrap
    
    SetJudge_MEDITOLISS = ""
    
    SQL = ""
    SQL = SQL & "SELECT REFER_VALUE                                 " & vbCr
    SQL = SQL & "  FROM MEDITOLISS..TOTRES                          " & vbCr
    SQL = SQL & " WHERE REQUEST_DATE    = '" & mResult.RsltDate & "'" & vbCr
    SQL = SQL & "   AND EXAM_NO         = '" & mResult.BarNo & "'   " & vbCr
    SQL = SQL & "   AND EXAM_CODE       = '" & asEquipCode & "'     " & vbCr
    
    Set RSJ = AdoCn.Execute(SQL)
    Do Until RSJ.EOF
        strRefVal = RSJ.Fields("REFER_VALUE").Value & ""
        If IsNumeric(asResult) And Len(strRefVal) > 0 Then
            If Val(Trim$(asResult)) < Val(Mid(strRefVal, 1, InStr(strRefVal, "~") - 1)) Then
                SetJudge_MEDITOLISS = "L"
            ElseIf Val(Trim$(asResult)) > Val(Mid(strRefVal, InStr(strRefVal, "~") + 1)) Then
                SetJudge_MEDITOLISS = "H"
            Else
                SetJudge_MEDITOLISS = ""
            End If
        End If
    Loop
                
    RSJ.Close
    
Exit Function

ErrorTrap:
    SetJudge_MEDITOLISS = ""
    
End Function

Function SetJudge_KCHART(asResult As String, asEquipCode As String) As String
    Dim RS1         As ADODB.Recordset
    Dim sEquipCode  As String
    Dim sEquipRes   As String
    Dim sResFlag    As String
    Dim strRefL     As String
    Dim strRefH     As String
    
    
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    sResFlag = ""
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    strRefL = ""
    strRefH = ""
    
'    SQL = SQL & "  L.진료검사ID AS R, " & vbCrLf
'    SQL = SQL & "  L.진료지원ID AS P, " & vbCrLf

    '성인남 참고치0~참고치1,
    '성인여 참고치2~참고치3,
    '소아남 참고치4~참고치5,
    '소아여 참고치6~참고치7
    
    SQL = ""
    SQL = SQL & "SELECT DISTINCT "
    SQL = SQL & "       A.환자성별 AS 성별                                          " & vbCr
    SQL = SQL & "     , L.참고치0, L.참고치1, L.참고치2, L.참고치3                  " & vbCr
    SQL = SQL & "     , L.참고치4, L.참고치5, L.참고치6, L.참고치7                  " & vbCr
    SQL = SQL & "     , (L.처방코드 + L.서브코드) AS ITEM                           " & vbCr
    SQL = SQL & "  FROM             TB_진료검사 L                                   " & vbCr
    SQL = SQL & "       INNER JOIN  TB_진료지원 J ON (L.진료지원ID = J.진료지원ID)  " & vbCr
    SQL = SQL & "       INNER JOIN  TB_진료일반 A ON (J.진료일자   = A.진료일자     " & vbCr
    SQL = SQL & "                                AND  J.챠트번호   = A.챠트번호     " & vbCr
    SQL = SQL & "                                AND  J.진료번호   = A.진료번호)    " & vbCr
    SQL = SQL & "  Where L.검체번호 = '" & mResult.BarNo & "'                       " & vbCr
    SQL = SQL & "    AND L.검사상태 < 5                                             " & vbCr
    SQL = SQL & "    AND (L.처방코드 + L.서브코드) = '" & sEquipCode & "'           " & vbCr
                                                                 

     Call SetSQLData("참고치조회", SQL)
     
     '-- Record Count 가져옴
     AdoCn.CursorLocation = adUseClient
     Set RS1 = AdoCn.Execute(SQL, , 1)
     If Not RS1.EOF = True And Not RS1.BOF = True Then
         Do Until RS1.EOF
            strRefL = ""
            strRefH = ""
            If Trim(RS1.Fields("성별")) & "" = "M" Then
                If Trim(RS1.Fields("참고치0")) & "" <> "" Then
                    strRefL = Trim(RS1.Fields("참고치0")) & ""
                    strRefH = Trim(RS1.Fields("참고치1")) & ""
                End If
            Else
                If Trim(RS1.Fields("성별")) & "" = "F" Then
                    If Trim(RS1.Fields("참고치2")) & "" <> "" Then
                        strRefL = Trim(RS1.Fields("참고치2")) & ""
                        strRefH = Trim(RS1.Fields("참고치3")) & ""
                    Else
                        strRefL = Trim(RS1.Fields("참고치0")) & ""
                        strRefH = Trim(RS1.Fields("참고치1")) & ""
                    End If
                End If
            End If
            RS1.MoveNext
        Loop
    
        If IsNumeric(sEquipRes) And IsNumeric(strRefL) = True And IsNumeric(strRefH) = True Then
            If CCur(sEquipRes) > CCur(strRefL) And CCur(sEquipRes) < CCur(strRefH) Then
                sResFlag = ""
            ElseIf CCur(strRefH) <= CCur(sEquipRes) Then
                sResFlag = "H"
            ElseIf CCur(strRefL) >= CCur(sEquipRes) Then
                sResFlag = "L"
            End If
        End If
    End If
    
    RS1.Clone
    
    SetJudge_KCHART = sResFlag
    
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

Function SetCutOffResult(asResult As String, asEquipCode As String, asTestCode As String)
    Dim RS_L        As ADODB.Recordset
    Dim i As Integer
    Dim sEquipCode As String
    Dim sEquipRes As String
    Dim sResult As String
'    Dim sPoint As Integer
'    Dim sResType As String
    
    Dim dblLow      As Double
    Dim dblHigh     As Double
    Dim strLComp    As String
    Dim strHComp    As String
    
    sResult = ""
    sEquipRes = Trim(asResult)
    sEquipCode = Trim(asEquipCode)
    
    If sEquipCode = "" Then
        Exit Function
    End If
    
    SQL = ""
    SQL = SQL & "SELECT RESULTTYPE, COLIN, COLCOMP, COLOUT, COHIN, COHCOMP, COHOUT, COMOUT   " & vbCr
    SQL = SQL & "  FROM EQPMASTER                                                " & vbCr
    SQL = SQL & " WHERE EQUIPCD     = '" & gHOSP.MACHCD & "'                     " & vbCr
    SQL = SQL & "   AND RSLTCHANNEL = '" & sEquipCode & "'                       " & vbCr
    SQL = SQL & "   AND TESTCODE    = '" & asTestCode & "'                       " & vbCr

    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
        If Trim(RS_L.Fields("COLCOMP") & "") <> "" And Trim(RS_L.Fields("COLIN") & "") <> "" Then
            If IsNumeric(Trim(RS_L.Fields("COLIN") & "")) Then
                dblLow = CCur(RS_L.Fields("COLIN"))
                strLComp = Trim(RS_L.Fields("COLCOMP"))
                If strLComp = "<" Then
                    If CCur(asResult) < dblLow Then
                        sResult = Trim(RS_L.Fields("COLOUT") & "")
                    Else
                        sResult = Trim(RS_L.Fields("COMOUT") & "")
                    End If
                ElseIf strLComp = "<=" Then
                    If CCur(asResult) <= dblLow Then
                        sResult = Trim(RS_L.Fields("COLOUT") & "")
                    Else
                        sResult = Trim(RS_L.Fields("COMOUT") & "")
                    End If
                End If
            End If
        ElseIf Trim(RS_L.Fields("COHCOMP") & "") <> "" And Trim(RS_L.Fields("COHIN") & "") <> "" Then
            If IsNumeric(Trim(RS_L.Fields("COHIN") & "")) Then
                dblHigh = CCur(RS_L.Fields("COHIN"))
                strHComp = Trim(RS_L.Fields("COHCOMP"))
                If strHComp = ">" Then
                    If CCur(asResult) < dblLow Then
                        sResult = Trim(RS_L.Fields("COHOUT") & "")
                    Else
                        sResult = Trim(RS_L.Fields("COMOUT") & "")
                    End If
                ElseIf strHComp = ">=" Then
                    If CCur(asResult) >= dblHigh Then
                        sResult = Trim(RS_L.Fields("COHOUT") & "")
                    Else
                        sResult = Trim(RS_L.Fields("COMOUT") & "")
                    End If
                End If
            End If
        End If
    End If
    
    If sResult <> "" Then
        Select Case Trim(RS_L.Fields("RESULTTYPE") & "")
            Case "변함없음"
                    sResult = Trim(asResult)
            Case "정량"
                    sResult = Trim(asResult)
            Case "정성"
                    sResult = Trim(sResult)
            Case "정량(정성)"
                    sResult = asResult & "(" & Trim(sResult) & ")"
            Case "정성(정량)"
                    sResult = sResult & "(" & Trim(asResult) & ")"
        End Select
    End If
    
    RS_L.Close
    
    SetCutOffResult = sResult
    
End Function
