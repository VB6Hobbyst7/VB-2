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
    'for LABOSPECT
    SmplID      As String
    SmplNo      As String
    RackType    As String
    ContType    As String
    Kind        As String
    OtherOrder  As String
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
    'for LABOSPECT
    SmplNo      As String
    RackType    As String
    ContType    As String
    PARTGBN     As String
    'LABSEQ      As String
    SPECIMENCD  As String
    'for ACK
    PSEX        As String
    PAGE        As String
    Delta       As String
    Panic       As String
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

Public gErr            As String

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
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = Trim(pBarno) Then
                    If Trim(GetText(frmMain.spdOrder, i, colSAVESEQ)) = "" Then
                        mResult.BarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                        mResult.PatNo = Trim(GetText(frmMain.spdOrder, i, colPID))
                        intRow = i
                        Exit For
                    End If
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
        Call SetText(.spdOrder, mResult.SmplNo, intRow, colSEQNO)
        Call SetText(.spdOrder, mResult.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mResult.TubePos, intRow, colPOSNO)
    
        '-- 환자정보 표시
        Call spdActiveCell(.spdOrder, intRow, colBARCODE)
        .spdOrder.Action = ActionActiveCell
        
        
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
        
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
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
                SQL = SQL & " WHERE WAITSEQNO   = '" & strPatID & "'" & vbCr
                SQL = SQL & "   AND MAP2SEQNO   = '" & strTestCd & "' " & vbCr
                SQL = SQL & "   AND ENTERDATE   = '" & strHospDate & "'"
                
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

Function SaveTransData_CHOSUN(ByVal argSpcRow As Integer) As Integer
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
        SaveTransData_CHOSUN = -1
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
                
        If Trim(strHospDate) = "" Then
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
            strOrdCd = Trim(GetText(.vasTemp, intRow, 2))       '항목코드
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
            
            If Mid(strBarcode, 5, 2) = "99" Then
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- 서버저장
                    If SP_UPCPL0820(sResult, Format(Now, "yyyymmdd"), Format(Now, "hhmm"), gHOSP.MACHCD, strTestCd, strBarcode) = True Then
                        SaveTransData_CHOSUN = 1
                    End If
                End If
            Else
                If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                    '-- 서버저장
                    SQL = ""
                    SQL = SQL & "INSERT INTO medi.CPL0891 " & vbCr
                    SQL = SQL & " ( SYS_DATE,USER_ID,UPD_DATE,JANGBI_NAME,  SAMPLE_ID, HANGMOG_CODE, CPL_RESULT,CHK_FLAG,RESULT_DATE,RESULT_SEQ ) " & vbCr
                    SQL = SQL & " VALUES " & vbCr
                    SQL = SQL & " ( SYSDATE,'" & gHOSP.USERID & "',SYSDATE,'" & gHOSP.MACHCD & "','" & strBarcode & "','" & strTestCd & "','" & sResult & "','N',TO_DATE('" & strDate & "','YYMMDD'), '' )"
    
    'INSERT INTO medi.CPL0891
    '( SYS_DATE,USER_ID,UPD_DATE,JANGBI_NAME,    SAMPLE_ID, HANGMOG_CODE, CPL_RESULT,CHK_FLAG,RESULT_DATE,RESULT_SEQ )
    'Values
    '( SYSDATE,'CHORUS',SYSDATE,'II','1304162030','HSV-G','12.4','N',TO_DATE('20130417','YYMMDD'),   '' )
                    Call SetSQLData("결과저장", SQL, "A")
                    
                    AdoCn.Execute SQL
                    
                    SaveTransData_CHOSUN = 1
                    
                End If
            End If
        Next intRow
        
        
    End With

Exit Function

ErrHandle:
    SaveTransData_CHOSUN = -1
    
End Function

'QC결과 Update
Public Function SP_UPCPL0820(ByVal stI_CPL_RESULT As String, ByVal stI_RESULT_DATE As String, _
                             ByVal stI_RESULT_TIME As String, ByVal stI_JANGBI_CODE As String, _
                             ByVal stI_JANGBI_OUT As String, ByVal stI_SPECIMEN_SER As String) As Boolean
    
    Dim adoCommand As ADODB.Command
    Dim stError As String
    
    On Error GoTo ErrHandler
'    If ConnectionCheck(CNNORA) = False Then
'       SP_UPCPL0820 = False
'       Exit Function
'    End If
   
    Set adoCommand = New ADODB.Command
    Set adoCommand.ActiveConnection = AdoCn

    adoCommand.CommandText = "MEDI.PR_CPL_UPDATE_CPL0820"
    adoCommand.CommandType = adCmdStoredProc
    
    With adoCommand
        .Parameters.Append .CreateParameter("I_CPL_RESULT", adVarChar, adParamInput, 30, stI_CPL_RESULT)
        .Parameters.Append .CreateParameter("I_RESULT_DATE", adVarChar, adParamInput, 12, stI_RESULT_DATE) 'Format(Date, "yyyymmdd")
        .Parameters.Append .CreateParameter("I_RESULT_TIME", adVarChar, adParamInput, 4, stI_RESULT_TIME)  'Format(Time, "hhmm")
        .Parameters.Append .CreateParameter("I_JANGBI_CODE", adVarChar, adParamInput, 3, stI_JANGBI_CODE)
        .Parameters.Append .CreateParameter("I_JANGBI_OUT", adVarChar, adParamInput, 20, stI_JANGBI_OUT)
        .Parameters.Append .CreateParameter("I_SPECIMEN_SER", adVarChar, adParamInput, 20, stI_SPECIMEN_SER)
    End With
    AdoCn.BeginTrans
    adoCommand.Execute
    AdoCn.CommitTrans
    Set adoCommand = Nothing

    SP_UPCPL0820 = True
    Exit Function
    
ErrHandler:
    AdoCn.RollbackTrans
    SP_UPCPL0820 = False
    Set adoCommand = Nothing
    If Err.Number <> 0 Then
        'If CheckOraConnection(Err.Description) = False Then fNetConnection = False
        stError = "SP_UPCPL0820(QC결과 Update) 실행 오류" & vbCrLf & "Error NO: " & Err.Number & vbCrLf & "Error  객체: " & Err.Source & vbCrLf _
                    & "Error 설명: " & Err.Description & vbCrLf
        'Call Errlog("DBErrlog", stError)
        'Err.Clear
    End If
End Function

'-- 검사결과 서버저장
Function SaveTransData(ByVal argSpcRow As Integer) As Integer
    
    SaveTransData = -1
    
    Select Case gEMR

        Case "ONITGUM"
            SaveTransData = SaveTransData_ONITGUM(argSpcRow)

        Case "ONITEMR"
            SaveTransData = SaveTransData_ONITEMR(argSpcRow)

        Case "CHOSUN"
            SaveTransData = SaveTransData_CHOSUN(argSpcRow)

        Case Else
            SaveTransData = -1
    End Select


End Function
                    
'-- 검사결과 서버저장
Function SaveTransDataR(ByVal argSpcRow As Integer) As Integer
    
    SaveTransDataR = -1
    
    Select Case gEMR
'        Case "ONITEMR"
'            SaveTransDataR = SaveTransDataR_ONITEMR(argSpcRow)

        Case Else
            SaveTransDataR = -1
    
    End Select

End Function



'Function SaveTransData_MCC_R(ByVal argSpcRow As Integer) As Integer
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
'    'Dim strReturn       As String
'    Dim intReturn       As Long
'    Dim strMSG          As String
'
'    Dim prm0 As New ADODB.Parameter
'    Dim prm1 As New ADODB.Parameter
'    Dim prm2 As New ADODB.Parameter
'    Dim prm3 As New ADODB.Parameter
'    Dim prm4 As New ADODB.Parameter
'    Dim prm5 As New ADODB.Parameter
'
'
'    Dim intBarno  As Double
'
'On Error GoTo ErrHandle
'
'    With frmMain
'        SaveTransData_MCC_R = -1
'        intRow = 0
'
'        strSaveSeq = Trim(GetText(.spdROrder, argSpcRow, colSAVESEQ))
'        strExamDate = Trim(GetText(.spdROrder, argSpcRow, colEXAMDATE))
'        strHospDate = Trim(GetText(.spdROrder, argSpcRow, colHOSPDATE))
'        strBarcode = Trim(GetText(.spdROrder, argSpcRow, colBARCODE))
'        strChartNo = Trim(GetText(.spdROrder, argSpcRow, colCHARTNO))
'        strPatID = Trim(GetText(.spdROrder, argSpcRow, colPID))
'
'        If IsNumeric(strBarcode) Then
'            intBarno = CDbl(strBarcode)
'        Else
'            Exit Function
'        End If
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
'                      SQL = "Exec UP_LIS_INTERFACE_U$001 " & intBarno
'                SQL = SQL & "," & strTestCd
'                SQL = SQL & "," & sResult
'                SQL = SQL & "," & gHOSP.MACHCD
'
'                'AdoCn.Execute SQL
'
'                Set AdoCmd = New ADODB.Command
'                Set AdoCmd.ActiveConnection = AdoCn
'                With AdoCmd
'                    .CommandTimeout = 15
'                    .CommandText = "UP_LIS_INTERFACE_U$001"
'                    .CommandType = adCmdStoredProc
'
'
'                    Set prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, intBarno)      '바코드번호
'                    .Parameters.Append prm1
'
'                    Set prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       '처방코드
'                    .Parameters.Append prm2
'
'                    Set prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '결과값
'                    .Parameters.Append prm3
'
'                    Set prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, gHOSP.MACHCD)    '장비코드
'                    .Parameters.Append prm4
'
'                    .Execute
'
'                End With
'
'                'Call SetSQLData("결과저장", SQL)
'
'                SaveTransData_MCC_R = 1
'
'            End If
'        Next intRow
'
'    End With
'
'Exit Function
'
'ErrHandle:
'    SaveTransData_MCC_R = -1
'
'End Function

Function SetJudge(asResult As String, asEquipCode As String)

    SetJudge = ""
    
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

Function SetCutOffResult(asResult As String, asEquipCode As String, asTestCode As String) As String
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
