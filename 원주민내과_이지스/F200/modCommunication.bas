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
Public Sub SetPatInfo(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    Call SetCommStatus("R", pBarNo, frmMain.spdComStatus)
    
    intRow = -1
    With frmMain
        Select Case pType
            Case "0"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colBARCODE) = pBarNo Then
                        If GetText(frmMain.spdOrder, i, colSAVESEQ) = mResult.RsltSeq Then
                            intRow = i
                            Exit For
                        End If
                    End If
                Next i
            '-- Seq
            Case "1"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colSEQNO) = mOrder.Seq Then
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Rack/Pos
            Case "2"
                For i = 1 To .spdOrder.DataRowCnt
                    If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                        intRow = i
                        Exit For
                    End If
                Next i
            '-- Check Top
            Case "3"
                For i = 1 To .spdOrder.DataRowCnt
                    If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" And GetText(frmMain.spdOrder, i, colSTATE) = "" Then
                        intRow = i
                        mOrder.BarNo = GetText(frmMain.spdOrder, i, colBARCODE)
                        mResult.BarNo = GetText(frmMain.spdOrder, i, colBARCODE)
                        Exit For
                    End If
                Next i
        End Select
        
        '-- 스프레드에서 못찾았음..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        
        '-- 장비결과인덱스 화면표시
        Call SetText(.spdOrder, "1", intRow, colCHECKBOX)
        Call SetText(.spdOrder, mResult.RsltDate, intRow, colEXAMDATE)
        Call SetText(.spdOrder, mResult.RsltTime, intRow, colEXAMTIME)
        Call SetText(.spdOrder, mResult.RsltSeq, intRow, colSAVESEQ)
        Call SetText(.spdOrder, mResult.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mResult.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mResult.TubePos, intRow, colPOSNO)
        Call SetText(.spdOrder, mResult.Seq, intRow, colSEQNO)
    
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        '.spdOrder.Col = 1
        '.spdOrder.Action = ActionActiveCell
        
        .spdOrder.RowHeight(-1) = 15
        
'        Call frmMain.spdOrder_Click(colBARCODE, intRow)
        
    End With
    
    '-- 현재 Row
    gRow = intRow
    
End Sub

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
    strSndMsg = STX & strSndMsg & ETX ' & GetChkSum(strSndMsg) & vbCr
    'strSndMsg = strSndMsg & vbCrLf
        
    SndMore = strSndMsg
    
End Function

'-- 검사결과 서버저장
Function SaveTransData(ByVal argSpcRow As Integer, ByVal SPD As Object) As Integer
    
    SaveTransData = -1
    
    Select Case gEMR
'        Case "AMIS"
'            SaveTransData = SaveTransData_AMIS(argSpcRow)
'
'        Case "BIGUBCARE"
'            SaveTransData = SaveTransData_BIGUBCARE(argSpcRow)
'
        Case "EONM"
            SaveTransData = SaveTransData_EONM(argSpcRow, SPD)
'
'        Case "BIT70"
'            SaveTransData = SaveTransData_BIT70(argSpcRow)
'
'        Case "EMEDI"
'            SaveTransData = SaveTransData_AMIS(argSpcRow)
'
'        Case "EONM"
'            SaveTransData = SaveTransData_EONM(argSpcRow)
'
        Case "EASYS"
            SaveTransData = SaveTransData_EASYS(argSpcRow, SPD)
'
'        Case "GINUS"
'            SaveTransData = SaveTransData_GINUS(argSpcRow)
'
'        Case "GSEN"
'            SaveTransData = SaveTransData_MSINFOTEC(argSpcRow)
'
'        Case "HWASAN"
'            SaveTransData = SaveTransData_HWASAN(argSpcRow)
'
'        Case "JAINCOM"
'            SaveTransData = SaveTransData_JAINCOM(argSpcRow)
'
'        Case "JWINFO"
'            SaveTransData = SaveTransData_JWINFO(argSpcRow)
'
'        Case "KCHART"
'            SaveTransData = SaveTransData_KCHART(argSpcRow)
'
'        Case "KOMAIN"
'            SaveTransData = SaveTransData_KOMAIN(argSpcRow)
'
'        Case "KYU"
'            SaveTransData = SaveTransData_KYU(argSpcRow)
'
'        Case "MEDICHART"
'            SaveTransData = SaveTransData_MEDICHART(argSpcRow)
'
'        Case "MEDIIT"
'            SaveTransData = SaveTransData_MEDIIT(argSpcRow)
'
'        Case "MEDITOLISS"
'            SaveTransData = SaveTransData_MEDITOLISS(argSpcRow)
'
'        Case "MCC"
'            SaveTransData = SaveTransData_MCC(argSpcRow)
'
'        Case "MOD"
'            SaveTransData = SaveTransData_MOD(argSpcRow)
'
'        Case "MSINFOTEC"
'            SaveTransData = SaveTransData_MSINFOTEC(argSpcRow)
'
'        Case "NEOSOFT"
'            SaveTransData = SaveTransData_NEOSOFT(argSpcRow)
'
'        Case "ONITGUM"
'            SaveTransData = SaveTransData_ONITGUM(argSpcRow)
'
'        Case "ONITEMR"
'            SaveTransData = SaveTransData_ONITEMR(argSpcRow)
'
'        Case "PLIS"
'            SaveTransData = SaveTransData_PLIS(argSpcRow)
'
'        Case "SY"
'            SaveTransData = SaveTransData_SY(argSpcRow)
'
'        Case "TWIN"
'            SaveTransData = SaveTransData_TWIN(argSpcRow)
'
'        Case "UBCARE"
'            SaveTransData = SaveTransData_UBCARE(argSpcRow)
'
'
'        Case Else
'            SaveTransData = -1
    End Select


End Function
                    
'-- 검사결과 서버저장
Function SaveTransDataR(ByVal argSpcRow As Integer) As Integer
    
    SaveTransDataR = -1
    
    Select Case gEMR
        Case "AMIS"
'            SaveTransDataR = SaveTransDataR_AMIS(argSpcRow)
        
'        Case "BIGUBCARE"
'            SaveTransDataR = SaveTransDataR_BIGUBCARE(argSpcRow)
'
'        Case "BIT"
'            SaveTransDataR = SaveTransDataR_BIT(argSpcRow)
'
'        Case "BIT70"
'            SaveTransDataR = SaveTransDataR_BIT70(argSpcRow)
'
'        Case "EMEDI"
'            SaveTransDataR = SaveTransDataR_AMIS(argSpcRow)
'
'        Case "EONM"
'            SaveTransDataR = SaveTransDataR_EONM(argSpcRow)
'
'        Case "EASYS"
'            SaveTransDataR = SaveTransDataR_EASYS(argSpcRow)
'
'        Case "GINUS"
'            SaveTransDataR = SaveTransDataR_GINUS(argSpcRow)
'
'        Case "GSEN"
'            SaveTransDataR = SaveTransDataR_MSINFOTEC(argSpcRow)
'
'        Case "HWASAN"
'            SaveTransDataR = SaveTransDataR_HWASAN(argSpcRow)
'
'        Case "JAINCOM"
'            SaveTransDataR = SaveTransDataR_JAINCOM(argSpcRow)
'
'        Case "JWINFO"
'            SaveTransDataR = SaveTransDataR_JWINFO(argSpcRow)
'
'        Case "KCHART"
'            SaveTransDataR = SaveTransDataR_KCHART(argSpcRow)
'
'        Case "KOMAIN"
'            SaveTransDataR = SaveTransDataR_KOMAIN(argSpcRow)
'
'        Case "KYU"
'            SaveTransDataR = SaveTransDataR_KYU(argSpcRow)
'
'        Case "MEDICHART"
'            SaveTransDataR = SaveTransDataR_MEDICHART(argSpcRow)
'
'        Case "MEDIIT"
'            SaveTransDataR = SaveTransDataR_MEDIIT(argSpcRow)
'
'        Case "MEDITOLISS"
'            SaveTransDataR = SaveTransDataR_MEDITOLISS(argSpcRow)
'
'        Case "MCC"
'            SaveTransDataR = SaveTransDataR_MCC(argSpcRow)
'
'        Case "MOD"
'            SaveTransDataR = SaveTransDataR_MOD(argSpcRow)
'
'        Case "MSINFOTEC"
'            SaveTransDataR = SaveTransDataR_MSINFOTEC(argSpcRow)
'
'        Case "NEOSOFT"
'            SaveTransDataR = SaveTransDataR_NEOSOFT(argSpcRow)
'
'        Case "ONITGUM"
'            SaveTransDataR = SaveTransDataR_ONITGUM(argSpcRow)
'
'        Case "ONITEMR"
'            SaveTransDataR = SaveTransDataR_ONITEMR(argSpcRow)
'
'        Case "PLIS"
'            SaveTransDataR = SaveTransDataR_PLIS(argSpcRow)
'
'        Case "SY"
'            SaveTransDataR = SaveTransDataR_SY(argSpcRow)
'
'        Case "TWIN"
'            SaveTransDataR = SaveTransDataR_TWIN(argSpcRow)
'
'        Case "UBCARE"
'            SaveTransDataR = SaveTransDataR_UBCARE(argSpcRow)

        
        Case Else
            SaveTransDataR = -1
    End Select

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


