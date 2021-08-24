VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISMicResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objMicSql As New clsLISSqlMicRst

Private mvarErrMsg As String

Public Property Get ErrMsg() As String
    ErrMsg = mvarErrMsg
End Property

Public Sub LoadNGRstCd(ByVal pCode As String, ByRef objList As Object, ByRef fNGCode() As Variant)
'pCode : 미생물 Worksheet코드 - 검체군정보의 field4에 정의
    Dim dsNG    As Recordset 'RecordsetobjMicSql.SqlUpdateOrderStatus
    Dim i       As Integer
    Dim j       As Integer

    ' Nogrowth result 로드
    Set dsNG = New Recordset
    dsNG.Open objMicSql.SQLGetNGRstCd(pCode), DBConn

    objList.Clear
    objList.AddItem LIS_Nothing
    
    If dsNG.EOF Then
        MsgBox "Nogrowth 결과코드가 등록되어 있지 않습니다.", vbExclamation, "기초자료 오류"
        Set dsNG = Nothing
        Exit Sub
    End If

    ReDim fNGCode(dsNG.RecordCount)

    fNGCode(0) = ""
    
    For i = 1 To dsNG.RecordCount
        objList.AddItem dsNG.Fields("Field1").Value
        fNGCode(i) = "" & dsNG.Fields("cdval2").Value
        dsNG.MoveNext
    Next i

    Set dsNG = Nothing

End Sub



'----------------------------------------------------------------------------------
'1. 미생물 Stain 결과등록 관련 메쏘드
'----------------------------------------------------------------------------------

Public Sub LoadRemark(ByRef objList As Object)
    Dim dsRM    As Recordset
    Dim i       As Integer
    Dim strTmp  As String

    ' 검체 Remark
    Set dsRM = New Recordset
    dsRM.Open objMicSql.SQLLoadRemark, DBConn

    objList.Clear
    objList.AddItem LIS_Nothing

    If dsRM.EOF Then Exit Sub

    dsRM.MoveFirst
    For i = 1 To dsRM.RecordCount
        strTmp = dsRM.Fields("cdval1").Value & Space(6)
        objList.AddItem Mid(strTmp, 1, 6) & dsRM.Fields("text1").Value
        dsRM.MoveNext
    Next i

    Set dsRM = Nothing

End Sub

Public Function GetRemark(ByVal pRmkCd As String) As String
    
    Dim dsRemark As Recordset

    Set dsRemark = New Recordset
    dsRemark.Open objMicSql.SqlGetRemark(pRmkCd), DBConn
    
    GetRemark = "" & dsRemark.Fields("text1").Value

    Set dsRemark = Nothing

End Function


Public Sub LoadWorksheetCode(ByVal lngCaseFg As Long, ByRef objList As Object, ByRef aryWS() As tpMicWorkSheet)
'lngCaseFg : 1-Nogrowth 배치, 2-Stain결과등록
    Dim dsSG        As Recordset
    Dim sqlSG       As String
    Dim sRstType    As String
    Dim sRTBuf      As String
    Dim i           As Integer
    Dim j           As Integer
    
    ' 검체군 로드
    Set dsSG = New Recordset
    dsSG.Open objMicSql.SQLGetWScode(lngCaseFg), DBConn

    If dsSG.EOF Then
        MsgBox "등록되어 있는 검체군이 없습니다."
        Set dsSG = Nothing
        Exit Sub
    End If

    ReDim aryWS(dsSG.RecordCount - 1)

    objList.Clear
    dsSG.MoveFirst
    For i = 1 To dsSG.RecordCount
        objList.AddItem dsSG.Fields("spcgrpnm").Value
        aryWS(i - 1).WsCode = "" & dsSG.Fields("spcgrp").Value
        aryWS(i - 1).WsName = "" & dsSG.Fields("spcgrpnm").Value
        aryWS(i - 1).WsType = "" & dsSG.Fields("wstype").Value
        sRTBuf = "" & dsSG.Fields("rsttype").Value
        
        sRstType = "": j = 1
        sRstType = medGetP(sRTBuf, j, ",")
        Do While (sRstType <> "")
           If j = 1 Then
              aryWS(i - 1).WsRstType = aryWS(i - 1).WsRstType & "'" & sRstType & "'"
           Else
              aryWS(i - 1).WsRstType = aryWS(i - 1).WsRstType & ",'" & sRstType & "'"
           End If
           j = j + 1: sRstType = medGetP(sRTBuf, j, ",")
        Loop
        dsSG.MoveNext
    Next i

    Set dsSG = Nothing

End Sub

Public Sub LoadMicWorkList(ByVal pWsCd As String, ByRef objList As Object, _
                           Optional ByVal pMfyFg As Boolean = False)
    Dim dsWSUnit    As Recordset
    Dim sSql        As String
    Dim sWsCd       As String
    Dim sWsUnit     As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim i, j        As Integer
    Dim bFlag       As Boolean

    Set dsWSUnit = New Recordset
    dsWSUnit.Open objMicSql.SQLReadWsUnit(pWsCd, pMfyFg), DBConn
    
    If dsWSUnit.EOF Then Set dsWSUnit = Nothing: Exit Sub
    
    objList.Clear: bFlag = False
    dsWSUnit.MoveFirst
    
    For i = 1 To dsWSUnit.RecordCount
        
        sWsCd = dsWSUnit.Fields("wscd").Value & ""
        sWsUnit = dsWSUnit.Fields("wsunit").Value & ""
        
        If LoadMicWorkListSub(sWsCd, sWsUnit) = True Then
            
            objList.AddItem dsWSUnit.Fields("wsunit").Value & "" & " " & _
                            dsWSUnit.Fields("workdt").Value & "" & " " & _
                            Format(dsWSUnit.Fields("worktm").Value & "", "0#:##:##")
                            
        End If
        
        dsWSUnit.MoveNext
    Next i
    
    Set dsWSUnit = Nothing
    
End Sub

'****************************************************************************************************CS
'   함수명      :   LoadMicWorkList_New
'   작성자      :   성원IT 개발부 부장 온승호 majestic@nate.com
'   작성일      :   2010.05.13
'   간략한 설명 :   조회조건 추가
'   인수        :   ByVal pWsCd As String
'                  ByVal pMonth As String
'                  ByRef objList As Object
'                  Optional ByVal pMfyFg As Boolean = False
'****************************************************************************************************CE
Public Sub LoadMicWorkList_New(ByVal pWsCd As String, ByVal pMonth As String, ByRef objList As Object, _
                           Optional ByVal pMfyFg As Boolean = False)
    Dim dsWSUnit    As Recordset
    Dim sSql        As String
    Dim sWsCd       As String
    Dim sWsUnit     As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim i, j        As Integer
    Dim bFlag       As Boolean

    Set dsWSUnit = New Recordset
    dsWSUnit.Open objMicSql.SQLReadWsUnit_New(pWsCd, pMonth, pMfyFg), DBConn
    
    If dsWSUnit.EOF Then Set dsWSUnit = Nothing: Exit Sub
    
    objList.Clear: bFlag = False
    dsWSUnit.MoveFirst
    
    For i = 1 To dsWSUnit.RecordCount
        
        sWsCd = dsWSUnit.Fields("wscd").Value & ""
        sWsUnit = dsWSUnit.Fields("wsunit").Value & ""
        
        If LoadMicWorkListSub(sWsCd, sWsUnit) = True Then
            
            objList.AddItem dsWSUnit.Fields("wsunit").Value & "" & " " & _
                            dsWSUnit.Fields("workdt").Value & "" & " " & _
                            Format(dsWSUnit.Fields("worktm").Value & "", "0#:##:##")
                            
        End If
        
        dsWSUnit.MoveNext
    Next i
    
    Set dsWSUnit = Nothing
    
End Sub

Private Function LoadMicWorkListSub(ByVal pWsCd As String, ByVal pWsUnit As String) As Boolean
    Dim Rs      As New ADODB.Recordset
    Dim strSql  As String
    
    strSql = objMicSql.SQLReadWsUnitSub(pWsCd, pWsUnit)
    
    Rs.Open strSql, DBConn, adOpenForwardOnly, adLockReadOnly
    
    LoadMicWorkListSub = IIf(Rs.EOF = False, True, False)
    
    Rs.Close
    Set Rs = Nothing
    
End Function

Public Sub DispWorksheetInfo(ByVal sWsCd As String, ByVal sWsUnit As String, _
                             ByRef strBltDate As String, ByRef strRcvDate As String)
    Dim dsWS    As Recordset
    Dim sqlWS   As String
    ' Work Sheet Info
    Set dsWS = New Recordset
    dsWS.Open objMicSql.SQLReadWSInfo(sWsCd, sWsUnit), DBConn

    ' 웤싵 정보 세팅
    If Not dsWS.EOF Then
        strBltDate = Format("" & dsWS.Fields("workdt").Value, "0000-00-00") & " " & _
                     Format("" & dsWS.Fields("worktm").Value, "00:00:00")
        strRcvDate = Format("" & dsWS.Fields("fnshdt").Value, "0000-00-00") & " " & _
                     Format("" & dsWS.Fields("fnshtm").Value, "00:00:00")
    End If
    
    Set dsWS = Nothing

End Sub

Public Sub DispWorksheetList(ByVal sWsCd As String, ByVal sWsUnit As String, _
                              ByRef objList As Object, _
                              Optional ByVal pMfySeq As Boolean = False)
    
    Dim dsSWS   As Recordset
    Dim sWA     As String
    Dim sWDt    As String
    Dim sWSeq   As String
    Dim iSWSCol As Integer
    
    objList.Clear
    Set dsSWS = New Recordset
    dsSWS.Open objMicSql.SQLReadWSList(sWsCd, sWsUnit, pMfySeq), DBConn
    
    If Not dsSWS.EOF Then
    
        Do Until dsSWS.EOF
            sWA = dsSWS.Fields("workarea").Value & ""
            sWDt = dsSWS.Fields("accdt").Value & ""
            sWDt = Mid$(sWDt, 3, Len(sWDt) - 2)
            sWSeq = dsSWS.Fields("accseq").Value & ""
            
            If LoadDCCheck(sWA, dsSWS.Fields("accdt").Value & "", sWSeq) = False Then
                objList.AddItem "" & sWA & "-" & sWDt & "-" & sWSeq
            End If
            
           dsSWS.MoveNext
        Loop

    End If
    Set dsSWS = Nothing

End Sub
Public Function LoadDCCheck(ByVal Workarea As String, ByVal Accdt As String, ByVal Accseq As String) As Boolean
    Dim Rs   As Recordset
    Dim sSql As String
    'Dc안난거 찾기
    sSql = "select * from " & T_LAB102 & " where " & DBW("workarea=", Workarea) & _
           " and " & DBW("accdt=", Accdt) & " and " & DBW("accseq=", Accseq) & _
           " and (dcfg='1')"
    
    Set Rs = New Recordset
    Rs.Open sSql, DBConn
    
    If Rs.EOF Then
        '다Dc가 안났어
        LoadDCCheck = False
    Else
        'dc가 났어
        LoadDCCheck = True
    End If
    Set Rs = Nothing

End Function

Public Sub LoadStainRstCd(ByVal pTestCd As String, ByRef objList1 As Object, ByRef objList2 As Object)
'Stain 결과코드 리스트 로드 : objList1 - 배치결과, objList2 - 일반결과

    Dim dsRstCd     As Recordset
    Dim sqlRstCd    As String
    Dim i           As Integer
    Dim iLen        As Integer
    
    Set dsRstCd = New Recordset
    dsRstCd.Open objMicSql.SQLGetStainRstCd(pTestCd), DBConn

    objList1.Clear: objList2.Clear
    If Not dsRstCd.EOF Then dsRstCd.MoveFirst
    Do Until dsRstCd.EOF
        If dsRstCd.Fields("field3").Value = "B" Then
            objList1.AddItem dsRstCd.Fields("cdval2").Value & Chr$(9) & dsRstCd.Fields("field1").Value
        Else
            objList2.AddItem dsRstCd.Fields("cdval2").Value & Chr$(9) & dsRstCd.Fields("field1").Value
        End If
        dsRstCd.MoveNext
    Loop

    Set dsRstCd = Nothing

End Sub

Public Function ResultCheck(ByVal pTestCd As String, ByVal pRstCd As String, ByRef pRstNm As String) As Boolean
    Dim dsRstCd     As Recordset
    Dim sqlRstCd    As String
    
    Set dsRstCd = New Recordset
    dsRstCd.Open objMicSql.SQLResultCheck(pTestCd, pRstCd), DBConn
    ResultCheck = Not dsRstCd.EOF
    
    If Not dsRstCd.EOF Then
        pRstNm = dsRstCd.Fields("field1").Value
    Else
        pRstNm = ""
    End If
    Set dsRstCd = Nothing
    
End Function

Public Function SaveStainResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                ByVal objRstTbl As Object, ByVal pWsCd As String, ByVal pEmpId As String, _
                                ByVal pFNote As String, ByVal pRmkCd As String) As Boolean
    Dim blnVerified     As Boolean
    Dim sDate           As String
    Dim sTime           As String
    Dim sTestCd         As String
    Dim sReqFg          As String
    Dim sDetSeq         As String
    Dim sResult         As String
    Dim sTestCd3        As String
    Dim sReqFg2         As String
    Dim sDetSeq2        As String
    Dim sResult2        As String
    Dim sResult3        As String
    Dim strStsCd        As String
    Dim strOrdInfo      As String
    Dim sPtId           As String
    Dim sOrdDt          As String
    Dim sOrdNo          As String
    Dim sOrdSeq         As String
    Dim strSql          As String
    Dim sqlUpdate()     As String
    Dim chkRequired     As Integer
    Dim chkAlternative1 As Integer
    Dim chkAlternative2 As Integer
    Dim iVerifyCount    As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim R1              As Integer
    Dim R2              As Integer
    Dim R3              As Integer
    Dim sHold           As Long

    '-- 전주예수병원 추가 변수=============
    Dim SqlTransfer()   As String
    Dim intTempSql      As Integer
    Dim clsTransfer     As New clsLISTransfer
    Dim strVfyDt        As String
    Dim strRstVal       As String
    Dim strOrdSQL       As String
    Dim strKey1         As String
    Dim strKey2         As String
    Dim strOCSNote      As String
    Dim strFromRef      As String
    Dim strToRef        As String
    Dim strRemark       As String
    '======================================
    
    ' 시스템 일자/시간 설정
    sDate = Format(GetSystemDate, CS_DateDbFormat)
    sTime = Format(GetSystemDate, CS_TimeDbFormat)
    
    ReDim sqlUpdate(0)
    
    '-- 전주예수병원 추가 변수=============
    intTempSql = 0: ReDim SqlTransfer(0)
    '======================================
    
    R1 = 1: iVerifyCount = 0
    
On Error GoTo DBExecError
    
    DBConn.BeginTrans
    
    With objRstTbl
        
        Do Until R1 > .MaxRows
            
            .Row = R1
            .Col = enSTAIN.tcTESTCD:   sTestCd = .Text
            .Col = enSTAIN.tcRSTDIV:   sReqFg = .Text
            .Col = enSTAIN.tcDETAIL:   sDetSeq = .Text
            .Col = enSTAIN.tcRSTCD:    sResult = Trim(.Text)
            ' 보류여부란이 static text인 경우, 처리과정에서 무시되므로 그냥 읽음
            .Col = enSTAIN.tcEXCPT:    sHold = .Value
            
            If sReqFg = "*" Then
                
                ' 당연히 Detail 항목은 대표항목 아래에 연속해서 존재하는 것으로 가정
                ' Display 할때 그렇게 처리함. 만약 아니라면 체크루틴 수정 필요
                chkRequired = 0: chkAlternative1 = 0: chkAlternative2 = 0
                R2 = R1 + 1
                For R2 = R1 + 1 To .MaxRows
                    .Row = R2
                    .Col = enSTAIN.tcRSTDIV: sReqFg2 = .Text
                    .Col = enSTAIN.tcDETAIL: sDetSeq2 = .Text
                    .Col = enSTAIN.tcRSTCD:  sResult2 = Trim(.Text)
                    If sDetSeq2 = sDetSeq Then
                        If sReqFg2 = LIS_RST_REQUIRED And sResult2 = "" Then chkRequired = chkRequired + 1
                        If sReqFg2 = LIS_RST_ALTERNATIVE And sResult2 = "" Then chkAlternative1 = chkAlternative1 + 1
                        If sReqFg2 = LIS_RST_ALTERNATIVE And sResult2 <> "" Then chkAlternative2 = chkAlternative2 + 1
                    Else
                        Exit For
                    End If
                Next R2
                
                blnVerified = IsAllVerified(pWorkArea, pAccDt, pAccSeq, sTestCd, strOrdInfo)
                sPtId = medGetP(strOrdInfo, 1, COL_DIV)
                sOrdDt = medGetP(strOrdInfo, 2, COL_DIV)
                sOrdNo = medGetP(strOrdInfo, 3, COL_DIV)
                sOrdSeq = medGetP(strOrdInfo, 4, COL_DIV)
                
                
                ' 대표항목,세부항목 일괄 처리
                If chkRequired = 0 And (chkAlternative1 = 0 Or chkAlternative2 > 0) And sHold = 0 Then                ' Verify
                    
                    strSql = objMicSql.SQLVerifyStainResult(pWorkArea, pAccDt, pAccSeq, sTestCd, "", _
                                                       enStsCd.StsCd_LIS_FinRst, sDate, sTime, pEmpId)
                    DBConn.Execute strSql
                    
                    '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.16=======================
                    '-- 결과전송 시 필요정보
                    Call clsTransfer.SqlOCS_Info_Mic(pWorkArea, pAccDt, pAccSeq)
                    
                    ' - 상세모항목 Insert => 수정 모항목 제외 OCS 요청 2004.11.10
'                    strVfyDt = sDate & sTime
'
'                    strSql = clsTransfer.SqlOCSRESULT_INSERT(sPtId, sOrdDt, sOrdNo, _
'                            sTestCd, clsTransfer.SlipCd, "", "", "", "", "", _
'                            "", clsTransfer.RcvDt & clsTransfer.RCVTM, strVfyDt, pEmpId, _
'                            strVfyDt, pEmpId, "", "", "", "", "")
'
'                    If strSql <> "" Then DBConn.Execute strSql
                    
                    '=================================================================================
                    
                    For R3 = R1 + 1 To R2 - 1
                        .Row = R3
                        .Col = enSTAIN.tcTESTCD: sTestCd3 = .Text
                        .Col = enSTAIN.tcRSTCD:  sResult3 = Trim(.Text)
                        ' 세부항목 결과 저장
'                        ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
                        strSql = objMicSql.SQLVerifyStainResult(pWorkArea, pAccDt, pAccSeq, sTestCd3, _
                                                        sResult3, enStsCd.StsCd_LIS_FinRst, sDate, sTime, pEmpId)
                        DBConn.Execute strSql
                        
                        '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.16
                        '==========================================================================
                        intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
                        
                        Call clsTransfer.ResultValue_Find(sTestCd3, sResult3)
                        
                        If pFNote <> "" Then
                            strOCSNote = "[FootNote]" & Chr(13) & pFNote & Chr(13)
                        Else
                            strOCSNote = ""
                        End If
                        
                        If pRmkCd <> "" Then
                            strRemark = clsTransfer.SqlGetRemark(pRmkCd)
                            If strRemark <> "" Then
                                strOCSNote = strOCSNote & "◈ Remark" & Chr(13) & strRemark
                            End If
                        End If
                        
                        strVfyDt = sDate & sTime
                        
                        '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
                        If strKey1 <> sPtId & sOrdDt & sOrdNo Then
                            If clsTransfer.WrkDiv = "3" Then
                                If clsTransfer.SqlOCSSU2EXAMT_Find(sPtId, sOrdDt, sOrdNo) = True Then
                                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE_New(sPtId, sOrdDt, sOrdNo, _
                                                            "N", clsTransfer.ItemVal, strOCSNote, "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "", enStsCd.StsCd_LIS_FinRst)
                                                            
                                    
                                Else
                                    GoTo DBExecError
                                End If
                            ElseIf clsTransfer.WrkDiv = "4" Then
                                    If clsTransfer.SqlOCSSG2EXAMT_Find(sPtId, sOrdDt, sOrdNo) = True Then
                                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE_NEW(sPtId, sOrdDt, sOrdNo, _
                                                                "N", clsTransfer.ItemVal, strOCSNote, "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "", enStsCd.StsCd_LIS_FinRst)
                                    Else
                                        GoTo DBExecError
                                    End If
                            Else
                                If clsTransfer.SqlOCSMDEXMORT_Find(sPtId, sOrdDt, sOrdNo) = True Then
                                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE_NEW(sPtId, sOrdDt, sOrdNo, _
                                                            "N", strVfyDt, pEmpId, "", "", "")
                                Else
                                    GoTo DBExecError
                                End If
                            End If
                            
                            strKey1 = sPtId & sOrdDt & sOrdNo
                            
                        End If
                        
                        '** 종건/일건 이외의 검사만 등록.
                        If clsTransfer.WrkDiv <> "" Then
                            GoTo OCSSKIP1
                        End If
        
                        '-- 2. 건진외결과 INSERT
                        '** 검사항목별 결과값 체크
                        intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
                        
                        If clsTransfer.SqlOCSRESULT_FLAG(sPtId, sOrdDt, sOrdNo, sTestCd3) = False Then
                            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(sPtId, sOrdDt, sOrdNo, _
                                                    sTestCd3, clsTransfer.SlipCd, clsTransfer.ItemVal, "", "", "", "", _
                                                    strOCSNote, clsTransfer.RcvDt & clsTransfer.RCVTM, strVfyDt, pEmpId, strVfyDt, pEmpId, "", clsTransfer.RsltType, _
                                                    "", "", "")
                        Else
                            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_SUB_UPDATE(sPtId, sOrdDt, sOrdNo, _
                                                    sTestCd3, clsTransfer.ItemVal, "", "", strVfyDt, pEmpId)
                        End If
                        
'                        If clsTransfer.SqlOCSRESULT_FLAG(sPtId, sOrdDt, sOrdNo, sTestCd) = False Then
'                            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
'
'                            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(sPtId, sOrdDt, sOrdNo, _
'                                                    sTestCd3, pWorkArea, strRstVal, "", "", "", "", _
'                                                    "", "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "", "", "")
'                        Else
'                            GoTo DBExecError
'                        End If
                        
OCSSKIP1:

                        For intTempSql = LBound(SqlTransfer) To UBound(SqlTransfer)
                            intTempSql = intTempSql
                            
                            If Trim(SqlTransfer(intTempSql)) <> "" Then DBConn.Execute (SqlTransfer(intTempSql))
                            
                        Next intTempSql
                        
                        Erase SqlTransfer
                        intTempSql = 0: ReDim SqlTransfer(0)
                        '//////////////////////////////////////////////////////////////////////////////
                        
                    Next R3
                    iVerifyCount = iVerifyCount + 1
                Else                                                            ' Only Save
                    For R3 = R1 To R2 - 1
                        .Row = R3
                        .Col = enSTAIN.tcTESTCD: sTestCd3 = .Text
                        .Col = enSTAIN.tcRSTCD:  sResult3 = Trim(.Text)
                        ' 세부항목 결과 저장
                        'ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
                        strSql = objMicSql.SQLVerifyStainResult(pWorkArea, pAccDt, pAccSeq, sTestCd3, _
                                                           sResult3, enStsCd.StsCd_LIS_InProcess, sDate, sTime, pEmpId)
                        If strSql <> "" Then
                            DBConn.Execute strSql
                        End If
                        
                    Next R3
                End If
                
                blnVerified = IsAllVerified(pWorkArea, pAccDt, pAccSeq, sTestCd, strOrdInfo)
                sPtId = medGetP(strOrdInfo, 1, COL_DIV)
                sOrdDt = medGetP(strOrdInfo, 2, COL_DIV)
                sOrdNo = medGetP(strOrdInfo, 3, COL_DIV)
                sOrdSeq = medGetP(strOrdInfo, 4, COL_DIV)
                
                ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
                If blnVerified Then
                    sqlUpdate(UBound(sqlUpdate)) = objMicSql.SqlUpdateOrderStatus(sPtId, sOrdDt, sOrdNo, sOrdSeq, _
                                                   enStsCd.StsCd_LIS_FinRst, sDate, sTime, pEmpId)
                Else
                    sqlUpdate(UBound(sqlUpdate)) = objMicSql.SqlUpdateOrderStatus(sPtId, sOrdDt, sOrdNo, sOrdSeq, _
                                                   enStsCd.StsCd_LIS_MidRst, sDate, sTime, pEmpId)
                End If
                ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
                sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLUpdateWoksheet(2, pWorkArea, pAccDt, pAccSeq, _
                                               pWsCd, "", MWS_Holding)
                
                R1 = R2
            Else
                
                ' 혹시 대표항목 없이 세부항목이 먼저 발견되면..
                If sDetSeq <> "" Then
                    mvarErrMsg = "검사항목 Display 순서에 오류가 있습니다. 전산실로 연락 바랍니다. (☎" & ObjSysInfo.HelpLine & ")"
                    GoTo DBExecError
                End If
                ' 세부항목을 가지는 대표항목과는 별개로 데이타 있으면 무조건 저장하고
                ' status 반영하고, 단 세부아이템은 그냥 데이타만 저장
                ' *** 결과 저장 & Status 반영(결과내역,처방바디)
                If sResult <> "" And sDetSeq = "" Then
                
                    strStsCd = IIf(sHold = 0, enStsCd.StsCd_LIS_FinRst, enStsCd.StsCd_LIS_InProcess)
                    
                    'ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
                    strSql = objMicSql.SQLVerifyStainResult(pWorkArea, pAccDt, pAccSeq, sTestCd, _
                                                   sResult, strStsCd, sDate, sTime, pEmpId)
                    DBConn.Execute strSql
                    
                    blnVerified = IsAllVerified(pWorkArea, pAccDt, pAccSeq, sTestCd, strOrdInfo)
                    sPtId = medGetP(strOrdInfo, 1, COL_DIV)
                    sOrdDt = medGetP(strOrdInfo, 2, COL_DIV)
                    sOrdNo = medGetP(strOrdInfo, 3, COL_DIV)
                    sOrdSeq = medGetP(strOrdInfo, 4, COL_DIV)
                    
                    '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.16
                    
                    '-- 결과전송 시 필요정보
                    Call clsTransfer.SqlOCS_Info_Mic(pWorkArea, pAccDt, pAccSeq)
                    
                    intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
                    
                    Call clsTransfer.ResultValue_Find(sTestCd, sResult)
                    
                    If pFNote <> "" Then
                        strOCSNote = "[FootNote]" & Chr(13) & pFNote & Chr(13)
                    Else
                        strOCSNote = ""
                    End If
                    
                    If pRmkCd <> "" Then
                        strRemark = clsTransfer.SqlGetRemark(pRmkCd)
                        If strRemark <> "" Then
                            strOCSNote = strOCSNote & "◈ Remark" & Chr(13) & strRemark
                        End If
                    End If
                        
                    strVfyDt = sDate & sTime
                    
                    '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
                    If strKey2 <> sPtId & sOrdDt & sOrdNo Then
                        If clsTransfer.WrkDiv = "3" Then
                            If clsTransfer.SqlOCSSU2EXAMT_Find(sPtId, sOrdDt, sOrdNo) = True Then
                                SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE(sPtId, sOrdDt, sOrdNo, _
                                                        "N", clsTransfer.ItemVal, strOCSNote, "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "")
                                
                                DBConn.Execute (SqlTransfer(intTempSql))
                            Else
                                GoTo DBExecError
                            End If
                        ElseIf clsTransfer.WrkDiv = "4" Then
                                If clsTransfer.SqlOCSSG2EXAMT_Find(sPtId, sOrdDt, sOrdNo) = True Then
                                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE(sPtId, sOrdDt, sOrdNo, _
                                                            "N", clsTransfer.ItemVal, strOCSNote, "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "")
                                                            
                                    DBConn.Execute (SqlTransfer(intTempSql))
                                Else
                                    GoTo DBExecError
                                End If
                        Else
                            If clsTransfer.SqlOCSMDEXMORT_Find(sPtId, sOrdDt, sOrdNo) = True Then
                                SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE(sPtId, sOrdDt, sOrdNo, _
                                                        "N", strVfyDt, pEmpId, "", "")
                            Else
                                GoTo DBExecError
                            End If
                        End If
                        
                        strKey2 = sPtId & sOrdDt & sOrdNo
                        
                    End If
                    
                    '** 종건/일건 이외의 검사만 등록.
                    If clsTransfer.WrkDiv <> "" Then
                        GoTo OCSSKIP2
                    End If
                        
                    '-- 2. 건진외결과 INSERT
                    '** 검사항목별 결과값 체크
                    Call clsTransfer.ResultValue_Find(sTestCd, sResult)
                    
                    intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
                    
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(sPtId, sOrdDt, sOrdNo, _
                                            sTestCd, clsTransfer.SlipCd, clsTransfer.ItemVal, "", "", "", "", _
                                            strOCSNote, clsTransfer.RcvDt & clsTransfer.RCVTM, strVfyDt, pEmpId, strVfyDt, pEmpId, "", clsTransfer.RsltType, _
                                            "", "", "")
                                                
'                    If clsTransfer.SqlOCSRESULT_FLAG(sPtId, sOrdDt, sOrdNo, sTestCd) = False Then
'                        intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
'
'                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(sPtId, sOrdDt, sOrdNo, _
'                                                sTestCd3, pWorkArea, strRstVal, "", "", "", "", _
'                                                "", "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "", "", "")
'                    Else
'                        GoTo DBExecError
'                    End If
                    
OCSSKIP2:
                    For intTempSql = LBound(SqlTransfer) To UBound(SqlTransfer)
                        intTempSql = intTempSql
                        
                        If Trim(SqlTransfer(intTempSql)) <> "" Then DBConn.Execute (SqlTransfer(intTempSql))
                        
                    Next intTempSql

                    Erase SqlTransfer
                    intTempSql = 0: ReDim SqlTransfer(0)
                    '//////////////////////////////////////////////////////////////////////////////
                    
                    
                    blnVerified = IsAllVerified(pWorkArea, pAccDt, pAccSeq, sTestCd, strOrdInfo)
                    sPtId = medGetP(strOrdInfo, 1, COL_DIV)
                    sOrdDt = medGetP(strOrdInfo, 2, COL_DIV)
                    sOrdNo = medGetP(strOrdInfo, 3, COL_DIV)
                    sOrdSeq = medGetP(strOrdInfo, 4, COL_DIV)
                    
                    
                    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
                    If blnVerified Then
                        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SqlUpdateOrderStatus(sPtId, sOrdDt, sOrdNo, sOrdSeq, _
                                                       enStsCd.StsCd_LIS_FinRst, sDate, sTime, pEmpId)
                    Else
                        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SqlUpdateOrderStatus(sPtId, sOrdDt, sOrdNo, sOrdSeq, _
                                                       enStsCd.StsCd_LIS_MidRst, sDate, sTime, pEmpId)
                    End If
                    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
                    sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLUpdateWoksheet(2, pWorkArea, pAccDt, pAccSeq, _
                                                   pWsCd, "", MWS_Holding, sHold)
                    If sHold = 0 Then iVerifyCount = iVerifyCount + 1
                    
                End If
                R1 = R1 + 1
            End If
    
        Loop
        
    End With
    
    ' *** Parent Status 반영 (접수내역, 미생물 WS)
    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
    sqlUpdate(UBound(sqlUpdate)) = SavePStatus(pWorkArea, pAccDt, pAccSeq, iVerifyCount, enStsCd.StsCd_LIS_FinRst, _
                                               sDate, sTime, pEmpId)

    ' *** Foot Note 저장 *** 검체 Remark 저장
    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
    sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLSaveFootNote(pWorkArea, pAccDt, pAccSeq, 0, "", "", 0)
    If Trim(pFNote) <> "" Then
        ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLSaveFootNote(pWorkArea, pAccDt, pAccSeq, 1, pEmpId, pFNote, 1)
        
        '-- 전주예수병원 추가 루틴=============
        intTempSql = 0: ReDim SqlTransfer(0)
        '======================================
    End If
    
    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
    sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLSaveRemark(pWorkArea, pAccDt, pAccSeq, 1, pRmkCd)

    For i = 1 To UBound(sqlUpdate)
        If Trim(sqlUpdate(i)) <> "" Then DBConn.Execute sqlUpdate(i)
    Next
    DBConn.CommitTrans
    
    SaveStainResult = True
    Exit Function

DBExecError:
    SaveStainResult = False
    DBConn.RollbackTrans
    
End Function

Private Function SavePStatus(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                             ByVal pVCnt As Integer, ByVal pStatus As String, ByVal pDate As String, _
                             ByVal pTime As String, ByVal pEmpId As String) As String
    Dim dsAcc           As Recordset
    Dim sqlAcc          As String
    Dim sqlUpdAcc       As String
    Dim sStatus         As String
    Dim upStatus        As String
    Dim iTotalCount     As Integer
    Dim iInputCount     As Integer
    Dim upInputCount    As Integer
    

    ' 중간결과를 사용하지 않는 검사인 경우에 성립
    ' 만약 감수성 검사 등의 경우라면 중간결과시 카운터에는 반영않지만 Status는 반영함

    If pStatus < enStsCd.StsCd_LIS_FinRst Then Exit Function

    ' 검사종류 확인
    Set dsAcc = New Recordset
    dsAcc.Open objMicSql.SqlGetTestCount(pWorkArea, pAccDt, pAccSeq), DBConn

    iTotalCount = Val("" & dsAcc.Fields("reqtotcnt").Value)
    iInputCount = Val("" & dsAcc.Fields("reqinputcnt").Value)
    sStatus = enStsCd.StsCd_LIS_MidRst  '"" & dsAcc.Fields("stscd").Value

    Set dsAcc = Nothing

    If iTotalCount <= iInputCount + pVCnt Then sStatus = pStatus                    ' 먼저 체크
    If pStatus = enStsCd.StsCd_LIS_FinRst Then iInputCount = iInputCount + pVCnt    ' 다음 체크 (순서 중요)

    SavePStatus = objMicSql.SQLUpdateStatus(pWorkArea, pAccDt, pAccSeq, sStatus, iInputCount, pDate, pTime, pEmpId)
    
End Function


Public Function DispSenDataByTestCd(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                    ByVal pAccSeq As String, ByVal pTestCd As String) As clsDictionary
    Dim dsPtInfo    As Recordset
    Dim dsFNote     As Recordset
    Dim sqlPtInfo   As String
    Dim sqlFNote    As String
    Dim sRemarkCd   As String
    Dim sRemarkIdx  As Integer

    Set DispSenDataByTestCd = Nothing

    Set dsPtInfo = New Recordset
    dsPtInfo.Open objMicSql.SQLGetPtInfoByTestCd(pWorkArea, pAccDt, pAccSeq, pTestCd), DBConn

    If dsPtInfo.EOF Then
        Set dsPtInfo = Nothing
        Exit Function
    End If

    Set DispSenDataByTestCd = New clsDictionary
    
    With DispSenDataByTestCd
        .Clear
' 08.11.07 양성현 처방의사, 처방Remark, 진료과명, 연락전화번호 등 추가 및 변경
        .FieldInialize "ptid", "ptnm,sex,ageday,sexage,deptcd,wardid,hosilid,location,spcnm," & _
                               "mfyseq,footnotefg,rmkcd,stscd,testcd,rsttype,majdoct,spccd,orddrnm,mesg,deptnm,phone,bedindt,orddoct,rcvdt"
        
        .AddNew "" & dsPtInfo.Fields("ptid").Value, ""
        .Fields("ptnm") = "" & dsPtInfo.Fields("ptnm").Value
        .Fields("sex") = "" & dsPtInfo.Fields("sex").Value
        .Fields("ageday") = "" & dsPtInfo.Fields("ageday").Value
        .Fields("sexage") = .Fields("sex") & "/" & Str$((Val("" & .Fields("ageday")) \ 365) + 1)
        .Fields("deptcd") = "" & dsPtInfo.Fields("deptcd").Value
        .Fields("wardid") = "" & dsPtInfo.Fields("wardid").Value
        .Fields("hosilid") = "" & dsPtInfo.Fields("hosilid").Value
        .Fields("location") = "" & dsPtInfo.Fields("wardid").Value & " - " & dsPtInfo.Fields("hosilid").Value '& " - " & dsPtInfo.Fields("bedid").Value
        .Fields("spcnm") = "" & dsPtInfo.Fields("spcnm").Value
        .Fields("spccd") = "" & dsPtInfo.Fields("spccd").Value
    
        .Fields("mfyseq") = "" & dsPtInfo.Fields("mfyseq").Value
        .Fields("footnotefg") = "" & dsPtInfo.Fields("footnotefg").Value
        .Fields("rmkcd") = "" & dsPtInfo.Fields("rmkcd").Value
        .Fields("stscd") = "" & dsPtInfo.Fields("stscd").Value
        .Fields("testcd") = "" & dsPtInfo.Fields("testcd").Value
        .Fields("rsttype") = "" & dsPtInfo.Fields("rsttype").Value
        .Fields("orddoct") = "" & dsPtInfo.Fields("orddoct").Value

' 08.11.07 양성현 처방의사, 처방Remark, 진료과명, 연락전화번호 등 추가 및 변경
        .Fields("orddrnm") = "" & dsPtInfo.Fields("orddrnm").Value
        .Fields("mesg") = "" & dsPtInfo.Fields("mesg").Value
        .Fields("deptnm") = "" & dsPtInfo.Fields("deptnm").Value
        .Fields("phone") = "" & dsPtInfo.Fields("phone").Value
        .Fields("bedindt") = "" & dsPtInfo.Fields("bedindt").Value
        .Fields("majdoct") = "" & dsPtInfo.Fields("majdoct").Value
        .Fields("rcvdt") = "" & dsPtInfo.Fields("rcvdt").Value
    End With
    
    Set dsPtInfo = Nothing

End Function



Public Function DispPtInfoByLabno(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                  ByVal pAccSeq As String, Optional ByVal pRstTp As String = "") As clsDictionary
    Dim dsPtInfo    As Recordset
    Dim sqlPtInfo   As String
    Dim sqlFNote    As String
    Dim sRemarkCd   As String
    Dim sRemarkIdx  As Integer
    
    Set DispPtInfoByLabno = Nothing

    Set dsPtInfo = New Recordset
    dsPtInfo.Open objMicSql.SQLGetPtInfoByLabno(pWorkArea, pAccDt, pAccSeq, pRstTp), DBConn

    If dsPtInfo.EOF Then
        Set dsPtInfo = Nothing
        Exit Function
    End If

    Set DispPtInfoByLabno = New clsDictionary
    
    With DispPtInfoByLabno
        .Clear
' 08.11.07 양성현 처방의사, 처방Remark, 진료과명, 연락전화번호 등 추가 및 변경
        .FieldInialize "ptid", "ptnm,sex,ageday,sexage,deptcd,wardid,hosilid,location,spcnm," & _
                               "mfyseq,footnotefg,rmkcd,stscd,testcd,rsttype,majdoct,orddrnm,mesg,deptnm,phone,orddoct,rcvdt"
        
        .AddNew "" & dsPtInfo.Fields("ptid").Value, ""
        .Fields("ptnm") = "" & dsPtInfo.Fields("ptnm").Value
        .Fields("sex") = "" & dsPtInfo.Fields("sex").Value
        .Fields("ageday") = "" & dsPtInfo.Fields("ageday").Value
        .Fields("sexage") = .Fields("sex") & "/" & Str$((Val("" & .Fields("ageday")) \ 365) + 1)
        .Fields("deptcd") = "" & dsPtInfo.Fields("deptcd").Value
        .Fields("wardid") = "" & dsPtInfo.Fields("wardid").Value
        .Fields("hosilid") = "" & dsPtInfo.Fields("hosilid").Value
        .Fields("location") = "" & dsPtInfo.Fields("wardid").Value & " - " & dsPtInfo.Fields("hosilid").Value '& " - " & dsPtInfo.Fields("bedid").Value
        .Fields("spcnm") = "" & dsPtInfo.Fields("spcnm").Value
    
        .Fields("mfyseq") = "" & dsPtInfo.Fields("mfyseq").Value
        .Fields("footnotefg") = "" & dsPtInfo.Fields("footnotefg").Value
        .Fields("rmkcd") = "" & dsPtInfo.Fields("rmkcd").Value
        .Fields("stscd") = "" & dsPtInfo.Fields("stscd").Value
        .Fields("testcd") = "" & dsPtInfo.Fields("testcd").Value
        .Fields("rsttype") = "" & dsPtInfo.Fields("rsttype").Value
        .Fields("orddoct") = "" & dsPtInfo.Fields("orddoct").Value
        
' 08.11.07 양성현 처방의사, 처방Remark, 진료과명, 연락전화번호 등 추가 및 변경
        .Fields("orddrnm") = "" & dsPtInfo.Fields("orddrnm").Value
        .Fields("mesg") = "" & dsPtInfo.Fields("mesg").Value
        .Fields("deptnm") = "" & dsPtInfo.Fields("deptnm").Value
        .Fields("phone") = "" & dsPtInfo.Fields("phone").Value
        .Fields("majdoct") = "" & dsPtInfo.Fields("majdoct").Value
        .Fields("rcvdt") = "" & dsPtInfo.Fields("rcvdt").Value
    End With
    
    Set dsPtInfo = Nothing

End Function

Public Function DispPtInfoByBarno(ByVal pSpcYY As String, ByVal pSpcNo As String, _
                                  Optional ByVal pRstTp As String = "") As clsDictionary
    Dim dsPtInfo    As Recordset
    Dim sqlPtInfo   As String
    Dim sqlFNote    As String
    Dim sRemarkCd   As String
    Dim sRemarkIdx  As Integer
    
    Set DispPtInfoByBarno = Nothing

    Set dsPtInfo = New Recordset
    dsPtInfo.Open objMicSql.SQLGetPtInfoByBarno(pSpcYY, pSpcNo, pRstTp), DBConn

    If dsPtInfo.EOF Then
        Set dsPtInfo = Nothing
        Exit Function
    End If

    Set DispPtInfoByBarno = New clsDictionary
    
    With DispPtInfoByBarno
        .Clear
' 08.11.07 양성현 처방의사, 처방Remark, 진료과명, 연락전화번호 등 추가 및 변경
        .FieldInialize "ptid", "ptnm,sex,ageday,sexage,deptcd,wardid,hosilid,location,spcnm," & _
                               "mfyseq,footnotefg,rmkcd,stscd,testcd,rsttype,majdoct,orddrnm,mesg,deptnm,phone,workarea,accdt,accseq"
        
        .AddNew "" & dsPtInfo.Fields("ptid").Value, ""
        .Fields("ptnm") = "" & dsPtInfo.Fields("ptnm").Value
        .Fields("sex") = "" & dsPtInfo.Fields("sex").Value
        .Fields("ageday") = "" & dsPtInfo.Fields("ageday").Value
        .Fields("sexage") = .Fields("sex") & "/" & Str$((Val("" & .Fields("ageday")) \ 365) + 1)
        .Fields("deptcd") = "" & dsPtInfo.Fields("deptcd").Value
        .Fields("wardid") = "" & dsPtInfo.Fields("wardid").Value
        .Fields("hosilid") = "" & dsPtInfo.Fields("hosilid").Value
        .Fields("location") = "" & dsPtInfo.Fields("wardid").Value & " - " & dsPtInfo.Fields("hosilid").Value '& " - " & dsPtInfo.Fields("bedid").Value
        .Fields("spcnm") = "" & dsPtInfo.Fields("spcnm").Value
    
        .Fields("mfyseq") = "" & dsPtInfo.Fields("mfyseq").Value
        .Fields("footnotefg") = "" & dsPtInfo.Fields("footnotefg").Value
        .Fields("rmkcd") = "" & dsPtInfo.Fields("rmkcd").Value
        .Fields("stscd") = "" & dsPtInfo.Fields("stscd").Value
        .Fields("testcd") = "" & dsPtInfo.Fields("testcd").Value
        .Fields("rsttype") = "" & dsPtInfo.Fields("rsttype").Value
        .Fields("majdoct") = "" & dsPtInfo.Fields("majdoct").Value
        .Fields("orddoct") = "" & dsPtInfo.Fields("orddoct").Value
        
' 08.11.07 양성현 처방의사, 처방Remark, 진료과명, 연락전화번호 등 추가 및 변경
        .Fields("orddrnm") = "" & dsPtInfo.Fields("orddrnm").Value
        .Fields("mesg") = "" & dsPtInfo.Fields("mesg").Value
        .Fields("deptnm") = "" & dsPtInfo.Fields("deptnm").Value
        .Fields("phone") = "" & dsPtInfo.Fields("phone").Value
        .Fields("workarea") = "" & dsPtInfo.Fields("workarea").Value
        .Fields("accdt") = "" & dsPtInfo.Fields("accdt").Value
        .Fields("accseq") = "" & dsPtInfo.Fields("accseq").Value
    End With
    
    Set dsPtInfo = Nothing

End Function

Public Function DispFootNote(ByVal pWorkArea As String, ByVal pAccDt As String, _
                             ByVal pAccSeq As String) As String
    Dim dsFNote     As Recordset
    Dim sqlFNote    As String

    ' Foot Note 내역
    DispFootNote = ""
    
    Set dsFNote = New Recordset
    dsFNote.Open objMicSql.SqlGetFootNote(pWorkArea, pAccDt, pAccSeq), DBConn

    While Not dsFNote.EOF
        DispFootNote = DispFootNote & dsFNote.Fields("rsttxt").Value & vbCrLf
        dsFNote.MoveNext
    Wend

    Set dsFNote = Nothing

End Function

Public Sub DispStainTable(ByRef objTable As Object, ByVal pWorkArea As String, ByVal pAccDt As String, _
                          ByVal pAccSeq As String, ByVal pRstTp As String, _
                          Optional ByVal pMfyFg As Boolean = False)
    Dim dsTest      As Recordset
    Dim sqlTest     As String
    Dim sRType      As String
    Dim sRDetail    As String
    Dim sPos        As String
    Dim iTestCol    As Integer
    Dim iCnt        As Integer

    Set dsTest = New Recordset
    dsTest.Open objMicSql.SQLLoadStainRst(pWorkArea, pAccDt, pAccSeq, pRstTp, pMfyFg), DBConn

    If dsTest.EOF Then
        MsgBox "해당 검사내역이 없습니다. 검체군이 맞게 선택되었는지 확인하십시오.", vbInformation, "메세지"
        Set dsTest = Nothing
        Exit Sub
    End If

    With objTable
        objTable.ReDraw = False
    
        iCnt = 0
        While (Not dsTest.EOF)
    
            iCnt = iCnt + 1
            .MaxRows = iCnt
            .Row = iCnt
            
            sRType = "" & dsTest.Fields("rstdiv").Value
            sRDetail = "" & dsTest.Fields("detailfg").Value
    
            sPos = SetTable(objTable, iCnt, sRType, sRDetail)
    
            .RowHeight(iCnt) = 12
            .Col = enSTAIN.tcTESTNM: .Text = sPos & dsTest.Fields("testnm").Value
            .Col = enSTAIN.tcTESTCD:  .Text = "" & dsTest.Fields("testcd").Value
            .Col = enSTAIN.tcRSTDIV:  .Text = "" & sRType
            .Col = enSTAIN.tcDETAIL:  .Text = "" & sRDetail
            
            '-- 최근결과 색깔변경 By M.G.Choi 2005.03.30
            .Col = enSTAIN.tcLASTRST: .Text = "" & dsTest.Fields("rst2").Value
            '==========================
            .ForeColor = DCM_LightRed
            '==========================
            
            .Col = enSTAIN.tcRSTCD:   .Text = "" & dsTest.Fields("rstcd").Value
            .Col = enSTAIN.tcOLDRST:  .Text = "" & dsTest.Fields("rstcd").Value
            .Col = enSTAIN.tcRSTNM:   .ForeColor = &HDF6A3E: .Text = "" & dsTest.Fields("rst1").Value
            
            '** 전주예수병원 추가루틴 By M.G.Choi 2004.10.05 ========================================
            .Col = enSTAIN.tcOLDRST + 1: .Text = "" & dsTest.Fields("vfydt").Value & dsTest.Fields("vfytm").Value
            .Col = enSTAIN.tcOLDRST + 2: .Text = "" & dsTest.Fields("vfyid").Value
            '========================================================================================
            
            dsTest.MoveNext
        Wend
    
        .MaxRows = dsTest.RecordCount
    
        .ReDraw = True
    End With
    
    DoEvents

    Set dsTest = Nothing

End Sub

Private Function SetTable(ByRef objTable As Object, ByVal pN As Integer, _
                          ByVal pRType As String, ByVal pRDetail As String) As String

    With objTable
        If pRType = "*" Then
            .Col = enSTAIN.tcTESTNM: .Row = pN
            .BackColor = DCM_LightGray
    
            .Col = enSTAIN.tcRSTCD: .Row = pN
            .CellType = 5   'CellTypeStaticText
    
            .Col = enSTAIN.tcEXCPT: .Row = pN
            .CellType = 10  'CellTypeCheckBox
            .TypeCheckCenter = True
            SetTable = Space$(1)
        Else
            .Col = enSTAIN.tcRSTCD: .Row = pN
            .CellType = 1   'CellTypeEdit
            .TypeEditCharCase = 2   'TypeEditCharCaseSetUpper
    
            If pRDetail = "" Then
                .Col = enSTAIN.tcEXCPT: .Row = pN
                .CellType = 10  'CellTypeCheckBox
                .TypeCheckCenter = True
                SetTable = Space$(1)
            Else
                .Col = enSTAIN.tcEXCPT: .Row = pN
                .CellType = 5   'CellTypeStaticText
                SetTable = Space$(10)
            End If
        End If
    End With
    
End Function

'----------------------------------------------------------------------------------
'2. 미생물 Stain 결과수정 관련 메쏘드
'----------------------------------------------------------------------------------

Public Function ModifyStainResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                                  ByVal ssRst As Object, ByVal pWsCd As String, ByVal pEmpId As String, _
                                  ByVal pFNSeq As Long, ByVal pFNote As String, ByVal pRmkCd As String) As Boolean

    Dim sDate       As String
    Dim sTime       As String
    Dim sTestCd     As String
    Dim sCRstCd     As String
    Dim sPRstCd     As String
    Dim sTestDiv    As String
    Dim sqlUpdate   As String
    Dim sMfySeq     As Long
    
    '-- 전주예수병원 추가 변수=============
    Dim blnVerified     As Boolean
    Dim bINUPFlag       As Boolean
    Dim SqlTransfer()   As String
    Dim intTempSql      As Integer
    Dim clsTransfer     As New clsLISTransfer
    Dim strVfyDt        As String
    Dim strRstVal       As String
    Dim strOrdSQL       As String
    Dim strOrdInfo      As String
    Dim sPtId           As String
    Dim sOrdDt          As String
    Dim sOrdNo          As String
    Dim strREmpID       As String
    Dim strRVfyDt       As String
    Dim strIPAdr        As String
    Dim strSql          As String
    Dim strSEQ          As String
    Dim strKey          As String
    Dim strWrkDiv       As String
    Dim strOCSNote      As String
    Dim strFromRef      As String
    Dim strToRef        As String
    Dim strRemark       As String
    '======================================
    
    ' 시스템 일자/시간 설정
    ' 시스템 일자/시간 설정
    sDate = Format(GetSystemDate, CS_DateDbFormat)
    sTime = Format(GetSystemDate, CS_TimeDbFormat)
    
    '-- 전주예수병원IP Address =============
    strIPAdr = clsTransfer.LocalIP_Address
    '=======================================
            
    If ChkModify(ssRst) = False Then
        MsgBox "결과 입력이 정확하지 않습니다. 확인 후 처리하십시오"
        Exit Function
    End If

On Error GoTo DBExecError

    DBConn.BeginTrans

    Dim i As Integer
    
    '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.16=======================
    
    '=================================================================================
    
    For i = 1 To ssRst.MaxRows
        
        ssRst.Row = i
        ssRst.Col = enSTAIN.tcTESTCD:  sTestCd = ssRst.Text
        ssRst.Col = enSTAIN.tcRSTCD:   sCRstCd = Trim(ssRst.Text)
        ssRst.Col = enSTAIN.tcOLDRST:  sPRstCd = Trim(ssRst.Text)        '수정전 결과
        
        '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.10.05=======================
        ssRst.Col = enSTAIN.tcOLDRST + 1: strRVfyDt = Trim(ssRst.Text)   '결과확인일시
        ssRst.Col = enSTAIN.tcOLDRST + 2: strREmpID = Trim(ssRst.Text)   '결과확인ID
        '=================================================================================
        
        '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.16
        'If PROJECT_HOSCD = "02" Then
        blnVerified = IsAllVerified(pWorkArea, pAccDt, pAccSeq, sTestCd, strOrdInfo)
        sPtId = medGetP(strOrdInfo, 1, COL_DIV)
        sOrdDt = medGetP(strOrdInfo, 2, COL_DIV)
        sOrdNo = medGetP(strOrdInfo, 3, COL_DIV)
        '========================================================================================
            
        '-- 추가 By M.G.Choi 2002.09.26
        ' - 대표항목코드 처방Body에 Update ======================================================
        ssRst.Col = enSTAIN.tcRSTDIV: sTestDiv = ssRst.Text
        If sTestDiv = "*" Then
            '처방내역 Update
            sqlUpdate = objMicSql.SqlUpdateOrder(pWorkArea, pAccDt, pAccSeq, sTestCd, _
                                                 enStsCd.StsCd_LIS_Modify, sDate, sTime, pEmpId)
            DBConn.Execute sqlUpdate
            
            '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.16=======================
            ' - 상세모항목 Insert
'            strVfyDt = sDate & sTime
'
'            strSql = clsTransfer.SqlOCSRESULT_UPDATE(sPtId, sOrdDt, sOrdNo, _
'                    sTestCd, pWorkArea, "", "", "", "", "", _
'                    "", "", "", "", "", "", "", pEmpId, strIPAdr, strVfyDt)
'
'            If strSql <> "" Then DBConn.Execute strSql
            '=================================================================================
                    
        End If
        '========================================================================================
        
        If sCRstCd <> sPRstCd Then
            
            If Not SetHistory(pWorkArea, pAccDt, pAccSeq, sTestCd, sMfySeq) Then GoTo DBExecError
            
            '결과내역 Update
            sqlUpdate = objMicSql.SQLVerifyStainResult(pWorkArea, pAccDt, pAccSeq, sTestCd, _
                                sCRstCd, enStsCd.StsCd_LIS_Modify, sDate, sTime, pEmpId, "", sMfySeq + 1)
            DBConn.Execute sqlUpdate
                        
            '처방내역 Update
            sqlUpdate = objMicSql.SqlUpdateOrder(pWorkArea, pAccDt, pAccSeq, sTestCd, _
                                                 enStsCd.StsCd_LIS_Modify, sDate, sTime, pEmpId)
            DBConn.Execute sqlUpdate
        End If
        
        '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.16
        '-- 결과등록자ID, 결과일자 가져오기
        'bINUPFlag = clsTransfer.SqlOCSRESULT_FLAG(sPtId, sOrdDt, sOrdNo, sTestCd)
        
        If sTestDiv = "*" Then
            GoTo Skip
        End If
        
        If pFNote <> "" Then
            strOCSNote = "[FootNote]" & Chr(13) & pFNote & Chr(13)
        Else
            strOCSNote = ""
        End If
        
        If pRmkCd <> "" Then
            strRemark = clsTransfer.SqlGetRemark(pRmkCd)
            If strRemark <> "" Then
                strOCSNote = strOCSNote & "◈ Remark" & Chr(13) & strRemark
            End If
        End If
                      
        strVfyDt = sDate & sTime
        
        '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
        If strKey <> sPtId & sOrdDt & sOrdNo Then
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(sPtId, sOrdDt, sOrdNo)
            If strWrkDiv = "3" Then
                If clsTransfer.SqlOCSSU2EXAMT_Find(sPtId, sOrdDt, sOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE(sPtId, sOrdDt, sOrdNo, _
                                            "N", "", strOCSNote, "", strRVfyDt, strREmpID, strRVfyDt, strREmpID, pEmpId, strIPAdr, strVfyDt)
                Else
                    GoTo DBExecError
                End If
            ElseIf strWrkDiv = "4" Then
                    If clsTransfer.SqlOCSSG2EXAMT_Find(sPtId, sOrdDt, sOrdNo) = True Then
                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE(sPtId, sOrdDt, sOrdNo, _
                                                "N", "", strOCSNote, "", strRVfyDt, strREmpID, strRVfyDt, strREmpID, pEmpId, strIPAdr, strVfyDt)
                    Else
                        GoTo DBExecError
                    End If
            Else
                If clsTransfer.SqlOCSMDEXMORT_Find(sPtId, sOrdDt, sOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE(sPtId, sOrdDt, sOrdNo, _
                                            "N", strRVfyDt, strREmpID, pEmpId, strIPAdr)
                Else
                    GoTo DBExecError
                End If
            End If
            
            strKey = sPtId & sOrdDt & sOrdNo
            
        End If
        
        '** 종건/일건 이외의 검사만 등록.
        If strWrkDiv <> "" Then
            GoTo OCSSKIP
        End If
        
        '-- 2. 건진외결과 INSERT
        strSEQ = clsTransfer.SqlOCSRESULT_HISTORY(sPtId, sOrdDt, sOrdNo, sTestCd)
        
        If strSEQ = "" Then
            strSEQ = "1"
        Else
            strSEQ = CInt(strSEQ) + 1
        End If
                
        '** 검사항목별 결과값 체크
        Call clsTransfer.ResultValue_Find(sTestCd, sCRstCd)
        
        '-- 결과전송 시 필요정보
        Call clsTransfer.SqlOCS_Info_Mic(pWorkArea, pAccDt, pAccSeq)
        
        If clsTransfer.SqlOCSRESULT_FLAG(sPtId, sOrdDt, sOrdNo, sTestCd) = True Then
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_UPDATE(sPtId, sOrdDt, sOrdNo, _
                                    sTestCd, pWorkArea, clsTransfer.ItemVal, "", "", "", "", _
                                    strOCSNote, "", "", "", "", "", "", pEmpId, strIPAdr, strVfyDt)
        Else
            '-- 임시 사용 By M.G.Choi 2007.06.28
            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(sPtId, sOrdDt, sOrdNo, _
                                                    sTestCd, pWorkArea, clsTransfer.ItemVal, "", "", "", "", _
                                                    strOCSNote, clsTransfer.RcvDt & clsTransfer.RCVTM, strVfyDt, pEmpId, strVfyDt, pEmpId, "", clsTransfer.RsltType, _
                                                    "", "", "")
'            GoTo DBExecError
        End If
                    
        intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
        
        '-- OCS 결과History테이블 업데이트
        SqlTransfer(intTempSql) = clsTransfer.SqlOCSRSLTHT_INSERT(sPtId, sOrdDt, sOrdNo, _
                    sTestCd, strSEQ, clsTransfer.SlipCd, clsTransfer.RstVal, clsTransfer.RstText, _
                    clsTransfer.RefValTo, clsTransfer.RefValFrom, clsTransfer.RstUnit, _
                    clsTransfer.FootNote, "", clsTransfer.VfyDt & clsTransfer.VfyTm, _
                    clsTransfer.VfyId, clsTransfer.ReadDt & clsTransfer.ReadTm, clsTransfer.VfyId, _
                    clsTransfer.VfyDt & clsTransfer.VfyTm, clsTransfer.RsltType, _
                    clsTransfer.MfyID, clsTransfer.IPAdr, clsTransfer.MfyDt & clsTransfer.MfyTm)
                                    
OCSSKIP:

        For intTempSql = LBound(SqlTransfer) To UBound(SqlTransfer)
            intTempSql = intTempSql
           
            If Trim(SqlTransfer(intTempSql)) <> "" Then DBConn.Execute (SqlTransfer(intTempSql))
            
        Next intTempSql
        
        Erase SqlTransfer
        intTempSql = 0: ReDim SqlTransfer(0)
        '=======================================================================================
        
Skip:

    Next i

    '## 5.0.9: 이상대(2005-01-31)
    '   - 201의 Status가 "5"이상일때만 201의 Status 업데이트
    '접수내역 Status반영
    Dim Rs  As Recordset
    Dim SQL As String
    Dim strStsCd As String
    
    SQL = objMicSql.SqlGetTestCount(pWorkArea, pAccDt, pAccSeq)
    Set Rs = New Recordset
    Rs.Open SQL, DBConn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not (Rs.BOF Or Rs.EOF) Then
        strStsCd = Rs.Fields("stscd").Value
    End If
    Rs.Close
    Set Rs = Nothing
    
    If strStsCd = enStsCd.StsCd_LIS_FinRst Or strStsCd = enStsCd.StsCd_LIS_Modify Then
        sqlUpdate = objMicSql.SQLUpdateStatus(pWorkArea, pAccDt, pAccSeq, enStsCd.StsCd_LIS_Modify, _
                                                0, sDate, sTime, pEmpId)
        DBConn.Execute sqlUpdate
    End If
    
    ' *** Foot Note 저장 *** 검체 Remark 저장
    If Not AppendFootnote(pWorkArea, pAccDt, pAccSeq, enStsCd.StsCd_LIS_Modify, _
                                        pFNSeq, pFNote, pRmkCd, pEmpId) Then GoTo DBExecError
                                        
    '결과보고대기내역 작성
    '===========================================================================================
    '결과보고대기내역 생성(2002/09/05)
    Dim strTmp      As String
    Dim tmpBussdiv  As String
    Dim tmpDept     As String
    Dim tmpDoct     As String
    Dim tmpPtid     As String
    
    strTmp = ""
    tmpBussdiv = "": tmpDept = "": tmpDoct = ""

    strTmp = Get_OrderInFo(pWorkArea, pAccDt, pAccSeq)
    tmpBussdiv = medGetP(strTmp, 1, COL_DIV)
    tmpDoct = medGetP(strTmp, 4, COL_DIV)
    tmpPtid = medGetP(strTmp, 5, COL_DIV)
    Select Case tmpBussdiv
        Case "1": tmpDept = medGetP(strTmp, 3, COL_DIV)
        Case "2": tmpDept = medGetP(strTmp, 2, COL_DIV)
        Case Else
            tmpDept = medGetP(strTmp, 2, COL_DIV)
            If tmpDept = "" Then tmpDept = medGetP(strTmp, 3, COL_DIV)
    End Select
    If Not SubmitVerifyList(tmpDept, sDate, sTime, tmpPtid, enStsCd.StsCd_LIS_Modify, ObjMyUser.EmpId, tmpDoct, tmpBussdiv) Then GoTo DBExecError
    '============================================================================================================================
    
    DBConn.CommitTrans
    Exit Function

DBExecError:
    Set Rs = Nothing
    DBConn.RollbackTrans
    MsgBox Err.Description, vbCritical, "오류"
End Function
    
Private Function ChkModify(ByVal ssRst As Object) As Boolean
    
    Dim sTestCd     As String
    Dim sReqFg      As String
    Dim sDetSeq     As String
    Dim sResult     As String
    Dim sReqFg2     As String
    Dim sDetSeq2    As String
    Dim sResult2    As String
    Dim R1          As Integer
    Dim R2          As Integer
    Dim R3          As Integer
    Dim chkDR       As Integer
    Dim chkDA1      As Integer
    Dim chkDA2      As Integer
    ChkModify = True
    
    R1 = 1
    Do Until R1 >= ssRst.MaxRows
        
        ssRst.Row = R1
        ssRst.Col = enSTAIN.tcTESTCD: sTestCd = ssRst.Text
        ssRst.Col = enSTAIN.tcRSTDIV: sReqFg = ssRst.Text
        ssRst.Col = enSTAIN.tcDETAIL: sDetSeq = ssRst.Text
        ssRst.Col = enSTAIN.tcRSTCD:  sResult = Trim(ssRst.Text)
    
        If sReqFg = "*" Then
            ' 당연히 Detail 항목은 대표항목 아래에 연속해서 존재하는 것으로 가정
            ' Display 할때 그렇게 처리함. 만약 아니라면 체크루틴 수정 필요
            chkDR = 0: chkDA1 = 0: chkDA2 = 0: R2 = R1 + 1
            For R2 = R1 + 1 To ssRst.MaxRows
                ssRst.Row = R2
                ssRst.Col = enSTAIN.tcRSTDIV: sReqFg2 = ssRst.Text
                ssRst.Col = enSTAIN.tcDETAIL: sDetSeq2 = ssRst.Text
                ssRst.Col = enSTAIN.tcRSTCD:  sResult2 = Trim(ssRst.Text)
                If sDetSeq2 = sDetSeq Then
                    If sReqFg2 = LIS_RST_REQUIRED And sResult2 = "" Then chkDR = chkDR + 1
                    If sReqFg2 = LIS_RST_ALTERNATIVE And sResult2 = "" Then chkDA1 = chkDA1 + 1
                    If sReqFg2 = LIS_RST_ALTERNATIVE And sResult2 <> "" Then chkDA2 = chkDA2 + 1
                Else
                    Exit For
                End If
            Next R2
            ' 대표항목,세부항목 일괄 처리
            If chkDR > 0 Or (chkDA1 > 0 And chkDA2 = 0) Then ChkModify = False
            R1 = R2
        Else
            ' 혹시 대표항목 없이 세부항목이 먼저 발견되면..
            If sDetSeq <> "" Then
                MsgBox "검사항목 Display 순서에 오류가 있습니다. 전산실로 연락 바랍니다(☎" & ObjSysInfo.HelpLine & ")", _
                        vbCritical, "기초자료오류"
                ChkModify = False
                Exit Function
            End If
            If sResult = "" And sDetSeq = "" Then ChkModify = False
            R1 = R1 + 1
        End If
    Loop

End Function


Private Function SetHistory(ByVal pWorkArea As String, ByVal pAccDt As String, _
                            ByVal pAccSeq As String, ByVal pTestCd As String, ByRef pMfySeq As Long) As Boolean
    Dim dsCurRst    As Recordset
    Dim sCurRst     As String
    Dim sqlDel407   As String
    Dim iSeq        As Long
    Dim sHead       As String
    Dim sData       As String
    Dim sqlIns407   As String
    Dim sMfySeq     As String
    Dim sRstCd      As String
    Dim sSenFg      As String
    Dim sRstDiv     As String
    Dim sDetailFg   As String
    Dim sVfyDt      As String
    Dim sVfyTm      As String
    Dim sVfyId      As String
    
    ' 현재 변경 대상 Data Read
    
    On Error GoTo Err_Trap
    
    Set dsCurRst = New Recordset
    dsCurRst.Open objMicSql.SQLGetCurrentRst(pWorkArea, pAccDt, pAccSeq, "'" & pTestCd & "'"), DBConn
    
    If Not dsCurRst.EOF Then
        sMfySeq = "" & dsCurRst.Fields("mfyseq").Value
        sRstCd = "" & dsCurRst.Fields("RstCd").Value
        sSenFg = "" & dsCurRst.Fields("SenFg").Value
        sRstDiv = "" & dsCurRst.Fields("RstDiv").Value
        sDetailFg = "" & dsCurRst.Fields("DetailFg").Value
        sVfyDt = "" & dsCurRst.Fields("VfyDt").Value
        sVfyTm = "" & dsCurRst.Fields("VfyTm").Value
        sVfyId = "" & dsCurRst.Fields("VfyId").Value
    End If
    
    Set dsCurRst = Nothing
    
    sqlIns407 = objMicSql.SQLInsertOldData(pWorkArea, pAccDt, pAccSeq, pTestCd, sMfySeq, sRstCd, _
                                           sSenFg, sRstDiv, sDetailFg, sVfyDt, sVfyTm, sVfyId)
    DBConn.Execute sqlIns407

    pMfySeq = Val(sMfySeq)
    
    SetHistory = True
    Exit Function

Err_Trap:
    SetHistory = False
    
End Function



Private Function AppendFootnote(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, ByVal pStatus As String, _
                                ByVal fFNSeq As Long, ByVal pFNote As String, ByVal pRmkCd As String, ByVal pEmpId As String) As Boolean
    Dim dsFN        As Recordset
    Dim sqlFNDel    As String
    Dim sqlFNIns    As String
    Dim sRemarkCd   As String
    Dim sqlUpdAcc   As String
    Dim sqlFN       As String
    Dim iFNCol      As Integer
    Dim sFSeq       As Integer
    Dim sMfyNo      As Integer
    
    'Foot Note 결과 저장

    On Error GoTo Err_Trap
    
    sFSeq = fFNSeq
    
    If Trim(pFNote) <> "" Then
        
        sFSeq = 2
        
        Set dsFN = New Recordset
        dsFN.Open objMicSql.SQLGetMaxNoOfFootnote(pWorkArea, pAccDt, pAccSeq), DBConn
        
        If dsFN.EOF Then
           sMfyNo = 1
        Else
           sMfyNo = Val(dsFN.Fields("mfyno").Value & "") + 1
        End If
        Set dsFN = Nothing

        sqlFNIns = objMicSql.SQLSaveFootNote(pWorkArea, pAccDt, pAccSeq, sMfyNo, pEmpId, pFNote, 1)
        DBConn.Execute sqlFNIns
    
    End If

    If sRemarkCd = LIS_Nothing Then sRemarkCd = ""

    ' 접수 내역에 Footnote Seq 및 검체 Remark Code 등록
    If pStatus >= enStsCd.StsCd_LIS_MidRst Then
        sqlUpdAcc = objMicSql.SQLSaveRemark(pWorkArea, pAccDt, pAccSeq, sFSeq, pRmkCd)
        DBConn.Execute sqlUpdAcc
    End If
    AppendFootnote = True
    Exit Function
    
Err_Trap:
    AppendFootnote = False

End Function



Public Sub LoadMicSpecies(ByRef cboSpecies As Object)
    Dim objSQL  As clsLISSqlStatistic
    Dim dsSC    As Recordset
    Dim sqlSC   As String
    Dim iSCCol  As Integer
    Dim strTemp As String
    
'    strTemp = " SELECT cdval1 mkind FROM " & T_LAB032 & _
'                 " WHERE " & DBW("cdindex=", LC3_Species) & " " & _
'                 " ORDER BY cdval1 asc"

    Set dsSC = New Recordset
    Set objSQL = New clsLISSqlStatistic
    dsSC.Open objSQL.GetSpecies, DBConn
    Set objSQL = Nothing

'    dsSC.Open strTemp, DBConn
    cboSpecies.Clear

    cboSpecies.AddItem CS_AllCaption
    
    While Not dsSC.EOF
        cboSpecies.AddItem "" & dsSC.Fields("mkind").Value
        dsSC.MoveNext
    Wend
    
    Set dsSC = Nothing

End Sub



Public Sub LoadAnti(ByVal pSCd As String, ByRef lstAnti As Object)
    
    Dim dsAnti As Recordset
    Dim objSQL As clsLISSqlStatistic

    Set objSQL = New clsLISSqlStatistic
    Set dsAnti = New Recordset
    dsAnti.Open objSQL.GetAntiBiotic(pSCd), DBConn

    lstAnti.Clear
    While Not dsAnti.EOF
        lstAnti.AddItem "" & dsAnti.Fields("anti").Value
        dsAnti.MoveNext
    Wend
    
    Set dsAnti = Nothing
    Set objSQL = Nothing

End Sub


'결과보고 대기내역
Public Function Get_Bussdiv(ByVal Workarea As String, ByVal Accdt As String, ByVal Accseq As String) As String
    Dim sRS     As Recordset
    Dim sSql    As String
    
    sSql = " select bussdiv from " & T_LAB101 & " a," & T_LAB102 & " b" & _
           " where " & DBW("b.workarea=", Workarea) & " and " & DBW("b.accdt=", Accdt) & " and " & DBW("b.accseq=", Accseq) & _
           " and a.ptid=b.ptid and a.orddt=b.orddt and a.ordno=b.ordno"
    
    Set sRS = New Recordset
    sRS.Open sSql, DBConn
    
    
    If Not sRS.EOF Then
        Get_Bussdiv = Trim(sRS.Fields("bussdiv").Value & "")
    End If
    Set sRS = Nothing
End Function
'결과보고 대기내역
Public Function Get_OrderInFo(ByVal Workarea As String, ByVal Accdt As String, ByVal Accseq As String) As String
    Dim sRS     As Recordset
    Dim sSql    As String
    
    Dim strTmp  As String
    sSql = " select a.bussdiv,a.wardid,a.deptcd,a.majdoct,a.ptid from " & T_LAB101 & " a," & T_LAB102 & " b" & _
           " where " & DBW("b.workarea=", Workarea) & " and " & DBW("b.accdt=", Accdt) & " and " & DBW("b.accseq=", Accseq) & _
           " and a.ptid=b.ptid and a.orddt=b.orddt and a.ordno=b.ordno"
    
    Set sRS = New Recordset
    sRS.Open sSql, DBConn
    
    
    If Not sRS.EOF Then
        Get_OrderInFo = Trim(sRS.Fields("bussdiv").Value & "") & COL_DIV & _
                        Trim(sRS.Fields("wardid").Value & "") & COL_DIV & _
                        Trim(sRS.Fields("deptcd").Value & "") & COL_DIV & _
                        Trim(sRS.Fields("majdoct").Value & "") & COL_DIV & _
                        Trim(sRS.Fields("ptid").Value & "")
    Else
        sSql = " select * from " & T_LAB201 & _
               " where " & DBW("workarea=", Workarea) & " and " & DBW("accdt=", Accdt) & " and " & DBW("accseq=", Accseq)
        Set sRS = Nothing
        Set sRS = New Recordset
        sRS.Open sSql, DBConn
        
        If Not sRS.EOF Then
            
            
            If Trim(sRS.Fields("wardid").Value & "") <> "" Then
                strTmp = enBussDiv.BussDiv_InPatient
            Else
                strTmp = enBussDiv.BussDiv_OutPatient
            End If
            Get_OrderInFo = strTmp & COL_DIV & _
                            Trim(sRS.Fields("wardid").Value & "") & COL_DIV & _
                            Trim(sRS.Fields("deptcd").Value & "") & COL_DIV & _
                            Trim(sRS.Fields("majdoct").Value & "") & COL_DIV & _
                            Trim(sRS.Fields("ptid").Value & "")
                
        End If
    End If
    
    Set sRS = Nothing
End Function

Public Function SubmitVerifyList(ByVal pDeptCd As String, ByVal pDate As String, ByVal pTime As String, _
                                 ByVal pPtId As String, ByVal pStatus As String, ByVal pEmpId As String, _
                                 ByVal pMajDoct As String, ByVal pBussDiv As String) As Boolean
    Dim dsSubmit    As Recordset
    Dim tmpSql      As clsLISSqlETest
    Dim sSubmit     As String
    Dim sMfyFg      As String
    Dim sExpire     As String
    
On Error GoTo Err_Trap

    SubmitVerifyList = True
    
    If pStatus < enStsCd.StsCd_LIS_Modify Then sMfyFg = "0"
    If pStatus >= enStsCd.StsCd_LIS_Modify Then sMfyFg = "1"
    
    Set tmpSql = New clsLISSqlETest
    Set dsSubmit = New Recordset
    dsSubmit.Open tmpSql.SqlReadVerifyList(pDeptCd, pDate, pPtId, sMfyFg), DBConn
     
    If dsSubmit.RecordCount >= 1 Then
    
        sExpire = "" & dsSubmit.Fields("donefg").Value
        If sExpire = "1" Then
        ' 갱신 루틴
           DBConn.Execute tmpSql.SqlUpdateVerifyList(pDeptCd, pDate, pTime, pPtId, sMfyFg, pEmpId)
        Else
        ' Skip 루틴
            Set dsSubmit = Nothing
            Set tmpSql = Nothing
            Exit Function
        End If
      
    Else
        ' 삽입 루틴
        DBConn.Execute tmpSql.SqlInsertVerifyList(pDeptCd, pDate, pTime, pPtId, sMfyFg, pEmpId, pMajDoct, pBussDiv)
    End If
    Set dsSubmit = Nothing
    Set tmpSql = Nothing
    Exit Function

Err_Trap:
    Set tmpSql = Nothing
    SubmitVerifyList = False

End Function


'-------------------------------
' 미생물 바코드 재발행
'-------------------------------
Public Sub LoadMedia(ByVal pCode As String, ByRef plblMediaList As Object)
    Dim dsNG    As Recordset
    Dim i       As Integer
    Dim j       As Integer
    

    ' Nogrowth result 로드
    Set dsNG = New Recordset
    dsNG.Open objMicSql.SQLGetMedias(pCode), DBConn
    
    plblMediaList.Caption = ""
    
    If dsNG.RecordCount > 0 Then
        dsNG.MoveFirst
        Do Until dsNG.EOF
            If plblMediaList.Caption = "" Then
                plblMediaList.Caption = dsNG.Fields("media").Value & ""
            Else
                plblMediaList.Caption = plblMediaList.Caption & ", " & dsNG.Fields("media").Value & ""
            End If
            dsNG.MoveNext
        Loop
    End If
    
    Set dsNG = Nothing

End Sub

Public Function DisPlayReBarCodeList(ByRef ssTable As Object, ByVal pWsCode As String, _
                                 ByVal pWsUnit As String, ByVal pRTs As String) As String
    Dim objSQL      As clsLISSqlStatement
    Dim tmpRs       As Recordset
    Dim dsNWS       As Recordset
    Dim sLabNo      As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim strLabNum   As String

    Dim SqlStmt     As String
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String
    Dim SvSpcNm     As String
    Dim SvOrdDoct   As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim tmpStatFg   As String
    Dim tmpTestFg   As String
    Dim tmpTestCd   As String
    Dim tmpTestNm   As String
    Dim tmpPtid     As String
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim jj          As Integer
    
    
    MouseRunning
    DoEvents
    Set objSQL = New clsLISSqlStatement
    Set dsNWS = New Recordset
    dsNWS.Open objMicSql.SQLBarcodeReprint(pWsCode, pWsUnit, pRTs), DBConn
    
    iCnt = 0
    i = 1
    strLabNum = ""
    
    With ssTable
        .ReDraw = False
        .MaxRows = 0
            
        If Not dsNWS.EOF Then
            
            If dsNWS.RecordCount < 22 Then
                .MaxRows = 22
            Else
                .MaxRows = dsNWS.RecordCount   '데이타 건수
            End If
            .RowHeight(-1) = 13
        
            While Not dsNWS.EOF
                sWorkArea = "" & dsNWS.Fields("workarea").Value
                sAccDt = "" & dsNWS.Fields("accdt").Value
                sAccSeq = "" & dsNWS.Fields("accseq").Value
                tmpTestNm = ""
                
                SqlStmt = objSQL.SqlBarReprint(2, sWorkArea, sAccDt, sAccSeq)
                Set tmpRs = Nothing
                Set tmpRs = New Recordset
                tmpRs.Open SqlStmt, DBConn
      
                If Not tmpRs.EOF Then
                
                    Do Until tmpRs.EOF
    '                   DoEvents
                       .Row = i
                        If tmpTestNm = "" Then
                            .Col = 1: .Value = 0
                            .Col = 9: .Value = Format("" & tmpRs.Fields("OrdDt").Value, CS_DateMask)    '처방일
                            .Col = 5: .Value = Trim("" & tmpRs.Fields("OrdNo").Value)                   '처방번호
                            .Col = 7: .Value = Trim("" & tmpRs.Fields("SpcNm").Value)                   '검체
                        
                            .Col = 6: .Value = Trim("" & tmpRs.Fields("TestNm").Value)                      '처방명
                                      .ForeColor = DCM_LightBlue                                         '약간 파란색
                            .Col = 8: .Value = Choose(Val("" & tmpRs.Fields("StatFg").Value) + 1, "", "Y")  '응급여부
                                         .ForeColor = DCM_Red                                               '빨간색
                                         
                            '## 5.0.12: 이상대(2005-05-30)
                            '   - 접수번호, 환자ID의 순서를 변경
                            .Col = 4: .Value = "" & dsNWS.Fields("ptid").Value & ""
                            .Col = 3: .Value = "" & dsNWS.Fields("ptnm").Value & ""
                            .Col = 2: .Value = Trim("" & tmpRs.Fields("LabNo").Value)       'LabNo
                            .Col = 10: .Value = Trim("" & tmpRs.Fields("OrdDt").Value)       '처방일
                            .Col = 11: .Value = Trim("" & tmpRs.Fields("OrdNo").Value)       '처방번호
                            .Col = 12: .Value = Trim("" & tmpRs.Fields("OrdSeq").Value)     '처방Seq
                            .Col = 13: .Value = Trim("" & tmpRs.Fields("OrdCd").Value)      '검사코드
                            .Col = 14: .Value = Trim("" & tmpRs.Fields("SpcNm").Value)      '검체명
                            .Col = 15: .Value = Trim("" & tmpRs.Fields("WorkArea").Value)   'WorkArea
                            .Col = 16: .Value = Trim("" & tmpRs.Fields("AccSeq").Value)     'AccSeq
                            .Col = 17: .Value = Trim("" & tmpRs.Fields("StoreCd").Value)    '보관코드
                            .Col = 18: .Value = Trim("" & tmpRs.Fields("AccDt").Value)      'AccDt  채혈일
                            .Col = 19: .Value = Trim("" & tmpRs.Fields("StatFg").Value)     '응급여부
                            .Col = 20: .Value = Trim("" & tmpRs.Fields("AbbrNm5").Value)    '약어명
                            .Col = 21: .Value = Trim("" & tmpRs.Fields("SpcYy").Value) & _
                                                Format(Val(tmpRs.Fields("SpcNo").Value), CS_BarFormat)  '검체번호
                            .Col = 22: .Value = Trim("" & tmpRs.Fields("BuildNm").Value)    '건물명
                            .Col = 23: .Value = Trim("" & tmpRs.Fields("HosilId").Value)    '호실코드
                            .Col = 24: .Value = Trim("" & tmpRs.Fields("DeptCd").Value)     '진료과코드
                            .Col = 25: .Value = Trim("" & tmpRs.Fields("StsCd").Value)      'status
                            .Col = 26: .Value = Mid(Trim("" & tmpRs.Fields("ReqTm").Value), 1, 2) & ":" & _
                                                Mid(Trim("" & tmpRs.Fields("ReqTm").Value), 3, 2)       '희망채혈일시
                            .Col = 29: .Value = Trim("" & tmpRs.Fields("WardId").Value)     '병동코드
                            .Col = 30: .Value = "1"
                            tmpTestNm = .Value
                       Else
                            .Col = 6: .Value = .Value & "," & Trim("" & tmpRs.Fields("TestNm").Value)       '처방명
                            .Col = 20: .Value = .Value & "," & Trim("" & tmpRs.Fields("AbbrNm5").Value)     '약어명
                            tmpTestNm = .Value
                       End If
                       tmpRs.MoveNext
                   Loop
                    i = i + 1
                
                End If
                
                dsNWS.MoveNext
            Wend
            .Row = 1: .row2 = 25
            .Col = 2: .col2 = 27
            .blockmode = True
            .Lock = True
            .Protect = True
            .blockmode = False
        End If
        
        .RowHeight(-1) = 11
        .ReDraw = True
    End With
    
    Set dsNWS = Nothing: Set tmpRs = Nothing: Set objSQL = Nothing
    MouseDefault
End Function


Public Function DisPlayReBarCodeList1(ByRef ssTable As Object, ByVal pStart As String, _
                                 ByVal pEnd As String, ByVal pWork As String, ByVal pTestCd As String) As String
    Dim objSQL      As clsLISSqlStatement
    Dim tmpRs       As Recordset
    Dim dsNWS       As Recordset
    Dim sLabNo      As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim strLabNum   As String

    Dim SqlStmt     As String
    Dim SvOrdDt     As String
    Dim SvOrdNo     As String
    Dim SvSpcNm     As String
    Dim SvOrdDoct   As String
    Dim tmpDate     As String
    Dim tmpTime     As String
    Dim tmpStatFg   As String
    Dim tmpTestFg   As String
    Dim tmpTestCd   As String
    Dim tmpTestNm   As String
    Dim tmpPtid     As String
    Dim i           As Integer
    Dim iCnt        As Integer
    Dim jj          As Integer
    
    
    MouseRunning
    DoEvents
    Set objSQL = New clsLISSqlStatement
    Set dsNWS = New Recordset
    dsNWS.Open objMicSql.SQLBarcodeReprint1(pStart, pEnd, pWork, pTestCd), DBConn
    
    iCnt = 0
    i = 1
    strLabNum = ""
    
    With ssTable
        .ReDraw = False
        .MaxRows = 0
            
        If Not dsNWS.EOF Then
            
            If dsNWS.RecordCount < 22 Then
                .MaxRows = 22
            Else
                .MaxRows = dsNWS.RecordCount   '데이타 건수
            End If
            .RowHeight(-1) = 13
        
            While Not dsNWS.EOF
                sWorkArea = "" & dsNWS.Fields("workarea").Value
                sAccDt = "" & dsNWS.Fields("accdt").Value
                sAccSeq = "" & dsNWS.Fields("accseq").Value
                tmpTestNm = ""
                
                SqlStmt = objSQL.SqlBarReprint(2, sWorkArea, sAccDt, sAccSeq)
                Set tmpRs = Nothing
                Set tmpRs = New Recordset
                tmpRs.Open SqlStmt, DBConn
      
                If Not tmpRs.EOF Then
                
                    Do Until tmpRs.EOF
    '                   DoEvents
                       .Row = i
                        If tmpTestNm = "" Then
                            .Col = 1: .Value = 0
                            .Col = 9: .Value = Format("" & tmpRs.Fields("OrdDt").Value, CS_DateMask)    '처방일
                            .Col = 5: .Value = Trim("" & tmpRs.Fields("OrdNo").Value)                   '처방번호
                            .Col = 7: .Value = Trim("" & tmpRs.Fields("SpcNm").Value)                   '검체
                        
                            .Col = 6: .Value = Trim("" & tmpRs.Fields("TestNm").Value)                      '처방명
                                      .ForeColor = DCM_LightBlue                                         '약간 파란색
                            .Col = 8: .Value = Choose(Val("" & tmpRs.Fields("StatFg").Value) + 1, "", "Y")  '응급여부
                                         .ForeColor = DCM_Red                                               '빨간색
                                         
                            '## 5.0.12: 이상대(2005-05-30)
                            '   - 접수번호, 환자ID의 순서를 변경
                            .Col = 4: .Value = "" & dsNWS.Fields("ptid").Value & ""
                            .Col = 3: .Value = "" & dsNWS.Fields("ptnm").Value & ""
                            .Col = 2: .Value = Trim("" & tmpRs.Fields("LabNo").Value)       'LabNo
                            .Col = 10: .Value = Trim("" & tmpRs.Fields("OrdDt").Value)       '처방일
                            .Col = 11: .Value = Trim("" & tmpRs.Fields("OrdNo").Value)       '처방번호
                            .Col = 12: .Value = Trim("" & tmpRs.Fields("OrdSeq").Value)     '처방Seq
                            .Col = 13: .Value = Trim("" & tmpRs.Fields("OrdCd").Value)      '검사코드
                            .Col = 14: .Value = Trim("" & tmpRs.Fields("SpcNm").Value)      '검체명
                            .Col = 15: .Value = Trim("" & tmpRs.Fields("WorkArea").Value)   'WorkArea
                            .Col = 16: .Value = Trim("" & tmpRs.Fields("AccSeq").Value)     'AccSeq
                            .Col = 17: .Value = Trim("" & tmpRs.Fields("StoreCd").Value)    '보관코드
                            .Col = 18: .Value = Trim("" & tmpRs.Fields("AccDt").Value)      'AccDt  채혈일
                            .Col = 19: .Value = Trim("" & tmpRs.Fields("StatFg").Value)     '응급여부
                            .Col = 20: .Value = Trim("" & tmpRs.Fields("AbbrNm5").Value)    '약어명
                            .Col = 21: .Value = Trim("" & tmpRs.Fields("SpcYy").Value) & _
                                                Format(Val(tmpRs.Fields("SpcNo").Value), CS_BarFormat)  '검체번호
                            .Col = 22: .Value = Trim("" & tmpRs.Fields("BuildNm").Value)    '건물명
                            .Col = 23: .Value = Trim("" & tmpRs.Fields("HosilId").Value)    '호실코드
                            .Col = 24: .Value = Trim("" & tmpRs.Fields("DeptCd").Value)     '진료과코드
                            .Col = 25: .Value = Trim("" & tmpRs.Fields("StsCd").Value)      'status
                            .Col = 26: .Value = Mid(Trim("" & tmpRs.Fields("ReqTm").Value), 1, 2) & ":" & _
                                                Mid(Trim("" & tmpRs.Fields("ReqTm").Value), 3, 2)       '희망채혈일시
                            .Col = 29: .Value = Trim("" & tmpRs.Fields("WardId").Value)     '병동코드
                            .Col = 30: .Value = "2"
                            tmpTestNm = .Value
                       Else
                            .Col = 6: .Value = .Value & "," & Trim("" & tmpRs.Fields("TestNm").Value)       '처방명
                            .Col = 20: .Value = .Value & "," & Trim("" & tmpRs.Fields("AbbrNm5").Value)     '약어명
                            tmpTestNm = .Value
                       End If
                       tmpRs.MoveNext
                   Loop
                    i = i + 1
                
                End If
                
                dsNWS.MoveNext
            Wend
            .Row = 1: .row2 = 25
            .Col = 2: .col2 = 27
            .blockmode = True
            .Lock = True
            .Protect = True
            .blockmode = False
        End If
        
        .RowHeight(-1) = 11
        .ReDraw = True
    End With
    
    Set dsNWS = Nothing: Set tmpRs = Nothing: Set objSQL = Nothing
    MouseDefault
End Function

Public Function IsAllVerified(ByVal pWorkArea As String, ByVal pAccDt As String, _
                              ByVal pAccSeq As String, ByVal pTestCd As String, _
                              ByRef pOrdInfo As String, _
                              Optional ByVal pMultiFg As Boolean = False) As Boolean
    Dim objRs   As Recordset
    Dim strSql  As String
    Dim sPtId   As String
    Dim sOrdDt  As String
    Dim sOrdNo  As String
    Dim sOrdSeq As String
    
    IsAllVerified = True
    pOrdInfo = ""
    
    strSql = objMicSql.SqlGetOrderNumber(pWorkArea, pAccDt, pAccSeq, pTestCd, pMultiFg)
    Set objRs = New Recordset
    objRs.Open strSql, DBConn
    
    If objRs.EOF Then
        Set objRs = Nothing
        Exit Function
    Else
        sPtId = "" & objRs.Fields("ptid").Value
        sOrdDt = "" & objRs.Fields("orddt").Value
        sOrdNo = "" & objRs.Fields("ordno").Value
        sOrdSeq = "" & objRs.Fields("ordseq").Value
        
        pOrdInfo = sPtId & COL_DIV & sOrdDt & COL_DIV & sOrdNo & COL_DIV & sOrdSeq
        
        Set objRs = Nothing
        Set objRs = New Recordset
        objRs.Open objMicSql.SqlOrdVerifyCheck(sPtId, sOrdDt, sOrdNo, sOrdSeq), DBConn
        If objRs.EOF Then
            IsAllVerified = True
        Else
            IsAllVerified = False
        End If
    End If
    Set objRs = Nothing
    
End Function
