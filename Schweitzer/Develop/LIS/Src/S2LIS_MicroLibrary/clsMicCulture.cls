VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISMicCulture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private objMicSql As New clsLISSqlMicRst
Private objMicLib As New clsLISMicroLib
'----------------------------------------------------------------------------------
'1. 미생물 Nogrowth Batch 결과등록 관련 메쏘드
'----------------------------------------------------------------------------------

'ssTable 스프레드 정보
'    1   |    2   |   3  | 4 |  5   |   6    |   7    | 8 |   9    |     10       |  11  | 12 | 13  |  14
'접수번호|환자번호|환자명|S/A|검체명|최근결과|현재결과|MIC|검사코드|검체군(WsCode)|WsUnit|보류|SpcCd|Warning

Public Function DispNogrowthList(ByRef ssTable As Object, ByVal pWsCode As String, _
                                 ByVal pWsUnit As String, ByVal pRTs As String) As Long
    Dim dsNWS       As Recordset
    Dim sLabNo      As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim strLabNum   As String
    
    Dim i           As Integer
    Dim iCnt        As Integer
    
    Set dsNWS = New Recordset
    dsNWS.Open objMicSql.SQLReadBatchWorkList(1, pWsCode, pWsUnit, pRTs), DBConn
    
    iCnt = 0
    strLabNum = ""
    
    With ssTable
        .ReDraw = False
        .MaxRows = 0    'dsNWS.RecordCount
            
        If Not dsNWS.EOF Then
            
            While Not dsNWS.EOF
                
                objMicLib.Workarea = dsNWS.Fields("workarea").Value & ""
                objMicLib.Accdt = dsNWS.Fields("accdt").Value & ""
                objMicLib.Accseq = dsNWS.Fields("accseq").Value & ""
                objMicLib.PtId = dsNWS.Fields("ptid").Value & ""
                objMicLib.TestCd = dsNWS.Fields("testcd").Value & ""
                objMicLib.SpcCd = dsNWS.Fields("spccd").Value & ""
                
                sWorkArea = "" & dsNWS.Fields("workarea").Value
                sAccDt = "" & dsNWS.Fields("accdt").Value
                sAccDt = Mid$(sAccDt, 3, Len(sAccDt) - 2)
                sAccSeq = "" & dsNWS.Fields("accseq").Value
                
                If strLabNum <> sWorkArea & "-" & sAccDt & "-" & sAccSeq Then
                    
                    iCnt = iCnt + 1
                    
                    .MaxRows = iCnt
                    .Row = iCnt
                    .Col = ColNo.cnAccNo:  .Text = "" & sWorkArea & "-" & sAccDt & "-" & sAccSeq
                    .Col = ColNo.cnPtid:   .Text = "" & dsNWS.Fields("ptid").Value
                    .Col = ColNo.cnPtNm:   .Text = "" & dsNWS.Fields("ptnm").Value
                    .Col = ColNo.cnSA:     .Text = "" & dsNWS.Fields("sex").Value & "/" & (Val("" & dsNWS.Fields("ageday").Value) \ 365) + 1
                    .Col = ColNo.cnSpcNm:  .Text = "" & dsNWS.Fields("spcnm").Value
                    
                    '## 5.0.12: 이상대(2005-05-30)
                    '   - 과거결과를 빨강색으로 표시
                    .Col = ColNo.cnLstRst: .ForeColor = vbRed: .Text = objMicLib.GetNoGrowthLatestRst
                    .Col = ColNo.cnCurRst: .Text = "" & dsNWS.Fields("rst").Value
                    .Col = ColNo.cnMic:
                    .Col = ColNo.cnTestCd: .Text = .Text & "'" & dsNWS.Fields("testcd").Value & "" & "'"
                    .Col = ColNo.cnWsCd:   .Text = pWsCode
                    .Col = ColNo.cnWsUnit: .Text = pWsUnit
                    .Col = ColNo.cnHold:
                    .Col = ColNo.cnSpcCd:  .Text = dsNWS.Fields("spccd").Value & ""
                    .Col = ColNo.cnWarn:
                    .Col = ColNo.cnBarCode:  .Text = dsNWS.Fields("barcode").Value

                    strLabNum = sWorkArea & "-" & sAccDt & "-" & sAccSeq
                Else
                    '** 중요 사용자 요청으로 인해 Fix 처리 함 By M.G.Choi 2008.02.19
                    '   Fix 처리 이유:미생물 워크시트 마스터 설정 Miss 로 변경해야 하는데 기존데이터로 인해 변경불가 함...
                    If dsNWS.Fields("testcd").Value & "" = "B4051" Then
                        .Col = ColNo.cnTestCd: .Text = .Text & ",'" & dsNWS.Fields("testcd").Value & "'" & ",'B4062'"
                    Else
                        .Col = ColNo.cnTestCd: .Text = .Text & ",'" & dsNWS.Fields("testcd").Value & "'"
                    End If
                    
                    '## 원본 =============================================================================
'                    .Col = ColNo.cnTestCd: .Text = .Text & ",'" & dsNWS.Fields("testcd").Value & "'"
                    '=====================================================================================
                End If
                
                .Col = ColNo.cnMic:        .Text = IIf(dsNWS.Fields("rsttype").Value = MRT_MicSen, "Y", "")    'MIC 감수성인 경우...

                dsNWS.MoveNext
            Wend
    
        End If
        
        .RowHeight(-1) = 11
        .ReDraw = True
    End With
    
    Set dsNWS = Nothing
    DispNogrowthList = iCnt
    
End Function

Public Function DispVitekNogrowth(ByRef ssTable As Object, ByVal pWsCode As String, _
                                  ByVal pWsUnit As String, ByVal pRTs As String) As Long
    Dim dsNWS       As Recordset
    Dim sLabNo      As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim strLabNum   As String
    Dim i           As Integer
    Dim iCnt        As Integer
    
    Set dsNWS = New Recordset
    dsNWS.Open objMicSql.SQLReadBatchWorkList(1, pWsCode, pWsUnit, pRTs), DBConn
    
    iCnt = 0
    strLabNum = ""
    
    With ssTable
        .ReDraw = False
        .MaxRows = 0    'dsNWS.RecordCount
            
        If Not dsNWS.EOF Then
            
            While Not dsNWS.EOF
        
                sWorkArea = "" & dsNWS.Fields("workarea").Value
                sAccDt = "" & dsNWS.Fields("accdt").Value
                sAccDt = Mid$(sAccDt, 3, Len(sAccDt) - 2)
                sAccSeq = "" & dsNWS.Fields("accseq").Value
                
                If strLabNum <> sWorkArea & "-" & sAccDt & "-" & sAccSeq Then
                    
                    iCnt = iCnt + 1
                    
                    .MaxRows = iCnt
                    .Row = iCnt
                    .Col = 1: .Text = "" & sWorkArea & "-" & sAccDt & "-" & sAccSeq
                    .Col = 2: .Text = "" & dsNWS.Fields("ptid").Value
                    .Col = 3: .Text = "" & dsNWS.Fields("ptnm").Value
                    .Col = 4: .Text = "" & dsNWS.Fields("sex").Value & "/" & (Val("" & dsNWS.Fields("ageday").Value) \ 365) + 1
                    .Col = 5: .Text = "" & dsNWS.Fields("spcnm").Value
                    .Col = 6: .Text = "" & dsNWS.Fields("rst").Value
                    .Col = 8: .Text = .Text & "'" & dsNWS.Fields("testcd").Value & "'"
                    .Col = 9: .Text = Format("" & dsNWS.Fields("dob").Value, CS_DateLongMask)
                    .Col = 10: .Text = "" & dsNWS.Fields("spccd").Value
                    .Col = 11: .Text = Format("" & dsNWS.Fields("coldt").Value, CS_DateLongMask)
                    .Col = 12: .Text = Format("" & dsNWS.Fields("coltm").Value, CS_TimeLongMask)
                    .Col = 13: .Text = Format("" & dsNWS.Fields("rcvdt").Value, CS_DateLongMask)
                    .Col = 14: .Text = Format("" & dsNWS.Fields("rcvtm").Value, CS_TimeLongMask)
                    
                    strLabNum = sWorkArea & "-" & sAccDt & "-" & sAccSeq
                Else
                    .Col = 8: .Text = .Text & ",'" & dsNWS.Fields("testcd").Value & "'"
                End If
                
                .Col = 7: .Text = IIf(dsNWS.Fields("rsttype").Value = MRT_MicSen, "Y", "")   'MIC 감수성인 경우...
                
                dsNWS.MoveNext
            Wend
    
        End If
        
        .RowHeight(-1) = 11
        .ReDraw = True
    End With
    
    dsNWS.Close: Set dsNWS = Nothing
    DispVitekNogrowth = iCnt
    
End Function

Public Function DispGrowthList(ByRef lstGList As Object, ByVal pWsCode As String, _
                               ByVal pWsUnit As String) As Long
    Dim dsGWS   As Recordset
    Dim sWA     As String
    Dim sWDt    As String
    Dim sWSeq   As String
    Dim iCnt    As Integer
    
    Set dsGWS = New Recordset
    dsGWS.Open objMicSql.SQLReadBatchWorkList(2, pWsCode, pWsUnit), DBConn
    
    lstGList.Clear
    iCnt = 0
    
    While Not dsGWS.EOF
        iCnt = iCnt + 1
    
        sWA = dsGWS.Fields("workarea").Value & ""
        sWDt = dsGWS.Fields("accdt").Value & ""
        sWDt = Mid$(sWDt, 3, Len(sWDt) - 2)
        sWSeq = dsGWS.Fields("accseq").Value & ""
        
        lstGList.AddItem "" & sWA & "-" & sWDt & "-" & sWSeq
        
        dsGWS.MoveNext
    Wend

    dsGWS.Close: Set dsGWS = Nothing
    
    DispGrowthList = iCnt
    
End Function

Public Function DispSensiList(ByRef lstGList As Object, ByVal pWsCode As String, _
                              ByVal pWsUnit As String) As Long
    Dim dsGWS   As Recordset
    Dim sWA     As String
    Dim sWDt    As String
    Dim sWSeq   As String
    Dim iCnt    As Integer
    
    Set dsGWS = New Recordset
    dsGWS.Open objMicSql.SQLReadBatchWorkList(5, pWsCode, pWsUnit), DBConn
    
    lstGList.Clear
    iCnt = 0
    
    While Not dsGWS.EOF
        iCnt = iCnt + 1
    
        sWA = dsGWS.Fields("workarea").Value & ""
        sWDt = dsGWS.Fields("accdt").Value & ""
        sWDt = Mid$(sWDt, 3, Len(sWDt) - 2)
        sWSeq = dsGWS.Fields("accseq").Value & ""
        
        lstGList.AddItem "" & sWA & "-" & sWDt & "-" & sWSeq
        
        dsGWS.MoveNext
    Wend

    dsGWS.Close: Set dsGWS = Nothing
    
    DispSensiList = iCnt
End Function

Public Function DispFinalList(ByRef lstFList As Object, ByVal pWsCode As String, _
                              ByVal pWsUnit As String) As Long
    Dim dsFWS   As Recordset
    Dim sWA     As String
    Dim sWDt    As String
    Dim sWSeq   As String
    Dim iCnt    As Integer
    
    Set dsFWS = New Recordset
    dsFWS.Open objMicSql.SQLReadBatchWorkList(3, pWsCode, pWsUnit), DBConn

    lstFList.Clear
    iCnt = 0
    
    While Not dsFWS.EOF
        iCnt = iCnt + 1
    
        sWA = dsFWS.Fields("workarea").Value & ""
        sWDt = dsFWS.Fields("accdt").Value & ""
        sWDt = Mid$(sWDt, 3, Len(sWDt) - 2)
        sWSeq = dsFWS.Fields("accseq").Value & ""
        
        lstFList.AddItem "" & sWA & "-" & sWDt & "-" & sWSeq
        dsFWS.MoveNext
    Wend
        
    dsFWS.Close: Set dsFWS = Nothing
    DispFinalList = iCnt
    
End Function

Public Function DispHoldingList(ByRef ssHTable As Object, ByVal pWsCode As String, _
                                 ByVal pWsUnit As String, ByVal pRTs As String, _
                                 Optional ByVal AppendFg As Boolean = False) As Long
    Dim objTmpDic   As clsDictionary
    Dim dsHWS       As Recordset
    Dim sLabNo      As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim strKey1     As String
    Dim strKey2     As String
    Dim strKey3     As String
    Dim i           As Integer
    Dim iCnt        As Integer
    
    Set objTmpDic = New clsDictionary
    Set dsHWS = New Recordset
    dsHWS.Open objMicSql.SQLReadBatchWorkList(4, pWsCode, pWsUnit, pRTs), DBConn
       
    ssHTable.ReDraw = False
    
    iCnt = 0

    objTmpDic.Clear
    objTmpDic.FieldInialize "workarea,accdt,accseq", "rowno"
    
    With ssHTable
    
        While Not dsHWS.EOF
    
            objMicLib.Workarea = dsHWS.Fields("workarea").Value & ""
            objMicLib.Accdt = dsHWS.Fields("accdt").Value & ""
            objMicLib.Accseq = dsHWS.Fields("accseq").Value & ""
            objMicLib.PtId = dsHWS.Fields("ptid").Value & ""
            objMicLib.TestCd = dsHWS.Fields("testcd").Value & ""
            objMicLib.SpcCd = dsHWS.Fields("spccd").Value & ""
            
            sWorkArea = "" & dsHWS.Fields("workarea").Value
            sAccDt = "" & dsHWS.Fields("accdt").Value
            sAccDt = Mid$(sAccDt, 3, Len(sAccDt) - 2)
            sAccSeq = "" & dsHWS.Fields("accseq").Value
        
            strKey1 = sWorkArea & "-" & sAccDt & "-" & sAccSeq
            strKey2 = pWsCode
            strKey3 = pWsUnit
            
            With ssHTable
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = ColNo.cnAccNo
                    If .Value = strKey1 Then
                        .Col = ColNo.cnWsCd
                        If .Value = strKey2 Then
                            .Col = ColNo.cnWsUnit
                            If .Value = strKey3 Then GoTo Skip
                        End If
                    End If
                Next
            End With
            
            If objTmpDic.Exists(medConcatString(COL_DIV, sWorkArea, sAccDt, sAccSeq)) Then
                objTmpDic.KeyChange (medConcatString(COL_DIV, sWorkArea, sAccDt, sAccSeq))
                .Row = Val(objTmpDic.Fields("rowno"))
                .Col = ColNo.cnMic: .Text = IIf(dsHWS.Fields("rsttype").Value & "" = MRT_MicSen, "Y", "")    'MIC 감수성인 경우...
                .Col = ColNo.cnTestCd: .Text = .Text & ",'" & dsHWS.Fields("testcd").Value & "'"
            Else
                iCnt = iCnt + 1
                
                If AppendFg Then
                    ssHTable.MaxRows = ssHTable.MaxRows + 1
                Else
                    ssHTable.MaxRows = iCnt
                End If
                
                .Row = .MaxRows
                .Col = ColNo.cnAccNo:  .Text = "" & sWorkArea & "-" & sAccDt & "-" & sAccSeq
                .Col = ColNo.cnPtid:   .Text = "" & dsHWS.Fields("ptid").Value
                .Col = ColNo.cnPtNm:   .Text = "" & dsHWS.Fields("ptnm").Value
                .Col = ColNo.cnSA:     .Text = "" & dsHWS.Fields("sex").Value & "/" & (Val("" & dsHWS.Fields("ageday").Value) \ 365) + 1
                .Col = ColNo.cnSpcNm:  .Text = "" & dsHWS.Fields("spcnm").Value
                .Col = ColNo.cnLstRst: .Text = objMicLib.GetNoGrowthLatestRst
                .Col = ColNo.cnCurRst: .Text = "" & dsHWS.Fields("rst").Value
                .Col = ColNo.cnMic:    .Text = IIf(dsHWS.Fields("rsttype").Value = MRT_MicSen, "Y", "")    'MIC 감수성인 경우...
                .Col = ColNo.cnTestCd: .Text = .Text & "'" & dsHWS.Fields("testcd").Value & "'"
                .Col = ColNo.cnWsCd:   .Text = pWsCode
                .Col = ColNo.cnWsUnit: .Text = pWsUnit
                .Col = ColNo.cnHold:
                .Col = ColNo.cnSpcCd:  .Text = dsHWS.Fields("spccd").Value & ""
               
                .Col = ColNo.cnWarn:   .Text = objMicLib.GetWarningForDisp(sWorkArea & "-" & sAccDt & "-" & sAccSeq, dsHWS.Fields("ptid").Value & "", dsHWS.Fields("testcd").Value & "", dsHWS.Fields("spccd").Value & "")
               
                objTmpDic.AddNew medConcatString(COL_DIV, sWorkArea, sAccDt, sAccSeq), iCnt
            End If
Skip:
            dsHWS.MoveNext
            
        Wend
        
        .ReDraw = True
        
    End With
    
    Set objTmpDic = Nothing
    dsHWS.Close: Set dsHWS = Nothing
    DispHoldingList = iCnt
    
End Function

Public Function DispVitekHolding(ByRef ssHTable As Object, ByVal pWsCode As String, _
                                 ByVal pWsUnit As String, ByVal pRTs As String, _
                                 Optional ByVal AppendFg As Boolean = False) As Long
    Dim objTmpDic   As clsDictionary
    Dim dsHWS       As Recordset
    Dim sLabNo      As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim strKey1     As String
    Dim strKey2     As String
    Dim strKey3     As String
    Dim i           As Integer
    Dim iCnt        As Integer
    
    Set objTmpDic = New clsDictionary
    Set dsHWS = New Recordset
    dsHWS.Open objMicSql.SQLReadBatchWorkList(4, pWsCode, pWsUnit, pRTs), DBConn
    
    ssHTable.ReDraw = False
    
    iCnt = 0

    objTmpDic.Clear
    objTmpDic.FieldInialize "workarea,accdt,accseq", "rowno"
    
    With ssHTable
    
        While Not dsHWS.EOF
    
            sWorkArea = "" & dsHWS.Fields("workarea").Value
            sAccDt = "" & dsHWS.Fields("accdt").Value
            sAccDt = Mid$(sAccDt, 3, Len(sAccDt) - 2)
            sAccSeq = "" & dsHWS.Fields("accseq").Value
        
            strKey1 = sWorkArea & "-" & sAccDt & "-" & sAccSeq
            strKey2 = pWsCode
            strKey3 = pWsUnit
            
            With ssHTable
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = 1
                    If .Value = strKey1 Then
                        .Col = 9
                        If .Value = strKey2 Then
                            .Col = 10
                            If .Value = strKey3 Then GoTo Skip
                        End If
                    End If
                Next
            End With
            
            If objTmpDic.Exists(medConcatString(COL_DIV, sWorkArea, sAccDt, sAccSeq)) Then
                objTmpDic.KeyChange (medConcatString(COL_DIV, sWorkArea, sAccDt, sAccSeq))
                .Row = Val(objTmpDic.Fields("rowno"))
                .Col = 7: .Text = IIf(dsHWS.Fields("rsttype").Value = MRT_MicSen, "Y", "")   'MIC 감수성인 경우...
                .Col = 8: .Text = .Text & ",'" & dsHWS.Fields("testcd").Value & "'"
            Else
                iCnt = iCnt + 1
                
                If AppendFg Then
                    ssHTable.MaxRows = ssHTable.MaxRows + 1
                Else
                    ssHTable.MaxRows = iCnt
                End If
                
                .Row = .MaxRows
                .Col = 1: .Text = "" & sWorkArea & "-" & sAccDt & "-" & sAccSeq
                .Col = 2: .Text = "" & dsHWS.Fields("ptid").Value
                .Col = 3: .Text = "" & dsHWS.Fields("ptnm").Value
                .Col = 4: .Text = "" & dsHWS.Fields("sex").Value & "/" & (Val("" & dsHWS.Fields("ageday").Value) \ 365) + 1
                '.col = 5: .Text = "" & dsHWS.Fields("spccd").Value
                .Col = 5: .Text = "" & dsHWS.Fields("spcnm").Value
                .Col = 6: .Text = "" & dsHWS.Fields("rst").Value
                .Col = 7: .Text = IIf(dsHWS.Fields("rsttype").Value = MRT_MicSen, "Y", "")   'MIC 감수성인 경우...
                .Col = 8: .Text = .Text & "'" & dsHWS.Fields("testcd").Value & "'"
                .Col = 9: .Text = pWsCode
                .Col = 10: .Text = pWsUnit
                .Col = 11: .Text = Format("" & dsHWS.Fields("coldt").Value, CS_DateLongMask)
                .Col = 12: .Text = Format("" & dsHWS.Fields("coltm").Value, CS_TimeLongMask)
                .Col = 13: .Text = Format("" & dsHWS.Fields("rcvdt").Value, CS_DateLongMask)
                .Col = 14: .Text = Format("" & dsHWS.Fields("rcvtm").Value, CS_TimeLongMask)
                .Col = 15: .Text = Format("" & dsHWS.Fields("dob").Value, CS_DateLongMask)
                .Col = 16: .Text = "" & dsHWS.Fields("spccd").Value
                .Col = 17: .Text = Mid(sAccDt, 3) & Format(sAccSeq, "0000")
                 
                objTmpDic.AddNew medConcatString(COL_DIV, sWorkArea, sAccDt, sAccSeq), iCnt
            End If
Skip:
            dsHWS.MoveNext
            
        Wend
        
        .ReDraw = True
        
    End With
    
    Set objTmpDic = Nothing
    Set dsHWS = Nothing
    DispVitekHolding = iCnt
    
End Function

Public Function SaveNogrowthBatch(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, ByVal pTestCds As String, _
                                  ByVal pResult As String, ByVal pStatus As String, ByVal pDate As String, ByVal pTime As String, _
                                  ByVal pEmpId As String, ByVal pWsCd As String) As Boolean
    
    Dim sqlUpRst()  As String
    Dim sqlUpOrd    As String
    Dim sqlUpWs     As String
    Dim sWsSts      As String
    Dim strSql      As String
    Dim strOrdInfo  As String
    Dim strTmp      As String
    Dim sPtId       As String
    Dim sOrdDt      As String
    Dim sOrdNo      As String
    Dim sOrdSeq     As String
    Dim blnVerified As Boolean
    Dim blnVerify   As Boolean
    Dim i           As Long
    Dim FileNo      As Long
    
    '-- 전주예수병원 추가 변수=============
    Dim clsTransfer As New clsLISTransfer
    Dim strVfyDt As String
    Dim strKey As String
    '** 검사항목별 결과값 체크
    Dim RsG     As New ADODB.Recordset
    Dim RsU     As New ADODB.Recordset
    Dim strQry  As String
    Dim strTest As String
'    Dim strSenVal As String
'    Dim strSEQ    As String
'    Dim strIPAdr  As String
'    Dim strWrkDiv As String
'    Dim bFlag     As Boolean
'    Dim bINUPFlag As Boolean
    '======================================
    
    ReDim sqlUpdate(0)
    strSql = objMicSql.SQLSaveNogrowthBatch(pWorkArea, pAccDt, pAccSeq, pTestCds, _
                                                                    pResult, pStatus, pDate, pTime, pEmpId)
    DBConn.Execute strSql
    
    blnVerified = IsAllVerified(pWorkArea, pAccDt, pAccSeq, pTestCds, strOrdInfo, True)
    strTmp = medShift(strOrdInfo, LINE_DIV)
    
    sPtId = medGetP(strTmp, 1, COL_DIV)
    sOrdDt = medGetP(strTmp, 2, COL_DIV)
    sOrdNo = medGetP(strTmp, 3, COL_DIV)
    
    '** 전주예수병원 추가 루틴 ==================================================================
    ' * 최종결과만 전송하기로 함 : pStatus = enStsCd.StsCd_LIS_FinRst
    '---------------------------
    ' * 변경 중간결과 및 최종결과 전송 By M.G.Choi 2005.11.17
    
    If pStatus = enStsCd.StsCd_LIS_MidRst Or pStatus = enStsCd.StsCd_LIS_FinRst Then
        
        '-- 결과전송 시 필요정보
        Call clsTransfer.SqlOCS_Info_Mic(pWorkArea, pAccDt, pAccSeq)
                    
        strVfyDt = pDate & pTime
        
        '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
        If clsTransfer.WrkDiv = "3" Then
            
            strQry = clsTransfer.NoGrowth_Test(pWorkArea, pAccDt, pAccSeq, pTestCds)
            
            RsU.Open strQry, DBConn, adOpenForwardOnly, adLockReadOnly
            
            If RsU.EOF = False Then
                
                strTest = RsU.Fields("testcd").Value & ""
            
                Call clsTransfer.ResultValue_NoGrowth(strTest, pResult)
            
                If clsTransfer.SqlOCSSU2EXAMT_Find(sPtId, sOrdDt, sOrdNo) = True Then
                    strSql = clsTransfer.SqlOCSSU2EXAMT_UPDATE_New(sPtId, sOrdDt, sOrdNo, _
                                            "N", clsTransfer.ItemVal, "", "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "", enStsCd.StsCd_LIS_FinRst)
                Else
                    GoTo DBExecError
                End If
            End If
            RsU.Close
        ElseIf clsTransfer.WrkDiv = "4" Then
                
                strQry = clsTransfer.NoGrowth_Test(pWorkArea, pAccDt, pAccSeq, pTestCds)
                
                RsG.Open strQry, DBConn, adOpenForwardOnly, adLockReadOnly
                
                If RsG.EOF = False Then
                    
                    strTest = RsG.Fields("testcd").Value & ""
                
                    Call clsTransfer.ResultValue_NoGrowth(strTest, pResult)
                    
                    If clsTransfer.SqlOCSSG2EXAMT_Find(sPtId, sOrdDt, sOrdNo) = True Then
                        strSql = clsTransfer.SqlOCSSG2EXAMT_UPDATE_NEW(sPtId, sOrdDt, sOrdNo, _
                                                "N", clsTransfer.ItemVal, "", "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "", enStsCd.StsCd_LIS_FinRst)
                    Else
                        GoTo DBExecError
                    End If
                End If
                RsG.Close
        Else
            If clsTransfer.SqlOCSMDEXMORT_Find(sPtId, sOrdDt, sOrdNo) = True Then
                strSql = clsTransfer.SqlOCSMDEXMORT_UPDATE_NEW(sPtId, sOrdDt, sOrdNo, _
                                        "N", strVfyDt, pEmpId, "", "", "")
            Else
                GoTo DBExecError
            End If
        End If
        
        Call DBConn.Execute(strSql)
        
        '** 종건/일건 이외의 검사만 등록.
        If clsTransfer.WrkDiv <> "" Then
            GoTo OCSSKIP
        End If
                
        '-- 2. 건진외결과 INSERT
        '** 검사항목별 결과값 체크
        Dim Rs      As New ADODB.Recordset
        
        strQry = clsTransfer.NoGrowth_Test(pWorkArea, pAccDt, pAccSeq, pTestCds)
        
        Rs.Open strQry, DBConn, adOpenForwardOnly, adLockReadOnly
        
        If Rs.BOF = False Then
            Do Until Rs.EOF = True
                
                strTest = Rs.Fields("testcd").Value & ""
                
                Call clsTransfer.ResultValue_NoGrowth(strTest, pResult)
                
                If clsTransfer.SqlOCSRESULT_FLAG(sPtId, sOrdDt, sOrdNo, strTest) = False Then
                    strSql = clsTransfer.SqlOCSRESULT_INSERT_NoGrowth(sPtId, sOrdDt, sOrdNo, _
                                            strTest, clsTransfer.SlipCd, clsTransfer.ItemVal, "", "", "", "", _
                                            "", clsTransfer.RcvDt & clsTransfer.RCVTM, strVfyDt, pEmpId, strVfyDt, pEmpId, "", clsTransfer.RsltType, _
                                            "", "", "")
                Else
                    strSql = clsTransfer.SqlOCSRESULT_UPDATE_NoGrowth(sPtId, sOrdDt, sOrdNo, _
                                            strTest, clsTransfer.SlipCd, clsTransfer.ItemVal, "", "", "", "", _
                                            "", clsTransfer.RcvDt & clsTransfer.RCVTM, strVfyDt, pEmpId, strVfyDt, pEmpId, "", clsTransfer.RsltType, _
                                            "", "", "")
                End If
                       
                Call DBConn.Execute(strSql)
                
                Rs.MoveNext
            Loop
        End If
        
        Rs.Close
        Set Rs = Nothing
        Set RsU = Nothing
        Set RsG = Nothing
    End If
    '============================================================================================
    
OCSSKIP:

    FileNo = FreeFile
    
    If Dir(App.Path & "\LOG", vbDirectory) = "" Then
        MkDir App.Path & "\LOG"
    End If
        
    Open App.Path & "\LOG\MIC_LOG_" & Format(GetSystemDate, "yyyyMMdd") & ".log" For Append As #FileNo   ' 파일 이름을 작성합니다.
    
    Print #FileNo, GetSystemDate
    Print #FileNo, " << ------------------------------ START LOG ------------------------------ >> "
    Print #FileNo, strSql

'리턴값이 False일때 최종을 눌렀어도 중간결과로 남게되겠구만..
    blnVerified = IsAllVerified(pWorkArea, pAccDt, pAccSeq, pTestCds, strOrdInfo, True)
    strTmp = medShift(strOrdInfo, LINE_DIV)
    
    While Trim(strTmp) <> ""
        sPtId = medGetP(strTmp, 1, COL_DIV)
        sOrdDt = medGetP(strTmp, 2, COL_DIV)
        sOrdNo = medGetP(strTmp, 3, COL_DIV)
        sOrdSeq = medGetP(strTmp, 4, COL_DIV)
        blnVerify = medGetP(strTmp, 5, COL_DIV)
    
        If blnVerify Then
            ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
            sqlUpdate(UBound(sqlUpdate)) = objMicSql.SqlUpdateOrderStatus(sPtId, sOrdDt, sOrdNo, sOrdSeq, _
                                                pStatus, pDate, pTime, pEmpId)
        Else
            ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
            sqlUpdate(UBound(sqlUpdate)) = objMicSql.SqlUpdateOrderStatus(sPtId, sOrdDt, sOrdNo, sOrdSeq, _
                                                enStsCd.StsCd_LIS_MidRst, pDate, pTime, pEmpId)
        
            If pStatus = "5" Then
                Print #FileNo, "**************최종결과를 눌렀는데 중간으로 남았네**************"
            End If
        End If
       
        strTmp = medShift(strOrdInfo, LINE_DIV)
    Wend
    
    ' *** Parent Status 반영 (접수내역, 미생물 WS)
    ' 접수내역에 조회정보 및 Status 반영
    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
    sqlUpdate(UBound(sqlUpdate)) = SaveMicStatus(pWorkArea, pAccDt, pAccSeq, 1, pStatus, _
                                               pDate, pTime, pEmpId)
        
    ' Wscd와 wsunit도 필요하다
    If pStatus = enStsCd.StsCd_LIS_MidRst Then sWsSts = MWS_Ready
    If pStatus = enStsCd.StsCd_LIS_FinRst Then sWsSts = MWS_Final
    
    If pStatus = enStsCd.StsCd_LIS_FinRst Or pStatus = enStsCd.StsCd_LIS_MidRst Then
        ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLUpdateWoksheet(1, pWorkArea, pAccDt, pAccSeq, _
                                                                      pWsCd, "", sWsSts)
    End If
    
On Error GoTo DBExecError
    
    For i = 1 To UBound(sqlUpdate)
        If Trim(sqlUpdate(i)) <> "" Then
            Print #FileNo, sqlUpdate(i)
            DBConn.Execute sqlUpdate(i)
        End If
    Next
    
    Print #FileNo, " << ------------------------------ END LOG ------------------------------ >> "
    Close #FileNo
    
    SaveNogrowthBatch = True
    Exit Function

DBExecError:
    SaveNogrowthBatch = False
End Function

Private Function SaveMicStatus(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, ByVal pVCnt As String, _
                              ByVal pStatus As String, ByVal pDate As String, ByVal pTime As String, ByVal pEmpId As String) As String
    Dim dsAcc           As Recordset
    Dim sqlAcc          As String
    Dim sqlUpdAcc       As String
    Dim sStatus         As String
    Dim upStatus        As String
    Dim upInputCount    As Integer
    Dim iTotalCount     As Integer
    Dim iInputCount     As Integer

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

    SaveMicStatus = objMicSql.SQLUpdateStatus(pWorkArea, pAccDt, pAccSeq, sStatus, iInputCount, pDate, pTime, pEmpId)

End Function


Public Function SaveHoldingList(ByVal ssObj As Object, ByVal pWsCd As String, ByVal pWsUnit As String) As Boolean
        
    Dim sAccNo      As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim sqlUpWs     As String
    Dim i           As Long
    
    With ssObj
        For i = 1 To .MaxRows
            .Col = 1: .Row = i: sAccNo = .Text
            sWorkArea = medGetP(sAccNo, 1, "-"): sAccDt = medGetP(sAccNo, 2, "-"): sAccSeq = medGetP(sAccNo, 3, "-")
            sAccDt = IIf(Mid(sAccDt, 1, 1) = "9", "19" & sAccDt, "20" & sAccDt)
            
            ' WorkSheet Update (Stain결과만 있는 경우는 이미 제외되었으므로 신경안써도 됨)
            sqlUpWs = objMicSql.SQLUpdateWoksheet(1, sWorkArea, sAccDt, sAccSeq, _
                                                     pWsCd, pWsUnit, MWS_Holding)
        On Error GoTo Err_Trap
            DBConn.Execute sqlUpWs
        Next
    End With
    SaveHoldingList = True
    Exit Function
Err_Trap:
    SaveHoldingList = False

End Function

Public Function SaveOneStatus(ByVal pWsCd As String, ByVal pWsUnit As String, ByVal pWorkArea As String, _
                              ByVal pAccDt As String, ByVal pAccSeq As String, ByVal pStatus As String) As Boolean
        
    Dim sAccNo      As String
    Dim sWorkArea   As String
    Dim sAccDt      As String
    Dim sAccSeq     As String
    Dim sqlUpWs     As String
    Dim i           As Long
    
    ' WorkSheet Update
    sqlUpWs = objMicSql.SQLUpdateWoksheet(1, pWorkArea, pAccDt, pAccSeq, _
                                             pWsCd, pWsUnit, pStatus)
    On Error GoTo Err_Trap
    DBConn.Execute sqlUpWs
    SaveOneStatus = True
    Exit Function
Err_Trap:
    SaveOneStatus = False

End Function


'----------------------------------------------------------------------------------
'2. 미생물 Culture 결과등록 관련 메쏘드
'----------------------------------------------------------------------------------

Public Sub LoadMicrobe(ByRef lstMicCd As Object, ByRef lstMicNm As Object, ByRef lstMGroup As Object)
    'lstMicCd : 미생물코드리스트, lstMicNm : 미생물명리스트, lstMGroup : 균종리스트
    Dim dsMic As Recordset

    ' 균 코드 로드
    Set dsMic = New Recordset
    dsMic.Open objMicSql.SQLLoadMicrobe, DBConn

    lstMicCd.Clear: lstMicNm.Clear: lstMGroup.Clear
    lstMicCd.AddItem LIS_Nothing
    lstMicNm.AddItem LIS_Nothing
    lstMGroup.AddItem LIS_Nothing

    While Not dsMic.EOF
        lstMicCd.AddItem "" & dsMic.Fields("cdval1").Value
        lstMicNm.AddItem "" & dsMic.Fields("text1").Value
        lstMGroup.AddItem "" & dsMic.Fields("field2").Value
        dsMic.MoveNext
    Wend

    Set dsMic = Nothing

End Sub

Public Sub LoadQuantity(ByRef lstQty As Object)
    
    Dim dsQty As Recordset

    Set dsQty = New Recordset
    dsQty.Open objMicSql.SQLLoadQuantity, DBConn

    If dsQty.RecordCount < 1 Then
        MsgBox "정도 코드가 등록되어 있지 않습니다.", vbExclamation, "정도코드"
        Set dsQty = Nothing
        Exit Sub
    End If

    lstQty.Clear
    dsQty.MoveFirst
    While Not dsQty.EOF
        lstQty.AddItem "" & dsQty.Fields("CdVal1").Value & vbTab & dsQty.Fields("field2").Value
        dsQty.MoveNext
    Wend

    Set dsQty = Nothing

End Sub

Public Function LoadAntibiotic(ByVal strSenType As String, ByVal strMicGroup As String, _
                               ByRef strAntiList As String) As Long
'항생제리스트와 갯수를 반환
    
    Dim dsMG As Recordset

    Set dsMG = New Recordset
    dsMG.Open objMicSql.SQLLoadAntibiotic(strSenType, strMicGroup), DBConn
    
    If dsMG.EOF Then
        LoadAntibiotic = 0
    Else
        strAntiList = dsMG.Fields("text1").Value
        LoadAntibiotic = Val(medShift(strAntiList, ";"))
        strAntiList = Replace(strAntiList, ";", vbTab)
    End If
    
    Set dsMG = Nothing
    
End Function

Public Function GetCultureType(ByVal pWorkArea As String, ByVal pAccDt As String, _
                               ByVal pAccSeq As String) As String
                               
    Dim objRs As Recordset
    
    Set objRs = New Recordset
    objRs.Open objMicSql.SQLLoadTestCdByLabNo(pWorkArea, pAccDt, pAccSeq), DBConn
    
    If objRs.EOF Then
        GetCultureType = MRT_GenSen
    Else
        If objRs.RecordCount > 1 Then
            GetCultureType = ""
        Else
            GetCultureType = objRs.Fields("RstType").Value
        End If
    End If
    Set objRs = Nothing
End Function

Public Function GetTestByLabNo(ByVal pWorkArea As String, ByVal pAccDt As String, _
                               ByVal pAccSeq As String, ByRef objList As Object) As Long
                               
    Dim objRs As Recordset
    
    objList.Clear
    GetTestByLabNo = 0
    
    Set objRs = New Recordset
    objRs.Open objMicSql.SQLLoadTestCdByLabNo(pWorkArea, pAccDt, pAccSeq), DBConn
    
    While Not objRs.EOF
        
        objList.AddItem "" & objRs.Fields("testcd").Value & vbTab & objRs.Fields("abbrnm10").Value
        objRs.MoveNext
        
        GetTestByLabNo = GetTestByLabNo + 1
    Wend
    
    Set objRs = Nothing

End Function

Public Function DispStainResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                ByVal pAccSeq As String, ByRef lstGramStain As Object) As String
    
    Dim dsStain As Recordset

    Set dsStain = New Recordset
    dsStain.Open objMicSql.SQLGetStainResult(pWorkArea, pAccDt, pAccSeq), DBConn
    
    DispStainResult = ""
    lstGramStain.Clear
    
    While Not dsStain.EOF
        If "" & dsStain.Fields("rstdiv").Value = "*" Then
            DispStainResult = "" & dsStain.Fields("abbrnm10").Value
        Else
            lstGramStain.AddItem "" & Format(dsStain.Fields("abbrnm10").Value, "@@@@@@@@@@!") & vbTab & _
                                      dsStain.Fields("rstnm").Value
        End If
        dsStain.MoveNext
    Wend

    Set dsStain = Nothing

End Function

Public Function DispGrowthResult(ByVal pWorkArea As String, ByVal pAccDt As String, _
                                 ByVal pAccSeq As String, ByVal pWsCd As String, ByVal pRstTp As String) As String
    
'pWsCd - 미생물 Worksheet코드

    Dim dsNG    As Recordset
    Dim sSenFg  As String
    Dim sNGRst  As String
    Dim strDate As String
    
    Set dsNG = New Recordset
    dsNG.Open objMicSql.SQLGetGrowthResult(pWorkArea, pAccDt, pAccSeq, pWsCd, pRstTp), DBConn

    If Not dsNG.EOF Then
        sNGRst = "" & dsNG.Fields("rst").Value
        sSenFg = "" & dsNG.Fields("senfg").Value
        strDate = "" & dsNG.Fields("vfydt").Value
    End If
    
    DispGrowthResult = sNGRst & COL_DIV & sSenFg & COL_DIV & strDate
    Set dsNG = Nothing

End Function

Public Function DispSensiResult(ByRef ssSusc As Object, ByVal pWorkArea As String, ByVal pAccDt As String, _
                                ByVal pAccSeq As String, ByVal pTestCd As String, _
                                Optional ByVal pMfySeq As String = "0") As String
        
    Dim dsGR    As Recordset
    Dim sBuffer As String
    Dim sTmp1   As String
    Dim sTmp2   As String
    Dim sTmp3   As String
    Dim sFld    As String
    Dim ssRst   As String
    Dim iCnt    As Long
    Dim iACnt   As Long
    Dim j       As Long
    
    Set dsGR = New Recordset
    dsGR.Open objMicSql.SQLGetSensiResult(pWorkArea, pAccDt, pAccSeq, pTestCd, pMfySeq), DBConn
    sBuffer = "": iCnt = 0
    
    With ssSusc
        
        
        
        While Not dsGR.EOF
    
            iCnt = iCnt + 1
    
            sTmp1 = "": sTmp2 = "": sTmp3 = ""
            iACnt = Val(dsGR.Fields("SCnt").Value)
            
            For j = 1 To iACnt
                sFld = "SRst" & j
                ssRst = "" & dsGR.Fields(sFld).Value
                sTmp1 = sTmp1 & vbTab & medGetP(ssRst, 1, ";")  '항생제코드
                sTmp2 = sTmp2 & vbTab & medGetP(ssRst, 2, ";")  '감수성결과
                sTmp3 = sTmp3 & vbTab & medGetP(ssRst, 3, ";")  'MIC결과
            Next j
            
            
            sBuffer = sBuffer & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & sTmp1 & vbCr
            sBuffer = sBuffer & _
                      dsGR.Fields("mnm").Value & vbTab & dsGR.Fields("mnmcd").Value & vbTab & dsGR.Fields("mdiv").Value & vbTab & _
                      dsGR.Fields("micfg").Value & vbTab & dsGR.Fields("mqt").Value & vbTab & _
                      dsGR.Fields("mqtcd").Value & vbTab & _
                      iACnt & sTmp2 & vbCr
            sBuffer = sBuffer & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & sTmp3 & vbCr
            dsGR.MoveNext
            
        Wend
        Set dsGR = Nothing

        .Col = 1: .col2 = .MaxCols
        .Row = 1: .row2 = .MaxRows
        .blockmode = True
        .Clip = sBuffer
        .blockmode = False
    
    End With
    
End Function

Public Function SaveNGResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                             ByVal pTestCd As String, ByVal pResult As String, ByVal pStatus As String, _
                             ByVal pEmpId As String, ByVal pFNote As String, ByVal pRmkCd As String, _
                             Optional ByVal pMfySeq As Long = 0) As Boolean
    Dim blnVerified As Boolean
    Dim sqlUpdate() As String
    Dim pDate       As String
    Dim pTime       As String
    Dim strSql      As String
    Dim strOrdInfo  As String
    Dim strTmp      As String
    Dim sPtId       As String
    Dim sOrdDt      As String
    Dim sOrdNo      As String
    Dim sOrdSeq     As String
    Dim i           As Long
    
    ' 시스템 일자/시간 설정
    pDate = Format(GetSystemDate, CS_DateDbFormat)
    pTime = Format(GetSystemDate, CS_TimeDbFormat)

    
    ReDim sqlUpdate(0)

    '*** 감수성결과 수정 시 미생물 결과 History 작업
    If pMfySeq > 0 Then Call SetHistory(pWorkArea, pAccDt, pAccSeq, pTestCd)
    '***

    strSql = objMicSql.SQLVerifyResult(pWorkArea, pAccDt, pAccSeq, pTestCd, _
                                         pResult, pStatus, pDate, pTime, pEmpId, "", pMfySeq)
    DBConn.Execute strSql
    

    blnVerified = IsAllVerified(pWorkArea, pAccDt, pAccSeq, pTestCd, strOrdInfo, True)
    strTmp = medShift(strOrdInfo, LINE_DIV)
    
    While Trim(strTmp) <> ""
        sPtId = medGetP(strTmp, 1, COL_DIV)
        sOrdDt = medGetP(strTmp, 2, COL_DIV)
        sOrdNo = medGetP(strTmp, 3, COL_DIV)
        sOrdSeq = medGetP(strTmp, 4, COL_DIV)
        
        If blnVerified Then
            ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
            sqlUpdate(UBound(sqlUpdate)) = objMicSql.SqlUpdateOrderStatus(sPtId, sOrdDt, sOrdNo, sOrdSeq, _
                                                pStatus, pDate, pTime, pEmpId)
        Else
            ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
            sqlUpdate(UBound(sqlUpdate)) = objMicSql.SqlUpdateOrderStatus(sPtId, sOrdDt, sOrdNo, sOrdSeq, _
                                                enStsCd.StsCd_LIS_MidRst, pDate, pTime, pEmpId)
        End If
        strTmp = medShift(strOrdInfo, LINE_DIV)
    Wend
    
    'WorkSheet Update (최종결과일때만 Update) & Mic감수성이 아닐경우만...
    If pStatus >= enStsCd.StsCd_LIS_FinRst Then  'And fMicFg = "" Then
        ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLUpdateWoksheet(3, pWorkArea, pAccDt, pAccSeq, "", "", MWS_Final)
    End If

    ' *** Parent Status 반영 (접수내역, 미생물 WS)
    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
    sqlUpdate(UBound(sqlUpdate)) = SavePStatus(pWorkArea, pAccDt, pAccSeq, 1, pStatus, _
                                               pDate, pTime, pEmpId)
'2003/10/08 수정전 원본
'    ' *** Foot Note 저장 *** 검체 Remark 저장
'    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
'    sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLSaveFootNote(pWorkArea, pAccDt, pAccSeq, 0, "", "", 0)
'    If Trim(pFNote) <> "" Then
'        ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
'        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLSaveFootNote(pWorkArea, pAccDt, pAccSeq, 1, pEmpId, pFNote, 1)
'    End If
'
'    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
'    sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLSaveRemark(pWorkArea, pAccDt, pAccSeq, 1, pRmkCd)
    
'Modify By legends 2003/10/08
'Footnote난 Remark가 없어도 수정하거나 등록만 하면 내역을 만들어주는 문제가 있네...

    ' *** Foot Note 저장 *** 검체 Remark 저장
    If pMfySeq = 0 Then
        ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLSaveFootNote(pWorkArea, pAccDt, pAccSeq, 0, "", "", 0)
    End If
    
    If Trim(pFNote) <> "" Then
        'FootNote저장
        ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLSaveFootNote(pWorkArea, pAccDt, pAccSeq, pMfySeq + 1, pEmpId, pFNote, 1)
        
        '검체 Remark저장
        ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLSaveRemark(pWorkArea, pAccDt, pAccSeq, 1, pRmkCd)
    End If
    
On Error GoTo DBExecError
    For i = 1 To UBound(sqlUpdate)
        DBConn.Execute sqlUpdate(i)
    Next
    SaveNGResult = True

    Exit Function

DBExecError:
    SaveNGResult = False
   
End Function

Public Function IsAllVerified(ByVal pWorkArea As String, ByVal pAccDt As String, _
                              ByVal pAccSeq As String, ByVal pTestCd As String, _
                              ByRef pOrdInfo As String, _
                              Optional ByVal pMultiFg As Boolean = False) As Boolean
    Dim objRs   As Recordset
    Dim objRs1  As Recordset
    Dim strSql  As String
    Dim sPtId   As String
    Dim sOrdDt  As String
    Dim sOrdNo  As String
    Dim sOrdSeq As String
    
    IsAllVerified = True
    pOrdInfo = ""
    
    strSql = objMicSql.SqlGetOrderNumber(pWorkArea, pAccDt, pAccSeq, pTestCd, pMultiFg) '처방번호 갖구 오는넘
    Set objRs = New Recordset
    objRs.Open strSql, DBConn
    
    Do Until objRs.EOF
        sPtId = "" & objRs.Fields("ptid").Value
        sOrdDt = "" & objRs.Fields("orddt").Value
        sOrdNo = "" & objRs.Fields("ordno").Value
        sOrdSeq = "" & objRs.Fields("ordseq").Value
        
        pOrdInfo = pOrdInfo & sPtId & COL_DIV & sOrdDt & COL_DIV & sOrdNo & COL_DIV & sOrdSeq & COL_DIV
        
        Set objRs1 = Nothing
        Set objRs1 = New Recordset
        objRs1.Open objMicSql.SqlOrdVerifyCheck(sPtId, sOrdDt, sOrdNo, sOrdSeq), DBConn
        
        If objRs1.EOF Then
            IsAllVerified = True
        Else
            IsAllVerified = False
        End If
                    
        pOrdInfo = pOrdInfo & sPtId & COL_DIV & sOrdDt & COL_DIV & sOrdNo & COL_DIV & sOrdSeq & COL_DIV & IsAllVerified & LINE_DIV
                    
        Set objRs1 = Nothing
            
        objRs.MoveNext
    Loop
    
    Set objRs = Nothing
    Set objRs1 = Nothing
End Function

Public Function IsSensiAllVerified(ByVal pWorkArea As String, ByVal pAccDt As String, _
                              ByVal pAccSeq As String, ByVal pTestCd As String, _
                              ByRef pOrdInfo As String, _
                              Optional ByVal pMultiFg As Boolean = False) As Boolean
    Dim objRs   As Recordset
    Dim objRs1  As Recordset
    Dim strSql  As String
    Dim sPtId   As String
    Dim sOrdDt  As String
    Dim sOrdNo  As String
    Dim sOrdSeq As String
    
    IsSensiAllVerified = True
    pOrdInfo = ""
    strSql = objMicSql.SqlGetSenOrdNumber(pWorkArea, pAccDt, pAccSeq, pTestCd, pMultiFg)
    Set objRs = New Recordset
    objRs.Open strSql, DBConn
    
    If objRs.EOF Then
        Exit Function
    Else
        While (Not objRs.EOF)
            sPtId = "" & objRs.Fields("ptid").Value
            sOrdDt = "" & objRs.Fields("orddt").Value
            sOrdNo = "" & objRs.Fields("ordno").Value
            sOrdSeq = "" & objRs.Fields("ordseq").Value
            
            pOrdInfo = pOrdInfo & sPtId & COL_DIV & sOrdDt & COL_DIV & sOrdNo & COL_DIV & sOrdSeq & LINE_DIV
            
            Set objRs1 = Nothing
            Set objRs1 = New Recordset
            objRs1.Open objMicSql.SqlOrdVerifyCheck(sPtId, sOrdDt, sOrdNo, sOrdSeq), DBConn
            If objRs1.EOF Then
                IsSensiAllVerified = True
            Else
                IsSensiAllVerified = False
            End If
            objRs.MoveNext
        Wend
    End If
    Set objRs = Nothing
    Set objRs1 = Nothing
End Function


Public Function SaveGRResult(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                            ByVal pTestCd As String, ByVal pResult As String, ByVal pStatus As String, _
                            ByVal pEmpId As String, ByVal pFNote As String, ByVal pRmkCd As String, _
                            Optional ByVal pMfySeq As Long = 0, Optional ByVal pRstType As String = "") As Boolean
    Dim sqlUpdate() As String
    Dim pDate       As String
    Dim pTime       As String
    Dim strSql      As String
    Dim strOrdInfo  As String
    Dim strTmp      As String
    Dim sPtId       As String
    Dim sOrdDt      As String
    Dim sOrdNo      As String
    Dim sOrdSeq     As String
    Dim blnVerified As Boolean
    Dim i           As Long
    
    ' 시스템 일자/시간 설정
    pDate = Format(GetSystemDate, CS_DateDbFormat)
    pTime = Format(GetSystemDate, CS_TimeDbFormat)

    
    ReDim sqlUpdate(0)
    
    '*** 감수성결과 수정 시 미생물 결과 History 작업
    If pMfySeq > 0 Then Call SetHistory(pWorkArea, pAccDt, pAccSeq, pTestCd)
    '***


    strSql = objMicSql.SQLVerifyResult(pWorkArea, pAccDt, pAccSeq, pTestCd, _
                                         pResult, pStatus, pDate, pTime, pEmpId, MRT_SenRst, pMfySeq, pRstType)
    
    DBConn.Execute strSql
 
    blnVerified = IsSensiAllVerified(pWorkArea, pAccDt, pAccSeq, pTestCd, strOrdInfo)

    strTmp = medShift(strOrdInfo, LINE_DIV)
    While Trim(strTmp) <> ""
        sPtId = medGetP(strTmp, 1, COL_DIV)
        sOrdDt = medGetP(strTmp, 2, COL_DIV)
        sOrdNo = medGetP(strTmp, 3, COL_DIV)
        sOrdSeq = medGetP(strTmp, 4, COL_DIV)
        
        If blnVerified Then
            ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
            sqlUpdate(UBound(sqlUpdate)) = objMicSql.SqlUpdateOrderStatus(sPtId, sOrdDt, sOrdNo, sOrdSeq, _
                                                pStatus, pDate, pTime, pEmpId)
        Else
            ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
            sqlUpdate(UBound(sqlUpdate)) = objMicSql.SqlUpdateOrderStatus(sPtId, sOrdDt, sOrdNo, sOrdSeq, _
                                                enStsCd.StsCd_LIS_MidRst, pDate, pTime, pEmpId)
        End If
        strTmp = medShift(strOrdInfo, LINE_DIV)
    Wend
    
    If pStatus >= enStsCd.StsCd_LIS_FinRst Then  'And fMicFg = "" Then               ' 최종결과이면 Final Flag
        ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLUpdateWoksheet(3, pWorkArea, pAccDt, pAccSeq, "", "", MWS_Final)
        
    Else                                        ' 중간결과이면 Growth Flag
        ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLUpdateWoksheet(3, pWorkArea, pAccDt, pAccSeq, "", "", MWS_Growth)
        
    End If

    ' *** Parent Status 반영 (접수내역, 미생물 WS)
    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
    sqlUpdate(UBound(sqlUpdate)) = SavePStatus(pWorkArea, pAccDt, pAccSeq, 1, pStatus, _
                                               pDate, pTime, pEmpId)

    ' *** Foot Note 저장 *** 검체 Remark 저장
    If pMfySeq = 0 Then
        ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLSaveFootNote(pWorkArea, pAccDt, pAccSeq, 0, "", "", 0)
    End If
    
    If Trim(pFNote) <> "" Then
        ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
'        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLSaveFootNote(pWorkArea, pAccDt, pAccSeq, 1, pEmpId, pFNote, 1)
'Modify By Legends 2003/09/23
'결과수정시 footnote를 추가적으로 기입할 경우 duplicate에러가 발생해서 수정했음.

        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLSaveFootNote(pWorkArea, pAccDt, pAccSeq, pMfySeq + 1, pEmpId, pFNote, 1)
    End If
'수정전 원본 2003/10/08
'    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
'    sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLSaveRemark(pWorkArea, pAccDt, pAccSeq, 1, pRmkCd)
    
'Modify By Legends 2003/10/08
'Footnote가 등록되지 않는데도 불구하고 접수내역의 Footnotefg='1'로 바꿀필요는 없쥐

    If Trim(pFNote) <> "" Then
        ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
        sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLSaveRemark(pWorkArea, pAccDt, pAccSeq, 1, pRmkCd)
    End If
    
On Error GoTo DBExecError
    For i = 1 To UBound(sqlUpdate)
        If Trim(sqlUpdate(i)) <> "" Then DBConn.Execute sqlUpdate(i)
    Next
    SaveGRResult = True

    Exit Function

DBExecError:
   
    
    SaveGRResult = False
   
End Function

Public Function SaveSenResult(ByVal ssSusc As Object, ByVal pWorkArea As String, ByVal pAccDt As String, _
                              ByVal pAccSeq As String, ByVal pTestCd As String, Optional ByVal pMfySeq = "0") As Boolean
    Dim sMicCd      As String
    Dim sSenType    As String
    Dim sQtyCd      As String
    
    Dim sInsHead    As String
    Dim sInsData    As String
    Dim sSen1       As String
    Dim sSen2       As String
    Dim sSen3       As String
    Dim sSenRst     As String
    Dim sqlDelete   As String
    Dim sqlInsert   As String
    Dim sqlUpdate() As String

    Dim i           As Integer
    Dim j           As Integer
    Dim iAntiCnt    As Integer


    ' 기존 감수성 결과를 삭제 (부분입력을 위해 갱신 할 수 있도록)  -- 수정 내역은 무조건 없슴
    
    ReDim sqlUpdate(0)
    

    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
    sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLDelSensiResult(pWorkArea, pAccDt, pAccSeq, pTestCd)


    ' 새로이 감수성 결과 삽입
    With ssSusc
    
        For i = 1 To .MaxRows Step 3
    
            .Col = 2: .Row = i + 1: sMicCd = .Text
            .Col = 4: .Row = i + 1: sSenType = .Text
            .Col = 6: .Row = i + 1: sQtyCd = .Text
            .Col = 7: .Row = i + 1: iAntiCnt = Val(.Value)
    
            ' 공백라인인 생기면 이후 입력된 결과는 무시한다. (Exit For --> Exit Sub)
            If sMicCd = "" Or sQtyCd = "" Then Exit For
    
            sInsHead = "(workarea,accdt,accseq,testcd,mfyseq,seq,mnmcd,micfg,mqtcd,scnt"
    
            sInsData = "(" & DBV("workarea", pWorkArea) & "," & DBV("accdt", pAccDt) & "," & _
                             DBV("accseq", pAccSeq) & "," & DBV("testcd", pTestCd) & "," & _
                             DBV("mfyseq", pMfySeq) & "," & DBV("seq", ((i + 2) / 3)) & "," & _
                             DBV("mnmcd", sMicCd) & "," & DBV("micfg", sSenType) & "," & _
                             DBV("mqtcd", sQtyCd) & "," & DBV("scnt", iAntiCnt)
    
            If iAntiCnt = 0 Then
                sInsHead = sInsHead & ")"
                sInsData = sInsData & ")"
            Else
                sInsHead = sInsHead & ","
                sInsData = sInsData & ","
            End If
    
    
                        
            Dim MnmWarn As String
            Dim QtyWarn As String
            Dim SenWarn As String
            Dim MedWarn As String
            
            For j = (7 + 1) To (iAntiCnt + 7)
    
                .Col = j: .Row = i: sSen1 = .Text
                .Col = j: .Row = i + 1: sSen2 = .Text
                .Col = j: .Row = i + 2: sSen3 = .Text
                
                .Col = 1: .Row = i + 1
                MnmWarn = IIf(.FontItalic, "M", "")
                
                .Col = 5: .Row = i + 1
                QtyWarn = IIf(.FontItalic, "Q", "")
                
                .Col = j: .Row = i
                SenWarn = IIf(.ForeColor = vbRed, "S", "")
                MedWarn = IIf(.FontItalic, "M", "")
                
                '항생제;감수성;농도;감수성 Warning여부(S);농도 Warning여부(M);미생물(M);정도(Q)
                sSenRst = sSen1 & ";" & sSen2 & ";" & sSen3 & ";" & SenWarn & ";" & MedWarn & ";" & MnmWarn & ";" & QtyWarn
    
                If j = (iAntiCnt + 7) Then
                    sInsHead = sInsHead & "srst" & (j - 7) & ")"
                    sInsData = sInsData & DBV("srst" & (j - 7), sSenRst) & ")"
                Else
                    sInsHead = sInsHead & "srst" & (j - 7) & ","
                    sInsData = sInsData & DBV("srst" & (j - 7), sSenRst) & ","
                End If
                
                
            
            Next j
                
    
            ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
            sqlUpdate(UBound(sqlUpdate)) = "INSERT INTO " & T_LAB405 & " " & sInsHead & " VALUES " & sInsData
            
            
            
        Next i
        
    End With
    
On Error GoTo DBExecError

    For i = 1 To UBound(sqlUpdate)
        DBConn.Execute sqlUpdate(i)
    Next
    SaveSenResult = True
    
    Exit Function

DBExecError:
    SaveSenResult = False
   
End Function

Public Function SaveSenResult_New(ByVal ssSusc As Object, ByVal pWorkArea As String, ByVal pAccDt As String, _
                                ByVal pAccSeq As String, ByVal pTestCd As String, ByVal pEmpId As String, _
                                ByVal pFNote As String, Optional ByVal pMfySeq = "0") As Boolean
    Dim sMicCd      As String
    Dim sSenType    As String
    Dim sQtyCd      As String
    
    Dim sInsHead    As String
    Dim sInsData    As String
    Dim sSen1       As String
    Dim sSen2       As String
    Dim sSen3       As String
    Dim sSenRst     As String
    Dim sqlDelete   As String
    Dim sqlInsert   As String
    Dim sqlUpdate() As String

    Dim i           As Integer
    Dim j           As Integer
    Dim iAntiCnt    As Integer

    '-- 전주예수병원 추가 변수=============
    Dim pDate As String
    Dim pTime As String
    Dim strOrdInfo As String, strTmp As String
    Dim strPtId As String, strOrdDt As String, strOrdNo As String
    Dim blnVerified As Boolean
    Dim SqlTransfer() As String
    Dim intTempSql As Integer
    Dim clsTransfer As New clsLISTransfer
    Dim strVfyDt As String
    Dim strRstVal As String
    Dim strSenVal As String
    Dim strSEQ    As String
    Dim strIPAdr  As String
    Dim strWrkDiv As String
    Dim bFlag     As Boolean
    Dim bINUPFlag As Boolean
    '======================================
    
    ' 기존 감수성 결과를 삭제 (부분입력을 위해 갱신 할 수 있도록)  -- 수정 내역은 무조건 없슴
    
    ReDim sqlUpdate(0)
    
    ' 시스템 일자/시간 설정
    pDate = Format(GetSystemDate, CS_DateDbFormat)
    pTime = Format(GetSystemDate, CS_TimeDbFormat)
    
    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
    sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLDelSensiResult(pWorkArea, pAccDt, pAccSeq, pTestCd)

    '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
    blnVerified = IsSensiAllVerified(pWorkArea, pAccDt, pAccSeq, pTestCd, strOrdInfo)
    
    strTmp = medShift(strOrdInfo, LINE_DIV)
    
    strPtId = medGetP(strTmp, 1, COL_DIV)
    strOrdDt = medGetP(strTmp, 2, COL_DIV)
    strOrdNo = medGetP(strTmp, 3, COL_DIV)
    '=================================================================================
    
    ' 새로이 감수성 결과 삽입
    With ssSusc
    
        For i = 1 To .MaxRows Step 3
    
            .Col = 2: .Row = i + 1: sMicCd = .Text
            .Col = 4: .Row = i + 1: sSenType = .Text
            .Col = 6: .Row = i + 1: sQtyCd = .Text
            .Col = 7: .Row = i + 1: iAntiCnt = Val(.Value)
    
            ' 공백라인인 생기면 이후 입력된 결과는 무시한다. (Exit For --> Exit Sub)
            If sMicCd = "" Or sQtyCd = "" Then Exit For
    
            sInsHead = "(workarea,accdt,accseq,testcd,mfyseq,seq,mnmcd,micfg,mqtcd,scnt"
    
            sInsData = "(" & DBV("workarea", pWorkArea) & "," & DBV("accdt", pAccDt) & "," & _
                             DBV("accseq", pAccSeq) & "," & DBV("testcd", pTestCd) & "," & _
                             DBV("mfyseq", pMfySeq) & "," & DBV("seq", ((i + 2) / 3)) & "," & _
                             DBV("mnmcd", sMicCd) & "," & DBV("micfg", sSenType) & "," & _
                             DBV("mqtcd", sQtyCd) & "," & DBV("scnt", iAntiCnt)
    
            If iAntiCnt = 0 Then
                sInsHead = sInsHead & ")"
                sInsData = sInsData & ")"
            Else
                sInsHead = sInsHead & ","
                sInsData = sInsData & ","
            End If
    
            '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
            ' * OCS측 결과는 균명을 넣어야 하기 때문에 균명과 정도를 가져온다.
            If bFlag = False Then
                strSenVal = clsTransfer.MicroNm_Find(sMicCd, sQtyCd)
                bFlag = True
            Else
                strSenVal = strSenVal & Chr(13) & clsTransfer.MicroNm_Find(sMicCd, sQtyCd)
            End If
            '=====================================================================================
                        
            Dim MnmWarn As String
            Dim QtyWarn As String
            Dim SenWarn As String
            Dim MedWarn As String
            
            For j = (7 + 1) To (iAntiCnt + 7)
    
                .Col = j: .Row = i: sSen1 = .Text
                .Col = j: .Row = i + 1: sSen2 = .Text
                .Col = j: .Row = i + 2: sSen3 = .Text
                
                .Col = 1: .Row = i + 1
                MnmWarn = IIf(.FontItalic, "M", "")
                
                .Col = 5: .Row = i + 1
                QtyWarn = IIf(.FontItalic, "Q", "")
                
                .Col = j: .Row = i
                SenWarn = IIf(.ForeColor = vbRed, "S", "")
                MedWarn = IIf(.FontItalic, "M", "")
                
                '항생제;감수성;농도;감수성 Warning여부(S);농도 Warning여부(M);미생물(M);정도(Q)
                sSenRst = sSen1 & ";" & sSen2 & ";" & sSen3 & ";" & SenWarn & ";" & MedWarn & ";" & MnmWarn & ";" & QtyWarn
    
                If j = (iAntiCnt + 7) Then
                    sInsHead = sInsHead & "srst" & (j - 7) & ")"
                    sInsData = sInsData & DBV("srst" & (j - 7), sSenRst) & ")"
                Else
                    sInsHead = sInsHead & "srst" & (j - 7) & ","
                    sInsData = sInsData & DBV("srst" & (j - 7), sSenRst) & ","
                End If
                
                '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
                ' * 적용항생제명과 결과정도를 넣어준다.
                strSenVal = strSenVal & Chr(13) & clsTransfer.SenNm_Find(sSen1) & " : " & sSen2 & " " & sSen3
                '=====================================================================================
            
            Next j
                
    
            ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
            sqlUpdate(UBound(sqlUpdate)) = "INSERT INTO " & T_LAB405 & " " & sInsHead & " VALUES " & sInsData
            
            
            
        Next i
        
    End With
    
On Error GoTo DBExecError
    
    '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
    
    'If pStatus >= enStsCd.StsCd_LIS_FinRst Then
        
'    End If
    
    
    strVfyDt = pDate & pTime
    
    '** 검사항목별 결과값 체크
    bINUPFlag = clsTransfer.SqlOCSRESULT_FLAG(strPtId, strOrdDt, strOrdNo, pTestCd)
    
    If bINUPFlag = False Then
        
        '-- 감수성 결과등록
        intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
        
        '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
        strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(strPtId, strOrdDt, strOrdNo)
        If strWrkDiv = "3" Then
            If clsTransfer.SqlOCSSU2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                        "N", "", pFNote, "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "")
            Else
                GoTo DBExecError
            End If
        ElseIf strWrkDiv = "4" Then
                If clsTransfer.SqlOCSSG2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                            "N", "", pFNote, "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "")
                Else
                    GoTo DBExecError
                End If
        Else
            If clsTransfer.SqlOCSMDEXMORT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                        "N", strVfyDt, pEmpId, "", "")
            Else
                GoTo DBExecError
            End If
        End If
        
        '** 종건/일건 이외의 검사만 등록.
        If strWrkDiv <> "" Then
            GoTo OCSSKIP
        End If
        
        intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
        
        SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(strPtId, strOrdDt, strOrdNo, _
                                pTestCd, clsTransfer.SlipCd, "", strSenVal, "", "", "", _
                                pFNote, "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "T", "", "", "")
    Else
        '-- 감수성 결과수정
        
        
        '-- 전주예수병원IP Address =============
        strIPAdr = clsTransfer.LocalIP_Address
        '=======================================
        
        '-- OCS 결과전송 시 필요한 정보 가져오기 ================
        'Call clsTransfer.SqlOCS_Info(strPtId, strOrdDt, strOrdNo, pTestCd)
        '========================================================
        
        intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
        
        '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
        'strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(strPtId, strOrdDt, strOrdNo)
        If clsTransfer.WrkDiv = "3" Then
            If clsTransfer.SqlOCSSU2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                        "N", "", pFNote, "", clsTransfer.VfyDt & clsTransfer.VfyTm, _
                                        clsTransfer.VfyId, clsTransfer.VfyDt & clsTransfer.VfyTm, _
                                        clsTransfer.VfyId, pEmpId, strIPAdr, pEmpId)
            Else
                GoTo DBExecError
            End If
        ElseIf clsTransfer.WrkDiv = "4" Then
                If clsTransfer.SqlOCSSG2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                            "N", "", pFNote, "", clsTransfer.VfyDt & clsTransfer.VfyTm, _
                                            clsTransfer.VfyId, clsTransfer.VfyDt & clsTransfer.VfyTm, _
                                            clsTransfer.VfyId, pEmpId, strIPAdr, pEmpId)
                Else
                    GoTo DBExecError
                End If
        Else
            If clsTransfer.SqlOCSMDEXMORT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                        "N", clsTransfer.VfyDt & clsTransfer.VfyTm, clsTransfer.VfyId, _
                                        pEmpId, strIPAdr)
            Else
                GoTo DBExecError
            End If
        End If
        
        '** 종건/일건 이외의 검사만 등록.
        If strWrkDiv <> "" Then
            GoTo OCSSKIP
        End If
        
        strSEQ = clsTransfer.SqlOCSRESULT_HISTORY(strPtId, strOrdDt, strOrdNo, pTestCd)
        
        If strSEQ = "" Then
            strSEQ = "1"
        Else
            strSEQ = CInt(strSEQ) + 1
        End If
        
        intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
        
        '-- OCS 결과테이블 업데이트
        SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                    pTestCd, clsTransfer.SlipCd, "", strSenVal, "", "", "", _
                    pFNote, "", "", "", "", "", "", pEmpId, strIPAdr, strVfyDt)
        
        intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
        
        '-- OCS 결과History테이블 INSERT
        SqlTransfer(intTempSql) = clsTransfer.SqlOCSRSLTHT_INSERT(strPtId, strOrdDt, strOrdNo, _
                    pTestCd, strSEQ, clsTransfer.SlipCd, "", clsTransfer.RstText, "", "", "", _
                    clsTransfer.FootNote, "", clsTransfer.VfyDt & clsTransfer.VfyTm, _
                    clsTransfer.VfyId, clsTransfer.VfyDt & clsTransfer.VfyTm, _
                    clsTransfer.VfyId, strVfyDt, clsTransfer.RsltType, pEmpId, strIPAdr, strVfyDt)
                    
    End If
    
    '/////////////////////////////////////////////////////////////////////////////////////////
    
OCSSKIP:
    
    For i = 1 To UBound(sqlUpdate)
        DBConn.Execute sqlUpdate(i)
    Next
    
    '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
    For intTempSql = LBound(SqlTransfer) To UBound(SqlTransfer)
        intTempSql = intTempSql
        
        If Trim(SqlTransfer(intTempSql)) <> "" Then DBConn.Execute (SqlTransfer(intTempSql))
        
    Next intTempSql
    '============================================================================================
    
    SaveSenResult_New = True
    
    Exit Function

DBExecError:
    SaveSenResult_New = False
   
End Function

Private Function SavePStatus(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, _
                             ByVal pVCnt As Integer, ByVal pStatus As String, ByVal pDate As String, _
                             ByVal pTime As String, ByVal pEmpId As String, Optional ByVal pCurStatus As String) As String
    Dim dsAcc           As Recordset
    Dim sqlAcc          As String
    Dim sStatus         As String
    Dim upStatus        As String
    Dim sqlUpdAcc       As String
    Dim upInputCount    As Integer
    Dim iTotalCount     As Integer
    Dim iInputCount     As Integer
    
    ' 중간결과를 사용하지 않는 검사인 경우에 성립
    ' 만약 감수성 검사 등의 경우라면 중간결과시 카운터에는 반영않지만 Status는 반영함

'    If pStatus < pCurStatus Then Exit Function

    ' 검사종류 확인
    Set dsAcc = New Recordset
    dsAcc.Open objMicSql.SqlGetTestCount(pWorkArea, pAccDt, pAccSeq), DBConn

    iTotalCount = Val("" & dsAcc.Fields("reqtotcnt").Value)
    iInputCount = Val("" & dsAcc.Fields("reqinputcnt").Value)
    sStatus = enStsCd.StsCd_LIS_MidRst  '"" & dsAcc.Fields("stscd").Value

    

    If iTotalCount <= iInputCount + pVCnt Then sStatus = pStatus                    ' 먼저 체크
    If pStatus = enStsCd.StsCd_LIS_FinRst Then iInputCount = iInputCount + pVCnt    ' 다음 체크 (순서 중요)

    SavePStatus = objMicSql.SQLUpdateStatus(pWorkArea, pAccDt, pAccSeq, sStatus, iInputCount, pDate, pTime, pEmpId)
    Set dsAcc = Nothing
End Function


'----------------------------------------------------------------------------------
'3. 미생물 감수성 결과수정 관련 메쏘드
'----------------------------------------------------------------------------------


Private Sub SetHistory(ByVal pWorkArea As String, ByVal pAccDt As String, ByVal pAccSeq As String, ByVal pTestCd As String)
    Dim dsCurRst    As Recordset
    Dim sCurRst     As String
    Dim sqlDel407   As String
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
    Dim sTestCd     As String
    
    
    ' 현재 변경 대상 Data Read
    '-----------------
    'Nogrowth결과수정할때 pTestCd에 검사항목이 여러개 넘어올때가 있어서 바꿈
    'Modify by legends 2003/07/30
    Set dsCurRst = New Recordset
    If InStr(pTestCd, "'") > 0 Then
        dsCurRst.Open objMicSql.SQLGetCurrentRst(pWorkArea, pAccDt, pAccSeq, pTestCd), DBConn
    Else
        dsCurRst.Open objMicSql.SQLGetCurrentRst(pWorkArea, pAccDt, pAccSeq, "'" & pTestCd & "'"), DBConn
    End If
    '------------------
    
'    Set dsCurRst = OpenRecordSet(objMicSql.SQLGetCurrentRst(pWorkArea, pAccDt, pAccSeq, pTestCd))

    ' 혹시 데이타가 언매치되면 접수내역을 기준으로 정리..
    ' 하지만 그럴일이 생길수 없으므로 사실 이 루틴은 필요 없슴
    ' 일단 막아놓고 나중에 필요 할 경우에 풀어줄 것.
    'sqlDel407 = "DELETE lab407 " & _
    '          " WHERE workarea='" & txtWorkArea & "' AND accdt='" & txtWorkArea & "' AND accseq='" & txtAccSeq & "'" & _
    '          " AND testcd='" & fTestCd & "' AND mfyseq>='" & "" & dsCurRst.Fields("mfyseq").Value & "'"
    'OraDB.ExecuteSQL (sqldel407)
    
    dsCurRst.MoveFirst
    Do Until dsCurRst.EOF
        sMfySeq = "" & dsCurRst.Fields("mfyseq").Value
        sRstCd = "" & dsCurRst.Fields("RstCd").Value
        sSenFg = "" & dsCurRst.Fields("SenFg").Value
        sRstDiv = "" & dsCurRst.Fields("RstDiv").Value
        sDetailFg = "" & dsCurRst.Fields("DetailFg").Value
        sVfyDt = "" & dsCurRst.Fields("VfyDt").Value
        sVfyTm = "" & dsCurRst.Fields("VfyTm").Value
        sVfyId = "" & dsCurRst.Fields("VfyId").Value
        sTestCd = "" & dsCurRst.Fields("testcd").Value
        
        sqlIns407 = objMicSql.SQLInsertOldData(pWorkArea, pAccDt, pAccSeq, sTestCd, sMfySeq, sRstCd, _
                                           sSenFg, sRstDiv, sDetailFg, sVfyDt, sVfyTm, sVfyId)
        DBConn.Execute sqlIns407
        
        dsCurRst.MoveNext
    Loop
    
    Set dsCurRst = Nothing
    
End Sub

Public Sub GetFinRstList(ByVal strMidVfyDt As String, ByRef objList As Object)
    
    Dim strSql  As String
    Dim objRs   As Recordset
    
    strSql = objMicSql.SqlGetFinList(strMidVfyDt)
    Set objRs = New Recordset
    objRs.Open strSql, DBConn
    
    objList.Clear
    While (Not objRs.EOF)
        objList.AddItem "" & objRs.Fields("workarea").Value & "-" & _
                        Mid(objRs.Fields("accdt").Value, 3) & "-" & _
                        Format(objRs.Fields("accseq").Value, "@@@@") & vbTab & _
                        objRs.Fields("ptnm").Value & vbTab & _
                        objRs.Fields("ptid").Value
        objRs.MoveNext
    Wend
    Set objRs = Nothing

End Sub

'** 사용자 요구사항에 따른 변경 By M.G.Choi 2006.04.05
'   감수성 최종결과보고 대상 조회 조건을 From ~ To 로 변경함
Public Sub GetFinRstList_New(ByVal strMidFVfyDt As String, _
                            ByVal strMidTVfyDt As String, _
                            ByRef objList As Object)
    
    Dim strSql  As String
    Dim objRs   As Recordset
    
    strSql = objMicSql.SqlGetFinList_New(strMidFVfyDt, strMidTVfyDt)
    
    Set objRs = New Recordset
    objRs.Open strSql, DBConn
    
    objList.Clear
    While (Not objRs.EOF)
        objList.AddItem "" & objRs.Fields("workarea").Value & "-" & _
                        Mid(objRs.Fields("accdt").Value, 3) & "-" & _
                        Format(objRs.Fields("accseq").Value, "@@@@") & vbTab & _
                        objRs.Fields("ptnm").Value & vbTab & _
                        objRs.Fields("ptid").Value
        objRs.MoveNext
    Wend
    Set objRs = Nothing

End Sub

'** 사용자 요구사항에 따른 변경 By 양성현 2008.12.17
'   Stain 최종결과보고 대상 조회 조건을 From ~ To 로 변경함
Public Sub GetFinRstList_Gra(ByVal strMidFVfyDt As String, ByVal strMidTVfyDt As String, ByRef objList As Object)
    
    Dim strSql  As String
    Dim objRs   As Recordset
    
    strSql = objMicSql.SqlGetFinList_Gra(strMidFVfyDt, strMidTVfyDt)
    
    Set objRs = New Recordset
    objRs.Open strSql, DBConn
    
    objList.Clear
    While (Not objRs.EOF)
        objList.AddItem "" & objRs.Fields("workarea").Value & "-" & _
                        Mid(objRs.Fields("accdt").Value, 3) & "-" & _
                        Format(objRs.Fields("accseq").Value, "@@@@") & vbTab & _
                        objRs.Fields("ptnm").Value & vbTab & _
                        objRs.Fields("ptid").Value
        objRs.MoveNext
    Wend
    Set objRs = Nothing

End Sub


Public Function GetLatestSensiAccNo(ByVal pPtId As String, ByVal pVfyDt As String, ByVal pSpcCd As String) As String
    Dim objRs       As Recordset
    Dim strSql      As String
    Dim blnAccNo    As Boolean
    
    blnAccNo = False
    
    strSql = objMicSql.SqlSensiAccNo(pPtId, pVfyDt, pSpcCd)
    Set objRs = New Recordset
    objRs.Open strSql, DBConn
    
    If Not objRs.EOF Then
        GetLatestSensiAccNo = "" & objRs.Fields("workarea").Value & "-" & _
                                   objRs.Fields("accdt").Value & "-" & _
                                   objRs.Fields("accseq").Value & "-" & _
                                   objRs.Fields("vfydt").Value & "-" & _
                                   objRs.Fields("vfytm").Value
    End If
    
    Set objRs = Nothing
    
End Function


Public Function SaveSenResult_New2(ByVal ssSusc As Object, ByVal pWorkArea As String, ByVal pAccDt As String, _
                                   ByVal pAccSeq As String, ByVal pTestCd As String, ByVal pEmpId As String, _
                                   ByVal pFNote As String, Optional ByVal pStatus As String = "", _
                                   Optional ByVal pMfySeq = "0") As Boolean
    Dim sMicCd      As String
    Dim sSenType    As String
    Dim sQtyCd      As String
    
    Dim sInsHead    As String
    Dim sInsData    As String
    Dim sSen1       As String
    Dim sSen2       As String
    Dim sSen3       As String
    Dim sSenRst     As String
    Dim sqlDelete   As String
    Dim sqlInsert   As String
    Dim sqlUpdate() As String

    Dim i           As Integer
    Dim j           As Integer
    Dim iAntiCnt    As Integer

    '-- 전주예수병원 추가 변수=============
    Dim pDate As String
    Dim pTime As String
    Dim strOrdInfo As String, strTmp As String
    Dim strPtId As String, strOrdDt As String, strOrdNo As String
    Dim blnVerified As Boolean
    Dim SqlTransfer() As String
    Dim intTempSql As Integer
    Dim clsTransfer As New clsLISTransfer
    Dim strVfyDt As String
    Dim strRstVal As String
    Dim strSenVal As String
    Dim strSEQ    As String
    Dim strIPAdr  As String
    Dim strWrkDiv As String
    Dim bFlag     As Boolean
    Dim bINUPFlag As Boolean
    '======================================
    
    ' 기존 감수성 결과를 삭제 (부분입력을 위해 갱신 할 수 있도록)  -- 수정 내역은 무조건 없슴
    
    ReDim sqlUpdate(0)
    
    ' 시스템 일자/시간 설정
    pDate = Format(GetSystemDate, CS_DateDbFormat)
    pTime = Format(GetSystemDate, CS_TimeDbFormat)
    
    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
    sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLDelSensiResult(pWorkArea, pAccDt, pAccSeq, pTestCd)

    '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
    blnVerified = IsSensiAllVerified(pWorkArea, pAccDt, pAccSeq, pTestCd, strOrdInfo)
    
    strTmp = medShift(strOrdInfo, LINE_DIV)
    
    strPtId = medGetP(strTmp, 1, COL_DIV)
    strOrdDt = medGetP(strTmp, 2, COL_DIV)
    strOrdNo = medGetP(strTmp, 3, COL_DIV)
    '=================================================================================
    
    ' 새로이 감수성 결과 삽입
    With ssSusc
    
        For i = 1 To .MaxRows Step 3
    
            .Col = 2: .Row = i + 1: sMicCd = .Text
            .Col = 4: .Row = i + 1: sSenType = .Text
            .Col = 6: .Row = i + 1: sQtyCd = .Text
            .Col = 7: .Row = i + 1: iAntiCnt = Val(.Value)
    
            ' 공백라인인 생기면 이후 입력된 결과는 무시한다. (Exit For --> Exit Sub)
            If sMicCd = "" Or sQtyCd = "" Then Exit For
    
            sInsHead = "(workarea,accdt,accseq,testcd,mfyseq,seq,mnmcd,micfg,mqtcd,scnt"
    
            sInsData = "(" & DBV("workarea", pWorkArea) & "," & DBV("accdt", pAccDt) & "," & _
                             DBV("accseq", pAccSeq) & "," & DBV("testcd", pTestCd) & "," & _
                             DBV("mfyseq", pMfySeq) & "," & DBV("seq", ((i + 2) / 3)) & "," & _
                             DBV("mnmcd", sMicCd) & "," & DBV("micfg", sSenType) & "," & _
                             DBV("mqtcd", sQtyCd) & "," & DBV("scnt", iAntiCnt)
    
            If iAntiCnt = 0 Then
                sInsHead = sInsHead & ")"
                sInsData = sInsData & ")"
            Else
                sInsHead = sInsHead & ","
                sInsData = sInsData & ","
            End If
    
            '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
            ' * OCS측 결과는 균명을 넣어야 하기 때문에 균명과 정도를 가져온다.
            If bFlag = False Then
                strSenVal = clsTransfer.MicroNm_Find(sMicCd, sQtyCd)
                bFlag = True
            Else
                strSenVal = strSenVal & Chr(13) & clsTransfer.MicroNm_Find(sMicCd, sQtyCd)
            End If
            '=====================================================================================
                        
            Dim MnmWarn As String
            Dim QtyWarn As String
            Dim SenWarn As String
            Dim MedWarn As String
            
            For j = (7 + 1) To (iAntiCnt + 7)
    
                .Col = j: .Row = i: sSen1 = .Text
                .Col = j: .Row = i + 1: sSen2 = .Text
                .Col = j: .Row = i + 2: sSen3 = .Text
                
                .Col = 1: .Row = i + 1
                MnmWarn = IIf(.FontItalic, "M", "")
                
                .Col = 5: .Row = i + 1
                QtyWarn = IIf(.FontItalic, "Q", "")
                
                .Col = j: .Row = i
                SenWarn = IIf(.ForeColor = vbRed, "S", "")
                MedWarn = IIf(.FontItalic, "M", "")
                
                '항생제;감수성;농도;감수성 Warning여부(S);농도 Warning여부(M);미생물(M);정도(Q)
                sSenRst = sSen1 & ";" & sSen2 & ";" & sSen3 & ";" & SenWarn & ";" & MedWarn & ";" & MnmWarn & ";" & QtyWarn
    
                If j = (iAntiCnt + 7) Then
                    sInsHead = sInsHead & "srst" & (j - 7) & ")"
                    sInsData = sInsData & DBV("srst" & (j - 7), sSenRst) & ")"
                Else
                    sInsHead = sInsHead & "srst" & (j - 7) & ","
                    sInsData = sInsData & DBV("srst" & (j - 7), sSenRst) & ","
                End If
                
                '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
                ' * 적용항생제명과 결과정도를 넣어준다.
                strSenVal = strSenVal & Chr(13) & clsTransfer.SenNm_Find(sSen1) & " : " & sSen2 & " " & sSen3
                '=====================================================================================
            
            Next j
                
    
            ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
            sqlUpdate(UBound(sqlUpdate)) = "INSERT INTO " & T_LAB405 & " " & sInsHead & " VALUES " & sInsData
            
            
            
        Next i
        
    End With
    
On Error GoTo DBExecError
    
    '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
    strVfyDt = pDate & pTime
    
    '-- OCS 결과전송 시 필요한 정보 가져오기 ================
    'Call clsTransfer.SqlOCS_Info(strPtId, strOrdDt, strOrdNo, pTestCd)
    '========================================================
    
    Select Case pStatus
        Case enStsCd.StsCd_LIS_FinRst '-- 결과등록
            '-- 감수성 결과등록
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
            strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(strPtId, strOrdDt, strOrdNo)
            If strWrkDiv = "3" Then
                If clsTransfer.SqlOCSSU2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                            "N", "", pFNote, "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "")
                Else
                    GoTo DBExecError
                End If
            ElseIf strWrkDiv = "4" Then
                    If clsTransfer.SqlOCSSG2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                                "N", "", pFNote, "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "")
                    Else
                        GoTo DBExecError
                    End If
            Else
                If clsTransfer.SqlOCSMDEXMORT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                            "N", strVfyDt, pEmpId, "", "")
                Else
                    GoTo DBExecError
                End If
            End If
            
            '** 종건/일건 이외의 검사만 등록.
            If strWrkDiv <> "" Then
                GoTo OCSSKIP
            End If
        
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(strPtId, strOrdDt, strOrdNo, _
                                    pTestCd, clsTransfer.SlipCd, "", strSenVal, "", "", "", _
                                    pFNote, clsTransfer.RcvDt & clsTransfer.RCVTM, strVfyDt, pEmpId, strVfyDt, pEmpId, "", "T", "", "", "")
            
        
        Case enStsCd.StsCd_LIS_Modify '-- 결과수정
            '-- 전주예수병원IP Address =============
            strIPAdr = clsTransfer.LocalIP_Address
            '=======================================
            
            '-- 전결과 정보 =======================================================
            Call clsTransfer.SqlOCSRESULT_OLD(strPtId, strOrdDt, strOrdNo, pTestCd)
            '======================================================================
    
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
    
            '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
            strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(strPtId, strOrdDt, strOrdNo)
            If strWrkDiv = "3" Then
                If clsTransfer.SqlOCSSU2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_MODIFY_SEN(strPtId, strOrdDt, strOrdNo, _
                                            "N", "", pFNote, pEmpId, strIPAdr, strVfyDt)
                Else
                    GoTo DBExecError
                End If
            ElseIf strWrkDiv = "4" Then
                    If clsTransfer.SqlOCSSG2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_MODIFY_SEN(strPtId, strOrdDt, strOrdNo, _
                                                "N", "", pFNote, pEmpId, strIPAdr, strVfyDt)
                    Else
                        GoTo DBExecError
                    End If
            Else
                If clsTransfer.SqlOCSMDEXMORT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE_NEW(strPtId, strOrdDt, strOrdNo, _
                                            "N", clsTransfer.VfyDt & clsTransfer.VfyTm, clsTransfer.VfyId, _
                                            pEmpId, strIPAdr, strVfyDt)
                Else
                    GoTo DBExecError
                End If
            End If
            
            '** 종건/일건 이외의 검사만 등록.
            If strWrkDiv <> "" Then
                GoTo OCSSKIP
            End If
        
            strSEQ = clsTransfer.SqlOCSRESULT_HISTORY(strPtId, strOrdDt, strOrdNo, pTestCd)
            
            If strSEQ = "" Then
                strSEQ = "1"
            Else
                strSEQ = CInt(strSEQ) + 1
            End If
            
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            '-- OCS 결과테이블 업데이트
            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_MODIFY(strPtId, strOrdDt, strOrdNo, _
                        pTestCd, "", strSenVal, pFNote, strVfyDt, pEmpId, strIPAdr, strVfyDt)
            
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            '-- OCS 결과History테이블 INSERT
            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRSLTHT_INSERT(strPtId, strOrdDt, strOrdNo, _
                        pTestCd, strSEQ, clsTransfer.SlipCd, "", clsTransfer.RstText, "", "", "", _
                        clsTransfer.FootNote, clsTransfer.RcvDt & clsTransfer.RCVTM, clsTransfer.VfyDt & clsTransfer.VfyTm, _
                        clsTransfer.VfyId, clsTransfer.VfyDt & clsTransfer.VfyTm, _
                        clsTransfer.VfyId, strVfyDt, clsTransfer.RsltType, pEmpId, strIPAdr, strVfyDt)
            
    End Select
    '/////////////////////////////////////////////////////////////////////////////////////////
    
OCSSKIP:
    
    For i = 1 To UBound(sqlUpdate)
        DBConn.Execute sqlUpdate(i)
    Next
    
    '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
    For intTempSql = LBound(SqlTransfer) To UBound(SqlTransfer)
        intTempSql = intTempSql
        
        If Trim(SqlTransfer(intTempSql)) <> "" Then DBConn.Execute (SqlTransfer(intTempSql))
        
    Next intTempSql
    '============================================================================================
    
    SaveSenResult_New2 = True
    
    Exit Function

DBExecError:
    SaveSenResult_New2 = False
   
End Function


Public Function SaveSenResult_New3(ByVal ssSusc As Object, ByVal pWorkArea As String, ByVal pAccDt As String, _
                                   ByVal pAccSeq As String, ByVal pTestCd As String, ByVal pEmpId As String, _
                                   ByVal pFNote As String, ByVal pRmkCd As String, Optional ByVal pStatus As String = "", _
                                   Optional ByVal pMfySeq = "0") As Boolean
    Dim sMicCd      As String
    Dim sSenType    As String
    Dim sQtyCd      As String
    
    Dim sInsHead    As String
    Dim sInsData    As String
    Dim sSen1       As String
    Dim sSen2       As String
    Dim sSen3       As String
    Dim sSenRst     As String
    Dim sqlDelete   As String
    Dim sqlInsert   As String
    Dim sqlUpdate() As String

    Dim i           As Integer
    Dim j           As Integer
    Dim iAntiCnt    As Integer

    '-- 전주예수병원 추가 변수=============
    Dim pDate As String
    Dim pTime As String
    Dim strOrdInfo As String, strTmp As String
    Dim strPtId As String, strOrdDt As String, strOrdNo As String
    Dim blnVerified     As Boolean
    Dim SqlTransfer()   As String
    Dim intTempSql      As Integer
    Dim clsTransfer     As New clsLISTransfer
    Dim strVfyDt        As String
    Dim strRstVal       As String
    Dim strSenVal       As String
    Dim strSEQ          As String
    Dim strIPAdr        As String
    Dim strWrkDiv       As String
    Dim strOCSNote      As String
    Dim strFromRef      As String
    Dim strToRef        As String
    Dim strRemark       As String
    Dim bFlag           As Boolean
    Dim bINUPFlag       As Boolean
    '======================================
    
    ' 기존 감수성 결과를 삭제 (부분입력을 위해 갱신 할 수 있도록)  -- 수정 내역은 무조건 없슴
    
    ReDim sqlUpdate(0)
    
    ' 시스템 일자/시간 설정
    pDate = Format(GetSystemDate, CS_DateDbFormat)
    pTime = Format(GetSystemDate, CS_TimeDbFormat)
    
    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
    sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLDelSensiResult(pWorkArea, pAccDt, pAccSeq, pTestCd)

    '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
    blnVerified = IsSensiAllVerified(pWorkArea, pAccDt, pAccSeq, pTestCd, strOrdInfo)
    
    strTmp = medShift(strOrdInfo, LINE_DIV)
    
    strPtId = medGetP(strTmp, 1, COL_DIV)
    strOrdDt = medGetP(strTmp, 2, COL_DIV)
    strOrdNo = medGetP(strTmp, 3, COL_DIV)
    '=================================================================================
    
    ' 새로이 감수성 결과 삽입
    With ssSusc
    
        For i = 1 To .MaxRows Step 3
    
            .Col = 2: .Row = i + 1: sMicCd = .Text
            .Col = 4: .Row = i + 1: sSenType = .Text
            .Col = 6: .Row = i + 1: sQtyCd = .Text
            .Col = 7: .Row = i + 1: iAntiCnt = Val(.Value)
    
            ' 공백라인인 생기면 이후 입력된 결과는 무시한다. (Exit For --> Exit Sub)
            If sMicCd = "" Or sQtyCd = "" Then Exit For
    
            sInsHead = "(workarea,accdt,accseq,testcd,mfyseq,seq,mnmcd,micfg,mqtcd,scnt"
    
            sInsData = "(" & DBV("workarea", pWorkArea) & "," & DBV("accdt", pAccDt) & "," & _
                             DBV("accseq", pAccSeq) & "," & DBV("testcd", pTestCd) & "," & _
                             DBV("mfyseq", pMfySeq) & "," & DBV("seq", ((i + 2) / 3)) & "," & _
                             DBV("mnmcd", sMicCd) & "," & DBV("micfg", sSenType) & "," & _
                             DBV("mqtcd", sQtyCd) & "," & DBV("scnt", iAntiCnt)
    
            If iAntiCnt = 0 Then
                sInsHead = sInsHead & ")"
                sInsData = sInsData & ")"
            Else
                sInsHead = sInsHead & ","
                sInsData = sInsData & ","
            End If
    
            '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
            ' * OCS측 결과는 균명을 넣어야 하기 때문에 균명과 정도를 가져온다.
            If bFlag = False Then
                strSenVal = clsTransfer.MicroNm_Find(sMicCd, sQtyCd)
                bFlag = True
            Else
                strSenVal = strSenVal & Chr(13) & clsTransfer.MicroNm_Find(sMicCd, sQtyCd)
            End If
            '=====================================================================================
                        
            Dim MnmWarn As String
            Dim QtyWarn As String
            Dim SenWarn As String
            Dim MedWarn As String
            
            For j = (7 + 1) To (iAntiCnt + 7)
    
                .Col = j: .Row = i: sSen1 = .Text
                .Col = j: .Row = i + 1: sSen2 = .Text
                .Col = j: .Row = i + 2: sSen3 = .Text
                
                .Col = 1: .Row = i + 1
                MnmWarn = IIf(.FontItalic, "M", "")
                
                .Col = 5: .Row = i + 1
                QtyWarn = IIf(.FontItalic, "Q", "")
                
                .Col = j: .Row = i
                SenWarn = IIf(.ForeColor = vbRed, "S", "")
                MedWarn = IIf(.FontItalic, "M", "")
                
                '항생제;감수성;농도;감수성 Warning여부(S);농도 Warning여부(M);미생물(M);정도(Q)
                sSenRst = sSen1 & ";" & sSen2 & ";" & sSen3 & ";" & SenWarn & ";" & MedWarn & ";" & MnmWarn & ";" & QtyWarn
    
                If j = (iAntiCnt + 7) Then
                    sInsHead = sInsHead & "srst" & (j - 7) & ")"
                    sInsData = sInsData & DBV("srst" & (j - 7), sSenRst) & ")"
                Else
                    sInsHead = sInsHead & "srst" & (j - 7) & ","
                    sInsData = sInsData & DBV("srst" & (j - 7), sSenRst) & ","
                End If
                
                '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
                ' * 적용항생제명과 결과정도를 넣어준다.
                strSenVal = strSenVal & Chr(13) & clsTransfer.SenNm_Find(sSen1) & " : " & sSen2 & " " & sSen3
                '=====================================================================================
            
            Next j
                
    
            ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
            sqlUpdate(UBound(sqlUpdate)) = "INSERT INTO " & T_LAB405 & " " & sInsHead & " VALUES " & sInsData
            
            
            
        Next i
        
    End With
    
On Error GoTo DBExecError
    
    '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
    strVfyDt = pDate & pTime
    
    strOCSNote = ""
    
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
                        
    '-- OCS 결과전송 시 필요한 정보 가져오기 ================
    'Call clsTransfer.SqlOCS_Info(strPtId, strOrdDt, strOrdNo, pTestCd)
    '========================================================
    
    Select Case pStatus
        Case enStsCd.StsCd_LIS_MidRst, enStsCd.StsCd_LIS_FinRst   '-- 결과등록
            '-- 감수성 결과등록
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
            strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(strPtId, strOrdDt, strOrdNo)
            If strWrkDiv = "3" Then
                If clsTransfer.SqlOCSSU2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                            "N", "", strOCSNote, "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "")
                Else
                    GoTo DBExecError
                End If
            ElseIf strWrkDiv = "4" Then
                    If clsTransfer.SqlOCSSG2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                                "N", "", strOCSNote, "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "")
                    Else
                        GoTo DBExecError
                    End If
            Else
                If clsTransfer.SqlOCSMDEXMORT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                            "N", strVfyDt, pEmpId, "", "")
                Else
                    GoTo DBExecError
                End If
            End If
            
            '** 종건/일건 이외의 검사만 등록.
            If strWrkDiv <> "" Then
                GoTo OCSSKIP
            End If
        
            Call clsTransfer.SqlOCS_Info_Mic(pWorkArea, pAccDt, pAccSeq)
            
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            '** 중간결과 처리로 인한 변경
            If clsTransfer.SqlOCSRESULT_FLAG(strPtId, strOrdDt, strOrdNo, pTestCd) = False Then
                SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(strPtId, strOrdDt, strOrdNo, _
                                        pTestCd, clsTransfer.SlipCd, "", strSenVal, "", "", "", _
                                        strOCSNote, clsTransfer.RcvDt & clsTransfer.RCVTM, strVfyDt, _
                                        pEmpId, strVfyDt, pEmpId, "", "T", "", "", "")
            Else
                SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_SUB_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                        pTestCd, "", strSenVal, strOCSNote, strVfyDt, pEmpId)
            End If
            
OCSSKIP:
            
            '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
            For intTempSql = LBound(SqlTransfer) To UBound(SqlTransfer)
                intTempSql = intTempSql
                
                If Trim(SqlTransfer(intTempSql)) <> "" Then DBConn.Execute (SqlTransfer(intTempSql))
                
            Next intTempSql
            '============================================================================================
        
        Case enStsCd.StsCd_LIS_Modify '-- 결과수정
            '-- 전주예수병원IP Address =============
            strIPAdr = clsTransfer.LocalIP_Address
            '=======================================
            
            '-- 전결과 정보 =======================================================
            Call clsTransfer.SqlOCSRESULT_OLD(strPtId, strOrdDt, strOrdNo, pTestCd)
            '======================================================================
    
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
    
            '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
            strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(strPtId, strOrdDt, strOrdNo)
            If strWrkDiv = "3" Then
                If clsTransfer.SqlOCSSU2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_MODIFY_SEN(strPtId, strOrdDt, strOrdNo, _
                                            "N", "", strOCSNote, pEmpId, strIPAdr, strVfyDt)
                                            
                    DBConn.Execute (SqlTransfer(intTempSql))
                Else
                    GoTo DBExecError
                End If
            ElseIf strWrkDiv = "4" Then
                    If clsTransfer.SqlOCSSG2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_MODIFY_SEN(strPtId, strOrdDt, strOrdNo, _
                                                "N", "", strOCSNote, pEmpId, strIPAdr, strVfyDt)
                                                
                        DBConn.Execute (SqlTransfer(intTempSql))
                    Else
                        GoTo DBExecError
                    End If
            Else
                If clsTransfer.SqlOCSMDEXMORT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE_NEW(strPtId, strOrdDt, strOrdNo, _
                                            "N", clsTransfer.VfyDt & clsTransfer.VfyTm, clsTransfer.VfyId, _
                                            pEmpId, strIPAdr, strVfyDt)
                Else
                    GoTo DBExecError
                End If
            End If
            
            '** 종건/일건 이외의 검사만 등록.
            If strWrkDiv <> "" Then
                GoTo OCSSKIP
            End If
        
            strSEQ = clsTransfer.SqlOCSRESULT_HISTORY(strPtId, strOrdDt, strOrdNo, pTestCd)
            
            If strSEQ = "" Then
                strSEQ = "1"
            Else
                strSEQ = CInt(strSEQ) + 1
            End If
            
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            '-- OCS 결과테이블 업데이트
            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_MODIFY(strPtId, strOrdDt, strOrdNo, _
                        pTestCd, "", strSenVal, strOCSNote, strVfyDt, pEmpId, strIPAdr, strVfyDt)
            
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            '-- OCS 결과History테이블 INSERT
            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRSLTHT_INSERT(strPtId, strOrdDt, strOrdNo, _
                        pTestCd, strSEQ, clsTransfer.SlipCd, "", clsTransfer.RstText, "", "", "", _
                        clsTransfer.FootNote, clsTransfer.RcvDt & clsTransfer.RCVTM, clsTransfer.VfyDt & clsTransfer.VfyTm, _
                        clsTransfer.VfyId, clsTransfer.VfyDt & clsTransfer.VfyTm, _
                        clsTransfer.VfyId, strVfyDt, clsTransfer.RsltType, pEmpId, strIPAdr, strVfyDt)
                        
            '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
            For intTempSql = LBound(SqlTransfer) To UBound(SqlTransfer)
                intTempSql = intTempSql
                
                If Trim(SqlTransfer(intTempSql)) <> "" Then DBConn.Execute (SqlTransfer(intTempSql))
                
            Next intTempSql
    '============================================================================================
            
    End Select
    '/////////////////////////////////////////////////////////////////////////////////////////
    
    For i = 1 To UBound(sqlUpdate)
        DBConn.Execute sqlUpdate(i)
    Next
    
    SaveSenResult_New3 = True
    
    Exit Function

DBExecError:
    SaveSenResult_New3 = False
   
End Function


'** 검체명 추가로 4차 수정함 By M.G.Choi 2004.12.08
Public Function SaveSenResult_New4(ByVal ssSusc As Object, ByVal pWorkArea As String, ByVal pAccDt As String, _
                                   ByVal pAccSeq As String, ByVal pTestCd As String, ByVal pEmpId As String, _
                                   ByVal pFNote As String, ByVal pRmkCd As String, Optional ByVal pStatus As String = "", _
                                   Optional ByVal pMfySeq = "0", Optional ByVal pSpcNm As String = "") As Boolean
    Dim sMicCd      As String
    Dim sSenType    As String
    Dim sQtyCd      As String
    Dim sInsHead    As String
    Dim sInsData    As String
    Dim sSen1       As String
    Dim sSen2       As String
    Dim sSen3       As String
    Dim sSenRst     As String
    Dim sqlDelete   As String
    Dim sqlInsert   As String
    Dim sqlUpdate() As String

    Dim i           As Integer
    Dim j           As Integer
    Dim iAntiCnt    As Integer

    '-- 전주예수병원 추가 변수=============
    Dim pDate As String
    Dim pTime As String
    Dim strOrdInfo As String, strTmp As String
    Dim strPtId As String, strOrdDt As String, strOrdNo As String
    Dim blnVerified     As Boolean
    Dim SqlTransfer()   As String
    Dim intTempSql      As Integer
    Dim clsTransfer     As New clsLISTransfer
    Dim strVfyDt        As String
    Dim strRstVal       As String
    Dim strSenVal       As String
    Dim strSEQ          As String
    Dim strIPAdr        As String
    Dim strWrkDiv       As String
    Dim strOCSNote      As String
    Dim strFromRef      As String
    Dim strToRef        As String
    Dim strRemark       As String
    Dim strSpcInfo      As String
    Dim strSql          As String
    Dim strAccNo        As String
    Dim strMicInfo      As String
    Dim bFlag           As Boolean
    Dim bINUPFlag       As Boolean
    '======================================
    
    ' 기존 감수성 결과를 삭제 (부분입력을 위해 갱신 할 수 있도록)  -- 수정 내역은 무조건 없슴
    
    ReDim sqlUpdate(0)
    
    ' 시스템 일자/시간 설정
    pDate = Format(GetSystemDate, CS_DateDbFormat)
    pTime = Format(GetSystemDate, CS_TimeDbFormat)
    
    ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
    sqlUpdate(UBound(sqlUpdate)) = objMicSql.SQLDelSensiResult(pWorkArea, pAccDt, pAccSeq, pTestCd)

    '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
    blnVerified = IsSensiAllVerified(pWorkArea, pAccDt, pAccSeq, pTestCd, strOrdInfo)
    
    strTmp = medShift(strOrdInfo, LINE_DIV)
    
    strPtId = medGetP(strTmp, 1, COL_DIV)
    strOrdDt = medGetP(strTmp, 2, COL_DIV)
    strOrdNo = medGetP(strTmp, 3, COL_DIV)
    
    strAccNo = pWorkArea & Mid(pAccDt, 3, 4) & pAccSeq
    
    '-- 추가루틴 : 미생물반응균결과전송 By M.G.Choi 2005.03.02
    '   3. SLXSENST : 미생물감수성결과       (DELETE)
    '   4. SLXBACTT : 미생물반응균결과       (DELETE)
    strSql = clsTransfer.SqlSLXSENST_DELETE(strAccNo)
    Call DBConn.Execute(strSql)
    
    strSql = clsTransfer.SqlSLXBACTT_DELETE(strAccNo)
    Call DBConn.Execute(strSql)
    '=================================================================================
    
    ' 새로이 감수성 결과 삽입
    With ssSusc
    
        For i = 1 To .MaxRows Step 3
    
            .Col = 2: .Row = i + 1: sMicCd = .Text
            .Col = 4: .Row = i + 1: sSenType = .Text
            .Col = 6: .Row = i + 1: sQtyCd = .Text
            .Col = 7: .Row = i + 1: iAntiCnt = Val(.Value)
    
            ' 공백라인인 생기면 이후 입력된 결과는 무시한다. (Exit For --> Exit Sub)
            '## 5.0.11: 이상대(2005-05-14)
            '   - 균코드, 정도코드가 비어있으면 해당 라인만 스킵하도록 수정
            If sMicCd = "" Or sQtyCd = "" Then GoTo Skip
    
            sInsHead = "(workarea,accdt,accseq,testcd,mfyseq,seq,mnmcd,micfg,mqtcd,scnt"
    
            sInsData = "(" & DBV("workarea", pWorkArea) & "," & DBV("accdt", pAccDt) & "," & _
                             DBV("accseq", pAccSeq) & "," & DBV("testcd", pTestCd) & "," & _
                             DBV("mfyseq", pMfySeq) & "," & DBV("seq", ((i + 2) / 3)) & "," & _
                             DBV("mnmcd", sMicCd) & "," & DBV("micfg", sSenType) & "," & _
                             DBV("mqtcd", sQtyCd) & "," & DBV("scnt", iAntiCnt)
    
            If iAntiCnt = 0 Then
                sInsHead = sInsHead & ")"
                sInsData = sInsData & ")"
            Else
                sInsHead = sInsHead & ","
                sInsData = sInsData & ","
            End If
    
            '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
            ' * OCS측 결과는 균명을 넣어야 하기 때문에 균명과 정도를 가져온다.
            If bFlag = False Then
                strSenVal = clsTransfer.MicroNm_Find(sMicCd, sQtyCd)
                bFlag = True
            Else
                strSenVal = strSenVal & Chr(13) & clsTransfer.MicroNm_Find(sMicCd, sQtyCd)
            End If
            '=====================================================================================
                        
            Dim MnmWarn As String
            Dim QtyWarn As String
            Dim SenWarn As String
            Dim MedWarn As String
            
            For j = (7 + 1) To (iAntiCnt + 7)
    
                .Col = j: .Row = i: sSen1 = .Text
                .Col = j: .Row = i + 1: sSen2 = .Text
                .Col = j: .Row = i + 2: sSen3 = .Text
                
                .Col = 1: .Row = i + 1
                MnmWarn = IIf(.FontItalic, "M", "")
                
                .Col = 5: .Row = i + 1
                QtyWarn = IIf(.FontItalic, "Q", "")
                
                .Col = j: .Row = i
                SenWarn = IIf(.ForeColor = vbRed, "S", "")
                MedWarn = IIf(.FontItalic, "M", "")
                
                '항생제;감수성;농도;감수성 Warning여부(S);농도 Warning여부(M);미생물(M);정도(Q)
                sSenRst = sSen1 & ";" & sSen2 & ";" & sSen3 & ";" & SenWarn & ";" & MedWarn & ";" & MnmWarn & ";" & QtyWarn
    
                If j = (iAntiCnt + 7) Then
                    sInsHead = sInsHead & "srst" & (j - 7) & ")"
                    sInsData = sInsData & DBV("srst" & (j - 7), sSenRst) & ")"
                Else
                    sInsHead = sInsHead & "srst" & (j - 7) & ","
                    sInsData = sInsData & DBV("srst" & (j - 7), sSenRst) & ","
                End If
                
                '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
                ' * 적용항생제명과 결과정도를 넣어준다.
                strSenVal = strSenVal & Chr(13) & clsTransfer.SenNm_Find(sSen1) & " : " & sSen2 & " " & sSen3
                
                '-- 추가루틴 : 미생물반응균결과전송 By M.G.Choi 2005.03.02
                '   3. SLXSENST : 미생물감수성결과       (I)
'                If clsTransfer.SqlSLXSENST_FLAG(strAccNo, sMicCd, sSen1) = False Then
'                    strSQL = clsTransfer.SqlSLXSENST_INSERT(strAccNo, sMicCd, sSen1, sSen2, sSen3, sQtyCd)
'                Else
'                    strSQL = clsTransfer.SqlSLXSENST_UPDATE(strAccNo, sMicCd, sSen1, sSen2, sSen3, sQtyCd)
'                End If
                
                strSql = clsTransfer.SqlSLXSENST_INSERT(strAccNo, sMicCd, sSen1, sSen2, sSen3, sQtyCd)
                Call DBConn.Execute(strSql)
                '=====================================================================================
            
            Next j
                
    
            ReDim Preserve sqlUpdate(UBound(sqlUpdate) + 1)
            sqlUpdate(UBound(sqlUpdate)) = "INSERT INTO " & T_LAB405 & " " & sInsHead & " VALUES " & sInsData
            
            '-- 추가루틴 : 미생물반응균결과전송 By M.G.Choi 2005.03.02
            '   4. SLXBACTT : 미생물반응균결과       (I) => 미생물접수번호/반응균코드만 생성
'            If clsTransfer.SqlSLXSENST_FLAG(strAccNo, sMicCd, sSen1) = False Then
'                strSQL = clsTransfer.SqlSLXSENST_INSERT(strAccNo, sMicCd, sSen1, sSen2, sSen3, sQtyCd)
'            Else
'                strSQL = clsTransfer.SqlSLXSENST_UPDATE(strAccNo, sMicCd, sSen1, sSen2, sSen3, sQtyCd)
'            End If
            
            strSql = clsTransfer.SqlSLXBACTT_INSERT(strAccNo, sMicCd)
            Call DBConn.Execute(strSql)
            '=====================================================================================
Skip:
        Next i
    End With
    
On Error GoTo DBExecError
    
    '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
    strVfyDt = pDate & pTime
    
    strOCSNote = ""
    
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
                        
    '-- OCS 결과전송 시 필요한 정보 가져오기 ================
    'Call clsTransfer.SqlOCS_Info(strPtId, strOrdDt, strOrdNo, pTestCd)
    '========================================================
    
    '** 검체정보
    If pSpcNm <> "" Then
        strSpcInfo = "◈ Specimen Name : " & pSpcNm
        
        strSenVal = strSenVal & Chr(13) & strSpcInfo
    End If
    
    '** 추가 By M.G.Choi 2005.11.17
    '   중간결과/최종결과 구분 OCS 전달
    If pStatus = enStsCd.StsCd_LIS_MidRst Then
        strSenVal = "♣중간결과♣" & Chr(13) & strSenVal
    ElseIf pStatus = enStsCd.StsCd_LIS_FinRst Then
            strSenVal = "♣최종결과♣" & Chr(13) & strSenVal
    End If
    
    Select Case pStatus
        Case enStsCd.StsCd_LIS_MidRst, enStsCd.StsCd_LIS_FinRst   '-- 결과등록
            '-- 감수성 결과등록
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
            strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(strPtId, strOrdDt, strOrdNo)
            If strWrkDiv = "3" Then
                If clsTransfer.SqlOCSSU2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                            "N", "", strOCSNote, "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "")
                Else
                    GoTo DBExecError
                End If
            ElseIf strWrkDiv = "4" Then
                    If clsTransfer.SqlOCSSG2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                                "N", "", strOCSNote, "", strVfyDt, pEmpId, strVfyDt, pEmpId, "", "", "")
                    Else
                        GoTo DBExecError
                    End If
            Else
                If clsTransfer.SqlOCSMDEXMORT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                            "N", strVfyDt, pEmpId, "", "")
                Else
                    GoTo DBExecError
                End If
            End If
            
            '** 종건/일건 이외의 검사만 등록.
            If strWrkDiv <> "" Then
                GoTo OCSSKIP
            End If
        
            Call clsTransfer.SqlOCS_Info_Mic(pWorkArea, pAccDt, pAccSeq)
            
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            '** 중간결과 처리로 인한 변경
            If clsTransfer.SqlOCSRESULT_FLAG(strPtId, strOrdDt, strOrdNo, pTestCd) = False Then
                SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_INSERT(strPtId, strOrdDt, strOrdNo, _
                                        pTestCd, clsTransfer.SlipCd, "", strSenVal, "", "", "", _
                                        strOCSNote, clsTransfer.RcvDt & clsTransfer.RCVTM, strVfyDt, _
                                        pEmpId, strVfyDt, pEmpId, "", "T", "", "", "")
            Else
                SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_SUB_UPDATE(strPtId, strOrdDt, strOrdNo, _
                                        pTestCd, "", strSenVal, strOCSNote, strVfyDt, pEmpId)
            End If
            
OCSSKIP:
            '** 추가루틴 : 전주예수병원 반응균 결과전송 By M.G.Choi 2005.03.02
            ' * OCS TABLE
            '   1. SLMICWMT : 미생물검사결과(MASTER) (I/U)
            '   2. SLMICWDT : 미생물검사결과(DETAIL) (I/U)
            '-- OCS 결과전송 시 필요한 정보 가져오기 ================
            ' - 검체코드/외래입원구분/진료과/병동/접수일시/보고일시
            strMicInfo = clsTransfer.SqlOCS_MICInfo(pWorkArea, pAccDt, pAccSeq)
            '========================================================
            
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            If clsTransfer.SqlSLMICWMT_FLAG(strAccNo) = False Then
                SqlTransfer(intTempSql) = clsTransfer.SqlSLMICWMT_INSERT(strAccNo, medGetP(strMicInfo, 1, vbTab), strPtId, _
                                        medGetP(strMicInfo, 2, vbTab), "", strOrdDt, medGetP(strMicInfo, 3, vbTab), _
                                        medGetP(strMicInfo, 4, vbTab), medGetP(strMicInfo, 5, vbTab), _
                                        medGetP(strMicInfo, 5, vbTab), strVfyDt)
            Else
                SqlTransfer(intTempSql) = clsTransfer.SqlSLMICWMT_UPDATE(strAccNo, medGetP(strMicInfo, 1, vbTab), strPtId, _
                                        medGetP(strMicInfo, 2, vbTab), "", strOrdDt, medGetP(strMicInfo, 3, vbTab), _
                                        medGetP(strMicInfo, 4, vbTab), medGetP(strMicInfo, 5, vbTab), _
                                        medGetP(strMicInfo, 5, vbTab), strVfyDt)
            End If
            
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            If clsTransfer.SqlSLMICWDT_FLAG(strAccNo, pTestCd, strOrdNo) = False Then
                SqlTransfer(intTempSql) = clsTransfer.SqlSLMICWDT_INSERT(strAccNo, pTestCd, strOrdNo, strSenVal)
            Else
                SqlTransfer(intTempSql) = clsTransfer.SqlSLMICWDT_UPDATE(strAccNo, pTestCd, strOrdNo, strSenVal)
            End If
            '반응균전송 끝 ==================================================================
            
            '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
            For intTempSql = LBound(SqlTransfer) To UBound(SqlTransfer)
                intTempSql = intTempSql
                
                If Trim(SqlTransfer(intTempSql)) <> "" Then DBConn.Execute (SqlTransfer(intTempSql))
                
            Next intTempSql
            '============================================================================================
            
        Case enStsCd.StsCd_LIS_Modify '-- 결과수정
            '-- 전주예수병원IP Address =============
            strIPAdr = clsTransfer.LocalIP_Address
            '=======================================
            
            '-- 전결과 정보 =======================================================
            Call clsTransfer.SqlOCSRESULT_OLD(strPtId, strOrdDt, strOrdNo, pTestCd)
            '======================================================================
    
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
    
            '-- 1. '3' 종합건진, '4' 일반검사, 이외검사로 구분을 한다.
            strWrkDiv = clsTransfer.SqlOCBussdiv_Flag(strPtId, strOrdDt, strOrdNo)
            If strWrkDiv = "3" Then
                If clsTransfer.SqlOCSSU2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSSU2EXAMT_MODIFY_SEN(strPtId, strOrdDt, strOrdNo, _
                                            "N", "", strOCSNote, pEmpId, strIPAdr, strVfyDt)
                                            
                    DBConn.Execute (SqlTransfer(intTempSql))
                Else
                    GoTo DBExecError
                End If
            ElseIf strWrkDiv = "4" Then
                    If clsTransfer.SqlOCSSG2EXAMT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                        SqlTransfer(intTempSql) = clsTransfer.SqlOCSSG2EXAMT_MODIFY_SEN(strPtId, strOrdDt, strOrdNo, _
                                                "N", "", strOCSNote, pEmpId, strIPAdr, strVfyDt)
                                                
                        DBConn.Execute (SqlTransfer(intTempSql))
                    Else
                        GoTo DBExecError
                    End If
            Else
                If clsTransfer.SqlOCSMDEXMORT_Find(strPtId, strOrdDt, strOrdNo) = True Then
                    SqlTransfer(intTempSql) = clsTransfer.SqlOCSMDEXMORT_UPDATE_NEW(strPtId, strOrdDt, strOrdNo, _
                                            "N", strVfyDt, clsTransfer.VfyId, _
                                            pEmpId, strIPAdr, strVfyDt)
                Else
                    GoTo DBExecError
                End If
            End If
            
            '** 종건/일건 이외의 검사만 등록.
            If strWrkDiv <> "" Then
                GoTo OCSSKIP
            End If
        
            strSEQ = clsTransfer.SqlOCSRESULT_HISTORY(strPtId, strOrdDt, strOrdNo, pTestCd)
            
            If strSEQ = "" Then
                strSEQ = "1"
            Else
                strSEQ = CInt(strSEQ) + 1
            End If
            
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            '-- OCS 결과테이블 업데이트
            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRESULT_MODIFY(strPtId, strOrdDt, strOrdNo, _
                        pTestCd, "", strSenVal, strOCSNote, strVfyDt, pEmpId, strIPAdr, strVfyDt)
            
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            '-- OCS 결과History테이블 INSERT
            SqlTransfer(intTempSql) = clsTransfer.SqlOCSRSLTHT_INSERT(strPtId, strOrdDt, strOrdNo, _
                        pTestCd, strSEQ, clsTransfer.SlipCd, "", clsTransfer.RstText, "", "", "", _
                        clsTransfer.FootNote, clsTransfer.RcvDt & clsTransfer.RCVTM, clsTransfer.VfyDt & clsTransfer.VfyTm, _
                        clsTransfer.VfyId, clsTransfer.VfyDt & clsTransfer.VfyTm, _
                        clsTransfer.VfyId, strVfyDt, clsTransfer.RsltType, pEmpId, strIPAdr, strVfyDt)
                        
                        
            '** 추가루틴 : 전주예수병원 반응균 결과전송 By M.G.Choi 2005.03.02
            ' * OCS TABLE
            '   1. SLMICWMT : 미생물검사결과(MASTER) (I/U)
            '   2. SLMICWDT : 미생물검사결과(DETAIL) (I/U)
            '-- OCS 결과전송 시 필요한 정보 가져오기 ================
            ' - 검체코드/외래입원구분/진료과/병동/접수일시/보고일시
            strMicInfo = clsTransfer.SqlOCS_MICInfo(pWorkArea, pAccDt, pAccSeq)
            '========================================================
            
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            If clsTransfer.SqlSLMICWMT_FLAG(strAccNo) = False Then
                SqlTransfer(intTempSql) = clsTransfer.SqlSLMICWMT_INSERT(strAccNo, medGetP(strMicInfo, 1, vbTab), strPtId, _
                                        medGetP(strMicInfo, 2, vbTab), "", strOrdDt, medGetP(strMicInfo, 3, vbTab), _
                                        medGetP(strMicInfo, 4, vbTab), medGetP(strMicInfo, 5, vbTab), _
                                        medGetP(strMicInfo, 5, vbTab), clsTransfer.VfyDt & clsTransfer.VfyTm)
            Else
                SqlTransfer(intTempSql) = clsTransfer.SqlSLMICWMT_UPDATE(strAccNo, medGetP(strMicInfo, 1, vbTab), strPtId, _
                                        medGetP(strMicInfo, 2, vbTab), "", strOrdDt, medGetP(strMicInfo, 3, vbTab), _
                                        medGetP(strMicInfo, 4, vbTab), medGetP(strMicInfo, 5, vbTab), _
                                        medGetP(strMicInfo, 5, vbTab), clsTransfer.VfyDt & clsTransfer.VfyTm)
            End If
            
            intTempSql = intTempSql + 1: ReDim Preserve SqlTransfer(intTempSql)
            
            If clsTransfer.SqlSLMICWDT_FLAG(strAccNo, pTestCd, strOrdNo) = False Then
                SqlTransfer(intTempSql) = clsTransfer.SqlSLMICWDT_INSERT(strAccNo, pTestCd, strOrdNo, strSenVal)
            Else
                SqlTransfer(intTempSql) = clsTransfer.SqlSLMICWDT_UPDATE(strAccNo, pTestCd, strOrdNo, strSenVal)
            End If
            '반응균전송 끝 ==================================================================
    
            '** 추가루틴 : 전주예수병원 결과전송 By M.G.Choi 2004.09.06
            For intTempSql = LBound(SqlTransfer) To UBound(SqlTransfer)
                intTempSql = intTempSql
                
                If Trim(SqlTransfer(intTempSql)) <> "" Then DBConn.Execute (SqlTransfer(intTempSql))
                
            Next intTempSql
            '============================================================================================
            
            
    End Select
    '/////////////////////////////////////////////////////////////////////////////////////////
    
    For i = 1 To UBound(sqlUpdate)
        DBConn.Execute sqlUpdate(i)
    Next
    
    SaveSenResult_New4 = True
    
    Exit Function

DBExecError:
    SaveSenResult_New4 = False
   
End Function




Private Sub Class_Terminate()
    Set objMicSql = Nothing
    Set objMicLib = Nothing
End Sub
