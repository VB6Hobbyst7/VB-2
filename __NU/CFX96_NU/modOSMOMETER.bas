Attribute VB_Name = "modOSMOMETER"
Option Explicit

Public Sub Phase_Serial_OSMOMETER()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
'        Select Case Asc(BufChar)
'            Case 5      'ENQ
'                frmMain.comEqp.Output = ACK
'            Case 0      '--
'                RcvBuffer = ""
'
'            Case 2      'STX
'                RcvBuffer = ""
'
'            Case 13      'ETX
'                Call SerialRcvData_OSMOMETER
'
'                RcvBuffer = ""
'                frmMain.comEqp.Output = ACK
'            Case 10
'                RcvBuffer = ""
'
'            Case 4      'EOT
'                RcvBuffer = ""
'
'            Case Else
'                RcvBuffer = RcvBuffer & BufChar
'        End Select
    
        Select Case Asc(BufChar)
            Case 13     'CR
                If Trim(RcvBuffer) <> "" Then
                    Call SerialRcvData_OSMOMETER
                    
                    RcvBuffer = ""
                End If
                
            Case 10     'LF
            
            Case Else
                RcvBuffer = RcvBuffer & BufChar

        End Select
    
    
    Next i

End Sub


Public Sub SerialRcvData_OSMOMETER()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strBarno        As String   '수신한 바코드번호
    Dim strSeq          As String   '수신한 Sequence
    Dim strRackNo       As String   '수신한 Rack Or Disk No
    Dim strTubePos      As String   '수신한 Tube Position
    Dim strIntBase      As String   '수신한 장비기준 검사명
    Dim strMachResult   As String   '수신한 장비결과
    Dim strResult       As String   '수신한 결과(정성)
    Dim strIntResult    As String   '수신한 결과(정량)
    Dim varResult       As Variant
    Dim strQCResult     As String   '수신한 결과(QC)
    Dim strFlag         As String   '수신한 Abnormal Flag
    Dim strComm         As String   '수신한 Comment
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strTmp          As String
    Dim strOldBarno     As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    With frmMain
        strRcvBuf = RcvBuffer

        '-- 테스트용 -----------------
        If .fraCommTest.Visible = False Then
            Call SetSQLData("RCV", strRcvBuf, "A")
        End If
        '-- 테스트용 -----------------
        
        '실제 전송되는 결과값
        '    Osmometer Ready
        '
        '  11/16/2006  02:21 AM
        'Osmolality  293 mOsm
        
        '1. Recall Results
        '#30: 284 mOsm     [PREV]
        'ID NONE
        '#29: 280 mOsm     [PREV]
        '#28: 296 mOsm     [PREV]
        '#27: 639 mOsm     [PREV]
        '#26: 288 mOsm     [PREV]
        '#25: 291 mOsm     [PREV]
        '#24: 381 mOsm     [PREV]
        '#23: 302 mOsm     [PREV]
            
        
        If InStr(strRcvBuf, "Ready") > 0 Or InStr(strRcvBuf, "Recall Results") > 0 Then
            Exit Sub
        End If
            
        'Realtime
        If InStr(strRcvBuf, "Osmolality") > 0 Then
            strResult = Trim(Mid(strRcvBuf, 11, 5))
            
        'Recall Result
        ElseIf Left(strRcvBuf, 1) = "#" And InStr(strRcvBuf, "mOsm") > 0 Then
            strResult = Trim(mGetP(mGetP(Trim(strRcvBuf), 2, ":"), 1, "mOsm"))
        Else
            Exit Sub
        End If
        
        strIntBase = "OSMO"
        

                
        With mResult
            .BarNo = ""
            .SpcPos = ""
            .Seq = ""
            .RackNo = ""
            .TubePos = ""
            'If strOldBarno <> strBarno Then
                'strOldBarno = strBarno
                .RsltDate = Format(Now, "yyyymmddhhmmss")
                .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
            'End If
        End With
        
        'Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
        If gRow <= 0 Then
            Exit Sub
        End If
        
                    
        If strIntBase <> "" And strResult <> "" Then
            If gPatOrdCd <> "" Then
                SQL = ""
                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                SQL = SQL & "  FROM EQPMASTER" & vbCr
                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                
                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                    lsTestCode = Trim(RS_L.Fields("TESTCODE"))
                    lsTestName = Trim(RS_L.Fields("TESTNAME"))
                    lsSeqNo = Trim(RS_L.Fields("SEQNO"))

                    '-- 결과Row 추가
                    lsRstRow = .spdResult.DataRowCnt + 1
                    If .spdResult.MaxRows < lsRstRow Then
                        .spdResult.MaxRows = lsRstRow
                    End If

                    '소수점 처리, 결과 형태 처리
                    strMachResult = strResult
                    strResult = SetResult(strResult, strIntBase)
                    strJudge = SetJudge(strResult, strIntBase)
                    
                    '진행상태 표시("결과")
                    SetText .spdOrder, "결과", gRow, colSTATE

                    '결과값 표시
                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                        If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                            SetText .spdOrder, strResult, gRow, intCol
                            Exit For
                        End If
                    Next

                    '-- 결과 List
                    SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                    SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                    SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                    SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                    SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                    SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                    SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                    SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                    
                    '-- 로컬 저장
                    SetLocalDB gRow, lsRstRow, "1", ""
                    
                    strState = "R"
                    
                    '-- BIORAD QC 저장
                    If mResult.Kind = "QC" Then
                        strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                        
                        Call SendBioRadQC(strQCData)
                    End If
            
                    '-- 결과Count
                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                        SetText .spdOrder, "1", gRow, colRCNT
                    Else
                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                    End If
                    
                End If
            Else
                SQL = ""
                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                SQL = SQL & "  FROM EQPMASTER" & vbCr
                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                
                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                    lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                    lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                    lsSeqNo = Trim(RS_L.Fields("SEQNO"))

                    '-- 결과Row 추가
                    lsRstRow = .spdResult.DataRowCnt + 1
                    If .spdResult.MaxRows < lsRstRow Then
                        .spdResult.MaxRows = lsRstRow
                    End If

                    '소수점 처리, 결과 형태 처리
                    strMachResult = strResult
                    strResult = SetResult(strResult, strIntBase)
                    strJudge = SetJudge(strResult, strIntBase)
                    
                    '진행상태 표시("결과")
                    SetText .spdOrder, "결과", gRow, colSTATE

                    '결과값 표시
                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                        If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                            SetText .spdOrder, strResult, gRow, intCol
                            Exit For
                        End If
                    Next

                    '-- 결과 List
                    SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                    SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                    SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                    SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                    SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                    SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                    SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                    SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                    
                    '-- 로컬 저장
                    SetLocalDB gRow, lsRstRow, "1", ""
                    
                    If strState <> "R" Then
                        strState = ""
                    End If

                    '-- BIORAD QC 저장
                    If mResult.Kind = "QC" Then
                        strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                        
                        Call SendBioRadQC(strQCData)
                    End If
                    
                    '-- 결과Count
                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
                        SetText .spdOrder, "1", gRow, colRCNT
                    Else
                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                    End If
                End If
                
            End If
            
        End If
                    
        .spdResult.RowHeight(-1) = 14
                    
        '## DB에 결과저장
        If .optTrans(0).Value = True And strState = "R" Then
            Res = SaveTransData_MCC(gRow)
            
            If Res = -1 Then
                '-- 저장 실패
                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "Failed", gRow, colSTATE
            Else
                '-- 저장 성공
                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX
                
                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                
                If DBExec(AdoCn_Local, SQL) Then
                    '-- 성공
                End If
            End If
            strState = ""
        End If
    End With

End Sub


