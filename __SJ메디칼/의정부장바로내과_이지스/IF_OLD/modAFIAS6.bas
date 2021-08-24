Attribute VB_Name = "modAFIAS6"
Option Explicit

Public Sub Phase_Serial_AFIAS6()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case "$" 'SOH
                Erase strRecvData
                
                intBufCnt = 1
                ReDim Preserve strRecvData(intBufCnt)
            Case vbCr
                Call SerialRcvData_AFIAS6
            Case Else
                If intBufCnt > 0 Then
                    strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                End If
        End Select
    Next i

End Sub


Public Sub SerialRcvData_AFIAS6()
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
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strRcvBuf = strRecvData(intCnt)
            strBarno = Trim(mGetP(strRcvBuf, 5, "|"))
            strRackNo = ""
            strTubePos = ""
            strSeq = ""
                        
            With mResult
                .BarNo = strBarno
                .SpcPos = strSeq
                .Seq = strSeq
                .RackNo = strRackNo
                .TubePos = strTubePos
                .RsltDate = Format(Now, "yyyymmddhhmmss")
                .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
            End With
            
            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
            If gRow <= 0 Then
                Exit Sub
            End If
                        
            strIntBase = mGetP(strRcvBuf, 8, "|")
            strResult = mGetP(strRcvBuf, 11, "|")
                        
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
        Next
    End With

End Sub
