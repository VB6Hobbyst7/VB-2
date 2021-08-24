Attribute VB_Name = "modHITACHI7020"
Option Explicit


'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput   As String     '송신할 데이터
    
    strOutput = ";" & mOrder.Function
    strOutput = strOutput & " 37"
    strOutput = strOutput & Mid(mOrder.Order, 1, 37)
    strOutput = strOutput & "00000"
    
    'COMMENT란에 BARCODE 표시
    'strOutput = strOutput & "100000" & Left(mOrder.BarNo & Space(30), 30)
    
    Call Sleep(100)
    
    '-- SPE Send(오더전송)
    frmMain.comEqp.Output = STX & strOutput & ETX & vbCr & vbLf
    
    SetRawData "[Tx]" & STX & strOutput & ETX & vbCr & vbLf


End Sub


Public Sub Phase_Serial_HITACHI7020()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(Buffer, i, 1)
        Select Case BufChar
            Case STX
                RcvBuffer = ""
                
            Case ETX
                 Call SerialRcvData_HITACHI7020
                 RcvBuffer = ""
            
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
            
End Sub

Public Sub SndMore()
    Dim strSndMsg As String
    
    strSndMsg = ">"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) ' & GetChkSum(strSndMsg) & vbCr
    strSndMsg = strSndMsg & vbCrLf
    
    frmMain.comEqp.Output = strSndMsg
    
    SetRawData "[Tx]" & strSndMsg
    
End Sub

Public Sub SndRec()
    Dim strSndMsg As String
    
    strSndMsg = "A"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) '& GetChkSum(strSndMsg)
    strSndMsg = strSndMsg & vbCrLf
    
    frmMain.comEqp.Output = strSndMsg
    
    SetRawData "[Tx]" & strSndMsg
    
End Sub


Private Sub SerialRcvData_HITACHI7020()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strType         As String   '수신한 Record Type
    Dim strOldBarno        As String   '수신한 바코드번호
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
    Dim strAspect       As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
    Dim lsOrderCode     As String   '처방코드
    Dim lsTestCode      As String   '검사코드
    Dim lsTestName      As String   '검사명
    Dim lsSeqNo         As String   '로컬DB 검사Seq
    
    Dim lsRstRow        As String   '결과스프레드 현재 Row
    Dim intCnt          As Integer  '통신 Frame 갯수
    Dim intCol          As Integer  '결과컬럼 갯수
    Dim strJudge        As String   '결과판정
    Dim Res             As Integer
    
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    Dim strTmp          As String
    Dim i               As Integer
    Dim iBCpos          As Integer
    
    Dim iTBlockNo   As Integer
    Dim iCBlockNo   As Integer
    Dim iItemNo     As Integer
    Dim strKind     As String
    Dim iPos        As Integer
    
    Dim varIntBase()    As String
    Dim varResult()     As String
    Dim varFlag()       As String

    'for H7020
    Dim strFunc       As String
    Dim strFunction   As String
    Dim strSendData   As String
    Dim strSndMsg     As String
    Dim strExamCode() As String
    
    With frmMain
        strRcvBuf = RcvBuffer
        
        '-- 테스트용 -----------------
        If .fraCommTest.Visible = False Then
            Call SetSQLData("RCV", strRcvBuf, "A")
        End If
        '-- 테스트용 -----------------
        
        strType = Mid$(strRcvBuf, 1, 1)
        If IsNumeric(strType) Then
            strType = Mid$(strRcvBuf, 2, 1)
        End If
        
        Select Case strType
            Case ">", "?", "@"      'ANY 수신
                Call SndMore        'MOR Send
                Do
                '   DoEvents
                Loop Until frmMain.comEqp.OutBufferCount = 0
            
            Case "?", "@"           'REP 수신
                Sleep (100)
                Call SndMore        'MOR Send
                Do
                '   DoEvents
                Loop Until frmMain.comEqp.OutBufferCount = 0
            
            Case ">", "?", "@"      'SUS 수신
                Sleep (100)
                Call SndMore        'MOR Send
                Do
                '   DoEvents
                Loop Until frmMain.comEqp.OutBufferCount = 0
            
            Case ";"                'SPE  전송(오더요청)
                strFunction = Mid(strRcvBuf, 2, 12) & String(13, "#") & Mid(strRcvBuf, 27, 15)
            
                strFunc = Mid(strRcvBuf, 2, 1)              'N
                strSeq = Mid(strRcvBuf, 4, 5)               '    1
                strRackNo = Mid(strRcvBuf, 9, 1)            '
                strTubePos = Mid(strRcvBuf, 10, 3)          '  1
                strBarno = Trim(Mid(strRcvBuf, 14, 13))

                With mOrder
                    .NoOrder = False
                    .BarNo = strBarno
                    .Func = strFunc
                    .Function = strFunction
                End With
                
                Call GetOrder(strBarno, gHOSP.RSTTYPE)
                
                Call SendOrder
        
            ' FR1 to FR9 (검사항목 25개 이상일 경우 처리)
            Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
                strFunc = Mid(strRcvBuf, 2, 1)
                
                If strFunc = "K" Or strFunc = "L" Or strFunc = "G" Or strFunc = "H" Then
                    Sleep (100)
                    Call SndMore        'MOR Send
                    Do
                    '   DoEvents
                    Loop Until frmMain.comEqp.OutBufferCount = 0
                    Exit Sub
                End If
                            
                Call SndMore            'MOR Send
                
                If strFunc <> "@" And strFunc <> "M" Then
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Trim(Mid(strRcvBuf, 10, 3))
                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
                    gRow = 0
                                      
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Kind = strKind
                        .Rerun = ""
                        If strOldBarno <> strBarno Then
                            strOldBarno = strBarno
                            .RsltDate = Format(Now, "yyyymmddhhmmss")
                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    
                            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                            
                        End If
                    End With
                    
                    'strTmp = Mid$(strRcvBuf, 29)
                    strTmp = Mid$(strRcvBuf, 45)
    
                    For i = 44 To Len(strRcvBuf) Step 10
                        strIntBase = Trim(Mid(strRcvBuf, i, 3))
                        strIntBase = Format(strIntBase, "00")
                        strResult = Trim(Mid(strRcvBuf, i + 3, 6))
                        
                        If strIntBase <> "" And strResult <> "" Then
                            If gPatOrdCd <> "" Then
                                SQL = ""
                                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
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
                                    
                                    '-- BIORAD QC 저장
'                                    If mResult.Kind = "QC" Then
'                                        strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                        Call SendBioRadQC(strQCData)
'                                    End If
                                    
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
                                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                                SQL = SQL & "  FROM EQPMASTER" & vbCr
                                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                                
                                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                    lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                    lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                    lsSeqNo = Trim(RS_L.Fields("SEQNO"))
            
                                    strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                    strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                    strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                    strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                    strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                    strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                    strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                    strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
            
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
                                    
                                    '-- BIORAD QC 저장
'                                    If mResult.Kind = "QC" Then
'
'                                        strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                        Call SendBioRadQC(strQCData)
'
'                                    End If
                                    
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
                        strTmp = Mid$(strTmp, 12)
                    Next
                    
                    .spdResult.RowHeight(-1) = 14
                
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_EASYS(gRow)
                        
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
                End If
                
            ' 결과 END
            Case ":"
                ':N     3   3 3                            4  4    17   5    16  11   9.3  12  0.89 
                strFunc = Mid(strRcvBuf, 2, 1)
                
                If strFunc = "K" Or strFunc = "L" Or strFunc = "G" Or strFunc = "H" Then
                    Sleep (100)
                    Call SndMore        'MOR Send
                    Do
                    '   DoEvents
                    Loop Until frmMain.comEqp.OutBufferCount = 0
                    Exit Sub
                End If
                
                If strFunc = "K" Or strFunc = "L" Then
                    Call SndMore        'MOR Send
                    Exit Sub
                End If
                
                Call SndMore            'MOR Send
                
                
                If strFunc <> "@" And strFunc <> "M" Then
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Trim(Mid(strRcvBuf, 10, 3))
                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
                    
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Kind = strKind
                        .Rerun = ""
                        If strOldBarno <> strBarno Then
                            strOldBarno = strBarno
                            .RsltDate = Format(Now, "yyyymmddhhmmss")
                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    
                            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                            
                        End If
                    End With
                    
                    'strTmp = Mid$(strRcvBuf, 29)
                    strTmp = Mid$(strRcvBuf, 45)
    
                    For i = 44 To Len(strRcvBuf) Step 10
                        strIntBase = Trim(Mid(strRcvBuf, i, 3))
                        strIntBase = Format(strIntBase, "00")
                        strResult = Trim(Mid(strRcvBuf, i + 3, 6))
                        
                        If strIntBase <> "" And strResult <> "" Then
                            If gPatOrdCd <> "" Then
                                SQL = ""
                                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
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
                                    
                                    '-- BIORAD QC 저장
                                    If mResult.Kind = "QC" Then
                                        strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                        
                                        Call SendBioRadQC(strQCData)
                                    End If
                                    
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
                                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
                                SQL = SQL & "  FROM EQPMASTER" & vbCr
                                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
                                
                                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
                                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
                                    lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
                                    lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
                                    lsSeqNo = Trim(RS_L.Fields("SEQNO"))
            
                                    strQCLab = Trim(RS_L.Fields("QCLab") & "")
                                    strQCLot = Trim(RS_L.Fields("QCLot") & "")
                                    strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
                                    strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
                                    strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
                                    strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
                                    strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
                                    strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
            
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
                                    
                                    '-- BIORAD QC 저장
                                    If mResult.Kind = "QC" Then
                                        
                                        strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                        
                                        Call SendBioRadQC(strQCData)
                                        
                                    End If
                                    
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
                        strTmp = Mid$(strTmp, 12)
                    Next
                    
                    .spdResult.RowHeight(-1) = 14
                
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_EASYS(gRow)
                        
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
                End If
        End Select
    End With

End Sub



'-----------------------------------------------------------------------------'
'   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
'                                 2. 장비수신정보 화면표시,
'                                 3. 처방코드 가져오기,
'                                 4. (처방코드로)검사오더 만들기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Private Sub GetOrder(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    
    '-- 1. 접수정보 조회
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
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
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
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
    
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
            
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
        
        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
        strItems = GetEquipExamCode_HITACHI7020(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- 진행상태(Order) 표시
            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
        End If


        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub

'검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
'한 장비 번호에 검사코드가 1개이상 존재
Private Function GetEquipExamCode_HITACHI7020(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode()   As String
    Dim sSpecNo         As String
    Dim iRow            As Long
    Dim SpecNo          As String
    
    Dim strSendData As String
    Dim ii          As Integer
    Dim strTestNum  As String
    
            
    GetEquipExamCode_HITACHI7020 = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    strSendData = String$(88, "0")
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    Erase strExamCode
    mOrder.SendCnt = 0
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strTestNum = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strTestNum <> "" Then
                ReDim Preserve strExamCode(ii)
                strExamCode(ii) = strTestNum
                ii = ii + 1
                mOrder.SendCnt = mOrder.SendCnt + 1
            End If
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    If gPatOrdCd <> "" And ii > 0 Then
        For i = 0 To UBound(strExamCode)
            If strExamCode(i) <> "" Then
                If strExamCode(i) <> "99" Then
                    Mid(strSendData, strExamCode(i), 1) = "1"
                End If
            End If
        Next
    End If
    
    GetEquipExamCode_HITACHI7020 = strSendData
    
End Function




