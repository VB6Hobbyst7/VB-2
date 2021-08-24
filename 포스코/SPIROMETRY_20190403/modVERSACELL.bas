Attribute VB_Name = "modVERSACELL"
Option Explicit

''-----------------------------------------------------------------------------'
''   기능 : 오더정보 전송
''-----------------------------------------------------------------------------'
'Private Sub SendOrder()
'    Dim strOutput   As String     '송신할 데이터
'
''1H|\^&||||62 Flanders-Bartley Road^Flanders^NJ^07921||973-927-2828|N81|||P|1|20161209210918
''6B
''2P|1|03217192|||Jo^ Yu Jeong^^|||U
''65
''3O|1|03217192||^^^wrCRP\^^^AMYLAS|R||||||||||1
''48
''4L|1|N
''07
'
'    Select Case intSndPhase
'        Case 1  '## Header
'            strOutput = intFrameNo & "H|\^&||||62 Flanders-Bartley Road^Flanders^NJ^07921||973-927-2828|N81|||P|1|" & Format(Now, "yyyymmddhhmmss") & "|" & vbCr & ETX
'            intSndPhase = 2
'            intFrameNo = intFrameNo + 1
'
'        Case 2  '## Patient
'            strOutput = intFrameNo & "P|1|" & mOrder.BarNo & "|||" & frmMain.Han2Eng.HanToEng(mOrder.PName) & "||||" & vbCr & ETX
'            intSndPhase = 3
'            intFrameNo = intFrameNo + 1
'
'        Case 3  '## Order
'            If mOrder.NoOrder = True Then
'                '## 접수정보가 없을경우
'                strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||||||" & mOrder.SPCCD
'                strOutput = intFrameNo & strOutput & vbCr & ETX
'                intSndPhase = 4
'
'            Else
'                '## 최초 보낼때
'                If mOrder.IsSending = False Then
'                    strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||||||1"
'
'                    If Len(strOutput) > 230 Then
'                        mOrder.IsSending = True
'                        mOrder.Order = Mid$(strOutput, 231)
'                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'                        intSndPhase = 3
'                    Else
'                        strOutput = intFrameNo & strOutput & vbCr & ETX
'                        intSndPhase = 4
'                    End If
'                '## 남은 문자열이 있을때
'                Else
'                    strOutput = mOrder.Order
'                    If Len(strOutput) > 230 Then
'                        mOrder.Order = Mid$(strOutput, 231)
'                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'                        intSndPhase = 3
'                    Else
'                        mOrder.IsSending = False
'                        strOutput = intFrameNo & strOutput & vbCr & ETX
'                        intSndPhase = 4
'                    End If
'                End If
'            End If
'            intFrameNo = intFrameNo + 1
'
'        Case 4  '## Termianator
'            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
'            intSndPhase = 5
'            intFrameNo = intFrameNo + 1
'
'        Case 5  '## EOT
'            strState = ""
'            frmMain.comEqp.Output = EOT
'            SetRawData "[Tx]" & EOT
'            intFrameNo = 1
'
'            Exit Sub
'    End Select
'
'    If intFrameNo = 8 Then
'        intFrameNo = 0
'    End If
'
'    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
'    frmMain.comEqp.Output = strOutput
'    SetRawData "[Tx]" & strOutput
'
'End Sub

'Public Sub Phase_Serial_VERSACELL()
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
'                        frmMain.comEqp.Output = ACK
'                        SetRawData "[Tx]" & ACK
'                    Case ACK
'                        If strState = "Q" Then
'                            Call SendOrder
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
'                        frmMain.comEqp.Output = ACK
'                        SetRawData "[Tx]" & ACK
'                End Select
'            Case 4      '## Termination Phase
'                Select Case BufChar
'                    Case STX
'                        intPhase = 2
'                    Case EOT
'                        Call SerialRcvData_VERSACELL
'                        If strState = "Q" Then
'                            intSndPhase = 1
'                            intFrameNo = 1
'                            frmMain.comEqp.Output = ENQ
'                            SetRawData "[Tx]" & ENQ
'                        End If
'                        intPhase = 1
'                End Select
'        End Select
'    Next i
'
'End Sub


'Private Sub SerialRcvData_VERSACELL()
'    Dim RS_L            As ADODB.Recordset
'    Dim strRcvBuf       As String   '수신한 Data
'    Dim strType         As String   '수신한 Record Type
'    'Dim strOldBarno        As String   '수신한 바코드번호
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
'    Dim strAspect       As String
'    Dim strEqpNm        As String
'
'    Dim strTemp1        As String
'    Dim strTemp2        As String
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
'    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
'    Dim strINTRResult   As String
'
'    '##################################
'    '##  1. 요소감소율 [URR]
'    '##  공식 : URR = 1 - ( PostBun [3730N1] / PreBUN [C3730N2] )
'    '##################################
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
'            strType = Mid$(strRcvBuf, 2, 1)
'            If strType = "|" Then
'                strType = Mid$(strRcvBuf, 1, 1)
'            End If
'
'            Select Case strType
'                Case "H"    '## Header
'                Case "P"    '## Patient
'                Case "Q"    '## Request Information
'                    If mGetP(strRcvBuf, 13, "|") = "A" Then Exit Sub
'                    strTemp1 = mGetP(strRcvBuf, 3, "|")
'                    strBarno = Trim$(mGetP(strTemp1, 2, "^"))
'
'                    With mOrder
'                        .NoOrder = False
'                        .BarNo = strBarno
'                        .Seq = mGetP(strTemp1, 3, "^")
'                        .RackNo = mGetP(strTemp1, 4, "^")
'                        .TubePos = mGetP(strTemp1, 5, "^")
'                    End With
'
'                    Call GetOrder(strBarno, gHOSP.RSTTYPE)
'                    strState = "Q"
'
'                Case "O"
'                    '3O|1|03498081||^^^FT4  |R||||||||||1|||||||||CENTAURXP|
'                    '3O|1|K1924282||^^^aHBs2|R|||||||||||||||||||CENTAURXP|
'                    '3O|1|K1924282||^^^aHBs2|R|||||||||||||||||||CENTAURXP|
'                    '3O|1|03498303||^^^Na   |R||||||||||1|||||||||ADVIA1800|
'                    '3O|1|03498300||^^^Na   |R||||||||||1|||||||||ADVIA1800|
'
'
'                    mResult.EqpCd = ""
'
'                    strBarno = mGetP(mGetP(strRcvBuf, 3, "|"), 1, "^")
'                    strRackNo = mGetP(mGetP(strRcvBuf, 3, "|"), 2, "^")
'                    strTubePos = mGetP(mGetP(strRcvBuf, 3, "|"), 3, "^")
'
'                    strEqpNm = mGetP(strRcvBuf, 25, "|")
'                    If strEqpNm = "" Then
'                        strEqpNm = mGetP(strRcvBuf, 26, "|")
'                    End If
'
'                    If strEqpNm <> "" Then
'                        If UCase(strEqpNm) = "CENTAURXP" Then
'                            mResult.EqpCd = gCENXPCD
'                        ElseIf UCase(strEqpNm) = "ADVIA1800" Then
'                            mResult.EqpCd = gADV18CD
'                        End If
'                    End If
'
'                    With mResult
'                        .BarNo = strBarno
'                        .SpcPos = strTubePos & "/" & strRackNo
'                        .Seq = strSeq
'                        .RackNo = mResult.EqpCd         'strRackNo
'                        .TubePos = Mid(strEqpNm, 1, 3)  'strTubePos
'                        If strOldBarno <> strBarno Then
'                            strOldBarno = strBarno
'                            .RsltDate = Format(Now, "yyyymmddhhmmss")
'                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'
'                            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'                        End If
'                    End With
'
'
'                Case "R"
'                    '6R|2|^^^aHBs2^^^1^COFF|1.00|mIU/mL||<|N|F||||20170831143313|CENTAURXP
'                    '4R|1|^^^CKMB^^^1^DOSE|2.30  |ng/mL|| |N|F||||20170831143543|CENTAURXP
'                    '5R|2|^^^CKMB^^^1^COFF|1.00  |ng/mL|| |N|F||||20170831143543|CENTAURXP
'                    '6R|3|^^^CKMB^^^1^RLU |14171 |     || |N|F||||20170831143543|CENTAURXP
'
'                    '4R|1|^^^Na|168||||                    N|F||||20170831051840|ADVIA1800
'
'                    strTemp1 = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
'                    strIntBase = strTemp1
'                    strAspect = mGetP(mGetP(strRcvBuf, 3, "|"), 8, "^")
'                    strTemp2 = mGetP(strRcvBuf, 4, "|")
'                    strFlag = mGetP(strRcvBuf, 7, "|")                  '<
'                    strIntResult = mGetP(strRcvBuf, 4, "|")
'
'                    'mResult.EqpNm = mGetP(strRcvBuf, 14, "|")           'CENTAURXP / ADVIA1800
'                    If mResult.EqpCd = gCENXPCD Then
'                        If strIntBase = "HBsII" Or strIntBase = "EHIV" Then 'INDX
'                            strIntBase = strIntBase & "_" & strAspect
'                            If strAspect = "INTR" Then  '정성결과
'                                strINTRResult = strIntResult
'                            End If
'                            If strAspect = "INDX" Then
'                                If UCase(strINTRResult) = "REACT" Then
'                                    strResult = "POSITIVE" & "(" & strIntResult & ")"
'                                Else
'                                    strResult = "NEGATIVE" & "(" & strIntResult & ")"
'                                End If
'                            End If
'
'                        ElseIf strIntBase = "aHBs2" Or strIntBase = "aHAVT" Or strIntBase = "aHAVM" Then
'                            strIntBase = strIntBase & "_" & strAspect
'                            If strAspect = "INTR" Then
'                                strINTRResult = strIntResult
'                            End If
'                            If strAspect = "DOSE" Then
'                                If UCase(strINTRResult) = "REACT" Then
'                                    strResult = "POSITIVE" & "(" & strIntResult & ")"
'                                Else
'                                    strResult = "NEGATIVE" & "(" & strIntResult & ")"
'                                End If
'                            End If
'
'                        ElseIf strIntBase = "aHCV" Then
'                            strIntBase = strIntBase & "_" & strAspect
'                            If strAspect = "INTR" Then
'                                strINTRResult = strIntResult
'                            End If
'                            If strAspect = "INDX" Then
'                                If UCase(strINTRResult) = "REACT" Then
'                                    strResult = "POSITIVE" & "(" & strIntResult & ")"
'                                Else
'                                    strResult = "NEGATIVE" & "(" & strIntResult & ")"
'                                End If
'                            End If
'                        Else
'                            If strAspect = "DOSE" Then
'                                strResult = strIntResult
'                            End If
'                        End If
'                    Else
'                        strResult = strIntResult
'                    End If
'
'                    If strIntBase <> "" And strResult <> "" Then
'                        If gPatOrdCd <> "" Then
'                            SQL = ""
'                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
'                            SQL = SQL & "  FROM EQPMASTER" & vbCr
'                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
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
'                                    If lsTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
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
'                                '-- 계산항목 처리
'                                Call CalProcess(gRow)
'
'                                '-- BIORAD QC 저장
''                                If Mid(strBarno, 1, 2) = "QC" Then
''                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
''                                End If
'
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
'                            SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
'                            SQL = SQL & "  FROM EQPMASTER" & vbCr
'                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
'
'                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                                lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                strQCLab = Trim(RS_L.Fields("QCLab") & "")
'                                strQCLot = Trim(RS_L.Fields("QCLot") & "")
'                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
'                                strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
'                                strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
'                                strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
'                                strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
'                                strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
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
'                                    If lsTestCode = Trim(gArrEQP(intCol - colSTATE, 2)) Then
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
'                                '-- BIORAD QC 저장
'                                If Mid(strBarno, 1, 2) = "QC" Then
'                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
'                                End If
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
''                Case "C"    '## Comment
''                    '## Abnormal 결과일때 Comment 저장
''                    If strFlag <> "N" Then
''                        strTemp1 = mGetP(strRcvBuf, 4, "|")
''                        strComm = mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")
''                    End If
''
''                Case "L"
''                    '## DB에 결과저장
'                    If .optTrans(0).Value = True And strState = "R" Then
'                        Res = SaveTransData_MCC_VERSACELL(gRow)
'
'                        If Res = -1 Then
'                            '-- 저장 실패
'                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                            SetText .spdOrder, "Failed", gRow, colSTATE
'                        Else
'                            '-- 저장 성공
'                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
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
'                            If DBExec(AdoCn_Local, SQL) Then
'                                '-- 성공
'                            End If
'                        End If
'                        strState = ""
'                    End If
'            End Select
'        Next
'    End With
'
'End Sub
'

''-----------------------------------------------------------------------------'
''   기능 : 해당 바코드번호에 대한 1. 접수정보 조회,
''                                 2. 장비수신정보 화면표시,
''                                 3. 처방코드 가져오기,
''                                 4. (처방코드로)검사오더 만들기
''   인수 :
''       - pBarNo : 바코드번호
''       - pType  : 바코드 미사용시 비교하는 대상
''                   1 : Seq
''                   2 : Rack/Pos
''                   3 : 체크된것중 제일 위에 것
''-----------------------------------------------------------------------------'
'Private Sub GetOrder(ByVal pBarno As String, ByVal pType As String)
'
'    Dim i           As Integer
'    Dim intRow      As Long
'    Dim strItems    As String
'    Dim strOrder    As String
'    Dim strDate     As String
'    Dim strInNum    As String
'    Dim strGumNum   As String
'
'    intRow = -1
'
'    '-- 1. 접수정보 조회
'    With frmMain
'        '-- 바코드 사용
'        If .optBarSeq(0).Value = True Then
'            For i = 1 To .spdOrder.DataRowCnt
'                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarno Then
'                    intRow = i
'                    Exit For
'                End If
'            Next i
'        Else
'            Select Case pType
'                '-- Seq
'                Case "1"
'                    For i = 1 To .spdOrder.DataRowCnt
'                        If Val(Trim(GetText(frmMain.spdOrder, i, colSEQNO))) = Val(mOrder.Seq) Then
'                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
'                            mOrder.BarNo = pBarno
'                            intRow = i
'                            Exit For
'                        End If
'                    Next i
'                '-- Rack/Pos
'                Case "2"
'                    For i = 1 To .spdOrder.DataRowCnt
'                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
'                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
'                            intRow = i
'                            Exit For
'                        End If
'                    Next i
'                '-- Check Top
'                Case "3"
'                    For i = 1 To .spdOrder.DataRowCnt
'                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
'                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
'                            mOrder.BarNo = pBarno
'                            intRow = i
'                            Exit For
'                        End If
'                    Next i
'            End Select
'        End If
'
'        '-- 스프레드에서 못찾았음..
'        If intRow < 0 Then
'            intRow = .spdOrder.DataRowCnt + 1
'            If .spdOrder.MaxRows < intRow Then
'                .spdOrder.MaxRows = intRow
'            End If
'        End If
'
'        '-- 장비수신정보 화면표시
'        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
'        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
'        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
'        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
'
'        '-- 결과스프레드 지우기
'        .spdResult.MaxRows = 0
'
'        '-- 검사자 정보 가져오기
'        Call GetSampleInfo(intRow, .spdOrder)
'
'        .spdOrder.RowHeight(-1) = 12
'
'        '-- 로컬테이블에서 검사항목에 해당하는 검사채널 찾아오기 (intRow = 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.)
'        strItems = GetEquipExamCode_VERSACELL(gHOSP.MACHCD, pBarno, intRow)
'
'        '-- 검사채널로 장비오더 만들기
'        If Trim(strItems) = "" Then
'            mOrder.NoOrder = True
'            mOrder.Order = ""
'
'            '-- 진행상태(Order) 표시
'            Call SetText(frmMain.spdOrder, "오더없음", intRow, colSTATE)
'        Else
'            mOrder.NoOrder = False
'            mOrder.Order = strItems
'
'            '-- 진행상태(Order) 표시
'            Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)
'        End If
'
'
'        '-- 현재 Row
'        gRow = intRow
'
'    End With
'
'End Sub

''검체번호에 존재하는 장비번호 해당하는 수가코드 가져오기
''한 장비 번호에 검사코드가 1개이상 존재
'Private Function GetEquipExamCode_VERSACELL(argEquipCode As String, argPID As String, Optional intRow As Long) As String
'    Dim i As Integer
'    Dim sExamCode As String
'    Dim strExamCode As String
'    Dim sSpecNo     As String
'    Dim iRow        As Long
'    Dim SpecNo      As String
'
'    GetEquipExamCode_VERSACELL = ""
'
'    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
'        Exit Function
'    End If
'
'    '-- 가져온 검사코드의 채널 찾기
'          SQL = "Select DISTINCT SENDCHANNEL "
'    SQL = SQL & "  From EQPMASTER "
'    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
'    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
'
'    strExamCode = ""
'
'    AdoCn_Local.CursorLocation = adUseClient
'    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
'    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
'        Do Until AdoRs_Local.EOF
'            If AdoRs_Local.Fields("SENDCHANNEL").Value & "" <> "990" Then
'                strExamCode = strExamCode & "\^^^" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
'            End If
'            AdoRs_Local.MoveNext
'        Loop
'    End If
'
'    AdoRs_Local.Close
'
'    GetEquipExamCode_VERSACELL = Mid(strExamCode, 2)
'
'End Function

