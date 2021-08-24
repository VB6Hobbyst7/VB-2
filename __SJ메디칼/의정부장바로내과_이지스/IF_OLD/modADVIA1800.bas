Attribute VB_Name = "modADVIA1800"
Option Explicit


'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput   As String     '송신할 데이터
    
    Select Case sSndState
        Case ""
            iIdleFlag = CStr(Val(iIdleFlag) + 1)
            
            '## Order 없는 경우
            If mOrder.NoOrder = True Then
                strOutput = ""
                strOutput = intFrameNo & "O" & " " & "0101"
                strOutput = strOutput & "000"                                                   'Sample Count
                strOutput = strOutput & "N"                                                     'Sample classification
                strOutput = strOutput & "2"                                                     'Registration data(0:New, 1:Add, 2:No Request, 3:Sample Delete)
                strOutput = strOutput & Left$(mOrder.BarNo & Space(13), 13)                     'Sample Number
                strOutput = strOutput & Space$(7) & Space$(16) & Space$(16) & "M" & Space$(3)   'Length = 45
                strOutput = strOutput & Space$(8) & " 1.0" & "1" & "1"                          '
                strOutput = strOutput & Space$(1) & ETX
            Else
                '1O 0101010N003498582                                            M            1.011 89M 81M 82M 90M 91M 85M106M103M104M105M 15
                strOutput = ""
                strOutput = intFrameNo & "O" & " " & "0101"
                strOutput = strOutput & Format$(mOrder.SendCnt, "000")                          'Sample Count
                strOutput = strOutput & "N"                                                     'Sample classification
                strOutput = strOutput & "0"                                                     'Registration data(0:New, 1:Add, 2:No Request, 3:Sample Delete)
                strOutput = strOutput & Left$(mOrder.BarNo & Space(13), 13)                     'Sample Number
                strOutput = strOutput & Space$(7) & Space$(16) & Space$(16) & "M" & Space$(3)   'Length = 45
                strOutput = strOutput & Space$(8)                                               '
                strOutput = strOutput & " 1.0"                                                  'Dilution coefficient(4)
                If mOrder.SPCCD = "2" Then                                                      'Sample classification(1:blood serum, 2:urine)
                    strOutput = strOutput & "2"
                Else
                    strOutput = strOutput & "1"
                End If
                strOutput = strOutput & "1"                                                     'Container classification
                strOutput = strOutput & mOrder.Order & Space$(1) & ETX
                
            End If
            
            'n개의 sSndPacket 구성
            ReDim Preserve sSndPacket(Val(iIdleFlag))
            sSndPacket(Val(iIdleFlag)) = STX & strOutput & GetChkSum(strOutput) & vbCr & vbLf
            
            intFrameNo = intFrameNo + 1
            
        Case "E"  '## 처음 Packet 전송
            iOrderFlag = 1
            frmMain.comEqp.Output = sSndPacket(iOrderFlag)
            SetRawData "[Tx]" & sSndPacket(iOrderFlag)
            
            If iOrderFlag = iTotQueryFlag Then
                sSndState = "L"
            Else
                sSndState = "P"
            End If
            
        Case "P"  '## Packet 전송
            iOrderFlag = iOrderFlag + 1
            frmMain.comEqp.Output = sSndPacket(iOrderFlag)
            SetRawData "[Tx]" & sSndPacket(iOrderFlag)
            
            If iOrderFlag = iTotQueryFlag Then
                sSndState = "L"
            Else
                sSndState = "P"
            End If
            
        Case "L"  '## EOT
            'strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            
            iOrderFlag = 0: iPendingFlag = 0: iIdleFlag = 0: iTotQueryFlag = 0
            intFrameNo = 1
            
            Exit Sub
    End Select
    
    If intFrameNo = 8 Then
        intFrameNo = 1
    End If
    
'    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
'    frmMain.comEqp.Output = strOutput
'    SetRawData "[Tx]" & strOutput

End Sub


Public Sub Phase_Serial_ADVIA1800()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        
                        sRcvState = "": sSndState = ""
                        intPhase = 2
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case Else
                        intPhase = 2
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case STX
                        intBufCnt = 1
                        Erase strRecvData
                        ReDim Preserve strRecvData(intBufCnt)
                    Case EOT
                        Select Case sRcvState
                            Case "Q"
                                intPhase = 3
                                iTotQueryFlag = iPendingFlag
                                iPendingFlag = 0
                                
                                'Order전송 Start
                                frmMain.comEqp.Output = ENQ
                                sSndState = "E"
                                
                            Case "R"
                                intPhase = 1
                        End Select
                        
                        sRcvState = ""
                    
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case vbLf
                        intPhase = 2
                        Call SerialRcvData_ADVIA1800
                        
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case vbCr
                    
                    Case ETB
                    
                    Case Else
                        intPhase = 2
                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar

                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case ACK
                        Select Case sSndState
                            Case "E"        '<ENQ> 전송 후의 상태
                                Call SendOrder
                        
                            Case "P"        '<Packet> 전송 후의 상태
                                Call SendOrder
                                                
                            Case "L"        '마지막 <Packet> 전송 후의 상태
                                Call SendOrder
                                
                                'Order관련 초기화
                                sSndState = ""
                                Erase sSndPacket
                                intPhase = 1
                        End Select
                    
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case NAK
                        Select Case sSndState
                            Case "E"
                                frmMain.comEqp.Output = Chr(5)
                                intPhase = 3
                            Case "P"
                                frmMain.comEqp.Output = sSndPacket(iOrderFlag)
                                intPhase = 3
                            Case "L"
                                frmMain.comEqp.Output = sSndPacket(iOrderFlag)
                                intPhase = 3
                        End Select
                        
                    Case 4      'EOT
                        Erase strRecvData
                        intPhase = 1
                        sRcvState = "": sSndState = ""
                        'Order관련 초기화
                        iPendingFlag = 0: iTotQueryFlag = 0
                        
                End Select
        End Select
    Next i
            
End Sub


Private Sub SerialRcvData_ADVIA1800()
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

    iBCpos = 2
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            strType = Mid$(strRcvBuf, iBCpos, 1)
            
            Select Case strType
                Case "q"    '## Request Information(Batch)
                    sRcvState = "Q"
                    sSndState = ""
                    
                Case "Q"    '## Request Information
                    sRcvState = "Q"
                    sSndState = ""
                
                    iTmpPendingFlag = Val(Mid$(strRcvBuf, iBCpos + 6, 2))
                    iPendingFlag = iPendingFlag + iTmpPendingFlag
                    
                    For i = 1 To iPendingFlag
                        strBarno = Trim$(Mid$(strRcvBuf, iBCpos + 9 + 13 * (i - 1), 13))
                        
                        With mOrder
                            .NoOrder = False
                            .BarNo = strBarno
                        End With
                        
                        Call GetOrder(strBarno, gHOSP.RSTTYPE)
                        Call SendOrder
                    Next
                
                Case "R"
                    sRcvState = "R"
                    
                    iTBlockNo = Val(Mid$(strRcvBuf, iBCpos + 2, 2))
                    iCBlockNo = Val(Mid$(strRcvBuf, iBCpos + 4, 2))
                    iItemNo = Val(Mid$(strRcvBuf, iBCpos + 6, 3))
                    
                    iBCpos = iBCpos + 6
                    
                    strKind = Mid$(strRcvBuf, iBCpos + 17, 1)       'N:Sample, C:Control
                    strBarno = Trim$(Mid$(strRcvBuf, iBCpos + 19, 13))
                                    
                    strTemp2 = Trim$(Mid$(strRcvBuf, iBCpos + 32, 7))
                    iPos = InStr(strTemp2, "-")
                             
                    If iPos = 0 Then
                        strRackNo = ""
                        strTubePos = ""
                    Else
                        strRackNo = Mid$(strTemp2, 1, iPos - 1)
                        strTubePos = Mid$(strTemp2, iPos + 1)
                    End If
                    
                    If strKind = "C" Then       'Control Result
                        strKind = "QC"
                    Else
                        strKind = ""
                    End If
                    
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
                    
                    ReDim Preserve varIntBase(iItemNo)
                    ReDim Preserve varResult(iItemNo)
                    ReDim Preserve varFlag(iItemNo)
                    
                    If iCBlockNo = 1 Then
                        For i = 1 To iItemNo
                            varIntBase(i) = Trim$(Mid(strRcvBuf, iBCpos + 89 + 19 * (i - 1), 3))
                            varResult(i) = Trim(Mid(strRcvBuf, iBCpos + 89 + 4 + 19 * (i - 1), 8))
                            varFlag(i) = Trim(Mid(strRcvBuf, iBCpos + 89 + 8 + 4 + 19 * (i - 1), 3))
                            
                            If InStr(varFlag(i), "R") > 0 Then
                                mResult.Rerun = "R"
                                varFlag(i) = Replace(varFlag(i), "R", "")
                            End If
                        Next i
                    Else
                        For i = 1 To iItemNo
                            varIntBase(i) = Trim$(Mid(strRcvBuf, iBCpos + 39 + 19 * (i - 1), 3))
                            varResult(i) = Trim(Mid(strRcvBuf, iBCpos + 39 + 4 + 19 * (i - 1), 8))
                            varFlag(i) = Trim(Mid(strRcvBuf, iBCpos + 39 + 8 + 4 + 19 * (i - 1), 3))
                            
                            If InStr(varFlag(i), "R") > 0 Then
                                mResult.Rerun = "R"
                                varFlag(i) = Replace(varFlag(i), "R", "")
                            End If
                        Next i
                    End If
                    
                    If mResult.Rerun = "R" Then       'Rerun Result
                        mResult.Kind = mResult.Kind & "R"
                    End If
                    
                    For i = 1 To iItemNo
                        strIntBase = varIntBase(i)
                        strResult = varResult(i)
                        
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
                    Next
                    
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
            End Select
        Next
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
        strItems = GetEquipExamCode_ADVIA1800(gHOSP.MACHCD, pBarno, intRow)

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
Private Function GetEquipExamCode_ADVIA1800(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_ADVIA1800 = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    '-- 가져온 검사코드의 채널 찾기
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    strExamCode = ""
    mOrder.SendCnt = 0
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            ' " 89M 81M 82M 90M 91M108M 85M"
            strExamCode = strExamCode & Right(Space(3) & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & ""), 3) & "M"
            mOrder.SendCnt = mOrder.SendCnt + 1
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_ADVIA1800 = Mid(strExamCode, 2)
    
End Function



