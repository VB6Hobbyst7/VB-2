Attribute VB_Name = "modRAPIDPOINT500"
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
                strOutput = strOutput & "000"
                strOutput = strOutput & "N"
                strOutput = strOutput & "2"
                strOutput = strOutput & Left$(mOrder.BarNo & Space(13), 13)
                strOutput = strOutput & Space$(7) & Space$(16) & Space$(16) & "M" & Space$(3)
                strOutput = strOutput & Space$(8) & " 1.0" & "1" & "1"
                strOutput = strOutput & Space$(1) & ETX
            Else
                strOutput = ""
                strOutput = intFrameNo & "O" & " " & "0101"
                strOutput = strOutput & Format$(mOrder.SendCnt, "000")
                strOutput = strOutput & "N"                                   'Sample classification
                strOutput = strOutput & "0"                                   'Registration data(0:New, 1:Add, 2:No Request, 3:Sample Delete)
                strOutput = strOutput & Left$(mOrder.BarNo & Space(13), 13)
                strOutput = strOutput & Space$(7) & Space$(16) & Space$(16) & "M" & Space$(3)
                strOutput = strOutput & Space$(8)
                strOutput = strOutput & " 1.0"        'Dilution coefficient(4)
                If mOrder.SPCCD = "2" Then
                    strOutput = strOutput & "2"           'Sample classification(1:blood serum, 2:urine)
                Else
                    strOutput = strOutput & "1"           'Sample classification(1:blood serum, 2:urine)
                End If
                strOutput = strOutput & "1"           'Container classification
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

Public Sub Phase_Serial_RAPIDPOINT500()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                AckOn = False
                RcvBuffer = BufChar
                
            Case EOT
                If AckOn = False Then
                    frmMain.comEqp.Output = STX & ACK & ETX & "0B" & EOT        'Ack Message
                    Call SerialRcvData_RAPIDPOINT500
                End If
            
            Case ACK
                AckOn = True
                RcvBuffer = RcvBuffer & BufChar
            
            Case Else
                RcvBuffer = RcvBuffer & BufChar
                
        End Select
    Next i
            
End Sub

Private Sub SendMessage_1200(ByVal MsgHead As String)
    Dim chksum As Integer
    Dim Buffer As String
    Dim C As Integer
'    Dim R As Integer

    Dim sSendData$
    
    Select Case MsgHead
        Case "ID_DATA"
            Buffer = STX & "ID_DATA" & FS & R_S _
                                    & "aMOD" & GS & "LIS" & GS & GS & GS & FS _
                                    & "iIID" & GS & "333" & GS & GS & GS & FS & R_S _
                                    & ETX
        Case "SMP_REQ"
            Buffer = STX & "SMP_REQ" & FS & R_S & "aMOD" & GS & aMod & GS & GS & GS _
                                        & FS & "iIID" & GS & iIID & GS & GS & GS _
                                        & FS & "rSEQ" & GS & Sample_Seq & GS & GS & GS _
                                        & FS & R_S & ETX
            
        Case "SMP_ORD"
    End Select
        
    For C = 1 To Len(Buffer)
        chksum = chksum + Asc(Mid(Buffer, C, 1))
    Next C
    
    sSendData = Buffer & Right("0" & Hex(chksum Mod 256), 2) & EOT
    
    frmMain.comEqp.Output = sSendData
    
End Sub

Private Sub GetaModiIID(ByVal sMsg As String)

    Dim tmpData()   As String
    
    '<STX>SYS_READY<FS><RS>aMOD<GS>1265<GS><GS><GS><FS>iIID
    '<GS>12345<GS><GS><GS><FS>aDATE<GS>20Jan2004<GS><GS><GS>
    '<FS>aTIME<GS>13:35:32<GS><GS><GS><FS>iOID<GS>3<GS><GS><GS><FS>
    '<ETX>{chksum}<EOT>

    tmpData() = Split(sMsg, GS)
    
    'aMod
    aMod = Trim(tmpData(1))
    
    'iIID
    iIID = Trim(tmpData(5))

End Sub


Private Function ConvertDateType(ByVal sDate As String) As String
    On Error GoTo ErrRtn
    
    Dim kk%
    Dim sTmp$
    Dim tmpYYYY$, tmpMM$, tmpDD$
    
    ConvertDateType = sDate
    
    tmpYYYY = Right(sDate, 4)
    sDate = Mid(sDate, 1, Len(sDate) - 4)
    
    For kk = 1 To Len(sDate)
        sTmp = Mid(sDate, kk, 1)
        If IsNumeric(sTmp) Then
            tmpDD = tmpDD & sTmp
        Else
            tmpMM = tmpMM & sTmp
        End If
    Next kk
    
    sTmp = tmpDD & Space(1) & tmpMM & Space(1) & tmpYYYY
    
    ConvertDateType = Format(sTmp, "YYYYMMDD")
    
ErrRtn:
    If Err <> 0 Then
        'RaiseEvent DispMsg("ConvertDateType - " & Err.Description)
    End If
End Function


Private Sub SerialRcvData_RAPIDPOINT500()
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

    Dim X   As Integer
    Dim C   As Integer
    Dim MsgID   As String
    
    Dim R   As Integer
    Dim x1  As Integer
    Dim x2  As Integer
    Dim AssayNm As String
    Dim RESULT  As String
    Dim EqCd    As String
    Dim OrdCd   As String
    Dim LabNo   As String
    Dim rSeq    As String
    Dim iPID    As String
    Dim iQID    As String

    Dim sRstDate$, sRstTime$
    Dim MsgBuf$
    
'    Dim strQCResult As String
    
    Dim iQLEV$, iQLOT$, strAnalyte$
    Dim db_tmp As String * 100
    
    With frmMain
        '-- 테스트용 -----------------
        If .fraCommTest.Visible = False Then
            Call SetSQLData("RCV", RcvBuffer, "A")
        End If
        '-- 테스트용 -----------------
        
        X = InStr(1, RcvBuffer, FS)
        If RcvBuffer <> "" Then
            MsgID = Mid(RcvBuffer, 2, X - 2)
        End If
        Select Case MsgID
            Case "ID_REQ"
                Call SendMessage_1200("ID_DATA")
            Case "SMP_START"
            Case "SMP_NEW_AV"
                Do Until X = 0
                    X = InStr(X, RcvBuffer, "r")
                    If X = 0 Then Exit Do
                    If Mid(RcvBuffer, X, 4) = "rSEQ" Then
                        X = X + 5
                        C = InStr(X, RcvBuffer, GS)
                        Sample_Seq = Mid(RcvBuffer, X, C - X)
                    End If
                    Call GetaModiIID(RcvBuffer)
                    Call SendMessage_1200("SMP_REQ")
                Loop
            
            Case "SYS_READY"
            Case "SYS_NOT_READY"
            Case "SMP_NEW_DATA", "SMP_EDIT_DATA"
                GoTo Rst
            Case "CAL_ABORT"
            Case "QC_START"
            Case "QC_NEW_AV"
                Do Until X = 0
                    X = InStr(X, RcvBuffer, "r")
                    If X = 0 Then Exit Do
                    If Mid(RcvBuffer, X, 4) = "rSEQ" Then
                        X = X + 5
                        C = InStr(X, RcvBuffer, GS)
                        Sample_Seq = Mid(RcvBuffer, X, C - X)
                    End If
                    Call GetaModiIID(RcvBuffer)
                    Call SendMessage_1200("SMP_REQ")
                Loop
            Case "QC_NEW_DATA", "QC_EDIT_DATA"
                GoTo Rst
        End Select
            
        Exit Sub

Rst:
        MsgBuf = RcvBuffer
        
        If MsgID = "SMP_NEW_DATA" Or MsgID = "SMP_EDIT_DATA" Then
            'aMod
            x1 = 1
            x1 = InStr(x1, MsgBuf, "aMod") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                aMod = Mid(MsgBuf, x1, x2 - x1)
            End If
        
            'iIID
            x1 = 1
            x1 = InStr(x1, MsgBuf, "iIID") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                iIID = Mid(MsgBuf, x1, x2 - x1)
            End If
        
            'rSEQ
            x1 = 1
            x1 = InStr(x1, MsgBuf, "rSEQ") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                rSeq = Mid(MsgBuf, x1, x2 - x1)
            End If
        
            'PID
            x1 = 1
            x1 = InStr(x1, MsgBuf, "iPID") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                iPID = Mid(MsgBuf, x1, x2 - x1)
            End If
            'DATE
            x1 = 1
            x1 = InStr(x1, MsgBuf, "rDATE") + 6
            If x1 <> 6 Then
                x2 = InStr(x1, MsgBuf, GS)
                sRstDate = Mid(MsgBuf, x1, x2 - x1)
                sRstDate = ConvertDateType(sRstDate)
            End If
            'TIME
            x1 = 1
            x1 = InStr(x1, MsgBuf, "rTIME") + 6
            If x1 <> 6 Then
                x2 = InStr(x1, MsgBuf, GS)
                sRstTime = Mid(MsgBuf, x1, x2 - x1)
                sRstTime = Format(sRstTime, "HHNNSS")
            End If
        
            x2 = 0
        
            '접수번호, SeqNo
            strBarno = Trim(iPID)
            strSeq = Trim(rSeq)
            
            If strBarno = "" Or Not IsNumeric(strBarno) Then
                Exit Sub
            End If
            
            With mResult
                .BarNo = strBarno
                .RackNo = strRackNo
                .TubePos = strTubePos
                .Rerun = ""
                If strOldBarno <> strBarno Then
                    strOldBarno = strBarno
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
            
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                End If
            End With
            
            strState = "O"
                        
            '----------------------------------------------------------------------------------------
            '   Measured Data
            '----------------------------------------------------------------------------------------
            x1 = 1
            Do While InStr(x1, MsgBuf, FS & "m") <> 0
                x1 = InStr(x1, MsgBuf, FS & "m")
                x2 = InStr(x1, MsgBuf, GS)
        
        '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                strIntBase = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))
        
                x2 = x2 + 1
                x1 = InStr(x2, MsgBuf, GS)
                strResult = Mid(MsgBuf, x2, x1 - x2)
        
                SetRawData "[결과]" & strIntBase & "," & strResult
                
                If strResult <> "" Then
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
            Loop
                    
                    
            '----------------------------------------------------------------------------------------
            '   Calibrated Data
            '----------------------------------------------------------------------------------------
            x1 = 1
            Do While InStr(x1, strRcvBuf, FS & "c") <> 0
                x1 = InStr(x1, strRcvBuf, FS & "c")
                x2 = InStr(x1, strRcvBuf, GS)
        
        '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                strIntBase = Mid(strRcvBuf, x1 + 1, x2 - (x1 + 1))
                x2 = x2 + 1
                x1 = InStr(x2, strRcvBuf, GS)
                strResult = Mid(strRcvBuf, x2, x1 - x2)
                
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
            Loop
            
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

            
        '>> If MsgID = "SMP_NEW_DATA" Or MsgID = "SMP_EDIT_DATA" Then
        
        ElseIf MsgID = "QC_NEW_DATA" Or MsgID = "QC_EDIT_DATA" Then
            '-- R348
'''            '## Type 조회
'''            Pos1 = InStr(strRcvBuf, "rTYPE")
'''            If Pos1 > 0 Then
'''                Pos2 = InStr(Mid$(strRcvBuf, Pos1), FS)
'''                strBarno = mGetP(Mid$(strRcvBuf, Pos1, Pos2), 2, GS)
'''                'strBarno = Val(strBarno)
'''            Else
'''                '## NOTE: WorkNo가 전송되지 않은 에러처리
'''                Exit Sub
'''            End If
'''
'''            '## Level 조회
'''            Pos1 = 0: Pos2 = 0
'''            Pos1 = InStr(strRcvBuf, "iQLEV")
'''            If Pos1 > 0 Then
'''                Pos2 = InStr(Mid$(strRcvBuf, Pos1), FS)
'''                strQCLevel = mGetP(Mid$(strRcvBuf, Pos1, Pos2), 2, GS)
'''            Else
'''                '## NOTE: 바코드번호가 전송되지 않은 에러처리
'''            End If
'''
'''
'''            '## QC 채널 조회
'''            Pos1 = 0: Pos2 = 0
'''            Pos1 = InStr(strRcvBuf, "iQFILE")
'''            If Pos1 > 0 Then
'''                Pos2 = InStr(Mid$(strRcvBuf, Pos1), FS)
'''                strQCChannel = mGetP(Mid$(strRcvBuf, Pos1, Pos2), 2, GS)
'''            Else
'''                '## NOTE: 바코드번호가 전송되지 않은 에러처리
'''            End If
            
            x1 = 1
            x1 = InStr(x1, MsgBuf, "aMod") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                aMod = Mid(MsgBuf, x1, x2 - x1)
            End If
        
            'iIID
            x1 = 1
            x1 = InStr(x1, MsgBuf, "iIID") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                iIID = Mid(MsgBuf, x1, x2 - x1)
            End If
        
            'rSEQ
            x1 = 1
            x1 = InStr(x1, MsgBuf, "rSEQ") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                rSeq = Mid(MsgBuf, x1, x2 - x1)
            End If
        
            'PID
            x1 = 1
            x1 = InStr(x1, MsgBuf, "iPID") + 5
            If x1 <> 5 Then
                x2 = InStr(x1, MsgBuf, GS)
                iPID = Mid(MsgBuf, x1, x2 - x1)
            End If
            'DATE
            x1 = 1
            x1 = InStr(x1, MsgBuf, "rDATE") + 6
            If x1 <> 6 Then
                x2 = InStr(x1, MsgBuf, GS)
                sRstDate = Mid(MsgBuf, x1, x2 - x1)
                sRstDate = ConvertDateType(sRstDate)
            End If
            'TIME
            x1 = 1
            x1 = InStr(x1, MsgBuf, "rTIME") + 6
            If x1 <> 6 Then
                x2 = InStr(x1, MsgBuf, GS)
                sRstTime = Mid(MsgBuf, x1, x2 - x1)
                sRstTime = Format(sRstTime, "HHNNSS")
            End If
        
            x2 = 0
        
            '접수번호, SeqNo
            strBarno = Trim(iPID)
            strSeq = Trim(rSeq)
            
            If strBarno = "" Or Not IsNumeric(strBarno) Then
                Exit Sub
            End If
            
            With mResult
                .BarNo = strBarno
                .RackNo = strRackNo
                .TubePos = strTubePos
                .Rerun = ""
                If strOldBarno <> strBarno Then
                    strOldBarno = strBarno
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
            
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                End If
            End With
            
            strState = "O"
                        
            '----------------------------------------------------------------------------------------
            '   Measured Data
            '----------------------------------------------------------------------------------------
            x1 = 1
            Do While InStr(x1, MsgBuf, FS & "m") <> 0
                x1 = InStr(x1, MsgBuf, FS & "m")
                x2 = InStr(x1, MsgBuf, GS)
        
        '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                strIntBase = Mid(MsgBuf, x1 + 1, x2 - (x1 + 1))
        
                x2 = x2 + 1
                x1 = InStr(x2, MsgBuf, GS)
                strResult = Mid(MsgBuf, x2, x1 - x2)
        
                SetRawData "[결과]" & strIntBase & "," & strResult
                
                If strResult <> "" Then
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
            Loop
                    
                    
            '----------------------------------------------------------------------------------------
            '   Calibrated Data
            '----------------------------------------------------------------------------------------
            x1 = 1
            Do While InStr(x1, strRcvBuf, FS & "c") <> 0
                x1 = InStr(x1, strRcvBuf, FS & "c")
                x2 = InStr(x1, strRcvBuf, GS)
        
        '        AssayNm = Mid(MsgBuf, x1 + 2, x2 - (x1 + 2))
                'Ca++의 경우 장비검사코드가 동일하기 때문에 Measured & Calibrated 의 구분이 필요...
                strIntBase = Mid(strRcvBuf, x1 + 1, x2 - (x1 + 1))
                x2 = x2 + 1
                x1 = InStr(x2, strRcvBuf, GS)
                strResult = Mid(strRcvBuf, x2, x1 - x2)
                
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
            Loop
            
            .spdResult.RowHeight(-1) = 14
            
            
        End If
        
        
        
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
        'strItems = GetEquipExamCode_ADVIA1800(gHOSP.MACHCD, pBarno, intRow)

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
Private Function GetEquipExamCode_RAPIDPOINT500(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_RAPIDPOINT500 = ""
    
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
    
    GetEquipExamCode_RAPIDPOINT500 = Mid(strExamCode, 2)
    
End Function





