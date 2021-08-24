Attribute VB_Name = "modCommunication"
Option Explicit

Public pBuffer As Variant

'-- 수신한 오더정보
Type RecvData
    BarNo       As String
    Seq         As String
    RackNo      As String
    TubePos     As String
    NoOrder     As Boolean
    Order       As String
    IsSending   As Boolean
    SendCnt     As Integer
    isresult    As Boolean
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
    TESTCD   As String
    DoctNm   As String
End Type

Public mResult As IntfData

Public Sub Serial_Protocol()

    Select Case UCase(gHOSP.MACHNM)
        Case "E411"
                Call Phase_Serial_E411
        Case "AU400"
                Call Phase_Serial_AU400
        Case "AU480"
                Call Phase_Serial_AU480
        Case "XN1000"
                Call Phase_Serial_XN1000
        Case Else
            
    End Select
    

End Sub

Public Sub TCP_Protocol()

    Select Case UCase(gHOSP.MACHNM)
        Case "BA400"
                Call Phase_TCP_BA400
        Case ""
        
    End Select
    
End Sub

Public Sub FILE_Protocol()

    Select Case UCase(gHOSP.MACHNM)
        Case "CFX96"
                Call Phase_FILE_CFX96
        Case ""
        
    End Select
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Public Sub SendOrder()
    Dim strOutput As String     '송신할 데이터
    
    '-- ASTM TYPE별 Define 해야함.
    '-- ASTM TYPE = Standard
    Select Case intSndPhase
        Case 1  '## Header
            'strOutput = intFrameNo & "H|\^&||| XN-10^00-14^15097^^^^AP795756||||||||E1394-97" & vbCr & ETX
            strOutput = intFrameNo & "H|\^&||||||||||P|1" & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
        Case 2  '## Patient
            'strOutput = intFrameNo & "P|1||||^^|||U|||||^||||||||||||^^^" & vbCr & ETX
            strOutput = intFrameNo & "P|1" & vbCr & ETX
            
            intSndPhase = 4
            intFrameNo = intFrameNo + 1
            
        Case 3  '## No Order
            
        Case 4  '## Order
            If mOrder.NoOrder = True Then
                    
                strOutput = intFrameNo & "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q"
                intSndPhase = 5
            
            Else
                If mOrder.IsSending = False Then   '## 최초 보낼때
                    strOutput = "O|1|" & mOrder.RackNo & "^" & mOrder.TubePos & "^" & Right(Space(15) & mOrder.BarNo, 15) & "^B||" & mOrder.Order & "|||||||N||||||||||||||Q"
                    
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Left(strOutput, 230) & vbCr & ETB
                        intSndPhase = 4
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 5
                    End If
                Else                        '## 남은 문자열이 있을때
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Left(strOutput, 230) & vbCr & ETB
                        intSndPhase = 4
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 5
                    End If
                End If
                
            End If
            
            intFrameNo = intFrameNo + 1
            
        Case 5  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 6
            intFrameNo = intFrameNo + 1
            
        Case 6  '## EOT
            strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1
            
            Exit Sub
    End Select
    
    
    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    frmMain.comEqp.Output = strOutput
'    Debug.Print strOutput
    SetRawData "[Tx]" & strOutput
    
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
'                                 3. 처방코드 가져오기,
'                                 4. (처방코드로)검사오더 만들기
'   인수 :
'       - pBarNo : 바코드번호
'       - pType  : 바코드 미사용시 비교하는 대상
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : 체크된것중 제일 위에 것
'-----------------------------------------------------------------------------'
Public Sub GetOrder(ByVal pBarno As String, ByVal pType As String)

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
        strItems = GetEquipExamCode_AU480(gHOSP.MACHCD, pBarno, intRow)

        '-- 검사채널로 장비오더 만들기
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""

            'S 003401 0019          1013001918    E
            SetRawData "[Tx]" & STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(26 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & ETX
            frmMain.comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(26 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & ETX

        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems

            '                    Rack     Pos          Seq      장비설정된 바코드 자리수만큼
            '                                                   예를들어 장비설정이 20자리고 바코드 자리가 12자리면 바코드번호앞에 스페이스 8자리를 줘야한다.
            '                                                                                   검사채널(채널당 2자리)

            'S 003401 0019          1013001918    E      01020304050607091011121415161719212632
            SetRawData "[Tx]" & STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(26 - Len(mOrder.BarNo)) & mOrder.BarNo & Space(4) & "E" & strItems & ETX
            frmMain.comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(26 - Len(mOrder.BarNo)) & mOrder.BarNo & Space(4) & "E" & strItems & ETX
        End If

        '-- 진행상태(Order) 표시
        Call SetText(frmMain.spdOrder, "오더전송", intRow, colSTATE)


        '-- 현재 Row
        gRow = intRow
        
    End With
    
End Sub

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
    Dim strState    As String
    
    intRow = -1
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
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" And GetText(frmMain.spdOrder, i, colSTATE) = "" Then
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
    
        Call SetText(.spdOrder, "1", intRow, colCHECKBOX)
        
        '-- 장비결과인덱스 화면표시
        Call SetText(.spdOrder, mResult.RsltSeq, intRow, colSAVESEQ)
        Call SetText(.spdOrder, mResult.RsltDate, intRow, colEXAMDATE)
        
        '-- 장비수신정보 화면표시
        Call SetText(.spdOrder, mResult.BarNo, intRow, colBARCODE)
        'Call SetText(.spdOrder, mResult.Seq, intRow, colSEQNO)
        'Call SetText(.spdOrder, mResult.RackNo, intRow, colRACKNO)
        'Call SetText(.spdOrder, mResult.TubePos, intRow, colPOSNO)
    
        '-- 환자정보 표시
        'Call vasActiveCell(.spdOrder, intRow, colBARCODE)
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        If .optCheck(0).Value = True Then
            strState = "0,2"
        ElseIf .optCheck(1).Value = True Then
            strState = "0"
        ElseIf .optCheck(2).Value = True Then
            strState = "2"
        End If
        
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder, "", strState)
        
        .spdOrder.RowHeight(-1) = 16
    
    End With
    
    '-- 현재 Row
    gRow = intRow
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Public Sub SendOrder_E411()
    
    
End Sub


Public Sub Phase_TCP_BA400()
 
End Sub
    

Public Sub Phase_Serial_E411()


End Sub

Public Sub Phase_Serial_AU400()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                intBufCnt = 1
                Erase strRecvData
                ReDim Preserve strRecvData(intBufCnt)
            Case ETB
            Case ETX
                Call SerialRcvData_AU400
            Case Else
                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
        End Select
    Next i

End Sub

Public Sub Phase_FILE_CFX96()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    'strRecvData = Split(pBuffer, vbLf)
    
    If UBound(strRecvData) > 0 Then
        Call FileRcvData_CFX96
    End If

End Sub

Public Sub Phase_Serial_AU480()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(Buffer, i, 1)
        Select Case BufChar
            Case STX
                intBufCnt = 1
                Erase strRecvData
                ReDim Preserve strRecvData(intBufCnt)
            Case ETB
            Case ETX
                Call SerialRcvData_AU480
            Case Else
                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
        End Select
    Next i

End Sub

Public Sub Phase_Serial_XN1000()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)
    With frmMain
        For i = 1 To lngBufLen
            BufChar = Mid$(pBuffer, i, 1)
            Select Case intPhase
                Case 1      '## Estabilshment Phase
                    Select Case BufChar
                        Case ENQ
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                            intPhase = 2
                            .comEqp.Output = ACK
                            DoEvents
                            SetRawData "[Tx]" & ACK
                        Case ACK
                            If strState = "Q" Then Call SendOrder
                    
                    End Select
                Case 2      '## Transfer Phase
                    Select Case BufChar
                        Case ENQ
                            Erase strRecvData
                            .comEqp.Output = ACK
                            DoEvents
                            SetRawData "[Tx]" & ACK
                        Case STX
                            intBufCnt = 1
                            Erase strRecvData
                            ReDim Preserve strRecvData(intBufCnt)
                        Case ETB
                            blnIsETB = True
                            intPhase = 3
                        Case ETX
                            intBufCnt = intBufCnt + 1
                            ReDim Preserve strRecvData(intBufCnt)
                            intPhase = 3
                        Case vbCr, vbLf
                        Case EOT
                            intPhase = 1
                        Case Else
                            If blnIsETB = False Then
                                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                            Else
                                blnIsETB = False
                            End If
                    End Select
                Case 3      '## Transfer Phase
                    Select Case BufChar
                        Case vbCr
                        Case vbLf
                            intPhase = 4
                            .comEqp.Output = ACK
                            DoEvents
                            SetRawData "[Tx]" & ACK
                    End Select
                Case 4      '## Termination Phase
                    Select Case BufChar
                        Case STX
                            intPhase = 2
                        Case EOT
                            Call SerialRcvData_XN1000
                            If strState = "Q" Then
                                intSndPhase = 1
                                intFrameNo = 1
                                .comEqp.Output = ENQ
                                DoEvents
                                SetRawData "[Tx]" & ENQ
                            End If
                            
                            intPhase = 1
                    End Select
            End Select
        Next i
    End With
        
End Sub


Public Sub SerialRcvData_XN1000()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '수신한 Data
    Dim strRcvBufOrd    As String
    Dim strRcvBufRst    As String
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
    
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
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
                Case "H"
                Case "Q"    '## Inquiry Order
                        strBarno = Trim(Mid(strRcvBuf, 14, 26))
                        strSeq = Mid(strRcvBuf, 9, 5)
                        strRackNo = Mid(strRcvBuf, 3, 4)
                        strTubePos = Mid(strRcvBuf, 7, 2)
                    
                        With mOrder
                            .BarNo = strBarno
                            .Seq = strSeq
                            .RackNo = strRackNo
                            .TubePos = strTubePos
                        End With
                        
                        If strBarno = "" Then
                            strBarno = "NoOrder_" & Trim(strSeq)
                        End If
                        
                        Call GetOrder(strBarno, gHOSP.RSTTYPE)
                        
                        strState = "Q"
    
                Case "P"
                
                Case "O"
                    '4O|1||1^6^          201404240002^B|^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NEUT#\^^^^LYMPH#\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH%\^^^^MONO%\^^^^EC|1||
                    
                    strRcvBufOrd = Trim$(mGetP(strRcvBuf, 4, "|"))
                    strBarno = Trim$(mGetP(strRcvBufOrd, 3, "^"))
                    strSeq = ""
                    strRackNo = Trim$(mGetP(strRcvBufOrd, 1, "^"))
                    strTubePos = Trim$(mGetP(strRcvBufOrd, 2, "^"))
                    
                    With mResult
                        .BarNo = strBarno
                        .SpcPos = strSeq
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    End With
                    
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
                    If gRow <= 0 Then
                        Exit Sub
                    End If
                    
                    strState = "O"
                    .spdResult.MaxRows = 0
                    
                Case "R"
                    strRcvBufRst = Trim(mGetP(strRcvBuf, 3, "|"))
                    strIntBase = Trim$(mGetP(strRcvBufRst, 5, "^"))
                    strResult = Trim(mGetP(strRcvBuf, 4, "|"))
                    
                    If strIntBase <> "" And strResult <> "" Then
                        If gPatOrdCd <> "" Then
                            SQL = ""
                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                            SQL = SQL & "  FROM EQPMASTER" & vbCr
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.HOSPCD & "' " & vbCr
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
                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.HOSPCD & "' " & vbCr
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
                            
                Case "L"
                    '## DB에 결과저장
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData(gRow)
                        
                        If Res = -1 Then
                            '-- 저장 실패
'                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                            SetText .spdOrder, "Failed", gRow, colSTATE
                        Else
                            '-- 저장 성공
'                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                            SetText .spdOrder, "저장완료", gRow, colSTATE
                            SetText .spdOrder, "0", gRow, colCHECKBOX
                            
                                  SQL = "Update PATRESULT Set " & vbCrLf
                            SQL = SQL & " sendflag = '2' " & vbCrLf
                            SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                            SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                            SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                            SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                            
'                            Res = SendQuery(gLocal, SQL)
                            If Res = -1 Then
'                                SaveQuery SQL
                                Exit Sub
                            End If
                        End If
                        strState = ""
                    End If
                
            End Select
        Next
    End With
    
End Sub


Public Sub SerialRcvData_AU480()
    

End Sub

Public Sub SerialRcvData_AU400()
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
            strType = Mid$(strRcvBuf, 1, 2)
            
            Select Case strType
                Case "R "    '## Inquiry Order
                        'R 000101 0001                1608270009
                        strBarno = Trim(Mid(strRcvBuf, 14, 26))
                        strSeq = Mid(strRcvBuf, 9, 5)
                        strRackNo = Mid(strRcvBuf, 3, 4)
                        strTubePos = Mid(strRcvBuf, 7, 2)
                        
                        With mOrder
                            .BarNo = strBarno
                            .Seq = strSeq
                            .RackNo = strRackNo
                            .TubePos = strTubePos
                        End With
                        
                        If strBarno = "" Then
                            strBarno = "NoOrder_" & Trim(strSeq)
                        End If
                        
                        Call GetOrder(strBarno, gHOSP.RSTTYPE)
                        
                        strState = "Q"
        
                Case "D "    '## Result
                        'D 000101 0001                1608270009    E001   9.3  002   5.8  
                        strBarno = Trim$(Mid$(strRcvBuf, 14, 10))
                        strRackNo = Mid(strRcvBuf, 3, 4)
                        strTubePos = Mid(strRcvBuf, 7, 2)
                        strSeq = Trim(Mid(strRcvBuf, 9, 5))
                        
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
                        
                        strTmp = Mid$(strRcvBuf, 29)
                        
                        Do While Len(strTmp) >= 11
                            strIntBase = Mid$(strTmp, 2, 2)
                            strResult = Mid$(strTmp, 4, 6)
                            strComm = Mid$(strTmp, 10, 1)
                        
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
                            strTmp = Mid$(strTmp, 12)
                        Loop
                        
                        .spdResult.RowHeight(-1) = 14
                        
                        '## DB에 결과저장
                        If .optTrans(0).Value = True And strState = "R" Then
                            Res = SaveTransData(gRow)
                            
                            If Res = -1 Then
                                '-- 저장 실패
    '                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                                SetText .spdOrder, "Failed", gRow, colSTATE
                            Else
                                '-- 저장 성공
    '                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                                SetText .spdOrder, "저장완료", gRow, colSTATE
                                SetText .spdOrder, "0", gRow, colCHECKBOX
                                
                                      SQL = "Update PATRESULT Set " & vbCrLf
                                SQL = SQL & " sendflag = '2' " & vbCrLf
                                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                                
    '                            Res = SendQuery(gLocal, SQL)
                                If Res = -1 Then
    '                                SaveQuery SQL
                                    Exit Sub
                                End If
                            End If
                            strState = ""
                        End If
                    
            End Select
        Next
    End With

End Sub

Public Sub FileRcvData_CFX96()
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
    
    Dim strTarget       As String
    Dim strVarTest      As String
    Dim strLGrp         As String
    Dim strHGrp         As String
    Dim strTotFlag      As String
    
    Dim strTmp          As String
    Dim strICVal        As String
    Dim strICVal1       As String
    Dim strICVal2       As String
    
    With frmMain
        For intCnt = 1 To 30 'UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            Debug.Print strRcvBuf
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            If intCnt = 1 Then
                strICVal = ""
                strBarno = mGetP(strRcvBuf, 1, ",")
                strRackNo = Mid(strRcvBuf, 3, 4)
                strTubePos = Mid(strRcvBuf, 7, 2)
                strSeq = Trim(Mid(strRcvBuf, 9, 5))
                
                With mResult
                    .BarNo = strBarno
                    .SpcPos = strSeq
                    .Seq = strSeq
                    .RackNo = strRackNo
                    .TubePos = strTubePos
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                End With
                If strBarno = "" Then
                    Exit For
                End If
                
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
            End If
            
            '마지막 환자의 경우 컨트롤을 데이터를 갖고 있다
            'If Mid(Trim(mGetP(strRcvBuf, 3, ",")), 1, 2) = "PC" Or Mid(Trim(mGetP(strRcvBuf, 3, ",")), 1, 2) = "NC" Then
            '    Exit Sub
            'End If
            
            strTarget = Trim(mGetP(strRcvBuf, 5, ","))      '-- Target(타입)
            strIntBase = "HPV"
            strResult = Trim(mGetP(strRcvBuf, 6, ","))      '-- 판정(정성)
            strIntResult = Trim(mGetP(strRcvBuf, 7, ","))   '-- 결과(정량)
                
'            If strTarget = "IC" And Len(strResult) = 1 Then
            If strTarget = "IC" Then
                If UCase(strICVal) <> "INVALID" Then
                    If intCnt = 15 Then
                        strICVal1 = strResult
                    Else
                        strICVal2 = strResult
                    End If
                    strICVal = strResult
                End If
            
                If strTarget = "IC" And Len(strResult) = 1 Then
                    strVarTest = "에러"
                End If
            End If
            
'            If strTarget = "11" Or strTarget = "40" Or strTarget = "42" Or strTarget = "43" Or strTarget = "44" Or strTarget = "54" Or strTarget = "6" Or strTarget = "61" then
            '-- 6, 11, 40, 42, 43, 44, 54, 61
            
            'If strTarget = "35" Then Stop
            'If strTarget = "11" Or strTarget = "40" Or strTarget = "42" Or strTarget = "43" Or strTarget = "44" Or strTarget = "54" Or strTarget = "6" Or strTarget = "61" Then
            If strTarget = "42" Or strTarget = "43" Or strTarget = "54" Or strTarget = "70" Or strTarget = "61" Or strTarget = "6" Or strTarget = "44" Or strTarget = "40" Or strTarget = "11" Then
                If strResult = "-" Then
                    strLGrp = strLGrp
                Else
                    strLGrp = strLGrp & strTarget & strResult & ","
                End If
            Else
                If strTarget <> "IC" Then
                    If strResult = "-" Then
                        strHGrp = strHGrp
                    Else
                        strHGrp = strHGrp & strTarget & strResult & ","
                    End If
                End If
            End If
        Next
        
        If strVarTest = "에러" Then
            strResult = ""
        Else
            If Len(strHGrp) = 0 And Len(strLGrp) = 0 Then
               strTotFlag = "Negative"
            Else
               strTotFlag = "Positive"
            End If
            
            strResult = strTotFlag
            
            If Len(strHGrp) > 0 Then
                strHGrp = Mid(strHGrp, 1, Len(strHGrp) - 1)
            End If
            
            If Len(strLGrp) > 0 Then
                strLGrp = Mid(strLGrp, 1, Len(strLGrp) - 1)
            End If
            
            strFlag = ""
'            If strTotFlag = "Negative" Then
'                          strTmp = "HPV High Risk Type : Negative"
'                          strFlag = "HPV High Risk Type : Negative" & Space(104 - Len(strTmp))
'                          strTmp = "HPV Low Risk Type : Negative"
'                strFlag = strFlag & "HPV Low Risk Type : Negative" & Space(104 - Len(strTmp))
'
'            ElseIf strTotFlag = "Positive" Then
'                If Len(strHGrp) > 0 And Len(strLGrp) > 0 Then
'                              strTmp = "HPV High Risk Type : Positive (" & strHGrp & ")"
'                              strFlag = "HPV High Risk Type : Positive (" & strHGrp & ")" & Space(104 - Len(strTmp))
'                              strTmp = "HPV Low Risk Type : Positive (" & strLGrp & ")"
'                    strFlag = strFlag & "HPV Low Risk Type : Positive (" & strLGrp & ")" & Space(104 - Len(strTmp))
'                ElseIf Len(strHGrp) > 0 And Len(strLGrp) <= 0 Then
'                              strTmp = "HPV High Risk Type : Positive (" & strHGrp & ")"
'                              strFlag = "HPV High Risk Type : Positive (" & strHGrp & ")" & Space(104 - Len(strTmp))
'                              strTmp = "HPV Low Risk Type : Negative"
'                    strFlag = strFlag & "HPV Low Risk Type : Negative" & Space(104 - Len(strTmp))
'                ElseIf Len(strHGrp) <= 0 And Len(strLGrp) > 0 Then
'                              strTmp = "HPV High Risk Type : Negative"
'                              strFlag = "HPV High Risk Type : Negative" & Space(104 - Len(strTmp))
'                              strTmp = "HPV Low Risk Type : Positive (" & strLGrp & ")"
'                    strFlag = strFlag & "HPV Low Risk Type : Positive (" & strLGrp & ")" & Space(104 - Len(strTmp))
'                End If
'            End If

            If strTotFlag = "Negative" Then
                          strTmp = "HPV High Risk Type : Negative"
                          strFlag = "HPV High Risk Type : Negative" & Chr(10)
                          strTmp = "HPV Low Risk Type : Negative"
                strFlag = strFlag & "HPV Low Risk Type : Negative" & Chr(10)

            ElseIf strTotFlag = "Positive" Then
                If Len(strHGrp) > 0 And Len(strLGrp) > 0 Then
                              strFlag = "HPV High Risk Type : Positive (" & strHGrp & ")" & Chr(10)
                    strFlag = strFlag & "HPV Low Risk Type : Positive (" & strLGrp & ")" & Chr(10)
                ElseIf Len(strHGrp) > 0 And Len(strLGrp) <= 0 Then
                              strFlag = "HPV High Risk Type : Positive (" & strHGrp & ")" & Chr(10)
                    strFlag = strFlag & "HPV Low Risk Type : Negative" & Chr(10)
                ElseIf Len(strHGrp) <= 0 And Len(strLGrp) > 0 Then
                              strFlag = "HPV High Risk Type : Negative" & Chr(10)
                    strFlag = strFlag & "HPV Low Risk Type : Positive (" & strLGrp & ")" & Chr(10)
                End If
            End If
        
'            If strTotFlag = "Negative" Then
'                          strTmp = "HPV High Risk Type : Negative"
'                          strFlag = "HPV High Risk Type : Negative" & Space(60 - Len(strTmp))
'                          strTmp = "HPV Low Risk Type : Negative"
'                strFlag = strFlag & "HPV Low Risk Type : Negative" & Space(60 - Len(strTmp))
'                          strFlag = "12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890" & vbNewLine
'                strFlag = strFlag & "09876543210987654321098765432109876543210987654321098765432109876543210987654321098765432109876543210987654321098765432109876543210987654321098765432109876543210987654321098765432109876543210987654321" & vbNewLine
'            ElseIf strTotFlag = "Positive" Then
'                If Len(strHGrp) > 0 And Len(strLGrp) > 0 Then
'                              strTmp = "HPV High Risk Type : Positive (" & strHGrp & ")"
'                              strFlag = "HPV High Risk Type : Positive (" & strHGrp & ")" & Space(60 - Len(strTmp))
'                              strTmp = "HPV Low Risk Type : Positive (" & strLGrp & ")"
'                    strFlag = strFlag & "HPV Low Risk Type : Positive (" & strLGrp & ")" & Space(60 - Len(strTmp))
'                ElseIf Len(strHGrp) > 0 And Len(strLGrp) <= 0 Then
'                              strTmp = "HPV High Risk Type : Positive (" & strHGrp & ")"
'                              strFlag = "HPV High Risk Type : Positive (" & strHGrp & ")" & Space(60 - Len(strTmp))
'                              strTmp = "HPV Low Risk Type : Negative"
'                    strFlag = strFlag & "HPV Low Risk Type : Negative" & Space(60 - Len(strTmp))
'                ElseIf Len(strHGrp) <= 0 And Len(strLGrp) > 0 Then
'                              strTmp = "HPV High Risk Type : Negative"
'                              strFlag = "HPV High Risk Type : Negative" & Space(60 - Len(strTmp))
'                              strTmp = "HPV Low Risk Type : Positive (" & strLGrp & ")"
'                    strFlag = strFlag & "HPV Low Risk Type : Positive (" & strLGrp & ")" & Space(60 - Len(strTmp))
'                End If
'            End If

        End If
        
        strResult = strFlag
'        If strFlag <> "" Then
'            strResult = strResult & "(" & strFlag & ")"
'        End If
        
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
                            If Len(strHGrp) > 0 Then
                                SetText .spdOrder, strHGrp, gRow, intCol + 1
                            End If
                            If Len(strLGrp) > 0 Then
                                SetText .spdOrder, strLGrp, gRow, intCol + 1
                            End If
                            SetText .spdOrder, strFlag, gRow, intCol + 2
                            Exit For
                        End If
                    Next
    
                    '-- 결과 List
                    SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                    SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                    SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                    SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                    SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                    SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                    SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '장비결과
                    SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS결과
                    'SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                    'SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                    SetText .spdResult, strICVal1, lsRstRow, colRJUDGE                     '판정
                    SetText .spdResult, strICVal2, lsRstRow, colRREF          '참고치
                    
                    '-- 로컬 저장
                    SetLocalDB gRow, lsRstRow, "1", ""
                    
                    strState = "R"
                    
                    '-- 결과Count
'                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                        SetText .spdOrder, "1", gRow, colRCNT
'                    Else
'                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                    End If
                    
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
                            If Len(strHGrp) > 0 Then
                                SetText .spdOrder, strHGrp, gRow, intCol + 1
                            End If
                            If Len(strLGrp) > 0 Then
                                SetText .spdOrder, strLGrp, gRow, intCol + 1
                            End If
                            SetText .spdOrder, strFlag, gRow, intCol + 2
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
                    'SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                    'SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                    SetText .spdResult, strICVal1, lsRstRow, colRJUDGE                     '판정
                    SetText .spdResult, strICVal2, lsRstRow, colRREF          '참고치
                    
                    '-- 로컬 저장
                    SetLocalDB gRow, lsRstRow, "1", ""
                    
                    If strState <> "R" Then
                        strState = ""
                    End If
    
                    '-- 결과Count
'                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                        SetText .spdOrder, "1", gRow, colRCNT
'                    Else
'                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                    End If
                End If
            End If
        End If
        
        .spdResult.RowHeight(-1) = 14
        
        '## DB에 결과저장
        If .optTrans(0).Value = True And strState = "R" Then
            Res = SaveTransData(gRow)
            
            If Res = -1 Then
                '-- 저장 실패
    '                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                SetText .spdOrder, "Failed", gRow, colSTATE
            Else
                '-- 저장 성공
    '                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                SetText .spdOrder, "저장완료", gRow, colSTATE
                SetText .spdOrder, "0", gRow, colCHECKBOX
                
                      SQL = "Update PATRESULT Set " & vbCrLf
                SQL = SQL & " sendflag = '2' " & vbCrLf
                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                
    '                            Res = SendQuery(gLocal, SQL)
                If Res = -1 Then
    '                                SaveQuery SQL
                    Exit Sub
                End If
            End If
            strState = ""
        End If
    
    End With

End Sub

Public Sub SerialRcvData_E411()
   

End Sub


Function SaveTransData(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    'Dim strPatSeq       As String
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
    
    Dim intState        As Integer
    Dim arg_Acptstatcd  As String
    Dim strRsltrgstno   As String
    Dim strAcptstatcd   As String
    Dim strState        As String
    
    Dim strTmp          As String
    
    
    
    '결과등록일자,결과등록시간,병리번호,결과등록번호,등록번호,결과등록자ID,검체명(spcm.spcnm),검사명(test.testengnm)
    Dim RgtDD, RgtTM, PTNO, Reg_NO, PID, USERID, SPCNM, TESTNM As String
    
On Error GoTo ErrHandle
    
    'RgtDD = Format(Now, "yyyymmdd")
    'RgtTM = Format(Now, "hhmmss")
    
    
    With frmMain
        SaveTransData = -1
        intRow = 0
        
        If .optCheck(0).Value = True Then
            strState = "0,2"
        ElseIf .optCheck(1).Value = True Then
            strState = "0"
        ElseIf .optCheck(2).Value = True Then
            strState = "2"
        End If

        Call GetSampleInfo_Save(argSpcRow, .spdOrder, "", strState)
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        
        strRsltrgstno = Trim(GetText(.spdOrder, argSpcRow, colOCNT))
        strAcptstatcd = Trim(GetText(.spdOrder, argSpcRow, colRCNT))
        
        'strPatSeq = strBarcode
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
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
        
        'AdoCn.BeginTrans
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '검사코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            'Step 1
            If GetSysDtTm Then
                'Step 2
                SQL = ""
                SQL = SQL & "SELECT acptstatcd          " & vbCr
                SQL = SQL & "  From lis.lpjmacpt        " & vbCr
                SQL = SQL & " WHERE instcd         = ?  " & vbCr
                SQL = SQL & "   AND ptno           = ?  " & vbCr
                SQL = SQL & "   AND acptdd         = ?  " & vbCr
                SQL = SQL & "   AND acptno         = ?  " & vbCr
                SQL = SQL & "   AND acptitemno     = ?  " & vbCr
                SQL = SQL & "   AND testcd         = ?  " & vbCr
                SQL = SQL & "   AND pid            = ?  " & vbCr
                'SQL = SQL & "   AND prcpdd         = ?  " & vbCr
                'SQL = SQL & "   AND execprcpuniqno = ?  "
                
                Call SetSQLData("결과저장", SQL, "A")
                
                Set AdoCmd = New ADODB.Command
                Set AdoCmd.ActiveConnection = AdoCn
                
                AdoCmd.CommandType = adCmdText
                AdoCmd.CommandText = SQL
                AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
                AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, strChartNo)
                AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptdd", adVarChar, adParamInput, 1000, strHospDate)
                AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptno", adVarChar, adParamInput, 1000, strBarcode)
                AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptitemno", adVarChar, adParamInput, 1000, 1)
                AdoCmd.Parameters.Append AdoCmd.CreateParameter("testcd", adVarChar, adParamInput, 1000, gAllTestCd1)
                AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, strPatID)
                'AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
                'AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
                
                strTmp = ""
                strTmp = strTmp & "instcd" & ":" & gHOSP.HOSPCD & vbCr
                strTmp = strTmp & "ptno" & ":" & strChartNo & vbCr
                strTmp = strTmp & "acptdd" & ":" & strHospDate & vbCr
                strTmp = strTmp & "acptno" & ":" & strBarcode & vbCr
                strTmp = strTmp & "acptitemno" & ":" & "1" & vbCr
                strTmp = strTmp & "testcd" & ":" & gAllTestCd1 & vbCr
                strTmp = strTmp & "pid" & ":" & strPatID & vbCr
                
                Call SetSQLData("결과저장", strTmp, "A")
                
                Set AdoRs = New ADODB.Recordset
                AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic
            
                If AdoRs.BOF = False Then
                    arg_Acptstatcd = AdoRs.Fields("acptstatcd").Value & ""
                End If
                
                Set AdoCmd = Nothing
                Set AdoRs = Nothing
                
                
                'param=[2, 10602673, 017, M17003176, 20170724, 33978, 1, PMO12040, 17488137, 20170724, 1151787391]|1 records|
                If arg_Acptstatcd = "0" Then
                    'Step 3 : 체번하기
                    strRsltrgstno = GetLastSeqNum
                    
                    If strRsltrgstno = "" Then
                        'Step 4 : 재저장,       insert
                        Call Regist_Result_Step2
                    Else
                        'Step 4 : 최초저장시,   update
                        Call Regist_Result_Step3(strRsltrgstno)
                    End If
                
                '-- 재입력
                ElseIf arg_Acptstatcd = "2" Then
                
                    strRsltrgstno = strRsltrgstno
                
                End If
                
                'Step 2-2
                If Not Regist_Result_Step1 Then
                    '저장실패
                End If
                
                '-- Step 5 결과저장(Header)
                If Regist_Result_Header(strChartNo, strHospDate, strRsltrgstno, strPatID, sResult) = 1 Then
                
                    '-- Step 6 결과저장(Detail)
                    If Regist_Result_Detail(strChartNo, strHospDate, strRsltrgstno, strPatID, sResult) = 1 Then
                    
                        '--Step 7 결과저장(T/M/P)
                        If Regist_Result_TMP(strChartNo, strHospDate, strRsltrgstno, strPatID, sResult) = 1 Then
                        
                            '--Step 8 결과저장(병리번호 발행이력 수정)
                            If Regist_Result_HIS(strChartNo, strHospDate, strRsltrgstno, strPatID, sResult) = 1 Then
                            
                                '--Step 9 결과저장(병리과 접수정보 수정)
                                If Regist_Result_RCPEDIT(strChartNo, strHospDate, strRsltrgstno, strPatID, sResult) = 1 Then
                                    
                                    intState = GetINOUT(strChartNo, strHospDate, strRsltrgstno, strPatID, sResult, strBarcode, strAcptstatcd)
                                    
                                    If intState = 0 Then
                                        SaveTransData = -1
                                    ElseIf intState = 1 Then
                                        SaveTransData = 1
                                    Else
                                        SaveTransData = -1
                                    End If
                                    
                                        '--Step 10 결과저장(Detail)
                                      '  If Regist_Result_Detail2(strChartNo, strHospDate, strRsltrgstno, strPatID, sResult, strBarcode) = 1 Then

                                            'If GetINOUT2(strChartNo, strHospDate, strRsltrgstno, strPatID, sResult, strBarcode) = 1 Then

'                                                SaveTransData = 1
                                                
                                            'End If
                                            
                                        'End If
                                   ' End If
                                End If
        
                            End If
                        End If
                    End If
                End If
                
                SaveTransData = 1
            
            End If
            

'            Call Regist_Result1(RgtDD, PTNO, Reg_NO, PID, USERID, F_RSLT)
'            Call Regist_Result_Item(PTNO, RgtDD, Reg_NO, RISKFLAGCD, ACPTDD, ACPTNO, TESTCD, ACPTITEMNO, F_RSLT, USERID)
'            Call Regist_Slide(RgtDD, RgtTM, USERID, PTNO)
'            Call Regist_Status_Insert710(PRCPDD, EXECPRCPUNIQNO, RgtDD, RgtTM, USERID)
'            Call Regist_Status_Insert700(PRCPDD, EXECPRCPUNIQNO, RgtDD, RgtTM, USERID)
'            Call Regist_Status_Update(Check_Kubun, USERID, PID, PRCPDD, EXECPRCPUNIQNO, PRCPNO)
'            Call Regist_Status_Update_ACPT(ACPTDD, ACPTNO, ACPTITEMNO, PTNO, PID)
        
'            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
'                      SQL = " Update SLA_LabResult  " & vbCr
'                SQL = SQL & "   Set Result     = '" & sResult & "' " & vbCr
'                SQL = SQL & "      ,NormalFlag = '0' " & vbCr
'                SQL = SQL & "      ,PanicFlag  = '0' " & vbCr
'                SQL = SQL & "      ,DeltaFlag  = '0' " & vbCr
'                SQL = SQL & "      ,TransFlag  = '1' " & vbCr
'                SQL = SQL & "      ,ResultID   = ''  " & vbCr
'                SQL = SQL & "      ,ResultDate = '" & Trim(Format(Now, "yyyy-mm-dd")) & "'" & vbCr
'                SQL = SQL & "      ,ResultTime = '" & Trim(Format(Time, "HH:MM:SS")) & "'" & vbCr
'                SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "'" & vbCr
'                SQL = SQL & "   And OrderCode = '" & gAllOrdCd & "'" & vbCr
'                SQL = SQL & "   And LabCode = '" & strTestCd & "'" & vbCr
'                SQL = SQL & "   And TransFlag < '2' "
'
'                Call SetSQLData("결과저장", SQL)
'                Call DBExec(AdoCn, SQL)
'
'                SaveTransData = 1
'
'            End If
        Next intRow
        
    End With

Exit Function

ErrHandle:
    SaveTransData = -1
    'AdoCn.RollbackTrans
    
End Function


'/*  결과순번 일련번호 채번을 위해 기준 Row를 Lock처리한다.
'himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml setseqnolock
'arg_seqflagcd = '4'  순번발생구분코드(검사결과) 고정입니다.
'param=[017, 4] */
Private Function Regist_Result_Step1() As Boolean
Dim strTmp  As String

On Error GoTo DBErr
    Regist_Result_Step1 = False
    
    SQL = ""
    SQL = SQL & "update lis.lpcmseqn" & vbCr
    SQL = SQL & "   Set lastgenrno = 1 " & vbCr
    SQL = SQL & " where instcd     = ? " & vbCr
    SQL = SQL & "   and seqgenryy  = '1900' " & vbCr
    SQL = SQL & "   and seqflagcd  = ? " & vbCr
    
    Call SetSQLData("결과저장", SQL, "A")
    
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    
    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("seqflagcd", adVarChar, adParamInput, 1000, "4")
    AdoCmd.Execute
    
    strTmp = ""
    strTmp = strTmp & "instcd" & ":" & gHOSP.HOSPCD & vbCr
    strTmp = strTmp & "seqflagcd" & ":" & "4" & vbCr
    
    Call SetSQLData("결과저장", strTmp, "A")
    
    Set AdoCmd = Nothing
    
    Regist_Result_Step1 = True
    
Exit Function

DBErr:
    Regist_Result_Step1 = False
    
End Function

'/*위에서 채번을 했는데 null 일경우 insert하고 lastgenrno = 1로 리턴 1부터 시작합니다.
'년도마다 새로 입력이 됩니다
Private Function Regist_Result_Step2() As Boolean

On Error GoTo DBErr
    Regist_Result_Step2 = False
    
    SQL = ""
    SQL = SQL & "INSERT INTO lis.lpcmseqn (seqgenryy, seqflagcd,  instcd,     lastgenrno," & vbCr
    SQL = SQL & "                          fstrgstdt, fstrgstrid, lastupdtdt, lastupdtrid)" & vbCr
    SQL = SQL & "                  VALUES (?, ?, ?, 1, SYSDATE, ?,  SYSDATE, ?)"
    
    Call SetSQLData("결과저장", SQL, "A")
    
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    
    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("seqgenryy", adVarChar, adParamInput, 1000, Format(Now, "yyyy"))
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("seqflagcd", adVarChar, adParamInput, 1000, "4")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
    AdoCmd.Execute
    
    Set AdoCmd = Nothing
    
    Regist_Result_Step2 = True
    
Exit Function

DBErr:
    Regist_Result_Step2 = False
    
End Function

Private Function Regist_Result_Step3(ByVal LastSeqNum As String) As Boolean

On Error GoTo DBErr
    Regist_Result_Step3 = False
    
    SQL = ""
    SQL = SQL & "Update lis.lpcmseqn" & vbCr
    SQL = SQL & "   set lastgenrno  = ? " & vbCr
    SQL = SQL & "      ,lastupdtdt  = sysdate" & vbCr
    SQL = SQL & "      ,lastupdtrid = ? " & vbCr
    SQL = SQL & " where instcd      = ? " & vbCr
    SQL = SQL & "   and seqgenryy   = ? " & vbCr
    SQL = SQL & "   and seqflagcd   = ? "
    
    Call SetSQLData("결과저장", SQL, "A")
    
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    
    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastgenrno", adVarChar, adParamInput, 1000, LastSeqNum)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("seqgenryy", adVarChar, adParamInput, 1000, Format(Now, "yyyy"))
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("seqflagcd", adVarChar, adParamInput, 1000, "4")
    
    AdoCmd.Execute
    
    Set AdoCmd = Nothing
    
    Regist_Result_Step3 = True
    
Exit Function

DBErr:
    Regist_Result_Step3 = False
    
End Function


'/* 결과순번 채번을 합니다. 여기서 조회된 lastgenrno가 뒤쪽에 rsltrgstno 로 처리됩니다.
'himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getlastseqno
'param=[017, 2017, 4]*/
Private Function GetLastSeqNum() As String
Dim strTmp  As String

On Error GoTo DBErr
    GetLastSeqNum = ""
    
    SQL = ""
    SQL = SQL & "select coalesce(lastgenrno+1, 1) as lastgenrno" & vbCr
    SQL = SQL & "  From lis.lpcmseqn" & vbCr
    SQL = SQL & " where instcd     = ? " & vbCr
    SQL = SQL & "   and seqgenryy  = ? " & vbCr
    SQL = SQL & "   and seqflagcd  = ? " & vbCr
    
    Call SetSQLData("결과저장", SQL, "A")
    
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    
    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("seqgenryy", adVarChar, adParamInput, 1000, Format(Now, "yyyy"))
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("seqflagcd", adVarChar, adParamInput, 1000, "4")
    
    strTmp = ""
    strTmp = strTmp & "instcd" & ":" & gHOSP.HOSPCD & vbCr
    strTmp = strTmp & "seqgenryy" & ":" & Format(Now, "yyyy") & vbCr
    strTmp = strTmp & "seqflagcd" & ":" & "4" & vbCr
    
    Call SetSQLData("결과저장", strTmp, "A")

    
    Set AdoRs = New ADODB.Recordset
    AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic

    If AdoRs.BOF = False Then
        GetLastSeqNum = AdoRs.Fields("lastgenrno").Value & ""
    Else
        GetLastSeqNum = ""
    End If
    
    Set AdoCmd = Nothing
    Set AdoRs = Nothing
    
Exit Function

DBErr:
    GetLastSeqNum = ""
    
End Function

'/* himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml getptnoprcpinfo
'param=[53, 53, 53,
'53, 53, 53,
'53, 53, 53,
'53, 53,
'53,
'017,
'M17003176,
'17488137]
'*/
Private Function GetINOUT(ByVal pPTNO As String, ByVal pREGDT As String, ByVal pPTSEQ As String, ByVal pPID As String, ByVal pRESULT As String, Optional ByVal pACPTNO As String, Optional ByVal pACPTSTATCD As String) As Integer

    Dim arg_Prcpdd              As String
    Dim arg_Execprcpuniqno      As String
    Dim arg_Prcpgenrflag        As String
    Dim arg_Flagcd              As String
    Dim arg_Tretdd              As String
    Dim arg_Trettm              As String
    Dim arg_Acptstatcd          As String
    Dim arg_Acptststcnt         As String
    Dim arg_Ptnocd              As String
    Dim arg_Prcpstatcd          As String
    Dim arg_Tretcnt             As String
    
    Dim strRsltCmt              As String
    
    '-- 저장
'                 strRsltCmt = "[Methods]" & Chr(10) 'vbNewLine
'    strRsltCmt = strRsltCmt & "Seegene HPV Real-Time PCR (Anyplex II HPV 28 Detection Real-time PCR)" & Chr(10) 'vbNewLine
'    strRsltCmt = strRsltCmt & Chr(10) 'vbNewLine
'    strRsltCmt = strRsltCmt & "[Result]" & Chr(10) 'vbNewLine
'    strRsltCmt = strRsltCmt & pRESULT
    
                 strRsltCmt = "[Result]" & Chr(10) 'vbNewLine
    strRsltCmt = strRsltCmt & pRESULT
    
    arg_Flagcd = "710"  '710 : 처방상태 예비결과 (저장된상태)
    
On Error GoTo DBErr
    
    GetINOUT = -1
    
    '53, 53, 53,
    '53, 53, 53,
    '53, 53, 53,
    '53, 53, 53,
    '017, M17003176, 17488137]|1 records|
    '-- 1 st Query
    SQL = ""
    SQL = SQL & "SELECT acpt.instcd, acpt.prcpdd, acpt.pid, acpt.prcpno, acpt.execprcpuniqno," & vbCr
    SQL = SQL & "       MIN(acpt.prcpgenrflag) AS prcpgenrflag,  'I' AS biztretflagcd," & vbCr
    SQL = SQL & "       Case WHEN ''||53||'' = '32' THEN '700'" & vbCr
    SQL = SQL & "            WHEN ''||53||'' = '52' THEN '740'" & vbCr
    SQL = SQL & "            WHEN ''||53||'' = '53' THEN '700'" & vbCr
    SQL = SQL & "       END AS newprcpstatcd," & vbCr
    SQL = SQL & "       Case WHEN ''||53||'' = '32' THEN '700'" & vbCr
    SQL = SQL & "            WHEN ''||53||'' = '52' THEN '740'" & vbCr
    SQL = SQL & "            WHEN ''||53||'' = '53' THEN '700'" & vbCr
    SQL = SQL & "       END AS bizflagcd," & vbCr
    SQL = SQL & "       Case WHEN ''||53||'' = '32' THEN '700'" & vbCr
    SQL = SQL & "            WHEN ''||53||'' = '52' THEN '740'" & vbCr
    SQL = SQL & "            WHEN ''||53||'' = '53' THEN '700'" & vbCr
    SQL = SQL & "       END AS tretflagcd," & vbCr
    SQL = SQL & "       CASE WHEN 53 = '32' THEN MAX(pnis.makeenddd) ELSE TO_CHAR(SYSDATE,'YYYYMMDD') END AS tretdd," & vbCr
    SQL = SQL & "       CASE WHEN 53 = '32' THEN MAX(pnis.makeendtm) ELSE TO_CHAR(SYSDATE,'HH24MISS') END AS trettm," & vbCr
    SQL = SQL & "       53 AS scrno" & vbCr
    SQL = SQL & "     , acpt.prcpgenrflag AS prcpgenrflagcd" & vbCr
    SQL = SQL & "  FROM lis.lpjmacpt acpt" & vbCr
    SQL = SQL & "     , lis.lpcmpnis pnis" & vbCr
    SQL = SQL & " WHERE acpt.instcd      = ?" & vbCr
    SQL = SQL & "   AND acpt.ptno        = ?" & vbCr
    SQL = SQL & "   AND acpt.pid         = ?" & vbCr
    SQL = SQL & "   AND acpt.acptstatcd IN ('0', '2', '3', '4', '9')" & vbCr
    SQL = SQL & "   AND acpt.instcd      = pnis.instcd" & vbCr
    SQL = SQL & "   AND acpt.ptno        = pnis.ptno" & vbCr
    SQL = SQL & "   AND pnis.delflagcd   = '0'" & vbCr
    SQL = SQL & "   AND acpt.acptstatcd  = ? " & vbCr
    SQL = SQL & " GROUP BY acpt.instcd, acpt.prcpdd, acpt.pid, acpt.prcpno," & vbCr
    SQL = SQL & "          acpt.execprcpuniqno , acpt.prcpgenrflag" & vbCr
    
    
    Call SetSQLData("결과저장", SQL, "A")
    
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    
    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptstatcd", adVarChar, adParamInput, 1000, pACPTSTATCD)
    
    Set AdoRs = New ADODB.Recordset
    AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic

    If AdoRs.BOF = False Then
        arg_Prcpdd = AdoRs.Fields("prcpdd").Value & ""
        arg_Execprcpuniqno = AdoRs.Fields("execprcpuniqno").Value & ""
        arg_Prcpgenrflag = AdoRs.Fields("prcpgenrflag").Value & ""
        arg_Tretdd = AdoRs.Fields("tretdd").Value & ""
        arg_Trettm = AdoRs.Fields("trettm").Value & ""
    Else
        MsgBox "해당 처방의 정보가 변경되었습니다. 확인바랍니다.", vbCritical + vbOKOnly, "처방정보변경"
        AdoCn.RollbackTrans
        Set AdoCmd = Nothing
        Set AdoRs = Nothing
        Exit Function
    End If
    
    Set AdoCmd = Nothing
    Set AdoRs = Nothing
    
    
'    Dim arg_Prcpstatcd      As String
    
    '-- 2 nd Query
    SQL = ""
    SQL = SQL & "select prcpstatcd" & vbCr
    SQL = SQL & "  from (" & vbCr
    SQL = SQL & "       SELECT b.prcpstatcd" & vbCr
    SQL = SQL & "         FROM emr.mmodexip a, emr.mmohiprc b " & vbCr '-- 입원
    SQL = SQL & "        WHERE a.instcd         = ? " & vbCr
    SQL = SQL & "          AND a.pid            = ? " & vbCr
    SQL = SQL & "          AND a.prcpdd         = ? " & vbCr
    SQL = SQL & "          AND a.execprcpuniqno = ? " & vbCr
    SQL = SQL & "          AND a.execprcphistcd = 'O'" & vbCr
    SQL = SQL & "          AND a.instcd         = b.instcd" & vbCr
    SQL = SQL & "          AND a.pid            = b.pid" & vbCr
    SQL = SQL & "          AND a.prcpdd         = b.prcpdd" & vbCr
    SQL = SQL & "          AND a.prcpno         = b.prcpno" & vbCr
    SQL = SQL & "          AND a.prcphistno     = b.prcphistno" & vbCr
    SQL = SQL & "          AND b.prcphistcd     = 'O'" & vbCr
    SQL = SQL & "          AND b.prcpclscd      = 'D2'" & vbCr
    SQL = SQL & "          AND b.tempprcpflag   = 'N'" & vbCr
    SQL = SQL & "        Union All" & vbCr
    SQL = SQL & "       SELECT b.prcpstatcd" & vbCr
    SQL = SQL & "         FROM emr.mmodexop a, emr.mmohoprc b " & vbCr   '-- 외래
    SQL = SQL & "        WHERE a.instcd         = ? " & vbCr
    SQL = SQL & "          AND a.pid            = ? " & vbCr
    SQL = SQL & "          AND a.prcpdd         = ? " & vbCr
    SQL = SQL & "          AND a.execprcpuniqno = ? " & vbCr
    SQL = SQL & "          AND a.execprcphistcd = 'O'" & vbCr
    SQL = SQL & "          AND a.instcd         = b.instcd" & vbCr
    SQL = SQL & "          AND a.pid            = b.pid" & vbCr
    SQL = SQL & "          AND a.prcpdd         = b.prcpdd" & vbCr
    SQL = SQL & "          AND a.prcpno         = b.prcpno" & vbCr
    SQL = SQL & "          AND a.prcphistno     = b.prcphistno" & vbCr
    SQL = SQL & "          AND b.prcphistcd     = 'O'" & vbCr
    SQL = SQL & "          AND b.prcpclscd      = 'D2'" & vbCr
    SQL = SQL & "          AND b.tempprcpflag   = 'N' )" & vbCr
    SQL = SQL & " Where rownum = 1"
    
    Call SetSQLData("결과저장", SQL, "A")
    
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    
    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
    
    Set AdoRs = New ADODB.Recordset
    AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic

    If AdoRs.BOF = False Then
        arg_Prcpstatcd = AdoRs.Fields("prcpstatcd").Value & ""
    End If
    
    Set AdoCmd = Nothing
    Set AdoRs = Nothing
    
    'param=[710, 10602673, 017, 17488137, 20170724, 1151787391]
    '-- 3 rd Query
    '--입원일경우 prcpgenrflag = I, D, E
    
    'I:  입원 , E: 응급 , D: DSC
    'emr.mmohiprc  -- 입원 의사처방
    'emr.mmodexip  -- 입원 진료지원부서 정보
    
    If arg_Prcpgenrflag = "I" Or arg_Prcpgenrflag = "D" Or arg_Prcpgenrflag = "E" Then
        SQL = ""
        SQL = SQL & "Update emr.mmohiprc"
        SQL = SQL & "   SET prcpstatcd  = ? ,"
        SQL = SQL & "       lastupdtdt  = SYSDATE,"
        SQL = SQL & "       lastupdtrid = ? "
        SQL = SQL & " WHERE (instcd, pid, prcpdd, prcpno, prcphistno) IN"
        SQL = SQL & "       (SELECT instcd, pid, prcpdd, prcpno, prcphistno"
        SQL = SQL & "          From emr.mmodexip"
        SQL = SQL & "         WHERE instcd         = ? "
        SQL = SQL & "           AND pid            = ? "
        SQL = SQL & "           AND prcpdd         = ? "
        SQL = SQL & "           AND execprcpuniqno = ? "
        SQL = SQL & "           AND execprcphistcd = 'O'"
        SQL = SQL & "       )"
        SQL = SQL & "   AND prcphistcd   = 'O'"
        SQL = SQL & "   AND prcpclscd    = 'D2'"
        SQL = SQL & "   AND tempprcpflag = 'N'"
    
        Call SetSQLData("결과저장", SQL, "A")
        
        Set AdoCmd = New ADODB.Command
        Set AdoCmd.ActiveConnection = AdoCn
        
        AdoCmd.CommandType = adCmdText
        AdoCmd.CommandText = SQL
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpstatcd", adVarChar, adParamInput, 1000, arg_Flagcd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
        
        AdoCmd.Execute
        
    '--외래일경우 prcpgenrflag = O, S
    
    'O: 외래 , S: 건진
    'emr.mmohipr -- 외래 의사처방
    'emr.mmodexop -- 외래 진료지원부서 정보
    
    ElseIf arg_Prcpgenrflag = "O" Or arg_Prcpgenrflag = "S" Then
        SQL = ""
        SQL = SQL & "Update emr.mmohoprc"
        SQL = SQL & "   SET prcpstatcd  = ? ,"
        SQL = SQL & "       lastupdtdt  = SYSDATE,"
        SQL = SQL & "       lastupdtrid = ? "
        SQL = SQL & " WHERE (instcd, pid, prcpdd, prcpno, prcphistno) IN"
        SQL = SQL & "       (SELECT instcd, pid, prcpdd, prcpno, prcphistno"
        SQL = SQL & "          From emr.mmodexop"
        SQL = SQL & "         WHERE instcd         = ? "
        SQL = SQL & "           AND pid            = ? "
        SQL = SQL & "           AND prcpdd         = ? "
        SQL = SQL & "           AND execprcpuniqno = ? "
        SQL = SQL & "           AND execprcphistcd = 'O'"
        SQL = SQL & "       )"
        SQL = SQL & "   AND prcphistcd   = 'O'"
        SQL = SQL & "   AND prcpclscd    = 'D2'"
        SQL = SQL & "   AND tempprcpflag = 'N'"
    
        
        Call SetSQLData("결과저장", SQL, "A")
        
        Set AdoCmd = New ADODB.Command
        Set AdoCmd.ActiveConnection = AdoCn
        
        AdoCmd.CommandType = adCmdText
        AdoCmd.CommandText = SQL
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpstatcd", adVarChar, adParamInput, 1000, arg_Flagcd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
        
        AdoCmd.Execute
    
    End If
    
    Set AdoCmd = Nothing


    '-- 4 st Query
    '--입원일경우 prcpgenrflag = I, D, E
    If arg_Prcpgenrflag = "I" Or arg_Prcpgenrflag = "D" Or arg_Prcpgenrflag = "E" Then
        SQL = ""
        SQL = SQL & "Update emr.mmodexip a"
        SQL = SQL & "  SET a.execprcpstatcd = ? "
        SQL = SQL & "     ,a.lastupdtdt     = SYSDATE "
        SQL = SQL & "     ,a.lastupdtrid    = ? "
        SQL = SQL & "WHERE a.instcd         = ? "
        SQL = SQL & "  AND a.pid            = ? "
        SQL = SQL & "  AND a.prcpdd         = ? "
        SQL = SQL & "  AND a.execprcpuniqno = ? "
        SQL = SQL & "  AND a.execprcphistcd = 'O'"
               
        Call SetSQLData("결과저장", SQL, "A")
        
        Set AdoCmd = New ADODB.Command
        Set AdoCmd.ActiveConnection = AdoCn
        
        AdoCmd.CommandType = adCmdText
        AdoCmd.CommandText = SQL
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpstatcd", adVarChar, adParamInput, 1000, arg_Flagcd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
        
        AdoCmd.Execute
        
    '--외래일경우 prcpgenrflag = O, S
    ElseIf arg_Prcpgenrflag = "O" Or arg_Prcpgenrflag = "S" Then
        SQL = ""
        SQL = SQL & "Update emr.mmodexop a"
        SQL = SQL & "  SET a.execprcpstatcd = ? "
        SQL = SQL & "     ,a.lastupdtdt     = SYSDATE "
        SQL = SQL & "     ,a.lastupdtrid    = ? "
        SQL = SQL & "WHERE a.instcd         = ? "
        SQL = SQL & "  AND a.pid            = ? "
        SQL = SQL & "  AND a.prcpdd         = ? "
        SQL = SQL & "  AND a.execprcpuniqno = ? "
        SQL = SQL & "  AND a.execprcphistcd = 'O'"
               
        Call SetSQLData("결과저장", SQL, "A")
        
        Set AdoCmd = New ADODB.Command
        Set AdoCmd.ActiveConnection = AdoCn
        
        AdoCmd.CommandType = adCmdText
        AdoCmd.CommandText = SQL
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpstatcd", adVarChar, adParamInput, 1000, arg_Flagcd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
        
        AdoCmd.Execute
    
    End If
    
    Set AdoCmd = Nothing


    '-- 5 st Query
'    Dim arg_Tretcnt         As Long
    
    SQL = ""
    SQL = SQL & "SELECT COUNT(prcpdd) AS tretcnt"
    SQL = SQL & "  From emr.mmodexpt"
    SQL = SQL & " WHERE instcd         = ? "
    SQL = SQL & "   AND execprcpuniqno = ? "
    SQL = SQL & "   AND prcpdd         = ? "
    SQL = SQL & "   AND tretflagcd     = ? "
    
    Call SetSQLData("결과저장", SQL, "A")

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    
    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("tretflagcd", adVarChar, adParamInput, 1000, arg_Flagcd)
    
    Set AdoRs = New ADODB.Recordset
    AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic

    If AdoRs.BOF = False Then
        arg_Tretcnt = AdoRs.Fields("tretcnt").Value
    Else
        arg_Tretcnt = 0
    End If
    
    Set AdoCmd = Nothing
    Set AdoRs = Nothing
    
    '-- 6 st Query '조회안되면 ... insert
    If arg_Tretcnt <= 0 Then
        SQL = ""
        SQL = SQL & "INSERT INTO emr.mmodexpt ("
        SQL = SQL & "prcpdd,       execprcpuniqno,"
        SQL = SQL & "tretflagcd,   instcd,"
        SQL = SQL & "tretdd,       trettm,    tretpsnid, fstrgstrid,   fstrgstdt,"
        SQL = SQL & "lastupdtrid,  lastupdtdt)"
        SQL = SQL & "VALUES ("
        SQL = SQL & " ?,  CAST(? AS INTEGER),"
        SQL = SQL & " ?,  ?,"
        SQL = SQL & " ?,  ?,  CASE WHEN ? IS NULL THEN ? ELSE ? END, "
        SQL = SQL & " ?,   SYSDATE,"
        SQL = SQL & " ?,   SYSDATE)"
        
        Call SetSQLData("결과저장", SQL, "A")
        
        Set AdoCmd = New ADODB.Command
        Set AdoCmd.ActiveConnection = AdoCn
        
        AdoCmd.CommandType = adCmdText
        AdoCmd.CommandText = SQL
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
        
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("tretflagcd", adVarChar, adParamInput, 1000, arg_Flagcd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
        
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("tretdd", adVarChar, adParamInput, 1000, arg_Tretdd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("trettm", adVarChar, adParamInput, 1000, arg_Trettm)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("tretpsnid", adVarChar, adParamInput, 1000, "")
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstdt", adVarChar, adParamInput, 1000, "")
        
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
        
        AdoCmd.Execute
    
        Set AdoCmd = Nothing
    
    End If
    
    '-- 7
    'param=[017, M17003176, 20170724, 32715, 9, null, [Methods]
    '   Seegene HPV Real-time PCR (Anyplex II HPV 28 Detection Real-time PCR)
    '
    '[Result]
    'HPV High Risk Type : POSITIVE (18+, 68+, 31+++)
    'HPV Low  Risk Type : POSITIVE (70+, 61+) , 10602673, 017, M17003176, 20170724, 32715, 9, PMO12040]|1 records|
    '-- 이력기록
    SQL = ""
    SQL = SQL & "INSERT INTO lis.lprmtrlt" & vbCr
    SQL = SQL & "       (ptno,       rsltrgstdd, rsltrgstno, rsltrgsthistno, riskflagcd, instcd," & vbCr
    SQL = SQL & "        acptdd,     acptno,     testcd,     acptitemno,     testrslt,  testrsltxml,  testrsltetc,  delflagcd," & vbCr
    SQL = SQL & "        fstrgstdt,  fstrgstrid," & vbCr
    SQL = SQL & "        lastupdtdt, lastupdtrid)" & vbCr
    SQL = SQL & " SELECT ptno,       rsltrgstdd, rsltrgstno," & vbCr
    SQL = SQL & "        (SELECT MAX(z.rsltrgsthistno)+1 FROM lis.lprmtrlt z" & vbCr
    SQL = SQL & "          WHERE z.instcd     = ?" & vbCr
    SQL = SQL & "            AND z.ptno       = ?" & vbCr
    SQL = SQL & "            AND z.rsltrgstdd = ?" & vbCr
    SQL = SQL & "            AND z.rsltrgstno = CAST(? AS DECIMAL(12,0))" & vbCr
    SQL = SQL & "            AND z.riskflagcd = ?" & vbCr
    SQL = SQL & "        )," & vbCr
    SQL = SQL & "        riskflagcd, instcd," & vbCr
    SQL = SQL & "        acptdd,     acptno,     testcd,     acptitemno, decode(nvl(?,'IN'),'SMLPU00700',?,testrslt), testrsltxml,  testrsltetc,  '1'," & vbCr
    SQL = SQL & "        fstrgstdt,  fstrgstrid," & vbCr
    SQL = SQL & "        SYSDATE,  ''||?||''" & vbCr
    SQL = SQL & "   From lis.lprmtrlt" & vbCr
    SQL = SQL & "  WHERE instcd         = ?" & vbCr
    SQL = SQL & "    AND ptno           = ?" & vbCr
    SQL = SQL & "    AND rsltrgstdd     = ?" & vbCr
    SQL = SQL & "    AND rsltrgstno     = CAST(? AS DECIMAL(12,0))" & vbCr
    SQL = SQL & "    AND riskflagcd     = ?" & vbCr
    SQL = SQL & "    AND testcd         = ?" & vbCr
    SQL = SQL & "    AND rsltrgsthistno = 1" & vbCr
                
    Call SetSQLData("결과저장", SQL, "A")

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstdd", adVarChar, adParamInput, 1000, pREGDT)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstno", adVarChar, adParamInput, 1000, pPTSEQ)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("riskflagcd", adVarChar, adParamInput, 1000, 9)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrslt", adVarChar, adParamInput, 1000, "")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrslt", adVarChar, adParamInput, 1000, strRsltCmt)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)

    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstdd", adVarChar, adParamInput, 1000, pREGDT)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstno", adVarChar, adParamInput, 1000, pPTSEQ)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("riskflagcd", adVarChar, adParamInput, 1000, 9)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testcd", adVarChar, adParamInput, 1000, gAllTestCd1)

    AdoCmd.Execute

    Set AdoCmd = Nothing
                    
    '-- 삭제
    SQL = ""
    SQL = SQL & "DELETE FROM lis.lprmtrlt" & vbCr
    SQL = SQL & " WHERE instcd         = ?" & vbCr
    SQL = SQL & "   AND ptno           = ?" & vbCr
    SQL = SQL & "   AND rsltrgstdd     = ?" & vbCr
    SQL = SQL & "   AND rsltrgstno     = CAST(? AS DECIMAL(12,0))" & vbCr
    SQL = SQL & "   AND rsltrgsthistno = 1" & vbCr
    SQL = SQL & "   AND riskflagcd     = ?" & vbCr
    SQL = SQL & "   AND testcd         = ?" & vbCr
    
    Call SetSQLData("결과저장", SQL, "A")

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstdd", adVarChar, adParamInput, 1000, pREGDT)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstno", adVarChar, adParamInput, 1000, pPTSEQ)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("riskflagcd", adVarChar, adParamInput, 1000, 9)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testcd", adVarChar, adParamInput, 1000, gAllTestCd1)
    
    AdoCmd.Execute

    Set AdoCmd = Nothing
    


    SQL = ""
    SQL = SQL & "INSERT INTO lis.lprmtrlt"
    SQL = SQL & "       (ptno,         rsltrgstdd,      rsltrgstno,           rsltrgsthistno,"
    SQL = SQL & "       riskflagcd,    instcd,"
    SQL = SQL & "       acptdd,        acptno,          testcd,"
    SQL = SQL & "       acptitemno,    testrslt,        testrsltxml,  testrsltetc, delflagcd,"
    SQL = SQL & "       fstrgstdt,     fstrgstrid,"
    SQL = SQL & "       lastupdtdt,    lastupdtrid)"
    SQL = SQL & "VALUES (?,  ?, CAST(? AS DECIMAL(12,0)), 1,"
    SQL = SQL & "        ?,  ?,"
    SQL = SQL & "        ?,  CAST(? AS DECIMAL(12,0)),   ?,"
    SQL = SQL & "        CAST(? AS SMALLINT),  ?,     ?  , ?  ,   '0',"
    SQL = SQL & "        SYSDATE, ?,"
    SQL = SQL & "        SYSDATE, ?)"

    Call SetSQLData("결과저장", SQL, "A")

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstdd", adVarChar, adParamInput, 1000, pREGDT)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstno", adVarChar, adParamInput, 1000, pPTSEQ)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("riskflagcd", adVarChar, adParamInput, 1000, 9)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)

    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptdd", adVarChar, adParamInput, 1000, pREGDT)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptno", adVarChar, adParamInput, 1000, pACPTNO)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testcd", adVarChar, adParamInput, 1000, gAllTestCd1)

    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptitemno", adVarChar, adParamInput, 1000, 1)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrslt", adVarChar, adParamInput, 4000, strRsltCmt)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrsltxml", adVarChar, adParamInput, 1000, "")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrsltetc", adVarChar, adParamInput, 1000, "")

    AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)

SQL = SQL & "ptno:" & pPTNO & vbCr
SQL = SQL & "rsltrgstdd:" & pREGDT & vbCr
SQL = SQL & "rsltrgstno:" & pPTSEQ & vbCr
SQL = SQL & "riskflagcd:" & 9 & vbCr
SQL = SQL & "instcd:" & gHOSP.HOSPCD & vbCr

SQL = SQL & "acptdd:" & pREGDT & vbCr
SQL = SQL & "acptno:" & pACPTNO & vbCr
SQL = SQL & "testcd:" & gAllTestCd1 & vbCr

SQL = SQL & "acptitemno:" & 1 & vbCr
SQL = SQL & "testrslt:" & strRsltCmt & vbCr
SQL = SQL & "testrsltxml:" & "" & vbCr
SQL = SQL & "testrsltetc:" & "" & vbCr
SQL = SQL & "fstrgstrid:" & gHOSP.USERID & vbCr
SQL = SQL & "lastupdtrid:" & gHOSP.USERID & vbCr
'
    Call SetSQLData("결과저장2", SQL, "A")
    
    AdoCmd.Execute

    Set AdoCmd = Nothing
    
    
    'param=[017, M17003176, 20170724, 33978, 1, PMO12040, 17488137, 20170724, 1151787391]|1 records|
    '-- 8
    SQL = ""
    SQL = SQL & "SELECT acptstatcd          " & vbCr
    SQL = SQL & "  From lis.lpjmacpt        " & vbCr
    SQL = SQL & " WHERE instcd         = ?  " & vbCr
    SQL = SQL & "   AND ptno           = ?  " & vbCr
    SQL = SQL & "   AND acptdd         = ?  " & vbCr
    SQL = SQL & "   AND acptno         = ?  " & vbCr
    SQL = SQL & "   AND acptitemno     = ?  " & vbCr
    SQL = SQL & "   AND testcd         = ?  " & vbCr
    SQL = SQL & "   AND pid            = ?  " & vbCr
    SQL = SQL & "   AND prcpdd         = ?  " & vbCr
    SQL = SQL & "   AND execprcpuniqno = ?  "
    
    Call SetSQLData("결과저장", SQL, "A")
    
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    
    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptdd", adVarChar, adParamInput, 1000, pREGDT)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptno", adVarChar, adParamInput, 1000, pACPTNO)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptitemno", adVarChar, adParamInput, 1000, 1)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testcd", adVarChar, adParamInput, 1000, gAllTestCd1)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
    
    Set AdoRs = New ADODB.Recordset
    AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic

    If AdoRs.BOF = False Then
        arg_Acptstatcd = AdoRs.Fields("acptstatcd").Value & ""
    End If
    
    Set AdoCmd = Nothing
    Set AdoRs = Nothing
    
    
    'param=[2, 10602673, 017, M17003176, 20170724, 33978, 1, PMO12040, 17488137, 20170724, 1151787391]|1 records|
    '-- 9
    If arg_Acptstatcd = "" Then
        GetINOUT = 0
    Else
        SQL = ""
        SQL = SQL & "Update lis.lpjmacpt " & vbCr
        SQL = SQL & "   SET acptstatcd  = ?, " & vbCr
        SQL = SQL & "       lastupdtdt  = SYSDATE, " & vbCr
        SQL = SQL & "       lastupdtrid = ? " & vbCr
        SQL = SQL & " WHERE instcd         = ? " & vbCr
        SQL = SQL & "   AND ptno           = ? " & vbCr
        SQL = SQL & "   AND acptdd         = ? " & vbCr
        SQL = SQL & "   AND acptno         = CAST(? AS DECIMAL(12,0)) " & vbCr
        SQL = SQL & "   AND acptitemno     = CAST(? AS SMALLINT) " & vbCr
        SQL = SQL & "   AND testcd         = ? " & vbCr
        SQL = SQL & "   AND pid            = ? " & vbCr
        SQL = SQL & "   AND prcpdd         = ? " & vbCr
        SQL = SQL & "   AND execprcpuniqno = CAST(? AS INTEGER)"
        
        
        Call SetSQLData("결과저장", SQL, "A")
        
        Set AdoCmd = New ADODB.Command
        Set AdoCmd.ActiveConnection = AdoCn
    
        AdoCmd.CommandType = adCmdText
        AdoCmd.CommandText = SQL
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptstatcd", adVarChar, adParamInput, 1000, 2) 'arg_Acptstatcd
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptdd", adVarChar, adParamInput, 1000, pREGDT)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptno", adVarChar, adParamInput, 1000, pACPTNO)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptitemno", adVarChar, adParamInput, 1000, 1)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("testcd", adVarChar, adParamInput, 1000, gAllTestCd1)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
        
        AdoCmd.Execute
        
        Set AdoCmd = Nothing
    End If
    
    

    
    GetINOUT = 1
    Exit Function
    
    'param=[017, 17488137, 20170724, 1151787391]|1 records|
    '-- 10
    SQL = ""
    SQL = SQL & "SELECT COUNT(distinct acptstatcd) AS acptststcnt , COUNT(distinct ptnocd) AS ptnocd " & vbCr
    SQL = SQL & "  From lis.lpjmacpt " & vbCr
    SQL = SQL & " WHERE instcd          = ? " & vbCr
    SQL = SQL & "   AND pid             = ? " & vbCr
    SQL = SQL & "   AND prcpdd          = ? " & vbCr
    SQL = SQL & "   AND execprcpuniqno  = CAST(? AS INTEGER) " & vbCr
    SQL = SQL & "   AND acptstatcd     IN ('0', '2', '3', '4', '9') " & vbCr
    SQL = SQL & " GROUP BY instcd, pid, prcpdd, execprcpuniqno "
    
        
    Call SetSQLData("결과저장", SQL, "A")
    
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn
    
    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
    
    Set AdoRs = New ADODB.Recordset
    AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic

    If AdoRs.BOF = False Then
        arg_Acptststcnt = AdoRs.Fields("acptststcnt").Value & ""
        arg_Ptnocd = AdoRs.Fields("ptnocd").Value & ""
    End If
    
    Set AdoCmd = Nothing
    Set AdoRs = Nothing
    
    
    'param=[017, 17488137, 20170724, 1151787391
    '11
    If arg_Acptststcnt < 0 Then
        SQL = ""
        SQL = SQL & "SELECT b.prcpstatcd " & vbCr
        SQL = SQL & "  FROM emr.mmodexip a, emr.mmohiprc b " & vbCr
        SQL = SQL & " WHERE a.instcd         = ? " & vbCr
        SQL = SQL & "   AND a.pid            = ? " & vbCr
        SQL = SQL & "   AND a.prcpdd         = ? " & vbCr
        SQL = SQL & "   AND a.execprcpuniqno = ? " & vbCr
        SQL = SQL & "   AND a.execprcphistcd = 'O' " & vbCr
        SQL = SQL & "   AND a.instcd         = b.instcd " & vbCr
        SQL = SQL & "   AND a.prcpdd         = b.prcpdd " & vbCr
        SQL = SQL & "   AND a.prcpno         = b.prcpno " & vbCr
        SQL = SQL & "   AND a.prcphistno     = b.prcphistno " & vbCr
        SQL = SQL & "   AND b.prcphistcd     = 'O' " & vbCr
        SQL = SQL & "   AND b.prcpclscd      = 'D2' " & vbCr
        SQL = SQL & "   AND b.tempprcpflag   = 'N' "
    
        Call SetSQLData("결과저장", SQL, "A")
        
        Set AdoCmd = New ADODB.Command
        Set AdoCmd.ActiveConnection = AdoCn
        
        AdoCmd.CommandType = adCmdText
        AdoCmd.CommandText = SQL
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
        
        Set AdoRs = New ADODB.Recordset
        AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic
    
        If AdoRs.BOF = False Then
            arg_Prcpstatcd = AdoRs.Fields("prcpstatcd").Value & ""
        End If
        
        Set AdoCmd = Nothing
        Set AdoRs = Nothing
    Else
        SQL = ""
        SQL = SQL & "SELECT b.prcpstatcd " & vbCr
        SQL = SQL & "  FROM emr.mmodexip a, emr.mmohiprc b " & vbCr
        SQL = SQL & " WHERE a.instcd         = ? " & vbCr
        SQL = SQL & "   AND a.pid            = ? " & vbCr
        SQL = SQL & "   AND a.prcpdd         = ? " & vbCr
        SQL = SQL & "   AND a.execprcpuniqno = ? " & vbCr
        SQL = SQL & "   AND a.execprcphistcd = 'O' " & vbCr
        SQL = SQL & "   AND a.instcd         = b.instcd " & vbCr
        SQL = SQL & "   AND a.prcpdd         = b.prcpdd " & vbCr
        SQL = SQL & "   AND a.prcpno         = b.prcpno " & vbCr
        SQL = SQL & "   AND a.prcphistno     = b.prcphistno " & vbCr
        SQL = SQL & "   AND b.prcphistcd     = 'O' " & vbCr
        SQL = SQL & "   AND b.prcpclscd      = 'D2' " & vbCr
        SQL = SQL & "   AND b.tempprcpflag   = 'N' "
    
        Call SetSQLData("결과저장", SQL, "A")
        
        Set AdoCmd = New ADODB.Command
        Set AdoCmd.ActiveConnection = AdoCn
        
        AdoCmd.CommandType = adCmdText
        AdoCmd.CommandText = SQL
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
        
        Set AdoRs = New ADODB.Recordset
        AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic
    
        If AdoRs.BOF = False Then
            arg_Prcpstatcd = AdoRs.Fields("prcpstatcd").Value & ""
        End If
        
        Set AdoCmd = Nothing
        Set AdoRs = Nothing
    End If
    
    'param=[710, 10602673, 017, 17488137, 20170724, 1151787391
    '12
    If arg_Prcpstatcd < 0 Then
        SQL = ""
        SQL = SQL & "Update emr.mmohiprc" & vbCr
        SQL = SQL & "   SET prcpstatcd  = ?," & vbCr
        SQL = SQL & "       lastupdtdt  = SYSDATE," & vbCr
        SQL = SQL & "       lastupdtrid = ? " & vbCr
        SQL = SQL & " WHERE (instcd, pid, prcpdd, prcpno, prcphistno) IN " & vbCr
        SQL = SQL & "       (SELECT instcd, pid, prcpdd, prcpno, prcphistno " & vbCr
        SQL = SQL & "          From emr.mmodexip " & vbCr
        SQL = SQL & "         WHERE instcd         = ? " & vbCr
        SQL = SQL & "           AND pid            = ? " & vbCr
        SQL = SQL & "           AND prcpdd         = ? " & vbCr
        SQL = SQL & "           AND execprcpuniqno = ? " & vbCr
        SQL = SQL & "           AND execprcphistcd = 'O' " & vbCr
        SQL = SQL & "       )" & vbCr
        SQL = SQL & "   AND prcphistcd   = 'O' " & vbCr
        SQL = SQL & "   AND prcpclscd    = 'D2' " & vbCr
        SQL = SQL & "   AND tempprcpflag = 'N' "
            
        Set AdoCmd = New ADODB.Command
        Set AdoCmd.ActiveConnection = AdoCn
    
        AdoCmd.CommandType = adCmdText
        AdoCmd.CommandText = SQL
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptstatcd", adVarChar, adParamInput, 1000, arg_Flagcd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
        
        AdoCmd.Execute
        
        Set AdoCmd = Nothing
    Else
        SQL = ""
        SQL = SQL & "Update emr.mmohiprc" & vbCr
        SQL = SQL & "   SET prcpstatcd  = ?," & vbCr
        SQL = SQL & "       lastupdtdt  = SYSDATE," & vbCr
        SQL = SQL & "       lastupdtrid = ? " & vbCr
        SQL = SQL & " WHERE (instcd, pid, prcpdd, prcpno, prcphistno) IN " & vbCr
        SQL = SQL & "       (SELECT instcd, pid, prcpdd, prcpno, prcphistno " & vbCr
        SQL = SQL & "          From emr.mmodexip " & vbCr
        SQL = SQL & "         WHERE instcd         = ? " & vbCr
        SQL = SQL & "           AND pid            = ? " & vbCr
        SQL = SQL & "           AND prcpdd         = ? " & vbCr
        SQL = SQL & "           AND execprcpuniqno = ? " & vbCr
        SQL = SQL & "           AND execprcphistcd = 'O' " & vbCr
        SQL = SQL & "       )" & vbCr
        SQL = SQL & "   AND prcphistcd   = 'O' " & vbCr
        SQL = SQL & "   AND prcpclscd    = 'D2' " & vbCr
        SQL = SQL & "   AND tempprcpflag = 'N' "
            
        Set AdoCmd = New ADODB.Command
        Set AdoCmd.ActiveConnection = AdoCn
    
        AdoCmd.CommandType = adCmdText
        AdoCmd.CommandText = SQL
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptstatcd", adVarChar, adParamInput, 1000, arg_Flagcd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
        
        AdoCmd.Execute
        
        Set AdoCmd = Nothing
    End If
    
    'param=[710, 10602673, 017, 17488137, 20170724, 1151787391]
    '13
    If arg_Prcpstatcd < 0 Then
        SQL = ""
        SQL = SQL & "Update emr.mmodexip a" & vbCr
        SQL = SQL & "   SET a.execprcpstatcd = ?," & vbCr
        SQL = SQL & "       a.lastupdtdt     = SYSDATE," & vbCr
        SQL = SQL & "       a.lastupdtrid    = ?" & vbCr
        SQL = SQL & " WHERE a.instcd         = ?" & vbCr
        SQL = SQL & "   AND a.pid            = ?" & vbCr
        SQL = SQL & "   AND a.prcpdd         = ?" & vbCr
        SQL = SQL & "   AND a.execprcpuniqno = ?" & vbCr
        SQL = SQL & "   AND a.execprcphistcd = 'O'" & vbCr
                    
        Set AdoCmd = New ADODB.Command
        Set AdoCmd.ActiveConnection = AdoCn
    
        AdoCmd.CommandType = adCmdText
        AdoCmd.CommandText = SQL
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpstatcd", adVarChar, adParamInput, 1000, arg_Flagcd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
        
        AdoCmd.Execute
        
        Set AdoCmd = Nothing
    Else
        SQL = ""
        SQL = SQL & "Update emr.mmodexip a" & vbCr
        SQL = SQL & "   SET a.execprcpstatcd = ?," & vbCr
        SQL = SQL & "       a.lastupdtdt     = SYSDATE," & vbCr
        SQL = SQL & "       a.lastupdtrid    = ?" & vbCr
        SQL = SQL & " WHERE a.instcd         = ?" & vbCr
        SQL = SQL & "   AND a.pid            = ?" & vbCr
        SQL = SQL & "   AND a.prcpdd         = ?" & vbCr
        SQL = SQL & "   AND a.execprcpuniqno = ?" & vbCr
        SQL = SQL & "   AND a.execprcphistcd = 'O'" & vbCr
                    
        Set AdoCmd = New ADODB.Command
        Set AdoCmd.ActiveConnection = AdoCn
    
        AdoCmd.CommandType = adCmdText
        AdoCmd.CommandText = SQL
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpstatcd", adVarChar, adParamInput, 1000, arg_Flagcd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
        
        AdoCmd.Execute
        
        Set AdoCmd = Nothing
    End If
    
    'param=[017, 1151787391, 20170724, 710]|1 records|
    '14
'''    SQL = ""
'''    SQL = SQL & "SELECT COUNT(prcpdd) AS tretcnt" & vbCr
'''    SQL = SQL & "  From emr.mmodexpt" & vbCr
'''    SQL = SQL & " WHERE instcd         = ?" & vbCr
'''    SQL = SQL & "   AND execprcpuniqno = ?" & vbCr
'''    SQL = SQL & "   AND prcpdd         = ?" & vbCr
'''    SQL = SQL & "   AND tretflagcd     = ?"
'''
'''    Set AdoCmd = New ADODB.Command
'''    Set AdoCmd.ActiveConnection = AdoCn
'''
'''    AdoCmd.CommandType = adCmdText
'''    AdoCmd.CommandText = SQL
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("tretflagcd", adVarChar, adParamInput, 1000, arg_Flagcd)
'''
'''    Set AdoRs = New ADODB.Recordset
'''    AdoRs.Open AdoCmd, , adOpenStatic, adLockBatchOptimistic
'''
'''    If AdoRs.BOF = False Then
'''        arg_Tretcnt = AdoRs.Fields("tretcnt").Value & ""
'''    End If
'''
'''    Set AdoCmd = Nothing
'''    Set AdoRs = Nothing
'''
'''
'''    '20170724, 1151787391, 710, 017, 20170724, 142613, null, 10602673, null, 10602673, 10602673]|1 records|
'''    '15
'''    If arg_Tretcnt < 0 Then
'''        SQL = ""
'''        SQL = SQL & "INSERT INTO emr.mmodexpt (prcpdd,       execprcpuniqno," & vbCr
'''        SQL = SQL & "                          tretflagcd,   instcd," & vbCr
'''        SQL = SQL & "                          tretdd,       trettm,    tretpsnid," & vbCr
'''        SQL = SQL & "                          fstrgstrid,   fstrgstdt," & vbCr
'''        SQL = SQL & "                          lastupdtrid,  lastupdtdt)" & vbCr
'''        SQL = SQL & "                  VALUES (?,       CAST(? AS INTEGER)," & vbCr
'''        SQL = SQL & "                          ?,       ?," & vbCr
'''        SQL = SQL & "                          ?,       ?,  CASE WHEN ? IS NULL THEN ? ELSE ? END," & vbCr
'''        SQL = SQL & "                          Print , SYSDATE," & vbCr
'''        SQL = SQL & "                          ?,    SYSDATE)" & vbCr
'''
'''        Set AdoCmd = New ADODB.Command
'''        Set AdoCmd.ActiveConnection = AdoCn
'''
'''        AdoCmd.CommandType = adCmdText
'''        AdoCmd.CommandText = SQL
'''        AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
'''        AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
'''        AdoCmd.Parameters.Append AdoCmd.CreateParameter("tretflagcd", adVarChar, adParamInput, 1000, arg_Flagcd)
'''        AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
'''        AdoCmd.Parameters.Append AdoCmd.CreateParameter("tretdd", adVarChar, adParamInput, 1000, arg_Tretdd)
'''        AdoCmd.Parameters.Append AdoCmd.CreateParameter("trettm", adVarChar, adParamInput, 1000, arg_Trettm)
'''        AdoCmd.Parameters.Append AdoCmd.CreateParameter("tretpsnid", adVarChar, adParamInput, 1000, "")
'''        AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
'''        AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstdt", adVarChar, adParamInput, 1000, "")
'''        AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
'''
'''        AdoCmd.Execute
'''    Else
'''
'''    End If
    
    'Select 후 Update를 Delete Insert 로 수정
        
    SQL = ""
    SQL = SQL & "Delete From emr.mmodexip a" & vbCr
    SQL = SQL & " WHERE a.instcd         = ?" & vbCr
    SQL = SQL & "   AND a.pid            = ?" & vbCr
    SQL = SQL & "   AND a.prcpdd         = ?" & vbCr
    SQL = SQL & "   AND a.execprcpuniqno = ?" & vbCr
    SQL = SQL & "   AND a.execprcphistcd = 'O'" & vbCr
                
    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
    
    AdoCmd.Execute
    
    Set AdoCmd = Nothing
        
        
        
    SQL = ""
    SQL = SQL & "INSERT INTO emr.mmodexpt (prcpdd,       execprcpuniqno," & vbCr
    SQL = SQL & "                          tretflagcd,   instcd," & vbCr
    SQL = SQL & "                          tretdd,       trettm,    tretpsnid," & vbCr
    SQL = SQL & "                          fstrgstrid,   fstrgstdt," & vbCr
    SQL = SQL & "                          lastupdtrid,  lastupdtdt)" & vbCr
    SQL = SQL & "                  VALUES (?,       CAST(? AS INTEGER)," & vbCr
    SQL = SQL & "                          ?,       ?," & vbCr
    SQL = SQL & "                          ?,       ?,  CASE WHEN ? IS NULL THEN ? ELSE ? END," & vbCr
    SQL = SQL & "                          Print , SYSDATE," & vbCr
    SQL = SQL & "                          ?,    SYSDATE)" & vbCr

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("prcpdd", adVarChar, adParamInput, 1000, arg_Prcpdd)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("execprcpuniqno", adVarChar, adParamInput, 1000, arg_Execprcpuniqno)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("tretflagcd", adVarChar, adParamInput, 1000, arg_Flagcd)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("tretdd", adVarChar, adParamInput, 1000, arg_Tretdd)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("trettm", adVarChar, adParamInput, 1000, arg_Trettm)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("tretpsnid", adVarChar, adParamInput, 1000, "")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstdt", adVarChar, adParamInput, 1000, "")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
    
    AdoCmd.Execute
        
    
    GetINOUT = 1
    
Exit Function

DBErr:
    GetINOUT = ""
    
End Function




    
'/* 검사결과(Header) 등록  저장합니다.
'himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml instestrslt
'param=[M17003176, 20170724, 32715, 017,
'17488137, 142613, null, null, null,
'HPV High Risk Type : Positive (18+, 68+, 31+++)
'HPV Low  Risk Type : Positive (70+, 61+) ,
'null,
'null, null, 세포검체 other, HPV genotyping real-time PCR, null, 0,
'0, 0, 0, 0, 0, 0, null,
'10602673,
'null, 0, null, null, null,
'null, null, 10602673,
'null, null, 10602673, null]
'*/
Private Function Regist_Result_Header(ByVal pPTNO As String, ByVal pREGDT As String, ByVal pPTSEQ As String, ByVal pPID As String, ByVal pRESULT As String) As String

    
On Error GoTo DBErr

    Regist_Result_Header = -1
    '-- 이력기록
    SQL = ""
    SQL = SQL & "INSERT INTO lis.lprmrslt (ptno,       rsltrgstdd,    rsltrgstno,     instcd,         rsltrgsthistno," & vbCr
    SQL = SQL & "                                      pid,        rsltrgsttm,    grostestrecdd,  grostestrectm,  grostestrecid," & vbCr
    SQL = SQL & "                                      readdd,     readtm,        readid,         extrpartcnts,   extrmthdcnts," & vbCr
    SQL = SQL & "                                      diagcnts,   diagcd," & vbCr
    SQL = SQL & "                                      spckeepflagcd, rslthideflagcd, conccaseflagcd, preprsltflagcd, ugcyalertflagcd, cnstcd," & vbCr
    SQL = SQL & "                                      rsltrgstid, cnclflagcd,    cnclresncd,     cncldd,         cncltm," & vbCr
    SQL = SQL & "                                      grospic,    keybloc,       tissbloct,      tissblocnt,     readgrade," & vbCr
    SQL = SQL & "                                      cnclid,     delflagcd,     signno," & vbCr
    SQL = SQL & "                                      fstrgstdt,  fstrgstrid," & vbCr
    SQL = SQL & "                                      lastupdtdt , lastupdtrid, cncrjudgflagcd" & vbCr
    SQL = SQL & "                                     )" & vbCr
    SQL = SQL & "            SELECT ptno,       rsltrgstdd,    rsltrgstno,   instcd," & vbCr
    SQL = SQL & "                   (SELECT MAX(z.rsltrgsthistno)+1" & vbCr
    SQL = SQL & "                      FROM lis.lprmrslt z" & vbCr
    SQL = SQL & "                     WHERE instcd         = ?" & vbCr
    SQL = SQL & "                       AND ptno           = ?" & vbCr
    SQL = SQL & "                       AND pid            = ?" & vbCr
    SQL = SQL & "                   )," & vbCr
    SQL = SQL & "                   pid,        rsltrgsttm,    grostestrecdd,  grostestrectm,  grostestrecid," & vbCr
    SQL = SQL & "                   readdd,     readtm,        readid,         extrpartcnts,   extrmthdcnts," & vbCr
    SQL = SQL & "                   diagcnts,   diagcd," & vbCr
    SQL = SQL & "                   spckeepflagcd, rslthideflagcd, conccaseflagcd, preprsltflagcd, ugcyalertflagcd,  cnstcd," & vbCr
    SQL = SQL & "                   rsltrgstid, cnclflagcd,    cnclresncd,     cncldd,         cncltm," & vbCr
    SQL = SQL & "                   grospic,    keybloc,       tissbloct,      tissblocnt,     readgrade," & vbCr
    SQL = SQL & "                   cnclid,     '1',    signno," & vbCr
    SQL = SQL & "                   fstrgstdt,  fstrgstrid," & vbCr
    SQL = SQL & "                   lastupdtdt , lastupdtrid, cncrjudgflagcd" & vbCr
    SQL = SQL & "              From lis.lprmrslt" & vbCr
    SQL = SQL & "             WHERE instcd         = ?" & vbCr
    SQL = SQL & "               AND ptno           = ?" & vbCr
    SQL = SQL & "               AND pid            = ?" & vbCr
    SQL = SQL & "               AND rsltrgsthistno = 1" & vbCr
    SQL = SQL & "               AND delflagcd      = '0'" & vbCr
    
    Call SetSQLData("결과저장", SQL, "A")

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)    ':arg_instcd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)            ':arg_ptno
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)                   ':arg_pid
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)    ':arg_instcd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)            ':arg_ptno
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)                   ':arg_pid
    
    AdoCmd.Execute
    
    
    Set AdoCmd = Nothing
    
    '-- 삭제
    SQL = ""
    SQL = SQL & "Delete From lis.lprmrslt " & vbCr
    SQL = SQL & "             WHERE instcd         = ?" & vbCr
    SQL = SQL & "               AND ptno           = ?" & vbCr
    SQL = SQL & "               AND pid            = ?" & vbCr
    SQL = SQL & "               AND rsltrgsthistno = 1" & vbCr
    SQL = SQL & "               AND delflagcd      = '0'" & vbCr
    
    Call SetSQLData("결과저장", SQL, "A")

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)    ':arg_instcd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)            ':arg_ptno
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)                   ':arg_pid
    
    AdoCmd.Execute
    
    Set AdoCmd = Nothing
    
    
    '-- 신규저장
    SQL = ""
    SQL = SQL & "insert into lis.lprmrslt (" & vbCr
    SQL = SQL & "ptno,          rsltrgstdd,     rsltrgstno,     rsltrgsthistno, instcd,                                 " & vbCr
    SQL = SQL & "pid,           rsltrgsttm,     grostestrecdd,  grostestrectm,  grostestrecid,                          " & vbCr
    SQL = SQL & "diagcnts,                                                                                              " & vbCr
    SQL = SQL & "readdd,        readtm,         readid,         extrpartcnts,   extrmthdcnts,   diagcd,                 " & vbCr
    SQL = SQL & "spckeepflagcd, rslthideflagcd, cncrjudgflagcd, conccaseflagcd, preprsltflagcd, ugcyalertflagcd, cnstcd," & vbCr
    SQL = SQL & "rsltrgstid,    cnclflagcd,     cnclresncd,     cncldd,         cncltm,                                 " & vbCr
    SQL = SQL & "grospic,       keybloc,        tissbloct,      tissblocnt,     readgrade,                              " & vbCr
    SQL = SQL & "cnclid,        delflagcd,                                                                              " & vbCr
    SQL = SQL & "fstrgstdt,     fstrgstrid,                                                                             " & vbCr
    SQL = SQL & "lastupdtdt,    lastupdtrid)                                                                            " & vbCr
    SQL = SQL & " values (" & vbCr
    SQL = SQL & "?, ?, ?, 1, ?," & vbCr
    SQL = SQL & "?, ?, ?, ?, ?," & vbCr
    SQL = SQL & "?," & vbCr
    SQL = SQL & "?, ?, ?, ?, ?, ?," & vbCr
    SQL = SQL & "?, ?, ?, ?, ?, ?, ?," & vbCr
    SQL = SQL & "CASE WHEN ? IS NULL THEN ? ELSE ? END ,       '-',       '-',       '-',       '-'," & vbCr
    SQL = SQL & "?, ?, ?, ?, ?," & vbCr
    SQL = SQL & "'-',   '0'," & vbCr
    SQL = SQL & "sysdate,  case when ? is null then ? else ? end ," & vbCr
    SQL = SQL & "sysdate,  case when ? is null then ? else ? end )" & vbCr
    
    Call SetSQLData("결과저장", SQL, "A")

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn


    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)            ':arg_ptno
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstdd", adVarChar, adParamInput, 1000, pREGDT)         ':arg_rsltrgstdd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstno", adVarChar, adParamInput, 1000, pPTSEQ)              ':arg_rsltrgstno
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)    ':arg_instcd
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)                   ':arg_pid
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgsttm", adVarChar, adParamInput, 1000, gSysTime)             ':arg_rsltrgsttm
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("grostestrecdd", adVarChar, adParamInput, 1000, "")             ':arg_grostestrecdd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("grostestrectm", adVarChar, adParamInput, 1000, "")             ':arg_grostestrectm
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("grostestrecid", adVarChar, adParamInput, 1000, "")             ':arg_grostestrecid
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("diagcnts", adVarChar, adParamInput, 1000, pRESULT)             ':arg_diagcnts
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("readdd", adVarChar, adParamInput, 1000, "")             ':arg_readdd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("readtm", adVarChar, adParamInput, 1000, "")             ':arg_readtm
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("readid", adVarChar, adParamInput, 1000, "")             ':arg_readid
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("extrpartcnts", adVarChar, adParamInput, 1000, "세포검체 other")             ':arg_extrpartcnts
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("extrmthdcnts", adVarChar, adParamInput, 1000, "HPV genotyping real-time PCR")             ':arg_extrmthdcnts
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("diagcd", adVarChar, adParamInput, 1000, "")             ':arg_diagcd
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("spckeepflagcd", adVarChar, adParamInput, 1000, 0)             ':arg_spckeepflagcd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rslthideflagcd", adVarChar, adParamInput, 1000, 0)             ':arg_rslthideflagcd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("cncrjudgflagcd", adVarChar, adParamInput, 1000, 0)             ':arg_cncrjudgflagcd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("conccaseflagcd", adVarChar, adParamInput, 1000, 0)             ':arg_conccaseflagcd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("preprsltflagcd", adVarChar, adParamInput, 1000, 0)             ':arg_preprsltflagcd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ugcyalertflagcd", adVarChar, adParamInput, 1000, 0)             ':arg_ugcyalertflagcd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("cnstcd", adVarChar, adParamInput, 1000, 0)             ':arg_cnstcd
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstid", adVarChar, adParamInput, 1000, gHOSP.USERID)              ':arg_cellrsltrgstid
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstid", adVarChar, adParamInput, 1000, gHOSP.USERID)              ':arg_userid
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstid", adVarChar, adParamInput, 1000, gHOSP.USERID)              ':arg_cellrsltrgstid
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("grospic", adVarChar, adParamInput, 1000, "0")             ':arg_grospic
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("keybloc", adVarChar, adParamInput, 1000, "")             ':arg_keybloc
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("tissbloct", adVarChar, adParamInput, 1000, "")             ':arg_tissbloct
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("tissblocnt", adVarChar, adParamInput, 1000, "")             ':arg_tissblocnt
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("readgrade", adVarChar, adParamInput, 1000, "")             ':arg_readgrade
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstrid", adVarChar, adParamInput, 1000, gHOSP.USERID)              ':arg_cellrsltrgstid
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstrid", adVarChar, adParamInput, 1000, gHOSP.USERID)              ':arg_userid
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstrid", adVarChar, adParamInput, 1000, gHOSP.USERID)              ':arg_cellrsltrgstid
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)              ':arg_cellrsltrgstid
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)              ':arg_userid
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)              ':arg_cellrsltrgstid
    
    AdoCmd.Execute
    
    Set AdoCmd = Nothing
    
    Regist_Result_Header = 1
    
Exit Function

DBErr:

    MsgBox Err.Description
    
    
End Function

'/* 검사결과내용 등록 합니다.
'himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml instestrsltcnts
'param=[
'M17003176,
'20170724,
'32715,
'017,
'17488137,
'[Methods]
'   Seegene HPV Real-time PCR (Anyplex II HPV 28 Detection Real-time PCR)
'
'[Result]
'HPV High Risk Type : POSITIVE (18+, 68+, 31+++)
'HPV Low  Risk Type : POSITIVE (70+, 61+) ,
'null,
'Adequate,
'null,
'null,
'null,
'3. Comment
'인유두종 바이러스 (Human papilloma virus)는 자궁경부암의 주요 위험인자로 알려져 있습니다. 암과의 관련성 정도에 따라 고위험군 (high risk)과 저위험군 (low risk)로 구분되며, 저위험군 HPV는 대개 시간이 지나면 소실되거나 사마귀 등의 양성변변의 원인이 되는 반면, 고위험군은 자궁경부암을 유발시키는데 관여합니다 (N Engl J Med. 2003 348:518).
'
'◆ 본 제품은 19종의 고위험군 HPV (16, 18, 26, 31, 33, 25, 29, 45, 51, 52, 53, 56, 58, 59, 66, 68, 73, 82형)과 9종의 저위험군 HPV (6, 11, 40, 42, 43, 44, 54, 61, 70형), 내부대조군의 타켓핵산을 검출합니다.
'◆ Viral load에서 +++:10^5 copies/reaction, ++:10^5~10^2 copies/reaction, +:10^2 copies/reaction의 농도로 해석될 수 있습니다. 이 중 “+”는 매우 낮은 농도로 감염 시기, 검체 채취 상태에 따라 반복 검사 시 재현되지 않을 수 있습니다.
'◆ PCR 검사는 검체 내 균수가 적거나 부적절한 검체 의석 또는 증폭 억제물질이 존재하는 경우 위음성이 나올 수 있습니다. 또한, PCR 검사는 유전자 유무를 검사하므로 생존균과 사균의 구분이 안되어 위양성의 가능성이 있습니다. 결과 해석 시 임상 양상과 연관지어 판단하시기 바랍니다.
'◆ 상기 검사는 검사 방법, 시약의 정도관리 및 검사 결과가 병리과 전문의에 의해 확인되었습니다.
'   (검사 담당자: 이재왕),
'10602673,
'10602673]
'*/
Private Function Regist_Result_Detail(ByVal pPTNO As String, ByVal pREGDT As String, ByVal pPTSEQ As String, ByVal pPID As String, ByVal pRESULT As String) As String
    Dim strRsltCmt  As String
    Dim strCmtCnt   As String
        
    
On Error GoTo DBErr

    Regist_Result_Detail = -1
    
    '-- 이력기록
    SQL = ""
    SQL = SQL & "INSERT INTO lis.lprmcnts (ptno,       rsltrgstdd, rsltrgstno, rsltrgsthistno, instcd, pid," & vbCr
    SQL = SQL & "                          rsltcnts1,  rsltcnts2,  rsltcnts3," & vbCr
    SQL = SQL & "                          rsltcnts4,  rsltcnts5,  rsltcnts6," & vbCr
    SQL = SQL & "                          cmtcnts,    delflagcd," & vbCr
    SQL = SQL & "                          fstrgstdt,              fstrgstrid," & vbCr
    SQL = SQL & "                          lastupdtdt,             lastupdtrid)" & vbCr
    SQL = SQL & "SELECT ptno,       rsltrgstdd, rsltrgstno," & vbCr
    SQL = SQL & "       (SELECT MAX(z.rsltrgsthistno)+1" & vbCr
    SQL = SQL & "          FROM lis.lprmcnts z" & vbCr
    SQL = SQL & "         WHERE instcd         = ?" & vbCr
    SQL = SQL & "           AND ptno           = ?" & vbCr
    SQL = SQL & "           AND pid            = ?" & vbCr
    SQL = SQL & "       )," & vbCr
    SQL = SQL & "       instcd,     pid," & vbCr
    SQL = SQL & "       rsltcnts1,  rsltcnts2,  rsltcnts3," & vbCr
    SQL = SQL & "       rsltcnts4,  rsltcnts5,  rsltcnts6," & vbCr
    SQL = SQL & "       cmtcnts,    '1'," & vbCr
    SQL = SQL & "       fstrgstdt,              fstrgstrid," & vbCr
    SQL = SQL & "       lastupdtdt , lastupdtrid" & vbCr
    SQL = SQL & "  From lis.lprmcnts" & vbCr
    SQL = SQL & " WHERE instcd         = ?" & vbCr
    SQL = SQL & "   AND ptno           = ?" & vbCr
    SQL = SQL & "   AND pid            = ?" & vbCr
    SQL = SQL & "   AND rsltrgsthistno = 1" & vbCr
    SQL = SQL & "   AND delflagcd      = '0'" & vbCr
    
    Call SetSQLData("결과저장", SQL, "A")

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)    ':arg_instcd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)            ':arg_ptno
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)                   ':arg_pid
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)    ':arg_instcd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)            ':arg_ptno
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)                   ':arg_pid
    
    AdoCmd.Execute
    
    Set AdoCmd = Nothing
    
    '-- 삭제
    SQL = ""
    SQL = SQL & "Delete From lis.lprmcnts" & vbCr
    SQL = SQL & " WHERE instcd         = ?" & vbCr
    SQL = SQL & "   AND ptno           = ?" & vbCr
    SQL = SQL & "   AND pid            = ?" & vbCr
    SQL = SQL & "   AND rsltrgsthistno = 1" & vbCr
    SQL = SQL & "   AND delflagcd      = '0'" & vbCr
    Call SetSQLData("결과저장", SQL, "A")

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)      ':arg_instcd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)               ':arg_ptno
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)                 ':arg_pid
        
    AdoCmd.Execute
    
    Set AdoCmd = Nothing
    
    '-- 저장
'    strRsltCmt = "[Methods]" & vbNewLine
'    strRsltCmt = strRsltCmt & "Seegene HPV Real-time PCR (Anyplex II HPV 28 Detection Real-time PCR)" & vbNewLine
'    strRsltCmt = strRsltCmt & vbNewLine
'    strRsltCmt = strRsltCmt & "[Result]" & vbNewLine
'    strRsltCmt = strRsltCmt & pRESULT
    
    strRsltCmt = "[Result]" & vbNewLine
    strRsltCmt = strRsltCmt & pRESULT
    
'                strCmtCnt = "인유두종 바이러스 (Human papilloma virus)는 자궁경부암의 주요 위험인자로 알려져 있습니다. 암과의 관련성 정도에 따라 고위험군 (high risk)과 저위험군 (low risk)로 구분되며, 저위험군 HPV는 대개 시간이 지나면 소실되거나 사마귀 등의 양성변변의 원인이 되는 반면, 고위험군은 자궁경부암을 유발시키는데 관여합니다 (N Engl J Med. 2003 348:518)."
'    strCmtCnt = strCmtCnt & "" & Chr(10)
'    strCmtCnt = strCmtCnt & "◆ 본 제품은 19종의 고위험군 HPV (16, 18, 26, 31, 33, 25, 29, 45, 51, 52, 53, 56, 58, 59, 66, 68, 73, 82형)과 9종의 저위험군 HPV (6, 11, 40, 42, 43, 44, 54, 61, 70형), 내부대조군의 타켓핵산을 검출합니다."
'    strCmtCnt = strCmtCnt & "" & Chr(10)
'    strCmtCnt = strCmtCnt & "◆ Viral load에서 +++:10^5 copies/reaction, ++:10^5~10^2 copies/reaction, +:10^2 copies/reaction의 농도로 해석될 수 있습니다. 이 중 “+”는 매우 낮은 농도로 감염 시기, 검체 채취 상태에 따라 반복 검사 시 재현되지 않을 수 있습니다."
'    strCmtCnt = strCmtCnt & "" & Chr(10)
'    strCmtCnt = strCmtCnt & "◆ PCR 검사는 검체 내 균수가 적거나 부적절한 검체 의석 또는 증폭 억제물질이 존재하는 경우 위음성이 나올 수 있습니다. 또한, PCR 검사는 유전자 유무를 검사하므로 생존균과 사균의 구분이 안되어 위양성의 가능성이 있습니다. 결과 해석 시 임상 양상과 연관지어 판단하시기 바랍니다."
'    strCmtCnt = strCmtCnt & "" & Chr(10)
'    strCmtCnt = strCmtCnt & "◆ 상기 검사는 검사 방법, 시약의 정도관리 및 검사 결과가 병리과 전문의에 의해 확인되었습니다."
'    strCmtCnt = strCmtCnt & "" & Chr(10)
'    strCmtCnt = strCmtCnt & "   (검사 담당자: " & gHOSP.USERNM & ")"
'
'                strCmtCnt = "인유두종 바이러스 (Human papilloma virus)는 자궁경부암의 주요 위험인자로 알려져 있습니다. 암과의 관련성 정도에 따라 고위험군 (high risk)과 저위험군 (low risk)로 구분되며, 저위험군 HPV는 대개 시간이 지나면 소실되거나 사마귀 등의 양성변변의 원인이 되는 반면, 고위험군은 자궁경부암을 유발시키는데 관여합니다 (N Engl J Med. 2003 348:518)."
'    strCmtCnt = strCmtCnt & "" & Chr(10)
'    strCmtCnt = strCmtCnt & "◆ 본 제품은 19종의 고위험군 HPV (16, 18, 26, 31, 33, 25, 29, 45, 51, 52, 53, 56, 58, 59, 66, 68, 73, 82형)과 9종의 저위험군 HPV (6, 11, 40, 42, 43, 44, 54, 61, 70형), 내부대조군의 타켓핵산을 검출합니다."
'    strCmtCnt = strCmtCnt & "" & Chr(10)
'    strCmtCnt = strCmtCnt & "◆ Viral load에서 +++:10^5 copies/reaction, ++:10^5~10^2 copies/reaction, +:10^2 copies/reaction의 농도로 해석될 수 있습니다. 이 중 “+”는 매우 낮은 농도로 감염 시기, 검체 채취 상태에 따라 반복 검사 시 재현되지 않을 수 있습니다."
'    strCmtCnt = strCmtCnt & "" & Chr(10)
'    strCmtCnt = strCmtCnt & "◆ PCR 검사는 검체 내 균수가 적거나 부적절한 검체 의석 또는 증폭 억제물질이 존재하는 경우 위음성이 나올 수 있습니다. 또한, PCR 검사는 유전자 유무를 검사하므로 생존균과 사균의 구분이 안되어 위양성의 가능성이 있습니다. 결과 해석 시 임상 양상과 연관지어 판단하시기 바랍니다."
'    strCmtCnt = strCmtCnt & "" & Chr(10)
'    strCmtCnt = strCmtCnt & "◆ 상기 검사는 검사 방법, 시약의 정도관리 및 검사 결과가 병리과 전문의에 의해 확인되었습니다."
'    strCmtCnt = strCmtCnt & "" & Chr(10)
'    strCmtCnt = strCmtCnt & "   (검사 담당자: " & gHOSP.USERNM & ")"
    
                strCmtCnt = "본 제품은 19종의 고위험군HPV 19종 (16, 18, 26, 31, 33, 35, 39, 45, 51, 52, 53, 56, 58, 59, 66, 68, 69, 73, 82형)과 9종의 저위험군 HPV 9종 (6, 11, 40, 42, 43, 44, 54, 61, 70형), 내부대조군(IC, Internal Control)의 타겟핵산을동시또는단독으로검출합니다."
    strCmtCnt = strCmtCnt & "" & Chr(10)
    strCmtCnt = strCmtCnt & "본 검사의 검출 한계치는 50 copies/reaction 입니다."
    strCmtCnt = strCmtCnt & "" & Chr(10)
    strCmtCnt = strCmtCnt & "분석 참고치"
    strCmtCnt = strCmtCnt & "" & Chr(10)
    'strCmtCnt = strCmtCnt & "+++ : 105 copies/reaction 이상 (Ct value 30 이하)"
    strCmtCnt = strCmtCnt & "+++ : 100000 copies/reaction 이상 (Ct value 30 이하)"
    strCmtCnt = strCmtCnt & "" & Chr(10)
    'strCmtCnt = strCmtCnt & "++   : 102 ~ 105 copies/reaction (Ct value 30 ~ 40)"
    strCmtCnt = strCmtCnt & "++   : 100 ~ 100000 copies/reaction (Ct value 30 ~ 40)"
    strCmtCnt = strCmtCnt & "" & Chr(10)
    strCmtCnt = strCmtCnt & "+     : 100 copies/reaction 이하 (Ct value 40 ~ 50 ; 검출한계치 이하의 HPV가 포함되어 있을 경우 재현되지 않을 수 있습니다.)"
    strCmtCnt = strCmtCnt & "" & Chr(10)
    strCmtCnt = strCmtCnt & "*CT Value: Cycle threshold"
    strCmtCnt = strCmtCnt & "" & Chr(10)
    strCmtCnt = strCmtCnt & "상기 검사는 검사 방법, 시약의 정도 관리 및 검사결과가 병리과 전문의에 의해 확인 되었습니다."
    strCmtCnt = strCmtCnt & "" & Chr(10)
    strCmtCnt = strCmtCnt & "검사자:" & gHOSP.USERNM & "           병리과 전문의: 신옥란"
    
    
    SQL = ""
    SQL = SQL & "insert into lis.lprmcnts (" & vbCr
    SQL = SQL & "ptno,       rsltrgstdd, rsltrgstno, rsltrgsthistno, instcd, pid," & vbCr
    SQL = SQL & "rsltcnts1,  rsltcnts2,  rsltcnts3," & vbCr
    SQL = SQL & "rsltcnts4,  rsltcnts5,  rsltcnts6," & vbCr
    SQL = SQL & "cmtcnts,    delflagcd," & vbCr
    SQL = SQL & "fstrgstdt,  fstrgstrid," & vbCr
    SQL = SQL & "lastupdtdt, lastupdtrid)" & vbCr
    SQL = SQL & " values (" & vbCr
    SQL = SQL & "?, ?, ?, 1, ?, ?," & vbCr
    SQL = SQL & "?, ?, ?," & vbCr
    SQL = SQL & "?, ?, ?," & vbCr
    SQL = SQL & "?,'0'," & vbCr
    SQL = SQL & "sysdate, ?," & vbCr
    SQL = SQL & "sysdate, ?)" & vbCr
    
    Call SetSQLData("결과저장", SQL, "A")

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)            ':arg_ptno
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstdd", adVarChar, adParamInput, 1000, pREGDT)         ':arg_rsltrgstdd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstno", adVarChar, adParamInput, 1000, pPTSEQ)              ':arg_rsltrgstno
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)    ':arg_instcd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)                   ':arg_pid
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltcnts1", adVarChar, adParamInput, 1000, strRsltCmt)             ':rsltcnts1
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltcnts2", adVarChar, adParamInput, 1000, "")             ':rsltcnts2
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltcnts3", adVarChar, adParamInput, 1000, "Adequate")             ':rsltcnts3
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltcnts4", adVarChar, adParamInput, 1000, "")             ':rsltcnts4
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltcnts5", adVarChar, adParamInput, 1000, "")             ':rsltcnts5
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltcnts6", adVarChar, adParamInput, 1000, "")             ':rsltcnts6
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("cmtcnts", adVarChar, adParamInput, 1000, strCmtCnt)             ':arg_diagcnts
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstrid", adVarChar, adParamInput, 1000, gHOSP.USERID)              ':arg_cellrsltrgstid
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)              ':arg_userid
    
    AdoCmd.Execute
    
    Set AdoCmd = Nothing
    
    Regist_Result_Detail = 1
    
Exit Function

DBErr:

    MsgBox Err.Description
    
    
End Function

'M17003176, 20170724, 32715, 9, 017, 20170724, 33978, PMO12040, 1,
'[Methods]
'Seegene HPV Real-time PCR (Anyplex II HPV 28 Detection Real-time PCR)
'
'[Result]
'HPV High Risk Type : POSITIVE (18+, 68+, 31+++)
'HPV Low  Risk Type : POSITIVE (70+, 61+) , null, null, 10602673, 10602673]|1 records|


'    SQL = ""
'    SQL = SQL & "insert into lis.lprmtrlt (" & vbCr
'    SQL = SQL & "ptno,          rsltrgstdd, rsltrgstno, rsltrgsthistno, riskflagcd, instcd, " & vbCr
'    SQL = SQL & "acptdd,        acptno,     testcd," & vbCr
'    SQL = SQL & "acptitemno,    testrslt,   delflagcd," & vbCr
'    SQL = SQL & "fstrgstrid,    fstrgstdt," & vbCr
'    SQL = SQL & "lastupdtrid,   lastupdtdt," & vbCr
'    SQL = SQL & "mig,testrsltxml,testrsltetc )" & vbCr
'    SQL = SQL & " values (" & vbCr
'    SQL = SQL & "?, ?, ?, 1, 9, ?," & vbCr
'    SQL = SQL & "?, ?, ?," & vbCr
'    SQL = SQL & "1, ?, ?," & vbCr
''    SQL = SQL & "?,'0'," & vbCr
'    SQL = SQL & "?,sysdate," & vbCr
'    SQL = SQL & "?,sysdate," & vbCr
'    SQL = SQL & "?, ?,?)" & vbCr
'
'    Call SetSQLData("결과저장", SQL)
'
'    Set AdoCmd = New ADODB.Command
'    Set AdoCmd.ActiveConnection = AdoCn
'
'
'    AdoCmd.CommandType = adCmdText
'    AdoCmd.CommandText = SQL
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstdd", adVarChar, adParamInput, 1000, pREGDT)
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstno", adVarChar, adParamInput, 1000, pPTSEQ)
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
'
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptdd", adVarChar, adParamInput, 1000, pREGDT)
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptno", adVarChar, adParamInput, 1000, pACPTNO)
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testcd", adVarChar, adParamInput, 1000, gAllTestCd1)
'
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrslt", adVarChar, adParamInput, 1000, strRsltCmt)
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("deflagcd", adVarChar, adParamInput, 1000, 0)
'
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
'
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
'
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("mig", adVarChar, adParamInput, 1000, "")
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrsltxml", adVarChar, adParamInput, 1000, "")
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrsltetc", adVarChar, adParamInput, 1000, "")
'
'    AdoCmd.Execute


'param=[M17003176, 20170724, 32715,
'9, 017,
'20170724, 33978, PMO12040,

'1, [Methods]
'   Seegene HPV Real-time PCR (Anyplex II HPV 28 Detection Real-time PCR)
'
'[Result]
'HPV High Risk Type : POSITIVE (18+, 68+, 31+++)
'HPV Low  Risk Type : POSITIVE (70+, 61+) , null, null, 10602673, 10602673]|1 records|
'''Private Function Regist_Result_Detail2(ByVal pPTNO As String, ByVal pREGDT As String, ByVal pPTSEQ As String, ByVal pPID As String, ByVal pRESULT As String, Optional ByVal pACPTNO As String) As String
'''    Dim strRsltCmt  As String
'''
'''
'''On Error GoTo DBErr
'''
'''    Regist_Result_Detail2 = -1
'''
'''    strRsltCmt = "[Methods]" & Chr(13) & Chr(10) 'vbNewLine
'''    strRsltCmt = strRsltCmt & "   Seegene HPV Real-time PCR (Anyplex II HPV 28 Detection Real-time PCR)" & Chr(13) & Chr(10) 'vbNewLine
'''   ' strRsltCmt = strRsltCmt & vbNewLine
'''    strRsltCmt = strRsltCmt & "[Result]" & Chr(13) & Chr(10) 'vbNewLine
'''    strRsltCmt = strRsltCmt & pRESULT
'''
'''    SQL = ""
'''    SQL = SQL & "INSERT INTO lis.lprmtrlt"
'''    SQL = SQL & "       (ptno,         rsltrgstdd,      rsltrgstno,           rsltrgsthistno,"
'''    SQL = SQL & "       riskflagcd,    instcd,"
'''    SQL = SQL & "       acptdd,        acptno,          testcd,"
'''
'''    SQL = SQL & "       acptitemno,    testrslt,        testrsltxml,  testrsltetc, delflagcd,"
'''    SQL = SQL & "       fstrgstdt,     fstrgstrid,"
'''    SQL = SQL & "       lastupdtdt,    lastupdtrid)"
'''
''''    SQL = SQL & "VALUES (?,  ?, CAST(? AS DECIMAL(12,0)), 1,"
'''    SQL = SQL & "VALUES (?,  ?, CAST(? AS DECIMAL(12,0)), " & vbCr
'''    SQL = SQL & "(SELECT MAX(z.rsltrgsthistno)+1" & vbCr
'''    SQL = SQL & "   FROM lis.lprmrslt z" & vbCr
'''    SQL = SQL & "  WHERE instcd         = ?" & vbCr
'''    SQL = SQL & "    AND ptno           = ?" & vbCr
'''    SQL = SQL & "    AND pid            = ?" & vbCr
'''    SQL = SQL & ")," & vbCr
'''    SQL = SQL & "        ?,  ?,"
'''    SQL = SQL & "        ?,  CAST(? AS DECIMAL(12,0)),   ?,"
'''    SQL = SQL & "        CAST(? AS SMALLINT),  ?,     ?  , ?  ,   '0',"
'''    SQL = SQL & "        SYSDATE, ?,"
'''    SQL = SQL & "        SYSDATE, ?)"
'''
'''
'''    Call SetSQLData("결과저장", SQL)
'''
'''    Set AdoCmd = New ADODB.Command
'''    Set AdoCmd.ActiveConnection = AdoCn
'''
'''
'''    AdoCmd.CommandType = adCmdText
'''    AdoCmd.CommandText = SQL
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstdd", adVarChar, adParamInput, 1000, pREGDT)
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstno", adVarChar, adParamInput, 1000, pPTSEQ)
'''
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
'''
'''
'''
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("riskflagcd", adVarChar, adParamInput, 1000, 9)
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
'''
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptdd", adVarChar, adParamInput, 1000, pREGDT)
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptno", adVarChar, adParamInput, 1000, pACPTNO)
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testcd", adVarChar, adParamInput, 1000, gAllTestCd1)
'''
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptitemno", adVarChar, adParamInput, 1000, 1)
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrslt", adVarChar, adParamInput, 1000, strRsltCmt)
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrsltxml", adVarChar, adParamInput, 1000, "")
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrsltetc", adVarChar, adParamInput, 1000, "")
'''
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
'''    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
'''
'''
'''    AdoCmd.Execute
'''
'''    Set AdoCmd = Nothing
'''
'''    Regist_Result_Detail2 = 1
'''
'''Exit Function
'''
'''DBErr:
'''
'''    MsgBox Err.Description
'''
'''
'''End Function

Private Function Regist_Result_TMP(ByVal pPTNO As String, ByVal pREGDT As String, ByVal pPTSEQ As String, ByVal pPID As String, ByVal pRESULT As String) As String
    Dim strRsltCmt  As String
    
    
    
'/* T/M/P 진단 설정
'  himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml updlastdiag
'param=[null, null, null,
'0,
'10602673,
'017,
'M17003176,
'17488137] */
    
    
On Error GoTo DBErr

    Regist_Result_TMP = -1
    
    SQL = ""
    SQL = SQL & "Update lis.lprmrslt" & vbCr
    SQL = SQL & "   set readdd          = ? " & vbCr
    SQL = SQL & "      ,readtm          = ? " & vbCr
    SQL = SQL & "      ,readid          = ? " & vbCr
    SQL = SQL & "      ,cnclflagcd      = '-' " & vbCr
    SQL = SQL & "      ,cnclresncd      = '-' " & vbCr
    SQL = SQL & "      ,cncldd          = '-' " & vbCr
    SQL = SQL & "      ,cncltm          = '-' " & vbCr
    SQL = SQL & "      ,cnclid          = '-' " & vbCr
    SQL = SQL & "      ,ugcyalertflagcd = ? " & vbCr
    SQL = SQL & "      ,lastupdtdt      = sysdate " & vbCr
    SQL = SQL & "     , lastupdtrid     = ? " & vbCr
    SQL = SQL & " where instcd          = ? " & vbCr
    SQL = SQL & "   and ptno            = ? " & vbCr
    SQL = SQL & "   and pid             = ? " & vbCr
    SQL = SQL & "   and rsltrgsthistno  = 1 " & vbCr
    SQL = SQL & "   and delflagcd       = '0'" & vbCr
    
    Call SetSQLData("결과저장", SQL, "A")

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("readdd", adVarChar, adParamInput, 1000, "")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("readtm", adVarChar, adParamInput, 1000, "")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("readid", adVarChar, adParamInput, 1000, "")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ugcyalertflagcd", adVarChar, adParamInput, 1000, 0)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
    
    AdoCmd.Execute
    
    Set AdoCmd = Nothing
    
    Regist_Result_TMP = 1
    
Exit Function

DBErr:

    MsgBox Err.Description
    
End Function

Private Function Regist_Result_HIS(ByVal pPTNO As String, ByVal pREGDT As String, ByVal pPTSEQ As String, ByVal pPID As String, ByVal pRESULT As String) As String
    Dim strRsltCmt  As String
    
    
    
'/* 병리번호 발행이력 수정
'himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml updlpcmpnis
'param=[세포검체 other,
'HPV genotyping real-time PCR,
'HPV High Risk Type : Positive (18+, 68+, 31+++)
'HPV Low  Risk Type : Positive (70+, 61+) ,
'null,
'10602673,
'017,
'M17003176] */

    
On Error GoTo DBErr

    Regist_Result_HIS = -1
    
    SQL = ""
    SQL = SQL & "Update lis.lpcmpnis" & vbCr
    SQL = SQL & "   set extrpartcnts = ? ," & vbCr
    SQL = SQL & "       extrmthdcnts = ? ," & vbCr
    SQL = SQL & "       diagcnts     = ? ," & vbCr
    SQL = SQL & "       diagcd       = ? ," & vbCr
    SQL = SQL & "       lastupdtdt   = sysdate," & vbCr
    SQL = SQL & "       lastupdtrid  = ? " & vbCr
    SQL = SQL & " where instcd       = ? " & vbCr
    SQL = SQL & "   and ptno         = ? " & vbCr
    SQL = SQL & "   and delflagcd    = '0'"
    
    Call SetSQLData("결과저장", SQL, "A")

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("extrpartcnts", adVarChar, adParamInput, 1000, "세포검체 other")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("extrmthdcnts", adVarChar, adParamInput, 1000, "HPV genotyping real-time PCR")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("diagcnts", adVarChar, adParamInput, 1000, pRESULT)
'    AdoCmd.Parameters.Append AdoCmd.CreateParameter("diagcnts", adVarChar, adParamInput, 1000, "")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("diagcd", adVarChar, adParamInput, 1000, "")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
    
    AdoCmd.Execute
    
    Set AdoCmd = Nothing
    
    Regist_Result_HIS = 1
    
Exit Function

DBErr:

    MsgBox Err.Description
    
End Function

Private Function Regist_Result_RCPEDIT(ByVal pPTNO As String, ByVal pREGDT As String, ByVal pPTSEQ As String, ByVal pPID As String, ByVal pRESULT As String) As String
    Dim strRsltCmt  As String
    
    
    
'/* 병리과 접수정보 수정
'himed/his/lis/plgyrsltmngtmgr/testrsltrgstmgt/dao/sqls/testrsltrgstdao_sqls.xml updexersltcomfirm
'param=[
'N,
'10602673,
'017,
'M17003176,
'17488137,
'20170724]*/

    
On Error GoTo DBErr

    Regist_Result_RCPEDIT = -1
    
    SQL = ""
    SQL = SQL & "Update lis.lpjmacpt" & vbCr
    SQL = SQL & "   set rsltstatcd  = nvl(?, 'Y')" & vbCr
    SQL = SQL & "     , lastupdtrid = ? " & vbCr
    SQL = SQL & "     , lastupdtdt  = sysdate" & vbCr
    SQL = SQL & " where instcd      = ? " & vbCr
    SQL = SQL & "   and ptno        = ? " & vbCr
    SQL = SQL & "   and pid         = ? " & vbCr
    SQL = SQL & "   and acptdd      = nvl(? , acptdd)" & vbCr
    
    Call SetSQLData("결과저장", SQL, "A")

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn

    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltstatcd", adVarChar, adParamInput, 1000, "N")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptdd", adVarChar, adParamInput, 1000, pREGDT)
    
    AdoCmd.Execute
    
    Set AdoCmd = Nothing
    
    Regist_Result_RCPEDIT = 1
    
Exit Function

DBErr:

    MsgBox Err.Description
    
End Function

Private Function Regist_Result(ByVal RgtDD, ByVal RgtTM, ByVal PTNO, ByVal Reg_NO, ByVal PID, ByVal USERID, ByVal SPCNM, ByVal TESTNM) As String
    Dim SQL2    As String
    
    '=====CHECK DATA BY SELECT SQL
    SQL = ""
    SQL = SQL & " SELECT ptno FROM lis.lprmrslt"
    SQL = SQL & " WHERE"
    SQL = SQL & " ptno='" & PTNO & "' AND"
    SQL = SQL & " instcd='017' AND"
    SQL = SQL & " pid='" & PID & "' AND"
    SQL = SQL & " rsltrgstid='" & USERID & "' AND"
    SQL = SQL & " fstrgstrid='" & USERID & "' AND"
    SQL = SQL & " lastupdtrid='" & USERID & "' AND"
    SQL = SQL & " EXTRPARTCNTS='" & SPCNM & "' AND"
    SQL = SQL & " EXTRMTHDCNTS='" & TESTNM & "'"
    
    'Set RS_L = AdoCn_Local.Execute(SQL, , 1)
    
    If AdoCn.Execute(SQL).EOF Then
        SQL2 = ""
        SQL2 = SQL2 & " INSERT INTO lis.lprmrslt ("
        SQL2 = SQL2 & " ptno," ''병리번호.........#ptno#
        SQL2 = SQL2 & " rsltrgstdd," '결과등록일자.........#rsltrgstdd#
        SQL2 = SQL2 & " rsltrgstno," '결과등록번호#rsltrgstno#
        SQL2 = SQL2 & " rsltrgsthistno," '결과등록이력번호...1
        SQL2 = SQL2 & " instcd," '기관코드:성빈센트병원(instcd = 017)는 고정입니다...#instcd#
        SQL2 = SQL2 & " pid," '등록번호..#pid#
        SQL2 = SQL2 & " rsltrgsttm," '결과등록시간..#rsltrgsttm#
        SQL2 = SQL2 & " grostestrecdd," '육안검사기록일자..#grostestrecdd#
        SQL2 = SQL2 & " grostestrectm," '육안검사기록시간..#grostestrectm#
        SQL2 = SQL2 & " grostestrecid," '육안검사기록자ID..#grostestrecid#
        
        SQL2 = SQL2 & " readdd," '판독일자..#readdd#
        SQL2 = SQL2 & " readtm," '판독시간..#readtm#
        SQL2 = SQL2 & " readid," '판독자ID..#readid#
        SQL2 = SQL2 & " EXTRPARTCNTS," '#extrpartcnts# 검체명(spcm.spcnm)
        SQL2 = SQL2 & " EXTRMTHDCNTS," '#extrmthdcnts# 검사명(test.testengnm)
        SQL2 = SQL2 & " DIAGCNTS," '진단코드.. #diagcd#
        SQL2 = SQL2 & " spckeepflagcd," '검체보관구분코드(0:미보관, 1:보관)..#spckeepflagcd#
        SQL2 = SQL2 & " rslthideflagcd," '결과숨기기구분코드(0:보여주기, 1:숨기기)..#rslthideflagcd#
        SQL2 = SQL2 & " conccaseflagcd," '관심증례구분코드(0:일반Case, 1:관심Case)..#conccaseflagcd#
        SQL2 = SQL2 & " preprsltflagcd," '예비결과구분코드(0:미표시, 1:표시)..#preprsltflagcd#
        
        SQL2 = SQL2 & " rsltrgstid," '결과등록자ID..#rsltrgstid#
        SQL2 = SQL2 & " cnclflagcd," '취소구분코드..'-'
        SQL2 = SQL2 & " cnclresncd," '취소사유코드..'-'
        SQL2 = SQL2 & " cncldd," '취소일자..'-'
        SQL2 = SQL2 & " cncltm," '취소시간..'-'
        SQL2 = SQL2 & " cnclid," '취소자ID..'-'
        SQL2 = SQL2 & " grospic," '육안사진(0:없음, 1:있음)..'0'
        SQL2 = SQL2 & " keybloc," '검체보관구분코드(0:미보관, 1:보관)..'-'
        SQL2 = SQL2 & " tissbloct," '조직은행 블록(T)..'-'
        SQL2 = SQL2 & " tissblocnt," '조직은행 블록(NT)..'-'
        
        SQL2 = SQL2 & " delflagcd," '삭제구분(0:미삭제, 1:삭제)..'0'
        SQL2 = SQL2 & " IMGRGSTDD," '..'-'
        SQL2 = SQL2 & " IMGNO," '..'0'
        SQL2 = SQL2 & " SIGNNO," '..'0'
        SQL2 = SQL2 & " HISTNO," '..'0'
        SQL2 = SQL2 & " PACSNO," '..'0'
        SQL2 = SQL2 & " readgrade," '판독등급..''
        SQL2 = SQL2 & " fstrgstdt," '최초등록일시(시스템정보)
        SQL2 = SQL2 & " fstrgstrid," '최초등록자ID(시스템정보)
        SQL2 = SQL2 & " lastupdtdt," '최종수정일시(시스템정보)
        
        SQL2 = SQL2 & " lastupdtrid," '최종수정자ID(시스템정보)
        SQL2 = SQL2 & " DIAGCD," '진단코드.. ''
        SQL2 = SQL2 & " UGCYALERTFLAGCD)" '진단코드.. '0'
        
        SQL2 = SQL2 & " VALUES('"
        SQL2 = SQL2 & PTNO & "','" ''병리번호.........
        SQL2 = SQL2 & RgtDD & "'," '결과등록일자.........
        SQL2 = SQL2 & Reg_NO & "," '결과등록번호
        SQL2 = SQL2 & " 1," '결과등록이력번호
        SQL2 = SQL2 & " '017','" '기관코드:성빈센트병원(instcd = 017)는 고정입니다.
        SQL2 = SQL2 & PID & "','"  '등록번호
        SQL2 = SQL2 & RgtTM & "'," '결과등록시간
        SQL2 = SQL2 & " '-'," '육안검사기록일자
        SQL2 = SQL2 & " '-'," '육안검사기록시간
        SQL2 = SQL2 & " '-','" '육안검사기록자ID
        
        SQL2 = SQL2 & "-','"  '판독일자
        SQL2 = SQL2 & "-','"  '판독시간
        SQL2 = SQL2 & "-','" '판독자ID
        SQL2 = SQL2 & SPCNM & "','" '#extrpartcnts# 검체명(spcm.spcnm)
        SQL2 = SQL2 & TESTNM & "'," '#extrmthdcnts# 검사명(test.testengnm)
        SQL2 = SQL2 & " ''," '진단코드.. #diagcd#
        SQL2 = SQL2 & " '0'," '검체보관구분코드(0:미보관, 1:보관)
        SQL2 = SQL2 & " '0'," '결과숨기기구분코드(0:보여주기, 1:숨기기)
        SQL2 = SQL2 & " '0'," '관심증례구분코드(0:일반Case, 1:관심Case)
        SQL2 = SQL2 & " '0','" '예비결과구분코드(0:미표시, 1:표시)
        
        SQL2 = SQL2 & USERID & "'," '결과등록자ID
        SQL2 = SQL2 & " '-'," '취소구분코드
        SQL2 = SQL2 & " '-'," '취소사유코드
        SQL2 = SQL2 & " '-'," '취소일자
        SQL2 = SQL2 & " '-'," '취소시간
        SQL2 = SQL2 & " '-'," '취소자ID
        SQL2 = SQL2 & " '0'," '육안사진(0:없음, 1:있음)
        SQL2 = SQL2 & " '-'," '검체보관구분코드(0:미보관, 1:보관)..'-'
        SQL2 = SQL2 & " '-'," '조직은행 블록(T)..'-'
        SQL2 = SQL2 & " '-'," '조직은행 블록(NT)..'-'
        
        SQL2 = SQL2 & " '0'," '판독등급
        SQL2 = SQL2 & " '-'," 'IMGRGSTDD..'-'
        SQL2 = SQL2 & " '0'," 'IMGNO..'0'
        SQL2 = SQL2 & " '0'," 'SIGNNO..'0'
        SQL2 = SQL2 & " '0'," 'HISTNO..'0'
        SQL2 = SQL2 & " '0'," 'PACSNO..'0'
        SQL2 = SQL2 & " ''," ''판독등급..''
        SQL2 = SQL2 & " SYSDATE,'" '최초등록일시(시스템정보)
        SQL2 = SQL2 & USERID & "',"  '최초등록자ID(시스템정보)
        SQL2 = SQL2 & " SYSDATE,'" '최종수정일시(시스템정보)
        
        SQL2 = SQL2 & USERID & "'," '최종수정자ID(시스템정보)
        SQL2 = SQL2 & " ''," 'DIAGCD진단코드.. ''
        SQL2 = SQL2 & " '0')" 'UGCYALERTFLAGCD긴급알림구분코드(0:미알림,1:알림).. '0'
    Else
        SQL2 = ""
    End If

End Function

Function SetJudge(asResult As String, asEquipCode As String)
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
    
    SQL = ""
    SQL = SQL & "SELECT REFLOW, REFHIGH " & vbCr
    SQL = SQL & "  FROM EQPMASTER " & vbCr
    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & "   AND RSLTCHANNEL = '" & sEquipCode & "'"

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
 
    SetJudge = sResFlag
    
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
