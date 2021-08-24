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
        Case "AU680"
                Call Phase_Serial_AU680
        Case "XN1000"
                Call Phase_Serial_XN1000
        Case Else
            
    End Select
    

End Sub

Public Sub TCP_Protocol()

    Select Case UCase(gHOSP.MACHNM)
        Case "BA400"
                Call Phase_TCP_BA400
        Case "VISIONC"
                Call Phase_TCP_VISIONC
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
    Debug.Print strOutput
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
Public Sub GetOrder(ByVal pBarNo As String, ByVal pType As String)

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
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
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
                            pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarNo
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" Then
                            pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarNo
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
        strItems = GetEquipExamCode_AU480(gHOSP.MACHCD, pBarNo, intRow)

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
Public Sub SetPatInfo(ByVal pBarNo As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    With frmMain
        '-- 바코드 사용
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If Trim(GetText(frmMain.spdOrder, i, colBARCODE)) = pBarNo Then
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
                            pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarNo
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
                    For i = 1 To .spdOrder.DataRowCnt
                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
                            pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Check Top
                Case "3"
                    For i = 1 To .spdOrder.DataRowCnt
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" And GetText(frmMain.spdOrder, i, colSTATE) = "" Then
                            pBarNo = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarNo
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
        Call SetText(.spdOrder, mResult.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mResult.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mResult.TubePos, intRow, colPOSNO)
    
        '-- 환자정보 표시
        'Call vasActiveCell(.spdOrder, intRow, colBARCODE)
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
    
    End With
    
    '-- 현재 Row
    gRow = intRow
    
End Sub

'-----------------------------------------------------------------------------'
'   기능 : 오더정보 전송
'-----------------------------------------------------------------------------'
Public Sub SendOrder_E411()
    
    
End Sub


Public Sub Phase_TCP_VISIONC()
    Dim varBuffers  As Variant
    Dim strBuffer   As String
    Dim strLastSeq  As String
    Dim strRcvSign  As String
    Dim strSendAck  As String
    Dim strRcvCnt   As String
    Dim strSendData As String
    Dim strNS       As String
    Dim strNE       As String
    
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long
 
    varBuffers = Split(pBuffer, vbLf)

    For i = 0 To UBound(varBuffers)
        strBuffer = varBuffers(i)
        If strBuffer = "" Then
            Exit For
        End If
        strLastSeq = mGetP(strBuffer, 1, vbTab)
        strRcvSign = mGetP(strBuffer, 2, vbTab)
        
        strSendAck = strLastSeq & vbTab & "ACK"
        
        Select Case UCase(strRcvSign)
            Case "RESULT"
                Call TCPRcvData_VISIONC
                pBuffer = ""
            Case "CONNECT"
                frmMain.wSck.SendData strSendAck & vbLf
                SetRawData "[Tx]" & strSendAck & vbLf
                
            Case "RESULTS"
                strRcvCnt = CInt(mGetP(strBuffer, 3, vbTab))
                
                strNS = strRcvCnt
                strNE = CInt(mGetP(strBuffer, 4, vbTab))
                
                strNS = strNS - strNE
                strNE = strNS + strNE
                
                strSendData = strLastSeq & vbTab & "GET" & vbTab & strNS & vbTab & strNE
                'strSendData = "0" & vbTab & "GET" & vbTab & "0" & vbTab & "0"
                
                frmMain.wSck.SendData strSendData & vbLf
                SetRawData "[Tx]" & strSendData & vbLf
                
'                strNS = mGetP(strTmp, 1, vbTab)
'                strSendData = "-" & strNS & vbTab & "GET" & vbTab & "0" & vbTab & "0"
                
        
        End Select
    Next i
    
 
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



Public Sub Phase_Serial_AU680()
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
                Call SerialRcvData_AU680
            Case Else
                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
        End Select
    Next i

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

Public Sub SerialRcvData_AU680()
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
                        strBarno = Trim$(Mid$(strRcvBuf, 14, 26))
                        strRackNo = Mid(strRcvBuf, 3, 4)
                        strTubePos = Mid(strRcvBuf, 7, 2)
                        strSeq = Trim(Mid(strRcvBuf, 10, 4))
                        
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
                        
                        'strTmp = Mid$(strRcvBuf, 29)
                        strTmp = Mid$(strRcvBuf, 45)
                        
                        'Do While Len(strTmp) >= 11
                        Do While Len(strTmp) >= 10
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

Public Sub TCPRcvData_VISIONC()
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
    Dim varRcvBuf       As Variant

    varRcvBuf = Split(pBuffer, vbLf)
    With frmMain
        For intCnt = 0 To UBound(varRcvBuf)
            strRcvBuf = varRcvBuf(intCnt)
            
            '-- 테스트용 -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- 테스트용 -----------------
            
            If Len(strRcvBuf) > 20 Then
                strBarno = mGetP(strRcvBuf, 7, vbTab)
                strSeq = mGetP(strRcvBuf, 1, vbTab)
                
                With mResult
                    .BarNo = strBarno
                    '.SpcPos = strSeq
                    .Seq = strSeq
                    '.RackNo = strRackNo
                    .TubePos = strTubePos
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                End With
                
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                
                If gRow <= 0 Then
                    Exit Sub
                End If
                
                strIntBase = "ESR"
                strResult = mGetP(strRcvBuf, 9, vbTab) 'ESR
                'strResult = mGetP(strRcvBuf, 10, vbTab) '18
                        
                If strIntBase <> "" And strResult <> "" Then
                    If gPatOrdCd <> "" Then
                        SQL = ""
                        SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
                        SQL = SQL & "  FROM EQPMASTER" & vbCr
                        SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
                        SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
                        SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
                        
                        'Call SetSQLData("검사항목조회", SQL)
                        
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
                        
                        'Call SetSQLData("검사항목조회N", SQL)
                        
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
                    
            End If
        Next
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
        
On Error GoTo ErrHandle

    With frmMain
        SaveTransData = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        
        
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
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                      SQL = " Update SLA_LabResult  " & vbCr
                SQL = SQL & "   Set Result     = '" & sResult & "' " & vbCr
                SQL = SQL & "      ,NormalFlag = '0' " & vbCr
                SQL = SQL & "      ,PanicFlag  = '0' " & vbCr
                SQL = SQL & "      ,DeltaFlag  = '0' " & vbCr
                SQL = SQL & "      ,TransFlag  = '1' " & vbCr
                SQL = SQL & "      ,ResultID   = ''  " & vbCr
                SQL = SQL & "      ,ResultDate = '" & Trim(Format(Now, "yyyy-mm-dd")) & "'" & vbCr
                SQL = SQL & "      ,ResultTime = '" & Trim(Format(Time, "HH:MM:SS")) & "'" & vbCr
                SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "'" & vbCr
                SQL = SQL & "   And OrderCode = '" & strTestCd & "'" & vbCr
                SQL = SQL & "   And LabCode = '" & strTestCd & "'" & vbCr
                SQL = SQL & "   And TransFlag < '2' "

                Call SetSQLData("결과저장", SQL)
                Call DBExec(AdoCn, SQL)
                
                SaveTransData = 1
                
            End If
        Next intRow
        
        If SaveTransData = 1 Then
                  SQL = " Update SLA_LabMaster " & vbCr
            SQL = SQL & "   Set JStatus = '2' " & vbCr
            SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "' " & vbCr
            SQL = SQL & "   And OrderCode = '" & strTestCd & "'" & vbCr
            SQL = SQL & "   And RECEIPTDATE = '" & Format(strHospDate, "yyyy-mm-dd") & "'" & vbCr
            SQL = SQL & "   And JStatus < '3' "
            
            Call SetSQLData("상태저장", SQL)
            Call DBExec(AdoCn, SQL)
            
        End If
        
        'AdoCn.CommitTrans
        
    
    End With

Exit Function

ErrHandle:
    SaveTransData = -1
    'AdoCn.RollbackTrans
    
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
