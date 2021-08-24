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

'    Select Case UCase(gHOSP.MACHNM)
'        Case "E411"
'                Call Phase_Serial_E411
'        Case "AU400"
'                Call Phase_Serial_AU400
'        Case "AU480"
'                Call Phase_Serial_AU480
'        Case "XN1000"
'                Call Phase_Serial_XN1000
'        Case Else
'
'    End Select
    

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
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSPECNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
                Case "2"
'                    For i = 1 To .spdOrder.DataRowCnt
'                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
'                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
'                            intRow = i
'                            Exit For
'                        End If
'                    Next i
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
'        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
'        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
'        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
            
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
                        If Val(Trim(GetText(frmMain.spdOrder, i, colSPECNO))) = Val(mOrder.Seq) Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mOrder.BarNo = pBarno
                            intRow = i
                            Exit For
                        End If
                    Next i
                '-- Rack/Pos
'                Case "2"
'                    For i = 1 To .spdOrder.DataRowCnt
'                        If Trim(GetText(frmMain.spdOrder, i, colRACKNO)) = mOrder.RackNo And Trim(GetText(frmMain.spdOrder, i, colPOSNO)) = mOrder.TubePos Then
'                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
'                            intRow = i
'                            Exit For
'                        End If
'                    Next i
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
        Call SetText(.spdOrder, mResult.BarNo, intRow, colJUBNO)
        
        '-- 결과스프레드 지우기
        .spdResult.MaxRows = 0
    
        '-- 검사자 정보 가져오기
        Call GetSampleInfo(intRow, .spdOrder, "", strState)
        
        
        For i = 1 To UBound(gTEMPLETE)
            .spdOrder.Col = colRELTEST
            .spdOrder.TypeComboBoxString = gTEMPLETE(i)
        Next
    
        .spdOrder.RowHeight(-1) = 24
    
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
'    Dim Buffer      As Variant
'    Dim BufChar     As String
'    Dim lngBufLen   As Long
'    Dim i           As Long
'
'    lngBufLen = Len(pBuffer)
'
'    For i = 1 To lngBufLen
'        BufChar = Mid$(pBuffer, i, 1)
'        Select Case BufChar
'            Case STX
'                intBufCnt = 1
'                Erase strRecvData
'                ReDim Preserve strRecvData(intBufCnt)
'            Case ETB
'            Case ETX
'                Call SerialRcvData_AU400
'            Case Else
'                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'        End Select
'    Next i

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
'    Dim Buffer      As Variant
'    Dim BufChar     As String
'    Dim lngBufLen   As Long
'    Dim i           As Long
'
'    lngBufLen = Len(pBuffer)
'
'    For i = 1 To lngBufLen
'        BufChar = Mid$(Buffer, i, 1)
'        Select Case BufChar
'            Case STX
'                intBufCnt = 1
'                Erase strRecvData
'                ReDim Preserve strRecvData(intBufCnt)
'            Case ETB
'            Case ETX
'                Call SerialRcvData_AU480
'            Case Else
'                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'        End Select
'    Next i

End Sub

'Public Sub Phase_Serial_XN1000()
'    Dim Buffer      As Variant
'    Dim BufChar     As String
'    Dim lngBufLen   As Long
'    Dim i           As Long
'
'    lngBufLen = Len(pBuffer)
'    With frmMain
'        For i = 1 To lngBufLen
'            BufChar = Mid$(pBuffer, i, 1)
'            Select Case intPhase
'                Case 1      '## Estabilshment Phase
'                    Select Case BufChar
'                        Case ENQ
'                            intBufCnt = 1
'                            Erase strRecvData
'                            ReDim Preserve strRecvData(intBufCnt)
'                            intPhase = 2
'                            .comEqp.Output = ACK
'                            DoEvents
'                            SetRawData "[Tx]" & ACK
'                        Case ACK
'                            If strState = "Q" Then Call SendOrder
'
'                    End Select
'                Case 2      '## Transfer Phase
'                    Select Case BufChar
'                        Case ENQ
'                            Erase strRecvData
'                            .comEqp.Output = ACK
'                            DoEvents
'                            SetRawData "[Tx]" & ACK
'                        Case STX
'                            intBufCnt = 1
'                            Erase strRecvData
'                            ReDim Preserve strRecvData(intBufCnt)
'                        Case ETB
'                            blnIsETB = True
'                            intPhase = 3
'                        Case ETX
'                            intBufCnt = intBufCnt + 1
'                            ReDim Preserve strRecvData(intBufCnt)
'                            intPhase = 3
'                        Case vbCr, vbLf
'                        Case EOT
'                            intPhase = 1
'                        Case Else
'                            If blnIsETB = False Then
'                                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'                            Else
'                                blnIsETB = False
'                            End If
'                    End Select
'                Case 3      '## Transfer Phase
'                    Select Case BufChar
'                        Case vbCr
'                        Case vbLf
'                            intPhase = 4
'                            .comEqp.Output = ACK
'                            DoEvents
'                            SetRawData "[Tx]" & ACK
'                    End Select
'                Case 4      '## Termination Phase
'                    Select Case BufChar
'                        Case STX
'                            intPhase = 2
'                        Case EOT
'                            Call SerialRcvData_XN1000
'                            If strState = "Q" Then
'                                intSndPhase = 1
'                                intFrameNo = 1
'                                .comEqp.Output = ENQ
'                                DoEvents
'                                SetRawData "[Tx]" & ENQ
'                            End If
'
'                            intPhase = 1
'                    End Select
'            End Select
'        Next i
'    End With
'
'End Sub



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
    Dim strICVal        As String
    Dim strICVal1       As String
    Dim strICVal2       As String
    
    Dim strTmp          As String
    Dim i               As Integer
    Dim blnSame         As Boolean
    
    strFlag = ""
    blnSame = False
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- 테스트용 -----------------
'            If .fraCommTest.Visible = False Then
'                Call SetSQLData("RCV", strRcvBuf, "A")
'            End If
            '-- 테스트용 -----------------
            
            If intCnt = 1 Then
                strICVal = ""
                strBarno = mGetP(strRcvBuf, 2, ",")
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
                
'                For i = 1 To .spdOrder.DataRowCnt
'                    If strBarno = GetText(frmMain.spdOrder, i, colJUBNO) Then
'                        blnSame = True
'                        Exit For
'                    End If
'                Next
'
'                If blnSame = False Then
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'                End If
                
                If gRow <= 0 Then
                    Exit Sub
                End If
            End If
                            
            strTarget = Trim(mGetP(strRcvBuf, 5, ","))      '-- Target(타입)
            strIntBase = "HPV"
            strResult = Trim(mGetP(strRcvBuf, 6, ","))      '-- 판정(정성)
            strIntResult = Trim(mGetP(strRcvBuf, 7, ","))   '-- 결과(정량)
            
            '-- IC 값 저장
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
            
            Else
                '-- 42,43,54,70,61,6,44,40,11   'Low
                If strTarget = "42" Or strTarget = "43" Or strTarget = "54" Or strTarget = "70" Or strTarget = "61" Or strTarget = "6" Or strTarget = "44" Or strTarget = "40" Or strTarget = "11" Then
                    If strResult <> "-" Then
                        strLGrp = strLGrp & strTarget & strResult & "," 'ex) 42++,54+++,70+++
                    End If
                Else        'High
                    If strResult <> "-" Then
                        strHGrp = strHGrp & strTarget & strResult & ","
                    End If
                End If
            End If
        Next
        
        If strVarTest = "에러" Then
            strResult = ""
            strResult = "INVALID"
            strFlag = "INVALID"
        Else
            '하나라도 양성이 있으면..
            If Len(strHGrp) = 0 And Len(strLGrp) = 0 Then
               strTotFlag = "Negative"
            Else
               strTotFlag = "Positive"
            End If
            
            strResult = strTotFlag
            
            '마지막 ',' 제거
            If Len(strHGrp) > 0 Then
                strHGrp = Mid(strHGrp, 1, Len(strHGrp) - 1)
            End If
            
            If Len(strLGrp) > 0 Then
                strLGrp = Mid(strLGrp, 1, Len(strLGrp) - 1)
            End If
            
            '-- 음성일 경우
            If strTotFlag = "Negative" Then
                          strFlag = "HPV High Risk Type : Negative" & vbNewLine
                strFlag = strFlag & "HPV Low Risk Type : Negative"
            
            '-- 양성일 경우
            ElseIf strTotFlag = "Positive" Then
                If Len(strHGrp) > 0 And Len(strLGrp) > 0 Then
                    strFlag = strFlag & "HPV High Risk Type : Positive (" & strHGrp & ")" & vbNewLine
                    strFlag = strFlag & "HPV Low Risk Type : Positive (" & strLGrp & ")"
                ElseIf Len(strHGrp) > 0 And Len(strLGrp) <= 0 Then
                    strFlag = strFlag & "HPV High Risk Type : Positive (" & strHGrp & ")" & vbNewLine
                    strFlag = strFlag & "HPV Low Risk Type : Negative"
                ElseIf Len(strHGrp) <= 0 And Len(strLGrp) > 0 Then
                    strFlag = strFlag & "HPV High Risk Type : Negative" & vbNewLine
                    strFlag = strFlag & "HPV Low Risk Type : Positive (" & strLGrp & ")"
                End If
            End If
            'IC값을 붙임
            strFlag = strFlag & vbNewLine & "IC " & strICVal1 & "/" & strICVal2
        End If
        
        Call SetSQLData("RCV", "bar:" & strBarno & ",result:" & strFlag & ",order:" & gPatOrdCd, "A")
        
        If strIntBase <> "" And strFlag <> "" Then
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
                    
                    '진행상태 표시("결과")
                    SetText .spdOrder, "결과", gRow, colSTATE
    
                    '결과값 표시
                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                        If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                            SetText .spdOrder, strResult, gRow, colRESULT
                            SetText .spdOrder, strFlag, gRow, intCol
                            '-- IC : 더 큰값을 사용한다
                            If Len(strICVal1) > Len(strICVal2) Then
                                SetText .spdOrder, strICVal1, gRow, colHPVIC
                            Else
                                SetText .spdOrder, strICVal2, gRow, colHPVIC
                            End If
                            
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
                    SetText .spdResult, strResult, lsRstRow, colRMACHRESULT     '장비결과
                    SetText .spdResult, strFlag, lsRstRow, colRLISRESULT          'LIS결과
                    SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                    
                    '-- 로컬 저장
                    SetLocalDB gRow, lsRstRow, "1", ""
                    
                    strState = "R"
                    
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
                    
                    '진행상태 표시("결과")
                    SetText .spdOrder, "결과", gRow, colSTATE
    
                    '결과값 표시
                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
                        If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                            SetText .spdOrder, strResult, gRow, colRESULT
                            SetText .spdOrder, strFlag, gRow, intCol
                            '-- IC : 더 큰값을 사용한다
                            If Len(strICVal1) > Len(strICVal2) Then
                                SetText .spdOrder, strICVal1, gRow, colHPVIC
                            Else
                                SetText .spdOrder, strICVal2, gRow, colHPVIC
                            End If
                            Exit For
                        End If
                    Next
    
                    '-- 결과 List
                    SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '순번
                    SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          '처방코드
                    SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '검사코드
                    SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '검사명
                    SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '장비채널
                    SetText .spdResult, strResult, lsRstRow, colRMACHRESULT     '장비결과
                    SetText .spdResult, strFlag, lsRstRow, colRLISRESULT          'LIS결과
                    SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '판정
                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '참고치
                    
                    '-- 로컬 저장
                    SetLocalDB gRow, lsRstRow, "1", ""
                    
                    If strState <> "R" Then
                        strState = ""
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
    Dim strJubNo        As String
    Dim strRstDate      As String
    Dim strDoctor       As String
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


    Dim prm(101)        As New ADODB.Parameter
    Dim i               As Integer

    Dim varData         As Variant
    Dim varHData        As Variant
    Dim varLData        As Variant
    Dim intINData       As Integer
    Dim intHCnt         As Integer
    Dim intLCnt         As Integer
    Dim strData         As String
    Dim j               As Integer
    Dim k               As Integer
    
On Error GoTo ErrHandle
    
    
    With frmMain
        SaveTransData = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))

        strJubNo = Trim(GetText(.spdOrder, argSpcRow, colJUBNO))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strRstDate = Trim(GetText(.spdOrder, argSpcRow, colRSTDATE))
        strDoctor = Trim(GetText(.spdOrder, argSpcRow, colDOCTOR))
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT TESTCD,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("TESTCD").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 3)
                'Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                'Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                'Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        intHCnt = 0
        intLCnt = 0
        strData = ""
        
        'AdoCn.BeginTrans
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 1))      '검사코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 2))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 3))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If

            varData = Split(sResult2, vbNewLine)
            'varHData =
            
            If strJubNo <> "" And strTestCd <> "" And sResult <> "" Then
                Set AdoCmd = New ADODB.Command
                Set AdoCmd.ActiveConnection = AdoCn
                With AdoCmd
                    .CommandTimeout = 15
                    .CommandText = "p_inter_k2a03Add"
                    .CommandType = adCmdStoredProc
                    
                    If frmMain.optSave(0).Value = True Then
                        Set prm(1) = .CreateParameter("@QUERY", adChar, adParamInput, 1, "A")               '작업구분(A:입력, E:수정)
                    Else
                        Set prm(1) = .CreateParameter("@QUERY", adChar, adParamInput, 1, "E")               '작업구분(A:입력, E:수정)
                    End If
                    Set prm(2) = .CreateParameter("@JUB_NO", adChar, adParamInput, 11, strJubNo)         '접수번호
                    Set prm(3) = .CreateParameter("@CHART_NO", adChar, adParamInput, 7, strChartNo)       '등록번호
                    Set prm(4) = .CreateParameter("@RESULT_RESID", adChar, adParamInput, 5, Mid(strDoctor, 1, 5))
                    Set prm(5) = .CreateParameter("@RESULT_STAFF", adChar, adParamInput, 5, Mid(strDoctor, 1, 5))
                    Set prm(6) = .CreateParameter("@RESULT_STAFF2", adChar, adParamInput, 5, "")
                    Set prm(7) = .CreateParameter("@RESULT_DATE", adChar, adParamInput, 8, Format(strRstDate, "yyyymmdd"))
                    
                    Set prm(8) = .CreateParameter("@M101", adSmallInt, adParamInput, 1, 0)
                    Set prm(9) = .CreateParameter("@M102", adSmallInt, adParamInput, 1, 0)
                    Set prm(10) = .CreateParameter("@M102P", adSmallInt, adParamInput, 1, 0)
                    Set prm(11) = .CreateParameter("@M102N", adSmallInt, adParamInput, 1, 0)
                    Set prm(12) = .CreateParameter("@M103", adSmallInt, adParamInput, 1, 0)
                    Set prm(13) = .CreateParameter("@M103P", adSmallInt, adParamInput, 1, 0)
                    Set prm(14) = .CreateParameter("@M103N", adSmallInt, adParamInput, 1, 0)
                    Set prm(15) = .CreateParameter("@M131", adSmallInt, adParamInput, 1, 0)
                    Set prm(16) = .CreateParameter("@M132", adSmallInt, adParamInput, 1, 0)
                    Set prm(17) = .CreateParameter("@M132P", adSmallInt, adParamInput, 1, 0)
                    Set prm(18) = .CreateParameter("@M132N", adSmallInt, adParamInput, 1, 0)
                    Set prm(19) = .CreateParameter("@M133", adSmallInt, adParamInput, 1, 0)
                    Set prm(20) = .CreateParameter("@M133P", adSmallInt, adParamInput, 1, 0)
                    Set prm(21) = .CreateParameter("@M133N", adSmallInt, adParamInput, 1, 0)
                    Set prm(22) = .CreateParameter("@M134", adSmallInt, adParamInput, 1, 0)
                    Set prm(23) = .CreateParameter("@M134P", adSmallInt, adParamInput, 1, 0)
                    Set prm(24) = .CreateParameter("@M134N", adSmallInt, adParamInput, 1, 0)
                    Set prm(25) = .CreateParameter("@M135", adSmallInt, adParamInput, 1, 0)
                    Set prm(26) = .CreateParameter("@M135P", adSmallInt, adParamInput, 1, 0)
                    Set prm(27) = .CreateParameter("@M135N", adSmallInt, adParamInput, 1, 0)
                    Set prm(28) = .CreateParameter("@M136", adSmallInt, adParamInput, 1, 0)
                    Set prm(29) = .CreateParameter("@M136P", adSmallInt, adParamInput, 1, 0)
                    Set prm(30) = .CreateParameter("@M136N", adSmallInt, adParamInput, 1, 0)
                    Set prm(31) = .CreateParameter("@M137", adSmallInt, adParamInput, 1, 0)
                    Set prm(32) = .CreateParameter("@M137P", adSmallInt, adParamInput, 1, 0)
                    Set prm(33) = .CreateParameter("@M137N", adSmallInt, adParamInput, 1, 0)
                    Set prm(34) = .CreateParameter("@M138", adSmallInt, adParamInput, 1, 0)
                    Set prm(35) = .CreateParameter("@M138P", adSmallInt, adParamInput, 1, 0)
                    Set prm(36) = .CreateParameter("@M138N", adSmallInt, adParamInput, 1, 0)
                    Set prm(37) = .CreateParameter("@M161", adSmallInt, adParamInput, 1, 0)
                    Set prm(38) = .CreateParameter("@M171", adSmallInt, adParamInput, 1, 0)
                    Set prm(39) = .CreateParameter("@M172", adSmallInt, adParamInput, 1, 0)
                    Set prm(40) = .CreateParameter("@M173", adSmallInt, adParamInput, 1, 0)
                    Set prm(41) = .CreateParameter("@M174", adSmallInt, adParamInput, 1, 0)
                    Set prm(42) = .CreateParameter("@M175", adSmallInt, adParamInput, 1, 0)
                    Set prm(43) = .CreateParameter("@M176", adSmallInt, adParamInput, 1, 0)
                    Set prm(44) = .CreateParameter("@M177", adSmallInt, adParamInput, 1, 0)
                    Set prm(45) = .CreateParameter("@M178", adSmallInt, adParamInput, 1, 0)
                    Set prm(46) = .CreateParameter("@M179", adSmallInt, adParamInput, 1, 0)
                    Set prm(47) = .CreateParameter("@M180", adSmallInt, adParamInput, 1, 0)
                    Set prm(48) = .CreateParameter("@M181", adSmallInt, adParamInput, 1, 0)
                    Set prm(49) = .CreateParameter("@M182", adSmallInt, adParamInput, 1, 0)
                    Set prm(50) = .CreateParameter("@M183", adSmallInt, adParamInput, 1, 0)
                    Set prm(51) = .CreateParameter("@M184", adSmallInt, adParamInput, 1, 0)
                    Set prm(52) = .CreateParameter("@M185", adSmallInt, adParamInput, 1, 0)
                    Set prm(53) = .CreateParameter("@M186", adSmallInt, adParamInput, 1, 0)
                    Set prm(54) = .CreateParameter("@M187", adSmallInt, adParamInput, 1, 0)
                    Set prm(55) = .CreateParameter("@M188", adSmallInt, adParamInput, 1, 0)
                    Set prm(56) = .CreateParameter("@M189", adSmallInt, adParamInput, 1, 0)
                    Set prm(57) = .CreateParameter("@M190", adSmallInt, adParamInput, 1, 0)
                    Set prm(58) = .CreateParameter("@M201", adSmallInt, adParamInput, 1, 0)
                    Set prm(59) = .CreateParameter("@M202", adSmallInt, adParamInput, 1, 0)
                    Set prm(60) = .CreateParameter("@M203", adSmallInt, adParamInput, 1, 0)
                    Set prm(61) = .CreateParameter("@M204", adSmallInt, adParamInput, 1, 0)
                    Set prm(62) = .CreateParameter("@M301", adSmallInt, adParamInput, 1, 0)
                    Set prm(63) = .CreateParameter("@M302", adSmallInt, adParamInput, 1, 0)
                    Set prm(64) = .CreateParameter("@M303", adSmallInt, adParamInput, 1, 0)
                    Set prm(65) = .CreateParameter("@M304", adSmallInt, adParamInput, 1, 0)
                    Set prm(66) = .CreateParameter("@M401", adSmallInt, adParamInput, 1, 0)
                    Set prm(67) = .CreateParameter("@M402", adSmallInt, adParamInput, 1, 0)
                    Set prm(68) = .CreateParameter("@M403", adSmallInt, adParamInput, 1, 0)
                    Set prm(69) = .CreateParameter("@M404", adSmallInt, adParamInput, 1, 0)
                    Set prm(70) = .CreateParameter("@M405", adSmallInt, adParamInput, 1, 0)
                    Set prm(71) = .CreateParameter("@M406", adSmallInt, adParamInput, 1, 0)
                    Set prm(72) = .CreateParameter("@M407", adSmallInt, adParamInput, 1, 0)
                    
                    Set prm(73) = .CreateParameter("@MNOTE", adVarChar, adParamInput, 1000, "")
                    
                    
                    intINData = 0
                    If mGetP(varData(0), 1, ":") = "HPV High Risk Type " Then
                        If Trim(UCase(mGetP(varData(0), 2, ":"))) = "NEGATIVE" Then
                            intINData = 0
                        Else
                            intINData = 1
                            varHData = mGetP(varData(0), 2, ":")
                            varHData = mGetP(varHData, 2, "(")
                            varHData = Replace(varHData, ")", "")
                            varHData = Split(varHData, ",")
                            intHCnt = UBound(varHData) + 1
                            
                            For j = 0 To UBound(varHData)
                                strData = strData & varHData(j) & "/"
                            Next
                        End If
                    End If
                    
                    'Set prm(74) = .CreateParameter("@M501", adSmallInt, adParamInput, 1, 0)
                    'Set prm(75) = .CreateParameter("@M5011", adSmallInt, adParamInput, 1, intINData)    '--DNA-Chip - High-Risk Positive
                    
                    If mGetP(varData(1), 1, ":") = "HPV Low Risk Type " Then
                        If Trim(UCase(mGetP(varData(1), 2, ":"))) = "NEGATIVE" Then
                            intINData = 0
                        Else
                            intINData = 1
                            varLData = mGetP(varData(1), 2, ":")
                            varLData = mGetP(varLData, 2, "(")
                            varLData = Replace(varLData, ")", "")
                            varLData = Split(varLData, ",")
                            intLCnt = UBound(varLData) + 1
                            For j = 0 To UBound(varLData)
                                strData = strData & varLData(j) & "/"
                            Next
                            
                        End If
                    End If
                    
                    'Set prm(76) = .CreateParameter("@M5012", adSmallInt, adParamInput, 1, intINData)    '--DNA-Chip - Low-Risk Positive
                    
                    If intHCnt + intLCnt > 1 Then
                        intINData = 1
                    Else
                        intINData = 0
                    End If
                    
'                    Set prm(77) = .CreateParameter("@M5013", adSmallInt, adParamInput, 1, intINData)    '--DNA-Chip - Multiple Infection
'                    Set prm(78) = .CreateParameter("@M5014", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - Other Type Positive
                    
                    Set prm(74) = .CreateParameter("@M501", adSmallInt, adParamInput, 1, 0)             '--Human papilloma virus - Positive
                    
                    If intHCnt > 1 Or intLCnt > 1 Then
                        Set prm(75) = .CreateParameter("@M5011", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - High-Risk Positive
                        Set prm(76) = .CreateParameter("@M5012", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - Low-Risk Positive
                        Set prm(77) = .CreateParameter("@M5013", adSmallInt, adParamInput, 1, 1)    '--DNA-Chip - Multiple Infection
                    Else
                        If intHCnt >= 1 Then
                            Set prm(75) = .CreateParameter("@M5011", adSmallInt, adParamInput, 1, 1)    '--DNA-Chip - High-Risk Positive
                            Set prm(76) = .CreateParameter("@M5012", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - Low-Risk Positive
                            Set prm(77) = .CreateParameter("@M5013", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - Multiple Infection
                        ElseIf intLCnt >= 1 Then
                            Set prm(75) = .CreateParameter("@M5011", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - High-Risk Positive
                            Set prm(76) = .CreateParameter("@M5012", adSmallInt, adParamInput, 1, 1)    '--DNA-Chip - Low-Risk Positive
                            Set prm(77) = .CreateParameter("@M5013", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - Multiple Infection
                        Else
                            Set prm(75) = .CreateParameter("@M5011", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - High-Risk Positive
                            Set prm(76) = .CreateParameter("@M5012", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - Low-Risk Positive
                            Set prm(77) = .CreateParameter("@M5013", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - Multiple Infection
                        End If
                    End If
                    Set prm(78) = .CreateParameter("@M5014", adSmallInt, adParamInput, 1, 0)            '--DNA-Chip - Other Type Positive
                    
                    
                    Set prm(79) = .CreateParameter("@M501T", adVarChar, adParamInput, 100, mGetP(strData, 1, "/"))      '-- HPV type - (1행,1열)
                    Set prm(80) = .CreateParameter("@M502T", adVarChar, adParamInput, 100, mGetP(strData, 2, "/"))      '-- HPV type - (1행,2열)
                    Set prm(81) = .CreateParameter("@M503T", adVarChar, adParamInput, 100, mGetP(strData, 3, "/"))      '-- HPV type - (1행,3열)
                    Set prm(82) = .CreateParameter("@M504T", adVarChar, adParamInput, 100, mGetP(strData, 4, "/"))      '-- HPV type - (1행,4열)
                    Set prm(83) = .CreateParameter("@M505T", adVarChar, adParamInput, 100, mGetP(strData, 5, "/"))      '-- HPV type - (1행,2열)
                    Set prm(84) = .CreateParameter("@M506T", adVarChar, adParamInput, 100, mGetP(strData, 6, "/"))      '-- HPV type - (2행,1열)
                    Set prm(85) = .CreateParameter("@M507T", adVarChar, adParamInput, 100, mGetP(strData, 7, "/"))      '-- HPV type - (2행,2열)
                    Set prm(86) = .CreateParameter("@M508T", adVarChar, adParamInput, 100, mGetP(strData, 8, "/"))      '-- HPV type - (2행,3열)
                    Set prm(87) = .CreateParameter("@M509T", adVarChar, adParamInput, 100, mGetP(strData, 9, "/"))      '-- HPV type - (2행,4열)
                    Set prm(88) = .CreateParameter("@M510T", adVarChar, adParamInput, 100, mGetP(strData, 10, "/"))     '-- HPV type - (2행,5열)

'                    If intHCnt + intLCnt = 0 Then
'                        intINData = 1
'                    Else
'                        intINData = 0
'                    End If
'
'                    Set prm(89) = .CreateParameter("@M502", adSmallInt, adParamInput, 1, intINData)     '--DNA-Chip Negative
                    
                    
                    If intHCnt = 0 And intLCnt = 0 Then
                        Set prm(89) = .CreateParameter("@M502", adSmallInt, adParamInput, 1, 1)     '--DNA-Chip Negative 여부
                    Else
                        Set prm(89) = .CreateParameter("@M502", adSmallInt, adParamInput, 1, 0)     '--DNA-Chip Negative 여부
                    End If
                    
                    Set prm(90) = .CreateParameter("@M503", adSmallInt, adParamInput, 1, 0)             '--DNA-Chip Insufficient Sample
                    
                    Set prm(91) = .CreateParameter("@M511", adSmallInt, adParamInput, 1, 0)
                    Set prm(92) = .CreateParameter("@M512", adSmallInt, adParamInput, 1, 0)
                    Set prm(93) = .CreateParameter("@M513", adSmallInt, adParamInput, 1, 0)
                    Set prm(94) = .CreateParameter("@M514", adSmallInt, adParamInput, 1, 0)
                    Set prm(95) = .CreateParameter("@M515", adSmallInt, adParamInput, 1, 0)
                    Set prm(96) = .CreateParameter("@M516", adSmallInt, adParamInput, 1, 0)
                    
                    Set prm(97) = .CreateParameter("@M517T", adChar, adParamInput, 8, "")
                    Set prm(98) = .CreateParameter("@M518T", adChar, adParamInput, 10, "")
                    Set prm(99) = .CreateParameter("@M519T", adChar, adParamInput, 10, "")
                    Set prm(100) = .CreateParameter("@M520T", adChar, adParamInput, 10, "")
                    Set prm(101) = .CreateParameter("@TAT_SAYU", adChar, adParamInput, 3, "")


                    For k = 1 To 101
                        .Parameters.Append prm(k)
                        Call SetSQLData("서버저장", k & ":" & prm(k), "A")
                    Next

                    
                    .Execute
                    
                End With


                SaveTransData = 1

            End If
        Next intRow
        
    End With

Exit Function

ErrHandle:
    SaveTransData = -1
    
    
    MsgBox Err.Number & vbCr & Err.Description
    
End Function

Function SaveTransDataR(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String

    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strJubNo        As String
    Dim strRstDate      As String
    Dim strDoctor       As String
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


    Dim prm(101)        As New ADODB.Parameter
    Dim i               As Integer

    Dim varData         As Variant
    Dim varHData        As Variant
    Dim varLData        As Variant
    Dim intINData       As Integer
    Dim intHCnt         As Integer
    Dim intLCnt         As Integer
    Dim strData         As String
    Dim j               As Integer
    Dim k               As Integer
    
On Error GoTo ErrHandle
    
    
    With frmMain
        SaveTransDataR = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdROrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdROrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdROrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdROrder, argSpcRow, colBARCODE))

        strJubNo = Trim(GetText(.spdROrder, argSpcRow, colJUBNO))
        strChartNo = Trim(GetText(.spdROrder, argSpcRow, colCHARTNO))
        strRstDate = Trim(GetText(.spdROrder, argSpcRow, colRSTDATE))
        strDoctor = Trim(GetText(.spdROrder, argSpcRow, colDOCTOR))
        
        '-- Local에서 환자별로 결과값 가져오기
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT TESTCD,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '장비코드
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '저장번호
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '검사일
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '바코드
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("TESTCD").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 3)
                'Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                'Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                'Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        intHCnt = 0
        intLCnt = 0
        strData = ""
        
        'AdoCn.BeginTrans
        
        '-- 서버로 결과값 저장하기
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 1))      '검사코드
            sResult1 = Trim(GetText(.vasTemp, intRow, 2))       '결과(장비결과)
            sResult2 = Trim(GetText(.vasTemp, intRow, 3))       '결과(수정결과)
                        
            '-- 장비결과적용
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If

            varData = Split(sResult2, vbNewLine)
            'varHData =
            
            If strJubNo <> "" And strTestCd <> "" And sResult <> "" Then
                Set AdoCmd = New ADODB.Command
                Set AdoCmd.ActiveConnection = AdoCn
                With AdoCmd
                    .CommandTimeout = 15
                    .CommandText = "p_inter_k2a03Add"
                    .CommandType = adCmdStoredProc
                    If frmMain.optSave(0).Value = True Then
                        Set prm(1) = .CreateParameter("@QUERY", adChar, adParamInput, 1, "A")               '작업구분(A:입력, E:수정)
                    Else
                        Set prm(1) = .CreateParameter("@QUERY", adChar, adParamInput, 1, "E")               '작업구분(A:입력, E:수정)
                    End If
                    Set prm(2) = .CreateParameter("@JUB_NO", adChar, adParamInput, 11, strJubNo)         '접수번호
                    Set prm(3) = .CreateParameter("@CHART_NO", adChar, adParamInput, 7, strChartNo)       '등록번호
                    Set prm(4) = .CreateParameter("@RESULT_RESID", adChar, adParamInput, 5, Mid(strDoctor, 1, 5))
                    Set prm(5) = .CreateParameter("@RESULT_STAFF", adChar, adParamInput, 5, Mid(strDoctor, 1, 5))
                    Set prm(6) = .CreateParameter("@RESULT_STAFF2", adChar, adParamInput, 5, "")
                    Set prm(7) = .CreateParameter("@RESULT_DATE", adChar, adParamInput, 8, Format(strRstDate, "yyyymmdd"))
                    
                    Set prm(8) = .CreateParameter("@M101", adSmallInt, adParamInput, 1, 0)
                    Set prm(9) = .CreateParameter("@M102", adSmallInt, adParamInput, 1, 0)
                    Set prm(10) = .CreateParameter("@M102P", adSmallInt, adParamInput, 1, 0)
                    Set prm(11) = .CreateParameter("@M102N", adSmallInt, adParamInput, 1, 0)
                    Set prm(12) = .CreateParameter("@M103", adSmallInt, adParamInput, 1, 0)
                    Set prm(13) = .CreateParameter("@M103P", adSmallInt, adParamInput, 1, 0)
                    Set prm(14) = .CreateParameter("@M103N", adSmallInt, adParamInput, 1, 0)
                    Set prm(15) = .CreateParameter("@M131", adSmallInt, adParamInput, 1, 0)
                    Set prm(16) = .CreateParameter("@M132", adSmallInt, adParamInput, 1, 0)
                    Set prm(17) = .CreateParameter("@M132P", adSmallInt, adParamInput, 1, 0)
                    Set prm(18) = .CreateParameter("@M132N", adSmallInt, adParamInput, 1, 0)
                    Set prm(19) = .CreateParameter("@M133", adSmallInt, adParamInput, 1, 0)
                    Set prm(20) = .CreateParameter("@M133P", adSmallInt, adParamInput, 1, 0)
                    Set prm(21) = .CreateParameter("@M133N", adSmallInt, adParamInput, 1, 0)
                    Set prm(22) = .CreateParameter("@M134", adSmallInt, adParamInput, 1, 0)
                    Set prm(23) = .CreateParameter("@M134P", adSmallInt, adParamInput, 1, 0)
                    Set prm(24) = .CreateParameter("@M134N", adSmallInt, adParamInput, 1, 0)
                    Set prm(25) = .CreateParameter("@M135", adSmallInt, adParamInput, 1, 0)
                    Set prm(26) = .CreateParameter("@M135P", adSmallInt, adParamInput, 1, 0)
                    Set prm(27) = .CreateParameter("@M135N", adSmallInt, adParamInput, 1, 0)
                    Set prm(28) = .CreateParameter("@M136", adSmallInt, adParamInput, 1, 0)
                    Set prm(29) = .CreateParameter("@M136P", adSmallInt, adParamInput, 1, 0)
                    Set prm(30) = .CreateParameter("@M136N", adSmallInt, adParamInput, 1, 0)
                    Set prm(31) = .CreateParameter("@M137", adSmallInt, adParamInput, 1, 0)
                    Set prm(32) = .CreateParameter("@M137P", adSmallInt, adParamInput, 1, 0)
                    Set prm(33) = .CreateParameter("@M137N", adSmallInt, adParamInput, 1, 0)
                    Set prm(34) = .CreateParameter("@M138", adSmallInt, adParamInput, 1, 0)
                    Set prm(35) = .CreateParameter("@M138P", adSmallInt, adParamInput, 1, 0)
                    Set prm(36) = .CreateParameter("@M138N", adSmallInt, adParamInput, 1, 0)
                    Set prm(37) = .CreateParameter("@M161", adSmallInt, adParamInput, 1, 0)
                    Set prm(38) = .CreateParameter("@M171", adSmallInt, adParamInput, 1, 0)
                    Set prm(39) = .CreateParameter("@M172", adSmallInt, adParamInput, 1, 0)
                    Set prm(40) = .CreateParameter("@M173", adSmallInt, adParamInput, 1, 0)
                    Set prm(41) = .CreateParameter("@M174", adSmallInt, adParamInput, 1, 0)
                    Set prm(42) = .CreateParameter("@M175", adSmallInt, adParamInput, 1, 0)
                    Set prm(43) = .CreateParameter("@M176", adSmallInt, adParamInput, 1, 0)
                    Set prm(44) = .CreateParameter("@M177", adSmallInt, adParamInput, 1, 0)
                    Set prm(45) = .CreateParameter("@M178", adSmallInt, adParamInput, 1, 0)
                    Set prm(46) = .CreateParameter("@M179", adSmallInt, adParamInput, 1, 0)
                    Set prm(47) = .CreateParameter("@M180", adSmallInt, adParamInput, 1, 0)
                    Set prm(48) = .CreateParameter("@M181", adSmallInt, adParamInput, 1, 0)
                    Set prm(49) = .CreateParameter("@M182", adSmallInt, adParamInput, 1, 0)
                    Set prm(50) = .CreateParameter("@M183", adSmallInt, adParamInput, 1, 0)
                    Set prm(51) = .CreateParameter("@M184", adSmallInt, adParamInput, 1, 0)
                    Set prm(52) = .CreateParameter("@M185", adSmallInt, adParamInput, 1, 0)
                    Set prm(53) = .CreateParameter("@M186", adSmallInt, adParamInput, 1, 0)
                    Set prm(54) = .CreateParameter("@M187", adSmallInt, adParamInput, 1, 0)
                    Set prm(55) = .CreateParameter("@M188", adSmallInt, adParamInput, 1, 0)
                    Set prm(56) = .CreateParameter("@M189", adSmallInt, adParamInput, 1, 0)
                    Set prm(57) = .CreateParameter("@M190", adSmallInt, adParamInput, 1, 0)
                    Set prm(58) = .CreateParameter("@M201", adSmallInt, adParamInput, 1, 0)
                    Set prm(59) = .CreateParameter("@M202", adSmallInt, adParamInput, 1, 0)
                    Set prm(60) = .CreateParameter("@M203", adSmallInt, adParamInput, 1, 0)
                    Set prm(61) = .CreateParameter("@M204", adSmallInt, adParamInput, 1, 0)
                    Set prm(62) = .CreateParameter("@M301", adSmallInt, adParamInput, 1, 0)
                    Set prm(63) = .CreateParameter("@M302", adSmallInt, adParamInput, 1, 0)
                    Set prm(64) = .CreateParameter("@M303", adSmallInt, adParamInput, 1, 0)
                    Set prm(65) = .CreateParameter("@M304", adSmallInt, adParamInput, 1, 0)
                    Set prm(66) = .CreateParameter("@M401", adSmallInt, adParamInput, 1, 0)
                    Set prm(67) = .CreateParameter("@M402", adSmallInt, adParamInput, 1, 0)
                    Set prm(68) = .CreateParameter("@M403", adSmallInt, adParamInput, 1, 0)
                    Set prm(69) = .CreateParameter("@M404", adSmallInt, adParamInput, 1, 0)
                    Set prm(70) = .CreateParameter("@M405", adSmallInt, adParamInput, 1, 0)
                    Set prm(71) = .CreateParameter("@M406", adSmallInt, adParamInput, 1, 0)
                    Set prm(72) = .CreateParameter("@M407", adSmallInt, adParamInput, 1, 0)
                    Set prm(73) = .CreateParameter("@MNOTE", adVarChar, adParamInput, 1000, "")
                    
                    
                    intINData = 0
                    If mGetP(varData(0), 1, ":") = "HPV High Risk Type " Then
                        If Trim(UCase(mGetP(varData(0), 2, ":"))) = "NEGATIVE" Then
                            intINData = 0
                        Else
                            intINData = 1
                            varHData = mGetP(varData(0), 2, ":")
                            varHData = mGetP(varHData, 2, "(")
                            varHData = Replace(varHData, ")", "")
                            varHData = Split(varHData, ",")
                            intHCnt = UBound(varHData) + 1
                            
                            For j = 0 To UBound(varHData)
                                strData = strData & varHData(j) & "/"
                            Next
                        End If
                    End If
                    
                    
                    If mGetP(varData(1), 1, ":") = "HPV Low Risk Type " Then
                        If Trim(UCase(mGetP(varData(1), 2, ":"))) = "NEGATIVE" Then
                            intINData = 0
                        Else
                            intINData = 1
                            varLData = mGetP(varData(1), 2, ":")
                            varLData = mGetP(varLData, 2, "(")
                            varLData = Replace(varLData, ")", "")
                            varLData = Split(varLData, ",")
                            intLCnt = UBound(varLData) + 1
                            For j = 0 To UBound(varLData)
                                strData = strData & varLData(j) & "/"
                            Next
                            
                        End If
                    End If
                    
                    
                    If intHCnt + intLCnt > 1 Then
                        intINData = 1
                    Else
                        intINData = 0
                    End If
                    
                    
                    Set prm(74) = .CreateParameter("@M501", adSmallInt, adParamInput, 1, 0)             '--Human papilloma virus - Positive
                    
                    If intHCnt > 1 Or intLCnt > 1 Then
                        Set prm(75) = .CreateParameter("@M5011", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - High-Risk Positive
                        Set prm(76) = .CreateParameter("@M5012", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - Low-Risk Positive
                        Set prm(77) = .CreateParameter("@M5013", adSmallInt, adParamInput, 1, 1)    '--DNA-Chip - Multiple Infection
                    Else
                        If intHCnt >= 1 Then
                            Set prm(75) = .CreateParameter("@M5011", adSmallInt, adParamInput, 1, 1)    '--DNA-Chip - High-Risk Positive
                            Set prm(76) = .CreateParameter("@M5012", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - Low-Risk Positive
                            Set prm(77) = .CreateParameter("@M5013", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - Multiple Infection
                        ElseIf intLCnt >= 1 Then
                            Set prm(75) = .CreateParameter("@M5011", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - High-Risk Positive
                            Set prm(76) = .CreateParameter("@M5012", adSmallInt, adParamInput, 1, 1)    '--DNA-Chip - Low-Risk Positive
                            Set prm(77) = .CreateParameter("@M5013", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - Multiple Infection
                        Else
                            Set prm(75) = .CreateParameter("@M5011", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - High-Risk Positive
                            Set prm(76) = .CreateParameter("@M5012", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - Low-Risk Positive
                            Set prm(77) = .CreateParameter("@M5013", adSmallInt, adParamInput, 1, 0)    '--DNA-Chip - Multiple Infection
                        End If
                    End If
                    Set prm(78) = .CreateParameter("@M5014", adSmallInt, adParamInput, 1, 0)            '--DNA-Chip - Other Type Positive
                    
                    
                    Set prm(79) = .CreateParameter("@M501T", adVarChar, adParamInput, 100, mGetP(strData, 1, "/"))      '-- HPV type - (1행,1열)
                    Set prm(80) = .CreateParameter("@M502T", adVarChar, adParamInput, 100, mGetP(strData, 2, "/"))      '-- HPV type - (1행,2열)
                    Set prm(81) = .CreateParameter("@M503T", adVarChar, adParamInput, 100, mGetP(strData, 3, "/"))      '-- HPV type - (1행,3열)
                    Set prm(82) = .CreateParameter("@M504T", adVarChar, adParamInput, 100, mGetP(strData, 4, "/"))      '-- HPV type - (1행,4열)
                    Set prm(83) = .CreateParameter("@M505T", adVarChar, adParamInput, 100, mGetP(strData, 5, "/"))      '-- HPV type - (1행,2열)
                    Set prm(84) = .CreateParameter("@M506T", adVarChar, adParamInput, 100, mGetP(strData, 6, "/"))      '-- HPV type - (2행,1열)
                    Set prm(85) = .CreateParameter("@M507T", adVarChar, adParamInput, 100, mGetP(strData, 7, "/"))      '-- HPV type - (2행,2열)
                    Set prm(86) = .CreateParameter("@M508T", adVarChar, adParamInput, 100, mGetP(strData, 8, "/"))      '-- HPV type - (2행,3열)
                    Set prm(87) = .CreateParameter("@M509T", adVarChar, adParamInput, 100, mGetP(strData, 9, "/"))      '-- HPV type - (2행,4열)
                    Set prm(88) = .CreateParameter("@M510T", adVarChar, adParamInput, 100, mGetP(strData, 10, "/"))     '-- HPV type - (2행,5열)

                    If intHCnt = 0 And intLCnt = 0 Then
                        Set prm(89) = .CreateParameter("@M502", adSmallInt, adParamInput, 1, 1)     '--DNA-Chip Negative 여부
                    Else
                        Set prm(89) = .CreateParameter("@M502", adSmallInt, adParamInput, 1, 0)     '--DNA-Chip Negative 여부
                    End If
                    
                    'Set prm(89) = .CreateParameter("@M502", adSmallInt, adParamInput, 1, intINData)     '--DNA-Chip Negative
                    Set prm(90) = .CreateParameter("@M503", adSmallInt, adParamInput, 1, 0)             '--DNA-Chip Insufficient Sample
                    
                    Set prm(91) = .CreateParameter("@M511", adSmallInt, adParamInput, 1, 0)
                    Set prm(92) = .CreateParameter("@M512", adSmallInt, adParamInput, 1, 0)
                    Set prm(93) = .CreateParameter("@M513", adSmallInt, adParamInput, 1, 0)
                    Set prm(94) = .CreateParameter("@M514", adSmallInt, adParamInput, 1, 0)
                    Set prm(95) = .CreateParameter("@M515", adSmallInt, adParamInput, 1, 0)
                    Set prm(96) = .CreateParameter("@M516", adSmallInt, adParamInput, 1, 0)
                    
                    Set prm(97) = .CreateParameter("@M517T", adChar, adParamInput, 8, "")
                    Set prm(98) = .CreateParameter("@M518T", adChar, adParamInput, 10, "")
                    Set prm(99) = .CreateParameter("@M519T", adChar, adParamInput, 10, "")
                    Set prm(100) = .CreateParameter("@M520T", adChar, adParamInput, 10, "")
                    Set prm(101) = .CreateParameter("@TAT_SAYU", adChar, adParamInput, 3, "")


                    For k = 1 To 101
                        .Parameters.Append prm(k)
                        Call SetSQLData("서버저장", k & ":" & prm(k), "A")
                    Next

                    
                    .Execute
                    
                End With


                SaveTransDataR = 1

            End If
        Next intRow
        
    End With

Exit Function

ErrHandle:
    SaveTransDataR = -1
    
    
    MsgBox Err.Number & vbCr & Err.Description
                
                
End Function

'/*  결과순번 일련번호 채번을 위해 기준 Row를 Lock처리한다.
'himed/his/lis/lib/plgycommonmgt/dao/sqls/plgycommondao_sqls.xml setseqnolock
'arg_seqflagcd = '4'  순번발생구분코드(검사결과) 고정입니다.
'param=[017, 4] */
Private Function Regist_Result_Step1() As Boolean

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
    
    '-- 저장
    strRsltCmt = "[Methods]" & vbCr
    strRsltCmt = strRsltCmt & "   Seegene HPV Real-time PCR (Anyplex II HPV 28 Detection Real-time PCR)" & vbCr
    strRsltCmt = strRsltCmt & vbCr
    strRsltCmt = strRsltCmt & "[Result]" & vbCr
    strRsltCmt = strRsltCmt & pRESULT

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
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrslt", adVarChar, adParamInput, 1000, strRsltCmt)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrsltxml", adVarChar, adParamInput, 1000, "")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrsltetc", adVarChar, adParamInput, 1000, "")

    AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)

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
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)    ':arg_instcd
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)            ':arg_ptno
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)                   ':arg_pid
    
    
    AdoCmd.Execute
    
    Set AdoCmd = Nothing
    
    '-- 저장
    strRsltCmt = "[Methods]" & vbCr
    strRsltCmt = strRsltCmt & "   Seegene HPV Real-time PCR (Anyplex II HPV 28 Detection Real-time PCR)" & vbCr
    strRsltCmt = strRsltCmt & vbCr
    strRsltCmt = strRsltCmt & "[Result]" & vbCr
    strRsltCmt = strRsltCmt & pRESULT
    
    strCmtCnt = "인유두종 바이러스 (Human papilloma virus)는 자궁경부암의 주요 위험인자로 알려져 있습니다. 암과의 관련성 정도에 따라 고위험군 (high risk)과 저위험군 (low risk)로 구분되며, 저위험군 HPV는 대개 시간이 지나면 소실되거나 사마귀 등의 양성변변의 원인이 되는 반면, 고위험군은 자궁경부암을 유발시키는데 관여합니다 (N Engl J Med. 2003 348:518)." & vbCr
    strCmtCnt = strCmtCnt & "" & vbCr
    strCmtCnt = strCmtCnt & "◆ 본 제품은 19종의 고위험군 HPV (16, 18, 26, 31, 33, 25, 29, 45, 51, 52, 53, 56, 58, 59, 66, 68, 73, 82형)과 9종의 저위험군 HPV (6, 11, 40, 42, 43, 44, 54, 61, 70형), 내부대조군의 타켓핵산을 검출합니다." & vbCr
    strCmtCnt = strCmtCnt & "◆ Viral load에서 +++:10^5 copies/reaction, ++:10^5~10^2 copies/reaction, +:10^2 copies/reaction의 농도로 해석될 수 있습니다. 이 중 “+”는 매우 낮은 농도로 감염 시기, 검체 채취 상태에 따라 반복 검사 시 재현되지 않을 수 있습니다." & vbCr
    strCmtCnt = strCmtCnt & "◆ PCR 검사는 검체 내 균수가 적거나 부적절한 검체 의석 또는 증폭 억제물질이 존재하는 경우 위음성이 나올 수 있습니다. 또한, PCR 검사는 유전자 유무를 검사하므로 생존균과 사균의 구분이 안되어 위양성의 가능성이 있습니다. 결과 해석 시 임상 양상과 연관지어 판단하시기 바랍니다." & vbCr
    strCmtCnt = strCmtCnt & "◆ 상기 검사는 검사 방법, 시약의 정도관리 및 검사 결과가 병리과 전문의에 의해 확인되었습니다." & vbCr
    strCmtCnt = strCmtCnt & "   (검사 담당자: " & gHOSP.USERNM & ")"
    
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
Private Function Regist_Result_Detail2(ByVal pPTNO As String, ByVal pREGDT As String, ByVal pPTSEQ As String, ByVal pPID As String, ByVal pRESULT As String, Optional ByVal pACPTNO As String) As String
    Dim strRsltCmt  As String
    
    
On Error GoTo DBErr

    Regist_Result_Detail2 = -1
    
    strRsltCmt = "[Methods]" & vbCr
    strRsltCmt = strRsltCmt & "   Seegene HPV Real-time PCR (Anyplex II HPV 28 Detection Real-time PCR)" & vbCr
    strRsltCmt = strRsltCmt & vbCr
    strRsltCmt = strRsltCmt & "[Result]" & vbCr
    strRsltCmt = strRsltCmt & pRESULT
    
    SQL = ""
    SQL = SQL & "INSERT INTO lis.lprmtrlt"
    SQL = SQL & "       (ptno,         rsltrgstdd,      rsltrgstno,           rsltrgsthistno,"
    SQL = SQL & "       riskflagcd,    instcd,"
    SQL = SQL & "       acptdd,        acptno,          testcd,"
    
    SQL = SQL & "       acptitemno,    testrslt,        testrsltxml,  testrsltetc, delflagcd,"
    SQL = SQL & "       fstrgstdt,     fstrgstrid,"
    SQL = SQL & "       lastupdtdt,    lastupdtrid)"
    
'    SQL = SQL & "VALUES (?,  ?, CAST(? AS DECIMAL(12,0)), 1,"
    SQL = SQL & "VALUES (?,  ?, CAST(? AS DECIMAL(12,0)), " & vbCr
    SQL = SQL & "(SELECT MAX(z.rsltrgsthistno)+1" & vbCr
    SQL = SQL & "   FROM lis.lprmrslt z" & vbCr
    SQL = SQL & "  WHERE instcd         = ?" & vbCr
    SQL = SQL & "    AND ptno           = ?" & vbCr
    SQL = SQL & "    AND pid            = ?" & vbCr
    SQL = SQL & ")," & vbCr
    SQL = SQL & "        ?,  ?,"
    SQL = SQL & "        ?,  CAST(? AS DECIMAL(12,0)),   ?,"
    SQL = SQL & "        CAST(? AS SMALLINT),  ?,     ?  , ?  ,   '0',"
    SQL = SQL & "        SYSDATE, ?,"
    SQL = SQL & "        SYSDATE, ?)"

    
    Call SetSQLData("결과저장", SQL)

    Set AdoCmd = New ADODB.Command
    Set AdoCmd.ActiveConnection = AdoCn


    AdoCmd.CommandType = adCmdText
    AdoCmd.CommandText = SQL
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstdd", adVarChar, adParamInput, 1000, pREGDT)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("rsltrgstno", adVarChar, adParamInput, 1000, pPTSEQ)
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("ptno", adVarChar, adParamInput, 1000, pPTNO)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("pid", adVarChar, adParamInput, 1000, pPID)
    
    
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("riskflagcd", adVarChar, adParamInput, 1000, 9)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("instcd", adVarChar, adParamInput, 1000, gHOSP.HOSPCD)
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptdd", adVarChar, adParamInput, 1000, pREGDT)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptno", adVarChar, adParamInput, 1000, pACPTNO)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testcd", adVarChar, adParamInput, 1000, gAllTestCd1)
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("acptitemno", adVarChar, adParamInput, 1000, 1)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrslt", adVarChar, adParamInput, 1000, strRsltCmt)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrsltxml", adVarChar, adParamInput, 1000, "")
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("testrsltetc", adVarChar, adParamInput, 1000, "")
    
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("fstrgstrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
    AdoCmd.Parameters.Append AdoCmd.CreateParameter("lastupdtrid", adVarChar, adParamInput, 1000, gHOSP.USERID)
    
    
    AdoCmd.Execute
    
    Set AdoCmd = Nothing
    
    Regist_Result_Detail2 = 1
    
Exit Function

DBErr:

    MsgBox Err.Description
    
    
End Function

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
