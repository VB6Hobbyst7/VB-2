Attribute VB_Name = "modComm"
Option Explicit

'===============================
Public Const STX As String = ""
Public Const ETX As String = ""
Public Const ENQ As String = ""
Public Const ACK As String = ""
Public Const NAK As String = ""
Public Const EOT As String = ""
Public Const ETB As String = ""
Public Const FS  As String = ""
Public Const Rst As String = ""
Public Const GS  As String = ""


Public strRecvData()    As String
Public intPhase         As Integer
Public strState         As String
Public intBufCnt        As Integer
Public blnIsETB         As Boolean
Public intSndPhase      As Integer
Public intFrameNo       As Integer
'===============================

Public gRow As Long

Public Sub CommDefine(ByVal pBuffer As Variant)
    
    Select Case gCommProtocol
        Case "1"    ' ASTM
            Call EditRcvData_ASTM(pBuffer)
        Case "2"    ' AU
            Call CommDefine_AU(pBuffer)
        Case "3"    ' XE
        Case "4"
        Case "5"
        Case "6"
    End Select
    
End Sub


Private Sub EditRcvData_ASTM(ByVal pBuffer As Variant)

End Sub



Private Sub CommDefine_AU(ByVal pBuffer As Variant)
    Dim lngBufLen   As Long
    Dim lngBufChar  As Long
    Dim BufChar     As String
    
    lngBufLen = Len(pBuffer)
    
    For lngBufChar = 1 To lngBufLen
        BufChar = Mid$(pBuffer, lngBufChar, 1)
        Select Case BufChar
            Case STX
                intBufCnt = 1
                Erase strRecvData
                ReDim Preserve strRecvData(intBufCnt)
            Case ETB
            Case ETX
'                Call EditRcvData_AU
            Case Else
                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
        End Select
    Next

End Sub
'
''===========================================================================
''-- AU exam receive data
''R 003201 0018          1013001917
''S 003201 0018          1013001917    E      13
''===========================================================================
'Private Sub EditRcvData_AU()
'    Dim strRcvBuf    As String   '수신한 Data
'    Dim strType      As String   '수신한 Record Type
'    Dim strBarno     As String   '수신한 바코드번호
'    Dim strSeq       As String   '수신한 Sequence
'    Dim strRackNo    As String   '수신한 Rack Or Disk No
'    Dim strTubePos   As String   '수신한 Tube Position
'    Dim strIntBase   As String   '수신한 장비기준 검사명
'    Dim strResult    As String   '수신한 결과
'    Dim strQualRslt  As String   '정성 결과
'    Dim strDoseRslt  As String   '정량 결과
'    Dim strQCResult  As String   '수신한 결과(QC)
'    Dim strFlag      As String   '수신한 Abnormal Flag
'    Dim strComment   As String   '수신한 Comment
'    Dim strTmp       As String
'
'    Dim strTemp1     As String
'    Dim strTemp2     As String
'    Dim intCnt       As Integer
'
'    Dim lsExamCode   As String
'    Dim lsExamName   As String
'    Dim lsResult     As String
'
'    Dim lsSeqNo As String
'    Dim lsExamDate As String
'    Dim lsEquipRes As String
'    Dim lsResRow    As String
'    Dim ii As Integer
''    Dim blnPSA       As Boolean
''    Dim blnfPSA      As Boolean
''    Dim strPSA       As String
''    Dim strfPSA      As String
'    Dim intIdx      As Integer
'
'    For intCnt = 1 To UBound(strRecvData)
'        strRcvBuf = strRecvData(intCnt)
'        strType = Mid$(strRcvBuf, 1, 2)
'
'        Select Case strType
'            Case "R "    '## Inquiry Order
'                strBarno = Trim(Mid(strRcvBuf, 14, 20))
'                strRackNo = Mid(strRcvBuf, 3, 4)
'                strTubePos = Mid(strRcvBuf, 7, 2)
'                strSeq = Mid(strRcvBuf, 9, 5)
'
'                If strBarno = "" And strSeq = "" Then
'                    Exit Sub
'                End If
'
'                '-- 오더객체에 담기
'                With mOrder
'                    .BarNo = strBarno
'                    .RackNo = strRackNo
'                    .TubePos = strTubePos
'                    .Seq = strSeq
'                End With
'
'                '-- 환자장보화면 표시및 오더 가져오기
'                Call GetOrder(strBarno)
'
'            Case "D "    '## Result
'                strBarno = Trim$(Mid$(strRcvBuf, 14, 10))
'                strRackNo = Mid(strRcvBuf, 3, 4)
'                strTubePos = Mid(strRcvBuf, 7, 2)
'                strSeq = Mid(strRcvBuf, 9, 5)
'
'                If strBarno = "" And strSeq = "" Then
'                    Exit Sub
'                End If
'
'                '-- 결과객체에 담기
'                With mResult
'                    .BarNo = strBarno
'                    .RackNo = strRackNo
'                    .TubePos = strTubePos
'                    .Seq = strSeq
'                End With
'
'                '-- 결과 시작위치
'                strTmp = Mid$(strRcvBuf, 29)
'
'                '-- 환자장보화면 표시
'                '   상태가 오더일 경우 같은자리
'                '   결과일 경우 MaxRows + 1 or 덮어쓰기
'                Call SetPatInfo(strBarno)
'
'                Do While Len(strTmp) >= 11
'
'                    strIntBase = Mid$(strTmp, 2, 2)
'                    strResult = Mid$(strTmp, 4, 6)
'                    strComment = Mid$(strTmp, 10, 1)
'
'                    If strResult <> "" Then
'                        SQL = ""
'                        SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
'                        SQL = SQL & "  FROM EQUIPEXAM"
'                        SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
'                        SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
'                        SQL = SQL & "   AND EXAMCODE in (" & gOrderExam & ") "
'
'                        Res = GetDBSelectColumn(gLocal, SQL)
'
'                        '-- 오더 없을경우 로컬에서 찾는다
'                        If Res <= 0 Then
'                            SQL = SQL & "SELECT EXAMCODE,EXAMNAME,SEQNO "
'                            SQL = SQL & "  FROM EQUIPEXAM"
'                            SQL = SQL & " WHERE EQUIPNO = '" & gEquip & "' "
'                            SQL = SQL & "   AND EQUIPCODE = '" & strIntBase & "' "
'
'                            Res = GetDBSelectColumn(gLocal, SQL)
'                        End If
'
'                        '-- 로컬에도 등록안된 검사일 경우는 Skip
'                        If Res > 0 Then
'                            lsExamCode = Trim(gReadBuf(0))
'                            lsExamName = Trim(gReadBuf(1))
'                            lsSeqNo = Trim(gReadBuf(2))
'
'                            lsResRow = frmInterface.spdResult.DataRowCnt + 1
'                            If frmInterface.spdResult.MaxRows < lsResRow Then
'                                frmInterface.spdResult.MaxRows = lsResRow
'                            End If
'
'                            '소수점 처리, 결과 형태 처리
'                            lsEquipRes = strResult
'                            lsResult = SetResult(strResult, strIntBase)
'
'                            '-- Order List
'                            SetText frmInterface.spdOrder, "Result", gRow, colState                 '진행상태
'
'                            '-- Result List
''                            SetText frmInterface.spdResult, lsExamName, lsResRow, colExamName       '검사명
'                            SetText frmInterface.spdResult, lsExamCode, lsResRow, colTestCd         '검사코드
'                            SetText frmInterface.spdResult, strIntBase, lsResRow, colChannel        '장비코드(채널)
'                            SetText frmInterface.spdResult, strResult, lsResRow, colEqpResult       '장비결과
'                            SetText frmInterface.spdResult, lsResult, lsResRow, colLisResult        'LIS결과
''                            SetText frmInterface.spdResult, lsSeqNo, lsResRow, colSeq               '순번
''                            SetText frmInterface.spdResult, strComm, lsResRow, 7                    'Flag
'                            '-- 로컬 저장
'                            SetLocalDB gRow, lsResRow, "1", lsEquipRes
'
'                            lsResult_Buff = ""
'                        End If
'
'                    End If
'                    strTmp = Mid$(strTmp, 12)
'                Loop
'                strState = "R"
'
'                If MnTransAuto.Checked = True Then
'
'                    Res = SaveTransDataW(gRow)
'
'                    If Res = -1 Then
'                        '-- 저장 실패
'                        SetForeColor spdOrder, gRow, gRow, 1, colState, 255, 0, 0
'                        SetText spdOrder, "Failed", gRow, colState
'                    Else
'                        '-- 저장 성공
'                        SetBackColor spdOrder, gRow, gRow, 1, colState, 202, 255, 112
'                        SetText spdOrder, "Trans", gRow, colState
'
'                        SQL = " Update pat_res Set " & vbCrLf & _
'                              " sendflag = '2' " & vbCrLf & _
'                              " Where equipno = '" & gEquip & "' " & vbCrLf & _
'                              " And barcode = '" & Trim(GetText(spdOrder, gRow, colBarcode)) & "' "
'                        Res = SendQuery(gLocal, SQL)
'                        If Res = -1 Then
'                            SaveQuery SQL
'                            Exit Sub
'                        End If
'                    End If
'                End If
'
'                SetText spdOrder, "Result", gRow, colState
'                strState = ""
'
'        End Select
'    Next
'
'End Sub

'Function SetResult(asResult As String, asEquipCode As String)
'    Dim i As Integer
'    Dim sLVal As String
'    Dim sHVal As String
'    Dim sEquipCode As String
'    Dim sEquipRes As String
'    Dim sResult As String
'    Dim sPoint As Integer
'    Dim sResType As String
'    Dim sResFlag As String
'
'
'    sEquipRes = Trim(asResult)
'    sEquipCode = Trim(asEquipCode)
'    sResFlag = ""
'
'    If sEquipCode = "" Then
'        Exit Function
'    End If
'
''    If IsNumeric(sEquipRes) = False Then
''        Exit Function
''    End If
'
'    SQL = "select resprec, reflow, refhigh from equipexam where equipcode = '" & sEquipCode & "' AND EQUIPNO = '" & gEquip & "' "
'    Res = GetDBSelectColumn(gLocal, SQL)
'
'    If IsNumeric(gReadBuf(0)) = True Then
'        sPoint = CInt(gReadBuf(0))
'        sResType = ""
'        For i = 0 To sPoint
'            If i = 0 Then
'                sResType = "#0"
'            ElseIf i = 1 Then
'                sResType = sResType & ".0"
'            Else
'                sResType = sResType & "0"
'            End If
'        Next
'
'        sResult = Format(sEquipRes, sResType)
'    Else
'        sResult = sEquipRes
'    End If
'
'''    If IsNumeric(gReadBuf(1)) = True Then
'''        sLVal = gReadBuf(1)
'''        If CCur(sLVal) > CCur(sEquipRes) Then
'''            sResFlag = "H"
'''        End If
'''    End If
'''
'''    If IsNumeric(gReadBuf(2)) = True Then
'''        sHVal = gReadBuf(2)
'''        If CCur(sHVal) < CCur(sEquipRes) Then
'''            sResFlag = ">"
'''        End If
'''    End If
'
'    If IsNumeric(gReadBuf(1)) = True And IsNumeric(gReadBuf(2)) = True Then
'        sLVal = gReadBuf(1)
'        sHVal = gReadBuf(2)
'        If CCur(sEquipRes) > CCur(sLVal) And CCur(sEquipRes) < CCur(sHVal) Then
'            sResFlag = ""
'        ElseIf CCur(sHVal) <= CCur(sEquipRes) Then
'            sResFlag = "H"
'        ElseIf CCur(sLVal) >= CCur(sEquipRes) Then
'            sResFlag = "L"
'        End If
'    End If
'
'    gsFlag = sResFlag
'    SetResult = sResult
'
'End Function

Private Sub GetOrder(ByVal pBarNo As String)
'    Dim i           As Integer
'    Dim intRow      As Long
'    Dim strItems    As String
'
'    intRow = -1
'    For i = 1 To spdorder.DataRowCnt
'        If Trim(GetText(spdorder, i, colBarcode)) = pBarNo Then
'            intRow = i
'            Exit For
'        End If
'    Next i
'
'    If intRow < 0 Then
'        intRow = spdorder.DataRowCnt + 1
'        If spdorder.MaxRows < intRow Then
'            spdorder.MaxRows = intRow
'        End If
'    End If
'
'    Call SetText(spdorder, pBarNo, intRow, colBarcode)         '2
'    Call SetText(spdorder, mOrder.RackNo, intRow, colRack)     '3
'    Call SetText(spdorder, mOrder.TubePos, intRow, colPos)     '4
'    Call vasActiveCell(spdorder, intRow, colBarcode)
'    Call ClearSpread(vasRes)
'
'    Call GetSampleInfoW(intRow)                            '5,6,7,8
'
'    gOrderExam = GetOrderExamCode_New(gEquip, pBarNo)
'
'    '-- 기존 검사했던 바코드가 다시 올라올 경우 위치를 못찾는다.
'    '-- intRow 추가
'    strItems = GetGetEquipExamCode_AU480(gEquip, pBarNo, intRow)
'
'    If Trim(strItems) = "" Then
'        mOrder.NoOrder = True
'        mOrder.Order = ""
'        'S 003401 0019          1013001918    E
'        comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & ETX
'        Debug.Print STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & Space(20 - Len(mOrder.BarNo)) & mOrder.BarNo & "    E" & ETX
'    Else
'        mOrder.NoOrder = False
'        mOrder.Order = strItems
'        'S 003401 0019          1013001918    E      01020304050607091011121415161719212632
'        comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & strItems & ETX
'        'comEqp.Output = STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E012" & ETX
'
'
'        Debug.Print STX & "S " & mOrder.RackNo & mOrder.TubePos & mOrder.Seq & mOrder.BarNo & "    E" & strItems & ETX
'
'    End If
    

End Sub

Private Sub SetPatInfo(ByVal pBarNo As String)
'    Dim i           As Integer
'    Dim intRow      As Long
'    Dim strItems    As String
'
'    intRow = -1
'    For i = 1 To spdorder.DataRowCnt
'        If Trim(GetText(spdorder, i, colBarcode)) = pBarNo Then
'            intRow = i
'            Exit For
'        End If
'    Next i
'
'    If intRow < 0 Then
'        intRow = spdorder.DataRowCnt + 1
'        If spdorder.MaxRows < intRow Then
'            spdorder.MaxRows = intRow
'        End If
'    End If
'
'    Call SetText(spdorder, pBarNo, intRow, colBarcode)             '2 Barcode
'    Call SetText(spdorder, mResult.RackNo, intRow, colRack)        '3 Rack
'    Call SetText(spdorder, mResult.TubePos, intRow, colPos)        '4 Pos
'    Call vasActiveCell(spdorder, intRow, colBarcode)
'
'    Call ClearSpread(vasRes)
'
'    Call GetSampleInfoW(intRow)                                '5,6,7,8
'
'    gRow = intRow
'
'    gOrderExam = GetOrderExamCode_New(gEquip, pBarNo)

End Sub
