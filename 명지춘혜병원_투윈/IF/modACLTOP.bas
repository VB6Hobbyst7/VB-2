Attribute VB_Name = "modACLTOP"
Option Explicit

''-----------------------------------------------------------------------------'
''   ��� : �������� ����
''-----------------------------------------------------------------------------'
'Private Sub SendOrder()
'
'
'    Dim strOutput   As String     '�۽��� ������
'    Dim blnLast     As Boolean
'    Dim intRow      As Integer
'    Dim strBarno    As String
'    Dim strItems    As String
'
'    blnLast = False
'
'    With frmMain.spdOrder
'        If intSndPhase <= 3 Then
'            For intRow = 1 To .DataRowCnt
'                If GetText(frmMain.spdOrder, intRow, colCHECKBOX) = "1" And GetText(frmMain.spdOrder, intRow, colSTATE) = "�����غ�" Then
'                    strBarno = Trim(GetText(frmMain.spdOrder, intRow, colBARCODE))
'                    strItems = Trim(GetText(frmMain.spdOrder, intRow, colKEY1))
'                    If intSndPhase = 3 Then
'                        .Row = intRow
'                        .Col = colCHECKBOX: .Text = "0"
'                        .Col = colSTATE:    .Text = "��������"
'
'                        If intRow = .DataRowCnt Then
'                            blnLast = True
'                        End If
'
'                    End If
'                    Exit For
'                End If
'            Next
'        End If
'    End With
'
'    If intRow = frmMain.spdOrder.DataRowCnt Then
'        blnLast = True
'    End If
'
'    Select Case intSndPhase
'        Case 1  '## Header
'        '''''            strOutput = "H|@^\|" & mOrder.MsgID & "||" & mOrder.Receiver & "|||||" & mOrder.Sender & "||P|" & mOrder.Version & "|" & Format(Now, "yyyyMMddHHmmss") & vbCr
'            strOutput = intFrameNo & "H|@^\|" & mOrder.MsgID & "||" & mOrder.Receiver & "|||||" & mOrder.Sender & "||P|" & mOrder.Version & "|" & Format(Now, "yyyyMMddHHmmss") & vbCr & ETB
'            intSndPhase = 2
'            intFrameNo = intFrameNo + 1
'
'        Case 2  '## Patient
''''''        strOutput = strOutput & "P|" & mPNo & "||||^||||||||" & vbCr
'            strOutput = intFrameNo & "P|" & mPNo & "||||^||||||||" & vbCr & ETB
'            intSndPhase = 3
'            intFrameNo = intFrameNo + 1
'            mPNo = mPNo + 1
'
'        Case 3  '## Order
'            '## ���� ������
'            If mOrder.IsSending = False Then
''''''         = strOutput & "O|1|" & strBarno & "||" & strItems & "|R|" & Format(Now, "yyyyMMddHHmmss") & "|||||A||||P||||||||||Q" & vbCr
'                strOutput = "O|1|" & strBarno & "||" & strItems & "|R|" & Format(Now, "yyyyMMddHHmmss") & "|||||A||||P||||||||||Q"
'                If Len(strOutput) > 230 Then
'                    mOrder.IsSending = True
'                    mOrder.Order = Mid$(strOutput, 231)
'                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'                    intSndPhase = 3
'                Else
'                    strOutput = intFrameNo & strOutput & vbCr & ETB
'                    If blnLast = True Then
'                        intSndPhase = 4
'                    Else
'                        intSndPhase = 2
'                    End If
'                End If
'            '## ���� ���ڿ��� ������
'            Else
'                strOutput = mOrder.Order
'                If Len(strOutput) > 230 Then
'                    mOrder.Order = Mid$(strOutput, 231)
'                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'                    intSndPhase = 3
'                Else
'                    mOrder.IsSending = False
'                    strOutput = intFrameNo & strOutput & vbCr & ETB
'                    If blnLast = True Then
'                        intSndPhase = 4
'                    Else
'                        intSndPhase = 2
'                    End If
'                End If
'            End If
'            intFrameNo = intFrameNo + 1
'
'        Case 4  '## Termianator
''''''            strOutput = strOutput & "L|1|N"
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

'Public Sub Phase_Serial_ACLTOP()
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
'                        intBufCnt = 0
'                        Erase strRecvData
'                        intPhase = 2
'                        frmMain.comEqp.Output = ACK
'                        SetRawData "[Tx]" & ACK
'                    Case ACK
'                        If strState = "Q" Then
'                            Call SendOrder
'                        Else
'                            frmMain.comEqp.Output = ACK
'                            SetRawData "[Tx]" & ACK
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
'                        Else
'                            intBufCnt = intBufCnt + 1
'                        End If
'                        ReDim Preserve strRecvData(intBufCnt)
'                    Case ETB
'                        blnIsETB = True
'                        intPhase = 3
'                    Case ETX
'                        intPhase = 3
'                    Case EOT
'                        intPhase = 1
'                    Case vbCr
'                        intBufCnt = intBufCnt + 1
'                        ReDim Preserve strRecvData(intBufCnt)
'                    Case vbLf
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
'                    Case vbLf
'                        intPhase = IIf(blnIsETB = False, 4, 2)
'                        frmMain.comEqp.Output = ACK
'                        SetRawData "[Tx]" & ACK
'                End Select
'            Case 4      '## Termination Phase
'                Select Case BufChar
'                    Case STX
'                        intPhase = 2
'                    Case EOT
'                        Call SerialRcvData_ACLTOP
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
'
'
'Private Sub SerialRcvData_ACLTOP()
'    Dim RS_L            As ADODB.Recordset
'    Dim strRcvBuf       As String   '������ Data
'    Dim strType         As String   '������ Record Type
'    Dim strBarno        As String   '������ ���ڵ��ȣ
'    Dim strSeq          As String   '������ Sequence
'    Dim strRackNo       As String   '������ Rack Or Disk No
'    Dim strTubePos      As String   '������ Tube Position
'    Dim strIntBase      As String   '������ ������ �˻��
'    Dim strMachResult   As String   '������ �����
'    Dim strResult       As String   '������ ���(����)
'    Dim strIntResult    As String   '������ ���(����)
'    Dim strQCResult     As String   '������ ���(QC)
'    Dim strFlag         As String   '������ Abnormal Flag
'    Dim strComm         As String   '������ Comment
'    Dim strAspect       As String
'
'    Dim strTemp1        As String
'    Dim strTemp2        As String
'
'    Dim lsOrderCode     As String   'ó���ڵ�
'    Dim lsTestCode      As String   '�˻��ڵ�
'    Dim lsTestName      As String   '�˻��
'    Dim lsSeqNo         As String   '����DB �˻�Seq
'
'    Dim lsRstRow        As String   '����������� ���� Row
'    Dim intCnt          As Integer  '��� Frame ����
'    Dim intCol          As Integer  '����÷� ����
'    Dim strJudge        As String   '�������
'    Dim Res             As Integer
'
'    Dim strTmp          As String
'    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
'    Dim varBarno        As Variant
'    Dim i               As Integer
'
'    Dim strUseRes       As String
'
'    With frmMain
'        For intCnt = 1 To UBound(strRecvData)
'            strRcvBuf = strRecvData(intCnt)
'
'            '-- �׽�Ʈ�� -----------------
'            If .fraCommTest.Visible = False Then
'                Call SetSQLData("RCV", strRcvBuf, "A")
'            End If
'            '-- �׽�Ʈ�� -----------------
'
'            strType = Mid$(strRcvBuf, 2, 1)
'            If strType = "|" Then
'                strType = Mid$(strRcvBuf, 1, 1)
'            End If
'
'            Select Case strType
'                Case "H"    '## Header
'                    '1H|@^\|<1504128210_21570><1504128210_21571>||acl|||||LIS||P|1394-97|20170830172330
'                    mOrder.MsgID = Trim(mGetP(strRcvBuf, 3, "|"))
'                    mOrder.Sender = Trim(mGetP(strRcvBuf, 5, "|"))
'                    mOrder.Receiver = Trim(mGetP(strRcvBuf, 10, "|"))
'                    mOrder.Version = Trim(mGetP(strRcvBuf, 13, "|"))
'
'                Case "P"    '## Patient
'                Case "Q"    '## Request Information
'
'                    'Q|1|^1001@^1002@^1003@^1004@^1005@^1006@^1008||||||||||O@N
'                    'Q|1|^198772||||||||||O@N
'                    'Q|1|^1310250941@^1310250867||||||||||O@N
'
'
'                    strTemp1 = mGetP(strRcvBuf, 3, "|")
'                    strTemp1 = Replace(strTemp1, "^", "")
'
'                    varBarno = Split(strTemp1, "@")
'
'                    For i = 0 To UBound(varBarno)
'                        mOrder.BarNo = varBarno(i)
'                        Call GetOrder(varBarno(i), gHOSP.RSTTYPE)
'                    Next
'
''                    With mOrder
''                        .NoOrder = False
''                        .BarNo = strBarno
''                        .Seq = mGetP(strTemp1, 3, "^")
''                        .RackNo = mGetP(strTemp1, 4, "^")
''                        .TubePos = mGetP(strTemp1, 5, "^")
''                    End With
'
'                   ' Call GetOrder(strBarno, gHOSP.RSTTYPE)
'                    strState = "Q"
'                    mPNo = 1
'
'                Case "O"
'                    strBarno = mGetP(strRcvBuf, 3, "|")
'
'                    With mResult
'                        .BarNo = strBarno
'                        '.SpcPos = strTubePos & "/" & strRackNo
'                        '.Seq = strSeq
'                        .RackNo = strRackNo
'                        .TubePos = strTubePos
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
'                Case "R"
'                    'R|1|^^^131|28.4|s||N||F@V||SysAdmin^SysAdmin||20170826150315|
'                    'R|1|^^^541|103.6|D mAbs||N||F@V||SysAdmin^SysAdmin||2017090108
'                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
'                    If strIntBase = "131" Then
'                        strIntBase = strIntBase & UCase(mGetP(strRcvBuf, 5, "|"))
'                    End If
'
'                    'R|1|^^^2241|0.3|microg/mLFEU||N||F@V||SysAdmin^SysAdmin||2017
'                    'R|2|^^^2241|172|ng/mL||N||F@V||SysAdmin^SysAdmin||20170901083115|acl^03^2
'
'                    ' D-Dimer
'                    If strIntBase = "2241" Then
'                        If mGetP(strRcvBuf, 5, "|") = "microg/mLFEU" Then
'                            strIntBase = strIntBase
'                        Else
'                            strIntBase = ""
'                        End If
'                    End If
'
'                    strResult = mGetP(strRcvBuf, 4, "|")
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
'                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""
'                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
'
'                                '-- ���Row �߰�
'                                lsRstRow = .spdResult.DataRowCnt + 1
'                                If .spdResult.MaxRows < lsRstRow Then
'                                    .spdResult.MaxRows = lsRstRow
'                                End If
'
'                                '-- �Ҽ��� ó��, ��� ���� ó��
'                                If strUseRes <> "" Then
'                                    strMachResult = strResult
'                                    strResult = SetResult(strResult, strIntBase)
'                                End If
'                                strJudge = SetJudge(strResult, strIntBase)
'
'                                '������� ǥ��("���")
'                                SetText .spdOrder, "���", gRow, colSTATE
'
'                                '����� ǥ��
'                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                        SetText .spdOrder, strResult, gRow, intCol
'                                        Exit For
'                                    End If
'                                Next
'
'                                '-- ��� List
'                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '����
'                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
'                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
'                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '�˻��
'                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '���ä��
'                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '�����
'                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS���
'                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '����
'                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
'
'                                '-- ���� ����
'                                SetLocalDB gRow, lsRstRow, "1", ""
'
'                                '-- BIORAD QC ����
'                                If Mid(strBarno, 1, 2) = "QC" Then
'                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
'                                End If
'
'
'                                strState = "R"
'
'                                '-- ���Count
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
'                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""
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
'                                '-- ���Row �߰�
'                                lsRstRow = .spdResult.DataRowCnt + 1
'                                If .spdResult.MaxRows < lsRstRow Then
'                                    .spdResult.MaxRows = lsRstRow
'                                End If
'
'                                '-- �Ҽ��� ó��, ��� ���� ó��
'                                If strUseRes <> "" Then
'                                    strMachResult = strResult
'                                    strResult = SetResult(strResult, strIntBase)
'                                End If
'                                strJudge = SetJudge(strResult, strIntBase)
'
'
'                                '������� ǥ��("���")
'                                SetText .spdOrder, "���", gRow, colSTATE
'
'                                '����� ǥ��
'                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                        SetText .spdOrder, strResult, gRow, intCol
'                                        Exit For
'                                    End If
'                                Next
'
'                                '-- ��� List
'                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '����
'                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
'                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
'                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '�˻��
'                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '���ä��
'                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '�����
'                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS���
'                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '����
'                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
'
'                                '-- ���� ����
'                                SetLocalDB gRow, lsRstRow, "1", ""
'
'                                '-- BIORAD QC ����
'                                If Mid(strBarno, 1, 2) = "QC" Then
'                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
'                                End If
'
'                                If strState <> "R" Then
'                                    strState = ""
'                                End If
'
'                                '-- ���Count
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
'                Case "C"    '## Comment
'                    '## Abnormal ����϶� Comment ����
'                    If strFlag <> "N" Then
'                        strTemp1 = mGetP(strRcvBuf, 4, "|")
'                        strComm = mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")
'                    End If
'
'                Case "L"
'                    '## DB�� �������
'                    If .optTrans(0).Value = True And strState = "R" Then
'                        Res = SaveTransData_MCC(gRow)
'
'                        If Res = -1 Then
'                            '-- ���� ����
'                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                            SetText .spdOrder, "Failed", gRow, colSTATE
'                        Else
'                            '-- ���� ����
'                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
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
'                                '-- ����
'                            End If
'                        End If
'                        strState = ""
'                    End If
'            End Select
'        Next
'    End With
'
'End Sub



''-----------------------------------------------------------------------------'
''   ��� : �ش� ���ڵ��ȣ�� ���� 1. �������� ��ȸ,
''                                 2. ���������� ȭ��ǥ��,
''                                 3. ó���ڵ� ��������,
''                                 4. (ó���ڵ��)�˻���� �����
''   �μ� :
''       - pBarNo : ���ڵ��ȣ
''       - pType  : ���ڵ� �̻��� ���ϴ� ���
''                   1 : Seq
''                   2 : Rack/Pos
''                   3 : üũ�Ȱ��� ���� ���� ��
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
'    '-- 1. �������� ��ȸ
'    With frmMain
'        '-- ���ڵ� ���
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
'        '-- �������忡�� ��ã����..
'        If intRow < 0 Then
'            intRow = .spdOrder.DataRowCnt + 1
'            If .spdOrder.MaxRows < intRow Then
'                .spdOrder.MaxRows = intRow
'            End If
'        End If
'
'        '-- ���������� ȭ��ǥ��
'        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
'        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
'        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
'        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
'
'        '-- ����������� �����
'        .spdResult.MaxRows = 0
'
'        '-- �˻��� ���� ��������
'        Call GetSampleInfo(intRow, .spdOrder)
'
'        .spdOrder.RowHeight(-1) = 12
'
'        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
'        strItems = GetEquipExamCode_ACLTOP(gHOSP.MACHCD, pBarno, intRow)
'
'        '-- �˻�ä�η� ������ �����
'        If Trim(strItems) = "" Then
'            mOrder.NoOrder = True
'            mOrder.Order = ""
'
'            '-- �������(Order) ǥ��
'            Call SetText(frmMain.spdOrder, "�����غ�", intRow, colSTATE)
'        Else
'            mOrder.NoOrder = False
'            mOrder.Order = strItems
'
'            '-- �������(Order) ǥ��
'            Call SetText(frmMain.spdOrder, "�����غ�", intRow, colSTATE)
'            '-- �������(Order) ǥ��
'            Call SetText(frmMain.spdOrder, strItems, intRow, colKEY1)
'        End If
'
'        SetText frmMain.spdOrder, "1", intRow, colCHECKBOX
'
'        '-- ���� Row
'        gRow = intRow
'
'    End With
'
'End Sub
'
''��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
''�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
'Private Function GetEquipExamCode_ACLTOP(argEquipCode As String, argPID As String, Optional intRow As Long) As String
'    Dim i As Integer
'    Dim sExamCode As String
'    Dim strExamCode As String
'    Dim sSpecNo     As String
'    Dim iRow        As Long
'    Dim SpecNo      As String
'
'    GetEquipExamCode_ACLTOP = ""
'
'    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
'        Exit Function
'    End If
'
'    '-- ������ �˻��ڵ��� ä�� ã��
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
'            strExamCode = strExamCode & "@^^^" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
'            AdoRs_Local.MoveNext
'        Loop
'    End If
'
'    AdoRs_Local.Close
'
'    GetEquipExamCode_ACLTOP = Mid(strExamCode, 2)
'
'End Function




'
'
'Option Explicit
'
''-----------------------------------------------------------------------------'
''   ��� : �������� ����
''-----------------------------------------------------------------------------'
'Private Sub SendOrder()
'
'
'    Dim strOutput   As String     '�۽��� ������
'    Dim blnLast     As Boolean
'    Dim intRow      As Integer
'    Dim strBarno    As String
'    Dim strItems    As String
'
'    blnLast = False
'
'    With frmMain.spdOrder
'        If intSndPhase <= 3 Then
'            For intRow = 1 To .DataRowCnt
'                If GetText(frmMain.spdOrder, intRow, colCHECKBOX) = "1" And GetText(frmMain.spdOrder, intRow, colSTATE) = "�����غ�" Then
'                    strBarno = Trim(GetText(frmMain.spdOrder, intRow, colBARCODE))
'                    strItems = Trim(GetText(frmMain.spdOrder, intRow, colKEY1))
'                    If intSndPhase = 3 Then
'                        .Row = intRow
'                        .Col = colCHECKBOX: .Text = "0"
'                        .Col = colSTATE:    .Text = "��������"
'
'                        If intRow = .DataRowCnt Then
'                            blnLast = True
'                        End If
'
'                    End If
'                    Exit For
'                End If
'            Next
'        End If
'    End With
'
'    If intRow = frmMain.spdOrder.DataRowCnt Then
'        blnLast = True
'    End If
'
'    Select Case intSndPhase
'        Case 1  '## Header
'        '''''            strOutput = "H|@^\|" & mOrder.MsgID & "||" & mOrder.Receiver & "|||||" & mOrder.Sender & "||P|" & mOrder.Version & "|" & Format(Now, "yyyyMMddHHmmss") & vbCr
'            strOutput = intFrameNo & "H|@^\|" & mOrder.MsgID & "||" & mOrder.Receiver & "|||||" & mOrder.Sender & "||P|" & mOrder.Version & "|" & Format(Now, "yyyyMMddHHmmss") & vbCr & ETB
'            intSndPhase = 2
'            intFrameNo = intFrameNo + 1
'
'        Case 2  '## Patient
''''''        strOutput = strOutput & "P|" & mPNo & "||||^||||||||" & vbCr
'            strOutput = intFrameNo & "P|" & mPNo & "||||^||||||||" & vbCr & ETB
'            intSndPhase = 3
'            intFrameNo = intFrameNo + 1
'            mPNo = mPNo + 1
'
'        Case 3  '## Order
'            '## ���� ������
'            If mOrder.IsSending = False Then
''''''         = strOutput & "O|1|" & strBarno & "||" & strItems & "|R|" & Format(Now, "yyyyMMddHHmmss") & "|||||A||||P||||||||||Q" & vbCr
'                strOutput = "O|1|" & strBarno & "||" & strItems & "|R|" & Format(Now, "yyyyMMddHHmmss") & "|||||A||||P||||||||||Q"
'                If Len(strOutput) > 230 Then
'                    mOrder.IsSending = True
'                    mOrder.Order = Mid$(strOutput, 231)
'                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'                    intSndPhase = 3
'                Else
'                    strOutput = intFrameNo & strOutput & vbCr & ETB
'                    If blnLast = True Then
'                        intSndPhase = 4
'                    Else
'                        intSndPhase = 2
'                    End If
'                End If
'            '## ���� ���ڿ��� ������
'            Else
'                strOutput = mOrder.Order
'                If Len(strOutput) > 230 Then
'                    mOrder.Order = Mid$(strOutput, 231)
'                    strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
'                    intSndPhase = 3
'                Else
'                    mOrder.IsSending = False
'                    strOutput = intFrameNo & strOutput & vbCr & ETB
'                    If blnLast = True Then
'                        intSndPhase = 4
'                    Else
'                        intSndPhase = 2
'                    End If
'                End If
'            End If
'            intFrameNo = intFrameNo + 1
'
'        Case 4  '## Termianator
''''''            strOutput = strOutput & "L|1|N"
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
'
'
'Public Sub Phase_Serial_ACLTOP()
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
'                        strFRcvState = ""
'                        strFSndState = ""
'                        strFRcvBuffer = ""
'
'                        frmMain.comEqp.Output = ACK
'                        SetRawData "[Tx]" & ACK
'                        intPhase = 2
'                    Case Else
'                        strFRcvState = ""
'                        strFSndState = ""
'                        intPhase = 1
'                End Select
'
'            Case 2      '## Transfer Phase
'                Select Case BufChar
'                    Case STX
'                    Case EOT
'                        Call SerialRcvData_ACLTOP
'                    Case ENQ
'                        strFRcvState = ""
'                        strFSndState = ""
'                        strFRcvBuffer = ""
'
'                        frmMain.comEqp.Output = ACK
'                        SetRawData "[Tx]" & ACK
'
'                    Case vbLf
'                        strFRcvBuffer = strFRcvBuffer & Mid(SavBuffer, 2, Len(SavBuffer) - 5)
'                        SavBuffer = ""
'
'                        frmMain.comEqp.Output = ACK
'                        SetRawData "[Tx]" & ACK
'
'                    Case NAK
'                        frmMain.comEqp.Output = ENQ
'                        SetRawData "[Tx]" & ENQ
'
'                    Case Else
'                        SavBuffer = SavBuffer & wkdat
'                End Select
'
'            Case 3      '## Transfer Phase
'                Select Case BufChar
'                    Case EOT
'                        intPhase = 1
'                    Case ACK
'                        If strFSndState = "E" Then
'                            ii_SendCnt = 0
'
'                            frmMain.comEqp.Output = m_aTemp(ii_SendCnt)
'
'                            If ii_SendCnt + 1 = miSendCnt Then
'                                strFSndState = "L"
'                            Else
'                                strFSndState = "P"
'                            End If
'
'                            intPhase = 3
'                            Exit Sub
'
'                        ElseIf strFSndState = "P" Then
'                            ii_SendCnt = ii_SendCnt + 1
'                            MSComm.Output = m_aTemp(ii_SendCnt)
'
'                            If ii_SendCnt + 1 = miSendCnt Then
'                                strFSndState = "L"
'                            Else
'                                strFSndState = "P"
'                            End If
'
'                            intPhase = 3
'                            Exit Sub
'
'                        ElseIf strFSndState = "L" Then
'                            'EOT ����
'                            MSComm.Output = Chr(4)
'
'                            m_iFrameN = 0
'                            strFSndState = ""
'                            ii_SendCnt = 0
'                            miSendCnt = 0
'                            intPhase = 1
'                            msAllBarCd = ""
'                            Erase maAllBarCd
'                            Erase m_aTemp
'
'                        End If
'                    Case NAK
'                        If strFSndState = "E" Then
'                            MSComm.Output = Chr(5)
'
'                            strFSndState = "E"
'                            intPhase = 3
'                            Exit Sub
'                        ElseIf strFSndState = "P" Or strFSndState = "L" Then
'                            MSComm.Output = m_aTemp(ii_SendCnt)
'
'                            If ii_SendCnt + 1 = miSendCnt Then
'                                strFSndState = "L"
'                            Else
'                                strFSndState = "P"
'                            End If
'
'                            intPhase = 3
'                            Exit Sub
'                        End If
'                    'ENQ
'                    Case 5
'                        strFRcvState = ""
'                        strFSndState = ""
'                        strFRcvBuffer = ""
'
'                        'ACK ����
'                        MSComm.Output = Chr(6)
'
'                        strFRcvBuffer = ""
'                        intPhase = 2
'                End Select
'        End Select
'    Next i
'
'End Sub
'
'
'Private Sub SerialRcvData_ACLTOP()
'    Dim RS_L            As ADODB.Recordset
'    Dim strRcvBuf       As String   '������ Data
'    Dim strType         As String   '������ Record Type
'    Dim strBarno        As String   '������ ���ڵ��ȣ
'    Dim strSeq          As String   '������ Sequence
'    Dim strRackNo       As String   '������ Rack Or Disk No
'    Dim strTubePos      As String   '������ Tube Position
'    Dim strIntBase      As String   '������ ������ �˻��
'    Dim strMachResult   As String   '������ �����
'    Dim strResult       As String   '������ ���(����)
'    Dim strIntResult    As String   '������ ���(����)
'    Dim strQCResult     As String   '������ ���(QC)
'    Dim strFlag         As String   '������ Abnormal Flag
'    Dim strComm         As String   '������ Comment
'    Dim strAspect       As String
'
'    Dim strTemp1        As String
'    Dim strTemp2        As String
'
'    Dim lsOrderCode     As String   'ó���ڵ�
'    Dim lsTestCode      As String   '�˻��ڵ�
'    Dim lsTestName      As String   '�˻��
'    Dim lsSeqNo         As String   '����DB �˻�Seq
'
'    Dim lsRstRow        As String   '����������� ���� Row
'    Dim intCnt          As Integer  '��� Frame ����
'    Dim intCol          As Integer  '����÷� ����
'    Dim strJudge        As String   '�������
'    Dim Res             As Integer
'
'    Dim strTmp          As String
'    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
'    Dim varBarno        As Variant
'    Dim i               As Integer
'
'    Dim strUseRes       As String
'
'    With frmMain
'        strRecvData = Split(RcvBuffer, Chr(13))
'        For intCnt = 0 To UBound(strRecvData)
'            strRcvBuf = strRecvData(intCnt)
'
'            '-- �׽�Ʈ�� -----------------
'            If .fraCommTest.Visible = False Then
'                Call SetSQLData("RCV", strRcvBuf, "A")
'            End If
'            '-- �׽�Ʈ�� -----------------
'
'            strType = Mid$(strRcvBuf, 2, 1)
'            If strType = "|" Then
'                strType = Mid$(strRcvBuf, 1, 1)
'            End If
'
'            Select Case strType
'                Case "H"    '## Header
'                    '1H|@^\|<1504128210_21570><1504128210_21571>||acl|||||LIS||P|1394-97|20170830172330
'                    mOrder.MsgID = Trim(mGetP(strRcvBuf, 3, "|"))
'                    mOrder.Sender = Trim(mGetP(strRcvBuf, 5, "|"))
'                    mOrder.Receiver = Trim(mGetP(strRcvBuf, 10, "|"))
'                    mOrder.Version = Trim(mGetP(strRcvBuf, 13, "|"))
'                Case "P"    '## Patient
'                Case "Q"    '## Request Information
'                    'Q|1|^1001@^1002@^1003@^1004@^1005@^1006@^1008||||||||||O@N
'                    'Q|1|^198772||||||||||O@N
'                    'Q|1|^1310250941@^1310250867||||||||||O@N
'                    strTemp1 = mGetP(strRcvBuf, 3, "|")
'                    strTemp1 = Replace(strTemp1, "^", "")
'
'                    strFRcvState = "Q"
'
'                    varBarno = Split(strTemp1, "@")
'
'                    For i = 0 To UBound(varBarno)
'                        If varBarno(i) <> "ALL" Then
'                            '
'                        Else
'                            m_p_iOrdCnt = 0
'                        End If
'                        mOrder.BarNo = varBarno(i)
'                        Call GetOrder(varBarno(i), gHOSP.RSTTYPE)
'                    Next
'                    strState = "Q"
'                    mPNo = 1
'
'                Case "O"
'                    strBarno = mGetP(strRcvBuf, 3, "|")
'
'                    With mResult
'                        .BarNo = strBarno
'                        '.SpcPos = strTubePos & "/" & strRackNo
'                        '.Seq = strSeq
'                        .RackNo = strRackNo
'                        .TubePos = strTubePos
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
'                Case "R"
'                    'R|1|^^^131|28.4|s||N||F@V||SysAdmin^SysAdmin||20170826150315|
'                    'R|1|^^^541|103.6|D mAbs||N||F@V||SysAdmin^SysAdmin||2017090108
'                    strIntBase = mGetP(mGetP(strRcvBuf, 3, "|"), 4, "^")
'                    If strIntBase = "131" Then
'                        strIntBase = strIntBase & UCase(mGetP(strRcvBuf, 5, "|"))
'                    End If
'
'                    'R|1|^^^2241|0.3|microg/mLFEU||N||F@V||SysAdmin^SysAdmin||2017
'                    'R|2|^^^2241|172|ng/mL||N||F@V||SysAdmin^SysAdmin||20170901083115|acl^03^2
'
'                    ' D-Dimer
'                    If strIntBase = "2241" Then
'                        If mGetP(strRcvBuf, 5, "|") = "microg/mLFEU" Then
'                            strIntBase = strIntBase
'                        Else
'                            strIntBase = ""
'                        End If
'                    End If
'
'                    strResult = mGetP(strRcvBuf, 4, "|")
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
'                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""
'                                strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
'
'                                '-- ���Row �߰�
'                                lsRstRow = .spdResult.DataRowCnt + 1
'                                If .spdResult.MaxRows < lsRstRow Then
'                                    .spdResult.MaxRows = lsRstRow
'                                End If
'
'                                '-- �Ҽ��� ó��, ��� ���� ó��
'                                If strUseRes <> "" Then
'                                    strMachResult = strResult
'                                    strResult = SetResult(strResult, strIntBase)
'                                End If
'                                strJudge = SetJudge(strResult, strIntBase)
'
'                                '������� ǥ��("���")
'                                SetText .spdOrder, "���", gRow, colSTATE
'
'                                '����� ǥ��
'                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                        SetText .spdOrder, strResult, gRow, intCol
'                                        Exit For
'                                    End If
'                                Next
'
'                                '-- ��� List
'                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '����
'                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
'                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
'                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '�˻��
'                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '���ä��
'                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '�����
'                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS���
'                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '����
'                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
'
'                                '-- ���� ����
'                                SetLocalDB gRow, lsRstRow, "1", ""
'
'                                '-- BIORAD QC ����
'                                If Mid(strBarno, 1, 2) = "QC" Then
'                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
'                                End If
'
'
'                                strState = "R"
'
'                                '-- ���Count
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
'                                strUseRes = Trim(RS_L.Fields("QCTEMP")) & ""
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
'                                '-- ���Row �߰�
'                                lsRstRow = .spdResult.DataRowCnt + 1
'                                If .spdResult.MaxRows < lsRstRow Then
'                                    .spdResult.MaxRows = lsRstRow
'                                End If
'
'                                '-- �Ҽ��� ó��, ��� ���� ó��
'                                If strUseRes <> "" Then
'                                    strMachResult = strResult
'                                    strResult = SetResult(strResult, strIntBase)
'                                End If
'                                strJudge = SetJudge(strResult, strIntBase)
'
'
'                                '������� ǥ��("���")
'                                SetText .spdOrder, "���", gRow, colSTATE
'
'                                '����� ǥ��
'                                For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                    If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                        SetText .spdOrder, strResult, gRow, intCol
'                                        Exit For
'                                    End If
'                                Next
'
'                                '-- ��� List
'                                SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '����
'                                SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
'                                SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
'                                SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '�˻��
'                                SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '���ä��
'                                SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '�����
'                                SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS���
'                                SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '����
'                                SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
'
'                                '-- ���� ����
'                                SetLocalDB gRow, lsRstRow, "1", ""
'
'                                '-- BIORAD QC ����
'                                If Mid(strBarno, 1, 2) = "QC" Then
'                                    Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
'                                End If
'
'                                If strState <> "R" Then
'                                    strState = ""
'                                End If
'
'                                '-- ���Count
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
'                Case "C"    '## Comment
'                    '## Abnormal ����϶� Comment ����
'                    If strFlag <> "N" Then
'                        strTemp1 = mGetP(strRcvBuf, 4, "|")
'                        strComm = mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")
'                    End If
'
'                Case "L"
'                    '## DB�� �������
'                    If .optTrans(0).Value = True And strState = "R" Then
'                        Res = SaveTransData_MCC(gRow)
'
'                        If Res = -1 Then
'                            '-- ���� ����
'                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                            SetText .spdOrder, "Failed", gRow, colSTATE
'                        Else
'                            '-- ���� ����
'                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                            SetText .spdOrder, "����Ϸ�", gRow, colSTATE
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
'                                '-- ����
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
'
'
''-----------------------------------------------------------------------------'
''   ��� : �ش� ���ڵ��ȣ�� ���� 1. �������� ��ȸ,
''                                 2. ���������� ȭ��ǥ��,
''                                 3. ó���ڵ� ��������,
''                                 4. (ó���ڵ��)�˻���� �����
''   �μ� :
''       - pBarNo : ���ڵ��ȣ
''       - pType  : ���ڵ� �̻��� ���ϴ� ���
''                   1 : Seq
''                   2 : Rack/Pos
''                   3 : üũ�Ȱ��� ���� ���� ��
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
'    '-- 1. �������� ��ȸ
'    With frmMain
'        '-- ���ڵ� ���
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
'        '-- �������忡�� ��ã����..
'        If intRow < 0 Then
'            intRow = .spdOrder.DataRowCnt + 1
'            If .spdOrder.MaxRows < intRow Then
'                .spdOrder.MaxRows = intRow
'            End If
'        End If
'
'        '-- ���������� ȭ��ǥ��
'        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
'        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
'        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
'        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
'
'        '-- ����������� �����
'        .spdResult.MaxRows = 0
'
'        '-- �˻��� ���� ��������
'        Call GetSampleInfo(intRow, .spdOrder)
'
'        .spdOrder.RowHeight(-1) = 12
'
'        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
'        strItems = GetEquipExamCode_ACLTOP(gHOSP.MACHCD, pBarno, intRow)
'
'        '-- �˻�ä�η� ������ �����
'        If Trim(strItems) = "" Then
'            mOrder.NoOrder = True
'            mOrder.Order = ""
'
'            '-- �������(Order) ǥ��
'            Call SetText(frmMain.spdOrder, "�����غ�", intRow, colSTATE)
'        Else
'            mOrder.NoOrder = False
'            mOrder.Order = strItems
'
'            '-- �������(Order) ǥ��
'            Call SetText(frmMain.spdOrder, "�����غ�", intRow, colSTATE)
'            '-- �������(Order) ǥ��
'            Call SetText(frmMain.spdOrder, strItems, intRow, colKEY1)
'        End If
'
'
'        '-- ���� Row
'        gRow = intRow
'
'    End With
'
'End Sub
'
''��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
''�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
'Private Function GetEquipExamCode_ACLTOP(argEquipCode As String, argPID As String, Optional intRow As Long) As String
'    Dim i As Integer
'    Dim sExamCode As String
'    Dim strExamCode As String
'    Dim sSpecNo     As String
'    Dim iRow        As Long
'    Dim SpecNo      As String
'
'    GetEquipExamCode_ACLTOP = ""
'
'    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
'        Exit Function
'    End If
'
'    '-- ������ �˻��ڵ��� ä�� ã��
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
'            strExamCode = strExamCode & "@^^^" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
'            AdoRs_Local.MoveNext
'        Loop
'    End If
'
'    AdoRs_Local.Close
'
'    GetEquipExamCode_ACLTOP = Mid(strExamCode, 2)
'
'End Function
'
'
'
