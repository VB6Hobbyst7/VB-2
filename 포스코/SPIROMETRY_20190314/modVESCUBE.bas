Attribute VB_Name = "modVESCUBE"
Option Explicit


'
'Private Sub VesMatic(asData As String)
'    Dim strHeader As String
'
'    strHeader = Trim(mGetP(asData, 1, "="))
'
'    If strHeader <> "" And IsNumeric(strHeader) Then
'        Call SerialRcvData_VESCUBE
'    Else
'        Exit Sub
'    End If
'
'End Sub
'
'Public Sub Phase_Serial_VESCUBE()
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
'        Select Case BufChar
'            Case vbCr
'                    Call VesMatic(RcvBuffer)
'                    RcvBuffer = ""
'            Case vbLf
'                    Call VesMatic(RcvBuffer)
'                    RcvBuffer = ""
'            Case Else
'                    RcvBuffer = RcvBuffer & BufChar
'        End Select
'    Next i
'
'End Sub
'
'Private Sub SerialRcvData_VESCUBE()
'    Dim RS_L            As ADODB.Recordset
'    Dim strRcvBuf       As String   '������ Data
'    Dim strType         As String   '������ Record Type
'    Dim strOldBarno        As String   '������ ���ڵ��ȣ
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
'
'    With frmMain
'        'For intCnt = 1 To UBound(strRecvData)
'        '-- �׽�Ʈ�� -----------------
'        If .fraCommTest.Visible = False Then
'            Call SetSQLData("RCV", RcvBuffer, "A")
'        End If
'        '-- �׽�Ʈ�� -----------------
'
'        '1 = O4ZU70QN0....  48
'        '1 = 199297.......   6
'        strTemp1 = mGetP(RcvBuffer, 2, "=")
'        strBarno = Trim(mGetP(strTemp1, 1, "......."))
'
'        If Trim(strBarno) <> "" And Len(strBarno) = 6 Then
'            strIntBase = "ESR"
'            strResult = Trim(Right(RcvBuffer, 5))
'
'            With mResult
'                .BarNo = strBarno
'                '.SpcPos = strTubePos & "/" & strRackNo
'                '.Seq = strSeq
'                '.RackNo = strRackNo
'                '.TubePos = strTubePos
'                If strOldBarno <> strBarno Then
'                    strOldBarno = strBarno
'                    .RsltDate = Format(Now, "yyyymmddhhmmss")
'                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'
'                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'                End If
'            End With
'
'            If gPatOrdCd <> "" Then
'                SQL = ""
'                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
'                SQL = SQL & "  FROM EQPMASTER" & vbCr
'                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
'                SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
'
'                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                    lsTestCode = Trim(RS_L.Fields("TESTCODE"))
'                    lsTestName = Trim(RS_L.Fields("TESTNAME"))
'                    lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                    '-- ���Row �߰�
'                    lsRstRow = .spdResult.DataRowCnt + 1
'                    If .spdResult.MaxRows < lsRstRow Then
'                        .spdResult.MaxRows = lsRstRow
'                    End If
'
'                    '�Ҽ��� ó��, ��� ���� ó��
'                    strMachResult = strResult
'                    strResult = SetResult(strResult, strIntBase)
'                    strJudge = SetJudge(strResult, strIntBase)
'
'                    '������� ǥ��("���")
'                    SetText .spdOrder, "���", gRow, colSTATE
'
'                    '����� ǥ��
'                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                        If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                            SetText .spdOrder, strResult, gRow, intCol
'                            Exit For
'                        End If
'                    Next
'
'                    '-- ��� List
'                    SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '����
'                    SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
'                    SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
'                    SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '�˻��
'                    SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '���ä��
'                    SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '�����
'                    SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS���
'                    SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '����
'                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
'
'                    '-- ���� ����
'                    SetLocalDB gRow, lsRstRow, "1", ""
'
'                    '-- BIORAD QC ����
'                    If Mid(strBarno, 1, 2) = "QC" Then
'                        Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
'                    End If
'
'
'                    strState = "R"
'
'                    '-- ���Count
'                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                        SetText .spdOrder, "1", gRow, colRCNT
'                    Else
'                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                    End If
'
'                End If
'            Else
'                SQL = ""
'                SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                SQL = SQL & "      ,QCLab, QCLot, QCAnalyte, QCMethod, QCInstrument,QCReagent, QCUnit, QCTemp" & vbCr
'                SQL = SQL & "  FROM EQPMASTER" & vbCr
'                SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
'
'                Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                    lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                    lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                    lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                    strQCLab = Trim(RS_L.Fields("QCLab") & "")
'                    strQCLot = Trim(RS_L.Fields("QCLot") & "")
'                    strQCAnalyte = Trim(RS_L.Fields("QCAnalyte") & "")
'                    strQCMethod = Trim(RS_L.Fields("QCMethod") & "")
'                    strQCInstrument = Trim(RS_L.Fields("QCInstrument") & "")
'                    strQCReagent = Trim(RS_L.Fields("QCReagent") & "")
'                    strQCUnit = Trim(RS_L.Fields("QCUnit") & "")
'                    strQCTemp = Trim(RS_L.Fields("QCTemp") & "")
'
'                    '-- ���Row �߰�
'                    lsRstRow = .spdResult.DataRowCnt + 1
'                    If .spdResult.MaxRows < lsRstRow Then
'                        .spdResult.MaxRows = lsRstRow
'                    End If
'
'                    '�Ҽ��� ó��, ��� ���� ó��
'                    strMachResult = strResult
'                    strResult = SetResult(strResult, strIntBase)
'                    strJudge = SetJudge(strResult, strIntBase)
'
'                    '������� ǥ��("���")
'                    SetText .spdOrder, "���", gRow, colSTATE
'
'                    '����� ǥ��
'                    For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                        If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                            SetText .spdOrder, strResult, gRow, intCol
'                            Exit For
'                        End If
'                    Next
'
'                    '-- ��� List
'                    SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '����
'                    SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
'                    SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
'                    SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '�˻��
'                    SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '���ä��
'                    SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '�����
'                    SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS���
'                    SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '����
'                    SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
'
'                    '-- ���� ����
'                    SetLocalDB gRow, lsRstRow, "1", ""
'
'                    '-- BIORAD QC ����
'                    If Mid(strBarno, 1, 2) = "QC" Then
'                        Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
'                    End If
'
'                    If strState <> "R" Then
'                        strState = ""
'                    End If
'
'                    '-- ���Count
'                    If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                        SetText .spdOrder, "1", gRow, colRCNT
'                    Else
'                        SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                    End If
'                End If
'
'            End If
'
'        End If
'
'        .spdResult.RowHeight(-1) = 14
'
'        '## DB�� �������
'        If .optTrans(0).Value = True And strState = "R" Then
'            Res = SaveTransData_MCC(gRow)
'
'            If Res = -1 Then
'                '-- ���� ����
'                SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                SetText .spdOrder, "Failed", gRow, colSTATE
'            Else
'                '-- ���� ����
'                SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                SetText .spdOrder, "����Ϸ�", gRow, colSTATE
'                SetText .spdOrder, "0", gRow, colCHECKBOX
'
'                      SQL = "Update PATRESULT Set " & vbCrLf
'                SQL = SQL & " sendflag = '2' " & vbCrLf
'                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
'                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
'
'                If DBExec(AdoCn_Local, SQL) Then
'                    '-- ����
'                End If
'            End If
'            strState = ""
'        End If
'    End With
'
'End Sub

