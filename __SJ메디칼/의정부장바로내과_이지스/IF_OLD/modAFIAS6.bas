Attribute VB_Name = "modAFIAS6"
Option Explicit

Public Sub Phase_Serial_AFIAS6()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)
    
    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case "$" 'SOH
                Erase strRecvData
                
                intBufCnt = 1
                ReDim Preserve strRecvData(intBufCnt)
            Case vbCr
                Call SerialRcvData_AFIAS6
            Case Else
                If intBufCnt > 0 Then
                    strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
                End If
        End Select
    Next i

End Sub


Public Sub SerialRcvData_AFIAS6()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strBarno        As String   '������ ���ڵ��ȣ
    Dim strSeq          As String   '������ Sequence
    Dim strRackNo       As String   '������ Rack Or Disk No
    Dim strTubePos      As String   '������ Tube Position
    Dim strIntBase      As String   '������ ������ �˻��
    Dim strMachResult   As String   '������ �����
    Dim strResult       As String   '������ ���(����)
    Dim strIntResult    As String   '������ ���(����)
    Dim strQCResult     As String   '������ ���(QC)
    Dim strFlag         As String   '������ Abnormal Flag
    Dim strComm         As String   '������ Comment
    
    Dim lsOrderCode     As String   'ó���ڵ�
    Dim lsTestCode      As String   '�˻��ڵ�
    Dim lsTestName      As String   '�˻��
    Dim lsSeqNo         As String   '����DB �˻�Seq
    
    Dim lsRstRow        As String   '����������� ���� Row
    Dim intCnt          As Integer  '��� Frame ����
    Dim intCol          As Integer  '����÷� ����
    Dim strJudge        As String   '�������
    Dim Res             As Integer
    
    Dim strTmp          As String
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- �׽�Ʈ�� -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- �׽�Ʈ�� -----------------
            
            strRcvBuf = strRecvData(intCnt)
            strBarno = Trim(mGetP(strRcvBuf, 5, "|"))
            strRackNo = ""
            strTubePos = ""
            strSeq = ""
                        
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
                        
            strIntBase = mGetP(strRcvBuf, 8, "|")
            strResult = mGetP(strRcvBuf, 11, "|")
                        
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

                        '-- ���Row �߰�
                        lsRstRow = .spdResult.DataRowCnt + 1
                        If .spdResult.MaxRows < lsRstRow Then
                            .spdResult.MaxRows = lsRstRow
                        End If

                        '�Ҽ��� ó��, ��� ���� ó��
                        strMachResult = strResult
                        strResult = SetResult(strResult, strIntBase)
                        strJudge = SetJudge(strResult, strIntBase)
                        
                        '������� ǥ��("���")
                        SetText .spdOrder, "���", gRow, colSTATE

                        '����� ǥ��
                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                SetText .spdOrder, strResult, gRow, intCol
                                Exit For
                            End If
                        Next

                        '-- ��� List
                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '����
                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '�˻��
                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '���ä��
                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '�����
                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS���
                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '����
                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
                        
                        '-- ���� ����
                        SetLocalDB gRow, lsRstRow, "1", ""
                        
                        strState = "R"
                        
                        '-- ���Count
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

                        '-- ���Row �߰�
                        lsRstRow = .spdResult.DataRowCnt + 1
                        If .spdResult.MaxRows < lsRstRow Then
                            .spdResult.MaxRows = lsRstRow
                        End If

                        '�Ҽ��� ó��, ��� ���� ó��
                        strMachResult = strResult
                        strResult = SetResult(strResult, strIntBase)
                        strJudge = SetJudge(strResult, strIntBase)
                        
                        '������� ǥ��("���")
                        SetText .spdOrder, "���", gRow, colSTATE

                        '����� ǥ��
                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
                                SetText .spdOrder, strResult, gRow, intCol
                                Exit For
                            End If
                        Next

                        '-- ��� List
                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '����
                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '�˻��
                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '���ä��
                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '�����
                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS���
                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '����
                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
                        
                        '-- ���� ����
                        SetLocalDB gRow, lsRstRow, "1", ""
                        
                        If strState <> "R" Then
                            strState = ""
                        End If

                        '-- ���Count
                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
                            SetText .spdOrder, "1", gRow, colRCNT
                        Else
                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
                        End If
                    End If
                    
                End If
                
            End If
                        
            .spdResult.RowHeight(-1) = 14
                        
            '## DB�� �������
            If .optTrans(0).Value = True And strState = "R" Then
                Res = SaveTransData_MCC(gRow)
                
                If Res = -1 Then
                    '-- ���� ����
                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
                    SetText .spdOrder, "Failed", gRow, colSTATE
                Else
                    '-- ���� ����
                    SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
                    SetText .spdOrder, "����Ϸ�", gRow, colSTATE
                    SetText .spdOrder, "0", gRow, colCHECKBOX
                    
                          SQL = "Update PATRESULT Set " & vbCrLf
                    SQL = SQL & " sendflag = '2' " & vbCrLf
                    SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
                    SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
                    SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
                    SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
                    
                    If DBExec(AdoCn_Local, SQL) Then
                        '-- ����
                    End If
                End If
                strState = ""
            End If
        Next
    End With

End Sub
