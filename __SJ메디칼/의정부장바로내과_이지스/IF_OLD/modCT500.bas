Attribute VB_Name = "modCT500"
Option Explicit

Public Sub Phase_Serial_CT500()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
        Select Case BufChar
            Case STX
                RcvBuffer = ""
                
                miLineNo = 0
            Case vbCr
                Call SerialRcvData_CT500
                
                miLineNo = 1
                
                RcvBuffer = ""
            
            Case ETX
                RcvBuffer = ""
                miLineNo = 0
                
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i

End Sub


Public Sub SerialRcvData_CT500()
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
    Dim varResult       As Variant
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
    Dim strOldBarno     As String
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    With frmMain
        strRcvBuf = RcvBuffer
        strRcvBuf = Replace(strRcvBuf, vbLf, "")
        
'#4-723      17-08-28
'ID = 3495464
'Color: STRAW
'Clarity:
'GLU NEGATIVE
'BIL NEGATIVE
'KET NEGATIVE
'SG 1.025
'BLO NEGATIVE
'pH 6#
'PRO NEGATIVE
'URO      0.2 E.U./dL
'NIT NEGATIVE
'LEU NEGATIVE
'


        '-- �׽�Ʈ�� -----------------
        If .fraCommTest.Visible = False Then
            Call SetSQLData("RCV", strRcvBuf, "A")
        End If
        '-- �׽�Ʈ�� -----------------
        
        If Mid(strRcvBuf, 1, 3) = "ID=" Then
            miLineNo = 1
            'strBarno = Trim(Mid(strRcvBuf, 5, 12))
            strBarno = Trim(Mid(strRcvBuf, 4))
            mResult.BarNo = strBarno
    
            With mResult
                .BarNo = strBarno
                .SpcPos = strSeq
                .Seq = strSeq
                .RackNo = strRackNo
                .TubePos = strTubePos
                If strOldBarno <> strBarno Then
                    strOldBarno = strBarno
                    .RsltDate = Format(Now, "yyyymmddhhmmss")
                    .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                End If
            End With
            
            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
            If gRow <= 0 Then
                Exit Sub
            End If
        Else
            If miLineNo = 1 Then
                strIntBase = Trim(mGetP(strRcvBuf, 1, Space$(1)))
                If Right(strIntBase, 1) = "*" Then
                    strIntBase = Mid(strIntBase, 1, Len(strIntBase) - 1)
                End If
                strResult = Trim(mGetP(strRcvBuf, 2, Space$(1)))
                If strResult = "" Then
                    If Len(strIntBase) = 3 Then
                        strResult = Trim(Mid(strRcvBuf, 8))
                    Else
                        strResult = Trim(Mid(strRcvBuf, 9))
                    End If
                End If
                strResult = Replace(strResult, "E.U./dL", "")
                strResult = Trim(strResult)
                
                '--QC
                If Len(mResult.BarNo) <= 5 Then
                    strResult = Replace(strResult, "<", "")
                    strResult = Replace(strResult, ">", "")
                    strResult = Replace(strResult, "=", "")
                End If
                
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
                            
                            '-- BIORAD QC ����
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
                            End If
                    
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
        
                            '-- BIORAD QC ����
                            If mResult.Kind = "QC" Then
                                strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                
                                Call SendBioRadQC(strQCData)
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
            End If
        End If
    End With

End Sub


