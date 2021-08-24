Attribute VB_Name = "modOSMOMETER"
Option Explicit

Public Sub Phase_Serial_OSMOMETER()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)
'        Select Case Asc(BufChar)
'            Case 5      'ENQ
'                frmMain.comEqp.Output = ACK
'            Case 0      '--
'                RcvBuffer = ""
'
'            Case 2      'STX
'                RcvBuffer = ""
'
'            Case 13      'ETX
'                Call SerialRcvData_OSMOMETER
'
'                RcvBuffer = ""
'                frmMain.comEqp.Output = ACK
'            Case 10
'                RcvBuffer = ""
'
'            Case 4      'EOT
'                RcvBuffer = ""
'
'            Case Else
'                RcvBuffer = RcvBuffer & BufChar
'        End Select
    
        Select Case Asc(BufChar)
            Case 13     'CR
                If Trim(RcvBuffer) <> "" Then
                    Call SerialRcvData_OSMOMETER
                    
                    RcvBuffer = ""
                End If
                
            Case 10     'LF
            
            Case Else
                RcvBuffer = RcvBuffer & BufChar

        End Select
    
    
    Next i

End Sub


Public Sub SerialRcvData_OSMOMETER()
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

        '-- �׽�Ʈ�� -----------------
        If .fraCommTest.Visible = False Then
            Call SetSQLData("RCV", strRcvBuf, "A")
        End If
        '-- �׽�Ʈ�� -----------------
        
        '���� ���۵Ǵ� �����
        '    Osmometer Ready
        '
        '  11/16/2006  02:21 AM
        'Osmolality  293 mOsm
        
        '1. Recall Results
        '#30: 284 mOsm     [PREV]
        'ID NONE
        '#29: 280 mOsm     [PREV]
        '#28: 296 mOsm     [PREV]
        '#27: 639 mOsm     [PREV]
        '#26: 288 mOsm     [PREV]
        '#25: 291 mOsm     [PREV]
        '#24: 381 mOsm     [PREV]
        '#23: 302 mOsm     [PREV]
            
        
        If InStr(strRcvBuf, "Ready") > 0 Or InStr(strRcvBuf, "Recall Results") > 0 Then
            Exit Sub
        End If
            
        'Realtime
        If InStr(strRcvBuf, "Osmolality") > 0 Then
            strResult = Trim(Mid(strRcvBuf, 11, 5))
            
        'Recall Result
        ElseIf Left(strRcvBuf, 1) = "#" And InStr(strRcvBuf, "mOsm") > 0 Then
            strResult = Trim(mGetP(mGetP(Trim(strRcvBuf), 2, ":"), 1, "mOsm"))
        Else
            Exit Sub
        End If
        
        strIntBase = "OSMO"
        

                
        With mResult
            .BarNo = ""
            .SpcPos = ""
            .Seq = ""
            .RackNo = ""
            .TubePos = ""
            'If strOldBarno <> strBarno Then
                'strOldBarno = strBarno
                .RsltDate = Format(Now, "yyyymmddhhmmss")
                .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
            'End If
        End With
        
        'Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                    
        If gRow <= 0 Then
            Exit Sub
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
    End With

End Sub


