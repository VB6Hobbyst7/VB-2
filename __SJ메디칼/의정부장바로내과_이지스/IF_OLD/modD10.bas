Attribute VB_Name = "modD10"
Option Explicit

'-----------------------------------------------------------------------------'
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput   As String     '�۽��� ������
    
    'intFrameNo = CStr(Val(iIdleFlag) Mod 8)
    
    Select Case intSndPhase
        Case 1  '## Header
            strOutput = intFrameNo & "H|@^\|<0_0><1025080549_50>||HOSTTOP|||||ACLTOP||P|1394-97|" & Format(Now, "YYYYMMDDHHMMSS") & vbCr & ETX
            intSndPhase = 2
            intFrameNo = intFrameNo + 1
        
        Case 2  '## Patient
            strOutput = intFrameNo & "P|1||" & mOrder.BarNo & "||^||||||||" & vbCr & ETX
            intSndPhase = 3
            intFrameNo = intFrameNo + 1
            
        Case 3  '## Order
            If mOrder.NoOrder = True Then
                '## ���������� �������
                strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||||||1"
                '-- 2015.10.15 �߰�
                strOutput = intFrameNo & strOutput & vbCr & ETX
                intSndPhase = 4
            
            Else
                '## ���� ������
                If mOrder.IsSending = False Then
                    strOutput = "O|1|" & mOrder.BarNo & "||" & mOrder.Order & "|R||||||A||||P||||||||||Q" & Chr(13)
                    
                    If Len(strOutput) > 230 Then
                        mOrder.IsSending = True
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                '## ���� ���ڿ��� ������
                Else
                    strOutput = mOrder.Order
                    If Len(strOutput) > 230 Then
                        mOrder.Order = Mid$(strOutput, 231)
                        strOutput = intFrameNo & Mid$(strOutput, 1, 230) & vbCr & ETB
                        intSndPhase = 3
                    Else
                        mOrder.IsSending = False
                        strOutput = intFrameNo & strOutput & vbCr & ETX
                        intSndPhase = 4
                    End If
                End If
            End If
            intFrameNo = intFrameNo + 1
            
        Case 4  '## Termianator
            strOutput = intFrameNo & "L|1|N" & vbCr & ETX
            intSndPhase = 5
            intFrameNo = intFrameNo + 1
            
        Case 5  '## EOT
            strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            intFrameNo = 1
            
            Exit Sub
    End Select
    
    If intFrameNo = 8 Then
        intFrameNo = 1
    End If
    
    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
    frmMain.comEqp.Output = strOutput
    SetRawData "[Tx]" & strOutput

End Sub


Public Sub Phase_Serial_D10()
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
            Case ENQ
                    RcvBuffer = ""
                    frmMain.comEqp.Output = ACK
                    SetRawData "[Tx]" & ACK
            
            Case vbLf
                    Call SerialRcvData_D10
                    RcvBuffer = ""
                    frmMain.comEqp.Output = ACK
                    SetRawData "[Tx]" & ACK
            Case EOT
                    RcvBuffer = ""
            Case Else
                    RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
            
End Sub


Private Sub SerialRcvData_D10()
    Dim RS_L            As ADODB.Recordset
    Dim strRcvBuf       As String   '������ Data
    Dim strType         As String   '������ Record Type
    Dim strOldBarno        As String   '������ ���ڵ��ȣ
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
    Dim strAspect       As String
    
    Dim strTemp1        As String
    Dim strTemp2        As String
    
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
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    With frmMain
        strRcvBuf = RcvBuffer
        
        '-- �׽�Ʈ�� -----------------
        If .fraCommTest.Visible = False Then
            Call SetSQLData("RCV", strRcvBuf, "A")
        End If
        '-- �׽�Ʈ�� -----------------
        
        strType = Mid$(strRcvBuf, 2, 1)
        If strType = "|" Then
            strType = Mid$(strRcvBuf, 1, 1)
        End If
        
        Select Case strType
            Case "H"    '## Header
            Case "P"    '## Patient
            Case "O"
                strBarno = mGetP(strRcvBuf, 3, "|")
                
                With mResult
                    .BarNo = strBarno
                    '.SpcPos = strTubePos & "/" & strRackNo
                    '.Seq = strSeq
                    '.RackNo = strRackNo
                    '.TubePos = strTubePos
                    If strOldBarno <> strBarno Then
                        strOldBarno = strBarno
                        .RsltDate = Format(Now, "yyyymmddhhmmss")
                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                
                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                        
                    End If
                End With
            
            Case "R"
                strTemp1 = mGetP(strRcvBuf, 3, "|")
                strIntBase = mGetP(strTemp1, 4, "^") & "^" & mGetP(strTemp1, 5, "^")
                strResult = mGetP(strRcvBuf, 4, "|")
                                            
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
                            
                            '-- BIORAD QC ����
                            If Mid(strBarno, 1, 2) = "QC" Then
                                Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
                            End If
                        
                            
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
                            
                            '-- BIORAD QC ����
                            If Mid(strBarno, 1, 2) = "QC" Then
                                Call MakeBioRadQC(gHOSP.MACHCD, strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp, strResult)
                            End If
                            
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
            
            Case "C"    '## Comment
                '## Abnormal ����϶� Comment ����
                If strFlag <> "N" Then
                    strTemp1 = mGetP(strRcvBuf, 4, "|")
                    strComm = mGetP(strTemp1, 1, "^") & ", " & mGetP(strTemp1, 2, "^")
                End If
            
            Case "L"
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
        End Select
    End With

End Sub



'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڵ��ȣ�� ���� 1. �������� ��ȸ,
'                                 2. ���������� ȭ��ǥ��,
'                                 3. ó���ڵ� ��������,
'                                 4. (ó���ڵ��)�˻���� �����
'   �μ� :
'       - pBarNo : ���ڵ��ȣ
'       - pType  : ���ڵ� �̻��� ���ϴ� ���
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : üũ�Ȱ��� ���� ���� ��
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
    
    '-- 1. �������� ��ȸ
    With frmMain
        '-- ���ڵ� ���
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
        
        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mOrder.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mOrder.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mOrder.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mOrder.TubePos, intRow, colPOSNO)
            
        '-- ����������� �����
        .spdResult.MaxRows = 0
    
        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
        
        '-- �������̺��� �˻��׸� �ش��ϴ� �˻�ä�� ã�ƿ��� (intRow = ���� �˻��ߴ� ���ڵ尡 �ٽ� �ö�� ��� ��ġ�� ��ã�´�.)
        strItems = GetEquipExamCode_D10(gHOSP.MACHCD, pBarno, intRow)

        '-- �˻�ä�η� ������ �����
        If Trim(strItems) = "" Then
            mOrder.NoOrder = True
            mOrder.Order = ""
            
            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        Else
            mOrder.NoOrder = False
            mOrder.Order = strItems
        
            '-- �������(Order) ǥ��
            Call SetText(frmMain.spdOrder, "��������", intRow, colSTATE)
        End If


        '-- ���� Row
        gRow = intRow
        
    End With
    
End Sub

'��ü��ȣ�� �����ϴ� ����ȣ �ش��ϴ� �����ڵ� ��������
'�� ��� ��ȣ�� �˻��ڵ尡 1���̻� ����
Private Function GetEquipExamCode_D10(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_D10 = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    '-- ������ �˻��ڵ��� ä�� ã��
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    strExamCode = ""
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strExamCode = strExamCode & "\^^^" & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_D10 = Mid(strExamCode, 2)
    
End Function



