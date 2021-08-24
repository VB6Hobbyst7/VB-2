Attribute VB_Name = "modHITACHI7020"
Option Explicit


'-----------------------------------------------------------------------------'
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput   As String     '�۽��� ������
    
    strOutput = ";" & mOrder.Function
    strOutput = strOutput & " 37"
    strOutput = strOutput & Mid(mOrder.Order, 1, 37)
    strOutput = strOutput & "00000"
    
    'COMMENT���� BARCODE ǥ��
    'strOutput = strOutput & "100000" & Left(mOrder.BarNo & Space(30), 30)
    
    Call Sleep(100)
    
    '-- SPE Send(��������)
    frmMain.comEqp.Output = STX & strOutput & ETX & vbCr & vbLf
    
    SetRawData "[Tx]" & STX & strOutput & ETX & vbCr & vbLf


End Sub


Public Sub Phase_Serial_HITACHI7020()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(Buffer, i, 1)
        Select Case BufChar
            Case STX
                RcvBuffer = ""
                
            Case ETX
                 Call SerialRcvData_HITACHI7020
                 RcvBuffer = ""
            
            Case Else
                RcvBuffer = RcvBuffer & BufChar
        End Select
    Next i
            
End Sub

Public Sub SndMore()
    Dim strSndMsg As String
    
    strSndMsg = ">"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) ' & GetChkSum(strSndMsg) & vbCr
    strSndMsg = strSndMsg & vbCrLf
    
    frmMain.comEqp.Output = strSndMsg
    
    SetRawData "[Tx]" & strSndMsg
    
End Sub

Public Sub SndRec()
    Dim strSndMsg As String
    
    strSndMsg = "A"
    strSndMsg = Chr(2) & strSndMsg & Chr(3) '& GetChkSum(strSndMsg)
    strSndMsg = strSndMsg & vbCrLf
    
    frmMain.comEqp.Output = strSndMsg
    
    SetRawData "[Tx]" & strSndMsg
    
End Sub


Private Sub SerialRcvData_HITACHI7020()
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
    
    Dim strQCData       As String
    Dim strQCRun, strQCLevel, strQCLab, strQCLot, strQCAnalyte, strQCMethod, strQCInstrument, strQCReagent, strQCUnit, strQCTemp As String
    
    Dim strTmp          As String
    Dim i               As Integer
    Dim iBCpos          As Integer
    
    Dim iTBlockNo   As Integer
    Dim iCBlockNo   As Integer
    Dim iItemNo     As Integer
    Dim strKind     As String
    Dim iPos        As Integer
    
    Dim varIntBase()    As String
    Dim varResult()     As String
    Dim varFlag()       As String

    'for H7020
    Dim strFunc       As String
    Dim strFunction   As String
    Dim strSendData   As String
    Dim strSndMsg     As String
    Dim strExamCode() As String
    
    With frmMain
        strRcvBuf = RcvBuffer
        
        '-- �׽�Ʈ�� -----------------
        If .fraCommTest.Visible = False Then
            Call SetSQLData("RCV", strRcvBuf, "A")
        End If
        '-- �׽�Ʈ�� -----------------
        
        strType = Mid$(strRcvBuf, 1, 1)
        If IsNumeric(strType) Then
            strType = Mid$(strRcvBuf, 2, 1)
        End If
        
        Select Case strType
            Case ">", "?", "@"      'ANY ����
                Call SndMore        'MOR Send
                Do
                '   DoEvents
                Loop Until frmMain.comEqp.OutBufferCount = 0
            
            Case "?", "@"           'REP ����
                Sleep (100)
                Call SndMore        'MOR Send
                Do
                '   DoEvents
                Loop Until frmMain.comEqp.OutBufferCount = 0
            
            Case ">", "?", "@"      'SUS ����
                Sleep (100)
                Call SndMore        'MOR Send
                Do
                '   DoEvents
                Loop Until frmMain.comEqp.OutBufferCount = 0
            
            Case ";"                'SPE  ����(������û)
                strFunction = Mid(strRcvBuf, 2, 12) & String(13, "#") & Mid(strRcvBuf, 27, 15)
            
                strFunc = Mid(strRcvBuf, 2, 1)              'N
                strSeq = Mid(strRcvBuf, 4, 5)               '    1
                strRackNo = Mid(strRcvBuf, 9, 1)            '
                strTubePos = Mid(strRcvBuf, 10, 3)          '  1
                strBarno = Trim(Mid(strRcvBuf, 14, 13))

                With mOrder
                    .NoOrder = False
                    .BarNo = strBarno
                    .Func = strFunc
                    .Function = strFunction
                End With
                
                Call GetOrder(strBarno, gHOSP.RSTTYPE)
                
                Call SendOrder
        
            ' FR1 to FR9 (�˻��׸� 25�� �̻��� ��� ó��)
            Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
                strFunc = Mid(strRcvBuf, 2, 1)
                
                If strFunc = "K" Or strFunc = "L" Or strFunc = "G" Or strFunc = "H" Then
                    Sleep (100)
                    Call SndMore        'MOR Send
                    Do
                    '   DoEvents
                    Loop Until frmMain.comEqp.OutBufferCount = 0
                    Exit Sub
                End If
                            
                Call SndMore            'MOR Send
                
                If strFunc <> "@" And strFunc <> "M" Then
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Trim(Mid(strRcvBuf, 10, 3))
                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
                    gRow = 0
                                      
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Kind = strKind
                        .Rerun = ""
                        If strOldBarno <> strBarno Then
                            strOldBarno = strBarno
                            .RsltDate = Format(Now, "yyyymmddhhmmss")
                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    
                            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                            
                        End If
                    End With
                    
                    'strTmp = Mid$(strRcvBuf, 29)
                    strTmp = Mid$(strRcvBuf, 45)
    
                    For i = 44 To Len(strRcvBuf) Step 10
                        strIntBase = Trim(Mid(strRcvBuf, i, 3))
                        strIntBase = Format(strIntBase, "00")
                        strResult = Trim(Mid(strRcvBuf, i + 3, 6))
                        
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
'                                    If mResult.Kind = "QC" Then
'                                        strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                        Call SendBioRadQC(strQCData)
'                                    End If
                                    
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
'                                    If mResult.Kind = "QC" Then
'
'                                        strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
'
'                                        Call SendBioRadQC(strQCData)
'
'                                    End If
                                    
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
                        strTmp = Mid$(strTmp, 12)
                    Next
                    
                    .spdResult.RowHeight(-1) = 14
                
                    '## DB�� �������
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_EASYS(gRow)
                        
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
                
            ' ��� END
            Case ":"
                ':N     3   3 3                            4  4    17   5    16  11   9.3  12  0.89 
                strFunc = Mid(strRcvBuf, 2, 1)
                
                If strFunc = "K" Or strFunc = "L" Or strFunc = "G" Or strFunc = "H" Then
                    Sleep (100)
                    Call SndMore        'MOR Send
                    Do
                    '   DoEvents
                    Loop Until frmMain.comEqp.OutBufferCount = 0
                    Exit Sub
                End If
                
                If strFunc = "K" Or strFunc = "L" Then
                    Call SndMore        'MOR Send
                    Exit Sub
                End If
                
                Call SndMore            'MOR Send
                
                
                If strFunc <> "@" And strFunc <> "M" Then
                    strRackNo = Mid(strRcvBuf, 9, 1)
                    strTubePos = Trim(Mid(strRcvBuf, 10, 3))
                    strBarno = Trim(Mid(strRcvBuf, 14, 13))
                    
                    With mResult
                        .BarNo = strBarno
                        .RackNo = strRackNo
                        .TubePos = strTubePos
                        .Kind = strKind
                        .Rerun = ""
                        If strOldBarno <> strBarno Then
                            strOldBarno = strBarno
                            .RsltDate = Format(Now, "yyyymmddhhmmss")
                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
                    
                            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
                            
                        End If
                    End With
                    
                    'strTmp = Mid$(strRcvBuf, 29)
                    strTmp = Mid$(strRcvBuf, 45)
    
                    For i = 44 To Len(strRcvBuf) Step 10
                        strIntBase = Trim(Mid(strRcvBuf, i, 3))
                        strIntBase = Format(strIntBase, "00")
                        strResult = Trim(Mid(strRcvBuf, i + 3, 6))
                        
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
                                    If mResult.Kind = "QC" Then
                                        strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                        
                                        Call SendBioRadQC(strQCData)
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
                                    If mResult.Kind = "QC" Then
                                        
                                        strQCData = GetQCResult_Detail(gHOSP.LABCD, strBarno, strQCAnalyte, strResult)
                                        
                                        Call SendBioRadQC(strQCData)
                                        
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
                        strTmp = Mid$(strTmp, 12)
                    Next
                    
                    .spdResult.RowHeight(-1) = 14
                
                    '## DB�� �������
                    If .optTrans(0).Value = True And strState = "R" Then
                        Res = SaveTransData_EASYS(gRow)
                        
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
        strItems = GetEquipExamCode_HITACHI7020(gHOSP.MACHCD, pBarno, intRow)

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
Private Function GetEquipExamCode_HITACHI7020(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode()   As String
    Dim sSpecNo         As String
    Dim iRow            As Long
    Dim SpecNo          As String
    
    Dim strSendData As String
    Dim ii          As Integer
    Dim strTestNum  As String
    
            
    GetEquipExamCode_HITACHI7020 = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    strSendData = String$(88, "0")
    
    '-- ������ �˻��ڵ��� ä�� ã��
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    Erase strExamCode
    mOrder.SendCnt = 0
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            strTestNum = Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & "")
            If strTestNum <> "" Then
                ReDim Preserve strExamCode(ii)
                strExamCode(ii) = strTestNum
                ii = ii + 1
                mOrder.SendCnt = mOrder.SendCnt + 1
            End If
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    If gPatOrdCd <> "" And ii > 0 Then
        For i = 0 To UBound(strExamCode)
            If strExamCode(i) <> "" Then
                If strExamCode(i) <> "99" Then
                    Mid(strSendData, strExamCode(i), 1) = "1"
                End If
            End If
        Next
    End If
    
    GetEquipExamCode_HITACHI7020 = strSendData
    
End Function




