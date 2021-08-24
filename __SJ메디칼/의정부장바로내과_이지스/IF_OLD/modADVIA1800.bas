Attribute VB_Name = "modADVIA1800"
Option Explicit


'-----------------------------------------------------------------------------'
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Private Sub SendOrder()
    Dim strOutput   As String     '�۽��� ������
    
    Select Case sSndState
        Case ""
            iIdleFlag = CStr(Val(iIdleFlag) + 1)
            
            '## Order ���� ���
            If mOrder.NoOrder = True Then
                strOutput = ""
                strOutput = intFrameNo & "O" & " " & "0101"
                strOutput = strOutput & "000"                                                   'Sample Count
                strOutput = strOutput & "N"                                                     'Sample classification
                strOutput = strOutput & "2"                                                     'Registration data(0:New, 1:Add, 2:No Request, 3:Sample Delete)
                strOutput = strOutput & Left$(mOrder.BarNo & Space(13), 13)                     'Sample Number
                strOutput = strOutput & Space$(7) & Space$(16) & Space$(16) & "M" & Space$(3)   'Length = 45
                strOutput = strOutput & Space$(8) & " 1.0" & "1" & "1"                          '
                strOutput = strOutput & Space$(1) & ETX
            Else
                '1O 0101010N003498582                                            M            1.011 89M 81M 82M 90M 91M 85M106M103M104M105M 15
                strOutput = ""
                strOutput = intFrameNo & "O" & " " & "0101"
                strOutput = strOutput & Format$(mOrder.SendCnt, "000")                          'Sample Count
                strOutput = strOutput & "N"                                                     'Sample classification
                strOutput = strOutput & "0"                                                     'Registration data(0:New, 1:Add, 2:No Request, 3:Sample Delete)
                strOutput = strOutput & Left$(mOrder.BarNo & Space(13), 13)                     'Sample Number
                strOutput = strOutput & Space$(7) & Space$(16) & Space$(16) & "M" & Space$(3)   'Length = 45
                strOutput = strOutput & Space$(8)                                               '
                strOutput = strOutput & " 1.0"                                                  'Dilution coefficient(4)
                If mOrder.SPCCD = "2" Then                                                      'Sample classification(1:blood serum, 2:urine)
                    strOutput = strOutput & "2"
                Else
                    strOutput = strOutput & "1"
                End If
                strOutput = strOutput & "1"                                                     'Container classification
                strOutput = strOutput & mOrder.Order & Space$(1) & ETX
                
            End If
            
            'n���� sSndPacket ����
            ReDim Preserve sSndPacket(Val(iIdleFlag))
            sSndPacket(Val(iIdleFlag)) = STX & strOutput & GetChkSum(strOutput) & vbCr & vbLf
            
            intFrameNo = intFrameNo + 1
            
        Case "E"  '## ó�� Packet ����
            iOrderFlag = 1
            frmMain.comEqp.Output = sSndPacket(iOrderFlag)
            SetRawData "[Tx]" & sSndPacket(iOrderFlag)
            
            If iOrderFlag = iTotQueryFlag Then
                sSndState = "L"
            Else
                sSndState = "P"
            End If
            
        Case "P"  '## Packet ����
            iOrderFlag = iOrderFlag + 1
            frmMain.comEqp.Output = sSndPacket(iOrderFlag)
            SetRawData "[Tx]" & sSndPacket(iOrderFlag)
            
            If iOrderFlag = iTotQueryFlag Then
                sSndState = "L"
            Else
                sSndState = "P"
            End If
            
        Case "L"  '## EOT
            'strState = ""
            frmMain.comEqp.Output = EOT
            SetRawData "[Tx]" & EOT
            
            iOrderFlag = 0: iPendingFlag = 0: iIdleFlag = 0: iTotQueryFlag = 0
            intFrameNo = 1
            
            Exit Sub
    End Select
    
    If intFrameNo = 8 Then
        intFrameNo = 1
    End If
    
'    strOutput = STX & strOutput & GetChkSum(strOutput) & vbCrLf
'    frmMain.comEqp.Output = strOutput
'    SetRawData "[Tx]" & strOutput

End Sub


Public Sub Phase_Serial_ADVIA1800()
    Dim Buffer      As Variant
    Dim BufChar     As String
    Dim lngBufLen   As Long
    Dim i           As Long

    lngBufLen = Len(pBuffer)

    For i = 1 To lngBufLen
        BufChar = Mid$(pBuffer, i, 1)

        Select Case intPhase
            Case 1      '## Estabilshment Phase
                Select Case BufChar
                    Case ENQ
                        Erase strRecvData
                        
                        sRcvState = "": sSndState = ""
                        intPhase = 2
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    Case Else
                        intPhase = 2
                End Select
            Case 2      '## Transfer Phase
                Select Case BufChar
                    Case STX
                        intBufCnt = 1
                        Erase strRecvData
                        ReDim Preserve strRecvData(intBufCnt)
                    Case EOT
                        Select Case sRcvState
                            Case "Q"
                                intPhase = 3
                                iTotQueryFlag = iPendingFlag
                                iPendingFlag = 0
                                
                                'Order���� Start
                                frmMain.comEqp.Output = ENQ
                                sSndState = "E"
                                
                            Case "R"
                                intPhase = 1
                        End Select
                        
                        sRcvState = ""
                    
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case vbLf
                        intPhase = 2
                        Call SerialRcvData_ADVIA1800
                        
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case vbCr
                    
                    Case ETB
                    
                    Case Else
                        intPhase = 2
                        strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar

                End Select
            Case 3      '## Transfer Phase
                Select Case BufChar
                    Case ACK
                        Select Case sSndState
                            Case "E"        '<ENQ> ���� ���� ����
                                Call SendOrder
                        
                            Case "P"        '<Packet> ���� ���� ����
                                Call SendOrder
                                                
                            Case "L"        '������ <Packet> ���� ���� ����
                                Call SendOrder
                                
                                'Order���� �ʱ�ȭ
                                sSndState = ""
                                Erase sSndPacket
                                intPhase = 1
                        End Select
                    
                    Case ENQ
                        Erase strRecvData
                        frmMain.comEqp.Output = ACK
                        SetRawData "[Tx]" & ACK
                    
                    Case NAK
                        Select Case sSndState
                            Case "E"
                                frmMain.comEqp.Output = Chr(5)
                                intPhase = 3
                            Case "P"
                                frmMain.comEqp.Output = sSndPacket(iOrderFlag)
                                intPhase = 3
                            Case "L"
                                frmMain.comEqp.Output = sSndPacket(iOrderFlag)
                                intPhase = 3
                        End Select
                        
                    Case 4      'EOT
                        Erase strRecvData
                        intPhase = 1
                        sRcvState = "": sSndState = ""
                        'Order���� �ʱ�ȭ
                        iPendingFlag = 0: iTotQueryFlag = 0
                        
                End Select
        End Select
    Next i
            
End Sub


Private Sub SerialRcvData_ADVIA1800()
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

    iBCpos = 2
    
    With frmMain
        For intCnt = 1 To UBound(strRecvData)
            strRcvBuf = strRecvData(intCnt)
            
            '-- �׽�Ʈ�� -----------------
            If .fraCommTest.Visible = False Then
                Call SetSQLData("RCV", strRcvBuf, "A")
            End If
            '-- �׽�Ʈ�� -----------------
            
            strType = Mid$(strRcvBuf, iBCpos, 1)
            
            Select Case strType
                Case "q"    '## Request Information(Batch)
                    sRcvState = "Q"
                    sSndState = ""
                    
                Case "Q"    '## Request Information
                    sRcvState = "Q"
                    sSndState = ""
                
                    iTmpPendingFlag = Val(Mid$(strRcvBuf, iBCpos + 6, 2))
                    iPendingFlag = iPendingFlag + iTmpPendingFlag
                    
                    For i = 1 To iPendingFlag
                        strBarno = Trim$(Mid$(strRcvBuf, iBCpos + 9 + 13 * (i - 1), 13))
                        
                        With mOrder
                            .NoOrder = False
                            .BarNo = strBarno
                        End With
                        
                        Call GetOrder(strBarno, gHOSP.RSTTYPE)
                        Call SendOrder
                    Next
                
                Case "R"
                    sRcvState = "R"
                    
                    iTBlockNo = Val(Mid$(strRcvBuf, iBCpos + 2, 2))
                    iCBlockNo = Val(Mid$(strRcvBuf, iBCpos + 4, 2))
                    iItemNo = Val(Mid$(strRcvBuf, iBCpos + 6, 3))
                    
                    iBCpos = iBCpos + 6
                    
                    strKind = Mid$(strRcvBuf, iBCpos + 17, 1)       'N:Sample, C:Control
                    strBarno = Trim$(Mid$(strRcvBuf, iBCpos + 19, 13))
                                    
                    strTemp2 = Trim$(Mid$(strRcvBuf, iBCpos + 32, 7))
                    iPos = InStr(strTemp2, "-")
                             
                    If iPos = 0 Then
                        strRackNo = ""
                        strTubePos = ""
                    Else
                        strRackNo = Mid$(strTemp2, 1, iPos - 1)
                        strTubePos = Mid$(strTemp2, iPos + 1)
                    End If
                    
                    If strKind = "C" Then       'Control Result
                        strKind = "QC"
                    Else
                        strKind = ""
                    End If
                    
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
                    
                    ReDim Preserve varIntBase(iItemNo)
                    ReDim Preserve varResult(iItemNo)
                    ReDim Preserve varFlag(iItemNo)
                    
                    If iCBlockNo = 1 Then
                        For i = 1 To iItemNo
                            varIntBase(i) = Trim$(Mid(strRcvBuf, iBCpos + 89 + 19 * (i - 1), 3))
                            varResult(i) = Trim(Mid(strRcvBuf, iBCpos + 89 + 4 + 19 * (i - 1), 8))
                            varFlag(i) = Trim(Mid(strRcvBuf, iBCpos + 89 + 8 + 4 + 19 * (i - 1), 3))
                            
                            If InStr(varFlag(i), "R") > 0 Then
                                mResult.Rerun = "R"
                                varFlag(i) = Replace(varFlag(i), "R", "")
                            End If
                        Next i
                    Else
                        For i = 1 To iItemNo
                            varIntBase(i) = Trim$(Mid(strRcvBuf, iBCpos + 39 + 19 * (i - 1), 3))
                            varResult(i) = Trim(Mid(strRcvBuf, iBCpos + 39 + 4 + 19 * (i - 1), 8))
                            varFlag(i) = Trim(Mid(strRcvBuf, iBCpos + 39 + 8 + 4 + 19 * (i - 1), 3))
                            
                            If InStr(varFlag(i), "R") > 0 Then
                                mResult.Rerun = "R"
                                varFlag(i) = Replace(varFlag(i), "R", "")
                            End If
                        Next i
                    End If
                    
                    If mResult.Rerun = "R" Then       'Rerun Result
                        mResult.Kind = mResult.Kind & "R"
                    End If
                    
                    For i = 1 To iItemNo
                        strIntBase = varIntBase(i)
                        strResult = varResult(i)
                        
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
                    Next
                    
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
            End Select
        Next
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
        strItems = GetEquipExamCode_ADVIA1800(gHOSP.MACHCD, pBarno, intRow)

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
Private Function GetEquipExamCode_ADVIA1800(argEquipCode As String, argPID As String, Optional intRow As Long) As String
    Dim i As Integer
    Dim sExamCode As String
    Dim strExamCode As String
    Dim sSpecNo     As String
    Dim iRow        As Long
    Dim SpecNo      As String

    GetEquipExamCode_ADVIA1800 = ""
    
    If Trim(argEquipCode) = "" Or gPatOrdCd = "" Then
        Exit Function
    End If
    
    '-- ������ �˻��ڵ��� ä�� ã��
          SQL = "Select DISTINCT SENDCHANNEL "
    SQL = SQL & "  From EQPMASTER "
    SQL = SQL & " Where EQUIPCD  = '" & Trim(gHOSP.MACHCD) & "' "
    SQL = SQL & "   and TESTCODE IN (" & Trim(gPatOrdCd) & ")"
    
    strExamCode = ""
    mOrder.SendCnt = 0
    
    AdoCn_Local.CursorLocation = adUseClient
    Set AdoRs_Local = AdoCn_Local.Execute(SQL, , 1)
    If Not AdoRs_Local.EOF = True And Not AdoRs_Local.BOF = True Then
        Do Until AdoRs_Local.EOF
            ' " 89M 81M 82M 90M 91M108M 85M"
            strExamCode = strExamCode & Right(Space(3) & Trim(AdoRs_Local.Fields("SENDCHANNEL").Value & ""), 3) & "M"
            mOrder.SendCnt = mOrder.SendCnt + 1
            AdoRs_Local.MoveNext
        Loop
    End If
    
    AdoRs_Local.Close
    
    GetEquipExamCode_ADVIA1800 = Mid(strExamCode, 2)
    
End Function



