Attribute VB_Name = "modCommunication"
Option Explicit

Public pBuffer As Variant

'-- ������ ��������
Type RecvData
    OrgBarNo    As String
    BarNo       As String
    Seq         As String
    RackNo      As String
    TubePos     As String
    NoOrder     As Boolean
    Order       As String
    IsSending   As Boolean
    SendCnt     As Integer
    isresult    As Boolean
    PID         As String
    SPCCD       As String
    SampleData  As String
    'for PLIS
    WA          As String
    AccSeq      As Long
    'for ACLTOP
    MsgID       As String
    Sender      As String
    Receiver    As String
    Version     As String
    PName       As String
    Count       As Integer
    Items()     As String
    'for H7180
    Func        As String
    Function    As String
    'for LH780
    BlkCnt      As Integer
    'for AU480
    SmpType     As String
End Type

Public mOrder As RecvData

'-- ������ �������
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
    TestCd   As String
    Kind     As String
    Rerun    As String
    IntBase  As String
    RESULT   As String
    EqpCd    As String
    '-- ��Ʈ��ũ����Ʈ��
    RESODRSEQ   As String
    RESSEQ      As String
    RESSUBSEQ   As String
End Type

Public mResult As IntfData

'for ADVIA1650
Public iPendingFlag    As Integer
Public iTotQueryFlag   As Integer
Public iTmpPendingFlag As Integer
Public iIdleFlag   As Integer
Public iOrderFlag  As Integer
Public iResultFlag As Integer
Public sRcvState   As String
Public sSndState   As String
Public sSndPacket()    As String
Public sQueryBarcd()   As String

'for ADVIA2120
Public Const mc_sSampleType    As String = "1"
Public Const mc_sPatInfo       As String = ""
Public Const mc_sSampInfo      As String = ""
Public Const mc_sSiteNm        As String = ""
Public Const mc_sRerunGbn      As String = ""
Public Const mc_bSerumIndex    As Integer = False
Public Const mc_sEqName        As String = ""
Public Const mc_bUseBarcode    As Boolean = False
Public Const mc_iPhase         As Integer = 1
Public Const mc_iSendPhase     As Integer = 1
Public Const mc_sTestMode      As String = "0"
Public Const mc_iFrameN        As Integer = 1
Public Const mc_sID            As String = ""
Public Const mc_sSeq           As String = ""
Public Const mc_sRack          As String = ""
Public Const mc_sPos           As String = ""
Public Const mc_iOrdCnt        As Integer = 0
Public Const mc_sTIFCd         As String = ""
Public Const mc_bPortOpen      As Boolean = False
Public Const mc_sOpenPW        As String = ""
Public Const mc_sEditPW        As String = ""
Public Const mc_bReserveEnd    As Boolean = False

'�Ӽ� ����:
Public mp_sSampleType          As String
Public mp_sPatInfo             As String
Public mp_sSampInfo            As String
Public mp_sSiteNm              As String
Public mp_sRerunGbn            As String
Public mp_bSerumIndex          As Boolean
Public mp_sEqName              As String
Public mp_bUseBarcode          As Boolean
Public mp_iPhase               As Integer
Public mp_iSendPhase           As Integer
Public mp_sTestMode            As String
Public mp_iFrameN              As Integer
Public mp_sID                  As String
Public mp_sSeq                 As String
Public mp_sRack                As String
Public mp_sPos                 As String
Public mp_iOrdCnt              As Integer
Public mp_sTIFCd               As String
Public mp_bPortOpen            As Boolean
Public mp_sOpenPW              As String
Public mp_sEditPW              As String
Public mp_bReserveEnd          As Boolean

Public Const mc_iMaxCnt     As Integer = 100
Public msMT                 As String
Public msTimerFlag          As String
Public msSndPacket          As String

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'for PFA-200, CT500
Public miLineNo             As Integer


'for RAPIDPOINT500
Public aMod                 As String
Public iIID                 As String
Public AckOn                As Boolean
Public Sample_Seq           As String

'for ACLTOP
Public mPNo                 As Integer

'===== User Define
'�������̽����� ���
Public strFRcvBuffer   As String
Public strFWkBuf       As String
Public strFState       As String
Public blnFSend        As Boolean
Public blnFEndChk      As Boolean
Public blnFSTXChk      As Boolean
Public strFRstEnd      As String

Public strFRcvState    As String
Public strFSndState    As String
Public msAllBarCd   As String
Public maAllBarCd() As String
Public TimerFlag    As Integer
Public SavBuffer    As String
Public ii_SendCnt   As Integer
Public m_aTemp()    As String
Public miSendCnt    As Integer
Public msSendBuff   As String

'�Ӽ� ����:
Public m_p_sPatInfo As Variant
Public m_EqName As String
Public m_bUseBarcode As Boolean
Public m_iPhase As Integer
Public m_iSendPhase As Integer
Public m_sTestMode As String
Public m_iFrameN As Integer
Public m_p_sID As String
Public m_p_sSeq As String
Public m_p_sRack As String
Public m_p_sPos As String
Public m_p_iOrdCnt As Integer
Public m_p_sTIFCd As String
Public m_PortOpen As Boolean
Public m_OpenPW As String
Public m_EditPW As String
Public m_IFMode As String

'for IF
Public strOldBarno          As String   '������ ���ڵ��ȣ

'for CT500
Public mColor               As Boolean

Public Sub Serial_Protocol()

    SetRawData "[Rx]" & pBuffer
        
    Select Case UCase(gHOSP.MACHNM)
        
'        Case "AFIAS6"
'                Call Phase_Serial_AFIAS6
                
'        Case "VERSACELL"
'                Call Phase_Serial_VERSACELL
                
'        Case "ADVIA1800-1", "ADVIA1800-2"
'                Call Phase_Serial_ADVIA1800
                
'        Case "ADVIA2120-1", "ADVIA2120-2"
'                Call Phase_Serial_ADVIA2120
                
'        Case "RAPIDLAB348"
'                Call Phase_Serial_RAPIDLAB348
                
'        Case "RAPIDPOINT500"
'                Call Phase_Serial_RAPIDPOINT500
        
'        Case "PFA200"
'                Call Phase_Serial_PFA200
                
'        Case "ACLTOP"
'                Call Phase_Serial_ACLTOP
                
'
'        Case "VESCUBE"
'                Call Phase_Serial_VESCUBE
                
                
'        Case "CT500"
'                Call Phase_Serial_CT500
                
        Case Else
            
    End Select
    

End Sub

'Public Sub TCP_Protocol()
'
'    Select Case UCase(gHOSP.MACHNM)
'        Case "BA400"
'                Call Phase_TCP_BA400
'        Case "OSMOPRO"
'                Call Phase_TCP_OSMOPRO
'        Case ""
'
'    End Select
'
'End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڿ��� CheckSum�� ����
'   �μ� :
'       - pMsg : ���ڿ�
'   ��ȯ : CheckSum
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

Public Function CheckSum_ADVIA2120(ByVal pMsg As String) As String
    Dim i%
    Dim sXOR$
    
    sXOR = ""
    sXOR = Mid(pMsg, 1, 1)
    
    For i = 2 To Len(pMsg)
        sXOR = Chr(Int(Asc(sXOR)) Xor Int(Asc(Mid(pMsg, i, 1))))
    Next
    
    If sXOR = Chr(3) Then
        sXOR = Chr(127)
    End If
    
    CheckSum_ADVIA2120 = Chr(2) & pMsg & sXOR & Chr(3)
    
End Function


'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڵ��ȣ�� ���� 1. �������� ��ȸ,
'                                 2. ���������� ȭ��ǥ��,
'                                 3. ó���ڵ� ��������
'   �μ� :
'       - pBarNo : ���ڵ��ȣ
'       - pType  : ���ڵ� �̻��� ���ϴ� ���
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : üũ�Ȱ��� ���� ���� ��
'-----------------------------------------------------------------------------'
Public Sub SetPatInfo(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    With frmMain
        '-- ���ڵ� ���
        If .optBarSeq(0).Value = True Then
            For i = 1 To .spdOrder.DataRowCnt
                If IsNumeric(pBarno) And IsNumeric(Trim(GetText(frmMain.spdOrder, i, colBARCODE))) Then
                    If Val(Trim(GetText(frmMain.spdOrder, i, colBARCODE))) = Val(pBarno) Then
                        If Trim(GetText(frmMain.spdOrder, i, colSTATE)) = "" Or InStr(GetText(frmMain.spdOrder, i, colSTATE), "����") > 0 Then
                            intRow = i
                            Exit For
                        End If
                    End If
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
                        If GetText(frmMain.spdOrder, i, colCHECKBOX) = "1" And GetText(frmMain.spdOrder, i, colSTATE) = "" Then
                            pBarno = Trim(GetText(frmMain.spdOrder, i, colBARCODE))
                            mResult.BarNo = pBarno
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
    
        Call SetText(.spdOrder, "1", intRow, colCHECKBOX)
        
        '-- ������ε��� ȭ��ǥ��
        Call SetText(.spdOrder, mResult.RsltSeq, intRow, colSAVESEQ)
        Call SetText(.spdOrder, mResult.RsltDate, intRow, colEXAMDATE)
        
        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mResult.BarNo, intRow, colBARCODE)
        Call SetText(.spdOrder, mResult.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mResult.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mResult.TubePos, intRow, colPOSNO)
    
        '-- ȯ������ ǥ��
        'Call vasActiveCell(.spdOrder, intRow, colBARCODE)
        
        '-- ����������� �����
        .spdResult.MaxRows = 0
    
        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
    
    End With
    
    '-- ���� Row
    gRow = intRow
    
End Sub

'-----------------------------------------------------------------------------'
'   ��� : �ش� ���ڵ��ȣ�� ���� 1. �������� ��ȸ,
'                                 2. ���������� ȭ��ǥ��,
'                                 3. ó���ڵ� ��������
'   �μ� :
'       - pBarNo : ���ڵ��ȣ
'       - pType  : ���ڵ� �̻��� ���ϴ� ���
'                   1 : Seq
'                   2 : Rack/Pos
'                   3 : üũ�Ȱ��� ���� ���� ��
'-----------------------------------------------------------------------------'
Public Sub SetPatInfo_H7080(ByVal pBarno As String, ByVal pType As String)

    Dim i           As Integer
    Dim intRow      As Long
    Dim strItems    As String
    Dim strOrder    As String
    Dim strDate     As String
    Dim strInNum    As String
    Dim strGumNum   As String
    
    intRow = -1
    With frmMain
        For i = 1 To .spdOrder.DataRowCnt
            If IsNumeric(pBarno) And IsNumeric(Trim(GetText(frmMain.spdOrder, i, colCHARTNO))) Then
                If Val(Trim(GetText(frmMain.spdOrder, i, colCHARTNO))) = Val(pBarno) Then
                    intRow = i
                    Exit For
                End If
            End If
        Next i
        
        '-- �������忡�� ��ã����..
        If intRow < 0 Then
            intRow = .spdOrder.DataRowCnt + 1
            If .spdOrder.MaxRows < intRow Then
                .spdOrder.MaxRows = intRow
            End If
        End If
    
        Call SetText(.spdOrder, "1", intRow, colCHECKBOX)
        
        '-- ������ε��� ȭ��ǥ��
        Call SetText(.spdOrder, mResult.RsltSeq, intRow, colSAVESEQ)
        Call SetText(.spdOrder, mResult.RsltDate, intRow, colEXAMDATE)
        
        '-- ���������� ȭ��ǥ��
        Call SetText(.spdOrder, mResult.BarNo, intRow, colBARCODE)
        'Call SetText(.spdOrder, mResult.Seq, intRow, colSEQNO)
        Call SetText(.spdOrder, mResult.RackNo, intRow, colRACKNO)
        Call SetText(.spdOrder, mResult.TubePos, intRow, colPOSNO)
    
        '-- ȯ������ ǥ��
        'Call vasActiveCell(.spdOrder, intRow, colBARCODE)
        
        '-- ����������� �����
        .spdResult.MaxRows = 0
    
        '-- �˻��� ���� ��������
        Call GetSampleInfo(intRow, .spdOrder)
        
        .spdOrder.RowHeight(-1) = 12
    
    End With
    
    '-- ���� Row
    gRow = intRow
    
End Sub


'-----------------------------------------------------------------------------'
'   ��� : �������� ����
'-----------------------------------------------------------------------------'
Public Sub SendOrder_E411()
    
    
End Sub

'
'Public Sub Phase_TCP_BA400()
'
'End Sub
'
'Public Sub Phase_TCP_OSMOPRO()
'
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
'                        Erase strRecvData
'                        intPhase = 2
'                        'frmMain.comEqp.Output = ACK
'                        frmMain.wSck.SendData ACK
'                        SetRawData "[Tx]" & ACK
'
'                    Case ACK
'                        If strState = "Q" Then
'                            'Call SendOrder_VERSACELL
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
'                            Erase strRecvData
'                            ReDim Preserve strRecvData(intBufCnt)
'                        Else
'                            intBufCnt = intBufCnt + 1
'                            ReDim Preserve strRecvData(intBufCnt)
'                        End If
'                    Case ETB
'                        blnIsETB = True
'                        intPhase = 3
'                    Case ETX
'                        intBufCnt = intBufCnt + 1
'                        ReDim Preserve strRecvData(intBufCnt)
'                        intPhase = 3
'                    Case vbCr
'                        intBufCnt = intBufCnt + 1
'                        ReDim Preserve strRecvData(intBufCnt)
'                    Case EOT
'                        intPhase = 1
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
'                        intPhase = 4
'                        'frmMain.comEqp.Output = ACK
'                        frmMain.wSck.SendData ACK
'                        SetRawData "[Tx]" & ACK
'                End Select
'            Case 4      '## Termination Phase
'                Select Case BufChar
'                    Case STX
'                        intPhase = 2
'                    Case EOT
'                        Call frmMain.SerialRcvData_OSMOPRO
''                        If strState = "Q" Then
''                            intSndPhase = 1
''                            intFrameNo = 1
''                            frmMain.comEqp.Output = ENQ
''                            SetRawData "[Tx]" & ENQ
''                        End If
'                        intPhase = 1
'                End Select
'        End Select
'    Next i
'
'End Sub
    

'Public Sub Phase_Serial_E411()
'
'
'End Sub
'
'Public Sub Phase_Serial_AU400()
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
'
'End Sub

'Public Sub Phase_Serial_AU680()
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
'                Call SerialRcvData_AU680
'            Case Else
'                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'        End Select
'    Next i
'
'End Sub

'Public Sub Phase_Serial_AFIAS6()
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
'            Case "$" 'SOH
'                intBufCnt = 1
'                Erase strRecvData
'                ReDim Preserve strRecvData(intBufCnt)
'            Case vbCr
'                Call SerialRcvData_AFIAS6
'            Case Else
'                strRecvData(intBufCnt) = strRecvData(intBufCnt) & BufChar
'        End Select
'    Next i
'
'End Sub


'Public Sub Phase_Serial_AU480()
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
'
'End Sub

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


'Public Sub SerialRcvData_XN1000()
'    Dim RS_L            As ADODB.Recordset
'    Dim strRcvBuf       As String   '������ Data
'    Dim strRcvBufOrd    As String
'    Dim strRcvBufRst    As String
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
'            strType = Mid$(strRcvBuf, 1, 1)
'            If IsNumeric(strType) Then
'                strType = Mid$(strRcvBuf, 2, 1)
'            End If
'
'            Select Case strType
'                Case "H"
'                Case "Q"    '## Inquiry Order
'                        strBarno = Trim(Mid(strRcvBuf, 14, 26))
'                        strSeq = Mid(strRcvBuf, 9, 5)
'                        strRackNo = Mid(strRcvBuf, 3, 4)
'                        strTubePos = Mid(strRcvBuf, 7, 2)
'
'                        With mOrder
'                            .BarNo = strBarno
'                            .Seq = strSeq
'                            .RackNo = strRackNo
'                            .TubePos = strTubePos
'                        End With
'
'                        If strBarno = "" Then
'                            strBarno = "NoOrder_" & Trim(strSeq)
'                        End If
'
'                        Call GetOrder(strBarno, gHOSP.RSTTYPE)
'
'                        strState = "Q"
'
'                Case "P"
'
'                Case "O"
'                    '4O|1||1^6^          201404240002^B|^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^RDW-SD\^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT\^^^^NEUT#\^^^^LYMPH#\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^NEUT%\^^^^LYMPH%\^^^^MONO%\^^^^EC|1||
'
'                    strRcvBufOrd = Trim$(mGetP(strRcvBuf, 4, "|"))
'                    strBarno = Trim$(mGetP(strRcvBufOrd, 3, "^"))
'                    strSeq = ""
'                    strRackNo = Trim$(mGetP(strRcvBufOrd, 1, "^"))
'                    strTubePos = Trim$(mGetP(strRcvBufOrd, 2, "^"))
'
'                    With mResult
'                        .BarNo = strBarno
'                        .SpcPos = strSeq
'                        .RackNo = strRackNo
'                        .TubePos = strTubePos
'                        .RsltDate = Format(Now, "yyyymmddhhmmss")
'                        .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'                    End With
'
'                    Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'                    If gRow <= 0 Then
'                        Exit Sub
'                    End If
'
'                    strState = "O"
'                    .spdResult.MaxRows = 0
'
'                Case "R"
'                    strRcvBufRst = Trim(mGetP(strRcvBuf, 3, "|"))
'                    strIntBase = Trim$(mGetP(strRcvBufRst, 5, "^"))
'                    strResult = Trim(mGetP(strRcvBuf, 4, "|"))
'
'                    If strIntBase <> "" And strResult <> "" Then
'                        If gPatOrdCd <> "" Then
'                            SQL = ""
'                            SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                            SQL = SQL & "  FROM EQPMASTER" & vbCr
'                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.HOSPCD & "' " & vbCr
'                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
'                            SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
'
'                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                lsTestCode = Trim(RS_L.Fields("TESTCODE"))
'                                lsTestName = Trim(RS_L.Fields("TESTNAME"))
'                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                '-- ���Row �߰�
'                                lsRstRow = .spdResult.DataRowCnt + 1
'                                If .spdResult.MaxRows < lsRstRow Then
'                                    .spdResult.MaxRows = lsRstRow
'                                End If
'
'                                '�Ҽ��� ó��, ��� ���� ó��
'                                strMachResult = strResult
'                                strResult = SetResult(strResult, strIntBase)
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
'                            SQL = SQL & "  FROM EQPMASTER" & vbCr
'                            SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.HOSPCD & "' " & vbCr
'                            SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
'
'                            Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                            If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                                lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                                lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                '-- ���Row �߰�
'                                lsRstRow = .spdResult.DataRowCnt + 1
'                                If .spdResult.MaxRows < lsRstRow Then
'                                    .spdResult.MaxRows = lsRstRow
'                                End If
'
'                                '�Ҽ��� ó��, ��� ���� ó��
'                                strMachResult = strResult
'                                strResult = SetResult(strResult, strIntBase)
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
'                Case "L"
'                    '## DB�� �������
'                    If .optTrans(0).Value = True And strState = "R" Then
'                        Res = SaveTransData(gRow)
'
'                        If Res = -1 Then
'                            '-- ���� ����
''                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                            SetText .spdOrder, "Failed", gRow, colSTATE
'                        Else
'                            '-- ���� ����
''                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
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
''                            Res = SendQuery(gLocal, SQL)
'                            If Res = -1 Then
''                                SaveQuery SQL
'                                Exit Sub
'                            End If
'                        End If
'                        strState = ""
'                    End If
'
'            End Select
'        Next
'    End With
'
'End Sub


'Public Sub SerialRcvData_AU480()
'
'
'End Sub

'Public Sub SerialRcvData_AU400()
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
'            strRcvBuf = strRecvData(intCnt)
'            strType = Mid$(strRcvBuf, 1, 2)
'
'            Select Case strType
'                Case "R "    '## Inquiry Order
'                        'R 000101 0001                1608270009
'                        strBarno = Trim(Mid(strRcvBuf, 14, 26))
'                        strSeq = Mid(strRcvBuf, 9, 5)
'                        strRackNo = Mid(strRcvBuf, 3, 4)
'                        strTubePos = Mid(strRcvBuf, 7, 2)
'
'                        With mOrder
'                            .BarNo = strBarno
'                            .Seq = strSeq
'                            .RackNo = strRackNo
'                            .TubePos = strTubePos
'                        End With
'
'                        If strBarno = "" Then
'                            strBarno = "NoOrder_" & Trim(strSeq)
'                        End If
'
'                        Call GetOrder(strBarno, gHOSP.RSTTYPE)
'
'                        strState = "Q"
'
'                Case "D "    '## Result
'                        'D 000101 0001                1608270009    E001   9.3  002   5.8  
'                        strBarno = Trim$(Mid$(strRcvBuf, 14, 10))
'                        strRackNo = Mid(strRcvBuf, 3, 4)
'                        strTubePos = Mid(strRcvBuf, 7, 2)
'                        strSeq = Trim(Mid(strRcvBuf, 9, 5))
'
'                        With mResult
'                            .BarNo = strBarno
'                            .SpcPos = strSeq
'                            .Seq = strSeq
'                            .RackNo = strRackNo
'                            .TubePos = strTubePos
'                            .RsltDate = Format(Now, "yyyymmddhhmmss")
'                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'                        End With
'
'                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'                        If gRow <= 0 Then
'                            Exit Sub
'                        End If
'
'                        strTmp = Mid$(strRcvBuf, 29)
'
'                        Do While Len(strTmp) >= 11
'                            strIntBase = Mid$(strTmp, 2, 2)
'                            strResult = Mid$(strTmp, 4, 6)
'                            strComm = Mid$(strTmp, 10, 1)
'
'                            If strIntBase <> "" And strResult <> "" Then
'                                If gPatOrdCd <> "" Then
'                                    SQL = ""
'                                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                                    SQL = SQL & "  FROM EQPMASTER" & vbCr
'                                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
'                                    SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
'
'                                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                        lsTestCode = Trim(RS_L.Fields("TESTCODE"))
'                                        lsTestName = Trim(RS_L.Fields("TESTNAME"))
'                                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                        '-- ���Row �߰�
'                                        lsRstRow = .spdResult.DataRowCnt + 1
'                                        If .spdResult.MaxRows < lsRstRow Then
'                                            .spdResult.MaxRows = lsRstRow
'                                        End If
'
'                                        '�Ҽ��� ó��, ��� ���� ó��
'                                        strMachResult = strResult
'                                        strResult = SetResult(strResult, strIntBase)
'                                        strJudge = SetJudge(strResult, strIntBase)
'
'                                        '������� ǥ��("���")
'                                        SetText .spdOrder, "���", gRow, colSTATE
'
'                                        '����� ǥ��
'                                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                                SetText .spdOrder, strResult, gRow, intCol
'                                                Exit For
'                                            End If
'                                        Next
'
'                                        '-- ��� List
'                                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '����
'                                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
'                                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
'                                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
'                                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '�˻��
'                                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '���ä��
'                                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '�����
'                                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS���
'                                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '����
'                                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
'
'                                        '-- ���� ����
'                                        SetLocalDB gRow, lsRstRow, "1", ""
'
'                                        strState = "R"
'
'                                        '-- ���Count
'                                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                                            SetText .spdOrder, "1", gRow, colRCNT
'                                        Else
'                                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                                        End If
'
'                                    End If
'                                Else
'                                    SQL = ""
'                                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                                    SQL = SQL & "  FROM EQPMASTER" & vbCr
'                                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
'
'                                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                        lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                                        lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                        '-- ���Row �߰�
'                                        lsRstRow = .spdResult.DataRowCnt + 1
'                                        If .spdResult.MaxRows < lsRstRow Then
'                                            .spdResult.MaxRows = lsRstRow
'                                        End If
'
'                                        '�Ҽ��� ó��, ��� ���� ó��
'                                        strMachResult = strResult
'                                        strResult = SetResult(strResult, strIntBase)
'                                        strJudge = SetJudge(strResult, strIntBase)
'
'                                        '������� ǥ��("���")
'                                        SetText .spdOrder, "���", gRow, colSTATE
'
'                                        '����� ǥ��
'                                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                                SetText .spdOrder, strResult, gRow, intCol
'                                                Exit For
'                                            End If
'                                        Next
'
'                                        '-- ��� List
'                                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '����
'                                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
'                                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
'                                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '�˻��
'                                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '���ä��
'                                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '�����
'                                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS���
'                                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '����
'                                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
'
'                                        '-- ���� ����
'                                        SetLocalDB gRow, lsRstRow, "1", ""
'
'                                        If strState <> "R" Then
'                                            strState = ""
'                                        End If
'
'                                        '-- ���Count
'                                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                                            SetText .spdOrder, "1", gRow, colRCNT
'                                        Else
'                                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                                        End If
'                                    End If
'
'                                End If
'
'                            End If
'                            strTmp = Mid$(strTmp, 12)
'                        Loop
'
'                        .spdResult.RowHeight(-1) = 14
'
'                        '## DB�� �������
'                        If .optTrans(0).Value = True And strState = "R" Then
'                            Res = SaveTransData(gRow)
'
'                            If Res = -1 Then
'                                '-- ���� ����
'    '                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                                SetText .spdOrder, "Failed", gRow, colSTATE
'                            Else
'                                '-- ���� ����
'    '                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                                SetText .spdOrder, "����Ϸ�", gRow, colSTATE
'                                SetText .spdOrder, "0", gRow, colCHECKBOX
'
'                                      SQL = "Update PATRESULT Set " & vbCrLf
'                                SQL = SQL & " sendflag = '2' " & vbCrLf
'                                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
'                                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'                                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'                                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
'
'    '                            Res = SendQuery(gLocal, SQL)
'                                If Res = -1 Then
'    '                                SaveQuery SQL
'                                    Exit Sub
'                                End If
'                            End If
'                            strState = ""
'                        End If
'
'            End Select
'        Next
'    End With
'
'End Sub
'
'Public Sub SerialRcvData_AU680()
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
'            strRcvBuf = strRecvData(intCnt)
'            strType = Mid$(strRcvBuf, 1, 2)
'
'            Select Case strType
'                Case "R "    '## Inquiry Order
'                        'R 000101 0001                1608270009
'                        strBarno = Trim(Mid(strRcvBuf, 14, 20))
'                        strSeq = Mid(strRcvBuf, 9, 5)
'                        strRackNo = Mid(strRcvBuf, 3, 4)
'                        strTubePos = Mid(strRcvBuf, 7, 2)
'
'                        With mOrder
'                            .BarNo = strBarno
'                            .Seq = strSeq
'                            .RackNo = strRackNo
'                            .TubePos = strTubePos
'                        End With
'
'                        If strBarno = "" Then
'                            strBarno = "NoOrder_" & Trim(strSeq)
'                        End If
'
'                        Call GetOrder(strBarno, gHOSP.RSTTYPE)
'
'                        strState = "Q"
'
'                Case "D "    '## Result
'                        '1234567890123456789012345678901234567890
'                        'D 000501 0001            02001035    E001  2.60Pr002  75.1Pr003    98r 004   7.1r 005   3.4Nr007    80Pr008    75Pr009    61r 011  0.68r 097   132Nr098   4.3r 099   101r
'                        'D 000101 0001                1608270009    E001   9.3  002   5.8  
'                        strBarno = Trim$(Mid$(strRcvBuf, 14, 20))
'                        strRackNo = Mid(strRcvBuf, 3, 4)
'                        strTubePos = Mid(strRcvBuf, 7, 2)
'                        strSeq = Trim(Mid(strRcvBuf, 10, 4))
'
'                        With mResult
'                            .BarNo = strBarno
'                            .SpcPos = strSeq
'                            .Seq = strSeq
'                            .RackNo = strRackNo
'                            .TubePos = strTubePos
'                            .RsltDate = Format(Now, "yyyymmddhhmmss")
'                            .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'                        End With
'
'                        Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'                        If gRow <= 0 Then
'                            Exit Sub
'                        End If
'
'                        'strTmp = Mid$(strRcvBuf, 29)
'                        strTmp = Mid$(strRcvBuf, 39)
'
'                        'Do While Len(strTmp) >= 11
'                        Do While Len(strTmp) >= 10
'                            strIntBase = Mid$(strTmp, 2, 2)
'                            strResult = Mid$(strTmp, 4, 6)
'                            strComm = Mid$(strTmp, 10, 1)
'
'                            If strIntBase <> "" And strResult <> "" Then
'                                If gPatOrdCd <> "" Then
'                                    SQL = ""
'                                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                                    SQL = SQL & "  FROM EQPMASTER" & vbCr
'                                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
'                                    SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
'
'                                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                        lsTestCode = Trim(RS_L.Fields("TESTCODE"))
'                                        lsTestName = Trim(RS_L.Fields("TESTNAME"))
'                                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                        '-- ���Row �߰�
'                                        lsRstRow = .spdResult.DataRowCnt + 1
'                                        If .spdResult.MaxRows < lsRstRow Then
'                                            .spdResult.MaxRows = lsRstRow
'                                        End If
'
'                                        '�Ҽ��� ó��, ��� ���� ó��
'                                        strMachResult = strResult
'                                        strResult = SetResult(strResult, strIntBase)
'                                        strJudge = SetJudge(strResult, strIntBase)
'
'                                        '������� ǥ��("���")
'                                        SetText .spdOrder, "���", gRow, colSTATE
'
'                                        '����� ǥ��
'                                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                                SetText .spdOrder, strResult, gRow, intCol
'                                                Exit For
'                                            End If
'                                        Next
'
'                                        '-- ��� List
'                                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '����
'                                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
'                                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
'                                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '�˻��
'                                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '���ä��
'                                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '�����
'                                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS���
'                                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '����
'                                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
'
'                                        '-- ���� ����
'                                        SetLocalDB gRow, lsRstRow, "1", ""
'
'                                        strState = "R"
'
'                                        '-- ���Count
'                                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                                            SetText .spdOrder, "1", gRow, colRCNT
'                                        Else
'                                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                                        End If
'
'                                    End If
'                                Else
'                                    SQL = ""
'                                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                                    SQL = SQL & "  FROM EQPMASTER" & vbCr
'                                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
'
'                                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                                        lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                                        lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                                        '-- ���Row �߰�
'                                        lsRstRow = .spdResult.DataRowCnt + 1
'                                        If .spdResult.MaxRows < lsRstRow Then
'                                            .spdResult.MaxRows = lsRstRow
'                                        End If
'
'                                        '�Ҽ��� ó��, ��� ���� ó��
'                                        strMachResult = strResult
'                                        strResult = SetResult(strResult, strIntBase)
'                                        strJudge = SetJudge(strResult, strIntBase)
'
'                                        '������� ǥ��("���")
'                                        SetText .spdOrder, "���", gRow, colSTATE
'
'                                        '����� ǥ��
'                                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                                SetText .spdOrder, strResult, gRow, intCol
'                                                Exit For
'                                            End If
'                                        Next
'
'                                        '-- ��� List
'                                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '����
'                                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
'                                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
'                                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '�˻��
'                                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '���ä��
'                                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '�����
'                                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS���
'                                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '����
'                                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
'
'                                        '-- ���� ����
'                                        SetLocalDB gRow, lsRstRow, "1", ""
'
'                                        If strState <> "R" Then
'                                            strState = ""
'                                        End If
'
'                                        '-- ���Count
'                                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                                            SetText .spdOrder, "1", gRow, colRCNT
'                                        Else
'                                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                                        End If
'                                    End If
'
'                                End If
'
'                            End If
'                            strTmp = Mid$(strTmp, 12)
'                        Loop
'
'                        .spdResult.RowHeight(-1) = 14
'
'                        '## DB�� �������
'                        If .optTrans(0).Value = True And strState = "R" Then
'                            Res = SaveTransData(gRow)
'
'                            If Res = -1 Then
'                                '-- ���� ����
'    '                            SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                                SetText .spdOrder, "Failed", gRow, colSTATE
'                            Else
'                                '-- ���� ����
'    '                            SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                                SetText .spdOrder, "����Ϸ�", gRow, colSTATE
'                                SetText .spdOrder, "0", gRow, colCHECKBOX
'
'                                      SQL = "Update PATRESULT Set " & vbCrLf
'                                SQL = SQL & " sendflag = '2' " & vbCrLf
'                                SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
'                                SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'                                SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'                                SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
'
'    '                            Res = SendQuery(gLocal, SQL)
'                                If Res = -1 Then
'    '                                SaveQuery SQL
'                                    Exit Sub
'                                End If
'                            End If
'                            strState = ""
'                        End If
'
'            End Select
'        Next
'    End With
'
'End Sub

'Public Sub SerialRcvData_AFIAS6()
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
'            strRcvBuf = strRecvData(intCnt)
'            strBarno = Trim(mGetP(strRcvBuf, 5, "|"))
'            strRackNo = ""
'            strTubePos = ""
'            strSeq = ""
'
'            With mResult
'                .BarNo = strBarno
'                .SpcPos = strSeq
'                .Seq = strSeq
'                .RackNo = strRackNo
'                .TubePos = strTubePos
'                .RsltDate = Format(Now, "yyyymmddhhmmss")
'                .RsltSeq = getMaxTestNum(Format(frmMain.dtpToday, "yyyymmdd"))
'            End With
'
'            Call SetPatInfo(strBarno, gHOSP.RSTTYPE)
'
'            If gRow <= 0 Then
'                Exit Sub
'            End If
'
'            strIntBase = mGetP(strRcvBuf, 8, "|")
'            strResult = mGetP(strRcvBuf, 11, "|")
'
'            If strIntBase <> "" And strResult <> "" Then
'                If gPatOrdCd <> "" Then
'                    SQL = ""
'                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                    SQL = SQL & "  FROM EQPMASTER" & vbCr
'                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' " & vbCr
'                    SQL = SQL & "   AND TESTCODE in (" & gPatOrdCd & ") "
'
'                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                        lsTestCode = Trim(RS_L.Fields("TESTCODE"))
'                        lsTestName = Trim(RS_L.Fields("TESTNAME"))
'                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                        '-- ���Row �߰�
'                        lsRstRow = .spdResult.DataRowCnt + 1
'                        If .spdResult.MaxRows < lsRstRow Then
'                            .spdResult.MaxRows = lsRstRow
'                        End If
'
'                        '�Ҽ��� ó��, ��� ���� ó��
'                        strMachResult = strResult
'                        strResult = SetResult(strResult, strIntBase)
'                        strJudge = SetJudge(strResult, strIntBase)
'
'                        '������� ǥ��("���")
'                        SetText .spdOrder, "���", gRow, colSTATE
'
'                        '����� ǥ��
'                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                SetText .spdOrder, strResult, gRow, intCol
'                                Exit For
'                            End If
'                        Next
'
'                        '-- ��� List
'                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '����
'                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
'                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
'                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '�˻��
'                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '���ä��
'                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '�����
'                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS���
'                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '����
'                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
'
'                        '-- ���� ����
'                        SetLocalDB gRow, lsRstRow, "1", ""
'
'                        strState = "R"
'
'                        '-- ���Count
'                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                            SetText .spdOrder, "1", gRow, colRCNT
'                        Else
'                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                        End If
'
'                    End If
'                Else
'                    SQL = ""
'                    SQL = SQL & "SELECT TESTCODE,TESTNAME,SEQNO,REFLOW,REFHIGH " & vbCr
'                    SQL = SQL & "  FROM EQPMASTER" & vbCr
'                    SQL = SQL & " WHERE EQUIPCD = '" & gHOSP.MACHCD & "' " & vbCr
'                    SQL = SQL & "   AND RSLTCHANNEL = '" & strIntBase & "' "
'
'                    Set RS_L = AdoCn_Local.Execute(SQL, , 1)
'                    If Not RS_L.EOF = True And Not RS_L.BOF = True Then
'                        lsTestCode = Trim(RS_L.Fields("TESTCODE") & "")
'                        lsTestName = Trim(RS_L.Fields("TESTNAME") & "")
'                        lsSeqNo = Trim(RS_L.Fields("SEQNO"))
'
'                        '-- ���Row �߰�
'                        lsRstRow = .spdResult.DataRowCnt + 1
'                        If .spdResult.MaxRows < lsRstRow Then
'                            .spdResult.MaxRows = lsRstRow
'                        End If
'
'                        '�Ҽ��� ó��, ��� ���� ó��
'                        strMachResult = strResult
'                        strResult = SetResult(strResult, strIntBase)
'                        strJudge = SetJudge(strResult, strIntBase)
'
'                        '������� ǥ��("���")
'                        SetText .spdOrder, "���", gRow, colSTATE
'
'                        '����� ǥ��
'                        For intCol = colSTATE + 1 To .spdOrder.MaxCols
'                            If lsTestCode = gArrEQP(intCol - colSTATE, 2) Then
'                                SetText .spdOrder, strResult, gRow, intCol
'                                Exit For
'                            End If
'                        Next
'
'                        '-- ��� List
'                        SetText .spdResult, lsSeqNo, lsRstRow, colRSEQNO                '����
'                        SetText .spdResult, lsOrderCode, lsRstRow, colRORDERCD          'ó���ڵ�
'                        SetText .spdResult, lsTestCode, lsRstRow, colRTESTCD            '�˻��ڵ�
'                        SetText .spdResult, lsTestName, lsRstRow, colRTESTNM            '�˻��
'                        SetText .spdResult, strIntBase, lsRstRow, colRCHANNEL           '���ä��
'                        SetText .spdResult, strMachResult, lsRstRow, colRMACHRESULT     '�����
'                        SetText .spdResult, strResult, lsRstRow, colRLISRESULT          'LIS���
'                        SetText .spdResult, strJudge, lsRstRow, colRJUDGE                     '����
'                        SetText .spdResult, Trim(RS_L.Fields("REFLOW")) & "~" & Trim(RS_L.Fields("REFHIGH")), lsRstRow, colRREF          '����ġ
'
'                        '-- ���� ����
'                        SetLocalDB gRow, lsRstRow, "1", ""
'
'                        If strState <> "R" Then
'                            strState = ""
'                        End If
'
'                        '-- ���Count
'                        If GetText(.spdOrder, gRow, colRCNT) = "" Then
'                            SetText .spdOrder, "1", gRow, colRCNT
'                        Else
'                            SetText .spdOrder, GetText(.spdOrder, gRow, colRCNT) + 1, gRow, colRCNT
'                        End If
'                    End If
'
'                End If
'
'            End If
'
'            .spdResult.RowHeight(-1) = 14
'
'            '## DB�� �������
'            If .optTrans(0).Value = True And strState = "R" Then
'                Res = SaveTransData_MCC(gRow)
'
'                If Res = -1 Then
'                    '-- ���� ����
'                    SetForeColor .spdOrder, gRow, gRow, 1, colSTATE, 255, 0, 0
'                    SetText .spdOrder, "Failed", gRow, colSTATE
'                Else
'                    '-- ���� ����
'                    SetBackColor .spdOrder, gRow, gRow, 1, colSTATE, 202, 255, 112
'                    SetText .spdOrder, "����Ϸ�", gRow, colSTATE
'                    SetText .spdOrder, "0", gRow, colCHECKBOX
'
'                          SQL = "Update PATRESULT Set " & vbCrLf
'                    SQL = SQL & " sendflag = '2' " & vbCrLf
'                    SQL = SQL & " Where equipno = '" & gHOSP.HOSPCD & "' " & vbCrLf
'                    SQL = SQL & "   And examdate = '" & Mid(Trim(GetText(.spdOrder, gRow, colEXAMDATE)), 1, 8) & "' " & vbCrLf
'                    SQL = SQL & "   And barcode = '" & Trim(GetText(.spdOrder, gRow, colBARCODE)) & "' " & vbCrLf
'                    SQL = SQL & "   And saveseq = " & Trim(GetText(.spdOrder, gRow, colSAVESEQ)) & vbCrLf
'
'                    If DBExec(AdoCn_Local, SQL) Then
'                        '-- ����
'                    End If
'                End If
'                strState = ""
'            End If
'        Next
'    End With
'
'End Sub


Public Sub SerialRcvData_E411()
   

End Sub


Function SaveTransData(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
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
        
On Error GoTo ErrHandle

    With frmMain
        SaveTransData = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        
        
        '-- Local���� ȯ�ں��� ����� ��������
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
        
        Call SetSQLData("���ð����ȸ", SQL)
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        'AdoCn.BeginTrans
        
        '-- ������ ����� �����ϱ�
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '�˻��ڵ�
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '���(�����)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '���(�������)
                        
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                      SQL = " Update SLA_LabResult  " & vbCr
                SQL = SQL & "   Set Result     = '" & sResult & "' " & vbCr
                SQL = SQL & "      ,NormalFlag = '0' " & vbCr
                SQL = SQL & "      ,PanicFlag  = '0' " & vbCr
                SQL = SQL & "      ,DeltaFlag  = '0' " & vbCr
                SQL = SQL & "      ,TransFlag  = '1' " & vbCr
                SQL = SQL & "      ,ResultID   = ''  " & vbCr
                SQL = SQL & "      ,ResultDate = '" & Trim(Format(Now, "yyyy-mm-dd")) & "'" & vbCr
                SQL = SQL & "      ,ResultTime = '" & Trim(Format(Time, "HH:MM:SS")) & "'" & vbCr
                SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "'" & vbCr
                SQL = SQL & "   And OrderCode IN (" & gAllOrdCd & ") " & vbCr
                SQL = SQL & "   And LabCode = '" & strTestCd & "'" & vbCr
                SQL = SQL & "   And TransFlag < '2' "

                Call SetSQLData("�������", SQL)
                Call DBExec(AdoCn, SQL)
                
                SaveTransData = 1
                
            End If
        Next intRow
        
        If SaveTransData = 1 Then
                  SQL = " Update SLA_LabMaster " & vbCr
            SQL = SQL & "   Set JStatus = '2' " & vbCr
            SQL = SQL & " Where SPECIMENNUM = '" & strBarcode & "' " & vbCr
            'SQL = SQL & "   And OrderCode = '" & strTestCd & "'" & vbCr
            SQL = SQL & "   And OrderCode IN (" & gAllOrdCd & ") " & vbCr
            SQL = SQL & "   And RECEIPTDATE = '" & Format(strHospDate, "yyyy-mm-dd") & "'" & vbCr
            SQL = SQL & "   And JStatus < '3' "
            
            Call SetSQLData("��������", SQL)
            Call DBExec(AdoCn, SQL)
            
        End If
        
        'AdoCn.CommitTrans
        
    
    End With

Exit Function

ErrHandle:
    SaveTransData = -1
    'AdoCn.RollbackTrans
    
End Function

Function SaveTransData_MCC(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
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
    
    'Dim strReturn       As String
    Dim intReturn       As Long
    Dim strMSG          As String
    
    Dim prm0 As New ADODB.Parameter
    Dim prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim prm3 As New ADODB.Parameter
    Dim prm4 As New ADODB.Parameter
    Dim prm5 As New ADODB.Parameter
    
    
    Dim intBarno  As Double
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_MCC = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        
        If IsNumeric(strBarcode) Then
            intBarno = CDbl(strBarcode)
        Else
            Exit Function
        End If
        
        '-- Local���� ȯ�ں��� ����� ��������
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
        
        Call SetSQLData("���ð����ȸ", SQL)
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- ������ ����� �����ϱ�
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '�˻��ڵ�
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '���(�����)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '���(�������)
                        
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                      SQL = "Exec UP_LIS_INTERFACE_U$001 " & intBarno
                SQL = SQL & "," & strTestCd
                SQL = SQL & "," & sResult
                SQL = SQL & "," & gHOSP.MACHCD

                'AdoCn.Execute SQL
                
                Set AdoCmd = New ADODB.Command
                Set AdoCmd.ActiveConnection = AdoCn
                With AdoCmd
                    .CommandTimeout = 15
                    .CommandText = "UP_LIS_INTERFACE_U$001"
                    .CommandType = adCmdStoredProc
                    
                    
                    Set prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, intBarno)      '���ڵ��ȣ
                    .Parameters.Append prm1

                    Set prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       'ó���ڵ�
                    .Parameters.Append prm2

                    Set prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '�����
                    .Parameters.Append prm3

                    Set prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, gHOSP.MACHCD)    '����ڵ�
                    .Parameters.Append prm4

                    .Execute
                    
                End With
                
                'Call SetSQLData("�������", SQL)
                
                SaveTransData_MCC = 1
                
            End If
        Next intRow
        
    End With

Exit Function

ErrHandle:
    SaveTransData_MCC = -1
    
End Function

Function SaveTransData_GINUS(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_GINUS = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        
        '-- Local���� ȯ�ں��� ����� ��������
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
        
        Call SetSQLData("���ð����ȸ", SQL)
        
        Set RS_L = New ADODB.Recordset
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- ������ ����� �����ϱ�
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '�˻��ڵ�
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '���(�����)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '���(�������)
                        
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            '-- ����� SEQ ã�ƿ���
            strSeqS = GetOrderSeqCode(Mid(strExamDate, 1, 8), strBarcode, strTestCd)
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- �������
                      SQL = "Update scrrslth" & vbCr
                SQL = SQL & " SET exam_stus    = decode(exam_stus, '0','1',exam_stus), " & vbCr
                SQL = SQL & "     exam_rslt    = '" & sResult & "', " & vbCr
                SQL = SQL & "     mach_rslt    = '" & sResult & "', " & vbCr
                SQL = SQL & "     exam_dt      = '" & Format(Now, "yyyymmddhhmm") & "', " & vbCr
                SQL = SQL & "     exam_empno   = '" & gHOSP.USERID & "'" & vbCr
                SQL = SQL & " WHERE hos_org_no = '" & gHOSP.HOSPCD & "'" & vbCr
                SQL = SQL & "   AND smp_no     = '" & strBarcode & "'" & vbCr
                SQL = SQL & "   AND cd         = '" & strTestCd & "'" & vbCr
                SQL = SQL & "   AND prcp_seq   = '" & mGetP(strSeqS, 1, "|") & "'" & vbCr
                SQL = SQL & "   AND exam_seq   = '" & mGetP(strSeqS, 2, "|") & "'" & vbCr
                SQL = SQL & "   AND rept_seq   = '" & mGetP(strSeqS, 3, "|") & "'"
                                  
                Call SetSQLData("�������", SQL, "A")
                AdoCn.Execute SQL
                
                '-- ���º���2
                'mosxpslh ó�� ���̺��� prcp_stus_cd ����    4
                
                      SQL = "UPDATE mosxpslh SET prcp_stus_cd = '4' " & vbCr
                SQL = SQL & " WHERE hos_org_no = '" & gHOSP.HOSPCD & "'" & vbCr
                SQL = SQL & "   AND smp_no     = '" & strBarcode & "'" & vbCr
                SQL = SQL & "   AND prcp_cd    = '" & strTestCd & "'" & vbCr
                
                Call SetSQLData("���º���2", SQL, "A")
                AdoCn.Execute SQL
                
                SaveTransData_GINUS = 1
            
            End If
        Next intRow
        
        
        '-- ���º���1
        'scrprexh �˻� ��� ���̺��� smp_stus ����   6
              
              SQL = "UPDATE scrprexh SET smp_stus = '6' " & vbCr            '���� : I/F = 6
        SQL = SQL & " WHERE hos_org_no = '" & gHOSP.HOSPCD & "'" & vbCr
        SQL = SQL & "   AND smp_no     = '" & strBarcode & "'" & vbCr
        
        Call SetSQLData("���º���1", SQL, "A")
        AdoCn.Execute SQL
        
        SaveTransData_GINUS = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_GINUS = -1
    
End Function

Function SaveTransData_JWINFO(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_JWINFO = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        
        '-- Local���� ȯ�ں��� ����� ��������
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
        
        Call SetSQLData("���ð����ȸ", SQL)
        
        Set RS_L = New ADODB.Recordset
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- ������ ����� �����ϱ�
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '�˻��ڵ�
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '���(�����)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '���(�������)
                        
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- �������
                               SQL = "Update SLA_LabResult  "
                SQL = SQL & vbCrLf & "   Set Result = '" & sResult & "', "
                SQL = SQL & vbCrLf & "       NormalFlag = '0', "
                SQL = SQL & vbCrLf & "       PanicFlag = '0', "
                SQL = SQL & vbCrLf & "       DeltaFlag = '0', "
                SQL = SQL & vbCrLf & "       TransFlag = '1', "
                SQL = SQL & vbCrLf & "       ResultID  = '', "
                SQL = SQL & vbCrLf & "       ResultDate = '" & Trim(Format(Now, "yyyy-mm-dd")) & "', "
                SQL = SQL & vbCrLf & "       ResultTime = '" & Trim(Format(Time, "HH:MM:SS")) & "' "
                SQL = SQL & vbCrLf & " Where SPECIMENNUM = '" & strBarcode & "' "
                'SQL = SQL & vbCrLf & "   AND ReceiptNo = '" & strChartNo & "' "
                'SQL = SQL & vbCrLf & "   AND OrderCode = '" & strTestCd & "'"
                SQL = SQL & vbCrLf & "   AND ORDERCODE IN (" & gAllOrdCd & ") " & vbCr
                SQL = SQL & vbCrLf & "   And LabCode = '" & strTestCd & "'"
                SQL = SQL & vbCrLf & "   And transflag < '2' "
                
                Call SetSQLData("�������", SQL, "A")
                AdoCn.Execute SQL
                
                SaveTransData_JWINFO = 1
            
            End If
        Next intRow
        
        '-- ���º���
                       SQL = "Update SLA_LabMaster "
        SQL = SQL & vbCrLf & "   Set JStatus = '2' "
        SQL = SQL & vbCrLf & " Where SPECIMENNUM = '" & strBarcode & "' "
        SQL = SQL & vbCrLf & "   AND OrderCode IN (" & gAllOrdCd & ") " & vbCr
        SQL = SQL & vbCrLf & "   And JStatus < '3' "
            
        Call SetSQLData("���º���", SQL, "A")
        AdoCn.Execute SQL
        
               
        SaveTransData_JWINFO = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_JWINFO = -1
    
End Function

Function SaveTransData_BIT(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim RS_S            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    Dim strSubCodes     As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_BIT = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        '-- Local���� ȯ�ں��� ����� ��������
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT,EXAMSUBCODE " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
        
        Call SetSQLData("���ð����ȸ", SQL)
        
        Set RS_L = New ADODB.Recordset
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                If RS_L.Fields("EXAMSUBCODE").Value & "" <> "" Then
                    Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 7)
                Else
                    strSubCodes = ""
                    '-- Seq ��������
                    SQL = ""
                    SQL = SQL & " SELECT R.ResOdrSeq, R.ResSeq, R.ResSubSeq " & vbCr
                    SQL = SQL & "   FROM RESINF AS R" & vbCr
                    SQL = SQL & " WHERE ltrim(rtrim(R.RESOCMNUM)) = '" & strBarcode & "'" & vbCr
                    SQL = SQL & "   AND R.RESLABCOD ='" & RS_L.Fields("EXAMCODE").Value & "'" & vbCr
            
                    'Call SetSQLData("SEQ��ȸ", SQL, "A")
                    
                    Set RS_S = New ADODB.Recordset
                    Set RS_S = AdoCn.Execute(SQL, , 1)
                    If Not RS_S.EOF = True And Not RS_S.BOF = True Then
                        Do Until RS_S.EOF
                            strSubCodes = Trim(RS_S.Fields("ResOdrSeq")) & "|" & Trim(RS_S.Fields("ResSeq")) & "|" & Trim(RS_S.Fields("ResSubSeq"))
                            mResult.RESODRSEQ = Trim(RS_S.Fields("ResOdrSeq"))
                            mResult.RESSEQ = Trim(RS_S.Fields("ResSeq"))
                            mResult.RESSUBSEQ = Trim(RS_S.Fields("ResSubSeq"))
                            RS_S.MoveNext
                        Loop
                    End If
                    RS_S.Close
                    
                    Call SetText(.vasTemp, strSubCodes, intRow, 7)
                    
                End If

                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        strSubCodes = ""
        
        '-- ������ ����� �����ϱ�
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '�˻��ڵ�
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '���(�����)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '���(�������)
            strSubCodes = Trim(GetText(.vasTemp, intRow, 7))      '����� �ڵ� : ex) 999|888|777
                        
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" And strSubCodes <> "" Then
                '-- �������
                      SQL = "Update ResInf "
                SQL = SQL & "   Set ResRltVal = '" & sResult & "'," & vbCr   '�˻���
                SQL = SQL & "       ResRepTyp = 'I' " & vbCrLf
                SQL = SQL & " Where ResOcmNum = '" & Space(10 - Len(strPatID)) & strPatID & "'" & vbCr            '10
                SQL = SQL & "   And ResLabCod = '" & strTestCd & "'"    '�˻��ڵ�
                SQL = SQL & "   And ResOdrSeq = '" & mGetP(strSubCodes, 1, "|") & "'" & vbCr
                SQL = SQL & "   And ResSeq    = '" & mGetP(strSubCodes, 2, "|") & "'" & vbCr
                SQL = SQL & "   And ResSubSeq = '" & mGetP(strSubCodes, 3, "|") & "'" & vbCr
                
                Call SetSQLData("�������", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
               
        SaveTransData_BIT = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_BIT = -1
    
End Function

Function SaveTransData_AMIS(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_AMIS = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        
        '-- Local���� ȯ�ں��� ����� ��������
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
        
        Call SetSQLData("���ð����ȸ", SQL)
        
        Set RS_L = New ADODB.Recordset
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- ������ ����� �����ϱ�
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '�˻��ڵ�
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '���(�����)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '���(�������)
                        
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- �������
                      SQL = " Update resultofnum Set"
                SQL = SQL & " resultindate = to_char(sysdate,'yyyymmdd')"
                SQL = SQL & " , resultintime = to_char(sysdate,'HH24MI')"
                SQL = SQL & " , resultinid = '" & gHOSP.USERID & "'"
                SQL = SQL & " , resultflag = '1' "
                'SQL = SQL & " , printflag  = '5' "
                SQL = SQL & " , textresultval= '" & sResult & "'"
                '-- ����� ��ġ���̸�
                If IsNumeric(sResult) Then
                    SQL = SQL & " , NUMRESULTVAL = '" & sResult & "'"
                End If
                'SQL = SQL & " , ANALYZERCODE= '" & gHOSP.MACHCD & "'"  'GEMINI = 43
                SQL = SQL & " where spcmno = '" & strBarcode & "'"
                SQL = SQL & " and resultitemcode = '" & strTestCd & "'"
                SQL = SQL & " and resultflag < '3' "
                
                Call SetSQLData("�������", SQL, "A")
                AdoCn.Execute SQL
                
                'SaveTransData_JWINFO = 1
            
            End If
        Next intRow
        
        '-- ���º���
        SQL = ""
        SQL = SQL & " UPDATE registinfos SET"
        SQL = SQL & " RESULTSTATE = '1'"
        SQL = SQL & " ,RsvAcptState = '4'"
        SQL = SQL & " where SPCMNO = '" & strBarcode & "'"
        SQL = SQL & "   AND ORDERCODE IN (" & gAllOrdCd & ") " & vbCr
        SQL = SQL & " and CLAS = 4"
        SQL = SQL & " and RESULTSTATE < '4'"
            
        Call SetSQLData("���º���", SQL, "A")
        AdoCn.Execute SQL
        
               
        SaveTransData_AMIS = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_AMIS = -1
    
End Function


Function SaveTransData_MSINFOTEC(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_MSINFOTEC = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local���� ȯ�ں��� ����� ��������
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
        
        Call SetSQLData("���ð����ȸ", SQL)
        
        Set RS_L = New ADODB.Recordset
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- ������ ����� �����ϱ�
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '�˻��ڵ�
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '���(�����)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '���(�������)
                        
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- H/L ����
                strRefVal = f_subSet_RefVal(strTestCd, sResult)
                
                '-- ��������
                SQL = ""
                SQL = SQL & " Update emr.LRESULT" & vbCr
                SQL = SQL & "   Set RSFL = 'Y'," & vbCr
                SQL = SQL & "       RSLT = '" & sResult & "'," & vbCr
                SQL = SQL & "       HLFL = '" & strRefVal & "'," & vbCr
                'SQL = SQL & "       RSDT = '" & Format(Now, "YYYYMMDD") & "'," & vbCr
                SQL = SQL & "       RSDT = sysdate," & vbCr
                SQL = SQL & "       RSID = '" & gHOSP.USERID & "'" & vbCr
                SQL = SQL & " Where SPNO = '" & strBarcode & "'" & vbCr
                SQL = SQL & "   And PAID = '" & strPatID & "'" & vbCr
                'SQL = SQL & "   And ORQN = " & strORQN & vbCr
                SQL = SQL & "   And ORCD = '" & strTestCd & "'" & vbCr
                SQL = SQL & "   And OKFL <> 'Y' "   '-- ���Ȯ������
                
                Call SetSQLData("�������", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
        
        SaveTransData_MSINFOTEC = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_MSINFOTEC = -1
    
End Function

Function SaveTransData_NAVY(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_NAVY = -1
        'intRow = 0
        'strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
'        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
'        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        
'        strDate = Format(Now, "yyyy-mm-dd")
'        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        If Trim(strPatID) = "" Then
            Exit Function
        End If
        
        '-- Local���� ȯ�ں��� ����� ��������
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
        
        Call SetSQLData("���ð����ȸ", SQL)
        
        Set RS_L = New ADODB.Recordset
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- ������ ����� �����ϱ�
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '�˻��ڵ�
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '���(�����)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '���(�������)
                        
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                '-- ��������
                SQL = ""
                SQL = SQL & " Update SLXWORKT"
                SQL = SQL & "   Set PROCSTAT = 'E' " & vbCr
                SQL = SQL & "     , RSLTTEXT = '" & sResult & "'" & vbCr
                SQL = SQL & " Where HOSPID = '" & gHOSP.HOSPCD & "'" & vbCr
                SQL = SQL & "   And PATNO = '" & strPatID & "'" & vbCr
                SQL = SQL & "   And WORKCODE = '" & strChartNo & "'" & vbCr
                SQL = SQL & "   And EXAMCODE = '" & strTestCd & "'" & vbCr
                'SQL = SQL & "   And SPCID = '" & strBarcode & "'" & vbCr
                'SQL = SQL & "   And ORDSEQNO = '" & strORSQ & "'"
                                
                Call SetSQLData("�������", SQL, "A")
                AdoCn.Execute SQL
                
            End If
        Next intRow
        
        SaveTransData_NAVY = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_NAVY = -1
    
End Function



Function SaveTransData_PLIS(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    Dim strTime         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
    Dim strIO           As String
    Dim strKey1         As String
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
    Dim strJudge        As String
    Dim blnSave         As Boolean
    Dim strSeqS         As String
    
    Dim strWorkArea     As String
    Dim strAccSeq       As String
    Dim lngAccSeq       As Long
    
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_PLIS = -1
        intRow = 0
        strJudge = ""
        blnSave = False
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        
        strIO = Trim(GetText(.spdOrder, argSpcRow, colINOUT))
        strKey1 = Trim(GetText(.spdOrder, argSpcRow, colKEY1))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        
        
        'strWorkArea = Trim(GetText(.spdOrder, argSpcRow, colRACKNO))
        'strAccSeq = Trim(GetText(.spdOrder, argSpcRow, colPOSNO))
        
        strDate = Format(Now, "yyyy-mm-dd")
        strTime = Format(Now, "hh:mm:ss")
        
        If Trim(strBarcode) = "" Then
            Exit Function
        End If
        
        
        '-- Local���� ȯ�ں��� ����� ��������
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT,PANICVALUE " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
        'SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
        
        Call SetSQLData("���ð����ȸ", SQL)
        
        Set RS_L = New ADODB.Recordset
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                Call SetText(.vasTemp, RS_L.Fields("PANICVALUE").Value & "", intRow, 7) 'accseq
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- ������ ����� �����ϱ�
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '�˻��ڵ�
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '���(�����)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '���(�������)
            lngAccSeq = Trim(GetText(.vasTemp, intRow, 7))      'accseq
            
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                If mResult.Kind = "QC" Then
                    '-- QC �������
                    SQL = ""
                    SQL = SQL & " UPDATE plis..s2lab026 SET " & vbCr
                    SQL = SQL & "    eqpcd   = '" & gHOSP.MACHCD & "'" & vbCr
                    If IsNumeric(sResult) And InStr(sResult, "+") <= 0 And InStr(sResult, "-") <= 0 Then
                        SQL = SQL & "  , rstval  = '" & sResult & "'" & vbCr
                    End If
                    SQL = SQL & "  , rstcd   = '" & sResult & "'" & vbCr
                    SQL = SQL & "  , rsttype = 'N' " & vbCr
                    SQL = SQL & " WHERE workarea = '" & mOrder.WA & "'" & vbCr
                    SQL = SQL & "   AND accdt    = '" & strHospDate & "'" & vbCr
                    SQL = SQL & "   AND accseq   = '" & lngAccSeq & "'" & vbCr
                    SQL = SQL & "   AND testcd   = '" & strTestCd & "'" & vbCr
                    SQL = SQL & "   And (vfydt IS NULL OR vfydt= '')"
                Else
                    '-- �������
                    SQL = ""
                    SQL = SQL & " UPDATE plis..s2lab302 SET " & vbCr
                    SQL = SQL & "    eqpcd   = '" & gHOSP.MACHCD & "'" & vbCr
                    If IsNumeric(sResult) And InStr(sResult, "+") <= 0 And InStr(sResult, "-") <= 0 Then
                        SQL = SQL & "  , rstval  = '" & sResult & "'" & vbCr
                    End If
                    SQL = SQL & "  , rstcd   = '" & sResult & "'" & vbCr
                    SQL = SQL & "  , rsttype = 'N' " & vbCr
                    SQL = SQL & " WHERE workarea = '" & mOrder.WA & "'" & vbCr
                    SQL = SQL & "   AND accdt    = '" & strHospDate & "'" & vbCr
                    SQL = SQL & "   AND accseq   = '" & lngAccSeq & "'" & vbCr
                    SQL = SQL & "   AND testcd   = '" & strTestCd & "'" & vbCr
                    SQL = SQL & "   And (vfydt IS NULL OR vfydt= '')"
                End If
                
                Call SetSQLData("�������", SQL, "A")
                AdoCn.Execute SQL
                
                SaveTransData_PLIS = 1
            
            End If
        Next intRow
               
        SaveTransData_PLIS = 1
        
    End With

Exit Function

ErrHandle:
    SaveTransData_PLIS = -1
    
End Function

Function GetOrderSeqCode(argExamDt As String, argPID As String, argPCD As String) As String
    Dim RS As ADODB.Recordset
    
    '-- SEQ ��������
    
          SQL = "SELECT /*+ INDEX(rslt scrrslth_ux1) INDEX (coif scccoifm_ix1) */" & vbCr
    SQL = SQL & "       rslt.smp_no, rslt.prcp_seq, rslt.exam_seq, rslt.rept_seq, rslt.cd, rslt.pt_no, rslt.exam_stus, rslt.mach_rslt, rslt.exam_rslt ," & vbCr
    SQL = SQL & "       coif.exam_nm, prex.acp_dt, ptbs.pt_nm, ptbs.ssn_1, ptbs.ssn_2, xpsl.pt_no, " & vbCr
    SQL = SQL & "       DECODE(xpsl.gnl_add_typ_cd,'3','I',xpsl.prcp_knd_cd), xpsl.adms_ymd, xpsl.mn_sub_typ_cd, xpsl.med_dpt_cd, xpsl.med_ymd, coif.spc_cd, codm.cd_desc" & vbCr
    SQL = SQL & "  FROM scrrslth rslt, scccoifm coif, scccodem codm, scrprexh prex, mosxpslh xpsl, pmcptbsm ptbs" & vbCr
    SQL = SQL & " WHERE rslt.hos_org_no   = '" & gHOSP.HOSPCD & "'" & vbCr & vbCr
    SQL = SQL & "  AND SUBSTR(prex.acp_dt,1,8) BETWEEN '" & argExamDt & "' AND '" & argExamDt & "'" & vbCr
    SQL = SQL & "  AND rslt.smp_no       = '" & argPID & "'" & vbCr
    SQL = SQL & "  AND rslt.cd           = '" & argPCD & "'" & vbCr
    SQL = SQL & "  AND rslt.exam_stus  IN ('0','1','2')" & vbCr
    SQL = SQL & "  AND coif.hos_org_no   = rslt.hos_org_no" & vbCr
    SQL = SQL & "  AND coif.exam_cd      = rslt.cd" & vbCr
    SQL = SQL & "  AND SUBSTR(prex.acp_dt,1,8) BETWEEN coif.fr_dt AND coif.to_dt" & vbCr
    SQL = SQL & "  AND coif.exam_mach_cd = '" & gHOSP.MACHCD & "'" & vbCr
    SQL = SQL & "  AND codm.hos_org_no   = coif.hos_org_no" & vbCr
    SQL = SQL & "  AND codm.typ_cd       = '02'" & vbCr
    SQL = SQL & "  AND codm.cd           = coif.spc_cd" & vbCr
    SQL = SQL & "  AND SUBSTR(prex.acp_dt,1,8) BETWEEN codm.fr_dt AND codm.to_dt" & vbCr
    SQL = SQL & "  AND prex.hos_org_no   = rslt.hos_org_no" & vbCr
    SQL = SQL & "  AND prex.smp_no       = rslt.smp_no" & vbCr
    SQL = SQL & "  AND prex.prcp_seq     = rslt.prcp_seq" & vbCr
    SQL = SQL & "  AND prex.exam_seq     = rslt.exam_seq" & vbCr
    SQL = SQL & "  AND xpsl.hos_org_no   = prex.hos_org_no" & vbCr
    SQL = SQL & "  AND xpsl.smp_no       = prex.smp_no" & vbCr
    SQL = SQL & "  AND xpsl.acp_no       = prex.prcp_seq" & vbCr
    SQL = SQL & "  AND xpsl.prcp_typ_cd IN ('O','C')" & vbCr
    SQL = SQL & "  AND ptbs.hos_org_no   = prex.hos_org_no" & vbCr
    SQL = SQL & "  AND ptbs.pt_no        = prex.pt_no" & vbCr

    Call SetSQLData("SEQã��", SQL)

    '-- Record Count ������
    AdoCn.CursorLocation = adUseClient
    Set RS = AdoCn.Execute(SQL, , 1)
    If Not RS.EOF = True And Not RS.BOF = True Then
        Do Until RS.EOF
            GetOrderSeqCode = GetOrderSeqCode & Trim(RS.Fields("prcp_seq")) & "|" & Trim(RS.Fields("exam_seq")) & "|" & Trim(RS.Fields("rept_seq")) & "|"
            RS.MoveNext
        Loop
    End If
    
    If GetOrderSeqCode <> "" Then
        GetOrderSeqCode = Mid(GetOrderSeqCode, 1, Len(GetOrderSeqCode) - 1)
    End If
    
    Set RS = Nothing
    
End Function


Function SaveTransData_MCC_R(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
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
    
    'Dim strReturn       As String
    Dim intReturn       As Long
    Dim strMSG          As String
    
    Dim prm0 As New ADODB.Parameter
    Dim prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim prm3 As New ADODB.Parameter
    Dim prm4 As New ADODB.Parameter
    Dim prm5 As New ADODB.Parameter
    
    
    Dim intBarno  As Double
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_MCC_R = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdROrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdROrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdROrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdROrder, argSpcRow, colBARCODE))
        strChartNo = Trim(GetText(.spdROrder, argSpcRow, colCHARTNO))
        strPatID = Trim(GetText(.spdROrder, argSpcRow, colPID))
        
        If IsNumeric(strBarcode) Then
            intBarno = CDbl(strBarcode)
        Else
            Exit Function
        End If
        
        '-- Local���� ȯ�ں��� ����� ��������
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
        
        Call SetSQLData("���ð����ȸ", SQL)
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- ������ ����� �����ϱ�
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '�˻��ڵ�
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '���(�����)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '���(�������)
                        
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                      SQL = "Exec UP_LIS_INTERFACE_U$001 " & intBarno
                SQL = SQL & "," & strTestCd
                SQL = SQL & "," & sResult
                SQL = SQL & "," & gHOSP.MACHCD

                'AdoCn.Execute SQL
                
                Set AdoCmd = New ADODB.Command
                Set AdoCmd.ActiveConnection = AdoCn
                With AdoCmd
                    .CommandTimeout = 15
                    .CommandText = "UP_LIS_INTERFACE_U$001"
                    .CommandType = adCmdStoredProc
                    
                    
                    Set prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, intBarno)      '���ڵ��ȣ
                    .Parameters.Append prm1

                    Set prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       'ó���ڵ�
                    .Parameters.Append prm2

                    Set prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '�����
                    .Parameters.Append prm3

                    Set prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, gHOSP.MACHCD)    '����ڵ�
                    .Parameters.Append prm4

                    .Execute
                    
                End With
                
                'Call SetSQLData("�������", SQL)
                
                SaveTransData_MCC_R = 1
                
            End If
        Next intRow
        
    End With

Exit Function

ErrHandle:
    SaveTransData_MCC_R = -1
    
End Function

Function SaveTransData_MCC_VERSACELL(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
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
    
    Dim intReturn       As Long
    Dim strMSG          As String
    
    Dim prm0 As New ADODB.Parameter
    Dim prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim prm3 As New ADODB.Parameter
    Dim prm4 As New ADODB.Parameter
    Dim prm5 As New ADODB.Parameter
    
    Dim intBarno        As Double
    
    Dim strMachCD       As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_MCC_VERSACELL = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdOrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdOrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdOrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdOrder, argSpcRow, colBARCODE))
        strChartNo = Trim(GetText(.spdOrder, argSpcRow, colCHARTNO))
        strPatID = Trim(GetText(.spdOrder, argSpcRow, colPID))
        
        If IsNumeric(strBarcode) Then
            intBarno = CDbl(strBarcode)
        Else
            Exit Function
        End If
        
        '-- Local���� ȯ�ں��� ����� ��������
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "      ,DISKNO "                                                        'VERSACELL ���� �������� ����ڵ带 �����س��� �ִ�.
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
        
        Call SetSQLData("���ð����ȸ", SQL)
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                Call SetText(.vasTemp, RS_L.Fields("DISKNO").Value & "", intRow, 7)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- ������ ����� �����ϱ�
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '�˻��ڵ�
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '���(�����)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '���(�������)
            strMachCD = Trim(GetText(.vasTemp, intRow, 7))      '����ڵ� : ADVIA1800, CENTAURXP
                        
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                      SQL = "Exec UP_LIS_INTERFACE_U$001 " & intBarno
                SQL = SQL & "," & strTestCd
                SQL = SQL & "," & sResult
                SQL = SQL & "," & strMachCD

                'AdoCn.Execute SQL
                
                Set AdoCmd = New ADODB.Command
                Set AdoCmd.ActiveConnection = AdoCn
                With AdoCmd
                    .CommandTimeout = 15
                    .CommandText = "UP_LIS_INTERFACE_U$001"
                    .CommandType = adCmdStoredProc
                                        
                    Set prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, intBarno)      '���ڵ��ȣ
                    .Parameters.Append prm1

                    Set prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       'ó���ڵ�
                    .Parameters.Append prm2

                    Set prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '�����
                    .Parameters.Append prm3

                    Set prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, strMachCD)    '����ڵ�
                    .Parameters.Append prm4

                    .Execute
                    
                End With
                
                Call SetSQLData("�������", SQL)
                
                SaveTransData_MCC_VERSACELL = 1
                
            End If
        Next intRow
                
    End With

Exit Function

ErrHandle:
    SaveTransData_MCC_VERSACELL = -1
    
End Function

Function SaveTransData_MCC_VERSACELL_R(ByVal argSpcRow As Integer) As Integer
    Dim RS_L            As ADODB.Recordset
    Dim intRow          As Integer
    Dim strDate         As String
    
    Dim strSaveSeq      As String
    Dim strExamDate     As String
    Dim strHospDate     As String
    Dim strBarcode      As String
    Dim strChartNo      As String
    Dim strPatID        As String
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
    
    Dim intReturn       As Long
    Dim strMSG          As String
    
    Dim prm0 As New ADODB.Parameter
    Dim prm1 As New ADODB.Parameter
    Dim prm2 As New ADODB.Parameter
    Dim prm3 As New ADODB.Parameter
    Dim prm4 As New ADODB.Parameter
    Dim prm5 As New ADODB.Parameter
    
    Dim intBarno        As Double
    
    Dim strMachCD       As String
    
On Error GoTo ErrHandle

    With frmMain
        SaveTransData_MCC_VERSACELL_R = -1
        intRow = 0
        
        strSaveSeq = Trim(GetText(.spdROrder, argSpcRow, colSAVESEQ))
        strExamDate = Trim(GetText(.spdROrder, argSpcRow, colEXAMDATE))
        strHospDate = Trim(GetText(.spdROrder, argSpcRow, colHOSPDATE))
        strBarcode = Trim(GetText(.spdROrder, argSpcRow, colBARCODE))
        strChartNo = Trim(GetText(.spdROrder, argSpcRow, colCHARTNO))
        strPatID = Trim(GetText(.spdROrder, argSpcRow, colPID))
        
        If IsNumeric(strBarcode) Then
            intBarno = CDbl(strBarcode)
        Else
            Exit Function
        End If
        
        '-- Local���� ȯ�ں��� ����� ��������
        .vasTemp.MaxRows = 0
        
              SQL = "SELECT EQUIPCODE,ORDERCODE,EXAMCODE,EXAMSUBCODE,EQUIPRESULT,RESULT " & vbCr
        SQL = SQL & "      ,DISKNO "                                                        'VERSACELL ���� �������� ����ڵ带 �����س��� �ִ�.
        SQL = SQL & "  FROM PATRESULT " & vbCr
        SQL = SQL & " WHERE EQUIPNO = '" & gHOSP.MACHCD & "'" & vbCr                      '����ڵ�
        SQL = SQL & "   AND SAVESEQ = " & strSaveSeq & vbCr                               '�����ȣ
        SQL = SQL & "   AND MID(EXAMDATE,1,8) = '" & Mid(strExamDate, 1, 8) & "'" & vbCr  '�˻���
        SQL = SQL & "   AND BARCODE = '" & strBarcode & "' " & vbCr                       '���ڵ�
        SQL = SQL & "   AND SENDFLAG <> '2' "
        Call SetSQLData("���ð����ȸ", SQL)
        
        Set RS_L = AdoCn_Local.Execute(SQL, , 1)
        If Not RS_L.EOF = True And Not RS_L.BOF = True Then
            .vasTemp.MaxRows = RS_L.RecordCount
            Do Until RS_L.EOF
                intRow = intRow + 1
                Call SetText(.vasTemp, RS_L.Fields("EQUIPCODE").Value & "", intRow, 1)
                Call SetText(.vasTemp, RS_L.Fields("ORDERCODE").Value & "", intRow, 2)
                Call SetText(.vasTemp, RS_L.Fields("EXAMCODE").Value & "", intRow, 3)
                Call SetText(.vasTemp, RS_L.Fields("EXAMSUBCODE").Value & "", intRow, 4)
                Call SetText(.vasTemp, RS_L.Fields("EQUIPRESULT").Value & "", intRow, 5)
                Call SetText(.vasTemp, RS_L.Fields("RESULT").Value & "", intRow, 6)
                Call SetText(.vasTemp, RS_L.Fields("DISKNO").Value & "", intRow, 7)
                RS_L.MoveNext
            Loop
        End If
        
        RS_L.Close
        
        sResult = ""
        sResult1 = ""
        sResult2 = ""
        
        '-- ������ ����� �����ϱ�
        For intRow = 1 To .vasTemp.DataRowCnt
            strTestCd = Trim(GetText(.vasTemp, intRow, 3))      '�˻��ڵ�
            sResult1 = Trim(GetText(.vasTemp, intRow, 5))       '���(�����)
            sResult2 = Trim(GetText(.vasTemp, intRow, 6))       '���(�������)
            strMachCD = Trim(GetText(.vasTemp, intRow, 7))      '����ڵ� : ADVIA1800, CENTAURXP
                        
            '-- ���������
            If .optSaveResult(0).Value = True Then
                sResult = sResult1
            Else
                sResult = sResult2
            End If
            
            If strBarcode <> "" And strTestCd <> "" And sResult <> "" Then
                      SQL = "Exec UP_LIS_INTERFACE_U$001 " & intBarno
                SQL = SQL & "," & strTestCd
                SQL = SQL & "," & sResult
                SQL = SQL & "," & strMachCD

                'AdoCn.Execute SQL
                
                Set AdoCmd = New ADODB.Command
                Set AdoCmd.ActiveConnection = AdoCn
                With AdoCmd
                    .CommandTimeout = 15
                    .CommandText = "UP_LIS_INTERFACE_U$001"
                    .CommandType = adCmdStoredProc
                    
                    
                    Set prm1 = .CreateParameter("BCODE_NO", adInteger, adParamInput, 30, intBarno)      '���ڵ��ȣ
                    .Parameters.Append prm1

                    Set prm2 = .CreateParameter("ORD_CD", adVarChar, adParamInput, 10, strTestCd)       'ó���ڵ�
                    .Parameters.Append prm2

                    Set prm3 = .CreateParameter("RESULT_NM", adVarChar, adParamInput, 4000, sResult)    '�����
                    .Parameters.Append prm3

                    Set prm4 = .CreateParameter("EQP_CD", adVarChar, adParamInput, 15, strMachCD)    '����ڵ�
                    .Parameters.Append prm4

                    .Execute
                    
                End With
                
                Call SetSQLData("�������", SQL)
                
                SaveTransData_MCC_VERSACELL_R = 1
                
            End If
        Next intRow
                
    End With

Exit Function

ErrHandle:
    SaveTransData_MCC_VERSACELL_R = -1
    
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
