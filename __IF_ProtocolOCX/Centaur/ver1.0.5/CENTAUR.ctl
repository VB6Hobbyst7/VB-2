VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl CENTAUR 
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   3330
   Begin VB.CommandButton cmdTest 
      Caption         =   "TEST"
      Height          =   375
      Left            =   210
      TabIndex        =   1
      Top             =   1725
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   1395
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   0
      Top             =   135
      Width           =   1365
   End
   Begin MSCommLib.MSComm msComm 
      Left            =   255
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "CENTAUR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�⺻ �Ӽ� ��:
Const m_def_EqName = "0"
Const m_def_bUseBarcode = 0
Const m_def_iPhase = 0
Const m_def_iSendPhase = 0
Const m_def_sTestMode = "0"
Const m_def_iFrameN = 0
Const m_def_p_sID = "0"
Const m_def_p_sSeq = "0"
Const m_def_p_sRack = "0"
Const m_def_p_sPos = "0"
Const m_def_p_iOrdCnt = 0
Const m_def_p_sTIFCd = "0"
Const m_def_PortOpen = 0
Const m_def_OpenPW = "0"
Const m_def_EditPW = "0"
'�Ӽ� ����:
Dim m_EqName As String
Dim m_bUseBarcode As Boolean
Dim m_iPhase As Integer
Dim m_iSendPhase As Integer
Dim m_sTestMode As String
Dim m_iFrameN As Integer
Dim m_p_sID As String
Dim m_p_sSeq As String
Dim m_p_sRack As String
Dim m_p_sPos As String
Dim m_p_iOrdCnt As Integer
Dim m_p_sTIFCd As String
Dim m_PortOpen As Boolean
Dim m_OpenPW As String
Dim m_EditPW As String
'�̺�Ʈ ����:
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTRstDT$, sOther1$)
Event RaiseError(sError$)
Event SendOrderOK(sID$, sSeqNo$, sRack$, sPos$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event RequestCurOrder(sID$, sRack$, sPos$)
Event DispMsg(sMsg$)
Event RequestNextOrder()
'Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)


'===== User Define
'�������̽����� ���
Dim RcvBuffer   As String
Dim wkBuf   As String
Public sState  As String
Dim sReqStatusCd    As String

'����ü ����
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'��Ÿ
Dim iSpaceCnt   As Integer

'For E-170/Hitachi7600
Dim bEndChk As Boolean
Dim bSTXChk As Boolean
Dim sNextSend   As String
Dim RstEnd      As String
Dim maSpcNo() As String
Dim maSendBuf() As String
Dim miSendCnt As Integer
Dim miIndex   As Integer

Dim msMsgID    As String
Dim msSender   As String
Dim msReceiver As String
Dim msVersion  As String

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=msComm,msComm,-1,CommPort
Public Property Get CommPort() As Integer
Attribute CommPort.VB_Description = "��� ��Ʈ ��ȣ�� ��ȯ�ϰų� �����մϴ�."
    CommPort = msComm.CommPort
End Property

Public Property Let CommPort(ByVal New_CommPort As Integer)
    msComm.CommPort() = New_CommPort
    PropertyChanged "CommPort"
End Property

Private Sub PhaseCfg_Protocol()

    '--- ����� Ȯ��
    If m_EditPW <> pEditPW Then
        MsgBox "��ϵ� ����ڰ� �ƴմϴ�. (��)���̾����̷� ������ �ֽʽÿ�!!!", vbCritical, "����� Ȯ��"
        Exit Sub
    End If
    '---------------
    
    If m_EqName = "0" Or m_EqName = "" Then
        RaiseEvent DispMsg("�˻������� ������ �ֽʽÿ�.!!!")
        Exit Sub
    End If
    
    Select Case UCase(m_EqName)
        Case "CENTAUR"
            If bUseBarcode = True Then
                Call PhaseCfg_Protocol_Centaur
            Else
                Call PhaseCfg_Protocol_Centaur_Batch
            End If
            
        Case "CENTAURCP_UNPACKED"
            Call PhaseCfg_Protocol_CentaurCP_UnPacked
            
        Case Else
            RaiseEvent DispMsg("�������� �ʴ� ��� �����߽��ϴ�.")
            
    End Select
    
End Sub

Private Sub PhaseCfg_Protocol_Centaur()
    On Error GoTo ErrRtn
    
    Dim wkDat   As String
    Dim ix1 As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        m_iPhase = 2
                        RstEnd = "Y"
                        bEndChk = True: bSTXChk = False

                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                    Case Else
                        m_iPhase = 1
                End Select

            Case 2
                Select Case Asc(wkDat)
                    Case 2      'STX
                        If bEndChk = True Then
                            RcvBuffer = ""
                        Else
                            bSTXChk = True
                        End If
                        bEndChk = True

                    Case 10     '<LF>
                        If bEndChk = True Then
                            Call DataEditResponse_Centaur
                            RcvBuffer = ""
                        End If
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                    Case 13     'CR
                        If bEndChk = True Then
                            Call DataEditResponse_Centaur
                            RcvBuffer = ""
                        End If

                    Case 4      'EOT
                        If sState = "Q" Then
                            msComm.Output = Chr(5)
                            
                            If m_sTestMode = "77" Then
                                RaiseEvent PrintSendLog(Chr(5))
                            End If
                            
                            m_iSendPhase = 1
                        End If
                        m_iPhase = 3

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = True
                        msComm.Output = Chr(6)   'Send ACK
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                    Case 21     'NAK
                        Call DataEditResponse_Centaur

                        m_iSendPhase = 1
                        m_iFrameN = 1

                    Case 23     ' ETB
                        bEndChk = False

                    Case Else
                        If bEndChk = True Then
                            If bSTXChk = True Then
                                bSTXChk = False
                            Else
                                RcvBuffer = RcvBuffer & wkDat
                            End If
                        End If

                End Select

            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        If sState = "Q" Then
                            Call SendOrder_Centaur
                        End If

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If
                        
                        m_iPhase = 2

                    Case 21     'NAK
                        m_iSendPhase = 1
                        m_iFrameN = 1
                        m_iPhase = 3

                    Case 4      'EOT
                        m_iPhase = 1

                End Select
        End Select
    Next ix1
    
ErrRtn:
    If Err <> 0 Then
        RcvBuffer = ""
        RaiseEvent DispMsg(Err.Description)
    End If
End Sub

Private Sub PhaseCfg_Protocol_Centaur_Batch()
    On Error GoTo ErrRtn
    
    Dim wkDat   As String
    Dim ix1 As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        m_iPhase = 2
                        RstEnd = "Y"
                        bEndChk = True: bSTXChk = False

                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                    Case Else
                        m_iPhase = 1
                End Select

            Case 2
                Select Case Asc(wkDat)
                    Case 2      'STX
                        If bEndChk = True Then
                            RcvBuffer = ""
                        Else
                            bSTXChk = True
                        End If
                        bEndChk = True

                    Case 10     '<LF>
                        If bEndChk = True Then
                            Call DataEditResponse_Centaur
                            RcvBuffer = ""
                        End If
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                    Case 13     'CR
                        If bEndChk = True Then
                            Call DataEditResponse_Centaur
                            RcvBuffer = ""
                        End If

                    Case 4      'EOT
                        If sState = "Q" Then
                            msComm.Output = Chr(5)
                            
                            If m_sTestMode = "77" Then
                                RaiseEvent PrintSendLog(Chr(5))
                            End If
                            
                            m_iSendPhase = 1
                        End If
                        m_iPhase = 3

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = True
                        msComm.Output = Chr(6)   'Send ACK
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                    Case 21     'NAK
                        Call DataEditResponse_Centaur

                        m_iSendPhase = 1
                        m_iFrameN = 1

''                        msComm.Output = Chr(5)   'Send ENQ
                        
                    Case 23     ' ETB
                        bEndChk = False

                    Case Else
                        If bEndChk = True Then
                            If bSTXChk = True Then
                                bSTXChk = False
                            Else
                                RcvBuffer = RcvBuffer & wkDat
                            End If
                        End If

                End Select

            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        If sState = "Q" Then
                            Call SendOrder_Centaur_Batch
                        End If

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If
                        
                        m_iPhase = 2

                    Case 21     'NAK
                        m_iSendPhase = 1
                        m_iFrameN = 1
''                        msComm.Output = Chr(5)
                        m_iPhase = 3

                    Case 4      'EOT
                        m_iPhase = 1

                End Select
        End Select
    Next ix1
    
ErrRtn:
    If Err <> 0 Then
        RcvBuffer = ""
        RaiseEvent DispMsg(Err.Description)
    End If
End Sub

Private Sub PhaseCfg_Protocol_CentaurCP_UnPacked()
    On Error GoTo ErrRtn
    
    Dim wkDat   As String
    Dim ix1 As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        m_iPhase = 2
                        RstEnd = "Y"
                        bEndChk = True: bSTXChk = False

                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                    Case Else
                        m_iPhase = 1
                End Select

            Case 2
                Select Case Asc(wkDat)
                    Case 2      'STX
                        If bEndChk = True Then
                            RcvBuffer = ""
                        Else
                            bSTXChk = True
                        End If
                        bEndChk = True

                    Case 10     '<LF>
                        If bEndChk = True Then
                            Call DataEditResponse_Centaur_UnPacked
                            RcvBuffer = ""
                        End If
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                    Case 13     'CR
                        If bEndChk = True Then
                            Call DataEditResponse_Centaur_UnPacked
                            RcvBuffer = ""
                        End If

                    Case 4      'EOT
                        If sState = "Q" Then
                            msComm.Output = Chr(5)
                            
                            If m_sTestMode = "77" Then
                                RaiseEvent PrintSendLog(Chr(5))
                            End If
                            
                            m_iSendPhase = 1
                        End If
                        m_iPhase = 3

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = True
                        msComm.Output = Chr(6)   'Send ACK
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                    Case 21     'NAK
                        Call DataEditResponse_Centaur_UnPacked

                        m_iSendPhase = 1
                        m_iFrameN = 1

''                        msComm.Output = Chr(5)   'Send ENQ
''
''                        If m_sTestMode = "77" Then
''                            RaiseEvent PrintSendLog(Chr(5))
''                        End If

                    Case 23     ' ETB
                        bEndChk = False

                    Case Else
                        If bEndChk = True Then
                            If bSTXChk = True Then
                                bSTXChk = False
                            Else
                                RcvBuffer = RcvBuffer & wkDat
                            End If
                        End If

                End Select

            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        If sState = "Q" Then
                            Call SendOrder_Centaur_UnPacked
                        End If

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If
                        
                        m_iPhase = 2

                    Case 21     'NAK
                        m_iSendPhase = 1
                        m_iFrameN = 1
                        ''msComm.Output = Chr(5)
                        m_iPhase = 3

                    Case 4      'EOT
                        m_iPhase = 1

                End Select
        End Select
    Next ix1
    
ErrRtn:
    If Err <> 0 Then
        RcvBuffer = ""
        RaiseEvent DispMsg(Err.Description)
    End If
End Sub


' *=====================================================*
' *               Data���� & ����ó��                   *
' *=====================================================*
Private Sub DataEditResponse_Centaur()
    On Error GoTo ErrRtn

    Dim RecType As String   'Record Type
    Dim ii      As Integer
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    Dim tmpField()  As String
    Dim tmpData()   As String
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$
    Dim sRstState   As String
    Dim tmpRstDT$

    ii = InStr(1, RcvBuffer, "|")
    If ii <> 0 Then
        RecType = Mid$(RcvBuffer, ii - 1, 1)
    Else
        Exit Sub
    End If

    Select Case RecType
        Case "H"        'Header Record
            sState = ""
            pSampleInfo.ID = ""
        
        Case "M"
        Case "P"        'Patient Record
            Call Init_pResultInfo

        Case "Q"        'Order Request Record
            tmpField() = Split(RcvBuffer, "|")
            sReqStatusCd = Trim(tmpField(12))    'Order Request Status Code

            If InStr(tmpField(2), "^") > 0 Then
                tmpData() = Split(tmpField(2), "^")
                tmpBarCd = Trim(tmpData(1))
            Else
                tmpBarCd = ""
            End If
            tmpSeqNo = ""
            tmpRack = ""
            tmpPos = ""

            If tmpBarCd <> "" Then    'BarCode ID�� �� �Ѿ�Դ��� �˻�
                sState = "Q"
                pSampleInfo.ID = UCase(tmpBarCd)
            Else
                sState = ""
                pSampleInfo.ID = ""
            End If
            
            If pSampleInfo.ID = "ALL" Then
                RaiseEvent RequestNextOrder
            Else
                pSampleInfo.SEQNO = tmpSeqNo
                pSampleInfo.RACK = tmpRack
                pSampleInfo.POS = tmpPos
            End If

        Case "O"
            tmpSeqNo = "": tmpBarCd = "": tmpRack = "": tmpPos = ""
            tmpField() = Split(RcvBuffer, "|")
            ii = InStr(1, tmpField(2), "^")
            If ii <> 0 Then
                tmpData() = Split(tmpField(2), "^")
                tmpBarCd = Trim(tmpData(0))
                tmpRack = Trim(tmpData(1))
                tmpPos = Trim(tmpData(2))
            Else
                tmpBarCd = Trim(tmpField(2))
            End If

            pSampleInfo.ID = UCase(tmpBarCd)
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos

        Case "R"        'Result Record
            '--- �������Ÿ ����
            '2:TEST ID
            '3:RESULT
            '4:UNITS
            '6:Result Abnormal Flags
            '8:Result Status
            '12:Result Date/Time
            tmpField() = Split(RcvBuffer, "|")

            tmpData() = Split(tmpField(2), "^")
            '2004/3/22 yk update
            sRstState = Trim(tmpData(7))
            tmpIFCd = Trim(tmpData(3)) & "^" & sRstState

            tmpRst = Trim(tmpField(3))
            tmpUnit = Trim(tmpField(4))
            tmpFlag = Trim(tmpField(6))

            tmpRstDT = Trim(tmpField(12))
            
            '--- ������� "^" �� ��� ����
            ii = InStr(1, tmpRst, "^")
            If ii <> 0 Then tmpRst = Mid(tmpRst, ii + 1)

            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If

            '������� ����ü�� ����
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS

                '����� ����
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
                .RSTDT = .RSTDT & tmpRstDT & Chr(124)  '����Ͻ�(2005/6/11) yk
            End With

        Case "C"        'Comment Record

        Case "L"
            '����� ���/ȭ�� ǥ�� ó��...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .RSTDT, "")
                End If
            End With

            Call Init_pResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit �����߻� - " & Err.Description)
    End If
End Sub

Private Sub DataEditResponse_Centaur_UnPacked()
    On Error GoTo ErrRtn

    Dim RecType As String   'Record Type
    Dim ii      As Integer
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    Dim tmpField()  As String
    Dim tmpData()   As String
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$
    Dim sRstState   As String
    Dim tmpRstDT$
    Dim tmpKind$

    ii = InStr(1, RcvBuffer, "|")
    If ii <> 0 Then
        RecType = Mid$(RcvBuffer, ii - 1, 1)
    Else
        Exit Sub
    End If

    Select Case RecType
        Case "H"        'Header Record
            sState = ""
            pSampleInfo.ID = ""
            
            '1H|\^&|||ACCP1|||||Host||P|1|20120619122819
            tmpField = Split(RcvBuffer, "|")

            msMsgID = Trim(tmpField(2))
            msSender = Trim(tmpField(4))
            msReceiver = Trim(tmpField(9))
            msVersion = Trim(tmpField(12))

        Case "M"
        Case "P"        'Patient Record
            Call Init_pResultInfo

        Case "Q"        'Order Request Record
            miIndex = 0
            
            tmpField() = Split(RcvBuffer, "|")
            sReqStatusCd = Trim(tmpField(12))    'Order Request Status Code

            If InStr(tmpField(2), "^") > 0 Then
                tmpData() = Split(tmpField(2), "^")
                ''tmpBarCd = Trim(tmpData(1))
                maSpcNo = Split(tmpData(1), "\")
                
                If UBound(maSpcNo) >= 0 Then
                    tmpBarCd = maSpcNo(0)
                Else
                    tmpBarCd = ""
                End If
            Else
                tmpBarCd = ""
            End If
            
            tmpSeqNo = ""
            tmpRack = ""
            tmpPos = ""

            If tmpBarCd <> "" Then    'BarCode ID�� �� �Ѿ�Դ��� �˻�
                sState = "Q"
                pSampleInfo.ID = UCase(tmpBarCd)
            Else
                sState = ""
                pSampleInfo.ID = ""
            End If
            
            If pSampleInfo.ID = "ALL" Then
                RaiseEvent RequestNextOrder
            Else
                pSampleInfo.SEQNO = tmpSeqNo
                pSampleInfo.RACK = tmpRack
                pSampleInfo.POS = tmpPos
            End If

        Case "O"
            tmpSeqNo = "": tmpBarCd = "": tmpRack = "": tmpPos = ""
            tmpField() = Split(RcvBuffer, "|")
            ii = InStr(1, tmpField(2), "^")
            If ii <> 0 Then
                tmpData() = Split(tmpField(2), "^")
                tmpBarCd = Trim(tmpData(0))
                tmpRack = Trim(tmpData(1))
                tmpPos = Trim(tmpData(2))
            Else
                tmpBarCd = Trim(tmpField(2))
            End If
            
            tmpKind = Trim(tmpField(11))
            
            pSampleInfo.ID = UCase(tmpBarCd)
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos
            pSampleInfo.Kind = tmpKind

        Case "R"        'Result Record
            '--- �������Ÿ ����
            '2:TEST ID
            '3:RESULT
            '4:UNITS
            '6:Result Abnormal Flags
            '8:Result Status
            '12:Result Date/Time
            tmpField() = Split(RcvBuffer, "|")

            tmpData() = Split(tmpField(2), "^")
            '2004/3/22 yk update
            sRstState = Trim(tmpData(7))
            tmpIFCd = Trim(tmpData(3)) & "^" & sRstState

            tmpRst = Trim(tmpField(3))
            tmpUnit = Trim(tmpField(4))
            tmpFlag = Trim(tmpField(6))

            tmpRstDT = Trim(tmpField(12))
            
            '--- ������� "^" �� ��� ����
            ii = InStr(1, tmpRst, "^")
            If ii <> 0 Then tmpRst = Mid(tmpRst, ii + 1)

            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If

            '������� ����ü�� ����
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS
                .Kind = pSampleInfo.Kind

                '����� ����
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
                .RSTDT = .RSTDT & tmpRstDT & Chr(124)  '����Ͻ�(2005/6/11) yk
            End With
            
            '����� ���/ȭ�� ǥ�� ó��...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .RSTDT, .Kind)
                End If
            End With

            Call Init_pResultInfo

        Case "C"        'Comment Record

        Case "L"
''            '����� ���/ȭ�� ǥ�� ó��...
''            With pResultInfo
''                If .RSTCNT > 0 Then
''                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .RSTDT, "")
''                End If
''            End With
''
''            Call Init_pResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit �����߻� - " & Err.Description)
    End If
End Sub

'
'   ȯ�� Order ����
'
Private Sub SendOrder_Centaur()
    On Error GoTo Err_Rtn

    Dim sSendBuff   As String
    Dim iCnt    As Integer
    Dim ChkSum  As String
    Dim sActionCd   As String
    
    Dim tmpBarCd    As String   '��񿡼� ���� ���ڵ� �ӽ� ����
    Dim tmpSpcNo    As String   '���ü��ȣ �ӽ� ����
    
    Select Case m_iSendPhase
        Case 1
            'Header Record
            sSendBuff = m_iFrameN & "H|\^&|||LIS_ID|||||NG_LIS||P|1" & vbCr
            
            tmpBarCd = pSampleInfo.ID
            
            '----- �˻��׸� ��ȸ
            RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)

            Call Get_OrderString
            
            '��񿡼� ���� �Ѿ�� ���ڵ�� ����
            If pSampleInfo.ID <> "" Then
                pSampleInfo.ID = tmpBarCd
            End If
            
            '�������� ���...2006/5/17 yk
            If pSampleInfo.ORDCNT = 0 Then
                'Terminator Record (N:Normal, F:Final, I:No Information Available)
                sSendBuff = sSendBuff & "L|1|I"
            
                sSendBuff = sSendBuff & Chr(13) & Chr(3)
                GoTo Send_Terminate
            End If
            
            'Patient Record
            '3(2):Practice Assigned Patient ID
            '6(5):Patient Name
            '14(13):Physician
            sSendBuff = sSendBuff & "P|1" & vbCr
'            sSendBuff = sSendBuff & "P|1|" & Trim(Left(pSampleInfo.PATINFO, 11)) & "|||" & _
'                            Trim(Left(pSampleInfo.SAMPINFO, 30)) & "||||||||" & tmpSpcNo & vbCr
                    
            'Order Record
            sSendBuff = sSendBuff & "O|1|" & Trim(pSampleInfo.ID) & "||"


            If pSampleInfo.ORDCNT = 0 Then
                sActionCd = "C"
            Else
                sActionCd = ""
            End If
            
            '�˻��׸� Order�ڵ� �߰�
            For iCnt = 1 To pSampleInfo.ORDCNT
                '���� ����
                sSendBuff = sSendBuff & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "\"
            Next iCnt
            If pSampleInfo.ORDCNT > 0 Then
                sSendBuff = Left(sSendBuff, Len(sSendBuff) - 1)      '"\" Cutting
            End If

'            sSendBuff = sSendBuff & "|R||||||" _
'                    & sActionCd & "||||||||||||||Q" & vbCr
            sSendBuff = sSendBuff & "|R||||||" _
                    & sActionCd & "||||||||||||||O\Q" & vbCr
                    
            'Terminator Record (N:Normal, F:Final, I:No Information Available)
            sSendBuff = sSendBuff & "L|1|F"
                        
            '--- Text�� ������ 240byte�� �Ѿ ��� ó�� �߰�...
            If Len(sSendBuff) >= 241 Then
                sNextSend = Mid(sSendBuff, 241)
                sSendBuff = Left(sSendBuff, 240)
                sSendBuff = sSendBuff & Chr(23)

                m_iFrameN = m_iFrameN + 1
                m_iSendPhase = 2
            Else
                sSendBuff = sSendBuff & Chr(13) & Chr(3)
                GoTo Send_Terminate
            End If

        Case 2
            sSendBuff = m_iFrameN & sNextSend & Chr(13) & Chr(3)
            sNextSend = ""

Send_Terminate:
            m_iSendPhase = 3

        Case 3      'EOT
            msComm.Output = Chr(4)   'EOT
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(4))
            End If
            
            m_iFrameN = 1
            m_iPhase = 3
            m_iSendPhase = 1

            sState = ""

            Exit Sub
    End Select

    ChkSum = ChkSum_ASTM(sSendBuff)
    sSendBuff = sSendBuff & ChkSum
    msComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)

    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If

    '���۵� ������ �ִ� ��� ȭ��ǥ��
    If pSampleInfo.ORDCNT > 0 And sReqStatusCd = "O" Then
        If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
        End If
    Else
        '��ȸ�� ������ ���� ��� ȯ������ ����ü �ʱ�ȭ
        Call Init_pResultInfo

        RaiseEvent SendOrderOK("", "", "", "")
    End If

Err_Rtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order ���۽� �����߻� - " & Err.Description)
    End If
End Sub

''Private Sub SendOrder_Centaur_Batch()
''    On Error GoTo Err_Rtn
''
''    Dim sSendBuff   As String
''    Dim iCnt    As Integer
''    Dim ChkSum  As String
''    Dim sActionCd   As String
''
''    Dim tmpBarCd    As String   '��񿡼� ���� ���ڵ� �ӽ� ����
''    Dim tmpSpcNo    As String   '���ü��ȣ �ӽ� ����
''
''    Select Case m_iSendPhase
''        Case 1
''            'Header Record
''            ''sSendBuff = m_iFrameN & "H|\^&|||LIS_ID|||||NG_LIS||P|1" & vbCr
''            sSendBuff = m_iFrameN & "H|\^&|||Host|||||NG_LIS||P|1" & vbCr
''
''            'tmpBarCd = pSampleInfo.ID
''
''            '----- �˻��׸� ��ȸ
''            'RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
''
''            Call Get_OrderString
''
''''            '��񿡼� ���� �Ѿ�� ���ڵ�� ����
''''            If pSampleInfo.ID <> "" Then
''''                pSampleInfo.ID = tmpBarCd
''''            End If
''
''            '�������� ���...2006/5/17 yk
''            If pSampleInfo.ORDCNT = 0 Then
''                'Terminator Record (N:Normal, F:Final, I:No Information Available)
''                sSendBuff = sSendBuff & "L|1|I"
''
''                sSendBuff = sSendBuff & Chr(13) & Chr(3)
''                GoTo Send_Terminate
''            End If
''
''            'Patient Record
''            '3(2):Practice Assigned Patient ID
''            '6(5):Patient Name
''            '14(13):Physician
''            sSendBuff = sSendBuff & "P|1" & vbCr
'''            sSendBuff = sSendBuff & "P|1|" & Trim(Left(pSampleInfo.PATINFO, 11)) & "|||" & _
'''                            Trim(Left(pSampleInfo.SAMPINFO, 30)) & "||||||||" & tmpSpcNo & vbCr
''
''            'Order Record
''            sSendBuff = sSendBuff & "O|1|" & Trim(pSampleInfo.ID) & "||"
''
''            If pSampleInfo.ORDCNT = 0 Then
''                sActionCd = "C"
''            Else
''                sActionCd = ""
''            End If
''
''            '�˻��׸� Order�ڵ� �߰�
''            For iCnt = 1 To pSampleInfo.ORDCNT
''                '���� ����
''                sSendBuff = sSendBuff & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "\"
''            Next iCnt
''            If pSampleInfo.ORDCNT > 0 Then
''                sSendBuff = Left(sSendBuff, Len(sSendBuff) - 1)      '"\" Cutting
''            End If
''
'''            sSendBuff = sSendBuff & "|R||||||" _
'''                    & sActionCd & "||||||||||||||Q" & vbCr
''            sSendBuff = sSendBuff & "|R||||||" _
''                    & sActionCd & "||||||||||||||O\Q" & vbCr
''
''            'Terminator Record (N:Normal, F:Final, I:No Information Available)
''            sSendBuff = sSendBuff & "L|1|F"
''
''            '--- Text�� ������ 240byte�� �Ѿ ��� ó�� �߰�...
''            If Len(sSendBuff) >= 241 Then
''                sNextSend = Mid(sSendBuff, 241)
''                sSendBuff = Left(sSendBuff, 240)
''                sSendBuff = sSendBuff & Chr(23)
''
''                m_iFrameN = m_iFrameN + 1
''                m_iSendPhase = 2
''            Else
''                sSendBuff = sSendBuff & Chr(13) & Chr(3)
''                GoTo Send_Terminate
''            End If
''
''        Case 2
''            sSendBuff = m_iFrameN & sNextSend & Chr(13) & Chr(3)
''            sNextSend = ""
''
''Send_Terminate:
''            m_iSendPhase = 3
''
''        Case 3      'EOT
''            msComm.Output = Chr(4)   'EOT
''
''            If m_sTestMode = "77" Then
''                RaiseEvent PrintSendLog(Chr(4))
''            End If
''
''            m_iFrameN = 1
''            m_iPhase = 1
''            m_iSendPhase = 1
''            sState = ""
''
''            RaiseEvent RequestNextOrder
''
''            If m_p_iOrdCnt > 0 Then
''                m_iPhase = 3
''                sState = "Q"
''                msComm.Output = Chr(5)  'ENQ
''
''                If m_sTestMode = "77" Then
''                    RaiseEvent PrintSendLog(Chr(5))
''                End If
''            End If
''
''            Exit Sub
''    End Select
''
''    ChkSum = ChkSum_ASTM(sSendBuff)
''    sSendBuff = sSendBuff & ChkSum
''    msComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)
''
''    If m_sTestMode = "77" Then
''        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
''    End If
''
''    '���۵� ������ �ִ� ��� ȭ��ǥ��
''    If pSampleInfo.ORDCNT > 0 Then
''        If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
''            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
''        End If
''    Else
''        '��ȸ�� ������ ���� ��� ȯ������ ����ü �ʱ�ȭ
''        Call Init_pResultInfo
''
''        RaiseEvent SendOrderOK("", "", "", "")
''    End If
''
''Err_Rtn:
''    If Err <> 0 Then
''        RaiseEvent DispMsg("Order ���۽� �����߻� - " & Err.Description)
''    End If
''End Sub

Private Sub SendOrder_Centaur_Batch()
    On Error GoTo Err_Rtn

    Dim sSendBuff   As String
    Dim iCnt    As Integer
    Dim ChkSum  As String
    Dim sActionCd   As String
    
    Dim tmpBarCd    As String   '��񿡼� ���� ���ڵ� �ӽ� ����
    Dim tmpSpcNo    As String   '���ü��ȣ �ӽ� ����
    
    If m_iFrameN > 7 Then
        m_iFrameN = 0
    End If
    
    Select Case m_iSendPhase
        Case 1
            'Header Record
            sSendBuff = m_iFrameN & "H|\^&|||Host|||||ACCP1||P|1|" & Format(Now, "yyyyMMddHHmmss") & Chr(13) & Chr(3)
            m_iSendPhase = 2
            
        Case 2
            Call Get_OrderString
            
            'Patient Record
            '3(2):Practice Assigned Patient ID
            '6(5):Patient Name
            '14(13):Physician
            sSendBuff = m_iFrameN & "P|1" & Chr(13) & Chr(3)
'            sSendBuff = sSendBuff & "P|1|" & Trim(Left(pSampleInfo.PATINFO, 11)) & "|||" & _
'                            Trim(Left(pSampleInfo.SAMPINFO, 30)) & "||||||||" & tmpSpcNo & vbCr

            '�������� ���...2006/5/17 yk
            If pSampleInfo.ORDCNT = 0 Then
                m_iSendPhase = 4
            Else
                m_iSendPhase = 3
            End If
        
        Case 3
            'Order Record
            sSendBuff = m_iFrameN & "O|1|" & Trim(pSampleInfo.ID) & "||"

            If pSampleInfo.ORDCNT = 0 Then
                sActionCd = "C"
            Else
                sActionCd = ""
            End If
            
            '�˻��׸� Order�ڵ� �߰�
            For iCnt = 1 To pSampleInfo.ORDCNT
                '���� ����
                sSendBuff = sSendBuff & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "\"
            Next iCnt
            
            If pSampleInfo.ORDCNT > 0 Then
                sSendBuff = Left(sSendBuff, Len(sSendBuff) - 1)      '"\" Cutting
            End If

            sSendBuff = sSendBuff & "|R||||||" & sActionCd & "||||||||||||||O\Q" & Chr(13) & Chr(3)
            
            ''m_iSendPhase = 4
            
            RaiseEvent RequestNextOrder
            
            If m_p_iOrdCnt > 0 Then
                m_iSendPhase = 2
            Else
                m_iSendPhase = 4
            End If
            
        Case 4
            If pSampleInfo.ORDCNT = 0 Then
                sSendBuff = m_iFrameN & "L|1|I" & Chr(13) & Chr(3)
            Else
                'Terminator Record (N:Normal, F:Final, I:No Information Available)
                sSendBuff = m_iFrameN & "L|1|F"
                            
                '--- Text�� ������ 240byte�� �Ѿ ��� ó�� �߰�...
                If Len(sSendBuff) >= 241 Then
                    sNextSend = Mid(sSendBuff, 241)
                    sSendBuff = Left(sSendBuff, 240)
                    sSendBuff = sSendBuff & Chr(23)
    
                    m_iFrameN = m_iFrameN + 1
                    m_iSendPhase = 2
                Else
                    sSendBuff = sSendBuff & Chr(13) & Chr(3)
                End If
            End If
            
            m_iSendPhase = 5

        Case 5      'EOT
            msComm.Output = Chr(4)   'EOT
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(4))
            End If
            
            m_iFrameN = 1
            m_iPhase = 1
            m_iSendPhase = 1
            sState = ""
            
            RaiseEvent RequestNextOrder
            
            If m_p_iOrdCnt > 0 Then
                m_iPhase = 3
                sState = "Q"
                msComm.Output = Chr(5)  'ENQ
                
                If m_sTestMode = "77" Then
                    RaiseEvent PrintSendLog(Chr(5))
                End If
            End If

            Exit Sub
    End Select

    ChkSum = ChkSum_ASTM(sSendBuff)
    sSendBuff = sSendBuff & ChkSum
    msComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)

    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If
    
    m_iFrameN = m_iFrameN + 1

    '���۵� ������ �ִ� ��� ȭ��ǥ��
    If pSampleInfo.ORDCNT > 0 Then
        If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
        End If
    Else
        '��ȸ�� ������ ���� ��� ȯ������ ����ü �ʱ�ȭ
        Call Init_pResultInfo

        RaiseEvent SendOrderOK("", "", "", "")
    End If

Err_Rtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order ���۽� �����߻� - " & Err.Description)
    End If
End Sub

'''
'''   ȯ�� Order ����
'''
''Private Sub SendOrder_Centaur_UnPacked()
''    On Error GoTo Err_Rtn
''
''    Dim sSendBuff   As String
''    Dim iCnt    As Integer
''    Dim ChkSum  As String
''    Dim sActionCd   As String
''
''    Dim tmpBarCd    As String   '��񿡼� ���� ���ڵ� �ӽ� ����
''    Dim tmpSpcNo    As String   '���ü��ȣ �ӽ� ����
''
''    Select Case m_iSendPhase
''        Case 1
''            'Header Record
''            sSendBuff = m_iFrameN & "H|\^&|||LIS_ID|||||NG_LIS||P|1" & vbCr
''
''            tmpBarCd = maSpcNo(miIndex)
''            miIndex = miIndex + 1
''
''            pSampleInfo.ID = tmpBarCd
''
''            '----- �˻��׸� ��ȸ
''            RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
''
''            Call Get_OrderString
''
''            '��񿡼� ���� �Ѿ�� ���ڵ�� ����
''            If pSampleInfo.ID <> "" Then
''                pSampleInfo.ID = tmpBarCd
''            End If
''
''            '�������� ���...2006/5/17 yk
''            If pSampleInfo.ORDCNT = 0 Then
''                'Terminator Record (N:Normal, F:Final, I:No Information Available)
''                sSendBuff = sSendBuff & "L|1|I"
''
''                sSendBuff = sSendBuff & Chr(13) & Chr(3)
''                GoTo Send_Terminate
''            End If
''
''            'Patient Record
''            '3(2):Practice Assigned Patient ID
''            '6(5):Patient Name
''            '14(13):Physician
''            sSendBuff = sSendBuff & "P|1" & vbCr
'''            sSendBuff = sSendBuff & "P|1|" & Trim(Left(pSampleInfo.PATINFO, 11)) & "|||" & _
'''                            Trim(Left(pSampleInfo.SAMPINFO, 30)) & "||||||||" & tmpSpcNo & vbCr
''
''            'Order Record
''            sSendBuff = sSendBuff & "O|1|" & Trim(pSampleInfo.ID) & "||"
''
''
''            If pSampleInfo.ORDCNT = 0 Then
''                sActionCd = "C"
''            Else
''                sActionCd = ""
''            End If
''
''            '�˻��׸� Order�ڵ� �߰�
''            For iCnt = 1 To pSampleInfo.ORDCNT
''                '���� ����
''                sSendBuff = sSendBuff & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "\"
''            Next iCnt
''            If pSampleInfo.ORDCNT > 0 Then
''                sSendBuff = Left(sSendBuff, Len(sSendBuff) - 1)      '"\" Cutting
''            End If
''
'''            sSendBuff = sSendBuff & "|R||||||" _
'''                    & sActionCd & "||||||||||||||Q" & vbCr
''            sSendBuff = sSendBuff & "|R||||||" _
''                    & sActionCd & "||||||||||||||O\Q" & vbCr
''
''            'Terminator Record (N:Normal, F:Final, I:No Information Available)
''            sSendBuff = sSendBuff & "L|1|F"
''
''            '--- Text�� ������ 240byte�� �Ѿ ��� ó�� �߰�...
''            If Len(sSendBuff) >= 241 Then
''                sNextSend = Mid(sSendBuff, 241)
''                sSendBuff = Left(sSendBuff, 240)
''                sSendBuff = sSendBuff & Chr(23)
''
''                m_iFrameN = m_iFrameN + 1
''                m_iSendPhase = 2
''            Else
''                sSendBuff = sSendBuff & Chr(13) & Chr(3)
''                GoTo Send_Terminate
''            End If
''
''        Case 2
''            sSendBuff = m_iFrameN & sNextSend & Chr(13) & Chr(3)
''            sNextSend = ""
''
''Send_Terminate:
''            m_iSendPhase = 3
''
''        Case 3      'EOT
''            msComm.Output = Chr(4)   'EOT
''
''            If m_sTestMode = "77" Then
''                RaiseEvent PrintSendLog(Chr(4))
''            End If
''
''            m_iFrameN = 1
''            m_iPhase = 3
''            m_iSendPhase = 1
''            miIndex = 0
''
''            sState = ""
''
''            Exit Sub
''    End Select
''
''    ChkSum = ChkSum_ASTM(sSendBuff)
''    sSendBuff = sSendBuff & ChkSum
''    msComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)
''
''    If m_sTestMode = "77" Then
''        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
''    End If
''
''    '���۵� ������ �ִ� ��� ȭ��ǥ��
''    If pSampleInfo.ORDCNT > 0 And sReqStatusCd = "O" Then
''        If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
''            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
''        End If
''    Else
''        '��ȸ�� ������ ���� ��� ȯ������ ����ü �ʱ�ȭ
''        Call Init_pResultInfo
''
''        RaiseEvent SendOrderOK("", "", "", "")
''    End If
''
''Err_Rtn:
''    If Err <> 0 Then
''        RaiseEvent DispMsg("Order ���۽� �����߻� - " & Err.Description)
''    End If
''End Sub


'
'   ȯ�� Order ����
'
Private Sub SendOrder_Centaur_UnPacked()
    On Error GoTo Err_Rtn

    Dim sSendBuff   As String
    Dim iCnt    As Integer
    Dim ChkSum  As String
    Dim sActionCd   As String
    
    Dim tmpBarCd    As String   '��񿡼� ���� ���ڵ� �ӽ� ����
    Dim tmpSpcNo    As String   '���ü��ȣ �ӽ� ����
    Dim iSndBufCnt  As Integer
    Dim iIdx        As Integer
    
    Select Case m_iSendPhase
        Case 1
            'Header Record
            'sSendBuff = "H|\^&|||Host|||||ACCP1||P|1|" & Chr(13) & Chr(3)
            sSendBuff = "H|\^&|||" & msReceiver & "|||||" & msSender & "||P|1|" & Chr(13) & Chr(3)
            
            tmpBarCd = maSpcNo(miIndex)
            miIndex = miIndex + 1
            
            pSampleInfo.ID = tmpBarCd
            
            '----- �˻��׸� ��ȸ
            RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)

            Call Get_OrderString
            
            '��񿡼� ���� �Ѿ�� ���ڵ�� ����
            If pSampleInfo.ID <> "" Then
                pSampleInfo.ID = tmpBarCd
            End If

            If pSampleInfo.ORDCNT = 0 Then
                m_iSendPhase = 5
            Else
                m_iSendPhase = 2
            End If
            
        Case 2
            'Patient Record
            '3(2):Practice Assigned Patient ID
            '6(5):Patient Name
            '14(13):Physician
            sSendBuff = "P|1" & Chr(13) & Chr(3)
'            sSendBuff = sSendBuff & "P|1|" & Trim(Left(pSampleInfo.PATINFO, 11)) & "|||" & _
'                            Trim(Left(pSampleInfo.SAMPINFO, 30)) & "||||||||" & tmpSpcNo & Chr(13) & Chr(3)

            m_iSendPhase = 3
            
        Case 3
            'Order Record
            sSendBuff = "O|1|" & Trim(pSampleInfo.ID) & "||"
            
            '�˻��׸� Order�ڵ� �߰�
            For iCnt = 1 To pSampleInfo.ORDCNT
                '���� ����
                sSendBuff = sSendBuff & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "\"
            Next iCnt
            sSendBuff = Left(sSendBuff, Len(sSendBuff) - 1)      '"\" Cutting

            sSendBuff = sSendBuff & "|R||||||" & sActionCd & "||||||||||||||O\Q"
                                    
            If Len(sSendBuff) > 240 Then
                iSndBufCnt = Len(sSendBuff) / 240
                
                ReDim maSendBuf(iSndBufCnt)
                
                iIdx = 1
                For iCnt = 0 To iSndBufCnt
                    If iCnt = iSndBufCnt Then
                        maSendBuf(iCnt) = Mid(sSendBuff, iIdx, 240) & Chr(13) & Chr(3)
                    Else
                        maSendBuf(iCnt) = Mid(sSendBuff, iIdx, 240) & Chr(23)
                    End If
                    
                    iIdx = iIdx + 240
                Next
                
                sSendBuff = maSendBuf(miSendCnt)
                miSendCnt = miSendCnt + 1
                
                m_iSendPhase = 4
            Else
                sSendBuff = sSendBuff & Chr(13) & Chr(3)
                m_iSendPhase = 5
            End If
                        
        Case 4
            sSendBuff = maSendBuf(miSendCnt)

            If UBound(maSendBuf) = miSendCnt Then
                m_iSendPhase = 5
            Else
                m_iSendPhase = 4
            End If
            
        Case 5
            'Terminator Record (N:Normal, F:Final, I:No Information Available)
            If pSampleInfo.ORDCNT = 0 Then
                sSendBuff = "L|1|I" & Chr(13) & Chr(3)
            Else
                sSendBuff = "L|1|F" & Chr(13) & Chr(3)
            End If
            
            miSendCnt = 0
            m_iSendPhase = 6
            
        Case 6      'EOT
            msComm.Output = Chr(4)   'EOT
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(4))
            End If
            
            If UBound(maSpcNo) >= miIndex Then
                m_iFrameN = 1: m_iPhase = 3: m_iSendPhase = 1: sState = "Q"
                msComm.Output = Chr(5)   'ENQ
                
                If m_sTestMode = "77" Then
                    RaiseEvent PrintSendLog(Chr(5))
                End If
            Else
                m_iFrameN = 1: m_iPhase = 1: m_iSendPhase = 1: sState = "": miIndex = 0
            End If
            
            '���۵� ������ �ִ� ��� ȭ��ǥ��
            If pSampleInfo.ORDCNT > 0 And sReqStatusCd = "O" Then
                If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
                    RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
                End If
            Else
                '��ȸ�� ������ ���� ��� ȯ������ ����ü �ʱ�ȭ
                Call Init_pResultInfo
                RaiseEvent SendOrderOK("", "", "", "")
            End If

            Exit Sub
    End Select

    ChkSum = ChkSum_ASTM(sSendBuff)
    sSendBuff = m_iFrameN & sSendBuff & ChkSum
    msComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)
    
    m_iFrameN = m_iFrameN + 1
    
    If m_iFrameN > 7 Then
        m_iFrameN = 0
    End If

    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If

Err_Rtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order ���۽� �����߻� - " & Err.Description)
    End If
End Sub

Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .SEQNO = m_p_sSeq
            .RACK = m_p_sRack
            .POS = m_p_sPos
            .ORDCNT = 0
        End With
    
        Exit Sub
    End If
    
    ReDim tmpData(m_p_iOrdCnt) As String
    tmpData() = Split(m_p_sTIFCd, Chr(124))
    
    With pSampleInfo
        .ID = m_p_sID
        .SEQNO = m_p_sSeq
        .RACK = m_p_sRack
        .POS = m_p_sPos
        .ORDCNT = m_p_iOrdCnt
        
        ReDim .IFCD(.ORDCNT)
        iCnt = 0
        For ii = 1 To .ORDCNT
            If Trim(tmpData(ii - 1)) <> "" Then
                iCnt = iCnt + 1
                .IFCD(iCnt) = tmpData(ii - 1)
            End If
        Next ii
        .ORDCNT = iCnt      '���� �˻� ������ �׸� ����
    End With
        
End Sub


'
'   ������� ����ü �ʱ�ȭ
'
Private Sub Init_pResultInfo()
    
    With pResultInfo
        .ID = ""
        .SEQNO = ""
        .RACK = ""
        .POS = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
        .RSTDT = ""
    End With
    
End Sub
'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=msComm,msComm,-1,RTSEnable
Public Property Get RTSEnable() As Boolean
Attribute RTSEnable.VB_Description = "���� ��û ���� ���������� ���θ� �����մϴ�."
    RTSEnable = msComm.RTSEnable
End Property

Public Property Let RTSEnable(ByVal New_RTSEnable As Boolean)
    msComm.RTSEnable() = New_RTSEnable
    PropertyChanged "RTSEnable"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=msComm,msComm,-1,RThreshold
Public Property Get RThreshold() As Integer
Attribute RThreshold.VB_Description = "������ ������ ���� ��ȯ�ϰų� �����մϴ�."
    RThreshold = msComm.RThreshold
End Property

Public Property Let RThreshold(ByVal New_RThreshold As Integer)
    msComm.RThreshold() = New_RThreshold
    PropertyChanged "RThreshold"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=msComm,msComm,-1,Settings
Public Property Get Settings() As String
Attribute Settings.VB_Description = "���� �ӵ�, �и�Ƽ, ������ ��Ʈ, �ߴ� ��Ʈ �Ű� ������ ��ȯ�ϰų� �����մϴ�."
    Settings = msComm.Settings
End Property

Public Property Let Settings(ByVal New_Settings As String)
    msComm.Settings() = New_Settings
    PropertyChanged "Settings"
End Property

Private Sub cmdTest_Click()

    wkBuf = Text1
    Call PhaseCfg_Protocol

End Sub

Private Sub msComm_OnComm()
        
    Select Case msComm.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            wkBuf = msComm.Input
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintRcvLog(wkBuf)
            End If
                                
            If iSpaceCnt = 30 Then
                iSpaceCnt = 0
            End If
            iSpaceCnt = iSpaceCnt + 2
            
            RaiseEvent DispMsg(Space(iSpaceCnt) & "���� Interface �۾� ��...")
            
            Call PhaseCfg_Protocol
            
        Case MSCOMM_EV_CTS      'j
        Case MSCOMM_EV_DSR      ' Change in the DSR line.
        Case MSCOMM_EV_CD       ' Change in the CD line.
        Case MSCOMM_EV_RING     ' Change in the Ring Indicator.
        ' Errors
        Case MSCOMM_ER_BREAK    ' A Break was received.
        ' Code to handle a BREAK goes here, and so on.
        Case MSCOMM_ER_CTSTO    ' CTS Timeout.
        Case MSCOMM_ER_DSRTO    ' DSR Timeout.
        Case MSCOMM_ER_FRAME    ' Framing Error.
        Case MSCOMM_ER_OVERRUN  ' Data Lost.
        Case MSCOMM_ER_CDTO     ' CD (RLSD) Timeout.
        Case MSCOMM_ER_RXOVER   ' Receive buffer overflow.
        Case MSCOMM_ER_RXPARITY ' Parity Error.
        Case MSCOMM_ER_TXFULL   ' Transmit buffer full.
    End Select
    
End Sub
'����ҿ��� �Ӽ����� �ε��մϴ�.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    msComm.CommPort = PropBag.ReadProperty("CommPort", 1)
    msComm.RTSEnable = PropBag.ReadProperty("RTSEnable", False)
    msComm.RThreshold = PropBag.ReadProperty("RThreshold", 0)
    msComm.Settings = PropBag.ReadProperty("Settings", "9600,n,8,1")
    m_PortOpen = PropBag.ReadProperty("PortOpen", m_def_PortOpen)
    m_OpenPW = PropBag.ReadProperty("OpenPW", m_def_OpenPW)
    m_EditPW = PropBag.ReadProperty("EditPW", m_def_EditPW)
    m_EqName = PropBag.ReadProperty("EqName", m_def_EqName)
    m_bUseBarcode = PropBag.ReadProperty("bUseBarcode", m_def_bUseBarcode)
    m_iPhase = PropBag.ReadProperty("iPhase", m_def_iPhase)
    m_iSendPhase = PropBag.ReadProperty("iSendPhase", m_def_iSendPhase)
    m_sTestMode = PropBag.ReadProperty("sTestMode", m_def_sTestMode)
    m_iFrameN = PropBag.ReadProperty("iFrameN", m_def_iFrameN)
    m_p_sID = PropBag.ReadProperty("p_sID", m_def_p_sID)
    m_p_sSeq = PropBag.ReadProperty("p_sSeq", m_def_p_sSeq)
    m_p_sRack = PropBag.ReadProperty("p_sRack", m_def_p_sRack)
    m_p_sPos = PropBag.ReadProperty("p_sPos", m_def_p_sPos)
    m_p_iOrdCnt = PropBag.ReadProperty("p_iOrdCnt", m_def_p_iOrdCnt)
    m_p_sTIFCd = PropBag.ReadProperty("p_sTIFCd", m_def_p_sTIFCd)
End Sub

'�Ӽ����� ����ҿ� ����մϴ�.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("CommPort", msComm.CommPort, 1)
    Call PropBag.WriteProperty("RTSEnable", msComm.RTSEnable, False)
    Call PropBag.WriteProperty("RThreshold", msComm.RThreshold, 0)
    Call PropBag.WriteProperty("Settings", msComm.Settings, "9600,n,8,1")
    Call PropBag.WriteProperty("PortOpen", m_PortOpen, m_def_PortOpen)
    Call PropBag.WriteProperty("OpenPW", m_OpenPW, m_def_OpenPW)
    Call PropBag.WriteProperty("EditPW", m_EditPW, m_def_EditPW)
    Call PropBag.WriteProperty("EqName", m_EqName, m_def_EqName)
    Call PropBag.WriteProperty("bUseBarcode", m_bUseBarcode, m_def_bUseBarcode)
    Call PropBag.WriteProperty("iPhase", m_iPhase, m_def_iPhase)
    Call PropBag.WriteProperty("iSendPhase", m_iSendPhase, m_def_iSendPhase)
    Call PropBag.WriteProperty("sTestMode", m_sTestMode, m_def_sTestMode)
    Call PropBag.WriteProperty("iFrameN", m_iFrameN, m_def_iFrameN)
    Call PropBag.WriteProperty("p_sID", m_p_sID, m_def_p_sID)
    Call PropBag.WriteProperty("p_sSeq", m_p_sSeq, m_def_p_sSeq)
    Call PropBag.WriteProperty("p_sRack", m_p_sRack, m_def_p_sRack)
    Call PropBag.WriteProperty("p_sPos", m_p_sPos, m_def_p_sPos)
    Call PropBag.WriteProperty("p_iOrdCnt", m_p_iOrdCnt, m_def_p_iOrdCnt)
    Call PropBag.WriteProperty("p_sTIFCd", m_p_sTIFCd, m_def_p_sTIFCd)
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=0,0,0,0
Public Property Get PortOpen() As Boolean
    PortOpen = m_PortOpen
End Property

Public Property Let PortOpen(ByVal New_PortOpen As Boolean)
    m_PortOpen = New_PortOpen
    PropertyChanged "PortOpen"
    
    '--- PortOpen�� ��ȣ Ȯ��
    If m_OpenPW <> pOpenPW Then
        MsgBox "��ϵ� ����ڰ� �ƴմϴ�. (��)���̾����̷� ������ �ֽʽÿ�!!!", vbCritical, "����� Ȯ��"
        Exit Property
    End If
    '-----------------------
    
    '���� �ʱ�ȭ(E-170/H-7600)
    RstEnd = "Y": bEndChk = True: bSTXChk = False
    
    
    On Error GoTo ErrPortOpen
    If m_PortOpen = True Then
        msComm.PortOpen = True
    End If
    On Error GoTo 0
ErrPortOpen:
    If Err <> 0 Then
        MsgBox "PortOpen Error!!! " & Err.Description, vbCritical
        RaiseEvent DispMsg(Err.Description)
    End If
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get OpenPW() As String
    OpenPW = m_OpenPW
End Property

Public Property Let OpenPW(ByVal New_OpenPW As String)
    m_OpenPW = New_OpenPW
    PropertyChanged "OpenPW"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get EditPW() As String
    EditPW = m_EditPW
End Property

Public Property Let EditPW(ByVal New_EditPW As String)
    m_EditPW = New_EditPW
    PropertyChanged "EditPW"
End Property

'����� ���� ��Ʈ�ѿ� ���� �Ӽ��� �ʱ�ȭ�մϴ�.
Private Sub UserControl_InitProperties()
    m_PortOpen = m_def_PortOpen
    m_OpenPW = m_def_OpenPW
    m_EditPW = m_def_EditPW
    m_EqName = m_def_EqName
    m_bUseBarcode = m_def_bUseBarcode
    m_iPhase = m_def_iPhase
    m_iSendPhase = m_def_iSendPhase
    m_sTestMode = m_def_sTestMode
    m_iFrameN = m_def_iFrameN
    m_p_sID = m_def_p_sID
    m_p_sSeq = m_def_p_sSeq
    m_p_sRack = m_def_p_sRack
    m_p_sPos = m_def_p_sPos
    m_p_iOrdCnt = m_def_p_iOrdCnt
    m_p_sTIFCd = m_def_p_sTIFCd
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get EqName() As String
    EqName = m_EqName
End Property

Public Property Let EqName(ByVal New_EqName As String)
    m_EqName = New_EqName
    PropertyChanged "EqName"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=0,0,0,0
Public Property Get bUseBarcode() As Boolean
    bUseBarcode = m_bUseBarcode
End Property

Public Property Let bUseBarcode(ByVal New_bUseBarcode As Boolean)
    m_bUseBarcode = New_bUseBarcode
    PropertyChanged "bUseBarcode"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,0
Public Property Get iPhase() As Integer
    iPhase = m_iPhase
End Property

Public Property Let iPhase(ByVal New_iPhase As Integer)
    m_iPhase = New_iPhase
    PropertyChanged "iPhase"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,0
Public Property Get iSendPhase() As Integer
    iSendPhase = m_iSendPhase
End Property

Public Property Let iSendPhase(ByVal New_iSendPhase As Integer)
    m_iSendPhase = New_iSendPhase
    PropertyChanged "iSendPhase"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get sTestMode() As String
    sTestMode = m_sTestMode
End Property

Public Property Let sTestMode(ByVal New_sTestMode As String)
    m_sTestMode = New_sTestMode
    PropertyChanged "sTestMode"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,0
Public Property Get iFrameN() As Integer
    iFrameN = m_iFrameN
End Property

Public Property Let iFrameN(ByVal New_iFrameN As Integer)
    m_iFrameN = New_iFrameN
    PropertyChanged "iFrameN"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get p_sID() As String
    p_sID = m_p_sID
End Property

Public Property Let p_sID(ByVal New_p_sID As String)
    m_p_sID = New_p_sID
    PropertyChanged "p_sID"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get p_sSeq() As String
    p_sSeq = m_p_sSeq
End Property

Public Property Let p_sSeq(ByVal New_p_sSeq As String)
    m_p_sSeq = New_p_sSeq
    PropertyChanged "p_sSeq"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get p_sRack() As String
    p_sRack = m_p_sRack
End Property

Public Property Let p_sRack(ByVal New_p_sRack As String)
    m_p_sRack = New_p_sRack
    PropertyChanged "p_sRack"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get p_sPos() As String
    p_sPos = m_p_sPos
End Property

Public Property Let p_sPos(ByVal New_p_sPos As String)
    m_p_sPos = New_p_sPos
    PropertyChanged "p_sPos"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,0
Public Property Get p_iOrdCnt() As Integer
    p_iOrdCnt = m_p_iOrdCnt
End Property

Public Property Let p_iOrdCnt(ByVal New_p_iOrdCnt As Integer)
    m_p_iOrdCnt = New_p_iOrdCnt
    PropertyChanged "p_iOrdCnt"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=13,0,0,0
Public Property Get p_sTIFCd() As String
    p_sTIFCd = m_p_sTIFCd
End Property

Public Property Let p_sTIFCd(ByVal New_p_sTIFCd As String)
    m_p_sTIFCd = New_p_sTIFCd
    PropertyChanged "p_sTIFCd"
End Property
'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=14
Public Function Send_Chr(iChr%) As Variant
    On Error GoTo ErrComm
    msComm.Output = Chr(iChr)
    On Error GoTo 0
ErrComm:
    If Err <> 0 Then
        RaiseEvent DispMsg("Send_Chr ���� - " & Err.Description)
    End If
End Function
