VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl ACUSTAR 
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
      ScrollBars      =   2  '수직
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
Attribute VB_Name = "ACUSTAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_p_sPatInfo = 0
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
'속성 변수:
Dim m_p_sPatInfo As Variant
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
'이벤트 선언:
Event SendOrderOK(sID$, sRack$, sPos$, iOrdCnt%)
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sQCGbn$, sINSTID$, sRstDt$, sOTHER$)
Event DispMsgComm(sMsg$)
'Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sQCGbn$)
'Event SendOrderOK(sID$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event RequestCurOrder(sID$, sRack$, sPos$)
Event DispMsg(sMsg$)

'===== User Define
'인터페이스에서 사용
Dim RcvBuffer   As String
Dim wkBuf   As String
Dim sState  As String
Dim sReqStatusCd    As String

'구조체 지정
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'기타
Dim iSpaceCnt   As Integer

'For E-170/Hitachi7600
Dim bEndChk As Boolean
Dim bSTXChk As Boolean
Dim sNextSend   As String
Dim RstEnd      As String

Dim msMsgID    As String
Dim msSender   As String
Dim msReceiver As String
Dim msVersion  As String

Dim maSpcNo() As String
Dim maEqSeq() As String
Dim miOrdIdx  As Integer

''Private Sub SendOrder_ACUSTAR()
''    On Error GoTo Err_Rtn
''
''    Dim sSendBuff   As String
''    Dim iCnt    As Integer
''    Dim i       As Integer
''    Dim ChkSum  As String
''    Dim sStat   As String
''    Dim sTOrder As String
''    Dim sPriority As String
''    Dim sSpcType As String
''    Dim iPatCnt As Integer
''
''    Select Case m_iSendPhase
''        Case 1
''            '----- 검사항목 조회
''            'Header Record
''            ''sSendBuff = m_iFrameN & "H|@^\|<0_0><1025080549_50>||" & msReceiver & "|||||" & msSender & "||P|" & msVersion & "|" & Format(Now, "yyyyMMddHHmmss") & Chr(13)
''            sSendBuff = "H|@^\|" & msMsgID & "||" & msReceiver & "|||||" & msSender & "||P|" & msVersion & "|" & Format(Now, "yyyyMMddHHmmss") & Chr(13)
''
''            pSampleInfo.ID = Trim(maSpcNo(miOrdIdx))
''            RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
''
''            miOrdIdx = miOrdIdx + 1
''
''            Call Get_OrderString
''
''            If pSampleInfo.ORDCNT > 0 Then
''                iPatCnt = iPatCnt + 1
''                'Patient Record
''                'sSendBuff = sSendBuff & "P|" & CStr(iPatCnt) & Chr(13)
''                sSendBuff = sSendBuff & "P|" & CStr(iPatCnt) & "||||^||||||||" & Chr(13)
''
''                'Order Record
''                sSendBuff = sSendBuff & "O|1|" & Left(Trim(pSampleInfo.ID), 16) & "|" & maEqSeq(i) & "|"
''
''                '검사항목 Order코드 추가
''                sTOrder = ""
''                For iCnt = 1 To pSampleInfo.ORDCNT
''                    '일반항목
''                    sTOrder = sTOrder & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "@"
''                Next iCnt
''                sTOrder = Left(sTOrder, Len(sTOrder) - 1)      '"@" Cutting
''
''                sSendBuff = sSendBuff & sTOrder & "|"
''
''                'S(Stat), R(normal)
''                If sPriority = "" Then
''                    sPriority = "R"
''                Else
''                End If
''
''                If pSampleInfo.SPCCD = "" Then
''                    sSpcType = "SER"    'Serum
''                Else
''                    sSpcType = pSampleInfo.SPCCD
''                End If
''
''                ''sSendBuff = sSendBuff & sPriority & "|" & Format(Now, "yyyyMMddHHmmss") & "|||||P||||" & sSpcType & "||||||||||O@I" & Chr(13)
''                sSendBuff = sSendBuff & sPriority & "||||||A||||" & sSpcType & "||||||||||Q" & Chr(13)
''
''                sSendBuff = sSendBuff & "L|1|F"
''            Else
''                sSendBuff = sSendBuff & "L|1|I"
''            End If
''
''            '--- Text의 내용이 240byte를 넘어갈 경우 처리 추가...
''            If Len(sSendBuff) >= 240 Then
''                sNextSend = Mid(sSendBuff, 241)
''                sSendBuff = Left(sSendBuff, 240)
''                sSendBuff = sSendBuff & Chr(23)
''
''                m_iSendPhase = 2
''            Else
''                sSendBuff = sSendBuff & Chr(13) & Chr(3)
''                 m_iSendPhase = 3
''            End If
''
''        Case 2
''            sSendBuff = sNextSend
''            sNextSend = ""
''
''            If Len(sSendBuff) >= 240 Then
''                sNextSend = Mid(sSendBuff, 241)
''                sSendBuff = Left(sSendBuff, 240)
''                sSendBuff = sSendBuff & Chr(23)
''
''                m_iSendPhase = 2
''            Else
''                sSendBuff = sSendBuff & Chr(13) & Chr(3)
''                m_iSendPhase = 3
''            End If
''
''        Case 3      'EOT
''            msComm.Output = Chr(4)   'EOT
''
''            If m_sTestMode = "77" Then
''                RaiseEvent PrintSendLog(Chr(4))
''            End If
''
''            If UBound(maSpcNo) < miOrdIdx Then
''                m_iFrameN = 1
''                m_iPhase = 1
''                m_iSendPhase = 1
''
''                sState = "": sReqStatusCd = "": miOrdIdx = 0
''            Else
''                m_iFrameN = 1
''                m_iSendPhase = 1
''
''                msComm.Output = Chr(5)
''
''                If m_sTestMode = "77" Then
''                    RaiseEvent PrintSendLog(Chr(5))
''                End If
''            End If
''
''            Exit Sub
''    End Select
''
''    sSendBuff = m_iFrameN & sSendBuff
''
''    ChkSum = ChkSum_ASTM(sSendBuff)
''    sSendBuff = sSendBuff & ChkSum
''    msComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)
''
''    If m_sTestMode = "77" Then
''        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
''    End If
''
''    m_iFrameN = m_iFrameN + 1
''
''    If m_iFrameN > 7 Then      'Frame Number가 8이상이면 0으로 바꿔줌
''        m_iFrameN = 0
''    End If
''
''    '전송된 오더가 있는 경우 화면표시
''    If pSampleInfo.ORDCNT > 0 Then
''        If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
''            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, pSampleInfo.ORDCNT)
''        End If
''    Else
''        '조회된 내용이 없는 경우 환자정보 구조체 초기화
''        Call Init_pResultInfo
''
''        RaiseEvent SendOrderOK("", "", "", 0)
''    End If
''
''Err_Rtn:
''    If Err <> 0 Then
''        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
''    End If
''End Sub

Private Sub SendOrder_ACUSTAR()
    On Error GoTo Err_Rtn

    Dim sSendBuff   As String
    Dim iCnt    As Integer
    Dim i       As Integer
    Dim ChkSum  As String
    Dim sStat   As String
    Dim sTOrder As String
    Dim sPriority As String
    Dim sSpcType As String
    Dim iPatCnt As Integer

    Select Case m_iSendPhase
        Case 1
            '----- 검사항목 조회
            'Header Record
            ''sSendBuff = m_iFrameN & "H|@^\|<0_0><1025080549_50>||" & msReceiver & "|||||" & msSender & "||P|" & msVersion & "|" & Format(Now, "yyyyMMddHHmmss") & Chr(13)
            sSendBuff = "H|@^\|" & msMsgID & "||" & msReceiver & "|||||" & msSender & "||P|" & msVersion & "|" & Format(Now, "yyyyMMddHHmmss") & Chr(13)

            If pSampleInfo.ID <> "ALL" Then
                For i = 0 To UBound(maSpcNo)
                    pSampleInfo.ID = Trim(maSpcNo(i))
                    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
    
                    Call Get_OrderString
    
                    If pSampleInfo.ORDCNT > 0 Then
                        iPatCnt = iPatCnt + 1
                        'Patient Record
                        'sSendBuff = sSendBuff & "P|" & CStr(iPatCnt) & Chr(13)
                        sSendBuff = sSendBuff & "P|" & CStr(iPatCnt) & "||||^||||||||" & Chr(13)
    
                        'Order Record
                        sSendBuff = sSendBuff & "O|1|" & Left(Trim(pSampleInfo.ID), 16) & "|" & maEqSeq(i) & "|"
    
                        '검사항목 Order코드 추가
                        sTOrder = ""
                        For iCnt = 1 To pSampleInfo.ORDCNT
                            '일반항목
                            sTOrder = sTOrder & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "@"
                        Next iCnt
                        sTOrder = Left(sTOrder, Len(sTOrder) - 1)      '"@" Cutting
    
                        sSendBuff = sSendBuff & sTOrder & "|"
    
                        'S(Stat), R(normal)
                        If sPriority = "" Then
                            sPriority = "R"
                        Else
                        End If
    
                        If pSampleInfo.SPCCD = "" Then
                            sSpcType = "SER"    'Serum
                        Else
                            sSpcType = pSampleInfo.SPCCD
                        End If
    
                        sSendBuff = sSendBuff & sPriority & "||||||A||||" & sSpcType & "||||||||||Q" & Chr(13)
                        
                        ''RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, pSampleInfo.ORDCNT)
                    End If
                Next
            End If

            'Terminator Record
            If iPatCnt > 0 Then
                sSendBuff = sSendBuff & "L|1|F"
            Else
                sSendBuff = sSendBuff & "L|1|I"
            End If

            '--- Text의 내용이 240byte를 넘어갈 경우 처리 추가...
            If Len(sSendBuff) >= 240 Then
                sNextSend = Mid(sSendBuff, 241)
                sSendBuff = Left(sSendBuff, 240)
                sSendBuff = sSendBuff & Chr(23)

                m_iSendPhase = 2
            Else
                sSendBuff = sSendBuff & Chr(13) & Chr(3)
                m_iSendPhase = 3
            End If

        Case 2
            sSendBuff = sNextSend
            sNextSend = ""

            If Len(sSendBuff) >= 240 Then
                sNextSend = Mid(sSendBuff, 241)
                sSendBuff = Left(sSendBuff, 240)
                sSendBuff = sSendBuff & Chr(23)

                m_iSendPhase = 2
            Else
                sSendBuff = sSendBuff & Chr(13) & Chr(3)
                m_iSendPhase = 3
            End If

        Case 3      'EOT
            msComm.Output = Chr(4)   'EOT

            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(Chr(4))
            End If

            m_iFrameN = 1
            m_iPhase = 1
            m_iSendPhase = 1

            sState = "": sReqStatusCd = ""

            Exit Sub
    End Select

    sSendBuff = m_iFrameN & sSendBuff

    ChkSum = ChkSum_ASTM(sSendBuff)
    sSendBuff = sSendBuff & ChkSum
    msComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)

    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If

     m_iFrameN = m_iFrameN + 1

    If m_iFrameN > 7 Then      'Frame Number가 8이상이면 0으로 바꿔줌
        m_iFrameN = 0
    End If

''    '전송된 오더가 있는 경우 화면표시
''    If pSampleInfo.ORDCNT > 0 Then
''        If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
''            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, pSampleInfo.ORDCNT)
''        End If
''    Else
''        '조회된 내용이 없는 경우 환자정보 구조체 초기화
''        Call Init_pResultInfo
''
''        RaiseEvent SendOrderOK("", "", "", 0)
''    End If

Err_Rtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub

''Private Sub SendOrder_ACUSTAR()
''    On Error GoTo Err_Rtn
''
''    Dim sSendBuff   As String
''    Dim iCnt    As Integer
''    Dim i       As Integer
''    Dim ChkSum  As String
''    Dim sStat   As String
''    Dim sTOrder As String
''    Dim sPriority As String
''    Dim sSpcType As String
''    Dim iPatCnt As Integer
''
''    Select Case m_iSendPhase
''        Case 1
''            '----- 검사항목 조회
''            'Header Record
''            ''sSendBuff = m_iFrameN & "H|@^\|<0_0><1025080549_50>||" & msReceiver & "|||||" & msSender & "||P|" & msVersion & "|" & Format(Now, "yyyyMMddHHmmss") & Chr(13)
''            sSendBuff = m_iFrameN & "H|@^\|" & msMsgID & "||" & msReceiver & "|||||" & msSender & "||P|" & msVersion & "|" & Format(Now, "yyyyMMddHHmmss") & Chr(13)
''
''            For i = 0 To UBound(maSpcNo)
''                pSampleInfo.ID = Trim(maSpcNo(i))
''                RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
''
''                Call Get_OrderString
''
''                If pSampleInfo.ORDCNT > 0 Then
''                    iPatCnt = iPatCnt + 1
''                    'Patient Record
''                    'sSendBuff = sSendBuff & "P|" & CStr(iPatCnt) & Chr(13)
''                    sSendBuff = sSendBuff & "P|1||||^||||||||" & Chr(13)
''
''                    'S(Stat), R(normal)
''                    If sPriority = "" Then
''                        sPriority = "R"
''                    Else
''                    End If
''
''                    If pSampleInfo.SPCCD = "" Then
''                        sSpcType = "SER"    'Serum
''                    Else
''                        sSpcType = pSampleInfo.SPCCD
''                    End If
''
''                    '검사항목 Order코드 추가
''                    For iCnt = 1 To pSampleInfo.ORDCNT
''                        'Order Record
''                        sSendBuff = sSendBuff & "O|" & CStr(iCnt) & "|" & Left(Trim(pSampleInfo.ID), 16) & "|" & maEqSeq(i) & "|" & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "|"
''                        sSendBuff = sSendBuff & sPriority & "||||||A||||" & sSpcType & "||||||||||Q" & Chr(13)
''                    Next iCnt
''                End If
''            Next
''
''            'Terminator Record
''            If iPatCnt > 0 Then
''                sSendBuff = sSendBuff & "L|1|F"
''            Else
''                sSendBuff = sSendBuff & "L|1|I"
''            End If
''
''            '--- Text의 내용이 240byte를 넘어갈 경우 처리 추가...
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
''            sSendBuff = sNextSend
''            sNextSend = ""
''
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
''
''            sState = "": sReqStatusCd = ""
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
''    '전송된 오더가 있는 경우 화면표시
''    If pSampleInfo.ORDCNT > 0 Then
''        If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
''            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, pSampleInfo.ORDCNT)
''        End If
''    Else
''        '조회된 내용이 없는 경우 환자정보 구조체 초기화
''        Call Init_pResultInfo
''
''        RaiseEvent SendOrderOK("", "", "", 0)
''    End If
''
''Err_Rtn:
''    If Err <> 0 Then
''        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
''    End If
''End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,CommPort
Public Property Get CommPort() As Integer
Attribute CommPort.VB_Description = "통신 포트 번호를 반환하거나 설정합니다."
    CommPort = msComm.CommPort
End Property

Public Property Let CommPort(ByVal New_CommPort As Integer)
    msComm.CommPort() = New_CommPort
    PropertyChanged "CommPort"
End Property

Private Sub PhaseCfg_Protocol()

    '--- 사용자 확인
    If m_EditPW <> pEditPW Then
        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
        Exit Sub
    End If
    '---------------
    
    If m_EqName = "0" Or m_EqName = "" Then
        RaiseEvent DispMsg("검사장비명을 지정해 주십시오.!!!")
        Exit Sub
    End If
    
    Select Case UCase(m_EqName)
        Case "ACUSTAR"
            Call PhaseCfg_Protocol_ACUSTAR
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub

Private Sub PhaseCfg_Protocol_ACUSTAR()
    On Error GoTo ErrRtn
    
    Dim wkDat   As String
    Dim ix1 As Integer
    Dim i   As Integer

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
                            Call DataEditResponse_ACUSTAR
                            RcvBuffer = ""
                        End If
                        
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If

                    Case 13     'CR
                        If bEndChk = True Then
                            Call DataEditResponse_ACUSTAR
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
                            Call SendOrder_ACUSTAR
                        End If

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        msComm.Output = Chr(6)
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(Chr(6))
                        End If
                        
                        m_iPhase = 2

                    Case 21     'NAK

                    Case 4      'EOT
                        m_iPhase = 1

                End Select
        End Select
    Next ix1

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
    End If
End Sub
' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_ACUSTAR()
    On Error GoTo ErrRtn

    Dim RecType     As String       'Record Type
    Dim sResType    As String
    Dim sResState   As String
    Dim sResData    As String

    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$, tmpQCGbn$
    Dim tmpIFCd$, tmpRst$, tmpRst2$, tmpUnit$, tmpRef$, tmpFlag$
    Dim sPriority As String
    Dim sSpcType As String
    Dim sRstState As String
    Dim sRstDt As String
    Dim sEqCd As String
    Dim sCmtCd As String
    Dim sCmtNm As String
    
    Dim sTmp As String
    Dim i As Integer
    Dim ii As Integer
    Dim aRow()  As String
    Dim tmpField()  As String
    Dim tmpData()   As String
    
    ii = InStr(1, RcvBuffer, "|")
    If ii <> 0 Then
        RecType = Mid$(RcvBuffer, ii - 1, 1)
    Else
        Exit Sub
    End If
    
    If InStr(1, RcvBuffer, Chr(13)) > 0 Then        '2007/6/22 yk
        aRow() = Split(RcvBuffer, Chr(13))
        RcvBuffer = aRow(0)
    End If

    Select Case RecType
        Case "H"        'Header Record
            '1H|@^\|081C0C75-B3D7-4C4E-93DC-EF20E446F11E||Alba11110228|||||LIS_001||P|1394-97|20120901165511
            Call Init_pResultInfo
            
            tmpData() = Split(RcvBuffer, Chr(124))
        
            msMsgID = Trim(tmpData(2))
            msSender = Trim(tmpData(4))
            msReceiver = Trim(tmpData(9))
            msVersion = Trim(tmpData(12))

        Case "P"        'Patient Record
            'P|1||||^||19700101|U|||||
            Call Init_pResultInfo

        Case "Q"        'Order Request Record
            'Q|1|^13020100773@^13013116006@^13020100789@^13020100774@^13020107860@^13020100786||||||||||O@N
            tmpData() = Split(RcvBuffer, "|")
            
            sTmp = Trim(tmpData(2))
            
            pSampleInfo.ID = sTmp
            
            If sTmp <> "ALL" Then
                tmpField = Split(tmpData(2), "@")
                
                For i = 0 To UBound(tmpField)
                    If Trim(tmpField(i)) <> "" Then
                        ReDim Preserve maEqSeq(i)
                        ReDim Preserve maSpcNo(i)
                        
                        If UBound(Split(tmpField(i), "^")) > 1 Then
                            maSpcNo(i) = Trim(Split(tmpField(i), "^")(1))
                            maEqSeq(i) = Trim(Split(tmpField(i), "^")(2))
                        Else
                            maSpcNo(i) = Trim(Split(tmpField(i), "^")(1))
                            maEqSeq(i) = ""
                        End If
                    End If
                Next
            End If
            
            sReqStatusCd = Trim(tmpData(12))    'Order Request Status Code
            
            If sReqStatusCd <> "A" Then
                sState = "Q"
                miOrdIdx = 0
            Else
                sState = ""
            End If

        Case "O"        'Order Record
            'O|1|13020100786|19948|^^^2269|R|20120901143900|||||||||SER||||||||||O@F
            'O|1|13020413329|20128|^^^aCL_IgG|R||||||P||||SER||||||||||O@I
            tmpData() = Split(RcvBuffer, "|")
            'BarCode
            tmpBarCd = Trim(tmpData(2))
            tmpSeqNo = Trim(tmpData(3))
            sPriority = Trim(tmpData(5))

            '일반/QC 결과 구분
            'Q (mandatory when quality control)
            tmpQCGbn = Trim(tmpData(11))
            
            sSpcType = Trim(tmpData(15))
            
            sRstState = Trim(tmpData(25))
            
            With pSampleInfo
                .ID = tmpBarCd
                .SEQNO = tmpSeqNo
                .QCGBN = tmpQCGbn
                .SPCCD = sSpcType
            End With
            
            If sRstState = "O@I" Then
                RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS, 1)
            End If

        Case "R"        'Result Record
            'R|1|^^^2269|FAILURE|U/mL||<||F@V||ACL^ACL||20120901163124|ALBA^D^5
            tmpData() = Split(RcvBuffer, "|")

            tmpIFCd = Trim(Split(tmpData(2), "^")(3))
            tmpRst = Trim(tmpData(3))
            tmpUnit = Trim(tmpData(4))
            
            'L (Below low normal)
            'H (Above high normal)
            'N (Normal)
            '< (Below absolute low)
            '> (Above absolute high)
            tmpFlag = Trim(tmpData(6))
            
            sRstState = Trim(tmpData(8))
            sRstDt = Trim(tmpData(12))
            
            tmpField = Split(tmpData(13), "^")
            
            sEqCd = Trim(tmpField(0))
            tmpRack = Trim(tmpField(1))
            tmpPos = Trim(tmpField(2))
           
            '결과정보 구조체에 저장
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = tmpRack
                .POS = tmpPos
                .QCGBN = pSampleInfo.QCGBN
                
                '결과값 누적
                .RSTCNT = 1
                .IFCD = tmpIFCd & Chr(124)
                .RST1 = tmpRst & Chr(124)
                .RST2 = tmpRst2 & Chr(124)
                .UNIT = tmpUnit & Chr(124)
                .FLAG = tmpFlag & Chr(124)
            End With

        Case "C"        'Comment Record
            'C|1|I|1025^reagenttemperaturewarning^HW|I
            'C|2|I|1030^cuvetteshuttletempwarning^HW|I
            tmpData() = Split(RcvBuffer, "|")
            
            sCmtCd = Trim(Split(tmpData(3), "^")(0))
            sCmtNm = Trim(Split(tmpData(3), "^")(1))

        Case "L"        'Msg Terminater Record
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, 1, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .QCGBN, "", sRstDt, .OTHER)
                End If
            End With

            Call Init_pResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub
Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .ORDCNT = 0
            Erase .IFCD
            .OTHER = ""     '2008/3/20 yk
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
        .ORDCNT = iCnt      '실제 검사 가능한 항목 갯수
        
        .OTHER = m_p_sPatInfo
    End With
        
End Sub

'
'   결과정보 구조체 초기화
'
Private Sub Init_pResultInfo()
    
    With pResultInfo
        .ID = ""
        .SEQNO = ""
        .RACK = ""
        .POS = ""
        .QCGBN = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
        .OTHER = ""
    End With
    
End Sub
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,RTSEnable
Public Property Get RTSEnable() As Boolean
Attribute RTSEnable.VB_Description = "전송 요청 줄이 가능한지의 여부를 결정합니다."
    RTSEnable = msComm.RTSEnable
End Property

Public Property Let RTSEnable(ByVal New_RTSEnable As Boolean)
    msComm.RTSEnable() = New_RTSEnable
    PropertyChanged "RTSEnable"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,RThreshold
Public Property Get RThreshold() As Integer
Attribute RThreshold.VB_Description = "수신할 문자의 수를 반환하거나 설정합니다."
    RThreshold = msComm.RThreshold
End Property

Public Property Let RThreshold(ByVal New_RThreshold As Integer)
    msComm.RThreshold() = New_RThreshold
    PropertyChanged "RThreshold"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MappingInfo=msComm,msComm,-1,Settings
Public Property Get Settings() As String
Attribute Settings.VB_Description = "전송 속도, 패리티, 데이터 비트, 중단 비트 매개 변수를 반환하거나 설정합니다."
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
            
            If sTestMode = "77" Then
                RaiseEvent PrintRcvLog(wkBuf)
            End If
                                
            If iSpaceCnt = 30 Then
                iSpaceCnt = 0
            End If
            iSpaceCnt = iSpaceCnt + 2
            
            RaiseEvent DispMsgComm(Space(iSpaceCnt) & "장비와 Interface 작업 중...")
            
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
'저장소에서 속성값을 로드합니다.
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
    m_p_sPatInfo = PropBag.ReadProperty("p_sPatInfo", m_def_p_sPatInfo)
End Sub

'속성값을 저장소에 기록합니다.
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
    Call PropBag.WriteProperty("p_sPatInfo", m_p_sPatInfo, m_def_p_sPatInfo)
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,0
Public Property Get PortOpen() As Boolean
    PortOpen = m_PortOpen
End Property

Public Property Let PortOpen(ByVal New_PortOpen As Boolean)
    m_PortOpen = New_PortOpen
    PropertyChanged "PortOpen"
    
    '--- PortOpen시 암호 확인
    If m_OpenPW <> pOpenPW Then
        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
        Exit Property
    End If
    '-----------------------
    
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

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get OpenPW() As String
    OpenPW = m_OpenPW
End Property

Public Property Let OpenPW(ByVal New_OpenPW As String)
    m_OpenPW = New_OpenPW
    PropertyChanged "OpenPW"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get EditPW() As String
    EditPW = m_EditPW
End Property

Public Property Let EditPW(ByVal New_EditPW As String)
    m_EditPW = New_EditPW
    PropertyChanged "EditPW"
End Property

'사용자 정의 컨트롤에 대한 속성을 초기화합니다.
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
    m_p_sPatInfo = m_def_p_sPatInfo
End Sub

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get EqName() As String
    EqName = m_EqName
End Property

Public Property Let EqName(ByVal New_EqName As String)
    m_EqName = New_EqName
    PropertyChanged "EqName"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,0
Public Property Get bUseBarcode() As Boolean
    bUseBarcode = m_bUseBarcode
End Property

Public Property Let bUseBarcode(ByVal New_bUseBarcode As Boolean)
    m_bUseBarcode = New_bUseBarcode
    PropertyChanged "bUseBarcode"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iPhase() As Integer
    iPhase = m_iPhase
End Property

Public Property Let iPhase(ByVal New_iPhase As Integer)
    m_iPhase = New_iPhase
    PropertyChanged "iPhase"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iSendPhase() As Integer
    iSendPhase = m_iSendPhase
End Property

Public Property Let iSendPhase(ByVal New_iSendPhase As Integer)
    m_iSendPhase = New_iSendPhase
    PropertyChanged "iSendPhase"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get sTestMode() As String
    sTestMode = m_sTestMode
End Property

Public Property Let sTestMode(ByVal New_sTestMode As String)
    m_sTestMode = New_sTestMode
    PropertyChanged "sTestMode"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iFrameN() As Integer
    iFrameN = m_iFrameN
End Property

Public Property Let iFrameN(ByVal New_iFrameN As Integer)
    m_iFrameN = New_iFrameN
    PropertyChanged "iFrameN"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sID() As String
    p_sID = m_p_sID
End Property

Public Property Let p_sID(ByVal New_p_sID As String)
    m_p_sID = New_p_sID
    PropertyChanged "p_sID"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sSeq() As String
    p_sSeq = m_p_sSeq
End Property

Public Property Let p_sSeq(ByVal New_p_sSeq As String)
    m_p_sSeq = New_p_sSeq
    PropertyChanged "p_sSeq"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sRack() As String
    p_sRack = m_p_sRack
End Property

Public Property Let p_sRack(ByVal New_p_sRack As String)
    m_p_sRack = New_p_sRack
    PropertyChanged "p_sRack"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sPos() As String
    p_sPos = m_p_sPos
End Property

Public Property Let p_sPos(ByVal New_p_sPos As String)
    m_p_sPos = New_p_sPos
    PropertyChanged "p_sPos"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get p_iOrdCnt() As Integer
    p_iOrdCnt = m_p_iOrdCnt
End Property

Public Property Let p_iOrdCnt(ByVal New_p_iOrdCnt As Integer)
    m_p_iOrdCnt = New_p_iOrdCnt
    PropertyChanged "p_iOrdCnt"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sTIFCd() As String
    p_sTIFCd = m_p_sTIFCd
End Property

Public Property Let p_sTIFCd(ByVal New_p_sTIFCd As String)
    m_p_sTIFCd = New_p_sTIFCd
    PropertyChanged "p_sTIFCd"
End Property
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function Send_Chr(iChr%) As Variant
    On Error GoTo ErrComm
    msComm.Output = Chr(iChr)
    On Error GoTo 0
ErrComm:
    If Err <> 0 Then
        RaiseEvent DispMsg("Send_Chr 에러 - " & Err.Description)
    End If
End Function

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get p_sPatInfo() As Variant
    p_sPatInfo = m_p_sPatInfo
End Property

Public Property Let p_sPatInfo(ByVal New_p_sPatInfo As Variant)
    m_p_sPatInfo = New_p_sPatInfo
    PropertyChanged "p_sPatInfo"
End Property

