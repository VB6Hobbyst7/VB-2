VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl VIDAS 
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   3405
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
Attribute VB_Name = "VIDAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
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
Event SendOrderOK(sID$, sSeq$, sRack$, sPos$)
'Event RequestNextOrder()
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event DispMsg(sMsg$)
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)


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

'For VIDAS
Dim iChkSumCnt  As Integer
Dim msMsgType   As String
Dim Trans_Flag  As String
Dim OrderFlag   As Integer

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
        Case "MINIVIDAS"
            Call PhaseCfg_Protocol_miniVIDAS
        
        Case "VIDAS"
            Call PhaseCfg_Protocol_VIDAS
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_VIDAS()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1          'DATA 수신
                Select Case Asc(wkDat)
                    Case 5      'ENQ 수신
                        msComm.Output = Chr(6)
                        RcvBuffer = ""
                        m_iPhase = 1

                    Case 2      'STX
                        RcvBuffer = ""

                    Case 29     'GS
                        Call DataEdit_VIDAS
                        msComm.Output = Chr(6)

                    Case 3, 4, 10, 13, 30   'RS
                    Case 38     '&
                    
                    Case 21
                        msComm.Output = Chr(5)
                        m_iPhase = 2

                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                End Select

            Case 2          'SEND ORDER 1st
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        '----- 검사항목 조회/편집
                        iSendPhase = 1
                        Call Get_OrderString
                        
                        Call SendOrder_VIDAS

                        m_iPhase = 3

                    Case 21     'NAK
                        msComm.Output = Chr(5)
                        m_iPhase = 3

                    Case 5      'ENQ
                        msComm.Output = Chr(6)
                        m_iPhase = 1

                    Case Else
                End Select
            
            Case 3          'SEND ORDER n
                Select Case Asc(wkDat)
                    Case 6      'ACK
'                        msComm.Output = Chr(3)
'                        Call Sleep(100)
                        
                        If iSendPhase > pSampleInfo.ORDCNT Then
                            msComm.Output = Chr(4)
                            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, "", "")
                            m_iPhase = 1
                        Else
                            Call SendOrder_VIDAS
                            m_iPhase = 3
                        End If

                    Case 5      'ENQ
                        msComm.Output = Chr(6)
                        m_iPhase = 1

                    Case 21     'NAK
                        msComm.Output = Chr(5)
                        m_iPhase = 2

                    Case Else
                        m_iPhase = 1
                End Select
        End Select
    Next ix1

End Sub

Private Sub Send_oos()
'    On Error GoTo ErrRtn
'
'    Dim sSend$, sChkS$
'
'    sSend = Chr(30) & "mtoos|" & Chr(29)
'
'    'CheckSum 계산
'    sChkS = ChkSum_ASTM(sSend)
'
'    sSend = Chr(2) & sSend & sChkS
'
'    msComm.Output = sSend
'
'    If sTestMode = "77" Then
'        RaiseEvent PrintSendLog(sSend)
'    End If
'
'ErrRtn:
'    If Err <> 0 Then
'        RaiseEvent DispMsg("SendOrder 에러 - " & Err.Description)
'    End If
End Sub


Private Sub DataEdit_VIDAS()
    On Error GoTo ErrRtn
    
    Dim tmpData()   As String
    Dim sSID$, sPID$, sIFCd$, sRst1$, sRst2$, sUnit$
    Dim ii%
    Dim sTmp$, sData$, sSign$
    Dim tmpRst()    As String
    
    Call Init_pResultInfo
    
    If Mid(RcvBuffer, 1, 5) <> "mtrsl" Then     'RESULT CHECK
        Exit Sub
    End If
    
    tmpData() = Split(RcvBuffer, "|")
    
    For ii = 1 To UBound(tmpData())
        If Trim(tmpData(ii)) = "" Then
            Exit For
        End If
        
        sTmp = Left(tmpData(ii), 2)
        sData = Trim(Mid(tmpData(ii), 3))
        
        Select Case sTmp
            Case "pi"   'PID
                sPID = sData
            Case "si"   'SID
                sSID = sData
            Case "rt"   '장비 검사코드
                sIFCd = sData
            Case "ci"
                sPID = sData
            Case "qn"   '정량결과
                sRst1 = sData
                sUnit = "": sSign = ""
                If Left(Trim(sRst1), 1) = "<" Or Left(Trim(sRst1), 1) = ">" Then
                    sSign = Left(Trim(sRst1), 1)
                    sRst1 = sSign & Trim(Mid(Trim(sRst1), 2))
                End If
                If InStr(Trim(sRst1), " ") > 0 Then
                    tmpRst() = Split(Trim(sRst1), " ")
                    sRst1 = Trim(tmpRst(0))
                    sUnit = Trim(tmpRst(1))
                End If
                
            Case "ql"   '정성결과
                sRst2 = sData
        End Select
    Next ii
            
    '결과정보 구조체에 저장
    With pResultInfo
        .ID = sPID
        .SEQNO = sSID
        .RACK = "": .POS = ""
        .RSTCNT = 1
        .IFCD = sIFCd & Chr(124)
        .RST1 = sRst1 & Chr(124)
        .RST2 = sRst2 & Chr(124)
        .UNIT = sUnit & Chr(124)
        .FLAG = "" & Chr(124)
    End With
            
    '결과값 등록/화면 표시 처리...
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With

    Call Init_pResultInfo
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_VIDAS()
    On Error GoTo ErrRtn
    
    Dim sSend$, sTmp$, ChkS$
    Dim ii%
    
    msComm.Output = Chr(2)
    
    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2))
    End If
    Call Sleep(100)
    
    sSend = "pi" & Trim(pSampleInfo.ID) & Chr(124)
    sSend = sSend & "pn" & Trim(pSampleInfo.ID) & Chr(124)
    sSend = sSend & "pb" & Chr(124)
    sSend = sSend & "ps" & Chr(124)
    sSend = sSend & "so" & Chr(124)
    sSend = sSend & "si" & Chr(124)
    sSend = sSend & "ci" & Trim(pSampleInfo.ID) & Chr(124)
    
    'ORDER CODE
    sSend = sSend & "rt" & pSampleInfo.IFCD(iSendPhase) & Chr(124)
    iSendPhase = iSendPhase + 1
    
    sSend = sSend & "qd1" & Chr(124)
    
    sTmp = Chr(30) & "mtmpr" & Chr(124) & sSend
    sSend = Chr(30) & "mtmpr" & Chr(124) & sSend & Chr(29)
    
    'CheckSum 계산
    ChkS = LCase(ChkSum_ASTM(sSend))
    
    sSend = sTmp        'sSend & ChkS
    
    msComm.Output = sSend
    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSend)
    End If
    Call Sleep(100)
    
    sSend = Chr(29) & ChkS
    msComm.Output = sSend
    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSend)
    End If
    
    msComm.Output = Chr(3)
    Call Sleep(100)
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러 - " & Err.Description)
    End If
End Sub
Private Sub Send_bis()
'    On Error GoTo ErrRtn
'
'    Dim sSend$, sChkS$
'
'    sSend = Chr(30) & "mtbis|" & Chr(29)
'
'    'CheckSum 계산
'    sChkS = ChkSum_ASTM(sSend)
'
'    sSend = Chr(2) & sSend & sChkS
'
'    msComm.Output = sSend
'
'    If sTestMode = "77" Then
'        RaiseEvent PrintSendLog(sSend)
'    End If
'
'ErrRtn:
'    If Err <> 0 Then
'        RaiseEvent DispMsg("SendOrder 에러 - " & Err.Description)
'    End If
End Sub

Private Sub PhaseCfg_Protocol_miniVIDAS()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
       
        Select Case Asc(wkDat)
            Case 4      'EOT
                Call DataEdit_miniVIDAS
                
                RcvBuffer = ""
            
            Case 5      'ENQ
                RcvBuffer = ""
                msComm.Output = Chr(6)
            
            Case 29     'GS
                msComm.Output = Chr(6)
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat

        End Select
    Next ix1
    
End Sub

Private Sub DataEdit_miniVIDAS()
    On Error GoTo ErrRtn
    
    Dim tmpData()   As String
    Dim sTmp$, sSampID$, sIFCd$, sRst$, sQL$, sQN$
    Dim sFlag$
    
    tmpData() = Split(RcvBuffer, "|")
    
    If UBound(tmpData()) = 0 Then
        GoTo ErrRcvData
    End If
    
    '4) Sample No. 구하기
    sTmp = Trim(tmpData(4))
    If Mid(sTmp, 2, 2) <> "ci" Then
        GoTo ErrRcvData
    End If
    sSampID = Trim(Mid(sTmp, 4))
    
    '5) 장비 검사코드
    sTmp = Trim(tmpData(5))
    If Mid(sTmp, 2, 2) <> "rt" Then
        GoTo ErrRcvData
    End If
    sIFCd = Trim(Mid(sTmp, 4))
    
    '9~10) 검사결과
    sTmp = Trim(tmpData(9))     '정성결과
    If Mid(sTmp, 2, 2) <> "ql" Then
        GoTo ErrRcvData
    End If
    sQL = Trim(Mid(sTmp, 4))
    
    sTmp = Trim(tmpData(10))    '정량결과
    If Mid(sTmp, 2, 2) <> "qn" Then
        GoTo ErrRcvData
    End If
    
    sTmp = Trim(Mid(sTmp, 4))
    
    If Left(sTmp, 1) = ">" Or Left(sTmp, 1) = "<" Then
        sFlag = Left(sTmp, 1)
        sTmp = Trim(Mid(sTmp, 2))
    End If
    
    If InStr(sTmp, " ") > 0 Then
        tmpData() = Split(sTmp, " ")
        
        If UBound(tmpData) > 0 Then
            sQN = tmpData(0)
        End If
        
        sQN = sFlag & sQN
        
        If Trim(sQL) <> "" And Trim(sQN) <> "" Then
            sRst = sQL & "(" & sQN & ")"
        ElseIf sQN <> "" Then
            sRst = sQN
        Else
            sRst = sQL
        End If
    Else
        sQN = sTmp
        
        If Trim(sQL) <> "" And Trim(sQN) <> "" Then
            sRst = sQL & "(" & sQN & ")"
        ElseIf sQN <> "" Then
            sRst = sQN
        Else
            sRst = sQL
        End If
    End If
        
    '결과정보 구조체에 저장
    With pResultInfo
        .ID = sSampID
        .SEQNO = ""
        .RACK = "": .POS = ""
        .RSTCNT = 1
        .IFCD = sIFCd & Chr(124)
        .RST1 = sRst & Chr(124)
    End With
    
    '결과값 등록/화면 표시 처리...
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
        End If
    End With

    Call Init_pResultInfo

    Exit Sub
ErrRcvData:
    RaiseEvent DispMsg("전송된 데이터에 오류가 있습니다!!")
    Exit Sub
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
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
        .ORDCNT = iCnt      '실제 검사 가능한 항목 갯수
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
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
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
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintRcvLog(wkBuf)
            End If
                                
            If iSpaceCnt = 30 Then
                iSpaceCnt = 0
            End If
            iSpaceCnt = iSpaceCnt + 2
            
            RaiseEvent DispMsg(Space(iSpaceCnt) & "장비와 Interface 작업 중...")
            
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
'    m_iSMPLen = PropBag.ReadProperty("iSMPLen", m_def_iSMPLen)
'    m_iBCLen = PropBag.ReadProperty("iBCLen", m_def_iBCLen)
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
'    Call PropBag.WriteProperty("iSMPLen", m_iSMPLen, m_def_iSMPLen)
'    Call PropBag.WriteProperty("iBCLen", m_iBCLen, m_def_iBCLen)
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
'    m_iSMPLen = m_def_iSMPLen
'    m_iBCLen = m_def_iBCLen
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
'MemberInfo=14
Public Function ConnectionMsg() As Variant

'    m_iPhase = 7
'    msMsgType = "bis"
'
'    msComm.Output = Chr(5)
    
End Function

