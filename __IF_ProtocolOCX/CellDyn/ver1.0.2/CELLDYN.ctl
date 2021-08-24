VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl CELLDYN 
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
Attribute VB_Name = "CELLDYN"
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
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTInstID$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
'Event RequestCurOrder(sID$)
Event DispMsg(sMsg$)
'Event RequestNextOrder()
'Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)


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
Dim sOpenPW$, sEditPW$
Dim iSpaceCnt   As Integer

'For MiditronM
Dim bRstFlag    As Boolean
Dim msSampleNo  As String       '단방향에서 장비번호 체크 사용
Dim msPreSampleNo   As String   '              "
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
        Case "CELLDYN1800"
            Call PhaseCfg_Protocol_CellDyn1800
            
        Case "CELLDYN3500"
            Call PhaseCfg_Protocol_CellDyn3500
            
        Case "CELLDYNRUBY"
            Call PhaseCfg_Protocol_CellDynRuby
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub

Private Sub PhaseCfg_Protocol_CellDynRuby()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
        
        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 5      'ENQ
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                    Case Else
                        m_iPhase = 1
                End Select
            
            Case 2
                Select Case Asc(wkDat)
                    Case 2      'STX
                        RcvBuffer = ""
                        
                    Case 3
                        msComm.Output = Chr(6)
                        
                    Case 10     '<LF>
                        Call DataEdit_CellDynRuby
                        RcvBuffer = ""
                        
                        msComm.Output = Chr(6)
                        
                    Case 13     '<CR>
                        Call DataEdit_CellDynRuby
                        RcvBuffer = ""
                        
'                        m_iPhase = 2
'                        msComm.Output = Chr(6)
                        
                    Case 4      'EOT
                        m_iPhase = 1
                                                
                    Case 5      'ENQ
                        msComm.Output = Chr(6)   'Send ACK
                        
                    Case 21     'NAK
                        msComm.Output = Chr(5)   'Send ENQ
                        m_iPhase = 1
                        
                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                        m_iPhase = 2
                End Select
            
'            Case 3
'                Select Case Asc(wkDat)
'                    Case 6      'ACK
'                        Call SendOrder_E411     'Order 전송
'
'                    Case 5      'ENQ
'                        m_iPhase = 2
'                        msComm.Output = Chr(6)
'
'                    Case 21     'NAK
'                        m_iSendPhase = m_iSendPhase - 1
'                        m_iFrameN = m_iFrameN - 1
'                        m_iPhase = 3
'
'                        Call SendOrder_E411     'Order 전송
'
'                    Case 4      'EOT
'                        m_iPhase = 1
                        
'                End Select
        End Select
    Next ix1
    
End Sub

Private Sub PhaseCfg_Protocol_CellDyn3500()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2         ' STX
                RcvBuffer = ""
                
            Case 3         ' ETX
                '--- 결과편집/등록
                Call DataEdit_CellDyn3500
                msComm.Output = Chr(6)
                
       
            Case Else      '
                RcvBuffer = RcvBuffer & wkDat
         End Select
    Next ix1
    
End Sub
Private Sub PhaseCfg_Protocol_CellDyn1800()

    Dim wkDat   As String
    Dim ix1     As Integer

    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2         ' STX
                RcvBuffer = ""
                
            Case 3         ' ETX
                '--- 결과편집/등록
                Call DataEdit_CellDyn1800
                msComm.Output = Chr(6)
                
            Case Else      '
                RcvBuffer = RcvBuffer & wkDat
         End Select
    Next ix1
    
End Sub

Private Sub DataEdit_CellDynRuby()
    On Error GoTo ErrRtn

    Dim tmpIFCd$, tmpRst$, tmpRst2$
    Dim tmpMsgType$, tmpSpcType$
    Dim tmpSeqNo$, tmpBarCd$, tmpRack$, tmpPos$
    Dim i%, ii%
    Dim RecType$
    Dim tmpData()   As String
    Dim tmpBuf()    As String
    
    ii = InStr(1, RcvBuffer, "|")
    If ii <> 0 Then
        RecType = Mid$(RcvBuffer, ii - 1, 1)
    Else
        Exit Sub
    End If

    Select Case RecType
        Case "H"        'Header Record
            sState = ""
            
        Case "M"
        Case "P"        'Patient Record
            Call Init_pResultInfo

        Case "Q"        'Order Request Record
'            tmpData() = Split(RcvBuffer, "|")
'            sReqStatusCd = Left(tmpData(12), 1)     'Order Request Status Code
'            tmpData() = Split(tmpData(2), "^")
'
'            tmpBarCd = Trim(tmpData(1))
'            tmpSeqNo = Trim(tmpData(2))
'            tmpRack = Trim(tmpData(3))
'            tmpPos = Trim(tmpData(4))
'
'            If tmpBarCd = "" Then           'BarCode ID가 잘 넘어왔는지 검사
'                sState = ""
'                sReqStatusCd = ""
'                pSampleInfo.ID = ""
'            End If
'
'
'            sState = "Q"
'            pSampleInfo.ID = tmpBarCd        'BarCode
'
'            '--- Rack 버전인 경우 SeqNo/Rack/Pos도 Order와 함께 전송함
'            pSampleInfo.SEQNO = tmpSeqNo
'            pSampleInfo.RACK = tmpRack
'            pSampleInfo.POS = tmpPos
            
        Case "O"
            tmpData() = Split(RcvBuffer, "|")
            tmpBarCd = Trim(tmpData(2))
            tmpSeqNo = Trim(tmpData(3))

            pSampleInfo.ID = tmpBarCd
            pSampleInfo.SEQNO = tmpSeqNo
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos

        Case "R"        'Result Record
            '--- 결과데이타 편집
            'tmpData(2): TESTCD
            '    "  (3): RESULT
            '    "  (4): UNIT
            '    "  (5): 참고치 범위
            '    "  (6): 참고치(N:Neg/H:Pos)
            tmpData() = Split(RcvBuffer, "|")
            If Trim(tmpData(2)) <> "" Then
                i = InStr(tmpData(2), "^")
                
                If i <> 0 Then
                    tmpBuf = Split(tmpData(2), "^")
                End If
            End If
            
            tmpIFCd = tmpBuf(6)
            
            tmpRst = Trim(tmpData(3))
            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If
            
            If tmpRst = "FLAG" Then Exit Sub
            
'            tmpUnit = Trim(tmpData(4))
'            tmpFlag = Trim(tmpData(6))
          

            '결과정보 구조체에 저장
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS

                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & tmpRst2 & Chr(124)
                .UNIT = .UNIT & "" & Chr(124)
                .FLAG = .FLAG & "" & Chr(124)
            End With

        Case "C"        'Comment Record

        Case "L"
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", .KIND, "", "")
                End If
            End With

            Call Init_pResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub DataEdit_CellDyn3500()
    On Error GoTo ErrRtn

    Dim sIFCd$, sRst1$
    Dim tmpMsgType$, tmpSpcType$
    Dim tmpSeqNo$, tmpBarCd$
    Dim ii%
    Dim tmpData()   As String

    '결과정보 구조체 초기화
    Call Init_pResultInfo

    'DATA 편집
    tmpData() = Split(RcvBuffer, ",")

    tmpMsgType = Trim(Replace(tmpData(0), """", ""))    'Message Type (Count Data: "   ", Retic Data: "RTC")

    If tmpMsgType <> "" And tmpMsgType <> "RTC" Then
        Exit Sub
    End If

    tmpSeqNo = Trim(tmpData(1))
    tmpSpcType = Trim(tmpData(3))   '0:Patient, 1:QC Sample
    
    If tmpSpcType <> "0" And tmpSpcType <> "1" Then
        Exit Sub
    End If
    
    tmpBarCd = Trim(tmpData(4))
    tmpBarCd = Replace(tmpBarCd, """", "")

    If tmpMsgType = "" Then
        'Count Data
        For ii = 9 To 30
            If Trim(tmpData(ii)) = "" Then
                Exit For
            End If
    
            If tmpData(ii) = "-----" Then
            Else
                sIFCd = Trim(ii - 8)
                sRst1 = Trim(tmpData(ii))
    
                If Left(sRst1, 1) = "." Then
                    sRst1 = "0" & sRst1
                End If
                If Right(sRst1, 1) = "." Then
                    sRst1 = sRst1 & "0"
                End If
    
                'Data 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & sIFCd & Chr(124)
                    .RST1 = .RST1 & sRst1 & Chr(124)
                    .RST2 = .RST2 & Chr(124)
                    .UNIT = .UNIT & Chr(124)
                    .FLAG = .FLAG & Chr(124)
                End With
            End If
        Next ii
    Else
        'Retic Data
        For ii = 11 To 12
            If Trim(tmpData(ii)) = "" Then
                Exit For
            End If
    
            If tmpData(ii) = "-----" Then
            Else
                sIFCd = Trim(ii + 12)
                sRst1 = Trim(tmpData(ii))
    
                If Left(sRst1, 1) = "." Then
                    sRst1 = "0" & sRst1
                End If
                If Right(sRst1, 1) = "." Then
                    sRst1 = sRst1 & "0"
                End If
    
                'Data 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & sIFCd & Chr(124)
                    .RST1 = .RST1 & sRst1 & Chr(124)
                    .RST2 = .RST2 & Chr(124)
                    .UNIT = .UNIT & Chr(124)
                    .FLAG = .FLAG & Chr(124)
                End With
            End If
        Next ii
    End If
    
    '결과값 등록/화면 표시 처리...
    With pResultInfo
        .SEQNO = tmpSeqNo
        .ID = tmpBarCd
        
        If tmpSpcType = "0" Then
            .KIND = ""
        ElseIf tmpSpcType = "1" Then
            .KIND = "QC"
        End If

        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", .KIND, "", "")
        End If
    End With

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub
Private Sub DataEdit_CellDyn1800()
    On Error GoTo ErrRtn
    
    Dim sIFCd$, sRst1$
    Dim tmpSeqNo$, tmpBarCd$
    Dim ii%
    Dim tmpData()   As String
    
    '결과정보 구조체 초기화
    Call Init_pResultInfo
    
    'DATA 편집
    tmpData() = Split(RcvBuffer, ",")
    
    tmpSeqNo = Trim(tmpData(3))
    tmpBarCd = Trim(tmpData(8))
    tmpBarCd = Replace(tmpBarCd, """", "")
    
    For ii = 16 To 36
        If Trim(tmpData(ii)) = "" Then
            Exit For
        End If
        
        If tmpData(ii) = "-----" Then
        Else
            sIFCd = Trim(ii)
            sRst1 = Format(tmpData(ii))
            
            Select Case ii
                Case 22, 31
                    sRst1 = Convert_Value(sRst1, 2)
                Case 29
                    sRst1 = Convert_Value(sRst1, 0)
                Case Else
                    sRst1 = Convert_Value(sRst1, 1)
            End Select
            
            If Left(sRst1, 1) = "." Then
                sRst1 = "0" & sRst1
            End If
            If Right(sRst1, 1) = "." Then
                sRst1 = sRst1 & "0"
            End If
            
            'Data 누적
            With pResultInfo
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & sIFCd & Chr(124)
                .RST1 = .RST1 & sRst1 & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .UNIT = .UNIT & Chr(124)
                .FLAG = .FLAG & Chr(124)
            End With
        End If
    Next ii
    
    '결과값 등록/화면 표시 처리...
    With pResultInfo
        .SEQNO = tmpSeqNo
        .ID = tmpBarCd
        
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", "", "", "", "")
        End If
    End With
                                            
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

'장비에서 넘어온 결과값을 Shift시켜서 자리수를 맞춘다.
Private Function Convert_Value(ByVal sValue As String, ByVal iShift As Integer) As String
    Convert_Value = CStr(Val(sValue) / (10 ^ iShift))
End Function
Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
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
        For ii = 1 To .ORDCNT
            .IFCD(ii) = tmpData(ii - 1)
        Next ii
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
        .KIND = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
        .INSTID = ""
        .ALARMCD = ""
        .RSTDT = ""
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
    
    '변수 초기화
    bRstFlag = False
    
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
    RaiseEvent DispMsg("Send_Chr 에러 - " & Err.Description)
End Function

