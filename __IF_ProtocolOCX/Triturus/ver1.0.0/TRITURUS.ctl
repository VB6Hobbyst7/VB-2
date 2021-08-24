VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl TRITURUS 
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
Attribute VB_Name = "TRITURUS"
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
Dim m_State As String
Dim m_iFrameNo As Integer
Dim m_iPatNo As Integer
Dim m_iTestCnt As Integer
Dim m_sRetrans As String
Dim m_p_sRerunGbn As String
Dim m_iEtbGbn As Integer
Dim m_sGbnBuf As String
Dim m_aTemp() As String
Dim m_iSndCnt As Integer
Dim m_sSavBuf As String

'이벤트 선언:
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
Event SendOrderOK(sID$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
'Event RequestCurOrder(sID$, sRack$, sPos$)
Event DispMsg(sMsg$)
Event RequestNextOrder()

'===== User Define
'인터페이스에서 사용
Dim RcvBuffer   As String
Dim SavBuffer   As String
Dim wkBuf   As String
Dim sState  As String
Dim sReqStatusCd    As String

'구조체 지정
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'기타
Dim iSpaceCnt   As Integer
Dim bEndChk As Boolean
Dim bSTXChk As Boolean
Dim sNextSend   As String
Dim RstEnd      As String

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
        Case "TRITURUS"
            If m_bUseBarcode = True Then
                '바코드 사용
                'Call PhaseCfg_Protocol_Triturus_BarcodeMode
            Else
                '바코드 사용 안함
                Call PhaseCfg_Protocol_Triturus
            End If
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub

Private Sub PhaseCfg_Protocol_Triturus()
    Dim wkDat   As String
    Dim ix1     As Integer
        
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 5      'ENQ
                msComm.Output = Chr(6)

            Case 2      'STX
                bEndChk = True
                SavBuffer = ""
            
            Case 10     '<LF>
                If bEndChk = True Then
                    Call DataEditResponse_Triturus  '데이터 Edit

                    msComm.Output = Chr(6)
                    RcvBuffer = ""
                End If

            Case 4      'EOT
                If m_State = "Q" Then   'Triturus로 부터 Q를 받았을때
                    msComm.Output = Chr(5)
                    m_State = "O"       'Order전송모드
                Else
                End If

            Case 6      'ACK
                If m_State = "O" Then   'Order전송모드
                    Call SendOrder_Triturus
                    
                ElseIf m_State = "S" Then   'Send모드
                
                    If m_aTemp(m_iSndCnt) <> "" Then
                        msComm.Output = m_aTemp(m_iSndCnt)
                        m_sRetrans = m_aTemp(m_iSndCnt)     '재전송을 위해 m_sRetrans에 저장
                        
                        If sTestMode = "77" Then
                            RaiseEvent PrintSendLog(m_aTemp(m_iSndCnt))
                        End If
                        
                        m_iSndCnt = m_iSndCnt + 1
                    Else
                        m_State = ""
                        m_iSndCnt = 0
                        msComm.Output = Chr(4)
                    End If
                Else
                End If

            Case 21     'NAK
                msComm.Output = m_sRetrans  '데이터 재전송
            
            Case 23     'ETB
                bEndChk = False
                RcvBuffer = RcvBuffer & Mid(SavBuffer, 2, Len(SavBuffer) - 1)
                msComm.Output = Chr(6)
                
            Case 3
                RcvBuffer = RcvBuffer & Mid(SavBuffer, 2, Len(SavBuffer) - 1)
                msComm.Output = Chr(6)
            
            Case Else
                If bEndChk = True Then
                    SavBuffer = SavBuffer & wkDat
                End If

        End Select
        
    Next ix1

End Sub

' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_Triturus()
    On Error GoTo ErrRtn

    Dim RecType As String   'Record Type
    Dim i       As Integer
    Dim ii      As Integer
    Dim tmpData()   As String
    Dim tmpField()  As String
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$, tmpKind$, tmpDate$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpRef$, tmpFlag$, tmpAlarmCd$
    
    '''1H|\^&|||Triturus^1^4.01|||||||P|1|20070419150438
    '''P|1||
    '''O|1||033-1|^^^PROTEIN C|R|20070419145303||||||||||||||||||Triturus^1^4.01|I|
    '''R|1|^^^PROTEIN C||IU/ml||||I|||||Triturus^1^4.01|

'''    1H|\^&|||Triturus^1^4.01|||||||P|1|20070418190204
'''    Q|1|ALL||ALL||||||||I|
'''    L|1|N
'''    C4
    
    Dim sCrSplit() As String
    Dim iCrCnt     As Integer
    
    sCrSplit() = Split(RcvBuffer, Chr(13))
    
    For iCrCnt = 0 To UBound(sCrSplit) - 1
    
        ii = InStr(1, sCrSplit(iCrCnt), "|")
        If ii <> 0 Then
            RecType = Mid$(sCrSplit(iCrCnt), ii - 1, 1)
        Else
            Exit Sub
        End If
    
        Select Case RecType
            Case "H"        'Header Record
                m_State = ""
    
            Case "P"        'Patient Record
                
                tmpData() = Split(sCrSplit(iCrCnt), "|")
                
                If tmpData(1) = "1" Then
                    Call Init_pResultInfo
                Else
                    With pResultInfo
                    If .RSTCNT > 0 Then
                        pResultInfo.ALARMCD = Replace(pResultInfo.ALARMCD, "####", "")      '2006/8/17 yk
    
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .ALARMCD, .KIND, "", "")
    
                    End If
                End With

                    Call Init_pResultInfo
                End If
                
            Case "Q"        'Order Request Record
                m_State = "Q"
    
            Case "O"
                'O|1|A0001234|912-1|^^^SQ CMV IgM|R|19990707144438||||||||||||||||||Triturus^1^2.00|F|
                'O|1||033-1|^^^PROTEIN C|R|20070419145303||||||||||||||||||Triturus^1^4.01|I|
                tmpData() = Split(sCrSplit(iCrCnt), "|")
                tmpBarCd = Trim(tmpData(2))
                tmpDate = Trim(tmpData(6))
    
                pSampleInfo.ID = tmpBarCd
    
            Case "R"        'Result Record
                '--- 결과데이타 편집
                '4R|1|^^^SQ CMV IgM|4.714286|Index||HH||F||Roser Ambros|||Triturus^1^4.00|
                'R|1|^^^PROTEIN C||IU/ml||||I|||||Triturus^1^4.01|
                'tmpData(2): TESTCD
                '    "  (3): RESULT
                '    "  (4): UNIT
                '    "  (6): 참고치(Negative = N 'normal'
                                   'Doubtful = H 'above high normal'
                                   'Positive = HH 'above panic high')
                '    "  (8): 결과상태(F = Final, validated result.
    '                                 X = Result not available because the test was cancelled.
    '                                 l = The result is not available yet, but the test is being executed.
    '                                 C = Modification of a previously sent result.)
    
                tmpData() = Split(sCrSplit(iCrCnt), "|")
                                
                tmpIFCd = Trim(Split(tmpData(2), "^")(3))
    
                tmpRst = Trim(tmpData(3))
                If Left$(tmpRst, 1) = "." Then
                    tmpRst = "0" & tmpRst
                End If
    
                tmpUnit = Trim(tmpData(4))
                tmpRef = Trim(tmpData(6))
                tmpFlag = Trim(tmpData(8))
                If tmpFlag = "N" Then
                    tmpFlag = ""
                End If
    
                '결과정보 구조체에 저장
                With pResultInfo
                    .ID = pSampleInfo.ID
                    .SEQNO = pSampleInfo.SEQNO
                    .RACK = pSampleInfo.RACK
                    .POS = pSampleInfo.POS
                    .KIND = pSampleInfo.KIND
    
                    '결과값 누적
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD & tmpIFCd & Chr(124)
                    .RST1 = .RST1 & tmpRst & Chr(124)
                    .RST2 = .RST2 & tmpRef & Chr(124)
                    .UNIT = .UNIT & tmpUnit & Chr(124)
                    .FLAG = .FLAG & tmpFlag & Chr(124)
                    .ALARMCD = .ALARMCD & "" & Chr(124)
                End With
                
            Case "C"        'Comment Record
    '           5C|1|I|>|I|
    '           6C|2|I|High positive|P|
    
                'Data Alarm 편집
                tmpData() = Split(sCrSplit(iCrCnt), "|")
                
                If Trim(tmpData(1)) = "1" Then
                    tmpAlarmCd = Trim(tmpData(3))
                End If
                    
                If tmpAlarmCd = ">" Then
                    pResultInfo.RST1 = tmpAlarmCd & pResultInfo.RST1
                    tmpAlarmCd = ""
                End If
    
    '            pResultInfo.ALARMCD = pResultInfo.ALARMCD & tmpAlarmCd & Chr(124)
                pResultInfo.ALARMCD = Replace(pResultInfo.ALARMCD, "####", tmpAlarmCd)
    
            Case "L"
                '결과값 등록/화면 표시 처리...
                With pResultInfo
                    If .RSTCNT > 0 Then
                        pResultInfo.ALARMCD = Replace(pResultInfo.ALARMCD, "####", "")      '2006/8/17 yk

                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .ALARMCD, .KIND, "", "")

                    End If
                End With

                Call Init_pResultInfo
    
        End Select
      
    Next
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub SendOrder_Triturus()
    On Error GoTo ErrRtn

    Dim sTmp    As String
    Dim ChkS    As String
    Dim TestDat As String
    Dim i       As Integer
    Dim sTmpData()  As String
    Dim sActionCd   As String
    Dim sReportType As String
    Dim iDiv As Integer
    Dim iCnt As Integer

'''    H|\^&
'''    P|1|A0001234
'''    O|1|A0001234||^^^ACA IgG
'''    O|2|A0001234||^^^ACA IgM
'''    P|2|A0001235
'''    O|1|A0001235||^^^MPO lgG
'''    O|2|A0001235||^^^PR3 lgG
    
    '<Order를 모두 모아서 240으로 나누어 전송>
    sTmp = "H|\^&" & Chr(13)   'Header Record
    
    Do
        RaiseEvent RequestNextOrder
    
        Call Get_OrderString
        
        If pSampleInfo.ID = "" Then
            sTmp = sTmp & "L|1|N"   'Last Record
            m_State = "S"
            Exit Do
            
        Else
            m_iPatNo = m_iPatNo + 1
            sTmp = sTmp & "P|" & m_iPatNo & "|" & pSampleInfo.ID & Chr(13)  'Patient Record
            
            For i = 1 To pSampleInfo.ORDCNT
                sTmp = sTmp & "O|" & i & "|" & pSampleInfo.ID & "||^^^" & pSampleInfo.IFCD(i) & Chr(13) 'Test Record
            Next
        
        End If
    Loop
           
    iDiv = Len(sTmp) / 240
    ReDim m_aTemp(iDiv + 1)
   
    For i = 0 To iDiv
        If Len(sTmp) > 240 Then
            ChkS = ChkSum_ASTM(m_iFrameN & Mid(sTmp, 1, 240) & Chr(23))
            m_aTemp(i) = Chr(2) & m_iFrameN & Mid(sTmp, 1, 240) & Chr(23) & ChkS & Chr(13) & Chr(10)
            sTmp = Replace(sTmp, Mid(sTmp, 1, 240), "")
        Else
            ChkS = ChkSum_ASTM(m_iFrameN & sTmp & Chr(3))
            m_aTemp(i) = Chr(2) & m_iFrameN & sTmp & Chr(3) & ChkS & Chr(13) & Chr(10)
        End If
        
        m_iFrameN = m_iFrameN + 1
        
        If m_iFrameN > 7 Then      'Frame Number가 8이상이면 0으로 바꿔줌
            m_iFrameN = 0
        End If
        
    Next
    
    msComm.Output = m_aTemp(m_iSndCnt)
    m_sRetrans = m_aTemp(m_iSndCnt)     '재전송을 위해 m_sRetrans에 저장
                            
    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(m_aTemp(m_iSndCnt))
    End If
    
    m_iSndCnt = m_iSndCnt + 1
    
    sTmp = ""
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러 - " & Err.Description)
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
    
    '변수 초기화(E-170/H-7600)
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
    If Err <> 0 Then
        RaiseEvent DispMsg("Send_Chr 에러 - " & Err.Description)
    End If
End Function

