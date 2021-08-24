VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl VITROS5600 
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
Attribute VB_Name = "VITROS5600"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_p_sDilution = ""
Const m_def_p_sSpcGbn = 0
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
Dim m_p_sDilution As String
Dim m_p_sSpcGbn As Variant
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
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
Event RequestNextOrder()
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
Event SendOrderOK(sID$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
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

'for Vitros5600/3600
Dim sDiluInfo   As String


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
        Case "VITROS5600", "VITROS3600"
            If m_bUseBarcode = True Then
                Call PhaseCfg_Protocol_VITROS5600
            Else
                Call PhaseCfg_Protocol_VITROS5600_Batch
            End If
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_VITROS5600_Batch()
    
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
                        
                    Case 10     '<LF>
                        Call DataEditResponse_VITROS5600
                        
                        m_iPhase = 2
                        msComm.Output = Chr(6)
                        
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
            
            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        Call SendOrder_VITROS5600       'Order 전송
                        
                    Case 5      'ENQ
                        m_iPhase = 2
                        msComm.Output = Chr(6)
                                                
                    Case 21     'NAK
                        m_iSendPhase = m_iSendPhase - 1
                        m_iFrameN = m_iFrameN - 1
                        m_iPhase = 3
                        
                        Call SendOrder_VITROS5600       'Order 전송
                        
                    Case 4      'EOT
                        m_iPhase = 1
                        
                End Select
        End Select
    Next ix1
    
End Sub

' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_VITROS5600()
    On Error GoTo ErrRtn

    Dim RecType As String   'Record Type
    Dim i       As Integer
    Dim tmpField()  As String
    Dim tmpData()   As String
    Dim tmpBarCd$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpMDil$, tmpTDil$
    Dim tmpRst1$, tmpUnit$, tmpFlag$, tmpRstDT$
    Dim aFlag() As String

    RecType = Mid$(RcvBuffer, 2, 1)

    Select Case RecType
        Case "H"        'Header Record
            sState = ""

        Case "M"
        Case "P"        'Patient Record
            Call Init_pResultInfo

        Case "Q"        'Order Request Record
            'Q|1|^123456789012345||ALL||||||||O
            tmpField() = Split(RcvBuffer, "|")

            sReqStatusCd = Left(tmpField(12), 1)    'Order Request Status Code

            tmpData() = Split(tmpField(2), "^")

            tmpBarCd = Trim(tmpData(1))
            tmpRack = ""
            tmpPos = ""

            If tmpBarCd = "" Then           'BarCode ID가 잘 넘어왔는지 검사
                sState = ""
                sReqStatusCd = ""
                pSampleInfo.ID = ""
            End If

            sState = "Q"

            pSampleInfo.ID = tmpBarCd        'BarCode
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos

        Case "O"
            'O|1|SampleID1234567^10^10||^^^1.0000+309+1.0\308+1.0\307+1.0\310+1.0\300+1.0\315+1.0\314+1.0\318+1.0|S||20030815073000||||N||||1||||||||||F||||||
            tmpField() = Split(RcvBuffer, "|")

            tmpData() = Split(tmpField(2), "^")
            tmpBarCd = Trim(tmpData(0))
            tmpRack = Trim(tmpData(1))
            tmpPos = Trim(tmpData(2))

            pSampleInfo.ID = tmpBarCd
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos

        Case "R"        'Result Record
            'R|2|^^^1.0000+950+1.0|500|||^0^AR||V|||20031015073055|20031015073800|J00000013
            'R|1|^^^1.0000+309+1.0|142|mg/dL||^0^UCLS\^0^pi\^0^ES\^J^||V|||20031015073055|20031015073800| J00000013

            '--- 결과데이타 편집
            'tmpData(2): TESTCD
            '    "  (3): RESULT
            '    "  (4): UNIT
            tmpField() = Split(RcvBuffer, "|")

            tmpData() = Split(tmpField(2), "^")
            tmpData() = Split(tmpData(3), "+")
            tmpMDil = Trim(tmpData(0))
            tmpIFCd = Trim(tmpData(1))      '검사코드
            tmpTDil = Trim(tmpData(2))

            tmpRst1 = Trim(tmpField(3))
            If Left$(tmpRst1, 1) = "." Then
                tmpRst1 = "0" & tmpRst1
            End If
            tmpUnit = Trim(tmpField(4))
            
            tmpFlag = Trim(tmpField(6))
            aFlag() = Split(tmpFlag, "^")
            tmpFlag = Trim(aFlag(1))
            If tmpFlag = "0" Then
                tmpFlag = ""
            End If
            
            If UBound(tmpField()) >= 12 Then
                tmpRstDT = Trim(tmpField(12))
            End If
            
            '결과정보 구조체에 저장
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = ""
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS

                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst1 & Chr(124)
                .RST2 = .RST2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
                .RSTDT = .RSTDT & tmpRstDT & Chr(124)
            End With

        Case "C"        'Comment Record

        Case "L"
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "", .KIND, .RSTDT, "")
                End If
            End With

            Call Init_pResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub PhaseCfg_Protocol_VITROS5600()

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

                    Case 10     '<LF>
                        Call DataEditResponse_VITROS5600

                        m_iPhase = 2
                        msComm.Output = Chr(6)
                        
                    Case 4      'EOT
                        If sState = "Q" Then
                            msComm.Output = Chr(5)
                            m_iSendPhase = 1
                        End If
                        m_iPhase = 3

                    Case 5      'ENQ
                        msComm.Output = Chr(6)   'Send ACK

                    Case 21     'NAK
                        msComm.Output = Chr(5)   'Send ENQ
                        m_iPhase = 1

                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                        m_iPhase = 2
                End Select

            Case 3
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        If sState = "Q" Then
                            Call SendOrder_VITROS5600   'Order 전송
                        End If

                    Case 5      'ENQ
                        msComm.Output = Chr(6)
                        m_iPhase = 2

                    Case 21     'NAK
                        msComm.Output = Chr(5)
                        m_iPhase = 3

                    Case 4      'EOT
                        m_iPhase = 1
                End Select

        End Select
    Next ix1
    
End Sub

Private Sub SendOrder_VITROS5600()
    On Error GoTo ErrRtn

    Dim sTmp    As String
    Dim ChkS    As String
    Dim TestDat As String
    Dim i       As Integer
    Dim sTmpData()  As String
    Dim sActionCd   As String

    Dim aDil()  As String
    Dim ix9     As Integer
    Dim bDilChk As Boolean
    
    If m_iFrameN > 7 Then
        m_iFrameN = 0
    End If

    Select Case m_iSendPhase
        Case 0
            m_iSendPhase = 1
            msComm.Output = Chr(5)
            Exit Sub

        Case 1      'H
            '                   H|\^&|||EM13||||||||E 1394-97|200209301154000
            sTmp = m_iFrameN & "H|\^&|||EM13||||||||E1394-97" & Chr(13) & Chr(3)
            m_iSendPhase = 2

        Case 2      'P
            '                   P|1|MRN123456789012             |||||||||||I||||||||||||||||||||||
            If m_bUseBarcode = True Then
                sTmp = m_iFrameN & "P|1|" & Trim(pSampleInfo.ID) & "||||||||||||||||||||||||||||||||||" & Chr(13) & Chr(3)
            Else
                sTmp = m_iFrameN & "P|1" & Chr(13) & Chr(3)
            End If
            m_iSendPhase = 3

        Case 3      'O
            'O|1|SampleID1234567^10^10||^^^1.0000+309+1.0\308+1.0\307+1.0\310+1.0\300+1.0\315+1.0\314+1.0\318+1.0|S||20030815073000||||N||||1||||||||||F||||||
            TestDat = ""
            sDiluInfo = ""
            '----- 검사항목 조회/편집
            If m_bUseBarcode = True Then
                RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
    
                Call Get_OrderString
    
                If pSampleInfo.ORDCNT = 0 Then
                    sActionCd = "C"
                    RaiseEvent DispMsg("인터페이스 오더 항목이 존재하지 않습니다!!")
                Else
                    sActionCd = "N"
                End If
            Else
                Call Get_OrderString
                
                sActionCd = "N"
            End If
            
            For i = 1 To pSampleInfo.ORDCNT
                '<S--- Auto Dilution...2009/10/6 yk
                bDilChk = False
                If Trim(sDiluInfo) <> "" Then
                    Erase aDil()
                    aDil() = Split(sDiluInfo, Chr(124))
                    For ix9 = 0 To UBound(aDil())
                        If Trim(aDil(ix9)) = "" Then Exit For
                        
                        If Trim(pSampleInfo.IFCD(i)) = Left(aDil(ix9), 3) And Trim(aDil(ix9)) <> "" Then
                            bDilChk = True
                            pSampleInfo.IFCD(i) = Trim(aDil(ix9))
                            
                            Exit For
                        End If
                    Next ix9
                End If
                '>E-----------------

                If bDilChk = True Then
                    TestDat = TestDat & Trim(pSampleInfo.IFCD(i)) & "\"     '희석배율 함께 전송
                Else
                    TestDat = TestDat & Trim(pSampleInfo.IFCD(i)) & "+1.0\"
                End If
            Next i
            If pSampleInfo.ORDCNT > 0 Then
                TestDat = "^^^1.0000+" & Left(TestDat, Len(TestDat) - 1)        '"\" Cutting
            End If
            '-------------------

            sTmp = m_iFrameN & "O|1|" & Trim(pSampleInfo.ID) & "^" & pSampleInfo.RACK & "^" & pSampleInfo.POS & "||" _
                & TestDat & "|R||||||" _
                & sActionCd & "||||" & pSampleInfo.SPCCD & "||||||||||O||||||" & vbCr & Chr(3)

            m_iSendPhase = 4

        Case 4      'T
            sTmp = m_iFrameN & "L|1|N" & vbCr & Chr(3)
            m_iSendPhase = 5

        Case 5      'EOT
            msComm.Output = Chr(4)   'EOT
            m_iFrameN = 1: m_iPhase = 1: m_iSendPhase = 1
            sState = ""

            If m_bUseBarcode = True Then
                'Barcode Mode인 경우 전송완료 이벤트 발생
                RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
            Else
                'BarCode Mode가 아닌 경우 다음 오더 조회
                RaiseEvent RequestNextOrder
            End If

            Exit Sub

    End Select

    'CheckSum 계산
    ChkS = ChkSum_ASTM(sTmp)

    msComm.Output = Chr(2) & sTmp & ChkS & Chr(13) & Chr(10)

    m_iFrameN = m_iFrameN + 1

    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sTmp & ChkS & Chr(13) & Chr(10))
    End If
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러 - " & Err.Description)
    End If
End Sub

Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    Dim sTmpSpc As String
    
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
        
        Select Case Trim(m_p_sSpcGbn)
            Case "2": .SPCCD = "2"      'CSF
            Case "3": .SPCCD = "3"      'Urine
            Case "4": .SPCCD = "4"      'W.B
            Case "5": .SPCCD = "5"      'Serum
            Case "6": .SPCCD = "6"      'Plasma
            Case "7": .SPCCD = "7"      'Amniotic Fluid
            Case Else: .SPCCD = "5"     'Serum
        End Select
        
        ReDim .IFCD(.ORDCNT)
        iCnt = 0
        For ii = 1 To .ORDCNT
            If Trim(tmpData(ii - 1)) <> "" Then
                iCnt = iCnt + 1
                .IFCD(iCnt) = tmpData(ii - 1)
            End If
        Next ii
        .ORDCNT = iCnt      '실제 검사 가능한 항목 갯수
        
        sDiluInfo = m_p_sDilution
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
        .KIND = ""
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
    m_p_sSpcGbn = PropBag.ReadProperty("p_sSpcGbn", m_def_p_sSpcGbn)
    m_p_sDilution = PropBag.ReadProperty("p_sDilution", m_def_p_sDilution)
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
    Call PropBag.WriteProperty("p_sSpcGbn", m_p_sSpcGbn, m_def_p_sSpcGbn)
    Call PropBag.WriteProperty("p_sDilution", m_p_sDilution, m_def_p_sDilution)
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
    m_p_sSpcGbn = m_def_p_sSpcGbn
    m_p_sDilution = m_def_p_sDilution
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
Public Property Get p_sSpcGbn() As Variant
    p_sSpcGbn = m_p_sSpcGbn
End Property

Public Property Let p_sSpcGbn(ByVal New_p_sSpcGbn As Variant)
    m_p_sSpcGbn = New_p_sSpcGbn
    PropertyChanged "p_sSpcGbn"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,
Public Property Get p_sDilution() As String
    p_sDilution = m_p_sDilution
End Property

Public Property Let p_sDilution(ByVal New_p_sDilution As String)
    m_p_sDilution = New_p_sDilution
    PropertyChanged "p_sDilution"
End Property

