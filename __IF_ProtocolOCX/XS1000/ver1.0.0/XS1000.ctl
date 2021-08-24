VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl XS1000 
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
Attribute VB_Name = "XS1000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_p_sPatInfo = "0"
Const m_def_p_sSampInfo = "0"
Const m_def_SiteNm = 0
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
Dim m_p_sPatInfo As String
Dim m_p_sSampInfo As String
Dim m_SiteNm As Variant
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
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)
Event SendOrderOK(sID$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event RequestCurOrder(sID$, sRack$, sPos$)
Event DispMsg(sMsg$)
Event RequestNextOrder()

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

Dim sNextSend As String
Dim bETBChk As Boolean


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
        Case "XS1000"
            Call PhaseCfg_Protocol_XS1000
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
End Sub

Private Sub PhaseCfg_Protocol_XS1000()
    Dim sWkDat$
    Dim i%
    
    For i = 1 To Len(wkBuf)
        sWkDat = Mid(wkBuf, i, 1)
        
        Select Case m_iPhase
            Case 1
                Select Case Asc(sWkDat)
                    Case 5
                    'ENQ 수신
                    
                        'ACK 송신
                        msComm.Output = Chr(6)
                        
                        m_iPhase = 2
                        
                        bETBChk = False
                    Case Else
                        m_iPhase = 1
                    End Select
            Case 2
                Select Case Asc(sWkDat)
                    Case 2
                    'STX 수신
                    
                        If bETBChk <> True Then
                            '초기화
                            RcvBuffer = ""
                        End If
                        bETBChk = False
                        
                    Case 10
                    'LF 수신
                        If bETBChk = False Then
                            Call DataEditResponse_XS1000
                            RcvBuffer = ""
                        End If
                        msComm.Output = Chr(6)
                    Case 13
                    'CR 수신
                        If bETBChk = False Then
                            Call DataEditResponse_XS1000
                            RcvBuffer = ""
                        End If
                            
                    Case 4
                    'EOT 수신
                        If sState = "Q" Then
                            'ENQ 송신
                            msComm.Output = Chr(5)
                            m_iSendPhase = 1
                        End If
                        m_iPhase = 3

                    Case 5
                    'ENQ 수신
                        bETBChk = False
                        
                        'ACK 송신
                        msComm.Output = Chr(6)

                    Case 21
                    'NAK 수신
                        Call DataEditResponse_XS1000
                        
                        m_iSendPhase = 1
                        m_iFrameN = 1
                        
                        'ENQ 송신
                        msComm.Output = Chr(5)

                    Case 23
                    'ETB 수신
                        bETBChk = True

                    Case Else
                        If bETBChk <> True Then
                            RcvBuffer = RcvBuffer & sWkDat
                        End If
                End Select

            Case 3
                Select Case Asc(sWkDat)
                    Case 6
                    'ACK 수신
                        If sState = "Q" Then
                            Call SendOrder_XS2100       'Order 전송
                        End If

                    Case 5      'ENQ
                        bETBChk = False
                        msComm.Output = Chr(6)
                        m_iPhase = 2

                    Case 21
                    'NAK 수신
                        m_iSendPhase = 1
                        m_iFrameN = 1
                        msComm.Output = Chr(5)
                        m_iPhase = 3

                    Case 4
                    'EOT 수신
                        m_iPhase = 1
                End Select
        End Select
    Next
End Sub

Private Sub DataEditResponse_XS1000()
    On Error GoTo ErrRtn
    
     
    Dim i As Integer
    
    Dim s_aRcvData() As String
    Dim sRcvData As String
    Dim sType As String
    
    
    Dim sID As String
    Dim sSeq As String
    Dim sRack As String
    Dim sPos As String
    Dim iCnt As Integer
    Dim sIFRstCd As String
    Dim sRst1 As String
    Dim sRst2 As String
    Dim sComment As String
    Dim sFlag As String
    Dim sUnit As String
    Dim sKind As String
     
    
     sType = Mid(RcvBuffer, 2, 1)


     Select Case sType
         Case "H"
         'Header Record
         
         '---------------------------------------
         '1H|\^&|||XS^00-01^11001^^^^12345678||||||||E1394 -97
         '---------------------------------------
         
         Case "Q"
         'Request Record
         
         '---------------------------------------
         '2Q|1|^^ 1234567890^B||||20011001153000
         '---------------------------------------
         
         sRcvData = Split(RcvBuffer, Chr(124))(2)
         
         sSeq = Split(RcvBuffer, Chr(124))(1)
         
         sID = Trim(Split(sRcvData, "^")(2))
         sRack = Trim(Split(sRcvData, "^")(0))
         sPos = Trim(Split(sRcvData, "^")(1))
         
         If (Trim(Split(sRcvData, "^")(3)) = "B") Then
            If sID = "" Then
                sState = ""
                pSampleInfo.ID = ""
            End If
         End If
         
         sState = "Q"
         
         pSampleInfo.ID = sID
         pSampleInfo.SEQNO = sSeq
         pSampleInfo.RACK = sRack
         pSampleInfo.POS = sPos
         
         Case "P"
         'Patient Record
         
         '---------------------------------------
         '2P|1|||100|^Heisei^Taro||20010820|M|||||^Dr.1||||||||||||^^^WEST
         '---------------------------------------
         
         Case "C"
         'Patient Comment
         '---------------------------------------
         '3C|1||patient_comments
         '---------------------------------------
         
         Case "O"
         'Order Record
         
         '---------------------------------------
         '4O|1||2^1^ 1234567890^B|^^^^WBC\^^^^RBC\^^^^HGB\^^^^HCT\^^^^MCV\^^^^MCH\^^^^MCHC\^^^^PLT\^^^^NEUT%\
         '^^^^LYMPH%\^^^^MONO%\^^^^EO%\^^^^BASO%\^^^^NEUT#\^^^^LYMPH#\^^^^MONO#\^^^^EO#\^^^^BASO#\^^^^RDW-SD\
         '^^^^RDW-CV\^^^^PDW\^^^^MPV\^^^^P-LCR\^^^^PCT|||||||N||||||||||||||F
         '---------------------------------------
         
         sRcvData = Split(RcvBuffer, Chr(124))(3)
         
         
         sID = Trim(Split(sRcvData, "^")(2))
         sRack = Trim(Split(sRcvData, "^")(0))
         sPos = Trim(Split(sRcvData, "^")(1))
         sKind = Split(RcvBuffer, Chr(124))(11)
         
         
         With pResultInfo
            .ID = sID
            .RACK = sRack
            .POS = sPos
            .KIND = sKind
         End With
         
         
         Case "R"
         'Result Recode
            
         '---------------------------------------
         '7R|1|^^^^WBC^1|7.81|10*3/uL||N||||||20010806120000
         '---------------------------------------
                  
         s_aRcvData = Split(RcvBuffer, Chr(124))
         
         sIFRstCd = Trim(Split(s_aRcvData(2), "^")(4))
         sRst1 = Trim(s_aRcvData(3))
         sFlag = Trim(s_aRcvData(6))
         sUnit = Trim(s_aRcvData(4))
         
         
         With pResultInfo
            .RSTCNT = .RSTCNT + 1
            .IFCD = .IFCD + sIFRstCd + Chr(124)
            .RST1 = .RST1 + sRst1 + Chr(124)
            .RST2 = .RST2 + sRst2 + Chr(124)
            .UNIT = .UNIT + sUnit + Chr(124)
            .FLAG = .FLAG + sFlag + Chr(124)
         End With
          
       
        Case "L"
         'Message Terminatior Record
         
        With pResultInfo
            If .RSTCNT > 0 Then
                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
            End If
        End With
        
        Call Init_pResultInfo
         
     End Select
     
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub
 

'< yjlee 2008-04-08
Private Sub SendOrder_XS2100()
    Dim sSendBuff   As String
    Dim sChkSum     As String
    Dim sTestDat    As String
    Dim i           As Integer
    Dim sReportType As String
    
    If m_iFrameN > 7 Then
        m_iFrameN = 0
    End If
    
    Select Case m_iSendPhase
        Case 1      'H
            sSendBuff = m_iFrameN & "H|\^&|||||||||||E1394-97" & Chr(13) & Chr(3)
            m_iSendPhase = 2
            sNextSend = ""

        Case 2      'P
            '----- 검사항목 조회
            If pSampleInfo.ID <> "" Then
                RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
            End If
            
            Call Get_OrderString
            
            If pSampleInfo.ORDCNT > 0 Then
                '오더 있는 경우
                sSendBuff = m_iFrameN & "P|1|||100|^Heisei^Taro||20010820|M|||||^Dr.1||||||||||||^^^WEST" & Chr(13) & Chr(3)
                m_iSendPhase = 3
            Else
                '오더 없는 경우
                sSendBuff = m_iFrameN & "P|1" & Chr(13) & Chr(3)
                m_iSendPhase = 4
            End If

        Case 3      'C(Patient comments)
            sSendBuff = m_iFrameN & "C|1||" & Chr(13) & Chr(3)
            m_iSendPhase = 4
            
        Case 4      'O
            sTestDat = ""
            '----- 검사항목 편집
            For i = 1 To pSampleInfo.ORDCNT
                sTestDat = sTestDat & "^^^^" & Trim(pSampleInfo.IFCD(i)) & "\"
            Next i
            If pSampleInfo.ORDCNT > 0 Then
                sTestDat = Left(sTestDat, Len(sTestDat) - 1)
                sReportType = "Q"
            Else
                sReportType = "Y"
            End If
            '-------------------

            sSendBuff = m_iFrameN & "O|1|" & pSampleInfo.RACK & "^" & pSampleInfo.POS & "^" & Right(Space(15) & Trim(pSampleInfo.ID), 15) & "^B||" _
                        & sTestDat & "||" & Format(Now, "YYYYMMDDHHNNSS") & "|||||N||||||||||||||" & sReportType
            
            '--- Text의 내용이 240byte를 넘어갈 경우 처리 추가...
            If Len(sSendBuff) >= 241 Then
                sNextSend = Mid(sSendBuff, 241)
                sSendBuff = Left(sSendBuff, 240)
                sSendBuff = sSendBuff & Chr(23)

                m_iSendPhase = 5
            Else
                sSendBuff = sSendBuff & Chr(13) & Chr(3)
                GoTo Send_Terminate
            End If
            
        Case 5
            sSendBuff = m_iFrameN & sNextSend
            If Len(sSendBuff) >= 241 Then
                sNextSend = Mid(sSendBuff, 241)
                sSendBuff = Left(sSendBuff, 240)
                sSendBuff = sSendBuff & Chr(23)
            Else
                sSendBuff = sSendBuff & Chr(13) & Chr(3)
                sNextSend = ""
Send_Terminate:
                If pSampleInfo.ORDCNT > 0 Then
                    m_iSendPhase = 6
                Else
                    m_iSendPhase = 7
                End If
            End If

        Case 6      'C(Specimen comments)
            sSendBuff = m_iFrameN & "C|1||" & Chr(13) & Chr(3)
            m_iSendPhase = 7
            
        Case 7      'T
            sSendBuff = m_iFrameN & "L|1|N" & vbCr & Chr(3)
            
            m_iSendPhase = 8

        Case 8      'EOT
            msComm.Output = Chr(4)   'EOT
            m_iFrameN = 1: m_iPhase = 1: m_iSendPhase = 1
            sState = ""

            '전송된 오더가 있는 경우 화면표시
            If pSampleInfo.ORDCNT > 0 Then
                RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)
            Else
                '조회된 내용이 없는 경우 환자정보 구조체 초기화
                Call Init_pResultInfo
        
                RaiseEvent SendOrderOK("", "", "")
            End If
            
            Exit Sub
    End Select

    'CheckSum 계산
    sChkSum = ChkSum_ASTM(sSendBuff)

    msComm.Output = Chr(2) & sSendBuff & sChkSum & Chr(13) & Chr(10)

    m_iFrameN = m_iFrameN + 1

    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & sChkSum & Chr(13) & Chr(10))
    End If

End Sub
'> yjlee 2008-04-08

Private Sub Get_OrderString()
    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .ORDCNT = 0
            Erase .IFCD
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

'결과정보 구조체 초기화
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
    m_SiteNm = PropBag.ReadProperty("SiteNm", m_def_SiteNm)
    m_p_sPatInfo = PropBag.ReadProperty("p_sPatInfo", m_def_p_sPatInfo)
    m_p_sSampInfo = PropBag.ReadProperty("p_sSampInfo", m_def_p_sSampInfo)
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
    Call PropBag.WriteProperty("SiteNm", m_SiteNm, m_def_SiteNm)
    Call PropBag.WriteProperty("p_sPatInfo", m_p_sPatInfo, m_def_p_sPatInfo)
    Call PropBag.WriteProperty("p_sSampInfo", m_p_sSampInfo, m_def_p_sSampInfo)
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
    m_SiteNm = m_def_SiteNm
    m_p_sPatInfo = m_def_p_sPatInfo
    m_p_sSampInfo = m_def_p_sSampInfo
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
Public Property Get SiteNm() As Variant
    SiteNm = m_SiteNm
End Property

Public Property Let SiteNm(ByVal New_SiteNm As Variant)
    m_SiteNm = New_SiteNm
    PropertyChanged "SiteNm"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sPatInfo() As String
    p_sPatInfo = m_p_sPatInfo
End Property

Public Property Let p_sPatInfo(ByVal New_p_sPatInfo As String)
    m_p_sPatInfo = New_p_sPatInfo
    PropertyChanged "p_sPatInfo"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sSampInfo() As String
    p_sSampInfo = m_p_sSampInfo
End Property

Public Property Let p_sSampInfo(ByVal New_p_sSampInfo As String)
    m_p_sSampInfo = New_p_sSampInfo
    PropertyChanged "p_sSampInfo"
End Property
