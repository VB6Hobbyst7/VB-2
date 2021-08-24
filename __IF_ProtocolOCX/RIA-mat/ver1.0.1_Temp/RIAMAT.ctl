VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl RIAMAT 
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   LockControls    =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   7245
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   870
      TabIndex        =   4
      Top             =   2175
      Width           =   660
   End
   Begin VB.TextBox txtSend 
      Height          =   3465
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   3
      Top             =   150
      Width           =   2325
   End
   Begin VB.TextBox txtLog 
      Height          =   3465
      Left            =   1785
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   2
      Top             =   150
      Width           =   2325
   End
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
      InputMode       =   1
   End
End
Attribute VB_Name = "RIAMAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
'Const m_def_iLenID = 0
Const m_def_iTotalItemCnt = 0
Const m_def_iOrderFlag = 0
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
'Dim m_iLenID As Integer
Dim m_iTotalItemCnt As Integer
Dim m_iOrderFlag As Integer
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
Event SendOrderOK(sID$, sSeqNo$, sRack$, sPos$)
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sOther1$)
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
'Event RequestCurOrder(sID$, sRack$, sPos$)
'Event SendOrderOK(sID$)
Event DispMsg(sMsg$)
'Event RequestNextOrder()
'Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)


'===== User Define
'인터페이스에서 사용
Dim RcvBuffer   As String
'Dim wkBuf   As String
Dim vwkBuf  As Variant
Dim sState  As String
Dim sReqStatusCd    As String

'구조체 지정
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'기타
Dim sOpenPW$, sEditPW$
Dim iSpaceCnt   As Integer

Private Sub PhaseCfg_Protocol_RIAMAT280()
    
    Dim ix1     As Integer
    Dim vSend
'    Dim bByte() As Byte

    Dim bAck(2) As Byte

    Dim sSend$
    Dim strTemp As String
    
    Dim vTmpBuf
    
    vTmpBuf = vwkBuf
    
    For ix1 = 0 To UBound(vTmpBuf) '- 1
        Select Case vTmpBuf(ix1)
            Case &H2        'STX
                RcvBuffer = ""
                
                'test
                txtLog = txtLog & Chr(2)
                
            Case &H3        'ETX
                Call Sleep(200)
                
                'test
                txtLog = txtLog & Chr(3)
                
'                vSend = ChrB(&H6)
'                bByte = vSend
'
'                msComm.Output = bByte

                bAck(0) = "&H6"
                msComm.Output = bAck
                
                Sleep (200)
                
                Call DataEditResponse_RIAMAT280
        
                RaiseEvent PrintRcvLog(RcvBuffer)
        
            Case &H6        'ACK
                'test
                txtLog = txtLog & Chr(6)
            
            Case Else   '문자 수신
                If vTmpBuf(ix1) >= &H80 Then
                    strTemp = "&H" & Hex(vTmpBuf(ix1))
                    ix1 = ix1 + 1
                    strTemp = strTemp & Hex(vTmpBuf(ix1))
                    RcvBuffer = RcvBuffer & Chr(Val(strTemp))
'                ElseIf vTmpBuf(ix1) = &H0 Then
                Else
                    RcvBuffer = RcvBuffer & Chr(vTmpBuf(ix1))
                End If
                
                'test
                txtLog = txtLog & Chr(vTmpBuf(ix1))

        End Select
    Next ix1
    
End Sub

Private Sub DataEditResponse_RIAMAT280()
    On Error GoTo ErrRtn

    Dim sMark$, sPNo$, sPID$
    Dim sRstData$
    Dim tmpIFCd$, tmpRst1$, tmpRst2$, tmpFlag$, tmpUnit$

    sMark = Mid$(RcvBuffer, 1, 1)

    Select Case sMark
        Case "I"        'Initialisation
'            Call Sleep(500)
            Call Sleep(300)
            Call Send_MarkI

        Case "N"        'Next patient
            sPNo = Mid(RcvBuffer, 2, 3)

            Call SendOrder_RIAMAT280(sPNo)

        Case "E"        'Result Set
            '결과정보 초기화
            Call Init_pResultInfo

            sPID = Trim(Mid(RcvBuffer, 2, 24))

            sRstData = Mid(RcvBuffer, 26)
            sRstData = Mid(sRstData, 1, Len(sRstData) - 2)

            Do Until sRstData = ""
                tmpIFCd = Mid(sRstData, 1, 4)
                tmpRst1 = Mid(sRstData, 5, 7)
                tmpFlag = Mid(sRstData, 12, 1)

                sRstData = Mid(sRstData, 13)

                '결과값 누적
                With pResultInfo
                    .RSTCNT = .RSTCNT + 1

                    .IFCD = .IFCD & Trim(tmpIFCd) & Chr(124)
                    .RST1 = .RST1 & Trim(tmpRst1) & Chr(124)
                    .RST2 = .RST2 & Chr(124)
                    .UNIT = .UNIT & Chr(124)
                    .FLAG = .FLAG & Trim(tmpFlag) & Chr(124)
                End With
            Loop

            '결과값 등록/화면 표시 처리...
            With pResultInfo
                .ID = sPID
                .SEQNO = "": .RACK = "": .POS = ""

                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, "")
                End If
            End With

            'Next Result Request
            Call Send_MarkW

        Case "S"        'End of List

        Case Else

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류 - (" & Err.Description & ")")
    End If
End Sub

'
'   Init Msg
'
Private Sub Send_MarkI()
    
'    Dim sSend$
'    Dim bSend() As Byte
'
'    sSend = Chr(2) & "I4;" & Chr(3)
'
'    Call fGetByteData(sSend, bSend)
'
'    'Init Msg
'    msComm.Output = bSend   'sSend
'
'    'Log 작성
'    If m_sTestMode = "77" Then
'        RaiseEvent PrintSendLog(sSend)
'    End If
    
'----------------
'''    If Trim(sData) <> "" Then
'''        For ii = 1 To Len(sData)
'''            sTmp = Mid(sData, ii, 1)
'''            vSend = vSend & ChrB(AscB(sTmp))
'''        Next ii
'''    End If
'''
'''    vSend = vCmd & vSend
'''
'''    Call ChkSum_CentaurLAS(vSend, vChkSum1, vChkSum2)
'''
'''    vSend = ChrB(&HF0) & vSend & ChrB("&h" & vChkSum1) & ChrB("&h" & vChkSum2) & ChrB(&HF8)
'''    bByte = vSend
'''
'''    Comm3.Output = bByte
'==================

    Dim sSend$
    Dim bSend(8) As Byte
    Dim ii%
    Dim sTmp$
    Dim vSend
    
    sSend = Chr(2) & "I4;" & Chr(3)
    
'    For ii = 1 To Len(sSend)
'        sTmp = Mid(sSend, ii, 1)
'        vSend = vSend & ChrB(AscB(sTmp))
'    Next ii

    bSend(0) = "&H2"
    bSend(1) = Asc("I")
    bSend(2) = Asc("4")
    bSend(3) = Asc(";")
    bSend(4) = "&H3"
'
''    Call fGetByteData(sSend, bSend)
    
    'Init Msg
'    msComm.Output = sSend
    msComm.Output = bSend   'sSend
    
    'Log 작성
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSend)
    End If
    
End Sub

'
'   Next Result Request
'
Private Sub Send_MarkW()

    Dim sSend$
    Dim bSend() As Byte

    sSend = Chr(2) & "W59" & Chr(3)
    
    Call fGetByteData(sSend, bSend)
    
    'Next Result Request
    msComm.Output = bSend   'sSend
    
    'Log 작성
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSend)
    End If
    
End Sub

Private Sub SendOrder_RIAMAT280(ByVal sPNo As String)
    On Error GoTo ErrSendOrd

    '환자의 Order 전송
'    Dim SendBuf As String
'    Dim sCS$
    
    Dim SendBuf$
    Dim sCS
    
    Dim sTestCd As String
    Dim ii      As Integer
    Dim iCnt    As Integer
    Dim tmpData()   As String

    Dim bSend() As Byte
    

    '현재 전송할 오더 조회
    RaiseEvent RequestCurOrder("", Trim(sPNo), "", "")
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .ORDCNT = 0
        End With
        
        Sleep (200)
        
        SendBuf = Chr(2) & "S55" & Chr(3)
        Call fGetByteData(SendBuf, bSend)
        
        msComm.Output = bSend       'Chr(2) & "S55" & Chr(3)

        'Log 작성
        If m_sTestMode = "77" Then
            RaiseEvent PrintSendLog(SendBuf)
        End If

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

    'Order 편집
    sTestCd = ""
    For ii = 1 To pSampleInfo.ORDCNT
        sTestCd = sTestCd & Left(pSampleInfo.IFCD(ii) & Space(4), 4)
    Next ii

'    'test
'    pSampleInfo.ID = Left(pSampleInfo.ID, 11)

    'Send Message 편집
    SendBuf = Chr(2)
    SendBuf = SendBuf & "P" & Right(Space(3) & Val(sPNo), 3)
'    SendBuf = SendBuf & Left(pSampleInfo.ID & Space(24), 24)
    SendBuf = SendBuf & LeftH(pSampleInfo.ID & Space(24), 24)
    SendBuf = SendBuf & sTestCd
        
    'CheckSum 계산
    sCS = ChkSum_RIAmat(SendBuf)
    
    SendBuf = SendBuf & sCS
    SendBuf = SendBuf & Chr(3)

    'test
'    SendBuf = "P  105121900011             AFP uE3 HCG 7="
    
'    ReDim bSend(LenH(SendBuf))
    
    'Byte 변환
    Call fGetByteData(SendBuf, bSend)

''    Dim iCnt%
'    Dim vSend
'    iCnt = UBound(bSend)
'    ReDim Preserve bSend(iCnt + 3)
'    bSend(iCnt - 2) = bSendCRC(0)
'    bSend(iCnt - 1) = bSendCRC(1)
'    vSend = ChrB(3)
'    bSend(iCnt) = vSend

'    Sleep (500)
    
    msComm.Output = bSend       'SendBuf & Chr(3)

    'Order 전송 완료
    RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, "", "")

    'Log 작성
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(SendBuf)
    End If

ErrSendOrd:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러발생 - " & Err.Description)
    End If
End Sub

Private Function ChkSum_RIAmat_2(ByVal Para As String) As String

    Dim i   As Integer
    Dim Tmp As Integer
    Dim ChkS1   As Integer
    Dim ChkS2   As String

    Dim sC1$, sC2$

    For i = 1 To Len(Para)
'    For i = 1 To LenB(Para)
        Tmp = Asc(Mid$(Para, i, 1))
'        Tmp = Asc(MidB$(Para, i, 1))
        ChkS1 = ChkS1 + Tmp
    Next i
    ChkS1 = ChkS1 Mod 256
    ChkS2 = Right$("0" & Hex$(ChkS1), 2)

    sC1 = Mid(ChkS2, 1, 1)
    sC1 = "3" & sC1
    sC1 = CDec("&H" & sC1)
    sC1 = Chr(Val(sC1))
    sC1 = Right(sC1, 1)

    sC2 = Mid(ChkS2, 2, 1)
    sC2 = "3" & sC2
    sC2 = CDec("&H" & sC2)
    sC2 = Chr(Val(sC2))
    sC2 = Right(sC2, 1)
'
'    ChkSum_RIAmat = sC1 & sC2
    
End Function
Private Function ChkSum_RIAmat_Err(ByVal sPara As String) As String
    On Error GoTo ErrTemp
    
    Dim i   As Integer
    Dim sC1, sC2
    
    Dim sHexaStr$
    Dim vChkS, vCS1, vCS2
    Dim sChkS$
    
    'Hexa값으로 변환
    sHexaStr = fGetHexaCode(sPara)
    
    vChkS = 0
    For i = 1 To Len(sHexaStr) Step 2
        vChkS = vChkS + ("&H" & Mid(sHexaStr, i, 2))
    Next i
    
    vChkS = vChkS Mod 256
    sChkS = Hex(vChkS)
    sChkS = Right("00" & sChkS, 2)
    
    vCS1 = "&H3" & Mid(sChkS, 1, 1)
    vCS2 = "&H3" & Mid(sChkS, 2, 1)
    
    sC1 = Chr(CDec(vCS1))
    sC2 = Chr(CDec(vCS2))
    
    ChkSum_RIAmat = sC1 & sC2
    
ErrTemp:
    If Err <> 0 Then
        Resume Next
    End If
End Function
Private Function ChkSum_RIAmat(ByVal sPara As String) As String
    On Error GoTo ErrTemp
    
    Dim i   As Integer
    Dim sC1, sC2
    
    Dim sHexaStr$
    Dim vChkS, vCS1, vCS2
    Dim sChkS$
    
    'Hexa값으로 변환
    sHexaStr = fGetHexaCode(sPara)
    
    vChkS = 0
    For i = 1 To Len(sHexaStr) Step 2
        vChkS = vChkS + ("&H" & Mid(sHexaStr, i, 2))
    Next i
    
    vChkS = vChkS Mod 256
    sChkS = Hex(vChkS)
    
    If Len(sChkS) < 2 Then
        vCS1 = "&H0"
        vCS2 = "&H" & sChkS
    Else
        vCS1 = "&H" & Mid(sChkS, 1, 1)
        vCS2 = "&H" & Mid(sChkS, 2, 1)
    End If
    
'    sChkS = Right("00" & sChkS, 2)
'
    vCS1 = "&H3" & Mid(sChkS, 1, 1)
    vCS2 = "&H3" & Mid(sChkS, 2, 1)
    
'    bSendCRC(0) = vCS1
'    bSendCRC(1) = vCS2
    
        
    
    sC1 = Chr(CDec(vCS1))
    sC2 = Chr(CDec(vCS2))

    ChkSum_RIAmat = sC1 & sC2
    
ErrTemp:
    If Err <> 0 Then
        Resume Next
    End If
End Function
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
        Case "RIAMAT280"
            Call PhaseCfg_Protocol_RIAMAT280
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
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
            .IFCD(ii) = tmpData(ii - 1)
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

    vwkBuf = Text1
    Call PhaseCfg_Protocol

End Sub

Private Sub Command1_Click()

    Call Send_MarkI
'    Call ChkSum_RIAmat("TEST")
    
End Sub

Private Sub msComm_OnComm()
        
    Select Case msComm.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            vwkBuf = msComm.Input
            
'            If sTestMode = "77" Then
'                RaiseEvent PrintRcvLog(Fu_Read_Name(vwkBuf))
'            End If
                                
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
    m_iOrderFlag = PropBag.ReadProperty("iOrderFlag", m_def_iOrderFlag)
    m_iTotalItemCnt = PropBag.ReadProperty("iTotalItemCnt", m_def_iTotalItemCnt)
'    m_iLenID = PropBag.ReadProperty("iLenID", m_def_iLenID)
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
    Call PropBag.WriteProperty("iOrderFlag", m_iOrderFlag, m_def_iOrderFlag)
    Call PropBag.WriteProperty("iTotalItemCnt", m_iTotalItemCnt, m_def_iTotalItemCnt)
'    Call PropBag.WriteProperty("iLenID", m_iLenID, m_def_iLenID)
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
'        With MSComm
'            .OutBufferSize = 1024
'            '- 송신 버퍼의 내용을 초기화합니다.
'            MSComm.OutBufferCount = 0
'            '- 수신 버퍼의 크기를 설정합니다.
'            MSComm.InBufferSize = 1024
'            '- 수신 버퍼의 내용을 초기화합니다.
'            MSComm.InBufferCount = 0
'            '- 수신시에 검색하는 데이터 형식을 설정.
'            MSComm.InputMode = comInputModeBinary
'            '- 1 Byte 수신 때마다 이벤트(OnComm) 발생을 설정.
'            MSComm.RThreshold = 1
'            '- 1 Byte식 데이터를 버퍼에서 읽도록 설정.
'            MSComm.InputLen = 0
'        End With
        
        msComm.PortOpen = True
    End If
    On Error GoTo 0
ErrPortOpen:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
        RaiseEvent RaiseError("PortOpen Error!!! " & Err.Description)
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
    m_iOrderFlag = m_def_iOrderFlag
    m_iTotalItemCnt = m_def_iTotalItemCnt
'    m_iLenID = m_def_iLenID
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

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iOrderFlag() As Integer
    iOrderFlag = m_iOrderFlag
End Property

Public Property Let iOrderFlag(ByVal New_iOrderFlag As Integer)
    m_iOrderFlag = New_iOrderFlag
    PropertyChanged "iOrderFlag"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iTotalItemCnt() As Integer
    iTotalItemCnt = m_iTotalItemCnt
End Property

Public Property Let iTotalItemCnt(ByVal New_iTotalItemCnt As Integer)
    m_iTotalItemCnt = New_iTotalItemCnt
    PropertyChanged "iTotalItemCnt"
End Property

