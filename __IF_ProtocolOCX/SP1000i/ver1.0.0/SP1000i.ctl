VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl SP1000i 
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
Attribute VB_Name = "SP1000i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_Rst_HCT = 0
Const m_def_Rst_WBC = 0
Const m_def_Rst_RBC = 0
Const m_def_No_Film = 0
Const m_def_PrtText1 = 0
Const m_def_PrtText2 = 0
Const m_def_PrtText3 = 0
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
Dim m_Rst_HCT As Variant
Dim m_Rst_WBC As Variant
Dim m_Rst_RBC As Variant
Dim m_No_Film As Variant
Dim m_PrtText1 As Variant
Dim m_PrtText2 As Variant
Dim m_PrtText3 As Variant
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
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sErrCd$, sKind$, sTRstDT$, sOther1$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event SendOrderOK(sID$, sRack$, sPos$, sState$)
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
Dim sOpenPW$, sEditPW$
Dim iSpaceCnt   As Integer

'for XE-2100/SE-9000
Dim miFlagCnt   As Integer
Dim msFlagCd  As String
Dim msFlagTot   As String

'for XT1800
Dim msOrdS2     As String

'for SP-1000i
Private pSlideInfo  As SLIDE_INFO


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
        Case "SP1000I"
            Call PhaseCfg_Protocol_SP1000i
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub


Private Sub SendOrder_SP1000i()
    On Error GoTo ErrRtn

    Dim sOrder$
    Dim iPos%, i%
    Dim sBuf$
    
    RaiseEvent RequestCurOrder(pSampleInfo.ID, "", "", "")

    Call Get_OrderString
    
    If pSlideInfo.NO_FILM = 0 Then
        RaiseEvent DispMsg("SLIDE 오더 항목이 존재하지 않습니다!!")
    End If

    sBuf = "S3"
    sBuf = sBuf & Right(Space(15) & pSampleInfo.ID, 15)
    'HCT
    If Trim(Rst_HCT) = "" Then
        sBuf = sBuf & "****"
    Else
        sBuf = sBuf & Right("0000" & Replace(Rst_HCT, ".", ""), 4)
    End If
    'WBC
    If Trim(Rst_WBC) = "" Then
        sBuf = sBuf & "******"
    Else
        sBuf = sBuf & Right("000000" & Replace(Rst_WBC, ".", ""), 6)
    End If
    'RBC
    If Trim(Rst_RBC) = "" Then
        sBuf = sBuf & "*****"
    Else
        sBuf = sBuf & Right("000000" & Replace(Rst_RBC, ".", ""), 5)
    End If
    
    'No of Films
'    sBuf = sBuf & "1"   '0"   '"1"
    sBuf = sBuf & Trim(pSlideInfo.NO_FILM)
    
    'print text1
'    sBuf = sBuf & Space(15)
    sBuf = sBuf & Left(PrtText1 & Space(15), 15)
    
    'print text2
    sBuf = sBuf & Left(PrtText2 & Space(15), 15)
    'print text3
    sBuf = sBuf & Left(PrtText3 & Space(15), 15)
    
    sBuf = sBuf & "000000"
    sBuf = sBuf & "00"
    
    'first slide glass
    sBuf = sBuf & "0"
    'second slide glass
    sBuf = sBuf & "0"
    
    'order reason
    sBuf = sBuf & "2"
    
    sBuf = sBuf & Space(15)     'print text4
    sBuf = sBuf & Space(15)     'print text5
    sBuf = sBuf & Space(15)     'print text6
    
    sBuf = sBuf & Space(2)
    
    'print information1,2
    sBuf = sBuf & Space(100)

    sBuf = Chr(2) & sBuf & Chr(3)
    
    msComm.Output = sBuf

    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sBuf)
    End If
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러 - " & Err.Description)
    End If
End Sub

Private Sub DataEdit_SP1000i()
    On Error GoTo ErrRtn

    Dim sBC$, sLC$
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$, tmpRstDt$, tmpState$
    Dim ii%
    Dim tmpRst()    As String       '결과 임시 저장
    Dim sRstDT$, sErrCd$, tmpErrCd$
    Dim sFlagBuf$
    
    sBC = Mid(RcvBuffer, 1, 2)
    sLC = Mid(RcvBuffer, 3, 1)

    Select Case sBC
        Case "R2"
            'R21     117362295100000000B
            With pSampleInfo
                .ID = Trim(Mid(RcvBuffer, 4, 15))
            End With

            Call SendOrder_SP1000i
            
            m_iPhase = 2

            Exit Sub

        Case "D1"   '밀기
            'D1520101U20111031143400022404     11736229510371111-10-31       O-02517        Yoo.HS         0                                                                           
            tmpRstDt = Trim(Mid(RcvBuffer, 10, 12))
            tmpRack = Trim(Mid(RcvBuffer, 22, 6))
            tmpPos = Trim(Mid(RcvBuffer, 28, 2))
            tmpBarCd = Trim(Mid(RcvBuffer, 30, 15))
            
''            "0": Smear was completed normally.
''            "1": Smear was interrupted or terminated due to the hardware error.
''            "2": Smear was not created due to no blood.
''            "3": Smear was not created due to a cancellation.
            tmpState = Trim(Mid(RcvBuffer, 95, 1))
            
            '결과정보 구조체에 저장
            With pSampleInfo
                .ID = tmpBarCd
                .SEQNO = ""
                .RACK = tmpRack
                .POS = tmpPos
            End With
            
            RaiseEvent SendOrderOK(pSampleInfo.ID, tmpRack, tmpPos, tmpState)

        Case "D2"   '염색
            'D2520101U20111031145200022404     11736229510371111-10-31       O-02517        Yoo.HS         0                                                                           
            tmpRstDt = Trim(Mid(RcvBuffer, 10, 12))
            tmpRack = Trim(Mid(RcvBuffer, 22, 6))
            tmpPos = Trim(Mid(RcvBuffer, 28, 2))
            tmpBarCd = Trim(Mid(RcvBuffer, 30, 15))
            
''            "0": Staining was completed normally.
''            "1": Staining was interrupted or terminated due to the hardware error.
            tmpState = Trim(Mid(RcvBuffer, 95, 1))
            
            '결과정보 구조체에 저장
            With pSampleInfo
                .ID = tmpBarCd
                .SEQNO = ""
                .RACK = tmpRack
                .POS = tmpPos
            End With
            
            RaiseEvent SendOrderOK(pSampleInfo.ID, tmpRack, tmpPos, tmpState)
            
        Case Else
    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 에러 발생 - " & Err.Description)
    End If
End Sub

Private Sub PhaseCfg_Protocol_SP1000i()

    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case m_iPhase
            Case 1
                Select Case Asc(wkDat)
                    Case 2      'STX
                        RcvBuffer = ""
                    
                    Case 3      'ETX
                        msComm.Output = Chr(6)       'ACK
                        
                        Call DataEdit_SP1000i
                        
                        RcvBuffer = ""
                        
                    Case Else
                        RcvBuffer = RcvBuffer & wkDat
                End Select
                
            Case 2
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        RaiseEvent SendOrderOK(pSampleInfo.ID, "", "", "")
                        
                        'Order를 보내고 다시 초기 상태
                        m_iPhase = 1
                        m_iOrderFlag = 0
                        
                    Case 21
                        Call SendOrder_SP1000i
                    
                    Case Else
                        m_iPhase = 1
                        m_iOrderFlag = 0
                End Select
        End Select
    Next ix1
    
End Sub

Private Sub Get_OrderString()
    
    If m_p_sID = "" Or m_No_Film = 0 Then
        With pSlideInfo
            .ID = m_p_sID
            .NO_FILM = 0
            
            .HCT = "": .WBC = "": .RBC = Rst_RBC
            .PRT1 = "": .PRT2 = "": .PRT3 = ""
        End With
        
        Exit Sub
    End If
    
    With pSlideInfo
        .ID = m_p_sID
        
        If m_No_Film > 3 Then
            m_No_Film = 1
        End If
        
        .NO_FILM = m_No_Film      'm_p_iOrdCnt
        
        .HCT = Rst_HCT
        .WBC = Rst_WBC
        .RBC = Rst_RBC
        
        .PRT1 = PrtText1
        .PRT2 = PrtText2
        .PRT3 = PrtText3
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
        .KIND = ""
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
    On Error GoTo ErrRtn
    
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
    
ErrRtn:
    If Err <> 0 Then
        MsgBox Err.Description
    End If
    m_Rst_HCT = PropBag.ReadProperty("Rst_HCT", m_def_Rst_HCT)
    m_Rst_WBC = PropBag.ReadProperty("Rst_WBC", m_def_Rst_WBC)
    m_Rst_RBC = PropBag.ReadProperty("Rst_RBC", m_def_Rst_RBC)
    m_No_Film = PropBag.ReadProperty("No_Film", m_def_No_Film)
    m_PrtText1 = PropBag.ReadProperty("PrtText1", m_def_PrtText1)
    m_PrtText2 = PropBag.ReadProperty("PrtText2", m_def_PrtText2)
    m_PrtText3 = PropBag.ReadProperty("PrtText3", m_def_PrtText3)
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
    Call PropBag.WriteProperty("Rst_HCT", m_Rst_HCT, m_def_Rst_HCT)
    Call PropBag.WriteProperty("Rst_WBC", m_Rst_WBC, m_def_Rst_WBC)
    Call PropBag.WriteProperty("Rst_RBC", m_Rst_RBC, m_def_Rst_RBC)
    Call PropBag.WriteProperty("No_Film", m_No_Film, m_def_No_Film)
    Call PropBag.WriteProperty("PrtText1", m_PrtText1, m_def_PrtText1)
    Call PropBag.WriteProperty("PrtText2", m_PrtText2, m_def_PrtText2)
    Call PropBag.WriteProperty("PrtText3", m_PrtText3, m_def_PrtText3)
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
    
    m_iOrderFlag = 0
    
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
    m_Rst_HCT = m_def_Rst_HCT
    m_Rst_WBC = m_def_Rst_WBC
    m_Rst_RBC = m_def_Rst_RBC
    m_No_Film = m_def_No_Film
    m_PrtText1 = m_def_PrtText1
    m_PrtText2 = m_def_PrtText2
    m_PrtText3 = m_def_PrtText3
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

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get Rst_HCT() As Variant
    Rst_HCT = m_Rst_HCT
End Property

Public Property Let Rst_HCT(ByVal New_Rst_HCT As Variant)
    m_Rst_HCT = New_Rst_HCT
    PropertyChanged "Rst_HCT"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get Rst_WBC() As Variant
    Rst_WBC = m_Rst_WBC
End Property

Public Property Let Rst_WBC(ByVal New_Rst_WBC As Variant)
    m_Rst_WBC = New_Rst_WBC
    PropertyChanged "Rst_WBC"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get Rst_RBC() As Variant
    Rst_RBC = m_Rst_RBC
End Property

Public Property Let Rst_RBC(ByVal New_Rst_RBC As Variant)
    m_Rst_RBC = New_Rst_RBC
    PropertyChanged "Rst_RBC"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get NO_FILM() As Variant
    NO_FILM = m_No_Film
End Property

Public Property Let NO_FILM(ByVal New_No_Film As Variant)
    m_No_Film = New_No_Film
    PropertyChanged "No_Film"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get PrtText1() As Variant
    PrtText1 = m_PrtText1
End Property

Public Property Let PrtText1(ByVal New_PrtText1 As Variant)
    m_PrtText1 = New_PrtText1
    PropertyChanged "PrtText1"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get PrtText2() As Variant
    PrtText2 = m_PrtText2
End Property

Public Property Let PrtText2(ByVal New_PrtText2 As Variant)
    m_PrtText2 = New_PrtText2
    PropertyChanged "PrtText2"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get PrtText3() As Variant
    PrtText3 = m_PrtText3
End Property

Public Property Let PrtText3(ByVal New_PrtText3 As Variant)
    m_PrtText3 = New_PrtText3
    PropertyChanged "PrtText3"
End Property

