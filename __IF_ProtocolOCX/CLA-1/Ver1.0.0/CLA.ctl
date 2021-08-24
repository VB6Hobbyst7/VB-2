VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl CLA 
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
   Begin MSCommLib.MSComm Comm 
      Left            =   255
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "CLA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mc_sSampleType    As String = "1"
Private Const mc_sPatInfo       As String = ""
Private Const mc_sSampInfo      As String = ""
Private Const mc_sSiteNm        As String = ""
Private Const mc_sRerunGbn      As String = ""
Private Const mc_bSerumIndex    As Integer = False
Private Const mc_sEqName        As String = ""
Private Const mc_bUseBarcode    As Boolean = False
Private Const mc_iPhase         As Integer = 1
Private Const mc_iSendPhase     As Integer = 1
Private Const mc_sTestMode      As String = "0"
Private Const mc_iFrameN        As Integer = 1
Private Const mc_sID            As String = ""
Private Const mc_sSeq           As String = ""
Private Const mc_sRack          As String = ""
Private Const mc_sPos           As String = ""
Private Const mc_iOrdCnt        As Integer = 0
Private Const mc_sTIFCd         As String = ""
Private Const mc_bPortOpen      As Boolean = False
Private Const mc_sOpenPW        As String = ""
Private Const mc_sEditPW        As String = ""

'속성 변수:
Dim mp_sSampleType          As String
Dim mp_sPatInfo             As String
Dim mp_sSampInfo            As String
Dim mp_sSiteNm              As String
Dim mp_sRerunGbn            As String
Dim mp_bSerumIndex          As Boolean
Dim mp_sEqName              As String
Dim mp_bUseBarcode          As Boolean
Dim mp_iPhase               As Integer
Dim mp_iSendPhase           As Integer
Dim mp_sTestMode            As String
Dim mp_iFrameN              As Integer
Dim mp_sID                  As String
Dim mp_sSeq                 As String
Dim mp_sRack                As String
Dim mp_sPos                 As String
Dim mp_iOrdCnt              As Integer
Dim mp_sTIFCd               As String
Dim mp_bPortOpen            As Boolean
Dim mp_sOpenPW              As String
Dim mp_sEditPW              As String

'이벤트 선언:
Public Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTRstDt$)
Public Event RequestCurOrder(sID$, sRack$, sPos$)
Public Event SendOrderOK(sID$, sSeqNo$, sRack$, sPos$)
Public Event RaiseError(sError$)
Public Event PrintRcvLog(sLog$)
Public Event PrintSendLog(sLog$)
Public Event DispMsg(sMsg$)

'< EqOcx Common
Dim msRcvBuffer       As String
Dim msWkBuf           As String

Dim miSpaceCnt        As Integer

Dim m_Sample_Info     As SAMPLE_INFO
Dim m_Result_Info     As RESULT_INFO
'> EqOcx Common

'< CLA
Private msSndH            As String
Private msSndP            As String
Private msSndO            As String
Private msSndL            As String

Private msRcvState        As String
Private msSndState        As String
'> CLA

Public Property Get CommPort() As Integer
Attribute CommPort.VB_Description = "통신 포트 번호를 반환하거나 설정합니다."
    CommPort = Comm.CommPort
End Property

Public Property Let CommPort(ByVal New_CommPort As Integer)
    Comm.CommPort() = New_CommPort
    PropertyChanged "CommPort"
End Property

Public Property Get EditPW() As String
    EditPW = mp_sEditPW
End Property

Public Property Let EditPW(ByVal New_EditPW As String)
    mp_sEditPW = New_EditPW
    PropertyChanged "EditPW"
End Property

Public Property Get EqName() As String
    EqName = mp_sEqName
End Property

Public Property Let EqName(ByVal New_EqName As String)
    mp_sEqName = New_EqName
    PropertyChanged "EqName"
End Property

Public Property Get FrameNo() As Integer
    FrameNo = mp_iFrameN
End Property

Public Property Let FrameNo(ByVal New_FrameNo As Integer)
    mp_iFrameN = New_FrameNo
    PropertyChanged "FrameNo"
End Property

Public Property Get ID() As String
    ID = mp_sID
End Property

Public Property Let ID(ByVal New_ID As String)
    mp_sID = New_ID
    PropertyChanged "ID"
End Property

Public Property Get OpenPW() As String
    OpenPW = mp_sOpenPW
End Property

Public Property Let OpenPW(ByVal New_OpenPW As String)
    mp_sOpenPW = New_OpenPW
    PropertyChanged "OpenPW"
End Property

Public Property Get OrderCnt() As Integer
    OrderCnt = mp_iOrdCnt
End Property

Public Property Let OrderCnt(ByVal New_OrderCnt As Integer)
    mp_iOrdCnt = New_OrderCnt
    PropertyChanged "OrderCnt"
End Property

Public Property Get PatientInfo() As String
    PatientInfo = mp_sPatInfo
End Property

Public Property Let PatientInfo(ByVal New_PatientInfo As String)
    mp_sPatInfo = New_PatientInfo
    PropertyChanged "PatientInfo"
End Property

Public Property Get Phase() As Integer
    Phase = mp_iPhase
End Property

Public Property Let Phase(ByVal New_Phase As Integer)
    mp_iPhase = New_Phase
    PropertyChanged "Phase"
End Property

Public Property Get PortOpen() As Boolean
    PortOpen = mp_bPortOpen
End Property

Public Property Let PortOpen(ByVal New_PortOpen As Boolean)
    mp_bPortOpen = New_PortOpen
    PropertyChanged "PortOpen"
    
    '--- PortOpen시 암호 확인
    If mp_sOpenPW <> gcOpenPW Then
        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
        Exit Property
    End If
    '-----------------------
    
    gsSiteNm = mp_sSiteNm
    
    On Error GoTo ErrPortOpen
    
    If mp_bPortOpen = True Then
        Comm.PortOpen = True
    Else
        Comm.PortOpen = False
    End If
    
ErrPortOpen:
    If Err <> 0 Then
        MsgBox "PortOpen Error!!! " & Err.Description, vbCritical
        RaiseEvent DispMsg(Err.Description)
    End If
End Property

Public Property Get POS() As String
    POS = mp_sPos
End Property

Public Property Let POS(ByVal New_Pos As String)
    mp_sPos = New_Pos
    PropertyChanged "Pos"
End Property

Public Property Get RACK() As String
    RACK = mp_sRack
End Property

Public Property Let RACK(ByVal New_Rack As String)
    mp_sRack = New_Rack
    PropertyChanged "Rack"
End Property

Public Property Get RerunGbn() As String
    RerunGbn = mp_sRerunGbn
End Property

Public Property Let RerunGbn(ByVal New_p_sRerunGbn As String)
    mp_sRerunGbn = New_p_sRerunGbn
    PropertyChanged "RerunGbn"
End Property

Public Property Get RTSEnable() As Boolean
    RTSEnable = Comm.RTSEnable
End Property

Public Property Let RTSEnable(ByVal New_RTSEnable As Boolean)
    Comm.RTSEnable = New_RTSEnable
    PropertyChanged "RTSEnable"
End Property

Public Property Get RThreshold() As Integer
    RThreshold = Comm.RThreshold
End Property

Public Property Let RThreshold(ByVal New_RThreshold As Integer)
    Comm.RThreshold = New_RThreshold
    PropertyChanged "RThreshold"
End Property

Public Property Get SampleInfo() As String
    SampleInfo = mp_sSampInfo
End Property

Public Property Let SampleInfo(ByVal New_SampleInfo As String)
    mp_sSampInfo = New_SampleInfo
    PropertyChanged "SampleInfo"
End Property

Public Property Get SampleType() As String
    SampleType = mp_sSampleType
End Property

Public Property Let SampleType(ByVal New_SampleType As String)
    mp_sSampleType = New_SampleType
    PropertyChanged "SampleType"
End Property

Public Property Get SendPhase() As Integer
    SendPhase = mp_iSendPhase
End Property

Public Property Let SendPhase(ByVal New_SendPhase As Integer)
    mp_iSendPhase = New_SendPhase
    PropertyChanged "SendPhase"
End Property

Public Property Get SEQNO() As String
    SEQNO = mp_sSeq
End Property

Public Property Let SEQNO(ByVal New_SeqNo As String)
    mp_sSeq = New_SeqNo
    PropertyChanged "SeqNo"
End Property
'
'Public Property Get SerumIndex() As Boolean
'    SerumIndex = mp_bSerumIndex
'End Property
'
'Public Property Let SerumIndex(ByVal New_SerumIndex As Boolean)
'    mp_bSerumIndex = New_SerumIndex
'    PropertyChanged "SerumIndex"
'End Property

Public Property Get Settings() As String
    Settings = Comm.Settings
End Property

Public Property Let Settings(ByVal New_Settings As String)
    Comm.Settings() = New_Settings
    PropertyChanged "Settings"
End Property

Public Property Get SiteNm() As String
    SiteNm = mp_sSiteNm
End Property

Public Property Let SiteNm(ByVal New_SiteNm As String)
    mp_sSiteNm = New_SiteNm
    PropertyChanged "SiteNm"
End Property

Public Property Get TestMode() As String
    TestMode = mp_sTestMode
End Property

Public Property Let TestMode(ByVal New_TestMode As String)
    mp_sTestMode = New_TestMode
    PropertyChanged "TestMode"
End Property

Public Property Get TotIFCd() As String
    TotIFCd = mp_sTIFCd
End Property

Public Property Let TotIFCd(ByVal New_TotIFCd As String)
    mp_sTIFCd = New_TotIFCd
    PropertyChanged "TotIFCd"
End Property

Public Property Get UseBarcode() As Boolean
    UseBarcode = mp_bUseBarcode
End Property

Public Property Let UseBarcode(ByVal New_UseBarcode As Boolean)
    mp_bUseBarcode = New_UseBarcode
    PropertyChanged "UseBarcode"
End Property

Private Sub PhaseCfg_Protocol()
    '--- 사용자 확인
    If mp_sEditPW <> gcEditPW Then
        MsgBox "등록된 사용자가 아닙니다. (주)에이씨케이로 문의해 주십시오!!!", vbCritical, "사용자 확인"
        Exit Sub
    End If
    '---------------
    
    If mp_sEqName = "0" Or mp_sEqName = "" Then
        RaiseEvent DispMsg("검사장비명을 지정해 주십시오.!!!")
        Exit Sub
    End If
    
    Select Case UCase(mp_sEqName)
        Case "CLA"
            Call PhaseCfg_Protocol_EQ
        
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
End Sub

Private Sub PhaseCfg_Protocol_EQ()
    Dim sWkDat   As String
    Dim i   As Integer
            
    For i = 1 To Len(msWkBuf)
        sWkDat = Mid(msWkBuf, i, 1)
        
        If sWkDat = "T" Then
            msRcvBuffer = msRcvBuffer & sWkDat
            
            If Right(msRcvBuffer, 10) = "BEGIN TEST" Then
                msRcvBuffer = "BEGIN TEST"
            ElseIf Right(msRcvBuffer, 8) = "END TEST" Then
                DataEditResponse
            End If
        Else
            msRcvBuffer = msRcvBuffer & sWkDat
        End If
    Next
End Sub

Private Sub DataEditResponse()
    On Error GoTo ErrRtn
    
    Dim sRxData As String
    
    Dim sIFCd As String
    Dim sRst As String
    Dim sRst2 As String
    Dim sRstDt As String
    
    Dim a_sBuf() As String
    
    Dim sPanelNo As String
    Dim iITC As Integer
    Dim iNLAC As Integer
    
    sRxData = msRcvBuffer
    
    '초기화
    sPanelNo = ""
    iITC = 0
    iNLAC = 0
    
    Call Init_m_Result_Info
    
    '<CR>로 Split
    a_sBuf = Split(sRxData, Chr(13))
    
    Dim i As Integer
    
    For i = 1 To UBound(a_sBuf) + 1
        If Left(a_sBuf(i - 1), Len("BEGIN TEST")) = "BEGIN TEST" Then
            GoTo continue
        End If
        
        If Trim(a_sBuf(i - 1)) = "" Then
            GoTo continue
        End If
        
        If Left(a_sBuf(i - 1), Len("DEVICE DIAGNOSTICS:")) = "DEVICE DIAGNOSTICS:" Then
            '이전 결과 존재 시 등록
            With m_Result_Info
                If .RSTCNT > 0 And .IFCD <> "" And .RST1 <> "" Then
                    If .ID = "" Then .ID = Format(Now, "HH:mm:ss")
                    
                    If .RSTCNT > 0 Then
                        .UNIT = String(.RSTCNT, Chr(124))
                        .FLAG = String(.RSTCNT, Chr(124))
                
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .RSTDT)
                    End If
                End If
                
                '초기화
                sPanelNo = ""
                iITC = 0
                iNLAC = 0
                
                Call Init_m_Result_Info
                
                If i > 2 Then
                    .SEQNO = Mid(a_sBuf(i - 3), 16, 8)
                End If
            End With
        End If
        
        If Left(a_sBuf(i - 1), Len("PANEL:")) = "PANEL:" Then
            sPanelNo = Replace(Trim(Mid(a_sBuf(i - 1), 7)), vbTab, "")
        End If
        
        If Left(a_sBuf(i - 1), Len("INTERNAL TEST CHECKS:")) = "INTERNAL TEST CHECKS:" Then
            iITC = i
        End If
        
        If iITC = 0 Then
            GoTo continue
        End If
        
        If i = iITC + 1 Or i = iITC + 2 Then
            sIFCd = sPanelNo + Mid(a_sBuf(i - 1), 1, 1)
            
            sRst = Trim(Mid(a_sBuf(i - 1), 11, 5))
            
            If Right(a_sBuf(i - 1), Len("PASS")) = "PASS" Then
                sRst2 = ""
            Else
                sRst2 = Mid(a_sBuf(i - 1), 20, 4)
            End If
            
            sRstDt = Format(Now, "yyyyMMddHHmmss")
            
            If sIFCd <> "" And sRst <> "" Then
                With m_Result_Info
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD + sIFCd + Chr(124)
                    .RST1 = .RST1 + sRst + Chr(124)
                    .RST2 = .RST2 + sRst2 + Chr(124)
                    .RSTDT = .RSTDT + sRstDt + Chr(124)
                End With
            End If
        End If
        
        If Left(a_sBuf(i - 1), Len("No. LU  Allergen  Class")) = "No. LU  Allergen  Class" Then
            iNLAC = i
        End If
        
        If iNLAC = 0 Then
            GoTo continue
        End If
        
        If i <= iNLAC + 1 Then
            GoTo continue
        End If
        
        'STX에 해당하는 DEVICE DIAGNOSTICS: 이전에 11-JUL-2005 처럼 날짜가 넘어와서
        '샘플결과처럼 해석될 수 있으므로
        '2 Line 앞이 DEVICE DIAGNOSTICS: 로 시작하면 continue
        If i <= UBound(a_sBuf) + 1 - 2 Then
            If Left(a_sBuf(i + 1), Len("DEVICE DIAGNOSTICS:")) = "DEVICE DIAGNOSTICS:" Then
                GoTo continue
            End If
        End If
            
        If Left(a_sBuf(i - 1), Len("END TEST")) = "END TEST" Then
            With m_Result_Info
                If .RSTCNT < 1 Then GoTo continue
                
                If .ID = "" Then .ID = Format(Now, "HH:mm:ss")
                
                If .RSTCNT > 0 Then
                    .UNIT = String(.RSTCNT, Chr(124))
                    .FLAG = String(.RSTCNT, Chr(124))
            
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .RSTDT)
                End If
                
                '초기화
                sPanelNo = ""
                iITC = 0
                iNLAC = 0
                
                Call Init_m_Result_Info
            End With
        Else
            sIFCd = sPanelNo + Trim(Mid(a_sBuf(i - 1), 1, 2))
            
            sRst = Left(Trim(Mid(a_sBuf(i - 1), 3, 4)) + Space(5), 5) + _
                    Right(Space(8) + "( " + Trim(Mid(a_sBuf(i - 1), 20, 4)) + " )", 8)
            
            sRst2 = ""
            
            sRstDt = Format(Now, "yyyyMMddHHmmss")
            
            If sIFCd <> "" And sRst <> "" Then
                With m_Result_Info
                    .RSTCNT = .RSTCNT + 1
                    .IFCD = .IFCD + sIFCd + Chr(124)
                    .RST1 = .RST1 + sRst + Chr(124)
                    .RST2 = .RST2 + sRst2 + Chr(124)
                    .RSTDT = .RSTDT + sRstDt + Chr(124)
                End With
            End If
        End If
        
continue:
    Next
            
    Exit Sub
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub sbOrder_Start()
    'ENQ 전송
    msSndState = "E"

    Comm.Output = Chr(5)

    If mp_sTestMode = "77" Then RaiseEvent PrintSendLog(Chr(5))
End Sub

Private Sub Init_m_Result_Info()
    With m_Result_Info
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
        .RSTDT = ""
        .UNIT = ""
        .FLAG = ""
        .ALARMCD = ""
        .INSTID = ""
    End With
End Sub

Private Sub Init_m_Sample_Info()
    With m_Sample_Info
        .ID = mc_sID
        Erase .IFCD
        ReDim .IFCD(mc_iOrdCnt)
        .ORDCNT = mc_iOrdCnt
        .PATINFO = mc_sPatInfo
        .POS = mc_sPos
        .RACK = mc_sRack
        .SAMPINFO = mc_sSampInfo
        .SAMPTYPE = mc_sSampleType
        .SEQNO = mc_sSeq
        .SINDEX = mc_bSerumIndex
    End With
End Sub

Private Sub Get_OrderString()
    Dim i           As Integer
    Dim tmpData()   As String
    Dim iCnt        As Integer

    If mp_sID = "" Or mp_iOrdCnt = 0 Then
        'ENQ 전송 후 EOT 전송하도록 함
        msSndH = Chr(4)

        Init_m_Sample_Info

        Exit Sub
    End If

    ReDim tmpData(mp_iOrdCnt) As String
    tmpData() = Split(mp_sTIFCd, Chr(124))

    With m_Sample_Info
        .ID = mp_sID
        .SEQNO = mp_sSeq
        .RACK = mp_sRack
        .POS = mp_sPos
        .SINDEX = mp_bSerumIndex
        .ORDCNT = mp_iOrdCnt
        .PATINFO = mp_sPatInfo
        .SAMPINFO = mp_sSampInfo
        Select Case mp_sSampleType
            Case "1", "2", "3", "4", "5"
                .SAMPTYPE = mp_sSampleType
            Case Else
                .SAMPTYPE = "1"
        End Select

        ReDim .IFCD(.ORDCNT)

        iCnt = 0

        For i = 1 To .ORDCNT
            If Trim(tmpData(i - 1)) <> "" Then
                iCnt = iCnt + 1
                .IFCD(iCnt) = tmpData(i - 1)
            End If
        Next

        .ORDCNT = iCnt      '실제 검사 가능한 항목 갯수
    End With

    'Packet 만들기
    msSndH = ""
    msSndP = ""
    msSndO = ""
    msSndL = ""
End Sub

Private Sub cmdTest_Click()
    msWkBuf = Text1
    Call PhaseCfg_Protocol
End Sub

Private Sub Comm_OnComm()
    Select Case Comm.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            msWkBuf = Comm.Input
            
            If mp_sTestMode = "77" Then
                RaiseEvent PrintRcvLog(msWkBuf)
            End If
                                
            If miSpaceCnt = 30 Then
                miSpaceCnt = 0
            End If
            miSpaceCnt = miSpaceCnt + 2
            
            RaiseEvent DispMsg(Space(miSpaceCnt) & "장비와 Interface 작업 중...")
            
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

Private Sub Text1_DblClick()

    Comm.Output = Text1
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Comm.CommPort = PropBag.ReadProperty("CommPort", 1)
    Comm.RTSEnable = PropBag.ReadProperty("RTSEnable", False)
    Comm.RThreshold = PropBag.ReadProperty("RThreshold", 0)
    Comm.Settings = PropBag.ReadProperty("Settings", "9600,n,8,1")
    mp_bPortOpen = PropBag.ReadProperty("PortOpen", mc_bPortOpen)
    mp_sOpenPW = PropBag.ReadProperty("OpenPW", mc_sOpenPW)
    mp_sEditPW = PropBag.ReadProperty("EditPW", mc_sEditPW)
    mp_sEqName = PropBag.ReadProperty("EqName", mc_sEqName)
    mp_bUseBarcode = PropBag.ReadProperty("UseBarcode", mc_bUseBarcode)
    mp_iPhase = PropBag.ReadProperty("Phase", mc_iPhase)
    mp_iSendPhase = PropBag.ReadProperty("SendPhase", mc_iSendPhase)
    mp_sTestMode = PropBag.ReadProperty("TestMode", mc_sTestMode)
    mp_iFrameN = PropBag.ReadProperty("FrameNo", mc_iFrameN)
    mp_sID = PropBag.ReadProperty("ID", mc_sID)
    mp_sSeq = PropBag.ReadProperty("SeqNo", mc_sSeq)
    mp_sRack = PropBag.ReadProperty("Rack", mc_sRack)
    mp_sPos = PropBag.ReadProperty("Pos", mc_sPos)
    mp_iOrdCnt = PropBag.ReadProperty("OrderCnt", mc_iOrdCnt)
    mp_sTIFCd = PropBag.ReadProperty("TotIFCd", mc_sTIFCd)
    mp_bSerumIndex = PropBag.ReadProperty("SerumIndex", mc_bSerumIndex)
    mp_sRerunGbn = PropBag.ReadProperty("RerunGbn", mc_sRerunGbn)
    mp_sSiteNm = PropBag.ReadProperty("SiteNm", mc_sSiteNm)
    mp_sPatInfo = PropBag.ReadProperty("PatientInfo", mc_sPatInfo)
    mp_sSampInfo = PropBag.ReadProperty("SampleInfo", mc_sSampInfo)
    mp_sSampleType = PropBag.ReadProperty("SampleType", mc_sSampleType)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CommPort", Comm.CommPort, 1)
    Call PropBag.WriteProperty("RTSEnable", Comm.RTSEnable, False)
    Call PropBag.WriteProperty("RThreshold", Comm.RThreshold, 0)
    Call PropBag.WriteProperty("Settings", Comm.Settings, "9600,n,8,1")
    Call PropBag.WriteProperty("PortOpen", mp_bPortOpen, mc_bPortOpen)
    Call PropBag.WriteProperty("OpenPW", mp_sOpenPW, mc_sOpenPW)
    Call PropBag.WriteProperty("EditPW", mp_sEditPW, mc_sEditPW)
    Call PropBag.WriteProperty("EqName", mp_sEqName, mc_sEqName)
    Call PropBag.WriteProperty("UseBarcode", mp_bUseBarcode, mc_bUseBarcode)
    Call PropBag.WriteProperty("Phase", mp_iPhase, mc_iPhase)
    Call PropBag.WriteProperty("SendPhase", mp_iSendPhase, mc_iSendPhase)
    Call PropBag.WriteProperty("TestMode", mp_sTestMode, mc_sTestMode)
    Call PropBag.WriteProperty("FrameNo", mp_iFrameN, mc_iFrameN)
    Call PropBag.WriteProperty("ID", mp_sID, mc_sID)
    Call PropBag.WriteProperty("SeqNo", mp_sSeq, mc_sSeq)
    Call PropBag.WriteProperty("Rack", mp_sRack, mc_sRack)
    Call PropBag.WriteProperty("Pos", mp_sPos, mc_sPos)
    Call PropBag.WriteProperty("OrderCnt", mp_iOrdCnt, mc_iOrdCnt)
    Call PropBag.WriteProperty("TotIFCd", mp_sTIFCd, mc_sTIFCd)
    Call PropBag.WriteProperty("SerumIndex", mp_bSerumIndex, mc_bSerumIndex)
    Call PropBag.WriteProperty("RerunGbn", mp_sRerunGbn, mc_sRerunGbn)
    Call PropBag.WriteProperty("SiteNm", mp_sSiteNm, mc_sSiteNm)
    Call PropBag.WriteProperty("PatientInfo", mp_sPatInfo, mc_sPatInfo)
    Call PropBag.WriteProperty("SampleInfo", mp_sSampInfo, mc_sSampInfo)
    Call PropBag.WriteProperty("SampleType", mp_sSampleType, mc_sSampleType)
End Sub

'사용자 정의 컨트롤에 대한 속성을 초기화합니다.
Private Sub UserControl_InitProperties()
    mp_bPortOpen = mc_bPortOpen
    mp_sOpenPW = mc_sOpenPW
    mp_sEditPW = mc_sEditPW
    mp_sEqName = mc_sEqName
    mp_bUseBarcode = mc_bUseBarcode
    mp_iPhase = mc_iPhase
    mp_iSendPhase = mc_iSendPhase
    mp_sTestMode = mc_sTestMode
    mp_iFrameN = mc_iFrameN
    mp_sID = mc_sID
    mp_sSeq = mc_sSeq
    mp_sRack = mc_sRack
    mp_sPos = mc_sPos
    mp_iOrdCnt = mc_iOrdCnt
    mp_sTIFCd = mc_sTIFCd
    mp_bSerumIndex = mc_bSerumIndex
    mp_sRerunGbn = mc_sRerunGbn
    mp_sSiteNm = mc_sSiteNm
    mp_sPatInfo = mc_sPatInfo
    mp_sSampInfo = mc_sSampInfo
    mp_sSampleType = mc_sSampleType
End Sub

Public Sub Send_Chr(ByVal aiChr As Integer)
    On Error GoTo ErrComm
    
    Comm.Output = Chr(aiChr)
    
    Exit Sub
    
ErrComm:
    RaiseEvent DispMsg("Send_Chr 에러 - " & Err.Description)
End Sub
