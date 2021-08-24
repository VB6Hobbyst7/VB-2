VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl ADVIA120 
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   3330
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1710
      Top             =   120
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
   Begin MSCommLib.MSComm Comm 
      Left            =   255
      Top             =   2370
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
End
Attribute VB_Name = "ADVIA120"
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
Private Const mc_bReserveEnd    As Boolean = False

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
Dim mp_bReserveEnd          As Boolean

'이벤트 선언:
Public Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)
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

'< ADVIA120
Private Const mc_iMaxCnt  As Integer = 100

Private msMT              As String
Private msTimerFlag       As String
Private msSndPacket       As String
'> ADVIA120

' ADVIA60
Private msPreSeq   As String

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

Public Property Get ReserveEnd() As Boolean
    ReserveEnd = mp_bReserveEnd
End Property

Public Property Let ReserveEnd(ByVal New_p_bReserveEnd As Boolean)
    mp_bReserveEnd = New_p_bReserveEnd
    PropertyChanged "ReserveEnd"
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

Public Property Get SerumIndex() As Boolean
    SerumIndex = mp_bSerumIndex
End Property

Public Property Let SerumIndex(ByVal New_SerumIndex As Boolean)
    mp_bSerumIndex = New_SerumIndex
    PropertyChanged "SerumIndex"
End Property

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
    
    If mp_sEqName = "" Then
        RaiseEvent DispMsg("검사장비명을 지정해 주십시오.!!!")
        Exit Sub
    End If
    
    Select Case UCase(mp_sEqName)
        Case "ADVIA120"
            Call PhaseCfg_Protocol_EQ
            
        Case "ADVIA60"
            Call PhaseCfg_Protocol_ADVIA60
        
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
End Sub

Private Sub PhaseCfg_Protocol_EQ()
    Dim sWkDat   As String
    Dim i   As Integer
    
    For i = 1 To Len(msWkBuf)
        sWkDat = Mid(msWkBuf, i, 1)
        
        Select Case mp_iPhase
            Case 1            ' 초기화 확인(MT 대기)
                Select Case sWkDat
                    Case Chr(2)     'STX
                        msRcvBuffer = ""
                        
                    Case Chr(21)    'NACK
                        Call InitialComm
                        mp_iPhase = 1
                        
                        Exit Sub
                        
                    Case msMT                    '즉 : '0'(&H30), Initialize Message 직후이므로
                        msRcvBuffer = ""
                        
                        'Transfer_Token 시도 후
                        Call TransferToken
                        
                        'Phase 옮김
                        mp_iPhase = 2
                        
                    Case Else
                        'MT이외의 경우(즉, NACK)
                        Call InitialComm
                        mp_iPhase = 1
                        
                        Exit Sub
                End Select
                
            Case 2            'Token Tranfer에 대한 MT 대기
                Select Case sWkDat
                    Case Chr(21)    'NACK
                        Call InitialComm
                        mp_iPhase = 1
                        
                        Exit Sub
                        
                    Case msMT
                        mp_iPhase = 3
                        
                    Case Chr(Asc(msMT) - 1)
                        msMT = Chr((msMT) - 1)

                        If Asc(msMT) < 0 Then
                            msMT = Chr(&H30)
                        End If

                        Call TransferToken      'Transfer_Token 시도 후
                        mp_iPhase = 2              'Phase 옮김
                        
                    Case Else
                        Call TransferToken
                        mp_iPhase = 2
                        
                End Select
                    
            Case 3            'CheckSum STX인 경우의 오류 방지를 위해, Phase 3 과 4 분리
                Select Case Asc(sWkDat)
                    Case 2
                        msRcvBuffer = ""
                        mp_iPhase = 4
                    Case Else
                        mp_iPhase = 3
                End Select
            
            Case 4            'DataEdit(장비쪽에서 보내는 S, Q, R 메세지에 대한) 대기,
                Select Case Asc(sWkDat)
                    Case 3            ' ETX
                        msMT = Left(msRcvBuffer, 1)
                                                                        
                        'Delay 0.025 sec --> LIS.ini의 Host_Tmt에 정의
                        Sleep 25
                        
                        Comm.Output = msMT
                                                                        
                        If mp_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(msMT)
                        End If
                        
                        Call DataEditResponse_ADVIA120
                        
                        mp_iPhase = 3
                        
                    Case Else
                        msRcvBuffer = msRcvBuffer & sWkDat
                        
                End Select
        
        End Select
    Next
End Sub

Private Sub PhaseCfg_Protocol_ADVIA60()
    Dim sWkDat   As String
    Dim i   As Integer
    
    For i = 1 To Len(msWkBuf)
        sWkDat = Mid(msWkBuf, i, 1)
        
        Select Case Asc(sWkDat)
            Case 2  'STX
                msRcvBuffer = ""
                
            Case 3  'ETX
                Call DataEditResponse_ADVIA60
                msRcvBuffer = ""
                
            Case Else
                msRcvBuffer = msRcvBuffer & sWkDat
                
        End Select
    Next
End Sub


Private Sub InitialComm()
    Dim sSendBuf$
    
    mp_iPhase = 1
    
    msMT = Chr(&H30)
    sSendBuf = msMT & "I " & vbCr & vbLf
    sSendBuf = CheckSum(sSendBuf)
    
    Comm.Output = sSendBuf
       
    If mp_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSendBuf)
    End If
End Sub

Private Function CheckSum(ByVal sSendBuf As String) As String
    Dim i%
    Dim sXOR$
    
    sXOR = ""
    
    sXOR = Mid(sSendBuf, 1, 1)
    
    For i = 2 To Len(sSendBuf)
'        sXOR = Chr(CInt(Asc(sXOR)) Xor CInt(Asc(Mid(sSendBuf, i, 1))))
        sXOR = Chr(Int(Asc(sXOR)) Xor Int(Asc(Mid(sSendBuf, i, 1))))
    Next
    
    If sXOR = Chr(3) Then
        sXOR = Chr(127)
    End If
    
    CheckSum = Chr(2) & sSendBuf & sXOR & Chr(3)
End Function

Private Sub TransferToken()
    Dim sSendBuf$
    
    '프로그램 종료 예정 --> 장비쪽이 Slave인 상태에서 종료되도록 TransferToken 안함
    If mp_bReserveEnd Then
        Comm.PortOpen = False
        mp_bPortOpen = False
        PropertyChanged "PortOpen"
        
        Exit Sub
    End If
    
    If msMT = "" Then msMT = Chr(&H30)
    
    msMT = Chr(Asc(msMT) + 1)
   
    If msMT > "Z" Then
        msMT = "0"
    End If
    
    sSendBuf = msMT & "S " & vbCr & vbLf
    
    sSendBuf = CheckSum(sSendBuf)
    
    '< rem freety 2005/04/21
    '# Sleep으로 처리하니 프로그램이 멎는 듯 보이는 현상때문에 Timer로 수정
'    'Delay 5 sec --> LIS.ini의 Host_Tsw에 정의
'    Sleep 5000
'
'    Comm.Output = sSendBuf
'
'    If mp_sTestMode = "77" Then
'        RaiseEvent PrintSendLog(sSendBuf)
'    End If
'    '> rem freety 2005/04/21
    
    '< add freety 2005/04/21
    '# Sleep으로 처리하니 프로그램이 멎는 듯 보이는 현상때문에 Timer로 수정
    msSndPacket = sSendBuf

''    Timer1.Interval = 5000
''    Timer1.Enabled = True
    '> add freety 2005/04/21
End Sub

Private Sub TransferOrderMsg(ByVal sTestBuf As String)
    Dim sSendBuf$
    
    If msMT = "" Then msMT = Chr(&H30)
    
    msMT = Chr(Asc(msMT) + 1)
    
    If msMT > "Z" Then
        msMT = "0"
    End If
    
    If sTestBuf = "" Then
    '해당 Work Order가 없을 경우
        sSendBuf = msMT & "N R " & m_Sample_Info.ID & vbCr & vbLf
    Else
    '해당 Work Order가 있는 경우
        sSendBuf = msMT & "Y     " & m_Sample_Info.ID & Space(42)
        sSendBuf = sSendBuf & Space(58)
        sSendBuf = sSendBuf & Space(14) & vbCr & vbLf
        sSendBuf = sSendBuf & sTestBuf
        sSendBuf = sSendBuf & vbCr & vbLf
    End If
        
    sSendBuf = CheckSum(sSendBuf)
    
    'Delay 2 sec --> 오더쿼리시간 고려하여 Delay 1 sec
    Sleep 1000
    
    Comm.Output = sSendBuf
   
    If mp_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSendBuf)
    End If
End Sub

Private Sub TransferResultValMsg()
    'Result Validation Message
    
    Dim sSendBuf$
    
    If msMT = "" Then msMT = Chr(&H30)

    msMT = Chr(Asc(msMT) + 1)
        
    If msMT > "Z" Then
        msMT = "0"
    End If
    
    sSendBuf = msMT & "Z                  0" & vbCr & vbLf
    sSendBuf = CheckSum(sSendBuf)
    
    'Delay 1.5 sec --> 결과등록시간 고려하여 Delay 1 sec
    Sleep 1000
    
    Comm.Output = sSendBuf
   
    If mp_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSendBuf)
    End If
End Sub

Private Sub DataEditResponse_ADVIA120()
    On Error GoTo ErrRtn
    
    Dim i%, iRstCnt%
    Dim sRxData$, sBarCd$, sSeqNo$, sRack$, sPos$
    Dim sBC$, sLC$
    
    Dim sRstCd$, sRst$, sRst2$
    Dim sTotIFCd$, sTotRst$, sTotRst2$
    Dim sTmp$
    
    Dim iStartPos As Integer
    
    sRxData = msRcvBuffer
    
    If Trim(sRxData) = "" Then Exit Sub

    ''sBC = Mid(sRxData, 2, 1)
    sBC = Mid(sRxData, 1, 1)

    Select Case sBC
        Case "S"
            Call TransferToken
            
        Case "Q"
            sBarCd = Mid(sRxData, 4, 14)
            
            With m_Sample_Info
                .ID = sBarCd
                .RACK = ""
                .POS = ""
                
                RaiseEvent RequestCurOrder(.ID, .RACK, .POS)
                    
                Call Get_OrderString
                
                If .ORDCNT > 0 Then
                    RaiseEvent SendOrderOK(.ID, .SEQNO, .RACK, .POS)
                End If
            End With
            
        Case "R"
            sBarCd = Mid(sRxData, 4, 14)
            sRack = Mid(sRxData, 19, 3)
            sPos = Mid(sRxData, 23, 2)
            
            Call Init_m_Result_Info
            
            'iStartPos : Result 시작 위치
            iStartPos = InStr(sRxData, vbCr)
        
            If iStartPos = 0 Then
                RaiseEvent DispMsg("DataEdit Error - " & "Cannot find position of result!!")
                Exit Sub
            End If
            
            iStartPos = iStartPos + 2
            
            For i = 1 To mc_iMaxCnt
                sTmp = Mid(sRxData, iStartPos + 9 * (i - 1), 1)
            
                If sTmp = vbCr Then Exit For
                
                sRstCd = CStr(Val(Trim(Mid(sRxData, iStartPos + 9 * (i - 1), 3))))
                sRst = Trim(Mid(sRxData, iStartPos + 9 * (i - 1) + 3, 5))
                sRst2 = Trim(Mid(sRxData, iStartPos + 9 * (i - 1) + 3 + 5, 1))
                
                If Left(sRst, 1) = "." Then
                    sRst = "0" & sRst
                End If
                
                If sRstCd <> "" And sRst <> "" Then
                    iRstCnt = iRstCnt + 1
                    sTotIFCd = sTotIFCd & sRstCd & Chr(124)
                    sTotRst = sTotRst & sRst & Chr(124)
                    sTotRst2 = sTotRst2 & sRst2 & Chr(124)
                End If
            Next
            
            '결과정보 구조체에 저장
            With m_Result_Info
                .ID = sBarCd
                .RACK = sRack
                .POS = sPos
                .RSTCNT = iRstCnt
                .IFCD = sTotIFCd
                .RST1 = sTotRst
                .RST2 = sTotRst2
                .UNIT = String(iRstCnt, Chr(124))
                .FLAG = String(iRstCnt, Chr(124))
                
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With

            Call Init_m_Result_Info
            
            Call TransferResultValMsg
        
    End Select
    
    Exit Sub
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub DataEditResponse_ADVIA60()
    On Error GoTo ErrRtn
    
    Dim i%, iRstCnt%
    Dim sRxData$, sBarCd$, sSeqNo$, sRack$, sPos$, sFlag$
    Dim sBC$, sLC$
    
    Dim sRstCd$, sRst$, sRst2$
    Dim sTotIFCd$, sTotRst$, sTotRst2$, sTotFlag$
    Dim sTmp$
    Dim sField$()
        
    Dim iStartPos As Integer
    
    sTotIFCd = "WBC|LYM#|LYM%|MON#|MON%|GRA#|GRA%|RBC|HGB|HCT|MCV|MCH|MCHC|RDW|PLT|MPV|"
    
    sRxData = msRcvBuffer

    sBC = Mid(sRxData, 1, 1)
    
    sField = Split(sRxData, Chr(13))

    Select Case sBC
        Case "R"
            sSeqNo = Val(sField(1))
            
            If msPreSeq = sSeqNo Then
                Call Init_m_Result_Info
                Exit Sub
            End If
            
            msPreSeq = sSeqNo
            
            'WBC|LYM#|LYM%|MON#|MON%|GRA#|GRA%
            For i = 4 To 10
                iRstCnt = iRstCnt + 1
                sRst = Val(Trim(Mid(sField(i), 1, 6)))
                sFlag = Trim(Mid(sField(i), 7, 2))
                
                sTotRst = sTotRst & sRst & Chr(124)
                sTotRst2 = sTotRst2 & sRst2 & Chr(124)
                sTotFlag = sTotFlag & sFlag & Chr(124)
            Next
            
            'RBC|HGB|HCT|MCV|MCH|MCHC|RDW
            For i = 25 To 31
                iRstCnt = iRstCnt + 1
                sRst = Val(Trim(Mid(sField(i), 1, 6)))
                sFlag = Trim(Mid(sField(i), 7, 2))
                
                sTotRst = sTotRst & sRst & Chr(124)
                sTotRst2 = sTotRst2 & sRst2 & Chr(124)
                sTotFlag = sTotFlag & sFlag & Chr(124)
            Next
            
            'PLT|MPV
            For i = 33 To 34
                iRstCnt = iRstCnt + 1
                sRst = Val(Trim(Mid(sField(i), 1, 6)))
                sFlag = Trim(Mid(sField(i), 7, 2))
                
                sTotRst = sTotRst & sRst & Chr(124)
                sTotRst2 = sTotRst2 & sRst2 & Chr(124)
                sTotFlag = sTotFlag & sFlag & Chr(124)
            Next
                        
            '결과정보 구조체에 저장
            With m_Result_Info
                .ID = ""
                .SEQNO = sSeqNo
                .RACK = ""
                .POS = ""
                .RSTCNT = iRstCnt
                .IFCD = sTotIFCd
                .RST1 = sTotRst
                .RST2 = sTotRst2
                .UNIT = String(iRstCnt, Chr(124))
                .FLAG = sTotFlag
                
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With

            Call Init_m_Result_Info
    End Select
    
    Exit Sub
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
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
    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    '< rem freety 2005/04/19
    '# Order없는 경우에도 Skip 하도록
'    If mp_sID = "" Or mp_iOrdCnt = 0 Then
'        Init_m_Sample_Info
'
'        Exit Sub
'    End If
    '> rem freety 2005/04/19
    
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
        For ii = 1 To .ORDCNT
            If Trim(tmpData(ii - 1)) <> "" Then
                iCnt = iCnt + 1
                .IFCD(iCnt) = tmpData(ii - 1)
            End If
        Next ii
        .ORDCNT = iCnt      '실제 검사 가능한 항목 갯수
    End With
    
    Dim sTmp As String
    Dim sTestBuf As String
    Dim sSendBuf As String
    
    With m_Sample_Info
        For ii = 1 To .ORDCNT
            sTmp = .IFCD(ii)
            
            If Not sTmp = "" Then
                sTestBuf = sTestBuf & Right(Space(3) & CStr(Val(sTmp)), 3)
            End If
        Next
        
        Call TransferOrderMsg(sTestBuf)
    End With
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

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Timer1.Interval = 0
    
    Select Case msTimerFlag
        Case "I"
            If msMT = "" Then
                InitialComm
                
                msSndPacket = ""
                msTimerFlag = ""
            Else
                mp_bReserveEnd = True
                PropertyChanged "ReserveEnd"
                
                If mp_bPortOpen = False Then
                    mp_bReserveEnd = False
                    PropertyChanged "ReserveEnd"
                    
                    Comm.PortOpen = True
                    mp_bPortOpen = True
                    PropertyChanged "PortOpen"
                    
                    Sleep 1000
                    
                    InitialComm
                    
                    msSndPacket = ""
                    msTimerFlag = ""
                Else
                    'Timer 재가동
                    msTimerFlag = "I"
''                    Timer1.Interval = 1000
''                    Timer1.Enabled = True
                End If
            End If
            
        Case Else
            If msSndPacket = "" Then Exit Sub
    
            Comm.Output = msSndPacket
        
            If mp_sTestMode = "77" Then
                RaiseEvent PrintSendLog(msSndPacket)
            End If
            
            msSndPacket = ""
            msTimerFlag = ""
            
    End Select
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
    mp_bReserveEnd = PropBag.ReadProperty("ReserveEnd", mc_bReserveEnd)
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
    Call PropBag.WriteProperty("ReserveEnd", mp_bReserveEnd, mc_bReserveEnd)
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
    mp_bReserveEnd = mc_bReserveEnd
End Sub

Public Function Send_Initial()
    
    msTimerFlag = "I"
''    Timer1.Interval = 1000
''    Timer1.Enabled = True

'    '2006/8/23 yk
'    Timer1.Enabled = False
'    Timer1.Interval = 0
'
'    InitialComm
End Function

Public Function Send_Chr(ByVal aiChr As Integer) As Variant
    On Error GoTo ErrComm
    
    Comm.Output = Chr(aiChr)
    
    Exit Function
    
ErrComm:
    RaiseEvent DispMsg("Send_Chr 에러 - " & Err.Description)
End Function
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function Send_Initial2() As Variant

    Timer1.Enabled = False
    Timer1.Interval = 0

    InitialComm

End Function

