VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl Qcm 
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
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
Attribute VB_Name = "Qcm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const HKEY_CURRENT_USER = &H80000001

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
Private m_p_sPatInfo As Variant
Private m_EqName As String
Private m_bUseBarcode As Boolean
Private m_iPhase As Integer
Private m_iSendPhase As Integer
Private m_sTestMode As String
Private m_iFrameN As Integer
Private m_p_sID As String
Private m_p_sSeq As String
Private m_p_sRack As String
Private m_p_sPos As String
Private m_p_iOrdCnt As Integer
Private m_p_sTIFCd As String
Private m_PortOpen As Boolean
Private m_OpenPW As String
Private m_EditPW As String
Private m_SendData As String

'이벤트 선언:
Event AppendData(sID As String, sSeq As String, sRack As String, sPos As String, iRstCnt As Integer, sTIFCd As String, sTRst1 As String, sTRst2 As String, sTRdt As String, sURid As String, sUnit As String, sTFlag As String, sQCGbn As String)
Event SendOrderOK(sID$, sSeqNo$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event RequestCurOrder(sSeqNo$)
Event DispMsg(sMsg$)

'===== User Define
'인터페이스에서 사용
Private f_strRcvBuffer  As String
Private f_strWkBuf      As String
Private f_strState      As String
Private f_blnSend       As Boolean
Private f_bEndChk       As Boolean
Private f_bSTXChk       As Boolean

'구조체 지정
Private f_typSampleInfo As SAMPLE_INFO
Private f_typResultInfo As RESULT_INFO

Private f_intSpaceCnt   As Integer

Private Sub SendOrder_BD()

    On Error GoTo ErrRtn
    
    Dim sTmp    As String
    Dim ChkS    As String
    Dim strDta1()   As String
    
    Dim i       As Integer
    
    If m_iFrameN >= 7 Then
        m_iFrameN = 1
    End If
    
    Do While True
        Select Case m_iSendPhase
            Case 1      'Header Record
                sTmp = m_iFrameN & "H|\^&|||Becton Dickinson||||||||V1.0|" & Format$(Now, "YYYYMMDD") & vbCr
                '----- 검사항목 조회/편집
                RaiseEvent RequestCurOrder(f_typSampleInfo.SEQNO)
    
                Call Get_OrderString
    
                '경우 오더가 없는 경우
'                If f_typSampleInfo.ORDCNT > 0 Then
                    m_iSendPhase = 2
'                Else
'                    m_iSendPhase = 4
'                End If
                
            Case 2      'Patient Record
                '-- 차트번호~성별~과~병동~검사코드
                If InStr(f_typSampleInfo.OTHER, "~") > 0 Then
                    strDta1 = Split(f_typSampleInfo.OTHER, "~")
                Else
                    ReDim strDta1(0 To 5) As String
                    strDta1(0) = f_typSampleInfo.OTHER  '-- 등록번호
                    strDta1(1) = "" '-- 성별
                    strDta1(2) = "" '-- 과
                    strDta1(3) = "" '-- 병동
                    strDta1(4) = "" '-- 검사코드
                End If
                
                sTmp = sTmp & "P|1||" & strDta1(0) & "||^^^^|||" & strDta1(1) & "||" & f_typSampleInfo.ID & "^^^^||" & strDta1(4) & "||" & f_typSampleInfo.ID & "|^^^^|||||^^^^||||||" & strDta1(3) & "|" & strDta1(2) & "|||||||" & vbCr
                m_iSendPhase = 3
                
            Case 3      'Order Record
                
                'BarCode 사용모드
                sTmp = sTmp & "O|1|" & f_typSampleInfo.SEQNO & "^^^||^^^^||||||^^||^^||^||^^|^^^^|||||^|||||" & vbCr
                m_iSendPhase = 5
                
            Case 4
'                msComm.Output = Chr(4) & Chr(5)
                sTmp = sTmp & "Q|1|||||||||||A" & vbCr
                
                m_iSendPhase = 5
                
            Case 5      'Terminator Record
                sTmp = sTmp & "L|1|N" & vbCr & Chr(3)
                
                msComm.Output = Chr(5) & Chr(2) & sTmp & ChkSum_ASTM(sTmp) & vbCrLf & Chr(4)
                m_iSendPhase = 6
                
                If sTestMode = "77" Then
                    RaiseEvent PrintSendLog(Chr(5) & Chr(2) & sTmp & ChkSum_ASTM(sTmp) & vbCrLf & Chr(4))
                End If
                
            Case 6      'EOT
                m_iFrameN = 1:  m_iPhase = 1:   m_iSendPhase = 7
                f_strState = ""
                
                'Barcode Mode인 경우 전송완료 이벤트 발생
                RaiseEvent SendOrderOK(f_typSampleInfo.ID, f_typSampleInfo.SEQNO)
                
                Exit Sub
        End Select
        
        If m_iSendPhase = 7 Then Exit Do
    Loop
    
    m_iFrameN = m_iFrameN + 1

    If sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sTmp & ChkSum_ASTM(sTmp) & vbCrLf & Chr(4))
    End If
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러 - " & Err.Description)
    End If
End Sub

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
        Case "ACCUCHEK"
            Call PhaseCfg_Protocol_AccuChek
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_AccuChek()
            
    Dim wkdat   As String
    Dim ix1     As Long
    
    For ix1 = 1 To Len(f_strWkBuf)
        wkdat = Mid$(f_strWkBuf, ix1, 1)
                 
        Select Case m_iPhase
            Case 1            'ENQ 대기
                Select Case Asc(wkdat)
                    Case 5
                        f_bEndChk = True: f_bSTXChk = False
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                    Case Else
                        m_iPhase = 1
                End Select
            
            Case 2      '<LF> 대기
                Select Case Asc(wkdat)
                    Case 2  '-STX
                        If f_bEndChk = True Then
'                            f_strRcvBuffer = ""
                        Else
                            f_bSTXChk = True
                        End If
                        f_bEndChk = True
                        
                    Case 3  'ETX
                        msComm.Output = Chr(6)
                    
                    Case 4  'EOT
                        If f_bEndChk = True Then
                            Call DataEditResponse_AccuChek
                            f_strRcvBuffer = ""
                        End If
                    
                    Case 5  '-ENQ
                        f_bEndChk = True: f_bSTXChk = False
                        msComm.Output = Chr(6)
                        
                    Case 21 '-NAK
                        
                    Case 23 '-ETB
                        f_bEndChk = False
                        msComm.Output = Chr(6)
                        
                    Case Else
                        If f_bEndChk = True Then
                            If f_bSTXChk = True Then
                                f_bSTXChk = False
                            Else
                                f_strRcvBuffer = f_strRcvBuffer & wkdat
                            End If
                        End If
                        
                End Select
            
            Case 3      'ACK 대기
                Select Case Asc(wkdat)
                    Case 6      'ACK
                        If f_strState = "Q" Then
                            Call SendOrder_BD
                        End If
                    
                    Case 5      'ENQ
                        f_bEndChk = True: f_bSTXChk = False
                        msComm.Output = Chr(6)
                        m_iPhase = 2
                        
                    Case 21     'NAK
                        msComm.Output = Chr(5)
                        m_iPhase = 3
                        
                    Case 3      'EOT
                        m_iPhase = 1
                End Select
                
        End Select
    Next ix1

End Sub


' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_AccuChek()

    On Error GoTo ErrHandler
    
    Dim strRecord() As String
    Dim intIdx  As Integer
    
    Dim strField As String
    
    Dim strData1$(), strData2$()

    If f_strRcvBuffer = "" Then Exit Sub
     
    strRecord = Split(f_strRcvBuffer, vbCrLf)
    
    For intIdx = 0 To UBound(strRecord) - 1
        
        If Mid(strRecord(intIdx), 2, 1) = "H" Or Mid(strRecord(intIdx), 3, 1) = "H" Then
            Call Init_f_typResultInfo
            
        Else
            Select Case Mid(strRecord(intIdx), 2, 1)
                Case "H"        'Header Record
                    Call Init_f_typResultInfo
                Case "M"
                Case "P"        'Patient Record
                Case "O"
                    
                Case "R"
                    strData1() = Split(strRecord(intIdx), "|")
                    strData2() = Split(strData1(2), "^")
                    
                    strField = strData2(3)
                    Select Case strField
                        Case "PATID"    '-- 등록번호
                                        f_typResultInfo.ID = strData1(3)
                                        
                        Case "ANALYZERNAME" '-- 장비코드
                                        f_typResultInfo.RACK = strData1(3)
                                            
                        Case "GLUCOSE"  '-- 검사결과
                        
                                        '결과정보 구조체에 저장
                                        With f_typResultInfo
                                            '결과값 누적
                                            .RSTCNT = .RSTCNT + 1
                                            .RST1 = .RST1 & strData1(3) & Chr(124)
                                            .RST2 = .RST2 & "" & Chr(124)
                                            .UNIT = .UNIT & "" & Chr(124)
                                            .FLAG = .FLAG & "" & Chr(124)
                                        End With
                                            
                        Case "TESTCODE" '-- AST 결과
                                        With f_typResultInfo
                                            '결과값 누적
                                            .IFCD = .IFCD & strData1(3) & Chr(124)
                                        End With
                        Case "OPID"     '-- 테스트 ID
                                        f_typResultInfo.INSTID = strData1(3)
                        
                        Case "ANALYZEDATETIME"
                                        '-- 검사일시
                                        f_typResultInfo.RSTDT = f_typResultInfo.RSTDT & strData1(3) & "|"
                        Case Else
                                        If InStr(strField, "COMMENT") > 0 Then
                                            If strData1(3) <> "" Then f_typResultInfo.ID = "-" & f_typResultInfo.ID
                                        End If
                    End Select
                        
                Case "L"
                        '결과값 등록/화면 표시 처리...
                        With f_typResultInfo
                            If .RSTCNT > 0 Then
                                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .RSTDT, .INSTID, .UNIT, .FLAG, .QCGBN)
                            End If
                        End With
                    
            End Select
        End If
        
    Next
    
    Exit Sub
    
ErrHandler:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit Error - " & Err.Description)
    End If
End Sub

Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With f_typSampleInfo
            .ID = m_p_sID
            .ORDCNT = 0
            Erase .IFCD
        End With
        
        Exit Sub
    End If
    
    With f_typSampleInfo
        .ID = m_p_sID
        .SEQNO = m_p_sSeq
        .RACK = m_p_sRack
        .POS = m_p_sPos
        .ORDCNT = 1      '실제 검사 가능한 항목 갯수
        ReDim Preserve .IFCD(1 To 1) As String
        .IFCD(1) = ""
        .OTHER = m_p_sPatInfo
    End With
        
End Sub

'
'   결과정보 구조체 초기화
'
Private Sub Init_f_typResultInfo()
    
    With f_typResultInfo
        .ID = ""
        .SEQNO = ""
        .RACK = ""
        .POS = ""
        .QCGBN = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .RSTDT = ""
        .INSTID = ""
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

    f_strWkBuf = Text1
    Call PhaseCfg_Protocol

End Sub

Private Sub msComm_OnComm()
        
    Select Case msComm.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            f_strWkBuf = msComm.Input
            
            If sTestMode = "77" Then
                RaiseEvent PrintRcvLog(f_strWkBuf)
            End If
                                
            If f_intSpaceCnt = 30 Then
                f_intSpaceCnt = 0
            End If
            f_intSpaceCnt = f_intSpaceCnt + 2
            
            RaiseEvent DispMsg(Space(f_intSpaceCnt) & "장비와 Interface 작업 중...")
            
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

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=0,0,0,0
Public Property Get SendData() As String
    SendData = m_SendData
End Property

Public Property Let SendData(ByVal QueryInfo As String)
    On Error GoTo ErrSendData
    
    Dim ADORS   As New ADODB.Recordset
    Dim sRetVal$: sRetVal = ""
    Dim SqlStr$: SqlStr = ""
    Dim stmpDT$
    Dim sData()     As String
    Dim aTmpRow()   As String
    Dim aTmpData()  As String
    Dim ii%
    
    m_SendData = QueryInfo
    PropertyChanged "SendData"
    
    If f_intSpaceCnt = 30 Then
        f_intSpaceCnt = 0
    End If
    f_intSpaceCnt = f_intSpaceCnt + 2
    
    RaiseEvent DispMsg(Space(f_intSpaceCnt) & "장비와 Interface 작업 중...")

    If m_SendData <> "" Then
        sData = Split(m_SendData, "|")
        
        If Trim(sData(1)) <> "" Then
            SqlStr = " SELECT distinct DPatId, dtTestDateTime, vcTest, vcResult, vcOperId, vcInstName, convert(bigint, attsTimeStamp)" _
                    & "  FROM rpt_vwPatRstList " _
                    & " WHERE attsTimeStamp > " & Trim(sData(1)) & " " _
                    & "   AND DPatId is not null " _
                    & "   AND fkchResultType = 'PAT'" _
                    & "   AND dtTestDateTime >= GETDATE() - 3" _
                    & " order by 7 "
        ElseIf Trim(sData(2)) <> "" Then
            stmpDT = Format(Trim(sData(2)), "YYYY-MM-DD HH:NN:SS")
            SqlStr = " SELECT distinct DPatId, dtTestDateTime, vcTest, vcResult, vcOperId, vcInstName, convert(bigint, attsTimeStamp)" _
                    & "  FROM rpt_vwPatRstList " _
                    & " WHERE dtTestDateTime > '" & Trim(stmpDT) & "'" _
                    & "   AND DPatId is not null " _
                    & "   AND fkchResultType = 'PAT'" _
                    & " order by 7 "
        End If
        
        ADORS.Open SqlStr, "Driver={SQL Server};Server=" & Trim(sData(0)) & ";Database=QCM3;Uid=QCM;Pwd=QCM", adOpenForwardOnly
        
        If ADORS.EOF <> True Then
            sRetVal = ADORS.GetString(adClipString, -1, Chr(124), Chr(3))
        End If
        ADORS.Close: Set ADORS = Nothing
        
        If sRetVal = "" Then
            GoTo ErrSendData
        End If
        
        aTmpRow() = Split(sRetVal, Chr(3))
        For ii = 0 To UBound(aTmpRow())
            If Trim(aTmpRow(ii)) = "" Then
                Exit For
            End If
        
            If f_intSpaceCnt = 30 Then
                f_intSpaceCnt = 0
            End If
            f_intSpaceCnt = f_intSpaceCnt + 2
            RaiseEvent DispMsg(Space(f_intSpaceCnt) & "장비와 Interface 작업 중...")
            
            Erase aTmpData()
            aTmpData() = Split(aTmpRow(ii), Chr(124))
            
            If UBound(aTmpData()) < 5 Then
                Exit For
            End If
            
            With f_typResultInfo
                .ID = Trim(aTmpData(0))     '환자번호,바코드번호
                .SEQNO = Trim(aTmpData(6))   '
                .RSTDT = Trim(aTmpData(1))  '검사일시
                .INSTID = Trim(aTmpData(4)) '사용자ID
                .RACK = Trim(aTmpData(5))   '장비코드
                '결과값 초기화
                .RSTCNT = 0
                .IFCD = ""
                .RST1 = ""
                .RST2 = ""
                .UNIT = ""
                .FLAG = ""
                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & Trim(aTmpData(2)) & Chr(124)
                .RST1 = .RST1 & Trim(aTmpData(3)) & Chr(124)
                .RST2 = .RST2 & "" & Chr(124)
                .UNIT = .UNIT & "" & Chr(124)
                .FLAG = .FLAG & "" & Chr(124)
                
                Call Sleep(3000)
                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .RSTDT, .INSTID, .UNIT, .FLAG, .QCGBN)
                Exit For
            End With
            
        Next ii
        RaiseEvent AppendData("EOF", "", "", "", 0, "", "", "", "", "", "", "", "")
    End If
    On Error GoTo 0

ErrSendData:
    RaiseEvent AppendData("EOF", "", "", "", 0, "", "", "", "", "", "", "", "")
    If Err <> 0 Then
        If ADORS.State = 1 Then
            ADORS.Close: Set ADORS = Nothing
        End If
        RaiseEvent DispMsg(Err.Description)
        RaiseEvent RaiseError("SendData Error!!! " & Err.Description)
    End If
End Property
