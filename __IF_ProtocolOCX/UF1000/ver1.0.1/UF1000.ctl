VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl UF1000 
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
      Handshaking     =   1
   End
End
Attribute VB_Name = "UF1000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
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
Dim m_iBCLen As Integer
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
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sKind$, sReviewFlag$, sErrFlag$, sOther1$)
Event RequestCurOrder(sID$, sSeq$, sRack$, sPos$)
Event RaiseError(sError$)
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event SendOrderOK(sID$, sSeq$, sRack$, sPos$)
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

'for UF-1000i
Dim msBarCd$, msRack$, msPos$
Dim msReviewFlag$, msEqErrFlag$
Dim msReview$, msRBCInfo$
Dim miRstCnt%
Dim msTotIFCd$, msTotRst$, msTotRst2$, msUnit$, msFlag$

Dim msSendBuf$, msSendBuf2$
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
        Case "UF1000"
            Call PhaseCfg_Protocol_UF1000
        
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub DataEditResponse_UF1000()
    On Error GoTo ErrRtn

    Dim sBC         As String
    Dim sLC         As String

    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    Dim ii          As Integer
    Dim sData()     As String
    Dim tmpIFCd$, tmpRst$, tmpRstDT$
    Dim sTIFCd$, sTRst$, sTRst2$, sTUnit$, sTFlag$
    Dim tmpCnt%
    Dim sUFNeed     As String
    Dim sMach       As String
    Dim sMode       As String
    
    Dim iChk%, iCnt%, sIFRstCd$, sRst$
    
    sBC = Mid(RcvBuffer, 1, 1)
    sLC = Mid(RcvBuffer, 2, 5)

    Select Case sBC
        Case "R"
            Call Sleep(200)     '500)

            sMode = Mid(RcvBuffer, 5, 1)
            pSampleInfo.RACK = Mid(RcvBuffer, 21, 6)
            pSampleInfo.POS = Mid(RcvBuffer, 27, 2)
            pSampleInfo.ID = Trim(Mid(RcvBuffer, 6, 15))

'            'Bacode Length 변수로 실제바코드번호얻음...
'            pSampleInfo.ID = Right(pSampleInfo.ID, m_iBCLen)

            Call SendOrder_UF1000
            
            '===============================
            '=== Analysis Order Format 1 ===
            '===============================
            msSendBuf = ""
            msSendBuf = msSendBuf & "S1441"
            msSendBuf = msSendBuf & Format(Now, "YYYYMMDD")
            msSendBuf = msSendBuf & Right(Space(15) & pSampleInfo.ID, 15)
            msSendBuf = msSendBuf & Right(Space(6) & pSampleInfo.RACK, 6)
            msSendBuf = msSendBuf & Right(Space(2) & pSampleInfo.POS, 2)
            '--- Inquiry Mode ---
            '1:Real-time, 2:Batch
            msSendBuf = msSendBuf & sMode
            '--- Order ---
            '0:Not alalyze, 1:Sediment(SEDch+BACch), 2:Only Bacteria(BACch)
            If pSampleInfo.ORDCNT = 0 Then
                msSendBuf = msSendBuf & "0"
            Else
                msSendBuf = msSendBuf & "1"
            End If
            '--- Patient ID ---
            msSendBuf = msSendBuf & Space(16)
            '--- Sample Comment ---
            msSendBuf = msSendBuf & Space(40)
            
            msSendBuf = msSendBuf & Format(Now, "YYYYMMDD")
            msSendBuf = msSendBuf & Format(Now, "HH:NN")
            msSendBuf = msSendBuf & "***"
            '--- Reserved ---
            msSendBuf = msSendBuf & String(143, "0")
            
            msSendBuf = Chr(2) & msSendBuf & Chr(3)

            '===============================
            '=== Analysis Order Format 2 ===
            '===============================
            msSendBuf2 = ""
            msSendBuf2 = msSendBuf2 & "S2441"
            msSendBuf2 = msSendBuf2 & Format(Now, "YYYYMMDD")
            msSendBuf2 = msSendBuf2 & Right(Space(15) & pSampleInfo.ID, 15)
            msSendBuf2 = msSendBuf2 & Right(Space(6) & pSampleInfo.RACK, 6)
            msSendBuf2 = msSendBuf2 & Right(Space(2) & pSampleInfo.POS, 2)
            '--- Inquiry Mode ---
            '1:Real-time, 2:Batch
            msSendBuf2 = msSendBuf2 & sMode
            '--- Patient ID ---
            msSendBuf2 = msSendBuf2 & Space(16)
            '--- Family Name ---
            msSendBuf2 = msSendBuf2 & Space(20)
            '--- Given Name ---
            msSendBuf2 = msSendBuf2 & Space(20)
            '--- Sex ---
            msSendBuf2 = msSendBuf2 & "0"
            '--- Date of Birth ---
            msSendBuf2 = msSendBuf2 & Space(8)
            '--- Patient Comment ---
            msSendBuf2 = msSendBuf2 & Space(100)
            '--- Attending Physician ---
            msSendBuf2 = msSendBuf2 & Space(20)
            '--- Ward ---
            msSendBuf2 = msSendBuf2 & Space(20)
            '--- Reserved ---
            msSendBuf2 = msSendBuf2 & String(11, "0")
            
            msSendBuf2 = Chr(2) & msSendBuf2 & Chr(3)
            
            '=== Analysis Order Format 1 장비로 전송 ===
            msComm.Output = msSendBuf
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintSendLog(msSendBuf)
            End If
            
            msSendBuf = ""
            
            Exit Sub

        Case "D"
            Select Case sLC
                Case "S4401"    'Sample Information Block
                    '결과정보 초기화
                    Call Init_pResultInfo

                    Call Init_RstVar   '임시변수 초기화
                    
                    msRack = Trim(Mid(RcvBuffer, 63, 6))
                    msPos = Trim(Mid(RcvBuffer, 69, 2))
                    msBarCd = Trim(Mid(RcvBuffer, 71, 15))
                    
'                    If IsNumeric(msBarCd) = True Then
''                        msBarCd = Trim(CStr(Val(msBarCd)))
'                        msBarCd = Right(msBarCd, m_iBCLen)
'                    End If
                    
                    'REVIEW
                    If Mid(RcvBuffer, 88, 1) = "1" Then msReviewFlag = "REVIEW"
                    'Analysis error
                    If Mid(RcvBuffer, 89, 1) = "1" Then msEqErrFlag = "ERROR"
                                        
                    'FLAG
                    If Mid(RcvBuffer, 140, 1) = "+" Then msReview = msReview & "RBC" & "^"
                    If Mid(RcvBuffer, 141, 1) = "+" Then msReview = msReview & "WBC" & "^"
                    If Mid(RcvBuffer, 142, 1) = "+" Then msReview = msReview & "EC" & "^"
                    If Mid(RcvBuffer, 143, 1) = "+" Then msReview = msReview & "CAST" & "^"
                    If Mid(RcvBuffer, 144, 1) = "+" Then msReview = msReview & "BACT" & "^"
                    If Mid(RcvBuffer, 145, 1) = "+" Then msReview = msReview & "Cond" & "^"
                    
                    If Mid(RcvBuffer, 140, 1) = "*" Then msReview = msReview & "RBC*" & "^"
                    If Mid(RcvBuffer, 141, 1) = "*" Then msReview = msReview & "WBC*" & "^"
                    If Mid(RcvBuffer, 142, 1) = "*" Then msReview = msReview & "EC*" & "^"
                    If Mid(RcvBuffer, 143, 1) = "*" Then msReview = msReview & "CAST*" & "^"
                    If Mid(RcvBuffer, 144, 1) = "*" Then msReview = msReview & "BACT*" & "^"
                    If Mid(RcvBuffer, 145, 1) = "*" Then msReview = msReview & "Cond*" & "^"
                    
                Case "P4402"    'Particle Count Block 1
                    iCnt = Mid(RcvBuffer, 48, 2)
                    RcvBuffer = Mid(RcvBuffer, 50)
                    
                    If IsNumeric(iCnt) Then
                        For ii = 1 To Val(iCnt)
                            sIFRstCd = Mid(RcvBuffer, 1 + 12 * (ii - 1), 4)
                            sRst = Mid(RcvBuffer, 1 + 12 * (ii - 1) + 4, 8)
                            
                            If IsNumeric(sRst) = True Then
                                sRst = Trim(CStr(Val(sRst)))
''                            Else
''                                sRst = "ERROR"
                            End If
                                                        
                            miRstCnt = miRstCnt + 1
                            msTotIFCd = msTotIFCd & sIFRstCd & Chr(124)
                            msTotRst = msTotRst & sRst & Chr(124)
                            msTotRst2 = msTotRst2 & Chr(124)
                            msUnit = msUnit & Chr(124)
                            msFlag = msFlag & Chr(124)
                        Next ii
                    End If
                
                Case "C4403"    'Comment Block 1
                    iCnt = Mid(RcvBuffer, 48, 2)
                    RcvBuffer = Mid(RcvBuffer, 50)
                    
                    If IsNumeric(iCnt) Then
                        For ii = 1 To Val(iCnt)
                            sIFRstCd = Mid(RcvBuffer, 1 + 4 * (ii - 1), 4)
                            Select Case sIFRstCd
                                Case "00D9"
                                    msReview = msReview & "P.CAST" & "^"
                                Case "0107"
                                    msReview = msReview & "SRC" & "^"
                                Case "0501"
                                    msReview = msReview & "SPERM" & "^"
                                Case "0300"
                                    msReview = msReview & "XTAL" & "^"
                                Case "0402"
                                    msReview = msReview & "YLC" & "^"
                                Case "00DA"
                                    msReview = msReview & "MUCUS" & "^"
                            End Select
                            
'                            miRstCnt = miRstCnt + 1
'                            msTotIFCd = msTotIFCd & sIFRstCd & "F" & Chr(124)
'                            msTotRst = msTotRst & "+" & Chr(124)
'                            msTotRst2 = msTotRst2 & Chr(124)
'                            msUnit = msUnit & Chr(124)
'                            msFlag = msFlag & Chr(124)
                        Next ii
                    End If
                    
                Case "Q4404"    'Particle Count Block 2
                    iCnt = Mid(RcvBuffer, 48, 2)
                    RcvBuffer = Mid(RcvBuffer, 50)
                    
                    If IsNumeric(iCnt) Then
                        For ii = 1 To Val(iCnt)
                            sIFRstCd = Mid(RcvBuffer, 1 + 12 * (ii - 1), 4)
                            sRst = Mid(RcvBuffer, 1 + 12 * (ii - 1) + 4, 8)
                            
                            If IsNumeric(sRst) = True Then
                                sRst = Trim(CStr(Val(sRst)))
''                            Else
''                                sRst = "ERROR"
                            End If
                                                        
                            miRstCnt = miRstCnt + 1
                            msTotIFCd = msTotIFCd & sIFRstCd & Chr(124)
                            msTotRst = msTotRst & sRst & Chr(124)
                            msTotRst2 = msTotRst2 & Chr(124)
                            msUnit = msUnit & Chr(124)
                            msFlag = msFlag & Chr(124)
                        Next ii
                    End If
                    
                    miRstCnt = miRstCnt + 1
                    msTotIFCd = msTotIFCd & "REVIEW" & Chr(124)
                    msTotRst = msTotRst & msReview & Chr(124)
                    msTotRst2 = msTotRst2 & Chr(124)
                    msUnit = msUnit & Chr(124)
                    msFlag = msFlag & Chr(124)
                
                Case "D4405"    'Comment Block 2
                    iCnt = Mid(RcvBuffer, 48, 2)
                    RcvBuffer = Mid(RcvBuffer, 50)
                    
                    If IsNumeric(iCnt) Then
                        For ii = 1 To Val(iCnt)
                            sIFRstCd = Mid(RcvBuffer, 1 + 12 * (ii - 1), 4)
                            sRst = Mid(RcvBuffer, 1 + 12 * (ii - 1) + 4, 8)
                            
                            If IsNumeric(sRst) = True Then
                                sRst = Trim(CStr(Val(sRst)))
''                            Else
''                                sRst = "ERROR"
                            End If
                            
                            miRstCnt = miRstCnt + 1
                            msTotIFCd = msTotIFCd & sIFRstCd & Chr(124)
                            msTotRst = msTotRst & sRst & Chr(124)
                            msTotRst2 = msTotRst2 & Chr(124)
                            msUnit = msUnit & Chr(124)
                            msFlag = msFlag & Chr(124)
                        Next ii
                    End If
                    
                    With pResultInfo
                        .ID = msBarCd
                        .RACK = msRack
                        .POS = msPos
                        .RSTCNT = miRstCnt
                        
                        .IFCD = msTotIFCd
                        .RST1 = msTotRst
                        .RST2 = msTotRst2
                        .UNIT = msUnit
                        .FLAG = msFlag
                    End With

                    '결과 처리
                    With pResultInfo
                        RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .KIND, msReviewFlag, msEqErrFlag, "")
                    End With
                    
                    Call Init_RstVar   '임시변수 초기화
                
            End Select

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 에러 발생 - " & Err.Description)
    End If
End Sub


'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=7,0,0,0
Public Property Get iBCLen() As Integer
    iBCLen = m_iBCLen
End Property

Public Property Let iBCLen(ByVal New_iBCLen As Integer)
    m_iBCLen = New_iBCLen
    PropertyChanged "iBCLen"
End Property

Private Sub PhaseCfg_Protocol_UF1000()
    On Error GoTo ErrHandler
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)

        Select Case Asc(wkDat)
            Case 2      'STX
                RcvBuffer = ""
            
            Case 3      'ETX
                Call Sleep(200)     '500)
                
                msComm.Output = Chr(6)       'ACK
                
                If sTestMode = "77" Then
                    RaiseEvent PrintSendLog(Chr(6))
                End If
                
                Call DataEditResponse_UF1000
                
            Case 6      'ACK
                If msSendBuf2 <> "" Then
                    Call Sleep(500)
                    
                    '=== Analysis Order Format 2 장비로 전송 ===
                    msComm.Output = msSendBuf2
                    
                    RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
                    
                    If m_sTestMode = "77" Then
                        RaiseEvent PrintSendLog(msSendBuf2)
                    End If
                    
                    msSendBuf = ""
                    msSendBuf2 = ""
                End If
                
            Case Else
                RcvBuffer = RcvBuffer & wkDat
        End Select
    Next ix1
    
    Exit Sub
ErrHandler:
    RaiseEvent DispMsg("PhaseCfg_Protocol_UF1000 에러 - " & Err.Description)
End Sub

Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    
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
        .KIND = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
        .RSTDT = ""
        .OTHER = ""
    End With
    
End Sub

Private Sub Init_RstVar()

    msBarCd = "": msRack = "": msPos = ""
    msReviewFlag = "": msEqErrFlag = ""
    msReview = "":  msRBCInfo = ""
    
    miRstCnt = 0
    msTotIFCd = "": msTotRst = "": msTotRst2 = "": msUnit = "": msFlag = ""
                   
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

Private Sub SendOrder_UF1000()
    On Error GoTo ErrRtn
    
    Dim SendBuf$, sBuf$
    Dim iPos%, i%
    Dim sOrder$
    
    RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
    
    Call Get_OrderString
    
    If pSampleInfo.ORDCNT = 0 Then
        RaiseEvent DispMsg("인터페이스 오더 항목이 존재하지 않습니다!!")
    Else
    
    End If
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("SendOrder 에러발생 - " & Err.Description)
    End If
End Sub

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
    On Error GoTo ErrHandler
        
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
    
    Exit Sub
    
ErrHandler:
    RaiseEvent DispMsg("msComm_OnComm 에러 - " & Err.Description)
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

