VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.UserControl LX 
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4065
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   4065
   Begin FPSpread.vaSpread spdBarCode 
      Height          =   1140
      Left            =   1620
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   225
      Width           =   2340
      _Version        =   196608
      _ExtentX        =   4128
      _ExtentY        =   2011
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      MaxRows         =   0
      ScrollBars      =   2
      SpreadDesigner  =   "LX.ctx":0000
      UserResize      =   1
      TextTip         =   1
      ScrollBarTrack  =   3
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
   End
End
Attribute VB_Name = "LX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_p_sSpcCd = "0"
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
Dim m_p_sSpcCd As String
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
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTAlarmCd$, sKind$)
Event RaiseError(sError$)
Event RequestCurOrder(sID$)
Event SendOrderOK(sID$, sSeqNo$, sRack$, sPos$)
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

'For CX
Private pCXInfo As CXINFO

Private Sub DataEditResponse_CX9()
'    On Error GoTo ErrRtn
'
'    Dim iPos1%, iPos2%, ix1%
'    Dim sSF     As String
'    Dim sFC     As String
'    Dim sRC     As String
'    Dim sHQ     As String
'
'    Dim tmpField()  As String
'    Dim tmpData()   As String
'
'    Dim tmpBarCd$, tmpRack$, tmpPos$, tmpSeqNo$, tmpSpcCd$
'    Dim tmpIFCd$, tmpRst$, tmpFlag$
'
'    iPos1 = InStr(RcvBuffer, "[")
'    iPos2 = InStr(RcvBuffer, "]")
'
'    Do While (iPos1 > 0)
'        sSF = Mid$(RcvBuffer, iPos1 + 4, 3)
'        sFC = Mid$(RcvBuffer, iPos1 + 8, 2)
'        sRC = Trim(Val(Trim(Mid$(RcvBuffer, iPos1 + 11, 2))))
'        sHQ = Mid$(RcvBuffer, iPos1 + 11, iPos2 - iPos1 - 11)   'HOST QUERY
'
'        Select Case sSF
'            ' ===== Order Arr.
'            Case "701"
'                Select Case sFC
'                    Case "02"
'                        Select Case sRC
'                            Case "0"
'                                With pSampleInfo
'                                    If .ORDCNT > 0 Then
'                                        RaiseEvent SendOrderOK(.ID, .SEQNO, .RACK, .POS)
'                                    Else
'                                        '조회된 내용이 없는 경우 환자정보 구조체 초기화
'                                        Call Init_pResultInfo
'
'                                        RaiseEvent SendOrderOK("", "", "", "")
'                                    End If
'                                End With
'
'                                Call SendNextOrder
'
'                            Case "1"
'                                RaiseEvent DispMsg("[701,02,01] SYNTAX ERROR")
'
'                            Case "2"
'                                RaiseEvent DispMsg("[701,02,02] BUSY")
'
'                            Case "3"
'                                RaiseEvent DispMsg("[701,02,03] INVALID CHEMISTRY REQUESTED")
'
'                            Case "4"
'                                RaiseEvent DispMsg("[701,02,04] INVALID ORDAC REQUESTED")
'
'                            Case "5"
'                                RaiseEvent DispMsg("[701,02,05] INVALID CHEMISTRY COMBINATION PROGRAMMED")
'
'                            Case "6"
'                                RaiseEvent DispMsg("[701,02,06] CONTROL NOT CONFIGURED")
'
'                            Case "7"
'                                RaiseEvent DispMsg("[701,02,07] CALIBRATOR SECTOR ONLY")
'
'                            Case "8"
'                                RaiseEvent DispMsg("[701,02,08] MODE MISMATCH")
'
'                            Case "9"
'                                RaiseEvent DispMsg("[701,02,09] CX7 ERROR")
'
'                            Case "10"
'                                RaiseEvent DispMsg("[701,02,10] COMPLETED SAMPLE")
'
'                            Case "11"
'                                RaiseEvent DispMsg("[701,02,11] Incompatible Fluid Types")
'
'                            Case "12"
'                                RaiseEvent DispMsg("[701,02,12] Incompatible Test Types")
'
'                            Case "13"
'                                RaiseEvent DispMsg("[701,02,13] Incompatible Patient Name")
'                        End Select
'
'                    Case "04"   'Clear Sector/Sample IDs가 없으므로 무의미 (나중을 위하여!)
'                        Select Case sRC
'                            Case "0"
'                                Sleep (500)
'
'                                m_iPhase = 5
'                                msComm.Output = Chr(4) & Chr(1) 'EOT+SOH 송신
'
'                            Case "1"
'                                RaiseEvent DispMsg("[701,04,01] BAD MESSAGE")
'
'                            Case "2"
'                                RaiseEvent DispMsg("[701,04,02] BUSY")
'
'                            Case "3"
'                                RaiseEvent DispMsg("[701,04,03] CX7 ERROR")
'
'                            Case "4"
'                                RaiseEvent DispMsg("[701,04,04] NOT EXISTENT ERROR")
'                        End Select
'
'                    Case "06"   'HOST QUERY
'                        With spdBarCode
'                            tmpField() = Split(sHQ, ",")
'
'                            For ix1 = 0 To UBound(tmpField())
'                                If ix1 > 6 Then Exit For
'
'                                If Trim(tmpField(ix1)) <> "" Then
'                                    .MaxRows = .MaxRows + 1
'                                    Call .SetText(1, .MaxRows, Trim(tmpField(ix1)))
'                                End If
'                            Next ix1
'                        End With
'
'                        Call SendNextOrder
'
'                End Select
'
'            ' ===== Result Arr.
'            Case "702"
'                Select Case sFC
'                    Case "01"
'                        Call Init_pResultInfo
'
'                        tmpField() = Split(RcvBuffer, ",")
'                        tmpSeqNo = Trim(tmpField(5))
'                        tmpRack = Trim(tmpField(7))
'                        tmpPos = Trim(tmpField(8))
'                        tmpSpcCd = Trim(tmpField(11))
'                        tmpBarCd = Trim(tmpField(12))
'
'                        With pResultInfo
'                            .ID = tmpBarCd
'                            .SEQNO = tmpSeqNo
'                            .RACK = tmpRack
'                            .POS = tmpPos
'                        End With
'
'                    Case "03"
'                        tmpField() = Split(RcvBuffer, ",")
'
'                        tmpIFCd = Trim(tmpField(10))
'                        tmpFlag = Trim(tmpField(22))
'                        If tmpFlag = "NA" Then
'                            tmpFlag = ""
'                        End If
'                        tmpRst = Trim(tmpField(15))     '25))
'
'                        If (IsNumeric(tmpRst) = False) Then
'                            tmpRst = ""
'                        End If
'                        If Left(tmpRst, 1) = "." Then
'                            tmpRst = "0" & tmpRst
'                        End If
'
'                        '결과정보 구조체에 저장
'                        With pResultInfo
'                            '결과값 누적
'                            .RSTCNT = .RSTCNT + 1
'                            .IFCD = .IFCD & tmpIFCd & Chr(124)
'                            .RST1 = .RST1 & tmpRst & Chr(124)
'                            .RST2 = .RST2 & Chr(124)
'                            .UNIT = .UNIT & Chr(124)
'                            .FLAG = .FLAG & tmpFlag & Chr(124)
'                        End With
'
'                    Case "05"
'                        '결과값 등록/화면 표시 처리...
'                        With pResultInfo
'                            If .RSTCNT > 0 Then
'                                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .kind)
'                            End If
'                        End With
'
'                        Call Init_pResultInfo
'
'                End Select
'        End Select
'
'        iPos1 = InStr(2, RcvBuffer, "[")
'        If iPos1 <> 0 Then
'            RcvBuffer = Mid(RcvBuffer, iPos1)
'            iPos1 = 1
'        End If
'    Loop
'
'ErrRtn:
'    If Err <> 0 Then
'        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
'    End If
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

Private Sub DispReturnMsg_LX20(ByVal sStream As String, ByVal sFunction As String, ByVal sRetCd As String)
    
    Select Case sStream
        Case "801"
            Select Case sFunction
                Case "02"
                    Select Case sRetCd
                        Case "1"
                            RaiseEvent DispMsg("[801,02,01] BAD MESSAGE")
                        Case "2"
                            RaiseEvent DispMsg("[801,02,02] BUSY")
                        Case "3"
                            RaiseEvent DispMsg("[801,02,03] NOT CONFIGURED")
                        Case "4"
                            RaiseEvent DispMsg("[801,02,04] NON ORDAC")
                        Case "5"
                            RaiseEvent DispMsg("[801,02,05] DILUTION ERROR")
                        Case "6"
                            RaiseEvent DispMsg("[801,02,06] CONTROL NOT CONFIGURED")
                        Case "7"
                            RaiseEvent DispMsg("[801,02,07] CALIBRATOR Rack ONLY")
                        Case "8"
                            RaiseEvent DispMsg("[801,02,08] Not Used")
                        Case "9"
                            RaiseEvent DispMsg("[801,02,09] LX20 ERROR")
                        Case "10"
                            RaiseEvent DispMsg("[801,02,10] Completed Sample")
                        Case "11"
                            RaiseEvent DispMsg("[801,02,11] Incompatible Fluid Types")
                        Case "12"
                            RaiseEvent DispMsg("[801,02,12] Incompatible Test Types")
                        Case "13"
                            RaiseEvent DispMsg("[801,02,13] Incompatible Patient Name")
                        Case "14"
                            RaiseEvent DispMsg("[801,02,14] Sample ID matches existing Control ID")
                        Case "15"
                            RaiseEvent DispMsg("[801,02,15] Rack Number too large")
                        Case Else
                            RaiseEvent DispMsg("[801,02," & sRetCd & "] Unknown Return Code")
                    End Select
                
                Case "04"
                    Select Case sRetCd
                        Case "1"
                            RaiseEvent DispMsg("[801,04,01] BAD MESSAGE")
                        Case "2"
                            RaiseEvent DispMsg("[801,04,02] BUSY")
                        Case "3"
                            RaiseEvent DispMsg("[801,04,03] LX20 ERROR")
                        Case "4"
                            RaiseEvent DispMsg("[801,04,04] Non-existent ERROR")
                        Case "5"
                            RaiseEvent DispMsg("[801,04,05] Rack Number too large")
                        Case Else
                            RaiseEvent DispMsg("[801,04," & sRetCd & "] Unknown Return Code")
                    End Select
                
            End Select
        
    End Select
    
End Sub

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
        Case "LX20"
            Call PhaseCfg_Protocol_LX20
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_LX20()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
             
        Select Case m_iPhase
            Case 1      '===== EOT 대기
                Select Case Asc(wkDat)
                    Case 4      '----- EOT 수신
                        m_iPhase = 2
                    Case Else
                        m_iPhase = 1
                End Select
                
            Case 2      '===== SOH 대기
                Select Case Asc(wkDat)
                    Case 1      '----- SOH 수신
                        msComm.Output = Chr(6)  'ACK 송신
                        piAckEtx = 2
                        m_iPhase = 3
                        RcvBuffer = ""
                End Select
                
            Case 3      '===== LF 대기
                Select Case Asc(wkDat)
                    Case 10     '----- LF 수신
                        Select Case piAckEtx
                            Case 1
                                msComm.Output = Chr(6)  'ACK 송신
                                piAckEtx = 2
                            Case 2
                                msComm.Output = Chr(3)  'ETX 송신
                                piAckEtx = 1
                        End Select
                        m_iPhase = 4
                        
                    Case Else   '----- 문자 수신
                        RcvBuffer = RcvBuffer & wkDat
                End Select
            
            Case 4      '===== EOT 대기
                Select Case Asc(wkDat)
                    Case 4      '----- EOT 수신
                        m_iPhase = 1
                        
                        ' Interface에서 받은 데이타 편집
                        Call DataEditResponse_LX20
                        
                        If pbContension = True Then
                            Sleep (500)
                            
                            msComm.Output = Chr(4) & Chr(1) 'EOT+SOH 송신
                            pbContension = False
                            m_iPhase = 5
                        End If
                        
                    Case Else   '----- 문자 수신
                        RcvBuffer = RcvBuffer + wkDat
                        m_iPhase = 3
                End Select
            
            Case 5      '===== ACK 대기
                Select Case Asc(wkDat)
                    Case 6      '----- ACK 수신
                        Call SendOrder_LX20
                        
                        m_iPhase = 6
                    Case 4      '----- EOT 수신
                        pbContension = True
                        m_iPhase = 2
                    Case Else
                End Select
            
            Case 6      '===== ETX 대기
                Select Case Asc(wkDat)
                    Case 3      '----- ETX 수신 (ORDER주었을 경우만 반응)
                        msComm.Output = Chr(4)      'EOT
                        m_iPhase = 1
                        
                    Case 21     '----- NAK 수신 !!!!!
                        Sleep (500)
                        
                        msComm.Output = Chr(4) & Chr(1)
                        m_iPhase = 7
                        
                    Case 4      'EOT
                        pbContension = True
                        m_iPhase = 2
                        
                End Select
    
            Case 7      'NAK가 온 경우 재전송 처리
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        msComm.Output = psNakBuf    'NAK message 재전송
                        m_iPhase = 6
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(psNakBuf)
                        End If
    
                    Case 4      '----- EOT 수신
                        pbContension = True
                        m_iPhase = 2
                        
                End Select
                
        End Select
    Next ix1
    
End Sub
Private Sub PhaseCfg_Protocol_CX9()
    
    Dim wkDat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
             
        Select Case m_iPhase
            Case 1      '===== EOT 대기
                Select Case Asc(wkDat)
                    Case 4      '----- EOT 수신
                        m_iPhase = 2
                    Case Else
                        m_iPhase = 1
                End Select
                
            Case 2      '===== SOH 대기
                Select Case Asc(wkDat)
                    Case 1      '----- SOH 수신
                        msComm.Output = Chr(6)  'ACK 송신
                        piAckEtx = 2
                        m_iPhase = 3
                        RcvBuffer = ""
                End Select
                
            Case 3      '===== LF 대기
                Select Case Asc(wkDat)
                    Case 10     '----- LF 수신
                        Select Case piAckEtx
                            Case 1
                                msComm.Output = Chr(6)  'ACK 송신
                                piAckEtx = 2
                            Case 2
                                msComm.Output = Chr(3)  'ETX 송신
                                piAckEtx = 1
                        End Select
                        m_iPhase = 4
                        
                    Case Else   '----- 문자 수신
                        RcvBuffer = RcvBuffer & wkDat
                End Select
            
            Case 4      '===== EOT 대기
                Select Case Asc(wkDat)
                    Case 4      '----- EOT 수신
                        m_iPhase = 1
                        
                        ' Interface에서 받은 데이타 편집
                        Call DataEditResponse_CX9
                        
                        If pbContension = True Then
                            Sleep (500)
                            
                            msComm.Output = Chr(4) & Chr(1) 'EOT+SOH 송신
                            pbContension = False
                            m_iPhase = 5
                        End If
                        
                    Case Else   '----- 문자 수신
                        RcvBuffer = RcvBuffer + wkDat
                        m_iPhase = 3
                End Select
            
            Case 5      '===== ACK 대기
                Select Case Asc(wkDat)
                    Case 6      '----- ACK 수신
                        Call SendOrder_CX9
                        
                        m_iPhase = 6
                    Case 4      '----- EOT 수신
                        pbContension = True
                        m_iPhase = 2
                    Case Else
                End Select
            
            Case 6      '===== ETX 대기
                Select Case Asc(wkDat)
                    Case 3      '----- ETX 수신 (ORDER주었을 경우만 반응)
                        msComm.Output = Chr(4)      'EOT
                        m_iPhase = 1
                        
                    Case 21     '----- NAK 수신 !!!!!
                        Sleep (500)
                        
                        msComm.Output = Chr(4) & Chr(1)
                        m_iPhase = 7
                        
                    Case 4      'EOT
                        pbContension = True
                        m_iPhase = 2
                        
                End Select
    
            Case 7      'NAK가 온 경우 재전송 처리
                Select Case Asc(wkDat)
                    Case 6      'ACK
                        msComm.Output = psNakBuf    'NAK message 재전송
                        m_iPhase = 6
                        
                        If m_sTestMode = "77" Then
                            RaiseEvent PrintSendLog(psNakBuf)
                        End If
    
                    Case 4      '----- EOT 수신
                        pbContension = True
                        m_iPhase = 2
                        
                End Select
                
        End Select
    Next ix1
    
End Sub


Private Sub Get_OrderString()

    Dim ii      As Integer
    Dim tmpData()   As String
    Dim iCnt    As Integer
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .SEQNO = m_p_sSeq
            .RACK = m_p_sRack
            .POS = m_p_sPos
            .SPCCD = m_p_sSpcCd
            .ORDCNT = 0
            Erase .IFCD     '2003/4/16
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
        .SPCCD = m_p_sSpcCd
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
            
            If m_sTestMode = "77" Then
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
    m_p_sSpcCd = PropBag.ReadProperty("p_sSpcCd", m_def_p_sSpcCd)
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
    Call PropBag.WriteProperty("p_sSpcCd", m_p_sSpcCd, m_def_p_sSpcCd)
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
    
    '변수 초기화(CX 계열)
    piAckEtx = 1: pbContension = False
    
    
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

Private Sub DataEditResponse_LX20()
    On Error GoTo ErrRtn
    
    Dim iPos1%, iPos2%, ix1%
    Dim sSF     As String
    Dim sFC     As String
    Dim sRC     As String
    Dim sHQ     As String
    
    Dim tmpField()  As String
    Dim tmpData()   As String
    
    Dim tmpBarCd$, tmpRack$, tmpPos$, tmpSeqNo$, tmpSpcCd$, tmpTestType$, tmpContNm$
    Dim tmpIFCd$, tmpRst$, tmpFlag$, tmpInstCd$
    
    Dim tmpRstErr$
    Dim ii%
        
    iPos1 = InStr(RcvBuffer, "[")
    iPos2 = InStr(RcvBuffer, "]")
    
    Do While (iPos1 > 0)
        sSF = Mid$(RcvBuffer, iPos1 + 4, 3)
        sFC = Mid$(RcvBuffer, iPos1 + 8, 2)
        sRC = Trim(Val(Trim(Mid$(RcvBuffer, iPos1 + 11, 2))))
        sHQ = Mid$(RcvBuffer, iPos1 + 11, iPos2 - iPos1 - 11)   'HOST QUERY
        
        Select Case sSF
            ' ===== Order Arr.
            Case "801"
                Select Case sFC
                    Case "02"
                        Select Case sRC
                            Case "0"        'OK
                                With pSampleInfo
                                    If .ORDCNT > 0 Then
                                        RaiseEvent SendOrderOK(.ID, .SEQNO, .RACK, .POS)
                                    Else
                                        '조회된 내용이 없는 경우 환자정보 구조체 초기화
                                        Call Init_pResultInfo
                                        
                                        RaiseEvent SendOrderOK("", "", "", "")
                                    End If
                                End With
                                
                                Call SendNextOrder
                            
                            Case Else
                                Call DispReturnMsg_LX20(sSF, sFC, sRC)
                            
                        End Select
                        
                    Case "04"   'Clear Sector/Sample IDs가 없으므로 무의미 (나중을 위하여!)
                        Select Case sRC
                            Case "0"
                                Sleep (500)
                                
                                m_iPhase = 5
                                msComm.Output = Chr(4) & Chr(1) 'EOT+SOH 송신
                            
                            Case Else
                                Call DispReturnMsg_LX20(sSF, sFC, sRC)
                                
                        End Select
                        
                    Case "06"   'HOST QUERY
                        With spdBarCode
                            Erase tmpField()
                            tmpField() = Split(sHQ, ",")
                        
                            For ix1 = 0 To UBound(tmpField())
                                If ix1 > 3 Then Exit For
                                
                                If Trim(tmpField(ix1)) <> "" Then
                                    .MaxRows = .MaxRows + 1
                                    Call .SetText(1, .MaxRows, Trim(tmpField(ix1)))
                                End If
                            Next ix1
                        End With
                        
                        Call SendNextOrder
                        
                End Select
                
            ' ===== Result Arr.
            Case "802"
                Select Case sFC
                    Case "01"
                        Call Init_pResultInfo
                        
                        Erase tmpField()
                        tmpField() = Split(RcvBuffer, ",")
                        
                        tmpSeqNo = Trim(tmpField(5))
                        tmpRack = Trim(tmpField(7))
                        tmpPos = Trim(tmpField(8))
                        tmpTestType = Trim(tmpField(9))
                        tmpSpcCd = Trim(tmpField(11))
                        tmpBarCd = Trim(tmpField(12))
                        tmpContNm = Trim(tmpField(13))  'control name
                        
                        If tmpTestType = "CO" Then      'Control 결과
                            tmpBarCd = tmpContNm
                        End If
                        
                        With pResultInfo
                            .ID = tmpBarCd
                            .SEQNO = tmpSeqNo
                            .RACK = tmpRack
                            .POS = tmpPos
                            .KIND = tmpTestType
                        End With
                        
                    Case "03"
                        Erase tmpField()
                        tmpField() = Split(RcvBuffer, ",")
                        
                        tmpIFCd = Trim(tmpField(10))
                        tmpFlag = Trim(tmpField(23))
                        If tmpFlag = "NA" Then
                            tmpFlag = ""
                        End If
                        tmpRst = Trim(tmpField(15))     '24))
                        
                        If (IsNumeric(tmpRst) = False) Then
                            If Trim(tmpRst) = "#########" Then
                                tmpRst = "#"
                            Else
                                tmpRst = ""
                            End If
                        End If
                        If Left(tmpRst, 1) = "." Then
                            tmpRst = "0" & tmpRst
                        End If
                                       
                        tmpRstErr = ""
                        For ii = 26 To 41   'Result error code
                            If Trim(tmpField(ii)) = "NO" Then
                            Else
                                tmpField(ii) = ConvertRstErrCd(tmpField(ii))
                                
                                If Trim(tmpRstErr) = "" Then
                                    tmpRstErr = Trim(tmpField(ii))
                                Else
                                    tmpRstErr = tmpRstErr & "," & Trim(tmpField(ii))
                                End If
                            End If
                        Next ii
                        
                        '결과정보 구조체에 저장
                        With pResultInfo
                            '결과값 누적
                            .RSTCNT = .RSTCNT + 1
                            .IFCD = .IFCD & tmpIFCd & Chr(124)
                            .RST1 = .RST1 & tmpRst & Chr(124)
                            .RST2 = .RST2 & Chr(124)
                            .UNIT = .UNIT & Chr(124)
                            .FLAG = .FLAG & tmpFlag & Chr(124)
                            .ALARMCD = .ALARMCD & tmpRstErr & Chr(124)
                        End With
            
                    Case "11"       'Calculations Result
                        Erase tmpField()
                        tmpField() = Split(RcvBuffer, ",")
                        
                        tmpIFCd = Trim(tmpField(10))
                        tmpFlag = ""
                        tmpInstCd = Trim(tmpField(11))      'Calc Status
                        If tmpInstCd = "OK" Then
                            tmpInstCd = ""
                        End If
                        tmpRst = Trim(tmpField(12))
                        
                        If (IsNumeric(tmpRst) = False) Then
                            If Trim(tmpRst) = "#########" Then
                                tmpRst = "#"
                            Else
                                tmpRst = ""
                            End If
                        End If
                        If Left(tmpRst, 1) = "." Then
                            tmpRst = "0" & tmpRst
                        End If
                                                
                        '결과정보 구조체에 저장
                        With pResultInfo
                            '결과값 누적
                            .RSTCNT = .RSTCNT + 1
                            .IFCD = .IFCD & tmpIFCd & Chr(124)
                            .RST1 = .RST1 & tmpRst & Chr(124)
                            .RST2 = .RST2 & Chr(124)
                            .UNIT = .UNIT & Chr(124)
                            .FLAG = .FLAG & tmpFlag & Chr(124)
                            .ALARMCD = .ALARMCD & tmpInstCd & Chr(124)
                        End With
                    
                    Case "05"
                        '결과값 등록/화면 표시 처리...
                        With pResultInfo
                            If .RSTCNT > 0 Then
                                RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .ALARMCD, .KIND)
                            End If
                        End With
            
                        Call Init_pResultInfo
                    
                End Select
        End Select
        
        iPos1 = InStr(2, RcvBuffer, "[")
        If iPos1 <> 0 Then
            RcvBuffer = Mid(RcvBuffer, iPos1)
            iPos1 = 1
        End If
    Loop
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub
Private Function ConvertRstErrCd(ByVal sErrCd As String) As String
        
    ConvertRstErrCd = ""
                
    Select Case Trim(sErrCd)
        Case "AB"
            ConvertRstErrCd = "Not all programmed tests have a valid result"
        Case "AX"
            ConvertRstErrCd = "Antigen excess"
        Case "BH"
            ConvertRstErrCd = "Blank absorbance high"
        Case "BL"
            ConvertRstErrCd = "Blank absorbance low"
        Case "BN"
            ConvertRstErrCd = "Blank noise"
        Case "BO"
            ConvertRstErrCd = "Blank outlier"
        Case "CH"
            ConvertRstErrCd = "Initial conductance high"
        Case "CL"
            ConvertRstErrCd = "Initial conductance low"
        Case "DH"
            ConvertRstErrCd = "Out of instrument range high"
        Case "DL"
            ConvertRstErrCd = "Out of instrument range low"
        Case "DR"
            ConvertRstErrCd = "Reference signal noise"
        Case "DS"
            ConvertRstErrCd = "Sample signal noise"
        Case "EA"
            ConvertRstErrCd = "Erratic ADC"
        Case "EC"
            ConvertRstErrCd = "Excessive reference drift"
        Case "ES"
            ConvertRstErrCd = "Excessive reference drift"
        Case "GH"
            ConvertRstErrCd = "URDAC high"
        Case "GL"
            ConvertRstErrCd = "URDAC low"
        Case "GT"
            ConvertRstErrCd = "Greater than upper instrument or reportable range"
        Case "HI"
            ConvertRstErrCd = "Initial ADC error high"
        Case "HR"
            ConvertRstErrCd = "Reaction absorbance high"
        Case "IA"
            ConvertRstErrCd = "Initial absorbance either too high or too low"
        Case "IK"
            ConvertRstErrCd = "Bad K value"
        Case "IL"
            ConvertRstErrCd = "Initial rate too low"
        Case "IN"
            ConvertRstErrCd = "Bad NA value"
        Case "IR"
            ConvertRstErrCd = "Initial rate too high"
        Case "LI"
            ConvertRstErrCd = "Initial ADC error low"
        Case "LR"
            ConvertRstErrCd = "Reaction absorbance low"
        Case "LT"
            ConvertRstErrCd = "Less than lower instrument or reportable range"
        Case "OF"
            ConvertRstErrCd = "Number overflow error"
        Case "OH"
            ConvertRstErrCd = "Out of instrument range ORDAC high"
        Case "OK"
            ConvertRstErrCd = "Result was calculated"
        Case "OL"
            ConvertRstErrCd = "Out of instrument range ORDAC low"
        Case "RE"
            ConvertRstErrCd = "Reaction error"
        Case "RH"
            ConvertRstErrCd = "Reaction rate high"
        Case "RL"
            ConvertRstErrCd = "Reaction rate low"
        Case "RN"
            ConvertRstErrCd = "Reaction noise"
        Case "RO"
            ConvertRstErrCd = "Reaction outlier"
        Case "SD"
            ConvertRstErrCd = "Substrate depleted"
        Case "SH"
            ConvertRstErrCd = "Blank rate high"
        Case "SL"
            ConvertRstErrCd = "Blank rate low"
        Case "TM"
            ConvertRstErrCd = "Temperature error"
        Case "UH"
            ConvertRstErrCd = "Out of reportable range high"
        Case "UL"
            ConvertRstErrCd = "Out of reportable range low"
        Case "UO"
            ConvertRstErrCd = "Out of ORDAC reportable range high"
        Case "ZD"
            ConvertRstErrCd = "Zero denominator"
        Case Else
            ConvertRstErrCd = sErrCd
    End Select
        
End Function
Private Sub SendNextOrder()
    On Error GoTo ErrRtn
    
    Dim vTmp
    Dim stmp$
        
NextOrder:
    With spdBarCode
        If .MaxRows = 0 Then
            pSampleInfo.ID = ""
            pSampleInfo.ORDCNT = 0
            Exit Sub
        End If
    
        Call .GetText(1, 1, vTmp)
        pSampleInfo.ID = Trim(vTmp)
        
        .Row = 1
        .Action = ActionDeleteRow
        .MaxRows = .MaxRows - 1
    End With
   
    If pSampleInfo.ID <> "" Then
        '----- 검사항목 조회
        RaiseEvent RequestCurOrder(pSampleInfo.ID)
    
        Call Get_OrderString
    Else
        pSampleInfo.ORDCNT = 0
    End If

    If pSampleInfo.ORDCNT > 0 Then
        Sleep (500)
        
        m_iPhase = 5
        msComm.Output = Chr(4) & Chr(1)
    Else
        If spdBarCode.MaxRows = 0 Then
'            msComm.Output = Chr(4)      'EOT
            m_iPhase = 1
        Else
            GoTo NextOrder
        End If
    End If
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order_Next 에러발생" & "(" & Err.Description & ")")
    End If
End Sub
Private Sub SendOrder_LX20()
    On Error GoTo ErrRtn
    
    Dim sSendBuff   As String
    Dim sTestCd     As String
    Dim iCnt%, iChk%, ix1
    Dim sChksums As String   '계산되어진 checksum
    
    Dim sSpcCd  As String
    
    Select Case Trim(pSampleInfo.SPCCD)
        Case "SE", "SF", "UR", "TU", "PL"
            sSpcCd = Trim(pSampleInfo.SPCCD)
        Case Else
            sSpcCd = "SE"
    End Select
    
    '----- Order 만들기
    sTestCd = ""
    For iCnt = 1 To pSampleInfo.ORDCNT
        sTestCd = sTestCd & "," & Trim$(pSampleInfo.IFCD(iCnt)) & " ,0"
    Next iCnt
   
    '----- 전송할 Order Format 만들기
    sSendBuff = "[ 0,801,01"
    sSendBuff = sSendBuff & ",0000"     'RACK
    sSendBuff = sSendBuff & ",00"       'POS
    sSendBuff = sSendBuff & ",0,RO"
    '2004/7/1 yk
    sSendBuff = sSendBuff & "," & sSpcCd    'Serum:SE, Urine:UR
'    sSendBuff = sSendBuff & ",SE"   'Serum:SE, Urine:UR
    sSendBuff = sSendBuff & "," & Left(Trim(Trim(pSampleInfo.ID)) & Space(15), 15)  '바코드번호(left justified)
    sSendBuff = sSendBuff & "," & String(20, " ")       'control name
    sSendBuff = sSendBuff & "," & String(12, " ")       'qc lot number
    sSendBuff = sSendBuff & "," & String(25, " ")       'sample comment code1
    sSendBuff = sSendBuff & "," & String(18, " ")       'name list
    sSendBuff = sSendBuff & "," & String(15, " ")       'name first
    sSendBuff = sSendBuff & "," & String(1, " ")        'name middle initial
    sSendBuff = sSendBuff & "," & String(15, " ")       'patient id
    sSendBuff = sSendBuff & "," & String(18, " ")       'doctor
    sSendBuff = sSendBuff & "," & String(8, " ")        'draw date
    sSendBuff = sSendBuff & "," & String(4, " ")        'draw time
    sSendBuff = sSendBuff & "," & String(20, " ")       'location
    sSendBuff = sSendBuff & ",000"                      'age
    sSendBuff = sSendBuff & ",5"                        'age unit
    sSendBuff = sSendBuff & "," & String(8, " ")        'birth date
    sSendBuff = sSendBuff & ",M"
    sSendBuff = sSendBuff & "," & String(45, " ")       'patient comments
    sSendBuff = sSendBuff & "," & String(7, " ")
    sSendBuff = sSendBuff & "," & String(4, " ")
    sSendBuff = sSendBuff & "," & String(4, " ")
    sSendBuff = sSendBuff & "," & String(2, " ")        'timed urine creatinine unit
    sSendBuff = sSendBuff & "," & String(6, " ")
    
    '----- 바코드번호에 부여된 Order 갯수
    sSendBuff = sSendBuff & "," & Format(Trim(pSampleInfo.ORDCNT), "000")
    
    '----- Order 붙이기
    sSendBuff = sSendBuff & sTestCd & "]"
    
    '----- checksum 계산
    iChk = 0
    For ix1 = 1 To Len(sSendBuff)
        iChk = iChk + Asc(Mid$(sSendBuff, ix1, 1))
    Next ix1
    iChk = iChk Mod 256
    iChk = 256 - iChk
    
    sChksums = Right$("0" & Hex$(iChk), 2)
    
    sSendBuff = sSendBuff & sChksums & Chr(13) & Chr(10)
    psNakBuf = sSendBuff
    
    msComm.Output = sSendBuff
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSendBuff)
    End If
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub
Private Sub SendOrder_CX9()
    On Error GoTo ErrRtn
    
    Dim sSendBuff   As String
    Dim sTestCd     As String
    Dim iCnt%, iChk%, ix1
    Dim sChksums As String   '계산되어진 checksum
    
    '----- Order 만들기
    sTestCd = ""
    For iCnt = 1 To pSampleInfo.ORDCNT
        sTestCd = sTestCd & "," & Trim$(pSampleInfo.IFCD(iCnt)) & " ,0"
    Next iCnt
   
    '----- 전송할 Order Format 만들기
    sSendBuff = "[ 0,701,01"
    sSendBuff = sSendBuff & ",00"   'RACK
    sSendBuff = sSendBuff & ",00"   'POS
    sSendBuff = sSendBuff & ",0,RO"
    sSendBuff = sSendBuff & ",SE"   'Serum:SE, Urine:UR
    sSendBuff = sSendBuff & "," & Left(Trim(Trim(pSampleInfo.ID)) & Space(11), 11)  '바코드번호(left justified)
    sSendBuff = sSendBuff & "," & String(20, " ")
    sSendBuff = sSendBuff & "," & String(25, " ")
    sSendBuff = sSendBuff & "," & String(25, " ")
    sSendBuff = sSendBuff & "," & String(18, " ")
    sSendBuff = sSendBuff & "," & String(15, " ")
    sSendBuff = sSendBuff & "," & String(1, " ")
    sSendBuff = sSendBuff & "," & String(12, " ")
    sSendBuff = sSendBuff & "," & String(18, " ")
    sSendBuff = sSendBuff & "," & String(6, " ")
    sSendBuff = sSendBuff & "," & String(4, " ")
    sSendBuff = sSendBuff & "," & String(20, " ")
    sSendBuff = sSendBuff & ",000"
    sSendBuff = sSendBuff & ",5"
    sSendBuff = sSendBuff & "," & String(6, " ")
    sSendBuff = sSendBuff & ",M"
    sSendBuff = sSendBuff & "," & String(25, " ")
    sSendBuff = sSendBuff & "," & String(7, " ")
    sSendBuff = sSendBuff & "," & String(4, " ")
    sSendBuff = sSendBuff & "," & String(4, " ")
    sSendBuff = sSendBuff & "," & String(6, " ")
    
    '----- 바코드번호에 부여된 Order 갯수
    sSendBuff = sSendBuff & "," & Format(Trim(pSampleInfo.ORDCNT), "000")
    
    '----- Order 붙이기
    sSendBuff = sSendBuff & sTestCd & "]"
    
    '----- checksum 계산
    iChk = 0
    For ix1 = 1 To Len(sSendBuff)
        iChk = iChk + Asc(Mid$(sSendBuff, ix1, 1))
    Next ix1
    iChk = iChk Mod 256
    iChk = 256 - iChk
    
    sChksums = Right$("0" & Hex$(iChk), 2)
    
    sSendBuff = sSendBuff & sChksums & Chr(13) & Chr(10)
    psNakBuf = sSendBuff
    
    msComm.Output = sSendBuff
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(sSendBuff)
    End If
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
End Sub
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
    m_p_sSpcCd = m_def_p_sSpcCd
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
'MemberInfo=14
Public Function Init_BarCode() As Variant

    spdBarCode.MaxRows = 0
    m_iPhase = 1
    
End Function

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sSpcCd() As String
    p_sSpcCd = m_p_sSpcCd
End Property

Public Property Let p_sSpcCd(ByVal New_p_sSpcCd As String)
    m_p_sSpcCd = New_p_sSpcCd
    PropertyChanged "p_sSpcCd"
End Property

