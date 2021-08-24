VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl VITROS 
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
Attribute VB_Name = "VITROS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
Const m_def_p_sSpcCd = 0
Const m_def_p_sTSVol = "0"
Const m_def_p_sRerunGbn = "0"
Const m_def_p_bSIndex = 0
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
Dim m_p_sSpcCd As Variant
Dim m_p_sTSVol As String
Dim m_p_sRerunGbn As String
Dim m_p_bSIndex As Boolean
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
Event RequestNextOrder()
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$, sTAlarmCd$, sKind$, sTRstDT$, sOther1$)
Event SendOrderOK(sID$, sSeqNo$, sRack$, sPos$)
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

'For VITROS
Dim sP_Type     As String   '패킷의 종류
Dim sKBuffer    As String   'Kermit Buffer
Dim miFileNo    As Integer  'File No
Dim msSendData  As String
Dim miOrdSeq    As Integer


Private Sub Build_SendData()

    Dim tmpData()   As String
    Dim ii%, iCnt%
    Dim sTestDat$
    Dim tmpIFCd$, tmpSpcCd$
    Dim aIFCd() As String
    Dim sCupPos$
    
    msSendData = ""
    
    If m_p_sID = "" Or m_p_iOrdCnt = 0 Then
        With pSampleInfo
            .ID = m_p_sID
            .SEQNO = m_p_sSeq
            .RACK = m_p_sRack
            .POS = m_p_sPos
            .ORDCNT = 0
        End With
    
        Exit Sub
    End If
    
    ReDim tmpData(m_p_iOrdCnt) As String
    tmpData() = Split(m_p_sTIFCd, Chr(124))
    
    With pSampleInfo
        .ID = m_p_sID
        .SEQNO = ""
        .RACK = ""
        .POS = ""
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
    
    '--- ORDER DATA 편집
    sTestDat = ""
    
    With pSampleInfo
        For ii = 1 To .ORDCNT
            tmpIFCd = Trim(.IFCD(ii))
            
            If InStr(tmpIFCd, "^") > 0 Then
                Erase aIFCd()
                aIFCd() = Split(Trim(.IFCD(ii)), "^")
                
                tmpIFCd = Trim(aIFCd(0))
                tmpSpcCd = Trim(aIFCd(1))
                If Len(tmpSpcCd) < 6 Then
                    tmpSpcCd = "11.000"
                End If
            Else
                tmpSpcCd = "11.000"     'DEFAULT 검체 및 DILUTION 지정
            End If
    
            If Chr(Trim(tmpIFCd)) = "#" Then
                sTestDat = sTestDat & "##"
            Else
                sTestDat = sTestDat & Chr(Trim(tmpIFCd))
            End If
        Next ii
        
        If m_bUseBarcode = True Then
            sCupPos = " "
        Else
            sCupPos = Chr(Val(m_p_sPos) + 32)
            If sCupPos = "#" Then
                sCupPos = "##"
            End If
        End If

        If m_bUseBarcode = True Then
            msSendData = Left$(.ID & Space(15), 15)
        Else
            msSendData = Chr(124) & Left$(m_p_sRack & Space(15), 15) & Left$(.ID & Space(15), 15)
        End If
        
        msSendData = msSendData & Mid(tmpSpcCd, 1, 1)   '1:Serum, 2:CSF, 3:Urine
        msSendData = msSendData & "0"                   '0:Rtn, 1:STAT
'        msSendData = msSendData & " "                   'Cup
        msSendData = msSendData & sCupPos               'Cup
        msSendData = msSendData & Mid(tmpSpcCd, 2, 5)   'Dil
        
        msSendData = msSendData & sTestDat & Chr(93)    'D 패킷 누적 'EOS
    End With
    
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
        Case "VITROS950", "VITROS250"
            Call PhaseCfg_Protocol_VITROS
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub

'
'   CheckSum 계산
'
Private Function ChkSum_Kermit(ByVal Para As String) As String

    Dim i  As Integer
    Dim Tmp   As Integer
    Dim ChkS  As Integer
    
    For i = 1 To Len(Para)
        Tmp = Tmp + Asc(Mid$(Para, i, 1))
    Next i
    
    ChkS = (Tmp And 192) / 64
    ChkS = (Tmp + ChkS) And 63
    
    ChkSum_Kermit = Chr(ChkS + 32)
    
End Function
'
'   Y 패킷 송신
'
Private Sub Send_Y()

    Dim Tmp     As String
    Dim RSeq    As String
    
    RSeq = Mid$(RcvBuffer, 2, 1)
    
    Tmp = RSeq & "Y"
    Tmp = ChkLen_Kermit(Tmp) & Tmp
    Tmp = Tmp & ChkSum_Kermit(Tmp)
    
    msComm.Output = Chr(1) & Tmp & Chr(13)

End Sub
'
'   S 패킷 송신
'
Private Sub Send_S()

    Dim TmpS As String
          '          + S~* @-#N1V
          '   Send := ' S~R @-#N1 ';
    TmpS = " S~R @-#N1 "
'   TmpS = " S~* @-#N1V"
    TmpS = ChkLen_Kermit(TmpS) & TmpS       '패킷의 길이
    TmpS = TmpS & ChkSum_Kermit(TmpS)       'Chkeck Sum

    msComm.Output = Chr(1) & TmpS & Chr(13) 'S Packet 송신

    miOrdSeq = 1        '패킷의 번호
    m_iPhase = 1
    m_iSendPhase = 1

End Sub
'
'   S 패킷에 대한 Y
'
Private Sub Response_S()

    Dim Tmp     As String
    Dim RSeq    As String
    
    RSeq = Mid$(RcvBuffer, 2, 1)
    
    'Tmp = ((SEQ + 32) Mod 64) & "Y~R @-#N1 "
    Tmp = RSeq & "Y~R @-#N1 "
    Tmp = ChkLen_Kermit(Tmp) & Tmp
    Tmp = Tmp & ChkSum_Kermit(Tmp)
    
    msComm.Output = Chr(1) & Tmp & Chr(13)

End Sub

'
'  결과값 편집
'
Private Sub EditData_VITROS950()
    On Error GoTo ErrHandler

    Dim ix1     As Integer

    Dim sTmpDat     As String
    Dim sTmpChr     As String
    Dim iStartN     As Integer
    Dim sPosiNo     As String
    Dim tmpBarCd$
    Dim tmpIFCd$, tmpRst$, tmpUnit$, tmpFlag$, tmpRstDT$
    Dim tmpDate$

    '결과정보 구조체 초기화
    Call Init_pResultInfo

    tmpBarCd = Trim(Mid$(sKBuffer, 26, 15))
    sPosiNo = Asc(Mid$(sKBuffer, 43, 1)) - 32

    tmpDate = Format(Now, "YYYYMMDDHHNNSS")     '임시...2005/7/14 YK

    '결과정보 구조체에 저장
    With pResultInfo
        .ID = tmpBarCd
        .SEQNO = ""
        .RACK = ""
        .POS = sPosiNo
    End With

    '----- 결과 편집
    sTmpDat = ""
    If sPosiNo = "3" Then   'CupNo가 3일 경우 '##'이므로 위치를 1증가시킨 후 시작
        iStartN = 51
    Else
        iStartN = 50
    End If

    For ix1 = iStartN To Len(sKBuffer)
        sTmpChr = Mid$(sKBuffer, ix1, 1)

        Select Case sTmpChr
            Case "}"
                tmpIFCd = Trim(Format$(Asc(Left$(sTmpDat, 1))))

                If tmpIFCd = "35" Then
                    tmpRst = Trim(Mid$(sTmpDat, 3, 9))
                    tmpFlag = Trim(Mid$(sTmpDat, 13, 1))
                Else
                    tmpRst = Trim(Mid$(sTmpDat, 2, 9))
                    tmpFlag = Trim(Mid$(sTmpDat, 12, 1))
                End If

                If tmpRst = "NO RESULT" Then
                    tmpRst = "NO RST"
                End If
                If tmpFlag = "0" Then
                    tmpFlag = ""
                End If
                
                If Left$(tmpRst, 1) = "-" Then
                    If Mid$(tmpRst, 2, 1) = "." Then
                        tmpRst = "-0" & Right$(tmpRst, Len(tmpRst) - 1)
                    ElseIf Right$(tmpRst, 1) = "." Then
                        tmpRst = Mid$(tmpRst, 1, Len(tmpRst) - 1)
                    End If
                Else
                    If Left$(tmpRst, 1) = "." Then
                        tmpRst = "0" & tmpRst
                    ElseIf Right$(tmpRst, 1) = "." Then
                        tmpRst = Mid$(tmpRst, 1, Len(tmpRst) - 1)
                    End If
                End If

                '---부등호처리---
                If tmpFlag = "4" Then
                    tmpRst = ">" & tmpRst
                ElseIf tmpFlag = "5" Then
                    tmpRst = "<" & tmpRst
                End If

                '결과값 누적
                If tmpIFCd <> "" Then
                    With pResultInfo
                        .RSTCNT = .RSTCNT + 1
                        .IFCD = .IFCD & tmpIFCd & Chr(124)
                        .RST1 = .RST1 & tmpRst & Chr(124)
                        .RST2 = .RST2 & Chr(124)
                        .UNIT = .UNIT & Chr(124)
                        .FLAG = .FLAG & tmpFlag & Chr(124)
                        .RSTDT = .RSTDT & tmpDate & Chr(124)    'TEMP....
                    End With
                End If

                sTmpDat = ""

            Case Else
                sTmpDat = sTmpDat & sTmpChr

        End Select
    Next ix1

    '결과값 등록/화면 표시 처리...
    With pResultInfo
        If .RSTCNT > 0 Then
            RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG, .ALARMCD, .KIND, .RSTDT, "")
        End If
    End With

    Call Init_pResultInfo

ErrHandler:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub

'
'   패킷의 길이 계산
'
Private Function ChkLen_Kermit(ByVal Para As String) As String

    Dim Tmp As Integer
    
    Tmp = Len(Para) + 33
    
    ChkLen_Kermit = Chr(Tmp)

End Function
Private Sub PhaseCfg_Protocol_VITROS()
    On Error GoTo ErrRtn
    
    Dim wkDat   As String
    Dim ix1     As Integer
             
    For ix1 = 1 To Len(wkBuf)
        wkDat = Mid$(wkBuf, ix1, 1)
        
        Select Case Asc(wkDat)
            Case 1          'SOH
                RcvBuffer = ""
                sP_Type = ""
                
            Case 13         '<CR>
                sP_Type = Mid$(RcvBuffer, 3, 1)     '패킷 Type
                
                Call Kermit_Phase                   'Phase 구문에 들어감
                
            Case Else       '그 외
                RcvBuffer = RcvBuffer & wkDat       '자료 누적
                
        End Select
    Next ix1
             
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg(Err.Description)
    End If
End Sub
'
'   패킷의 종류에 따른 Phase 처리
'
Private Sub Kermit_Phase()

    Dim sTmpS   As String
    Dim k       As Integer

    Select Case m_iPhase
        Case 1          'Y 대기
            Select Case sP_Type
                Case "Y"
                    Call SendOrder_VITROS     'Order 전송

                Case "E"
                    m_iPhase = 3
            End Select

        Case 2          'Y 대기
            Select Case sP_Type
                Case "Y"
                    m_iPhase = 3
                    RaiseEvent SendOrderOK(pSampleInfo.ID, "", "", "")
                    
                Case "E"
                    m_iPhase = 3
            End Select

        Case 3          'S 대기
            Select Case sP_Type
                Case "S"
                    Call Response_S     'S 에 대한 Y 송신
                    m_iPhase = 4
            End Select

        Case 4          'Z 대기
            Select Case sP_Type
                Case "S"
                    Call Response_S     'S 에 대한 Y 송신
                    m_iPhase = 4

                Case "Z"
                    Call EditData_VITROS950       'Data편집
                    Call Send_Y
                    sKBuffer = ""
                    m_iPhase = 4

                Case "B"
                    Call Send_Y
                    m_iPhase = 3

                Case "F"
                    Call Send_Y
                    m_iPhase = 4

                Case "E"
                    m_iPhase = 3

                Case "D"
                    Call Send_Y
                    sKBuffer = sKBuffer & Mid$(RcvBuffer, 4, (Len(RcvBuffer) - 4))
                    
            End Select
    End Select

End Sub
'
'   환자 Order 전송
'
Private Sub SendOrder_VITROS()
    On Error GoTo ErrRtn
    
    Dim sTmp$
    
    Select Case m_iSendPhase
        Case 1      'F 패킷, SendData 편집
            sTmp = Chr$((miOrdSeq + 32) Mod 64) & "FS" & Format$(miFileNo, "0000000")
            miFileNo = miFileNo + 1
            m_iSendPhase = 2
            
            'SendData 편집
            Call Build_SendData
            
        Case 2      'D 패킷
            '----- 누적된 D 패킷 Data를 90 Byte로 나누어 송신
            If Len(msSendData) < 90 Then
                sTmp = msSendData
                m_iSendPhase = 4
            Else
                sTmp = Left$(msSendData, 89)
                msSendData = Mid$(msSendData, 90, (Len(msSendData) - 89))
                m_iSendPhase = 3
            End If
            '------------------------------------------------
            sTmp = Chr$((miOrdSeq + 32) Mod 64) & "D" & sTmp

        Case 3      'D 패킷
            If Len(msSendData) < 90 Then
                sTmp = msSendData
                m_iSendPhase = 4
            Else
                sTmp = Left$(msSendData, 89)
                msSendData = Mid$(msSendData, 90, Len(msSendData) - 89)
                m_iSendPhase = 3
            End If
            sTmp = Chr$((miOrdSeq + 32) Mod 64) & "D" & sTmp

        Case 4      'Z 패킷
            sTmp = Chr$((miOrdSeq + 32) Mod 64) & "Z"
            m_iSendPhase = 5

        Case 5      'B 패킷(EOT)
            sTmp = Chr$((miOrdSeq + 32) Mod 64) & "B"
            m_iPhase = 2
            m_iSendPhase = 1

    End Select

    sTmp = ChkLen_Kermit(sTmp) & sTmp
    sTmp = sTmp & ChkSum_Kermit(sTmp)
    
    msComm.Output = Chr(1) & sTmp & Chr(13)
    
    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(1) & sTmp & Chr(13))
    End If
    
    miOrdSeq = miOrdSeq + 1
    
ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
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
        .QCGBN = ""
        .KIND = ""
        .RSTCNT = 0
        .IFCD = ""
        .RST1 = ""
        .RST2 = ""
        .UNIT = ""
        .FLAG = ""
        .INSTID = ""
        .ALARMCD = ""
        .RSTDT = ""
        .OTHER = ""
    End With
    
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
            .ORDCNT = 0
            .SINDEX = False
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
        .SINDEX = m_p_bSIndex
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
    m_p_bSIndex = PropBag.ReadProperty("p_bSIndex", m_def_p_bSIndex)
    m_p_sRerunGbn = PropBag.ReadProperty("p_sRerunGbn", m_def_p_sRerunGbn)
    m_p_sTSVol = PropBag.ReadProperty("p_sTSVol", m_def_p_sTSVol)
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
    Call PropBag.WriteProperty("p_bSIndex", m_p_bSIndex, m_def_p_bSIndex)
    Call PropBag.WriteProperty("p_sRerunGbn", m_p_sRerunGbn, m_def_p_sRerunGbn)
    Call PropBag.WriteProperty("p_sTSVol", m_p_sTSVol, m_def_p_sTSVol)
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
    
    'For VITROS
    miFileNo = 1
    
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
'    m_iStartSampleNo = m_def_iStartSampleNo
    m_p_bSIndex = m_def_p_bSIndex
    m_p_sRerunGbn = m_def_p_sRerunGbn
    m_p_sTSVol = m_def_p_sTSVol
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
'MemberInfo=0,0,0,0
Public Property Get p_bSIndex() As Boolean
    p_bSIndex = m_p_bSIndex
End Property

Public Property Let p_bSIndex(ByVal New_p_bSIndex As Boolean)
    m_p_bSIndex = New_p_bSIndex
    PropertyChanged "p_bSIndex"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sRerunGbn() As String
    p_sRerunGbn = m_p_sRerunGbn
End Property

Public Property Let p_sRerunGbn(ByVal New_p_sRerunGbn As String)
    m_p_sRerunGbn = New_p_sRerunGbn
    PropertyChanged "p_sRerunGbn"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=13,0,0,0
Public Property Get p_sTSVol() As String
    p_sTSVol = m_p_sTSVol
End Property

Public Property Let p_sTSVol(ByVal New_p_sTSVol As String)
    m_p_sTSVol = New_p_sTSVol
    PropertyChanged "p_sTSVol"
End Property

'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14,0,0,0
Public Property Get p_sSpcCd() As Variant
    p_sSpcCd = m_p_sSpcCd
End Property

Public Property Let p_sSpcCd(ByVal New_p_sSpcCd As Variant)
    m_p_sSpcCd = New_p_sSpcCd
    PropertyChanged "p_sSpcCd"
End Property
'
'경고! 주석으로 되어 있는 다음 줄은 제거하거나 수정하지 마십시오!
'MemberInfo=14
Public Function Send_Packet(sPacket$) As Variant
    
    Select Case sPacket
        Case "S"
            Call Send_S
    End Select
    
End Function

