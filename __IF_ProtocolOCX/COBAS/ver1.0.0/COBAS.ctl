VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.UserControl COBAS 
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   3330
   Begin VB.Timer Timer2 
      Left            =   1440
      Top             =   2295
   End
   Begin VB.Timer Timer1 
      Left            =   975
      Top             =   2280
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
Attribute VB_Name = "COBAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'기본 속성 값:
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
Event PrintRcvLog(sLog$)
Event PrintSendLog(sLog$)
Event RequestCurOrder(sID$, sRack$, sPos$)
Event DispMsg(sMsg$)
Event RequestNextOrder()
Event AppendData(sID$, sSeq$, sRack$, sPos$, iRstCnt%, sTIFCd$, sTRst1$, sTRst2$, sTUnit$, sTFlag$)


'===== User Define
'인터페이스에서 사용
Dim RcvBuffer   As String
Dim Wkbuf   As String
Dim sState  As String
Dim sReqStatusCd    As String

'구조체 지정
Private pSampleInfo As SAMPLE_INFO
Private pResultInfo As RESULT_INFO

'기타
Dim iSpaceCnt   As Integer

'For COBAS 계열
Dim iRstFlag    As Integer
Dim iOrdFlag    As Integer
Dim iIdleFlag   As Integer
Dim iPendFlag   As Integer
Private Sub PhaseCfg_Protocol_Integra700()
    
    Dim Wkdat   As String
    Dim ix1     As Integer
    
    For ix1 = 1 To Len(WkBuf1)
        Wkdat = Mid$(WkBuf1, ix1, 1)
                          
        Select Case Asc(Wkdat)
            Case 1         ' SOH
                RcvBuffer = ""
                
            Case 4         ' EOT
                Call DataEditResponse_Integra700
                
                RcvBuffer = ""
              
            Case 17, 19    ' DC1, DC3 (XON, XOFF) 삭제
           
            Case Else      ' Data
                RcvBuffer = RcvBuffer & Wkdat
                
        End Select
    Next ix1
    
End Sub
Private Sub DataEditResponse_Integra700()
    On Error GoTo ErrRtn
    
    Dim tmpBarCd$, tmpSeqNo$, tmpRack$, tmpPos$
    Dim tmpIFCd$, tmpRst$, tmpRst2$, tmpUnit$, tmpFlag$
    Dim tmpSign$, tmpExp$
    
    'For COBAS 계열
    Dim sBC     As String
    Dim sLC     As String
    Dim iBCpos  As Integer
    Dim iLCpos  As Integer
    
    Dim iErrCode        As Integer
    Dim sGeneralErrCode As String
    '--------------
    
    iErrCode = 0
    iBCpos = 22
    sBC = Mid$(RcvBuffer, iBCpos, 2)
    
    Select Case Trim(sBC)
        Case "19"           '### Order Manipulation response Block ###
            iErrCode = 99
        Case "00"           '### Idle Block ###
            iIdleFlag = 1
    '!--- Result Output Mode 에서 Transmit 되는 것이 Samples Only로 설정시 해당없음---!
        Case "02"           '### CAL Result Block ###
        Case "03"           '### Control Result Block ###
    '!--------------------------------------------------------------------------------!
        Case "04"           '### Patient Result Block ###
        Case "49"           '### Archive Manipulation response Block ###
        Case "62"           '### pending Sample Tubes Block ###
            iPendFlag = 1
        Case "69"           '### No More pending Sample Tubes Response Block ###
    End Select
        
    iLCpos = iBCpos + 5
    
    Do
        If Asc(Mid$(RcvBuffer, iLCpos, 1)) = 3 Then     'ETX(END OF DATA BLOCK)
            Exit Do
        End If
        
        sLC = Mid$(RcvBuffer, iLCpos, 2)
        
        Select Case sLC
            Case "00"       'RESULT DATA
                tmpSign = Trim(Mid$(RcvBuffer, iLCpos + 3, 1))
                tmpRst = Trim(Mid$(RcvBuffer, iLCpos + 4, 8))
                tmpExp = Mid$(RcvBuffer, iLCpos + 12, 4)
                tmpUnit = ""
                
                If tmpSign = "-" Then
                    If tmpRst = "9.999999" And Mid$(tmpExp, 3, 2) = "99" Then
                        tmpRst = "LOWER LIMIT"
                    Else
                        tmpRst = "-" & ConvertResult_Cobas(Mid$(tmpExp, 2, 1), Mid$(tmpExp, 3, 2), tmpRst)
                    End If
                Else
                    If tmpRst = "9.999999" And Mid$(tmpExp, 3, 2) = "99" Then
                        tmpRst = "UPPER LIMIT"
                    Else
                        tmpRst = ConvertResult_Cobas(Mid$(tmpExp, 2, 1), Mid$(tmpExp, 3, 2), tmpRst)
                    End If
                End If
                               
                iRstFlag = 1
                
                Exit Do
                
            Case "01"       'Result Time --> CAL, QC 일때만 전송됨
                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
                
            Case "02"       'Control ID --> CAL, QC 일때만 전송됨
                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
                
            Case "03"       'Standard Rates --> CAL, QC 일때만 전송됨
                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
                
            Case "04"       'Calibration Curve --> CAL, QC 일때만 전송됨
                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
            
            Case "07"       'ABS Sample Check --> CAL, QC 일때만 전송됨
                Exit Do     '전송 모드를 샘플모드로 셋팅시 전송안됨
                
            Case "41"       'Slot State
                '[41] + [Space] + [Rack Number Slot 1 (I3)]
                '     + [Space] + [Rack Number Slot 2 (I3)]
                '     + [Space] + [Rack Number Slot 3 (I3)]
                '     + [Space] + [Rack Number Slot 4 (I3)]
                '     + [Space] + [Rack Number Slot 5 (I3)] + [LF]
                
                'Example "41 023 128 000 000 050<LF>"
                Exit Do
                
            Case "42"       'Tube Information
                'Example "42 .21 25 1 ...AB(012:3456) URI<LF>"
                'Integra700
                tmpBarCd = Trim$(Mid$(RcvBuffer, iLCpos + 12, 15))
                
                If Trim(tmpBarCd) = "" Then
                Else
                    pSampleInfo.ID = tmpBarCd
                    
                    'Integra700
                    pSampleInfo.RACK = Trim$(Mid$(RcvBuffer, iLCpos + 3, 3))
                    pSampleInfo.POS = Trim$(Mid$(RcvBuffer, iLCpos + 7, 2))
                    
'                    'Order 가져오는 부분
'                    Call Order_Input("B")
                End If
                
                'Integra700
                iLCpos = iLCpos + 28
                
            Case "43"       'Test State
                'Example "43 032 1<LF>"
                
            Case "44"       'Cal/CS State
            
            Case "50"       'Patient ID
            
            Case "51"       'Patient Information
            
            Case "52"       'Special Order Selection
            
            Case "53"       'Order ID
                'Version 1.0
                'slipno = Trim$(Mid$(RcvBuffer, LCpos + 3, 9))
                
                'Version 2.0
                tmpBarCd = Trim(Mid(RcvBuffer, iLCpos + 3, 15))
                
                pResultInfo.ID = tmpBarCd
                
                'Version 1.0
                'LCpos = LCpos + 24  'Sample type 옵션을 No
                'LCpos = LCpos + 28  'Sample type 옵션을 Ok
                
                'Version 2.0
                iLCpos = iLCpos + 30  'Sample type 옵션을 No
                'LCpos = LCpos + 34  'Sample type 옵션을 Ok
                
            Case "55"       'Test ID
                tmpIFCd = Trim$(Mid$(RcvBuffer, iLCpos + 3, 3))
                
                iLCpos = iLCpos + 7
                
            Case "96"       'Error Code
                If iOrdFlag = 0 Then
                    'Pending Sample Request후 Response에 대한 것
                    If Mid(RcvBuffer, iLCpos + 3, 2) = "61" Then
'                        TimerFlag = 0
'                        Exit Do
                    End If
                    Exit Do
                Else
                    'Order를 내린 후 Response에 대한 것
                    If Mid$(RcvBuffer, iLCpos + 3, 2) = "00" Then
                        iErrCode = 0     'Order Input Accepted
                        Exit Do
                    Else
                        If Mid(RcvBuffer, iLCpos + 3, 2) = "22" Then
                            iErrCode = 1     'Order already available
                            Exit Do
                        ElseIf Mid(RcvBuffer, iLCpos + 3, 2) = "24" Then
                            'Test not defined - all other tests will be performed
                            iErrCode = 0
                            RaiseEvent DispMsg("일부 항목의 IF 오더코드가 잘못 설정되었습니다!!")
                            Exit Do
                        Else
                            iErrCode = 2     '기타 에러로 검사중, ID 오류, ORDER NO 오류, SAMPLE TYPE 오류 등의 에러
                            RaiseEvent DispMsg("Order 전송 중 에러발생!! 에러코드 : " & Mid$(RcvBuffer, iLCpos + 3, 2))
                            Exit Do
                        End If
                    End If
                End If
                
            Case "98"       'Protocol Version
                MsgBox Mid$(RcvBuffer, iLCpos + 3, 4)
                Exit Do
                
            Case "99"       'General Error Code
                sGeneralErrCode = Mid$(RcvBuffer, iLCpos + 3, 2)
                RaiseEvent DispMsg("General Error가 일어났습니다. 에러코드 : " & sGeneralErrCode)
                Exit Do
                
            Case Else
                Exit Do
                
        End Select
    Loop
    
'### Pending Sample Request ##############################################
    If iPendFlag = 1 And sBC = "62" Then
        PendingFlag = 0
    End If
    
'### CONNECTION CHECK ##########################################################
    If iIdleFlag = 1 And sBC = "00" Then
        iIdleFlag = 0
        
        'Ver 1.0
        Timer1.Interval = 11000     '6000
        
        'Ver 2.0
        'Timer1.Interval = 3000
    End If
    
'### NO MORE PENDING SAMPLE #####################################################
    If iPendFlag = 1 And sBC = "69" Then
        iPendFlag = 0
    End If
    
'### ORDER INPUT RESPONSE ################################################################
    'OrderFlag = 1 --> From Host To Integra : Sample Order 내린 상태
    'OrderFlag = 2 --> From Host To Integra : Order Delete를 요청한 상태
    'OrderFlag = 0 --> Order 전송이 제대로 끝난 상태
    If sBC = "19" And iErrCode = 0 Then
        If iOrdFlag = 1 Then
'            pnlStatus.Caption = gOrderTable.sSampID & "   Order OK!"
            iOrdFlag = 0   'Order 전송이 제대로 끝난 상태
'            Call Order_Next
            RaiseEvent SendOrderOK(pSampleInfo.ID, "", "")
            
            TimerFlag = 0
            
        ElseIf iOrdFlag = 2 Then
'            pnlStatus.Caption = gOrderTable.sSampID & "   Delete OK!"
'            Call Order_Input
            Call SENDORDER_INTEGRA800
        End If
        
    ElseIf sBC = "19" And iErrCode = 1 Then
        'LineCode 22의 에러발생
        RaiseEvent DispMsg("지금 Rack/Pos에 Order가 이미 존재하거나 Full(50개)인 상태입니다.!!")
        TimerFlag = 0
        RcvBuffer = ""
        Call cmdInitial.DoClick
        Exit Sub
        
    ElseIf sBC = "19" And iErrCode = 2 Then
        'LineCode 22를 제외한 에러발생
        RaiseEvent DispMsg("Order 거부!! " & _
                        "TestNo Err, Already Running, ID Err, OrderNo Err, SampleType Err 등의 에러발생...")
        TimerFlag = 0
        RcvBuffer = ""
        Call cmdInitial.DoClick
        
        Exit Sub
    End If
    
'### SAMPLE RESULT 보기 & 등록 #####################################################
    If tmpBarCd <> "" And tmpIFCd <> "" Then
        If iRstFlag = 1 And sBC = "04" Then
            '결과정보 구조체에 저장
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS

                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & tmpRst2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
            End With
                        
            TimerFlag = 0
            
            ResultFlag = 0
            
            Call Init_pResultInfo
            
            'Ver 1.0
            'Timer1.Interval = 6000
            
            'Ver 2.0
            Timer1.Interval = 1000
            Timer1.Enabled = True
        End If
    Else
        If iRstFlag = 1 And sBC = "04" Then
            iRstFlag = 0
            Timer1.Enabled = True
        End If
    End If

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
        Call Send_Initial
    End If
End Sub

Private Function ConvertResult_Cobas(ByVal sSign As String, ByVal sExp As String, ByVal sRst As String) As String
    
    Dim i%
    Dim sDot$, sDotGbn$
    Dim sValue$, sTmpVal$
    
    If IsNumeric(sRst) = False Then
        ConvertResult_Cobas = sRst
        Exit Function
    End If
    
    If sSign = "" Then
        sSign = "+"
    End If
    
    If sSign = "+" Then
        sValue = CStr(Val(sRst) * (10 ^ Val(sExp)))
    ElseIf sSign = "-" Then
        sValue = CStr(Val(sRst) / (10 ^ Val(sExp)))
    End If
    
    If Left(sValue, 1) = "." Then
        sValue = "0" & sValue
    End If
    
    ConvertResult_Cobas = sValue
    
End Function
Private Sub Order_Input()
    On Error GoTo ErrHandler
    
    ' 바코드가 없는 경우 , 환자의 오더를 일괄 전송하는 방법
    Dim SendBuff As String
    Dim i%
    
    SendBuff = ""
    
    SendBuff = Chr(1) & Chr(10)     '<SOH><LF>
    'Integra 800
    SendBuff = SendBuff & "20" & " " & "COBAS INTEGRA800" & " " & "10" & Chr(10)     '<LF>
    SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>
    SendBuff = SendBuff & "50" & " " & Space(15) & Chr(10)
  
    SendBuff = SendBuff & "53" & " " & gOrderTable.sSampID & Space(4) & _
                          " " & Format(DT1.Value, "dd/mm/yyyy") & Chr(10)   '<LF>
                          
    SendBuff = SendBuff & "54" & " " & "5" & gOrderTable.sRack & " " & gOrderTable.sPos & " " & _
                          "A" & " " & Space(21) & " " & Space(21) & Chr(10)   '<LF>
                          'gOrderTable.sOrdOpt & " " & Space(21) & " " & Space(21) & Chr(10)   '<LF>
                          
    For i = 1 To Val(gOrderTable.iOrdCnt)
        If gOrderTable.sIFTestCd(i) <> "" Then
            If Trim(Left(Trim(gOrderTable.sIFTestCd(i)) & "   ", 3)) <> "" _
                    And Trim(gOrderTable.sIFTestCd(i)) <> "930" Then
                SendBuff = SendBuff & "55" & " " & String(3 - Len(Trim(gOrderTable.sIFTestCd(i))), " ") & Trim(gOrderTable.sIFTestCd(i)) & Chr(10)
'                SendBuff = SendBuff & "55" & " " & Left(Trim(gOrderTable.sIFTestCd(i)) & "   ", 3) & Chr(10)
            End If
        End If
    Next
    
    SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>
    SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>
    
    Comm1.Output = SendBuff
    
'    Print #2, SendBuff;
    
    'OrderFlag = 1 --> From Host To Integra : Sample Order 내린 상태
    'OrderFlag = 2 --> From Host To Integra : Order Delete를 요청한 상태
    'OrderFlag = 0 --> Order 전송이 제대로 끝난 상태
    
    OrderFlag = 1
    
    Exit Sub
    
ErrHandler:
    pnlStatus.Caption = "Order_Input 에러발생" & "(" & CStr(Err.Number) & ")"
End Sub
Private Sub Timer1_Timer()
    Dim SendBuff As String
           
 '########### ALL TYPES OF FINAL RESULTS ARE TRANSFFERD TO THE HOST ######################
    SendBuff = ""
    
    SendBuff = Chr(1) & Chr(10)     '<SOH><LF>
 
 '--- HEADER BLOCK ---------------------------------------------------------------------
'    SendBuff = SendBuff & "09" & " " & "COBAS INTEGRA   " & " " & "09" & Chr(10)
'    'IC(2)^ID(16)^BC(2)<LF> -- COBASCORE IC^ID^RESULT REQUEST/RESPONSE^LF
    SendBuff = SendBuff & "20" & " " & "COBAS INTEGRA800" & " " & "09" & Chr(10)
    
    SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>
 
 '--- DATA BLOCK ------------------------------------------------------------------------
    SendBuff = SendBuff & "10" & " " & "01" & Chr(10)
    'LC(2)^RESULT SELECTOR(2)^<LF> -- RESULT REQUEST^NONSPECIFIC RESULT REQUEST^LF
    SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>
    SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>

    Comm1.Output = SendBuff
    
    '결과요구 메세지 전송
    MsgCompFlag = 1
'===============================
'    Dim SendBuff As String
'
'    If TimerFlag = 1 Then
'        Exit Sub
'    End If
'
' '########### ALL TYPES OF FINAL RESULTS ARE TRANSFFERD TO THE HOST ######################
'    SendBuff = ""
'
'    SendBuff = Chr(1) & Chr(10)     '<SOH><LF>
'
''    SendBuff = SendBuff & "06" & " " & "COBAS CORE II   " & " " & "09" & Chr(10)
''    SendBuff = SendBuff & "14" & " " & "COBAS INTEGRA400" & " " & "09" & Chr(10)
''    SendBuff = SendBuff & "09" & " " & "COBAS INTEGRA700" & " " & "09" & Chr(10)
'    SendBuff = SendBuff & "20" & " " & "COBAS INTEGRA800" & " " & "09" & Chr(10)
'
'    SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>
'    SendBuff = SendBuff & "10" & " " & "01" & Chr(10)
'    SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>
'    SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>
'
'    Comm1.Output = SendBuff
'
'    If iTestMode = 2 Then
'        Print #2, SendBuff;
'    End If
'
'    TimerFlag = 1
End Sub

Private Sub cmdInitial_Click()
    On Error GoTo ErrRtn
    
    Dim SendBuff    As String
    
    Timer1.Enabled = False
    
 '########### CONNECTION ESTABLISH ######################
    
    SendBuff = ""
    
    SendBuff = Chr(1) & Chr(10)     '<SOH><LF>
 
 '--- HEADER BLOCK ---------------------------------------------------------------------
'    SendBuff = SendBuff & "06" & " " & "COBAS INTEGRA   " & " " & "00" & Chr(10)
'    'IC(2)^ID(16)^BC(2)<LF> -- COBASCORE IC^ID^IDLE BLOCK^LF
    SendBuff = SendBuff & "20" & " " & "COBAS INTEGRA800" & " " & "00" & Chr(10)
    
    SendBuff = SendBuff & Chr(2) & Chr(10)      '<STX><LF>
    SendBuff = SendBuff & Chr(3) & Chr(10)      '<ETX><LF>
    SendBuff = SendBuff & Chr(4) & Chr(10)      '<EOT><LF>

    Comm1.Output = SendBuff
    
    IdleFlag = 1
'    '========= Add 2000.11.3 Start =============='
'    OrdSndFlag = 0
'    Timer1.Enabled = True
'    cmdSendOrder.Enabled = True
'    MsgCompFlag = 0
'    '========= Add 2000.11.3 End   =============='

ErrRtn:

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
        Case "INTEGRA700"
            Call PhaseCfg_Protocol_Integra700
            
        Case Else
            RaiseEvent DispMsg("지원되지 않는 장비를 선택했습니다.")
            
    End Select
    
End Sub
Private Sub PhaseCfg_Protocol_E170()
        
    Dim Wkdat   As String
    Dim ix1 As Integer

    For ix1 = 1 To Len(Wkbuf)
        Wkdat = Mid$(Wkbuf, ix1, 1)

        Select Case m_iPhase
            Case 1
                Select Case Asc(Wkdat)
                    Case 5      'ENQ
                        m_iPhase = 2
                        RstEnd = "Y"
                        bEndChk = True: bSTXChk = False

                        msComm.Output = Chr(6)

                    Case Else
                        m_iPhase = 1
                End Select

            Case 2
                Select Case Asc(Wkdat)
                    Case 2      'STX
                        If bEndChk = True Then
                            RcvBuffer = ""
                        Else
                            bSTXChk = True
                        End If
                        bEndChk = True

                    Case 10     '<LF>
                        If bEndChk = True Then
                            Call DataEditResponse_E170
                            RcvBuffer = ""
                        End If
                        msComm.Output = Chr(6)

                    Case 13     'CR
                        If bEndChk = True Then
                            Call DataEditResponse_E170
                            RcvBuffer = ""
                        End If

                    Case 4      'EOT
                        If sState = "Q" Then
                            msComm.Output = Chr(5)
                            m_iSendPhase = 1
                        End If
                        m_iPhase = 3

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = True
                        msComm.Output = Chr(6)   'Send ACK

                    Case 21     'NAK
                        Call DataEditResponse_E170

                        m_iSendPhase = 1
                        m_iFrameN = 1

                        msComm.Output = Chr(5)   'Send ENQ

                    Case 23     ' ETB
                        bEndChk = False

                    Case Else
                        If bEndChk = True Then
                            If bSTXChk = True Then
                                bSTXChk = False
                            Else
                                RcvBuffer = RcvBuffer & Wkdat
                            End If
                        End If

                End Select

            Case 3
                Select Case Asc(Wkdat)
                    Case 6      'ACK
                        If sState = "Q" Then
                            Call SendOrder_E170
                        End If

                    Case 5      'ENQ
                        bEndChk = True: bSTXChk = False
                        msComm.Output = Chr(6)
                        m_iPhase = 2

                    Case 21     'NAK
                        m_iSendPhase = 1
                        m_iFrameN = 1
                        msComm.Output = Chr(5)
                        m_iPhase = 3

                    Case 4      'EOT
                        m_iPhase = 1

                End Select
        End Select
    Next ix1
    
End Sub

' *=====================================================*
' *               Data편집 & 응답처리                   *
' *=====================================================*
Private Sub DataEditResponse_E170()
    On Error GoTo ErrRtn

    Dim RecType As String   'Record Type
    Dim ii      As Integer
    Dim tmpBarCd    As String
    Dim tmpSeqNo    As String
    Dim tmpRack     As String
    Dim tmpPos      As String
    Dim tmpData()   As String
    Dim tmpIFCd$, tmpRst$, tmpRst2$, tmpUnit$, tmpFlag$
   
    
    ii = InStr(1, RcvBuffer, "|")
    If ii <> 0 Then
        RecType = Mid$(RcvBuffer, ii - 1, 1)
    Else
        Exit Sub
    End If
    
    Select Case RecType
        Case "H"        'Header Record
        Case "M"
        Case "P"        'Patient Record
            Call Init_pResultInfo
            
        Case "Q"        'Order Request Record
            tmpData() = Split(RcvBuffer, "|")
            sReqStatusCd = Trim(tmpData(12))    'Order Request Status Code
            tmpData() = Split(tmpData(2), "/")
            
            tmpBarCd = Trim(tmpData(1))
            tmpSeqNo = tmpData(0)
            tmpRack = tmpData(3)
            tmpPos = tmpData(4)
            tmpData() = Split(tmpSeqNo, "^")
            tmpSeqNo = Trim(tmpData(2))
            
            If tmpBarCd <> "" Then    'BarCode ID가 잘 넘어왔는지 검사
                sState = "Q"
                pSampleInfo.ID = UCase(tmpBarCd)
            Else
                sState = ""
                pSampleInfo.ID = ""
            End If
                
            pSampleInfo.SEQNO = tmpSeqNo
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos
            
        Case "O"
            tmpSeqNo = "": tmpBarCd = "": tmpRack = "": tmpPos = ""
            tmpData() = Split(RcvBuffer, "|")
            ii = InStr(1, tmpData(2), "^")
            If ii <> 0 Then
                tmpData() = Split(tmpData(2), "^")
                tmpSeqNo = Trim(tmpData(0))
                tmpBarCd = Trim(tmpData(1))
                tmpRack = Trim(tmpData(3))
                tmpPos = Trim(tmpData(4))
            End If

            pSampleInfo.ID = UCase(tmpBarCd)
            pSampleInfo.RACK = tmpRack
            pSampleInfo.POS = tmpPos
                                    
        Case "R"        'Result Record
            '--- 결과데이타 편집
            '2:TEST ID
            '3:RESULT
            '4:UNITS
            '5:Reference Ranges
            '6:Result Abnormal Flags
            '8:Result Status(F:First,C:Rerun)
            tmpData() = Split(RcvBuffer, "|")
            
            tmpIFCd = Trim(tmpData(2))
            tmpIFCd = Mid(tmpIFCd, 4)
            tmpIFCd = Mid(tmpIFCd, 1, InStr(1, tmpIFCd, "/") - 1)
            tmpRst = Trim(tmpData(3))
            tmpRst2 = ""
            tmpUnit = Trim(tmpData(4))
            tmpFlag = Trim(tmpData(6))

            '--- 결과값에 "^" 들어갈 경우 편집
            ii = InStr(1, tmpRst, "^")
            If ii <> 0 Then
                tmpRst2 = Mid(tmpRst, 1, ii - 1)
                tmpRst = Mid(tmpRst, ii + 1)
            End If

            If Left$(tmpRst, 1) = "." Then
                tmpRst = "0" & tmpRst
            End If
            
            '결과정보 구조체에 저장
            With pResultInfo
                .ID = pSampleInfo.ID
                .SEQNO = pSampleInfo.SEQNO
                .RACK = pSampleInfo.RACK
                .POS = pSampleInfo.POS

                '결과값 누적
                .RSTCNT = .RSTCNT + 1
                .IFCD = .IFCD & tmpIFCd & Chr(124)
                .RST1 = .RST1 & tmpRst & Chr(124)
                .RST2 = .RST2 & tmpRst2 & Chr(124)
                .UNIT = .UNIT & tmpUnit & Chr(124)
                .FLAG = .FLAG & tmpFlag & Chr(124)
            End With

        Case "C"        'Comment Record
        
        Case "L"
            '결과값 등록/화면 표시 처리...
            With pResultInfo
                If .RSTCNT > 0 Then
                    RaiseEvent AppendData(.ID, .SEQNO, .RACK, .POS, .RSTCNT, .IFCD, .RST1, .RST2, .UNIT, .FLAG)
                End If
            End With

            Call Init_pResultInfo

    End Select

ErrRtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("DataEdit 오류발생 - " & Err.Description)
    End If
End Sub
'
'   환자 Order 전송
'
Private Sub SendOrder_E170()
    On Error GoTo Err_Rtn

    Dim sSendBuff   As String
    Dim iCnt    As Integer
    Dim ChkSum  As String

    Select Case m_iSendPhase
        Case 1
            'Header Record
            sSendBuff = m_iFrameN & "H|\^&|||HOST^2|||||E170^1|TSDWN^REPLY|P|1" & vbCr

            'Patient Record
            sSendBuff = sSendBuff & "P|1" & vbCr
                    
            'Order Record
            sSendBuff = sSendBuff & "O|1|" & pSampleInfo.SEQNO & "^" & Trim(pSampleInfo.ID) & "^1^" _
                    & Trim(pSampleInfo.RACK) & "^" & Trim(pSampleInfo.POS) & "|R1|"

            '----- 검사항목 조회
            RaiseEvent RequestCurOrder(pSampleInfo.ID, pSampleInfo.RACK, pSampleInfo.POS)

            Call Get_OrderString

            '검사항목 Order코드 추가
            For iCnt = 1 To pSampleInfo.ORDCNT
                'Request Information Code에 따라 검사항목을 추가하거나 취소한다.
                If Trim(sReqStatusCd) = "O" Then
                    '정상 오더
                    sSendBuff = sSendBuff & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "/\"
                ElseIf Trim(sReqStatusCd) = "A" Then
                    '오더 취소
                    sSendBuff = sSendBuff & "^^^" & Trim$(pSampleInfo.IFCD(iCnt)) & "/Clr\"
                End If
            Next iCnt
            If pSampleInfo.ORDCNT > 0 Then
                sSendBuff = Left(sSendBuff, Len(sSendBuff) - 1)      '"\" Cutting
            End If

            sSendBuff = sSendBuff & "|R||" & Format(Now, "YYYYMMDDHHNNSS") & "||||N||^^||||||" _
                    & "^^^^||||||O" & vbCr

            'Terminator Record
            sSendBuff = sSendBuff & "L|1|N"


            '--- Text의 내용이 240byte를 넘어갈 경우 처리 추가...
            If Len(sSendBuff) >= 242 Then
                sNextSend = Mid(sSendBuff, 242)
                sSendBuff = Left(sSendBuff, 241)
                sSendBuff = sSendBuff & Chr(23)

                m_iFrameN = m_iFrameN + 1
                m_iSendPhase = 2
            Else
                sSendBuff = sSendBuff & Chr(13) & Chr(3)
                GoTo Send_Terminate
            End If

        Case 2
            sSendBuff = m_iFrameN & sNextSend & Chr(13) & Chr(3)
            sNextSend = ""

Send_Terminate:
            m_iSendPhase = 3

        Case 3      'EOT
            msComm.Output = Chr(4)   'EOT
            m_iFrameN = 1
            m_iPhase = 3
            m_iSendPhase = 1

            sState = "": sReqStatusCd = ""

            Exit Sub
    End Select

    ChkSum = ChkSum_ASTM(sSendBuff)
    sSendBuff = sSendBuff & ChkSum
    msComm.Output = Chr(2) & sSendBuff & Chr(13) & Chr(10)

    If m_sTestMode = "77" Then
        RaiseEvent PrintSendLog(Chr(2) & sSendBuff & Chr(13) & Chr(10))
    End If

    '전송된 오더가 있는 경우 화면표시
    If pSampleInfo.ORDCNT > 0 And sReqStatusCd = "O" Then
        If Trim(sNextSend) = "" And m_iSendPhase <> 2 Then
            RaiseEvent SendOrderOK(pSampleInfo.ID, pSampleInfo.SEQNO, pSampleInfo.RACK, pSampleInfo.POS)
        End If
    Else
        '조회된 내용이 없는 경우 환자정보 구조체 초기화
        Call Init_pResultInfo

        RaiseEvent SendOrderOK("", "", "", "")
    End If

Err_Rtn:
    If Err <> 0 Then
        RaiseEvent DispMsg("Order 전송시 오류발생 - " & Err.Description)
    End If
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

    Wkbuf = Text1
    Call PhaseCfg_Protocol

End Sub

Private Sub msComm_OnComm()
        
    Select Case msComm.CommEvent
       ' Events
        Case MSCOMM_EV_SEND     ' There are SThreshold number of
                                ' character in the transmit buffer.
        Case MSCOMM_EV_RECEIVE  ' Received RThreshold # of chars.
            Wkbuf = msComm.Input
            
            If m_sTestMode = "77" Then
                RaiseEvent PrintRcvLog(Wkbuf)
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
    
    '변수 초기화(E-170/H-7600)
    
    
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
Public Function Send_Initial() As Variant
    
    TimerFlag = 0
    
    RcvBuffer = ""
    cmdInitial.DoClick
    
End Function

